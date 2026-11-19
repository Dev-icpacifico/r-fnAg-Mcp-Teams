import json
import logging
import os
import re

import azure.functions as func
import requests

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

MAX_RESPUESTA_LENGTH = 20_000
MAX_CORRELATION_ID_LENGTH = 256
MAX_DESTINO_LENGTH = 100
MAX_MENSAJE_LENGTH = 20_000

CALLBACK_ERROR_MESSAGES = {
    404: "la solicitud no existe",
    409: "la solicitud ya fue respondida",
    410: "la solicitud expiró",
}


class InvalidArguments(ValueError):
    pass


# Identidad de quien llama al MCP. APIM la toma del token de Entra ya
# validado (claim oid) y la entrega en el encabezado CALLER_HEADER y en el
# argumento CALLER_ARGUMENT, sobrescribiendo cualquier valor del cliente.
CALLER_HEADER = "x-icp-caller-oid"
CALLER_ARGUMENT = "_icp_caller_oid"
GUID_PATTERN = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)


def _content(context):
    try:
        content = json.loads(context)
    except (TypeError, json.JSONDecodeError) as exc:
        raise InvalidArguments("contexto MCP inválido") from exc

    if not isinstance(content, dict) or not isinstance(content.get("arguments"), dict):
        raise InvalidArguments("argumentos MCP inválidos")

    return content


def _arguments(context):
    return _content(context)["arguments"]


def _transport_headers(content):
    """Encabezados HTTP de la sesión MCP, si la extensión los expone."""
    for container in (content.get("transport"), content):
        if not isinstance(container, dict):
            continue
        props = container.get("properties")
        for source in (props if isinstance(props, dict) else {}, container):
            headers = source.get("headers") if isinstance(source, dict) else None
            if isinstance(headers, dict):
                return {str(k).lower(): v for k, v in headers.items()}
    return {}


def _caller_oid(content):
    """oid de Entra de quien llama. Prefiere el encabezado (el modelo no
    puede fijarlo); si la extensión no expone encabezados, usa el argumento
    que APIM inyecta sobre el cuerpo JSON-RPC."""
    value = _transport_headers(content).get(CALLER_HEADER)
    if isinstance(value, list):
        value = value[0] if value else None
    source = "header"
    if not value:
        value = content["arguments"].get(CALLER_ARGUMENT)
        source = "argument"
    if isinstance(value, str) and GUID_PATTERN.match(value.strip()):
        return value.strip().lower(), source
    return None, "missing"


MAX_MENCIONES_LENGTH = 2_000
MAX_DESTINATARIO_LENGTH = 320
MAX_NOMBRE_ARCHIVO_LENGTH = 200
# 4 MB en base64 ocupan ~5,6 millones de caracteres.
MAX_CONTENIDO_BASE64_LENGTH = 5_600_000

MENCIONES_PROPERTY = {
    "propertyName": "menciones",
    "propertyType": "string",
    "description": (
        "Opcional. Correos (o Object ID de Entra) de las personas a mencionar, "
        "separados por coma. En el mensaje escribe @correo donde quieras la "
        "mención; si no aparece, se menciona al inicio. Sólo se puede mencionar "
        "a miembros del equipo del canal. Teams no permite mencionar el canal ni "
        "a todo el equipo."
    ),
    "isRequired": False,
}


def _optional_text(args, name, max_length):
    value = args.get(name)
    if value in (None, ""):
        return None
    if isinstance(value, list):
        value = ",".join(str(v) for v in value)
    if not isinstance(value, str):
        raise InvalidArguments(f"{name} debe ser texto")
    if len(value) > max_length:
        raise InvalidArguments(f"{name} supera el tamaño permitido")
    return value.strip() or None


def _adapter_detail(response, known):
    """Mensaje para el agente: el conocido para el código o el del adaptador."""
    detail = known.get(response.status_code)
    if detail:
        return detail
    try:
        body = response.json()
    except Exception:  # noqa: BLE001
        return None
    error = body.get("error") if isinstance(body, dict) else None
    return str(error)[:300] if error else None


def _adapter_url(env_name, suffix):
    url = os.getenv(env_name)
    if url:
        return url
    base = (os.getenv("TEAMS_PUBLISH_URL") or "").rstrip("/")
    if base.endswith("/api/teams-publish"):
        return base[: -len("/api/teams-publish")] + suffix
    return None


def _required_text(args, name, max_length):
    value = args.get(name)
    if not isinstance(value, str) or not value.strip():
        raise InvalidArguments(f"falta {name} o no es texto")
    if len(value) > max_length:
        raise InvalidArguments(f"{name} supera el tamaño permitido")
    return value

teams_tool_properties = json.dumps([
    {
        "propertyName": "respuesta",
        "propertyType": "string",
        "description": "Respuesta final generada por el agente que debe enviarse a Microsoft Teams.",
        "isRequired": True
    },
    {
        "propertyName": "correlation_id",
        "propertyType": "string",
        "description": "Identificador de correlación de la conversación de Teams.",
        "isRequired": True
    },
    MENCIONES_PROPERTY,
])


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="enviar_respuesta_teams",
    description="Envía la respuesta final del agente a la conversación de Microsoft Teams asociada al correlation_id.",
    tool_properties=teams_tool_properties,
)
def enviar_respuesta_teams(context) -> str:
    try:
        args = _arguments(context)
        respuesta = _required_text(args, "respuesta", MAX_RESPUESTA_LENGTH)
        correlation_id = _required_text(args, "correlation_id", MAX_CORRELATION_ID_LENGTH)
        menciones = _optional_text(args, "menciones", MAX_MENCIONES_LENGTH)

        callback_url = os.getenv("TEAMS_CALLBACK_URL")
        callback_key = os.getenv("TEAMS_CALLBACK_KEY")

        if not callback_url:
            return "Error: TEAMS_CALLBACK_URL no configurada"
        if not callback_key:
            return "Error: TEAMS_CALLBACK_KEY no configurada"

        payload = {"respuesta": respuesta, "correlation_id": correlation_id}
        if menciones:
            payload["menciones"] = menciones
        response = requests.post(
            callback_url,
            json=payload,
            headers={"x-callback-key": callback_key},
            timeout=30,
            allow_redirects=False,
        )

        if not 200 <= response.status_code < 300:
            logging.error(
                "Error enviando respuesta a Teams. correlation_id=%r status=%s",
                correlation_id[:128],
                response.status_code,
            )
            detail = _adapter_detail(response, CALLBACK_ERROR_MESSAGES)
            return (
                f"Error enviando respuesta a Teams. HTTP {response.status_code}"
                + (f": {detail}" if detail else "")
            )

        logging.info("Respuesta enviada a Teams. correlation_id=%r", correlation_id[:128])
        return f"Respuesta enviada correctamente a Teams. correlation_id={correlation_id}"

    except InvalidArguments as exc:
        return f"Error: {exc}"
    except requests.RequestException as exc:
        logging.error("Error de conexión al callback de Teams: %s", type(exc).__name__)
        return "Error de conexión al callback de Teams"
    except Exception as exc:
        logging.error("Error ejecutando enviar_respuesta_teams: %s", type(exc).__name__)
        return "Error ejecutando enviar_respuesta_teams"


publicar_teams_properties = json.dumps([
    {
        "propertyName": "destino",
        "propertyType": "string",
        "description": (
            "Alias del destino autorizado de Microsoft Teams. "
            "Por ejemplo: transformacion_digital."
        ),
        "isRequired": True
    },
    {
        "propertyName": "mensaje",
        "propertyType": "string",
        "description": "Contenido del mensaje que se publicará en Microsoft Teams.",
        "isRequired": True
    },
    {
        "propertyName": "correlation_id",
        "propertyType": "string",
        "description": (
            "Identificador de la solicitud de Teams que originó la publicación. "
            "Enviarlo siempre que exista: el destino debe pertenecer a la misma "
            "área que esa solicitud (ver destinos_permitidos en la solicitud)."
        ),
        "isRequired": False
    },
    MENCIONES_PROPERTY,
])

PUBLISH_ERROR_MESSAGES = {
    403: "el destino no está autorizado para esta solicitud o su área",
    404: "el destino no existe",
}


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="publicar_mensaje_teams",
    description=(
        "Publica un mensaje en un destino autorizado de Microsoft Teams "
        "utilizando su alias de destino. Sólo se permiten destinos del área "
        "de la solicitud indicada por correlation_id."
    ),
    tool_properties=publicar_teams_properties,
)
def publicar_mensaje_teams(context) -> str:
    try:
        args = _arguments(context)

        destino = _required_text(
            args,
            "destino",
            MAX_DESTINO_LENGTH
        )

        mensaje = _required_text(
            args,
            "mensaje",
            MAX_MENSAJE_LENGTH
        )

        correlation_id = None
        if args.get("correlation_id") not in (None, ""):
            correlation_id = _required_text(
                args,
                "correlation_id",
                MAX_CORRELATION_ID_LENGTH
            )

        publish_url = os.getenv("TEAMS_PUBLISH_URL")
        publish_key = os.getenv("TEAMS_PUBLISH_KEY")

        if not publish_url:
            return "Error: TEAMS_PUBLISH_URL no configurada"

        if not publish_key:
            return "Error: TEAMS_PUBLISH_KEY no configurada"

        payload = {
            "destino": destino,
            "mensaje": mensaje
        }

        if correlation_id:
            payload["correlation_id"] = correlation_id

        menciones = _optional_text(args, "menciones", MAX_MENCIONES_LENGTH)
        if menciones:
            payload["menciones"] = menciones

        response = requests.post(
            publish_url,
            json=payload,
            headers={
                "x-publish-key": publish_key
            },
            timeout=30,
            allow_redirects=False,
        )

        if not 200 <= response.status_code < 300:
            logging.error(
                "Error publicando en Teams. destino=%r status=%s",
                destino[:100],
                response.status_code,
            )

            detail = _adapter_detail(response, PUBLISH_ERROR_MESSAGES)

            return (
                "Error publicando mensaje en Teams. "
                f"HTTP {response.status_code}"
                + (f": {detail}" if detail else "")
            )

        logging.info(
            "Mensaje publicado correctamente en Teams. destino=%r",
            destino[:100],
        )

        return (
            "Mensaje publicado correctamente en Teams. "
            f"destino={destino}"
        )

    except InvalidArguments as exc:
        return f"Error: {exc}"

    except requests.RequestException as exc:
        logging.error(
            "Error de conexión al Teams Adapter: %s",
            type(exc).__name__
        )
        return "Error de conexión al Teams Adapter"

    except Exception as exc:
        logging.error(
            "Error ejecutando publicar_mensaje_teams: %s",
            type(exc).__name__
        )
        return "Error ejecutando publicar_mensaje_teams"


publicar_canal_properties = json.dumps([
    {
        "propertyName": "destino",
        "propertyType": "string",
        "description": "Alias del canal de Teams donde publicar. Por ejemplo: td_publicaciones.",
        "isRequired": True
    },
    {
        "propertyName": "mensaje",
        "propertyType": "string",
        "description": "Contenido del mensaje que se publicará en el canal.",
        "isRequired": True
    },
    MENCIONES_PROPERTY,
])

DIRECT_PUBLISH_ERROR_MESSAGES = {
    401: "no se pudo identificar a la cuenta que publica",
    403: "la cuenta conectada no está autorizada para publicar en ese canal",
    404: "el destino no existe",
}


def _direct_publish_url():
    url = os.getenv("TEAMS_PUBLISH_DIRECT_URL")
    if url:
        return url
    base = os.getenv("TEAMS_PUBLISH_URL") or ""
    if base.rstrip("/").endswith("/api/teams-publish"):
        return base.rstrip("/") + "-direct"
    return None


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="publicar_en_canal_teams",
    description=(
        "Publica un mensaje en un canal de Microsoft Teams cuando la publicación "
        "se inicia desde ChatGPT (una tarea programada o un pedido en esta "
        "conversación) y no existe una solicitud de Teams con correlation_id. "
        "Sólo funciona para cuentas autorizadas como publicadoras del canal. "
        "Si la solicitud llegó desde Teams (hay correlation_id), usa "
        "publicar_mensaje_teams."
    ),
    tool_properties=publicar_canal_properties,
)
def publicar_en_canal_teams(context) -> str:
    try:
        content = _content(context)
        args = content["arguments"]
        destino = _required_text(args, "destino", MAX_DESTINO_LENGTH)
        mensaje = _required_text(args, "mensaje", MAX_MENSAJE_LENGTH)

        caller_oid, source = _caller_oid(content)
        if caller_oid is None:
            logging.error(
                "Publicación directa sin identidad. claves_contexto=%s",
                sorted(content.keys()),
            )
            return "Error: no se pudo identificar a la cuenta que publica"

        publish_url = _direct_publish_url()
        publish_key = os.getenv("TEAMS_PUBLISH_KEY")
        if not publish_url:
            return "Error: TEAMS_PUBLISH_DIRECT_URL no configurada"
        if not publish_key:
            return "Error: TEAMS_PUBLISH_KEY no configurada"

        payload = {"destino": destino, "mensaje": mensaje, "caller_oid": caller_oid}
        menciones = _optional_text(args, "menciones", MAX_MENCIONES_LENGTH)
        if menciones:
            payload["menciones"] = menciones
        response = requests.post(
            publish_url,
            json=payload,
            headers={"x-publish-key": publish_key},
            timeout=30,
            allow_redirects=False,
        )

        if not 200 <= response.status_code < 300:
            logging.error(
                "Error en publicación directa. destino=%r status=%s origen_identidad=%s",
                destino[:100],
                response.status_code,
                source,
            )
            detail = _adapter_detail(response, DIRECT_PUBLISH_ERROR_MESSAGES)
            return (
                "Error publicando mensaje en Teams. "
                f"HTTP {response.status_code}"
                + (f": {detail}" if detail else "")
            )

        logging.info(
            "Publicación directa correcta. destino=%r origen_identidad=%s",
            destino[:100],
            source,
        )
        return f"Mensaje publicado correctamente en Teams. destino={destino}"

    except InvalidArguments as exc:
        return f"Error: {exc}"
    except requests.RequestException as exc:
        logging.error("Error de conexión al Teams Adapter: %s", type(exc).__name__)
        return "Error de conexión al Teams Adapter"
    except Exception as exc:
        logging.error("Error ejecutando publicar_en_canal_teams: %s", type(exc).__name__)
        return "Error ejecutando publicar_en_canal_teams"


DIRECT_MESSAGE_ERROR_MESSAGES = {
    401: "no se pudo identificar a la cuenta que envía",
}

PERSONAL_DESCRIPTION = (
    "Puede recibirlo cualquier persona de la empresa (no invitados). Si la "
    "solicitud llegó desde Teams, envía su correlation_id. Si no (pedido "
    "desde ChatGPT), indica tu alias de agente en «agente»: la cuenta "
    "conectada debe tener acceso a ese agente."
)

AGENTE_PROPERTY = {
    "propertyName": "agente",
    "propertyType": "string",
    "description": "Alias del agente que envía (por ejemplo noticias). Obligatorio si no hay correlation_id.",
    "isRequired": False,
}

mensaje_directo_properties = json.dumps([
    {
        "propertyName": "destinatario",
        "propertyType": "string",
        "description": "Correo (o Object ID de Entra) de la persona que recibirá el mensaje.",
        "isRequired": True,
    },
    {
        "propertyName": "mensaje",
        "propertyType": "string",
        "description": "Texto del mensaje (admite formato markdown y enlaces).",
        "isRequired": True,
    },
    {
        "propertyName": "correlation_id",
        "propertyType": "string",
        "description": "Identificador de la solicitud de Teams que originó el envío, si existe.",
        "isRequired": False,
    },
    AGENTE_PROPERTY,
])

archivo_properties = json.dumps([
    {
        "propertyName": "destinatario",
        "propertyType": "string",
        "description": "Correo (o Object ID de Entra) de la persona que recibirá el archivo.",
        "isRequired": True,
    },
    {
        "propertyName": "titulo",
        "propertyType": "string",
        "description": "Sólo con «contenido»: título del PDF simple que genera el bot.",
        "isRequired": False,
    },
    {
        "propertyName": "contenido",
        "propertyType": "string",
        "description": (
            "Alternativa sin archivo propio: texto en markdown (títulos #, listas, "
            "tablas, enlaces) con el que el bot genera un PDF simple. No replica el "
            "diseño del boletín; prefiere contenido_base64."
        ),
        "isRequired": False,
    },
    {
        "propertyName": "nombre_archivo",
        "propertyType": "string",
        "description": "Nombre del archivo con extensión, por ejemplo Boletin_ICP_2026-11-19.pdf. Obligatorio con contenido_base64.",
        "isRequired": False,
    },
    {
        "propertyName": "contenido_base64",
        "propertyType": "string",
        "description": "Recomendado: el archivo que generaste (por ejemplo el PDF del boletín) codificado en base64, completo. Máximo 4 MB.",
        "isRequired": False,
    },
    {
        "propertyName": "mensaje",
        "propertyType": "string",
        "description": "Opcional. Texto que acompaña al archivo.",
        "isRequired": False,
    },
    {
        "propertyName": "correlation_id",
        "propertyType": "string",
        "description": "Identificador de la solicitud de Teams que originó el envío, si existe.",
        "isRequired": False,
    },
    AGENTE_PROPERTY,
])

MAX_CONTENIDO_LENGTH = 200_000


def _personal_payload(content, args):
    """Destinatario, solicitud de origen e identidad de quien llama."""
    payload = {
        "destinatario": _required_text(args, "destinatario", MAX_DESTINATARIO_LENGTH).strip(),
    }
    correlation_id = _optional_text(args, "correlation_id", MAX_CORRELATION_ID_LENGTH)
    if correlation_id:
        payload["correlation_id"] = correlation_id
    caller_oid, source = _caller_oid(content)
    if caller_oid:
        payload["caller_oid"] = caller_oid
    elif not correlation_id:
        logging.error("Envío personal sin identidad ni correlation_id. origen=%s", source)
        raise InvalidArguments("no se pudo identificar a la cuenta que envía")
    agente = _optional_text(args, "agente", 32)
    if agente:
        payload["agente"] = agente
    elif not correlation_id:
        raise InvalidArguments("sin correlation_id hay que indicar el alias del agente en «agente»")
    return payload


def _post_personal(env_name, suffix, payload, tool):
    url = _adapter_url(env_name, suffix)
    key = os.getenv("TEAMS_PUBLISH_KEY")
    if not url:
        return None, f"Error: {env_name} no configurada"
    if not key:
        return None, "Error: TEAMS_PUBLISH_KEY no configurada"
    response = requests.post(
        url, json=payload, headers={"x-publish-key": key}, timeout=60, allow_redirects=False,
    )
    if not 200 <= response.status_code < 300:
        logging.error("Error en %s. status=%s", tool, response.status_code)
        detail = _adapter_detail(response, DIRECT_MESSAGE_ERROR_MESSAGES)
        return None, f"Error en {tool}. HTTP {response.status_code}" + (f": {detail}" if detail else "")
    return response, None


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="enviar_mensaje_directo_teams",
    description=(
        "Envía un mensaje al chat personal de una persona en Microsoft Teams, "
        "aunque no haya escrito antes al bot. " + PERSONAL_DESCRIPTION
    ),
    tool_properties=mensaje_directo_properties,
)
def enviar_mensaje_directo_teams(context) -> str:
    try:
        content = _content(context)
        args = content["arguments"]
        payload = _personal_payload(content, args)
        payload["mensaje"] = _required_text(args, "mensaje", MAX_MENSAJE_LENGTH)
        _, error = _post_personal(
            "TEAMS_DIRECT_MESSAGE_URL", "/api/teams-direct-message", payload, "enviar_mensaje_directo_teams"
        )
        if error:
            return error
        logging.info("Mensaje directo enviado")
        return f"Mensaje directo enviado a {payload['destinatario']}."
    except InvalidArguments as exc:
        return f"Error: {exc}"
    except requests.RequestException as exc:
        logging.error("Error de conexión al Teams Adapter: %s", type(exc).__name__)
        return "Error de conexión al Teams Adapter"
    except Exception as exc:
        logging.error("Error ejecutando enviar_mensaje_directo_teams: %s", type(exc).__name__)
        return "Error ejecutando enviar_mensaje_directo_teams"


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="enviar_archivo_teams",
    description=(
        "Envía un archivo al chat personal de una persona en Microsoft Teams. "
        "Envía el archivo que generaste en nombre_archivo + contenido_base64 "
        "(máximo 4 MB). Sólo si no tienes archivo, titulo + contenido (markdown) "
        "para que el bot genere un PDF simple. Teams le pide aceptarlo y se guarda "
        "en su OneDrive. "
        + PERSONAL_DESCRIPTION
    ),
    tool_properties=archivo_properties,
)
def enviar_archivo_teams(context) -> str:
    try:
        content = _content(context)
        args = content["arguments"]
        payload = _personal_payload(content, args)
        contenido = _optional_text(args, "contenido", MAX_CONTENIDO_LENGTH)
        if contenido:
            payload["formato"] = "pdf"
            payload["contenido"] = contenido
            payload["titulo"] = _required_text(args, "titulo", 150).strip()
            nombre = _optional_text(args, "nombre_archivo", MAX_NOMBRE_ARCHIVO_LENGTH)
            if nombre:
                payload["nombre_archivo"] = nombre
        else:
            payload["nombre_archivo"] = _required_text(args, "nombre_archivo", MAX_NOMBRE_ARCHIVO_LENGTH).strip()
            payload["contenido_base64"] = _required_text(args, "contenido_base64", MAX_CONTENIDO_BASE64_LENGTH)
        mensaje = _optional_text(args, "mensaje", 2_000)
        if mensaje:
            payload["mensaje"] = mensaje
        _, error = _post_personal("TEAMS_FILE_URL", "/api/teams-file", payload, "enviar_archivo_teams")
        if error:
            return error
        logging.info("Archivo ofrecido por chat personal")
        return (
            f"Archivo {payload.get('nombre_archivo') or payload.get('titulo')} enviado a {payload['destinatario']}. "
            "La persona debe aceptarlo en Teams para recibirlo."
        )
    except InvalidArguments as exc:
        return f"Error: {exc}"
    except requests.RequestException as exc:
        logging.error("Error de conexión al Teams Adapter: %s", type(exc).__name__)
        return "Error de conexión al Teams Adapter"
    except Exception as exc:
        logging.error("Error ejecutando enviar_archivo_teams: %s", type(exc).__name__)
        return "Error ejecutando enviar_archivo_teams"
