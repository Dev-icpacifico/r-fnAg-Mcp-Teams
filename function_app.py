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
    }
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

        callback_url = os.getenv("TEAMS_CALLBACK_URL")
        callback_key = os.getenv("TEAMS_CALLBACK_KEY")

        if not callback_url:
            return "Error: TEAMS_CALLBACK_URL no configurada"
        if not callback_key:
            return "Error: TEAMS_CALLBACK_KEY no configurada"

        response = requests.post(
            callback_url,
            json={"respuesta": respuesta, "correlation_id": correlation_id},
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
            detail = CALLBACK_ERROR_MESSAGES.get(response.status_code)
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
    }
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

            detail = PUBLISH_ERROR_MESSAGES.get(response.status_code)

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
    }
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

        response = requests.post(
            publish_url,
            json={"destino": destino, "mensaje": mensaje, "caller_oid": caller_oid},
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
            detail = DIRECT_PUBLISH_ERROR_MESSAGES.get(response.status_code)
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
