import json
import logging
import os

import azure.functions as func
import requests

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

MAX_RESPUESTA_LENGTH = 20_000
MAX_CORRELATION_ID_LENGTH = 256
MAX_TITULO_LENGTH = 200
MAX_MENSAJE_LENGTH = 20_000


class InvalidArguments(ValueError):
    pass


def _arguments(context):
    try:
        content = json.loads(context)
    except (TypeError, json.JSONDecodeError) as exc:
        raise InvalidArguments("contexto MCP inválido") from exc

    if not isinstance(content, dict) or not isinstance(content.get("arguments"), dict):
        raise InvalidArguments("argumentos MCP inválidos")

    return content["arguments"]


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
            return f"Error enviando respuesta a Teams. HTTP {response.status_code}"

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
        "propertyName": "titulo",
        "propertyType": "string",
        "description": "Título del mensaje que se publicará en Microsoft Teams.",
        "isRequired": True
    },
    {
        "propertyName": "mensaje",
        "propertyType": "string",
        "description": "Contenido del mensaje que se publicará en Microsoft Teams.",
        "isRequired": True
    }
])


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="publicar_mensaje_teams",
    description=(
        "Publica un mensaje en un canal autorizado de Microsoft Teams "
        "mediante un Teams Workflow."
    ),
    tool_properties=publicar_teams_properties,
)
def publicar_mensaje_teams(context) -> str:
    try:
        args = _arguments(context)
        titulo = _required_text(args, "titulo", MAX_TITULO_LENGTH)
        mensaje = _required_text(args, "mensaje", MAX_MENSAJE_LENGTH)

        webhook_url = os.getenv("TEAMS_CHANNEL_WEBHOOK_URL")
        if not webhook_url:
            return "Error: TEAMS_CHANNEL_WEBHOOK_URL no configurada"

        payload = {
            "type": "message",
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.card.adaptive",
                    "content": {
                        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                        "type": "AdaptiveCard",
                        "version": "1.4",
                        "body": [
                            {
                                "type": "TextBlock",
                                "text": titulo,
                                "weight": "Bolder",
                                "size": "Medium",
                                "wrap": True,
                            },
                            {
                                "type": "TextBlock",
                                "text": mensaje,
                                "wrap": True,
                            },
                        ],
                    },
                }
            ],
        }

        response = requests.post(
            webhook_url, json=payload, timeout=30, allow_redirects=False
        )

        if not 200 <= response.status_code < 300:
            logging.error("Error publicando en Teams. status=%s", response.status_code)
            return f"Error publicando mensaje en Teams. HTTP {response.status_code}"

        logging.info("Solicitud de publicación aceptada por Teams Workflow")
        return "Solicitud de publicación aceptada por Teams Workflow."

    except InvalidArguments as exc:
        return f"Error: {exc}"
    except requests.RequestException as exc:
        logging.error("Error de conexión al webhook de Teams: %s", type(exc).__name__)
        return "Error de conexión al webhook de Teams"
    except Exception as exc:
        logging.error("Error ejecutando publicar_mensaje_teams: %s", type(exc).__name__)
        return "Error ejecutando publicar_mensaje_teams"
