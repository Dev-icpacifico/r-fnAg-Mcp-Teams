import azure.functions as func
import json
import logging
import os
import requests

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

tool_properties = json.dumps([
    {
        "propertyName": "respuesta",
        "propertyType": "string",
        "description": "Respuesta final generada por el agente.",
        "isRequired": True
    },
    {
        "propertyName": "correlation_id",
        "propertyType": "string",
        "description": "Identificador utilizado para trazabilidad.",
        "isRequired": True
    }
])


@app.mcp_tool_trigger(
    arg_name="context",
    tool_name="enviar_respuesta_prueba",
    description="Recibe la respuesta final del agente para validar la integración MCP.",
    tool_properties=tool_properties,
)
def enviar_respuesta_prueba(context) -> str:
    try:
        content = json.loads(context)
        args = content.get("arguments", {})

        respuesta = args.get("respuesta")
        correlation_id = args.get("correlation_id")

        logging.info(
            "MCP ejecutado. correlation_id=%s respuesta=%s",
            correlation_id,
            respuesta
        )

        return (
            f"Respuesta recibida correctamente. "
            f"correlation_id={correlation_id}"
        )

    except Exception as e:
        logging.exception("Error ejecutando herramienta MCP")
        return f"Error ejecutando herramienta MCP: {str(e)}"
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
        content = json.loads(context)
        args = content.get("arguments", {})

        respuesta = args.get("respuesta")
        correlation_id = args.get("correlation_id")

        if not respuesta or not correlation_id:
            return "Error: faltan respuesta o correlation_id"

        callback_url = os.getenv("TEAMS_CALLBACK_URL")
        callback_key = os.getenv("TEAMS_CALLBACK_KEY")

        if not callback_url:
            return "Error: TEAMS_CALLBACK_URL no configurada"

        if not callback_key:
            return "Error: TEAMS_CALLBACK_KEY no configurada"

        response = requests.post(
            callback_url,
            json={
                "respuesta": respuesta,
                "correlation_id": correlation_id
            },
            headers={
                "x-callback-key": callback_key
            },
            timeout=30
        )

        if response.status_code != 200:
            logging.error(
                "Error enviando respuesta a Teams. "
                "correlation_id=%s status=%s body=%s",
                correlation_id,
                response.status_code,
                response.text
            )

            return (
                f"Error enviando respuesta a Teams. "
                f"HTTP {response.status_code}"
            )

        logging.info(
            "Respuesta enviada a Teams. correlation_id=%s",
            correlation_id
        )

        return (
            f"Respuesta enviada correctamente a Teams. "
            f"correlation_id={correlation_id}"
        )

    except Exception as e:
        logging.exception("Error ejecutando enviar_respuesta_teams")
        return f"Error ejecutando enviar_respuesta_teams: {str(e)}"
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
        content = json.loads(context)
        args = content.get("arguments", {})

        titulo = args.get("titulo")
        mensaje = args.get("mensaje")

        if not titulo or not mensaje:
            return "Error: faltan titulo o mensaje"

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
                                "wrap": True
                            },
                            {
                                "type": "TextBlock",
                                "text": mensaje,
                                "wrap": True
                            }
                        ]
                    }
                }
            ]
        }

        response = requests.post(
            webhook_url,
            json=payload,
            timeout=30
        )

        if response.status_code not in (200, 202):
            logging.error(
                "Error publicando en Teams. status=%s body=%s",
                response.status_code,
                response.text
            )

            return (
                f"Error publicando mensaje en Teams. "
                f"HTTP {response.status_code}"
            )

        logging.info(
            "Mensaje publicado correctamente en Teams. titulo=%s",
            titulo
        )

        return "Mensaje publicado correctamente en Microsoft Teams."

    except Exception as e:
        logging.exception("Error ejecutando publicar_mensaje_teams")
        return f"Error ejecutando publicar_mensaje_teams: {str(e)}"