import importlib.util
import json
import os
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


azure = types.ModuleType("azure")
azure_functions = types.ModuleType("azure.functions")


class FunctionApp:
    def __init__(self, **kwargs):
        self.tool_names = []

    def mcp_tool_trigger(self, **kwargs):
        self.tool_names.append(kwargs["tool_name"])
        return lambda function: function


azure_functions.FunctionApp = FunctionApp
azure_functions.AuthLevel = types.SimpleNamespace(FUNCTION="function")
azure.functions = azure_functions
sys.modules["azure"] = azure
sys.modules["azure.functions"] = azure_functions

requests = types.ModuleType("requests")
requests.RequestException = type("RequestException", (Exception,), {})
requests.post = Mock()
sys.modules["requests"] = requests

spec = importlib.util.spec_from_file_location(
    "function_app", Path(__file__).resolve().parents[1] / "function_app.py"
)
function_app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(function_app)


def context(**arguments):
    return json.dumps({"arguments": arguments})


class TeamsToolsTests(unittest.TestCase):
    def setUp(self):
        requests.post.reset_mock()
        requests.post.side_effect = None

    def test_only_teams_tools_are_registered(self):
        self.assertEqual(
            set(function_app.app.tool_names),
            {"enviar_respuesta_teams", "publicar_mensaje_teams"},
        )

    def test_server_metadata_describes_teams(self):
        host_path = Path(__file__).resolve().parents[1] / "host.json"
        mcp = json.loads(host_path.read_text(encoding="utf-8"))["extensions"]["mcp"]
        self.assertEqual(mcp["serverName"], "MCPTeams")
        self.assertIn("Microsoft Teams", mcp["instructions"])

    def test_invalid_arguments_do_not_send_requests(self):
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            for invalid_context in (
                "{",
                "[]",
                '{"arguments": []}',
                context(respuesta="   ", correlation_id="abc"),
                context(respuesta="hola", correlation_id=42),
                context(respuesta="x" * 20_001, correlation_id="abc"),
            ):
                result = function_app.enviar_respuesta_teams(invalid_context)
                self.assertIn("Error:", result)
        requests.post.assert_not_called()

    def test_callback_sends_expected_payload(self):
        requests.post.return_value = types.SimpleNamespace(status_code=204)
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            result = function_app.enviar_respuesta_teams(
                context(respuesta="hola", correlation_id="abc")
            )
        self.assertIn("correctamente", result)
        requests.post.assert_called_once_with(
            "https://example.test",
            json={"respuesta": "hola", "correlation_id": "abc"},
            headers={"x-callback-key": "key"},
            timeout=30,
            allow_redirects=False,
        )

    def test_success_logs_do_not_include_message_content(self):
        requests.post.return_value = types.SimpleNamespace(status_code=200)
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            with self.assertLogs(level="INFO") as logs:
                function_app.enviar_respuesta_teams(
                    context(respuesta="contenido privado", correlation_id="abc")
                )
        self.assertNotIn("contenido privado", " ".join(logs.output))

    def test_redirect_is_not_reported_as_success(self):
        requests.post.return_value = types.SimpleNamespace(status_code=302)
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            with self.assertLogs(level="ERROR"):
                result = function_app.enviar_respuesta_teams(
                    context(respuesta="hola", correlation_id="abc")
                )
        self.assertIn("HTTP 302", result)
        self.assertFalse(requests.post.call_args.kwargs["allow_redirects"])

    def test_http_error_does_not_log_response_body(self):
        requests.post.return_value = types.SimpleNamespace(
            status_code=500, text="secret-in-body"
        )
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            with self.assertLogs(level="ERROR") as logs:
                result = function_app.enviar_respuesta_teams(
                    context(respuesta="hola", correlation_id="abc")
                )
        self.assertIn("HTTP 500", result)
        self.assertNotIn("secret-in-body", " ".join(logs.output))

    def test_transport_error_does_not_expose_publish_url(self):
        requests.post.side_effect = requests.RequestException(
        "https://example.test/api/teams-publish?secret=private"
    )

        with patch.dict(
        os.environ,
        {
            "TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
            "TEAMS_PUBLISH_KEY": "private-key",
        }
    ):
            with self.assertLogs(level="ERROR") as logs:
                result = function_app.publicar_mensaje_teams(
                context(
                    destino="transformacion_digital",
                    mensaje="hola"
                )
            )

        self.assertIn("Error de conexión", result)
        self.assertNotIn("private-key", result + " ".join(logs.output))
        self.assertNotIn(
        "api/teams-publish?secret=private",
        result + " ".join(logs.output)
    )

    def test_unexpected_error_does_not_expose_details(self):
        requests.post.side_effect = RuntimeError("private-token")
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            with self.assertLogs(level="ERROR") as logs:
                result = function_app.enviar_respuesta_teams(
                    context(respuesta="hola", correlation_id="abc")
                )
        self.assertNotIn("private-token", result + " ".join(logs.output))

    def test_publish_sends_expected_payload(self):
        requests.post.return_value = types.SimpleNamespace(status_code=200)

        with patch.dict(
        os.environ,
        {
            "TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
            "TEAMS_PUBLISH_KEY": "publish-key",
        }
    ):
            result = function_app.publicar_mensaje_teams(
            context(
                destino="transformacion_digital",
                mensaje="hola"
            )
        )

        self.assertIn("publicado correctamente", result)

        requests.post.assert_called_once_with(
        "https://example.test/api/teams-publish",
        json={
            "destino": "transformacion_digital",
            "mensaje": "hola",
        },
        headers={
            "x-publish-key": "publish-key",
        },
        timeout=30,
        allow_redirects=False,
    )
    def test_publish_accepts_other_success_statuses(self):
        requests.post.return_value = types.SimpleNamespace(status_code=201)

        with patch.dict(
        os.environ,
        {
            "TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
            "TEAMS_PUBLISH_KEY": "publish-key",
        }
    ):
            result = function_app.publicar_mensaje_teams(
            context(
                destino="transformacion_digital",
                mensaje="hola"
            )
        )

        self.assertIn("publicado correctamente", result)


if __name__ == "__main__":
    unittest.main()
