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
            {"enviar_respuesta_teams", "publicar_mensaje_teams", "publicar_en_canal_teams",
             "enviar_mensaje_directo_teams", "enviar_archivo_teams"},
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



class PublishCorrelationTests(unittest.TestCase):
    ENV = {
        "TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
        "TEAMS_PUBLISH_KEY": "publish-key",
    }

    def setUp(self):
        requests.post.reset_mock()
        requests.post.side_effect = None

    def test_publish_forwards_correlation_id(self):
        requests.post.return_value = types.SimpleNamespace(status_code=200)
        with patch.dict(os.environ, self.ENV):
            result = function_app.publicar_mensaje_teams(
                context(destino="rrhh_reportes", mensaje="informe", correlation_id="abc-123")
            )
        self.assertIn("publicado correctamente", result)
        self.assertEqual(
            requests.post.call_args.kwargs["json"],
            {"destino": "rrhh_reportes", "mensaje": "informe", "correlation_id": "abc-123"},
        )

    def test_empty_correlation_id_is_omitted(self):
        requests.post.return_value = types.SimpleNamespace(status_code=200)
        with patch.dict(os.environ, self.ENV):
            function_app.publicar_mensaje_teams(
                context(destino="rrhh_reportes", mensaje="informe", correlation_id="")
            )
        self.assertNotIn("correlation_id", requests.post.call_args.kwargs["json"])

    def test_invalid_correlation_id_does_not_send(self):
        with patch.dict(os.environ, self.ENV):
            for value in (42, "   ", "x" * 257):
                result = function_app.publicar_mensaje_teams(
                    context(destino="rrhh_reportes", mensaje="informe", correlation_id=value)
                )
                self.assertIn("Error:", result)
        requests.post.assert_not_called()

    def test_forbidden_destination_is_explained(self):
        requests.post.return_value = types.SimpleNamespace(status_code=403, text="detalle interno")
        with patch.dict(os.environ, self.ENV):
            with self.assertLogs(level="ERROR"):
                result = function_app.publicar_mensaje_teams(
                    context(destino="contabilidad", mensaje="informe", correlation_id="abc")
                )
        self.assertIn("HTTP 403", result)
        self.assertIn("no está autorizado", result)
        self.assertNotIn("detalle interno", result)

    def test_callback_already_answered_is_explained(self):
        requests.post.return_value = types.SimpleNamespace(status_code=409)
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test", "TEAMS_CALLBACK_KEY": "key"}):
            with self.assertLogs(level="ERROR"):
                result = function_app.enviar_respuesta_teams(
                    context(respuesta="hola", correlation_id="abc")
                )
        self.assertIn("ya fue respondida", result)


if __name__ == "__main__":
    unittest.main()


OID = "AbCdEf01-2345-6789-abcd-ef0123456789"
DIRECT_ENV = {
    "TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
    "TEAMS_PUBLISH_KEY": "key",
}


def direct_context(arguments, headers=None, transport=True):
    content = {"arguments": arguments}
    if headers is not None:
        if transport:
            content["transport"] = {"name": "http-streamable", "properties": {"headers": headers}}
        else:
            content["headers"] = headers
    return json.dumps(content)


class DirectPublishTests(unittest.TestCase):
    def setUp(self):
        requests.post.reset_mock()
        requests.post.side_effect = None
        requests.post.return_value = types.SimpleNamespace(status_code=200)

    def call(self, ctx, env=DIRECT_ENV):
        with patch.dict(os.environ, env, clear=True):
            return function_app.publicar_en_canal_teams(ctx)

    def test_uses_caller_from_transport_header(self):
        result = self.call(direct_context(
            {"destino": "td_publicaciones", "mensaje": "hola"},
            headers={"X-ICP-Caller-Oid": OID},
        ))
        self.assertIn("correctamente", result)
        requests.post.assert_called_once_with(
            "https://example.test/api/teams-publish-direct",
            json={"destino": "td_publicaciones", "mensaje": "hola", "caller_oid": OID.lower()},
            headers={"x-publish-key": "key"},
            timeout=30,
            allow_redirects=False,
        )

    def test_header_wins_over_argument(self):
        other = "11111111-2222-3333-4444-555555555555"
        self.call(direct_context(
            {"destino": "td", "mensaje": "hola", "_icp_caller_oid": other},
            headers={"x-icp-caller-oid": [OID]},
        ))
        self.assertEqual(requests.post.call_args.kwargs["json"]["caller_oid"], OID.lower())

    def test_falls_back_to_injected_argument(self):
        self.call(direct_context({"destino": "td", "mensaje": "hola", "_icp_caller_oid": OID}))
        self.assertEqual(requests.post.call_args.kwargs["json"]["caller_oid"], OID.lower())

    def test_top_level_headers_are_read(self):
        self.call(direct_context({"destino": "td", "mensaje": "hola"},
                                 headers={"x-icp-caller-oid": OID}, transport=False))
        self.assertEqual(requests.post.call_args.kwargs["json"]["caller_oid"], OID.lower())

    def test_missing_or_invalid_caller_does_not_send(self):
        for ctx in (
            direct_context({"destino": "td", "mensaje": "hola"}),
            direct_context({"destino": "td", "mensaje": "hola", "_icp_caller_oid": "admin"}),
            direct_context({"destino": "td", "mensaje": "hola"}, headers={"x-icp-caller-oid": ""}),
        ):
            with self.assertLogs(level="ERROR"):
                result = self.call(ctx)
            self.assertIn("no se pudo identificar", result)
        requests.post.assert_not_called()

    def test_explicit_direct_url_is_used(self):
        env = {**DIRECT_ENV, "TEAMS_PUBLISH_DIRECT_URL": "https://other.test/direct"}
        self.call(direct_context({"destino": "td", "mensaje": "hola", "_icp_caller_oid": OID}), env)
        self.assertEqual(requests.post.call_args.args[0], "https://other.test/direct")

    def test_missing_configuration(self):
        ctx = direct_context({"destino": "td", "mensaje": "hola", "_icp_caller_oid": OID})
        self.assertIn("TEAMS_PUBLISH_DIRECT_URL", self.call(ctx, {"TEAMS_PUBLISH_KEY": "key"}))
        self.assertIn("TEAMS_PUBLISH_KEY", self.call(ctx, {"TEAMS_PUBLISH_URL": DIRECT_ENV["TEAMS_PUBLISH_URL"]}))
        requests.post.assert_not_called()

    def test_forbidden_is_explained_without_leaking_message(self):
        requests.post.return_value = types.SimpleNamespace(status_code=403, text="detalle interno")
        with self.assertLogs(level="ERROR") as logs:
            result = self.call(direct_context(
                {"destino": "td", "mensaje": "contenido privado", "_icp_caller_oid": OID}))
        self.assertIn("HTTP 403", result)
        self.assertIn("no está autorizada", result)
        joined = " ".join(logs.output)
        self.assertNotIn("contenido privado", joined)
        self.assertNotIn("detalle interno", joined)

    def test_invalid_arguments_do_not_send(self):
        for ctx in ("{", direct_context({"destino": "", "mensaje": "hola", "_icp_caller_oid": OID})):
            self.assertIn("Error:", self.call(ctx))
        requests.post.assert_not_called()


PERSONAL_ENV = {"TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish", "TEAMS_PUBLISH_KEY": "key"}


class PersonalToolsTests(unittest.TestCase):
    def setUp(self):
        requests.post.reset_mock()
        requests.post.side_effect = None
        requests.post.return_value = types.SimpleNamespace(status_code=200)

    def call(self, tool, ctx, env=PERSONAL_ENV):
        with patch.dict(os.environ, env, clear=True):
            return getattr(function_app, tool)(ctx)

    def test_direct_message_with_correlation(self):
        result = self.call("enviar_mensaje_directo_teams", context(
            destinatario="lpizarro@icp.cl", mensaje="Hola", correlation_id="corr-1"))
        self.assertIn("enviado", result)
        requests.post.assert_called_once_with(
            "https://example.test/api/teams-direct-message",
            json={"destinatario": "lpizarro@icp.cl", "correlation_id": "corr-1", "mensaje": "Hola"},
            headers={"x-publish-key": "key"},
            timeout=60,
            allow_redirects=False,
        )

    def test_direct_message_without_correlation_sends_caller_and_agent(self):
        self.call("enviar_mensaje_directo_teams", direct_context(
            {"destinatario": "otro@icp.cl", "mensaje": "Hola", "agente": "noticias"},
            headers={"x-icp-caller-oid": OID}))
        body = requests.post.call_args.kwargs["json"]
        self.assertEqual((body["caller_oid"], body["agente"]), (OID.lower(), "noticias"))

    def test_without_correlation_the_agent_alias_is_required(self):
        result = self.call("enviar_mensaje_directo_teams", direct_context(
            {"destinatario": "otro@icp.cl", "mensaje": "Hola"}, headers={"x-icp-caller-oid": OID}))
        self.assertIn("agente", result)
        requests.post.assert_not_called()

    def test_pdf_from_markdown_content(self):
        result = self.call("enviar_archivo_teams", context(
            destinatario="a@icp.cl", titulo="Noticias 19-11", contenido="# Hoy\n- dato",
            correlation_id="corr-1"))
        self.assertIn("Noticias 19-11", result)
        body = requests.post.call_args.kwargs["json"]
        self.assertEqual((body["formato"], body["titulo"], body["contenido"]), ("pdf", "Noticias 19-11", "# Hoy\n- dato"))
        self.assertNotIn("contenido_base64", body)
        missing_title = self.call("enviar_archivo_teams", context(
            destinatario="a@icp.cl", contenido="x", correlation_id="c"))
        self.assertIn("titulo", missing_title)

    def test_direct_message_without_identity_is_rejected(self):
        with self.assertLogs(level="ERROR"):
            result = self.call("enviar_mensaje_directo_teams", context(destinatario="a@icp.cl", mensaje="x"))
        self.assertIn("identificar", result)
        requests.post.assert_not_called()

    def test_file_is_forwarded_with_its_own_url(self):
        result = self.call("enviar_archivo_teams", context(
            destinatario="a@icp.cl", nombre_archivo="informe.pdf", contenido_base64="JVBERg==",
            mensaje="Informe", correlation_id="corr-1"))
        self.assertIn("debe aceptarlo", result)
        self.assertEqual(requests.post.call_args.args[0], "https://example.test/api/teams-file")
        body = requests.post.call_args.kwargs["json"]
        self.assertEqual((body["nombre_archivo"], body["contenido_base64"], body["mensaje"]),
                         ("informe.pdf", "JVBERg==", "Informe"))

    def test_adapter_error_message_reaches_the_agent(self):
        requests.post.return_value = types.SimpleNamespace(
            status_code=403, json=lambda: {"error": "La persona no tiene acceso a este agente"})
        with self.assertLogs(level="ERROR"):
            result = self.call("enviar_mensaje_directo_teams", context(
                destinatario="a@icp.cl", mensaje="x", correlation_id="c"))
        self.assertIn("HTTP 403: La persona no tiene acceso", result)

    def test_mentions_are_forwarded_only_when_present(self):
        with patch.dict(os.environ, {"TEAMS_PUBLISH_URL": "https://example.test/api/teams-publish",
                                     "TEAMS_PUBLISH_KEY": "key"}, clear=True):
            function_app.publicar_mensaje_teams(context(destino="td", mensaje="hola @a@icp.cl",
                                                        menciones=["a@icp.cl", "b@icp.cl"]))
            self.assertEqual(requests.post.call_args.kwargs["json"]["menciones"], "a@icp.cl,b@icp.cl")
            function_app.publicar_mensaje_teams(context(destino="td", mensaje="hola"))
            self.assertNotIn("menciones", requests.post.call_args.kwargs["json"])
        with patch.dict(os.environ, {"TEAMS_CALLBACK_URL": "https://example.test/cb",
                                     "TEAMS_CALLBACK_KEY": "k"}, clear=True):
            function_app.enviar_respuesta_teams(context(respuesta="r", correlation_id="c", menciones="a@icp.cl"))
            self.assertEqual(requests.post.call_args.kwargs["json"]["menciones"], "a@icp.cl")
