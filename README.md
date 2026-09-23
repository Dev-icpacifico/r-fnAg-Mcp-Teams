# MCP Teams

Azure Function en Python 3.12 que expone herramientas MCP para enviar respuestas a conversaciones de Microsoft Teams y publicar mensajes mediante un Teams Workflow.

## Despliegue automático

El workflow `.github/workflows/deploy-mcp-teams.yml` ejecuta las pruebas en cada pull request hacia `main` y, después de pasar las pruebas, despliega en la Function App `mcp-teams` cada vez que se envían cambios a `main`. También se puede ejecutar manualmente desde GitHub Actions. El despliegue usa compilación remota para Flex Consumption.

Configura una vez lo siguiente:

1. En Azure Portal, abre la Function App **mcp-teams**. En **Settings > Configuration**, comprueba que **SCM Basic Auth Publishing Credentials** esté habilitado. Luego, en **Overview**, usa **Get publish profile**.
2. En el repositorio de GitHub, abre **Settings > Secrets and variables > Actions > New repository secret**. Crea `AZURE_FUNCTIONAPP_PUBLISH_PROFILE` y pega el contenido XML completo del archivo `.PublishSettings`. No agregues ese archivo al repositorio.
3. En la configuración de la Function App en Azure, comprueba que estén definidos `TEAMS_CALLBACK_URL`, `TEAMS_CALLBACK_KEY` y `TEAMS_CHANNEL_WEBHOOK_URL` con los valores del entorno correspondiente. GitHub Actions despliega código; no configura estos valores.
4. Haz commit y push a `main`. Revisa las tareas **Validate Python 3.12 Function** y **Deploy to Azure Flex Consumption** en la pestaña **Actions** del repositorio. Un commit sólo local no dispara GitHub Actions.

Si se renueva el publish profile, reemplaza también el secreto de GitHub.

## Conexión MCP desde VS Code

La configuración compartida está en `.vscode/mcp.json` y apunta al dominio de la Function App `mcp-teams`. Al iniciar el servidor, introduce la clave de sistema `mcp_extension` de **Functions > App keys > System keys**. La clave se solicita como entrada local y no está guardada en el repositorio.

## Pruebas locales

Con Python 3.12 y las dependencias de `requirements.txt` instaladas:

```bash
python -m unittest discover -s tests -v
```
