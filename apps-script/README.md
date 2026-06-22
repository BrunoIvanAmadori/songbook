# Apps Script deployment

1. Crea un proyecto de Google Apps Script y copia `Code.gs` y `appsscript.json`.
2. Despliega una `Web app`.
3. Elige `Execute as: User accessing the web app`.
4. Elige `Who has access: Anyone with Google account`.
5. Comparte la Google Sheet solo con las cuentas que deban editar.
6. Copia la URL `.../exec` desplegada.
7. Pega esa URL en `APPS_SCRIPT_URL_DEFAULT` dentro de `index.html`.

La app lee canciones desde la Google Sheet publica con `gviz`. Para login y guardado, abre la Web App en un popup y recibe el resultado con `postMessage`, evitando el bloqueo CORS de `fetch`.
