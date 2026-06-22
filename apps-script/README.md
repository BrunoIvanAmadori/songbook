# Apps Script deployment

1. Crea un proyecto de Google Apps Script y copia `Code.gs` y `appsscript.json`.
2. Despliega una `Web app`.
3. Elige `Execute as: User accessing the web app`.
4. Elige `Who has access: Anyone with Google account`.
5. Comparte la Google Sheet solo con las cuentas que deban editar.
6. Copia la URL `.../exec` desplegada.
7. Pega esa URL en `APPS_SCRIPT_URL_DEFAULT` dentro de `index.html`.
8. Publica tambien `apps-script-callback.html` junto a `index.html`.

La app lee canciones desde la Google Sheet publica con `gviz`. Para login, abre la Web App en un popup. Para guardar, envia el formulario a un iframe oculto si la sesion ya esta activa. Apps Script redirige ambos flujos a `apps-script-callback.html`, y esa pagina devuelve el resultado al cancionero con `postMessage`, evitando el bloqueo CORS de `fetch`.

Cada cambio en `Code.gs` requiere redeployar la Web App.
