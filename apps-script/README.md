# Apps Script deployment

1. Creá un proyecto de Google Apps Script y copiá `Code.gs` y `appsscript.json`.
2. Desplegá una `Web app`.
3. Elegí `Execute as: User accessing the web app`.
4. Elegí `Who has access: Anyone with Google account`.
5. Compartí la Google Sheet solo con las cuentas que deban editar.
6. Copiá la URL `.../exec` desplegada.
7. Abrí la app con `?api=URL_DE_TU_WEB_APP` o guardá esa URL en `localStorage.songbook.apiUrl`.

La app seguirá pudiendo leer desde la Sheet pública si la Web App falla, pero el guardado y la verificación de permisos dependen de esta Web App.
