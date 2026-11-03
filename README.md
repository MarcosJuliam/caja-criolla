# Caja Criolla

App web instalable para registrar ventas, gastos y merma del negocio. Cada registro se envía a una hoja de Google Sheets; sin internet se guarda en el equipo y se envía cuando vuelve la conexión.

- `index.html`, `sync.js`, `sw.js`, `manifest.webmanifest`, `icons/` — la app (GitHub Pages).
- `apps-script/Code.gs` — el conector que va en la hoja de Google (Extensiones → Apps Script). Cambia `CLAVE` antes de implementarlo.

La clave y los datos no están en este repositorio: viven en la hoja de Google y en cada equipo.
