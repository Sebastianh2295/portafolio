# Portafolio PMO — complemento de Excel

App web (HTML + JavaScript) que corre dentro de Excel en línea y usa `Portafolio.xlsx` como base de datos.
No necesita servidor, registro de aplicación ni aprobación de TI.

## Archivos
| Archivo | Para qué |
|---|---|
| `index.html` | La app. Fuera de Excel abre en modo demostración con datos de prueba. |
| `taskpane.html` | Panel de Excel: abre la ventana grande y hace de puente con el libro. |
| `commands.html` | Requerido por Office; no se toca. |
| `manifest.xml` | "Instalador" del botón **Portafolio PMO** en Excel. |
| `js/datos.js` | Única parte que lee y escribe Excel (Office.js). |
| `js/reglas.js` | Roles, frecuencias, estado de seguimiento, IDs, KPIs. |
| `js/app.js` | Pantallas y formularios. |
| `css/estilos.css` | Diseño (paleta FSFB aproximada). |

## Puesta en marcha
1. Cree el repositorio público `portafolio` en GitHub y suba todo el contenido de esta carpeta.
2. Settings → Pages → Deploy from a branch → `main` / `(root)` → Save.
3. Abra `https://TU-USUARIO.github.io/portafolio/` — debe verse la app en modo demostración.
4. En `manifest.xml` reemplace `TU-USUARIO` por su usuario de GitHub (buscar y reemplazar todo) y súbalo de nuevo.
5. Abra `Portafolio.xlsx` en Excel en línea → Inicio → Complementos → Más complementos → Mis complementos → Cargar mi complemento → elija `manifest.xml`.
6. Clic en **Portafolio PMO** (pestaña Inicio) → **Abrir Portafolio**.

Cada usuario repite el paso 5 una sola vez.
