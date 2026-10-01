# Convención de Dorcas 2026 · IEPI LAN-C

Proyecto independiente para inscripciones y hospedaje. No tiene relación con el portal del IBC.

| Página | Dirección | Para qué sirve |
|---|---|---|
| `index.html` | https://profesorrudy.github.io/dorcas2026/ | Formulario público de inscripción (nombre, apellido, edad, iglesia, teléfono opcional). |
| `admin.html` | https://profesorrudy.github.io/dorcas2026/admin.html | Panel con clave: lista de inscritas, hospedadores, asignación y tarjetas para imprimir. |

## Modo de prueba
Mientras `EVENTO_URL` (en `comun.js`) esté vacío, todo se guarda solo en el navegador
donde se usa y la clave del panel es `prueba`. Sirve para probar, no para recibir inscripciones reales.

## Ponerlo en marcha
1. Sigue las instrucciones al inicio de `apps-script.gs`: pégalo en un proyecto nuevo de
   https://script.google.com, agrega la propiedad `CLAVE_ADMIN`, ejecuta `probar` e implementa
   como aplicación web.
2. Pega la URL que termina en `/exec` en `EVENTO_URL`, dentro de `comun.js`.
3. Completa en `comun.js` las fechas, el lugar y el contacto del evento.

Los datos quedan en la planilla "Convención Dorcas 2026 - Inscripciones" de tu Google Drive.
