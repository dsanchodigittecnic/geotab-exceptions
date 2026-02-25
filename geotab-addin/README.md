# Geotab Exceptions Add-In

Add-in web que replica la lógica de `app3.py` y descarga un Excel con 5 hojas:

- `00_Parametros`
- `01_Resumen_grupo`
- `02_Por_vehiculo`
- `03_Segmentos_1`
- `04_Excepciones`

## Despliegue

1. Publica la carpeta `geotab-addin` en un host HTTPS.
2. Ajusta en `manifest.json`:
   - `url`
   - `icon`
   - `supportEmail`
3. Registra el add-in en MyGeotab con ese `manifest.json`.

## Notas

- Usa el objeto `api` del propio add-in de Geotab, sin usuario/contraseña en código.
- La exportación de Excel se hace con SheetJS (`xlsx.full.min.js`).
