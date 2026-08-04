# Configuración de Grids V2

Haga doble clic en `Open Settings.bat` para abrir la configuración sin dejar una ventana de PowerShell visible.

La interfaz solo modifica:

- Idioma mostrado por PowerShell: inglés o español
- Diferencia y etiqueta UTC/GMT
- Configuración automática de color y nombre de la impresora
- Cuántas hojas compatibles formatea cada tipo de grilla: `X` para todas, o un número usando First/Last

Los tamaños de formato son fijos dentro del formateador de cada tipo de grilla y no se editan en esta interfaz V1 simplificada.

Al pulsar **Guardar Configuración**, la interfaz:

1. Crea un respaldo de `FormatGrids.ps1`.
2. Reemplaza solamente el bloque protegido de configuración.
3. Valida el script actualizado de PowerShell.
4. Lo activa únicamente si la validación es correcta.

`FormatGrids.ps1` no lee `settings.json` durante el formateo normal. El formateador continúa funcionando aunque se elimine toda la carpeta Settings.
