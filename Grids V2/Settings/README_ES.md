# Configuracion del Formateador de Grillas V2

1. Haga doble clic en **Open Settings.bat**.
2. Cambie el idioma, la diferencia UTC/GMT, el nombre mostrado, la configuracion de color o la seleccion de pestanas.
3. Use **X** para procesar todas las pestanas elegibles, o escriba un numero y elija First/Last.
4. Presione **Save Settings**.

Al guardar, solo se actualiza el bloque protegido dentro de `../FormatGrids.ps1`. El formateador no lee `settings.json` al procesar grillas. Si se elimina toda la carpeta Settings, el script activo continua funcionando.

`printer_color_set.txt` solo guarda el ultimo dia en que se configuro el color. Los paquetes de operador sin Settings usan automaticamente un archivo de estado local de Windows.
