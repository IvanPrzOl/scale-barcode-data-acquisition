# Scale plus barcode data acquisition App
***
Programa para la adquisición de datos en las estaciones de pesado (báscula + escaner de código de barras) del grupo de fisiología de trigo en CIMMYT.

Escrito en python usando tkinter para construir la interfaz gráfica de usuario y pySerial para la comunicación de dispositivos.

Los datos se almacenan en archivos de excel usando el paquete xlwings

El programa consiste en tres módulos principales:
- GUI
- Serial gateway (código tomado de <http://www.hessmer.org/blog/2010/11/21/sending-data-from-arduino-to-ros/> )
- Excel Bridge