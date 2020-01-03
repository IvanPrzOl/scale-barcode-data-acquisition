# Scale plus barcode data collection App
***
Programa para la adquisición de datos en las estaciones de pesado (báscula + escaner de código de barras) del grupo de fisiología de trigo en CIMMYT.

El programa hace uso de los siguientes módulos:
- Tkinter, para la interfaz gráfica.
- pySerial, para la conexión de dispositivos.
- SerialDataGateway, interfaz encargada de recibir los datos del dispositivo y tratar la información como líneas individuales, código base creado por el [Dr. Rainer Hessmer](http://www.hessmer.org/blog/2010/11/21/sending-data-from-arduino-to-ros/).
- xlwings, interfaz para manipular archivos de excel con python.

Los datos se almacenan en archivos de excel usando un formato personalizado que se incluye como ejemplo en el repositorio.
