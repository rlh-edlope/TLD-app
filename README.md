# TLD-app

<p align="center">
  <img src="GUI/splash.png" alt="logo" width="500">
</p>

![Version](https://img.shields.io/badge/version-V2.5.1-blue)
[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20Windows-lightgrey.svg)

Programa para trabajar con curvas de brillo

Esta herramienta es una simplificación para visualizar y trabajar con curvas de brillo provenientes de equipos termoluminiscentes Harshaw 3500.

Para utilizar esta herramienta se necesita que el equipo esté configurado para exportar los datos con el siguiente formato de exportación de la curvas. 

```
INSTRUMENT_TYPE
GROUP_ID
DOSIMETER_ID "\"%-16.16s\", "
READING_DATE "%Y%m%d, "
READING_TIME "%H%M%S, "
TTP_NUMBER
READING_FUNCTION
ELEMENT_FLAGS
RCF_VALUE "%6.3f, " 1
ECC_VALUE "%6.3f, " 1
RESPONSE "%9.3f, " 1
RESPONSE_UNIT "\"%s\"\n" 
HEAT_CURVE "%d\n" 1 1-200
GLOW_CURVE "%ld\n" 1 1-199
GLOW_CURVE "%ld" 1 200-200
```
## Funciones principales

### Visualización
Al abrir un archivo se pueden visualizar los datos de cada una de las curvas ya sea individualmente o múltiples curvas a la vez. 

### Normalización
Las curvas de brillo se pueden autonormalizar o normalizar respecto al máximo de alguna otra curva de brillo. 

### Resta de fondo

El programa permite hacer la resta del fondo exponencial que muchas veces se mete en la curva de brillo. Para realizar esta resta se puede utilizar el menú de análisis o el atajo `Ctrl+R` para hacer el procedimiento.

Este método utiliza 4 canales que se definen donde se define la altura del fondo y otros dos donde se muestra la sección exponencial. Ambos canales se pueden modificar en la configuración. 

Al restar un fondo se puede presionar el botón de **Guardar Curva** para mantener esta curva.

### Exportación de curvas de brillo

Las curvas de brillo se pueden exportar en un archivo Excel para ser utilizadas en este formato. En este archivo se pueden guardar las curvas de brillo sin el fondo que se resta en la función **Resta de Fondo**.  
