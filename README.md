# Generador de Casos de Prueba a Excel

Este es un script de Python que lee casos de prueba desde un archivo de texto (`.txt`), los procesa y los exporta a un archivo Excel (`.xlsx`) con formato profesional.

## Características

-   Detecta automáticamente si el formato del archivo de entrada es texto simple o texto delimitado por tabuladores.
-   Genera un archivo Excel con estilos: cabeceras en negrita, autoajuste de celdas y bordes.
-   Nombra el archivo de salida con una marca de tiempo para evitar sobreescribir resultados anteriores.

## Requisitos

-   Python 3.x
-   Librería `openpyxl`

## Instalación

1.  Clona este repositorio:
    ```bash
    git clone [https://github.com/jogonzal79/generador-casos-prueba.git](https://github.com/jogonzal79/generador-casos-prueba.git)
    ```
2.  Navega al directorio del proyecto:
    ```bash
    cd nombre-del-repo
    ```
3.  Instala las dependencias necesarias:
    ```bash
    pip install openpyxl
    ```

## Modo de Uso

1.  Coloca tu archivo de casos de prueba en formato `.txt` dentro de la carpeta `generador/input/`.
2.  Ejecuta el script desde la raíz del proyecto:
    ```bash
    python generador/test_case_parser.py
    ```
3.  El archivo Excel generado se guardará en la carpeta `generador/output/`.