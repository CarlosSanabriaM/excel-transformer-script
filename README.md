# Excel transformer script

![Transformation example](docs/transformation-example.png)

The `excel-transformer-script` project is a Python-based tool designed to modify Excel files through a terminal-based user interface.

Utilizing the `openpyxl` library for Excel file manipulations, the script performs specific data transformations while preserving cell formatting and colors. It adjusts integer and percentage pairs in the spreadsheet, setting values below 100 to zero and marking these changes with red text.

## Pasos para ejecutar

1. Mover la Excel a modificar a la carpeta del script
2. Editar el script, cambiando los valores de las siguientes 3 variables
	- `input_excel_file`: Nombre del fichero excel a modificar (ejemplo: `input.xlsx`)
	- `output_excel_file`: Nombre del fichero excel de salida (ejemplo: `output.xlsx`)
	- `start_column_name` = Nombre de la primera columna a partir de la cual se deben hacer las transformaciones (ejemplo: `Eventos GFP+ (CD45+)`)
3. Abrir una terminal en la carpeta del script y ejecutar los siguientes comandos

```shell
pipenv shell
pip install openpyxl
python script.py input.xlsx
```
4. Abrir la Excel de salida y verificar que todo está OK