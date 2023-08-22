# Documentacion
## Librerias
- xlrd
  - esta libreria es  para el trabajo con los .xlsx, osea versiones de 20016 en adelante
- openpyxl
  - esta libreria es para el trabajo con los .xls, osea versiones de 2007 en adelante
- csv
  - esta libreria es para crear los archivos .csv

## Funciones
### read_folder():
> Funcionamiento: esta funcionalidad es la que permite seleccionar la carpeta en la cual se trabajara
```python
def read_folder():
    route_folder = filedialog.askdirectory()
    file_names = os.listdir(route_folder)
    file_paths = [os.path.join(route_folder, name_file) for name_file in file_names if
                  os.path.isfile(os.path.join(route_folder, name_file))]

    return file_paths
```
### write_data()
> Parametros:
> - year, month = fechas para luego realizar la creacion del documento
> -  data_fo_f2, data_fo_f2_correct  = datos que seran incluidos dentro del archivo .csv
```python
def write_data(year, month, data_fo_f2, data_fo_f2_correct):
```

> Variables:
> - data_time = es el rango de horas que existen, desde 0-23
> - header = es la cabecera de los elementos del .csv
> - file = es el nombre del archivo final
```python
data_time = [i for i in range(24)]
header = ['year', 'month', 'day', 'time', 'foF2', 'foF2_']
file = str(year) + " " + str(month) + ".csv"
```
> Funcionamiento:
> - Se crea el archivo .csv, y en este se comenzaran a ingresar los archivos que son recopilados
> - Primero sera ingresada la fecha de creacion, y luego mediante bucles se ingresaran los demas datos
> - Si todo es ingresado de forma segura se imprimira por pantalla un "file Created :)"
```python
with open(file, 'w', newline='') as archivo_csv:
    writer = csv.writer(archivo_csv)
    writer.writerow(header)
    i = 1
    for row_fo_f2, row_fo_f2_c in zip(data_fo_f2, data_fo_f2_correct):
        for fo_f2, fo_f2_c, time in zip(row_fo_f2, row_fo_f2_c, data_time):
            print("day: ", i, "\t time: ", time, "\tfo_f2: ", fo_f2, "\tfo_f2_correct: ", fo_f2_c)
            writer.writerow([year, month, i, time, fo_f2, fo_f2_c])
        i += 1

    archivo_csv.close()
```

### read_and_write_xlsx():
