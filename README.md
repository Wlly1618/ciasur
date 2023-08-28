# Documentacion
## Librerias
> - xlrd<br>
  esta libreria es  para el trabajo con los .xlsx, osea versiones de 20016 en adelante
> - openpyxl<br>
  esta libreria es para el trabajo con los .xls, osea versiones de 2007 en adelante
> - csv<br>
  esta libreria es para crear los archivos .csv

## Funciones
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
> - Los parametros son enviados desde dos funciones diferentes, cada una  recibe los datos diferentes formas
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
#### read_and_write:
> Funcionamiento:
> - Primero se inicializa el  libro, luego se crean dos hojas  con las cuales  se  trabajara
> -  Luego se obtiene las fechas  para obtener las celdas a recorrer
> - Por ultimo  se obtiene los datos de  las diferentes  hojas y  se envian a write_data()
```python
        # para archivos .xlxs
        workbook = xlrd.open_workbook(route)
        worksheets_names = workbook.sheet_names()

        worksheet_fof2 = workbook.sheet_by_name(worksheets_names[0])
        worksheet_fof2_correct = workbook.sheet_by_name(worksheets_names[2])

        month: int = int(worksheet_fof2.cell_value(5, 5))
        year: int = int(worksheet_fof2.cell_value(5, 7))
        amo_days = calendar.monthrange(year=year, month=month)

        data_fo_f2 = get_data_fo_f2_xls(worksheet_fof2, amo_days[1])
        data_fo_f2_correct = get_data_fo_f2_xls(worksheet_fof2_correct, amo_days[1])

        write_data(year, month, data_fo_f2, data_fo_f2_correct, folder)
```
```python
        # para archivos .cl
        workbook = openpyxl.load_workbook(route)
        worksheet_fof2 = workbook[workbook.sheetnames[0]]
        worksheet_fof2_correct = workbook[workbook.sheetnames[2]]

        month: int = int(worksheet_fof2['F6'].value)
        year: int = int(worksheet_fof2['H6'].value)
        amo_days = calendar.monthrange(year=year, month=month)

        data_fo_f2 = get_data_fo_f2_xlsx(worksheet_fof2, amo_days[1])
        data_fo_f2_correct = get_data_fo_f2_xlsx(worksheet_fof2_correct, amo_days[1])

        write_data(year, month, data_fo_f2, data_fo_f2_correct, folder)
        workbook.close()
```

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

### main()
> La funcion principal se encarga de usar la funcion read_folder  para recuperar los datos
> Luego de controla las extensiones de los archivos para mandar con la funcion correspondiente
> Tambien pregunta seleccionar la carpeta donde se guardaran los .csv
```python
    list_routes = read_folder()
    route_folder = filedialog.askdirectory()
    print(route_folder)
    for route in list_routes:
        ext = os.path.splitext(route)[1]
        print(route)
        if '.xlsx' == ext:
            read_and_write_xlsx(route, route_folder)
        elif '.xls' == ext:
            read_and_write_xls(route, route_folder)
        else:
            print('No excel file')
```
