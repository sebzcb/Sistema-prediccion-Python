import xlrd
import pandas as pd
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import SimpleExpSmoothing
import warnings
warnings.filterwarnings("ignore")


def menu():
    while True:
        mostrarTitulo("MENU DE PREDICCIONES")
        print("1. Predicción de ventas para 2023")
        print("2. Salir")
        print("")
        opcion = input("Elige una opción (1/2): ")
        mostrarTitulo("SELECCIONE CATEGORIA")
        if opcion == "1":
            mostrarCategorias()
            numCategoria = int(input("Introduce el numero de la categoria: "))
            categoria = categoriaSeleccionada(numCategoria)
            if(categoria == ""):
                print("No existe la opcion " + str(numCategoria))
                break
            mostrarTitulo("SELECCIONE MES")
            mostrarMes()
            numMes = int(input("Introduce el mes: "))
            mes = seleccionarMes(numMes)
            if(mes ==""):
                print("No existe la opcion " + str(mes))
                break

            mostrarTitulo("RESULTADO DE LA PREDICCION")
            print(" CATEGORIA : ",categoria)
            print(" MES : ",mes)
            print(" N° DE PRODUCTOS A VENDER PARA ",mes," DE 2023: ",predecirVentaMes(categoria,mes) ," unidades")
            print("")

        elif opcion == "2":
            print("¡Hasta luego!")
            break
        else:
            print("Opción no válida. Elige 1, o 2.")
def mostrarTitulo(titulo):
    print("")
    print("==================================================",titulo,"==================================================")
    print("")
def mostrartTituloResultado(mes,categoria):
    print("")
    print("====================  ================================================")
    print("")
def categoriaSeleccionada(num):
    seleccion = ""
    if(num == 1):
        seleccion = "Accesorios"
    elif (num == 2):
        seleccion = "Juguete"
    elif (num == 3):
        seleccion = "Casa"
    elif (num == 4):
        seleccion = "Electronicos"
    elif (num == 5):
        seleccion = "Deportes"
    elif (num == 6):
        seleccion = "Papeleria"
    elif (num == 7):
        seleccion = "Vestuario"
    return seleccion    

def seleccionarMes(num):
    seleccion = ""
    if(num == 1):
        seleccion = "ene"
    elif (num == 2):
        seleccion = "feb"
    elif (num == 3):
        seleccion = "mar"
    elif (num == 4):
        seleccion = "abr"
    elif (num == 5):
        seleccion = "may"
    elif (num == 6):
        seleccion = "jun"
    elif (num == 7):
        seleccion = "jul"
    elif (num == 8):
        seleccion = "ago"
    elif (num == 9):
        seleccion = "sept"
    elif (num == 10):
        seleccion = "oct"
    elif (num == 11):
        seleccion = "nov"
    elif (num == 12):
        seleccion = "dic"
    return seleccion    
def mostrarCategorias():
    print("")
    print("1.Accesorios", end=" ")
    print("2.Juguete", end=" ")
    print("3.Casa", end=" ")
    print("4.Electronicos", end=" ")
    print("5.Deportes", end=" ")
    print("6.Papeleria", end=" ")
    print("7.Vestuario")
    print("")
def mostrarMes():
    print("")
    print("1.Enero", end=" ")
    print("2.Febrero", end=" ")
    print("3.Marzo", end=" ")
    print("4.Abril", end=" ")
    print("5.Mayo", end=" ")
    print("6.Junio", end=" ")
    print("7.Julio", end=" ")
    print("8.Agosto", end=" ")
    print("9.Septiembre", end=" ")
    print("10.Octubre", end=" ")
    print("11.Noviembre", end=" ")
    print("12.Diciembre")
    print("")


def crearDiccionarioDesdeHojaExcel(hojaExcel):
    diccionario = {}
    # Obtiene las filas y columnas
    filas = hojaExcel.nrows
    columnas = hojaExcel.ncols
    # Obtiene los nombres de los accesorios de la primera fila
    nombres_accesorios = hojaExcel.row_values(0)
    # Itera a través de las filas de datos
    for fila in range(1, filas):
        datos_fila = hojaExcel.row_values(fila)
        accesorio = datos_fila[0]  # El nombre del accesorio está en la primera columna

        if accesorio not in diccionario:
            diccionario[accesorio] = {}

        # Itera a través de las columnas (años) y asigna la cantidad vendida
        for columna in range(1, columnas):
            anio = nombres_accesorios[columna]
            cantidad_vendida = datos_fila[columna]
            diccionario[accesorio][anio] = cantidad_vendida

    return diccionario

def predecirVentaMes(categoriaSolicitada, mesSolicitado):
    anios = [2017, 2018, 2019, 2020, 2021, 2022]

    # Crear una lista de cantidad vendida para la categoría y mes solicitados
    cantidad_vendida = [datos_2017[categoriaSolicitada][mesSolicitado],
                    datos_2018[categoriaSolicitada][mesSolicitado],
                    datos_2019[categoriaSolicitada][mesSolicitado],
                    datos_2020[categoriaSolicitada][mesSolicitado],
                    datos_2021[categoriaSolicitada][mesSolicitado],
                    datos_2022[categoriaSolicitada][mesSolicitado]]
    if categoriaSolicitada == "todasLasCategorias":
        cantidad_vendida = [datos_2017[categoriaSolicitada][mesSolicitado],
                    datos_2018[categoriaSolicitada][mesSolicitado],
                    datos_2019[categoriaSolicitada][mesSolicitado],
                    datos_2020[categoriaSolicitada][mesSolicitado],
                    datos_2021[categoriaSolicitada][mesSolicitado],
                    datos_2022[categoriaSolicitada][mesSolicitado]]
    # Crear un DataFrame a partir de los datos
    data = {
        'Año': anios,
        'Cantidad_Vendida': cantidad_vendida
    }
    df = pd.DataFrame(data)

    # Establece 'Año' como el índice
    df.set_index('Año', inplace=True)

    # Visualiza los datos
    plt.plot(df)
    plt.xlabel('Año')
    plt.ylabel('Cantidad Vendida')
    plt.title(f'Ventas de {categoriaSolicitada} en {mesSolicitado}')
    #plt.show()
    # Ajusta un modelo de Suavizamiento Exponencial Simple (SES) a los datos
    model = SimpleExpSmoothing(df['Cantidad_Vendida'])
    model_fit = model.fit()
    # Realiza la predicción para enero de 2023
    prediction = model_fit.forecast(steps=1)
    # Suma la predicción a la última observación para obtener la predicción original
    # Obtiene la última observación directamente
    last_observation = df['Cantidad_Vendida'].iloc[-1]
    # Suma la predicción a la última observación para obtener la predicción original
    predicted_value = last_observation + prediction
    return int(predicted_value.values[0])

# Abre el archivo XLS
archivo_xls = xlrd.open_workbook('base de datos.xls')
    # Selecciona una hoja específica (opcional)
hoja2017 = archivo_xls.sheet_by_name('2017')
hoja2018 = archivo_xls.sheet_by_name('2018')
hoja2019 = archivo_xls.sheet_by_name('2019')
hoja2020 = archivo_xls.sheet_by_name('2020')
hoja2021 = archivo_xls.sheet_by_name('2021')
hoja2022 = archivo_xls.sheet_by_name('2022')

    # Crea el diccionario de datos
datos_2017 = crearDiccionarioDesdeHojaExcel(hoja2017)
datos_2018 = crearDiccionarioDesdeHojaExcel(hoja2018)
datos_2019 = crearDiccionarioDesdeHojaExcel(hoja2019)
datos_2020 = crearDiccionarioDesdeHojaExcel(hoja2020)
datos_2021 = crearDiccionarioDesdeHojaExcel(hoja2021)
datos_2022 = crearDiccionarioDesdeHojaExcel(hoja2022)

menu()

