from openpyxl import load_workbook

archivo_excel = "../Include/asistencia.xlsx"  # Cargar el archivo de Excel
libro = load_workbook(archivo_excel)
hoja = libro["Asistencia"]        # Obtener la hoja "Asistencia"

ultima_fila = hoja.max_row  # Obtener la última fila con datos

nombre = input("Ingresa tu nombre: ")  # Solicitar datos al usuario
fecha = input("Ingresa la fecha (yyyy-mm-dd): ")
hora = input("Ingresa la hora de entrada (hh:mm): ")

nueva_fila = ultima_fila + 1  # Agregar una nueva fila después de la última fila con datos

hoja.cell(row=nueva_fila, column=1).value = nombre   # Escribir los datos en la nueva fila
hoja.cell(row=nueva_fila, column=2).value = fecha
hoja.cell(row=nueva_fila, column=3).value = hora

libro.save(archivo_excel)    # Guardar los cambios en el archivo

print(f"Datos guardados exitosamente en la fila {nueva_fila}.")
