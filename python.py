import openpyxl

def crear_archivo_excel():
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Gastos"
    workbook.save("informe_gastos.xlsx")

def ingresar_gastos():
    gastos = []
    while True:
        fecha = input("Ingrese la fecha del gasto (o 'fin' para terminar): ")
        if fecha.lower() == 'fin':
            break
        descripcion = input("Ingrese la descripción del gasto: ")
        monto = float(input("Ingrese el monto del gasto: "))
        gastos.append((fecha, descripcion, monto))
    return gastos

def resumen_gastos(gastos):
    total_gastos = sum(monto for _, _, monto in gastos)
    max_gasto = max(gastos, key=lambda x: x[2])
    min_gasto = min(gastos, key=lambda x: x[2])

    print("Resumen de gastos:")
    print(f"Número total de gastos: {len(gastos)}")
    print(f"Gasto más caro: Fecha: {max_gasto[0]}, Descripción: {max_gasto[1]}, Monto: {max_gasto[2]}")
    print(f"Gasto más barato: Fecha: {min_gasto[0]}, Descripción: {min_gasto[1]}, Monto: {min_gasto[2]}")
    print(f"Monto total de gastos: {total_gastos}")

    return total_gastos, max_gasto, min_gasto

def guardar_en_excel(gastos):
    workbook = openpyxl.load_workbook("informe_gastos.xlsx")
    worksheet = workbook["Gastos"]
    
    for gasto in gastos:
        worksheet.append(gasto)
    
    workbook.save("informe_gastos.xlsx")

crear_archivo_excel()
print("Ingrese los detalles de sus gastos:")
lista_gastos = ingresar_gastos()
total, max_gasto, min_gasto = resumen_gastos(lista_gastos)
guardar_en_excel(lista_gastos)
print("Informe de gastos guardado en 'informe_gastos.xlsx'.")
