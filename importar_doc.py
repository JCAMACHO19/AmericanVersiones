import pandas as pd
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

# Define las rutas de los archivos
file_path_importar = r'C:\Users\jcamacho\Desktop\PRUEBA IMPORTE DE DOCUMENTOS\doc_importar.xls'
file_path_datos = r'C:\Users\jcamacho\Desktop\PRUEBA IMPORTE DE DOCUMENTOS\datos_facturas.xlsx'
file_path_cuenta_contable = r'C:\Users\jcamacho\Desktop\PRUEBA IMPORTE DE DOCUMENTOS\Cuenta_contable.xlsx'

# Lee los archivos Excel con los datos a rellenar
df_datos = pd.read_excel(file_path_datos)
df_cuenta_contable = pd.read_excel(file_path_cuenta_contable, dtype={'Centro Costos': str, 'IVA': str})

# Asegúrate de que los valores de NIT y Nit del Emisor: estén en formato de texto
df_datos['Nit del Emisor:'] = df_datos['Nit del Emisor:'].astype(str)
df_cuenta_contable['NIT'] = df_cuenta_contable['NIT'].astype(str)

# Reemplazar NaN en 'Fecha de Vencimiento:' con los valores de 'Fecha de Emisión:'
df_datos['Fecha de Vencimiento:'] = df_datos['Fecha de Vencimiento:'].fillna(df_datos['Fecha de Emisión:'])

# Abre el archivo .xls para leer y copiar
rb = open_workbook(file_path_importar, formatting_info=True)
wb = copy(rb)
sheet = wb.get_sheet(0)

# Fila destino en el archivo de importación
dest_row = 1  # Comienza desde la segunda fila en el archivo de importación

# Itera sobre todas las filas del archivo datos_facturas.xlsx
for i, row in df_datos.iterrows():
    nit_del_emisor = str(row['Nit del Emisor:'])
    numero_factura = str(row['Número de Factura:'])
    descripcion = str(row['descripcion'])
    fecha_emision = str(row['Fecha de Emisión:'])
    fecha_vencimiento = str(row['Fecha de Vencimiento:'])
    
    # Valores de las columnas para la verificación
    total_bruto_factura = row.get('Total Bruto Factura', 0)
    total_factura = row.get('Total factura (=)', 0)
    iva = row.get('IVA', 0)
    inc = row.get('INC', 0)
    
    # Definir los valores para la columna mCuenta
    cuenta_iva = df_cuenta_contable.loc[df_cuenta_contable['NIT'] == nit_del_emisor, 'IVA']
    
    mcuenta_values = {
        'IVA': cuenta_iva.values[0] if not cuenta_iva.empty else '24081001',
        'Total factura (=)': '23359505',
        'INC': '51159801'
    }
    
    # Contar cuántas columnas tienen valores mayores a cero
    columnas_mayores_cero = [
        (total_bruto_factura, 'mDebito', 'Total Bruto Factura'), 
        (iva, 'mDebito', 'IVA'), 
        (inc, 'mDebito', 'INC'), 
        (total_factura, 'mCredito', 'Total factura (=)')
    ]
    
    # Escribe los datos en el archivo de importación
    sheet.write(dest_row, 0, 'CP')  # dTipoDocumento
    sheet.write(dest_row, 1, None)  # dConsecutivo
    sheet.write(dest_row, 2, nit_del_emisor)  # dTercero
    sheet.write(dest_row, 3, f"{numero_factura} {descripcion}")  # dDescripcion
    sheet.write(dest_row, 4, fecha_emision)  # dFecha
    sheet.write(dest_row, 5, fecha_vencimiento)  # dVencimiento
    sheet.write(dest_row, 6, numero_factura)  # dReferencia
    sheet.write(dest_row, 7, None)  # mCuenta (vacío por ahora)

    # Incrementa la fila destino considerando las filas vacías
    dest_row += 1
    
    # Rellena las filas vacías
    for value, col_name, mcuenta_key in columnas_mayores_cero:
        if value > 0:
            # Deja una fila vacía
            sheet.write(dest_row, 0, None)  # dTipoDocumento
            sheet.write(dest_row, 1, None)  # dConsecutivo
            
            if col_name == 'mDebito':
                sheet.write(dest_row, 8, value)  # mDebito
            elif col_name == 'mCredito':
                sheet.write(dest_row, 9, value)  # mCredito
            
            # Rellena mCuenta según el tipo de dato
            if mcuenta_key and mcuenta_key != 'Total Bruto Factura':
                sheet.write(dest_row, 7, mcuenta_values[mcuenta_key])  # mCuenta
            elif mcuenta_key == 'Total Bruto Factura':
                # Busca el valor correspondiente en Cuenta_contable.xlsx
                cuenta_contable = df_cuenta_contable.loc[df_cuenta_contable['NIT'] == nit_del_emisor, 'Cuenta Contable Moda']
                if not cuenta_contable.empty:
                    sheet.write(dest_row, 7, str(cuenta_contable.values[0]))  # mCuenta (convertido a cadena)
            
            # Rellenar mDescripcion, mNit, mBase, mCentroC, mSegmento solo en las filas correspondientes
            sheet.write(dest_row, 10, f"{numero_factura} {descripcion}")  # mDescripcion
            sheet.write(dest_row, 11, nit_del_emisor)  # mNit
            
            if mcuenta_key == 'IVA':
                mBase_value = round((value * 100) / 19, 2)
                sheet.write(dest_row, 12, mBase_value)  # mBase
            else:
                sheet.write(dest_row, 12, None)  # mBase
                
            # Rellenar mCentroC basado en el valor de 'Centro Costos' como texto
            centro_costos = df_cuenta_contable.loc[df_cuenta_contable['NIT'] == nit_del_emisor, 'Centro Costos']
            if not centro_costos.empty:
                sheet.write(dest_row, 13, str(centro_costos.values[0]))  # mCentroC (convertido a cadena)
            else:
                sheet.write(dest_row, 13, None)  # mCentroC
            
            sheet.write(dest_row, 14, None)  # mSegmento
            
            dest_row += 1

# Guarda el archivo modificado
wb.save(file_path_importar)

print(f"El archivo se ha actualizado correctamente en {file_path_importar}")
