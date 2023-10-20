from styleframe import StyleFrame, Styler
import pandas as pd
import numpy as np
import msvcrt as m
from sys import exit
import xlsxwriter

def wait():
   m.getch()
   
print("Leyendo archivo \"ALMACEN 1.xls\"...\n")
try:
    df = pd.read_excel('ALMACEN 1.xlsx')
except:
    print("ERROR: No se ha podido leer el archivo \"ALMACEN 1.xlsx\". ¿Tiene ese nombre y lo tienes en la misma carpeta que este ejecutable?\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    wait()
    exit(1)
df.fillna(value = 0, inplace = True) 

print("Leyendo archivo \"ALMACEN 2.xls\"...\n")
try:
    df2 = pd.read_excel('ALMACEN 2.xlsx')
except:
    print("ERROR: No se ha podido leer el archivo \"ALMACEN 2.xlsx\". ¿Tiene ese nombre y lo tienes en la misma carpeta que este ejecutable?\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    wait()
    exit(1)
    
df2.fillna(value = 0, inplace = True) 

print("Procesando archivos...\n")

try:
    a1 = df["Artículo"].unique()
    a1 = a1[a1 != "Artículo"]
    a1 = [a for a in a1 if isinstance(a, str)]

    a2 = df2["Artículo"].unique()
    a2 = a2[a2 != "Artículo"]
    a2 = [a for a in a2 if isinstance(a, str)]

    Artículos = a1 + list(set(a2) - set(a1)) #combina las dos listas de artículos

    for index,a in enumerate(Artículos):
        Artículos[index] = {}
        Artículos[index]["Artículo"] = a
        Artículos[index]["Descripción"] = df.loc[df["Artículo"] == a, "Descripción"].unique()[0]
        
        c1 = set(df.loc[df["Artículo"] == a, "Color"].unique())
        c2 = set(df2.loc[df2["Artículo"] == a, "Color"].unique())
        
        colores = sorted(list(set(c1).union(set(c2))))
          
        Artículos[index]["Color"] = {}
        
        for c in colores:
            Artículos[index]["Color"][c] = {}
            if len(df2.loc[(df2["Artículo"] == a) & (df2["Color"] == c), "Desc. Color"].unique()) > 0:
                Artículos[index]["Color"][c]["Desc. Color"] = df2.loc[(df2["Artículo"] == a) & (df2["Color"] == c), "Desc. Color"].unique()[0]
            else:
                Artículos[index]["Color"][c]["Desc. Color"] = df.loc[(df["Artículo"] == a) & (df["Color"] == c), "Desc. Color"].unique()[0]

            Artículos[index]["Color"][c]["Stock 1"] = df.loc[(df["Artículo"] == a) & (df["Color"] == c) & (df["Tipo Fila"] == "STOCK") , ["S","M","L","XL","2XL","3XL","4XL","5XL"]].to_dict('records')
            if not Artículos[index]["Color"][c]["Stock 1"]:
                Artículos[index]["Color"][c]["Stock 1"] = {'S': 0, 'M': 0, 'L': 0, 'XL': 0, '2XL': 0, '3XL': 0, '4XL': 0, '5XL': 0}
            else:
                Artículos[index]["Color"][c]["Stock 1"] = Artículos[index]["Color"][c]["Stock 1"][0]
              
            Artículos[index]["Color"][c]["Stock 2"] = df2.loc[(df2["Artículo"] == a) & (df2["Color"] == c) & (df2["Tipo Fila"] == "STOCK") , ["S","M","L","XL","2XL","3XL","4XL","5XL"]].to_dict('records')
            if not Artículos[index]["Color"][c]["Stock 2"]:
                Artículos[index]["Color"][c]["Stock 2"] = {'S': 0, 'M': 0, 'L': 0, 'XL': 0, '2XL': 0, '3XL': 0, '4XL': 0, '5XL': 0}
            else:
                Artículos[index]["Color"][c]["Stock 2"] = Artículos[index]["Color"][c]["Stock 2"][0]
         
            Artículos[index]["Color"][c]["Pend. Servir"] = df.loc[(df["Artículo"] == a) & (df["Color"] == c) & (df["Tipo Fila"] == "PEND. SERVIR") , ["S","M","L","XL","2XL","3XL","4XL","5XL"]].to_dict('records')
            if not Artículos[index]["Color"][c]["Pend. Servir"]:
                Artículos[index]["Color"][c]["Pend. Servir"] = df2.loc[(df2["Artículo"] == a) & (df2["Color"] == c) & (df2["Tipo Fila"] == "PEND. SERVIR") , ["S","M","L","XL","2XL","3XL","4XL","5XL"]].to_dict('records')
            
            Artículos[index]["Color"][c]["Pend. Servir"] = Artículos[index]["Color"][c]["Pend. Servir"][0]
except:
    print("ERROR: Ha habido un error leyendo los datos del archivo. ¿Está en el mismo formato que le enviaste a Cristian?\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    exit(1)
    
print("Procesando movimientos del almacén 1 al 2...\n")

column_names = ["Movim.","Artículo","Descripción","Color","Desc. Color","S","M","L","XL","2XL","3XL","4XL","5XL"]
dfexcel = pd.DataFrame(columns = column_names)

for Artículo in Artículos:
    for color in Artículo["Color"]:
        
        pend_servir = Artículo["Color"][color]["Pend. Servir"]
        stock1 = Artículo["Color"][color]["Stock 1"]
        stock2 = Artículo["Color"][color]["Stock 2"]
        
        mov21 = {}
        for talla in pend_servir:
            virtual = Artículo["Color"][color]["Stock 2"][talla] - pend_servir[talla]
            if virtual < 0:
                if Artículo["Color"][color]["Stock 1"][talla] > 0:
                    total = virtual + Artículo["Color"][color]["Stock 1"][talla]
                    if total >= 0:
                        mov21[talla] = virtual * -1
                    else:
                        mov21[talla] = Artículo["Color"][color]["Stock 1"][talla]
                else:
                    mov21[talla] = 0
            else:
                mov21[talla] = 0
                
        row = [ "1 -> 2",
                Artículo["Artículo"],
                Artículo["Descripción"],
                color, 
                Artículo["Color"][color]["Desc. Color"], 
                mov21["S"],
                mov21["M"],
                mov21["L"],
                mov21["XL"],
                mov21["2XL"],
                mov21["3XL"],
                mov21["4XL"],
                mov21["5XL"]]
        if not all(mov21[v] == 0 for v in mov21):
            dfexcel_length = len(dfexcel)
            dfexcel.loc[dfexcel_length] = row
    
print("Procesando movimientos del almacén 2 al 1...\n")

for Artículo in Artículos:
    for color in Artículo["Color"]:
        
        pend_servir = Artículo["Color"][color]["Pend. Servir"]
        stock1 = Artículo["Color"][color]["Stock 1"]
        stock2 = Artículo["Color"][color]["Stock 2"]
        
        mov21 = {}
        for talla in pend_servir:
            virtual = Artículo["Color"][color]["Stock 2"][talla] - pend_servir[talla]
            
            if virtual >= 10:
                if Artículo["Color"][color]["Stock 1"][talla] < 10:
                    mov21[talla] = 10 - Artículo["Color"][color]["Stock 1"][talla]
                else:
                    mov21[talla] = 0
            else:
                mov21[talla] = 0
                
        row = [ "2 -> 1",
                Artículo["Artículo"],
                Artículo["Descripción"],
                color, 
                Artículo["Color"][color]["Desc. Color"], 
                mov21["S"],
                mov21["M"],
                mov21["L"],
                mov21["XL"],
                mov21["2XL"],
                mov21["3XL"],
                mov21["4XL"],
                mov21["5XL"]]
        if not all(mov21[v] == 0 for v in mov21):
            dfexcel_length = len(dfexcel)
            dfexcel.loc[dfexcel_length] = row

output_workbook = 'MOVIMIENTOS.xlsx'

writer = pd.ExcelWriter(output_workbook, engine='xlsxwriter')

dfexcel.to_excel(writer, index=False, sheet_name='Movimientos')

workbook  = writer.book
worksheet = writer.sheets['Movimientos']

# Cabeceras
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': False,
    'valign': 'top',
    'fg_color': '#000000',
    'font_color': '#ffffff'})
    
for col_num, value in enumerate(dfexcel.columns.values):
    worksheet.write(0, col_num, value, header_format)


size13 = workbook.add_format({'font_size': 13})

# Tamaños de columna
worksheet.set_column(0, 0, 7.2) #A
worksheet.set_column(1, 1, 7.3) #B
worksheet.set_column(2, 2, 37.7) #C
worksheet.set_column(3, 3, 6.3) #D
worksheet.set_column(4, 4, 23.3) #E
worksheet.set_column(5, 12, 3.45, size13) #F-M

filas_impares = workbook.add_format({
    'bg_color': '#DDDDDD'})

# Colorear filas impares
worksheet.conditional_format('A2:M1048576', {'type': 'formula',
                                            'criteria': '=AND((MOD(ROW(A2),2)=0), (A2<>""))',
                                            'format': filas_impares})

#Perfil de impresión
worksheet.set_margins(0.74, 0.74, 0, 0)
worksheet.print_area('A1:M1048576')
worksheet.fit_to_pages(1, 0)

print("Escribiendo archivo \"MOVIMIENTOS.xlsx\"...\n")
try:
    workbook.close()
    print("Proceso completado.")
except:
    print("ERROR: No se ha podido guardar el archivo. ¿Lo tienes abierto en Excel?\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    wait()
    exit(1)

''' after output_workbook =...
excel_writer = StyleFrame.ExcelWriter(output_workbook)

st = Styler(
    font='Calibri',
    font_size=11.0,
    horizontal_alignment='left',
    border_type='none',
    bg_color=None,
    wrap_text=False
)

st_bold = Styler(
    font='Calibri',
    font_size=11.0,
    bold=True,
    horizontal_alignment='left',
    border_type='none',
    font_color = "white",
    bg_color= "black"
)

st_highlight = Styler(
    font='Calibri',
    font_size=11.0,
    horizontal_alignment='left',
    border_type='none',
    bg_color='#DDDDDD'
)

st_highlight_center = Styler(
    font='Calibri',
    font_size=13.0,
    horizontal_alignment='center',
    border_type='none',
    bg_color='#DDDDDD'
)

st_highlight_center_odd = Styler(
    font='Calibri',
    font_size=13.0,
    horizontal_alignment='center',
    border_type='none',
    bg_color=None
)

st_highlight_center_black = Styler(
    font='Calibri',
    font_size=11.0,
    bold=True,
    horizontal_alignment='center',
    border_type='none',
    font_color = "white",
    bg_color= "black"
)

sf = StyleFrame(dfexcel, st)

sf.apply_headers_style(st_bold)
sf.apply_column_style(["S","M","L","XL","2XL","3XL","4XL","5XL"], st_highlight_center_black,style_header=True)
#sf.apply_column_style('Movim.', st, False, True, 10)
#sf.apply_column_style('Color', st, False, True, 8)
#sf.apply_column_style(["S","M","L","XL","2XL","3XL","4XL","5XL"], st, False, True, 4.86)

r = range(len(sf))
sf.apply_style_by_indexes(sf.index[[*r][::2]], st_highlight)
sf.apply_style_by_indexes(indexes_to_style=sf.index[[*r][::2]], cols_to_style=["S","M","L","XL","2XL","3XL","4XL","5XL"], styler_obj=st_highlight_center)
sf.apply_style_by_indexes(indexes_to_style=sf.index[[*r][1::2]], cols_to_style=["S","M","L","XL","2XL","3XL","4XL","5XL"], styler_obj=st_highlight_center_odd)


sf.set_column_width('Movim.', 10.0)
sf.set_column_width('Color', 7.56)
sf.set_column_width(["S","M","L","XL","2XL","3XL","4XL","5XL"], 4.86)

sf.A_FACTOR = 3
sf.P_FACTOR = 1.5

try:
    sf.to_excel(
        excel_writer=excel_writer, 
        best_fit=["Artículo","Descripción","Desc. Color"] #dfexcel.columns.tolist()
    )
except:
    print("¡No he encontrado ningún movimiento posible entre almacenes!\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    wait()
    exit(1)

print("Escribiendo archivo \"MOVIMIENTOS.xlsx\"...\n")
try:
    excel_writer.save()
    print("Proceso completado.")
except:
    print("ERROR: No se ha podido guardar el archivo. ¿Lo tienes abierto en Excel?\n")
    print("Pulsa cualquier tecla para cerrar esta pantalla...")
    wait()
    exit(1)'''