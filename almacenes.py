from styleframe import StyleFrame, Styler
import pandas as pd

def insert_row(row_num, orig_df, row_to_add):
    row_num= min(max(0, row_num), len(orig_df))
    df_part_1 = orig_df.loc[0:row_num]
    df_part_2 = orig_df.loc[row_num+1:]
    df_final = df_part_1.append(row_to_add, ignore_index = True)
    df_final = df_final.append(df_part_2, ignore_index = True)
    return df_final

df = pd.read_excel('ALMACEN 1.xlsx')
df2 = pd.read_excel('ALMACEN 2.xlsx')
output_workbook = 'ALMACENES.xlsx'

tallas = [['S','S'], ['M','M'], ['L','L'], ['XL','XL'], ['_12','2XL'], ['_13','3XL'], ['_14','4XL'], ['_15','5XL']]

addedrows = 0

for row in df.itertuples(index = True):
    rowindex = getattr(row, "Index")
    if (getattr(row, "_6") == "STOCK VIRTUAL"):
        df = insert_row(rowindex+addedrows, df, df2.loc[rowindex-1])
        addedrows = addedrows + 1
        
        df.loc[rowindex+addedrows, 'Tipo Fila'] = 'MOV. ALMACEN'
        df.loc[rowindex+addedrows, ['Total','S','M','L','XL','2XL','3XL','4XL','5XL']] = 0
    
        df = insert_row(rowindex+addedrows, df, df2.loc[rowindex-1])
        addedrows = addedrows + 1
        
        df.loc[rowindex+addedrows, 'Tipo Fila'] = 'A CORTAR'
        df.loc[rowindex+addedrows, ['Total','S','M','L','XL','2XL','3XL','4XL','5XL']] = 0
        
        total_mov = 0
        total_cut = 0
        for talla, talla_name in tallas:
            need = getattr(row, talla)
            stock = df2.loc[rowindex-1, talla_name]
            
            if need < 0 and stock >= need * -1:
                df.loc[rowindex+addedrows-1, talla_name] = need*-1
                total_mov = total_mov + need*-1
            elif need < 0 and stock > 0:
                df.loc[rowindex+addedrows-1, talla_name] = stock
                total_mov = total_mov + stock
            
            if need+stock < 0:
                df.loc[rowindex+addedrows, talla_name] = (need+stock)*-1
                total_cut = total_cut + (need+stock)*-1
        
        df.loc[rowindex+addedrows-1, 'Total'] = total_mov
        df.loc[rowindex+addedrows, 'Total'] = total_cut

df = df[df.columns.drop(list(df.filter(regex='Unnamed')))]

excel_writer = StyleFrame.ExcelWriter(output_workbook)

st = Styler(
    font='Calibri',
    font_size=11.0,
    horizontal_alignment='left',
    border_type='none',
    bg_color=None
)

st_bold = Styler(
    font='Calibri',
    font_size=11.0,
    bold=True,
    horizontal_alignment='left',
    border_type='none',
    bg_color=None
)

st_highlight = Styler(
    font='Calibri',
    font_size=11.0,
    horizontal_alignment='left',
    border_type='none',
    bg_color='#fffda1'
)

sf = StyleFrame(df, st)

sf.apply_headers_style(st_bold)
sf.apply_column_style('Tipo Fila', st_bold)
sf.apply_style_by_indexes(sf[sf['Tipo Fila'].str.contains('MOV. ALMACEN')], st_highlight)
sf.apply_style_by_indexes(sf[sf['Tipo Fila'].str.contains('A CORTAR')], st_highlight)


sf.A_FACTOR = 3
sf.P_FACTOR = 1
sf.to_excel(
    excel_writer=excel_writer, 
    best_fit=df.columns.tolist()
)
excel_writer.save()