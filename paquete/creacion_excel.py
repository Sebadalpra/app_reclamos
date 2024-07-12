import pandas as pd
from tkinter import messagebox

def crear_nuevo_excel():
    try:
        # Cargar los archivos Excel
        df1 = pd.read_excel('./excels/produccion.xls', engine='xlrd')
        df2 = pd.read_excel('./excels/db_fuente.xlsx')

        # Seleccionar solo la columna "CORREOS" junto con "CIA" para el merge
        df2 = df2[['CIA', 'CORREOS']]

        # Fusionar los dataframes en la columna "Cia ID"/"CIA"
        merged_df = pd.merge(df1, df2, left_on='Cia ID', right_on='CIA', how='left')

        # Eliminar la columna "CIA" del dataframe resultante
        merged_df = merged_df.drop(columns=['CIA'])

        # Guardar el dataframe fusionado en un nuevo archivo Excel
        merged_df.to_excel('cias_sin_presentar.xlsx', index=False)
        print("Excel generado con exito")
    except Exception:
        print(f"Error en la union de los excels: {Exception}")
        messagebox.showerror("Error", f"Error en la union de los excels 'db_fuente' y 'produccion'")
