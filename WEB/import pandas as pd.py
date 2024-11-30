import pandas as pd

# Cargar el archivo Excel
df_2 = pd.read_excel(r"C:\Users\Maria Fernanda\Proyecto_POO\ANALIZADOR\REPORTE_PROVINCIA.xlsx")
print(df_2[['DEPARTAMENTO', 'PROVINCIA']].dropna().head())