import pandas as pd
import plotly.express as px
from flask import Flask, render_template, request
from flask import send_file
from openpyxl.utils import get_column_letter
import io

# Configuración de Flask
app = Flask(__name__, static_url_path="", template_folder=".", static_folder='.')

# Cargar el archivo Excel y obtener la lista de departamentos
df_1 = pd.read_excel(r"C:\Users\Maria Fernanda\Proyecto_POO\ANALIZADOR\REPORTE_DEPARTAMENTO.xlsx")
# Obtener los departamentos únicos
departamentos = df_1['DEPARTAMENTO'].dropna().unique()
# Cargar el archivo Excel y obtener la lista de departamentos

df_2 = pd.read_excel(r"C:\Users\Maria Fernanda\Proyecto_POO\ANALIZADOR\REPORTE_PROVINCIA.xlsx")

# Ruta principal
@app.route('/')
def index():
    return render_template('index.html')

# Ruta para NIVEL_DEPARTAMENTO
@app.route('/NIVEL_DEPARTAMENTO.html', methods=['GET', 'POST'])
def nivel_departamento():
    # Cargar el DataFrame
    # Pasar los departamentos al template
    return render_template('NIVEL_DEPARTAMENTO.html', departamentos=departamentos)

@app.route('/RESULTADO_DEPARTAMENTO.html', methods=['POST'])
def resultado_departamento():
    # Obtener el departamento o distrito seleccionado
    departamento_seleccionado = request.form['departamento']
    
    # Filtrar los datos para el departamento seleccionado
    df_filtrado = df_1[df_1['DEPARTAMENTO'] == departamento_seleccionado]
    
    # Crear un archivo CSV en memoria
    output = io.BytesIO()
    df_filtrado.to_csv(output, index=False, encoding='utf-8')
    output.seek(0)  # Volver al inicio del archivo en memoria

    # Extraer las columnas que necesitas mostrar
    info = df_filtrado[['DEPARTAMENTO',
                        'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G', 
                        'CALIFICACION', 'CANT_EB_TOTAL', 'POBLACION_TOTAL',  'ALCANCE_EB', "POBLACION_CUBIERTA", "POBLACION_NO_CUBIERTA", "EB_NECESARIAS", "EB_FALTANTES", 'PROPUESTA_EB']]
    propuesta = info["PROPUESTA_EB"].iloc[0]
    poblacion_data = pd.DataFrame({
        'Categoria': ['Población Cubierta', 'Población No Cubierta'],
        'Valor': [df_filtrado['POBLACION_CUBIERTA'].sum(), df_filtrado['POBLACION_NO_CUBIERTA'].sum()]
    })

    eb_data = pd.DataFrame({
    'Categoria': ["2G", "3G", "4G", "5G"],
    'Valor': [
        df_filtrado['CANT_EB_2G'].iloc[0],  # Se toma el valor directo
        df_filtrado['CANT_EB_3G'].iloc[0],  # Se toma el valor directo
        df_filtrado['CANT_EB_4G'].iloc[0],  # Se toma el valor directo
        df_filtrado['CANT_EB_5G'].iloc[0]   # Se toma el valor directo
    ]
    })

    fig1 = px.bar(df_filtrado, 
                 x='DEPARTAMENTO', 
                 y=['CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G'],
                 title=f'Estaciones Base por Tecnología en {departamento_seleccionado}',
                 labels={'value': 'Cantidad de Estaciones Base', 'variable': 'Tecnología'}, # Añadir etiquetas para mejorar la visualización
                 barmode='group')  # Agrupar las barras por distrito
    # Crear el gráfico circular con los valores correctos
    fig2 = px.pie(eb_data, 
                names='Categoria',  # Nombres de las categorías
                values='Valor',  # Valores de cada tecnología
                title=f'Distribución de Estaciones Base por Tecnología en {departamento_seleccionado}',
                color_discrete_map={'2G': '#6379F2', 
                                    '3G': '#F24C3D', 
                                    '4G': '#04BF8A', 
                                    '5G': '#9F63F2'})  # Colores personalizados

    # Crear el gráfico circular
    fig3 = px.pie(poblacion_data, 
                  names='Categoria', 
                  values='Valor',
                  title=f'Comparación de Población Cubierta vs. No Cubierta en {departamento_seleccionado}',
                  color='Categoria',
                  color_discrete_map={'Población Cubierta': '#6379F2', 'Población No Cubierta': '#F24C3D'})  # Colores personalizados
    
    fig4 = px.bar(info, 
              x='DEPARTAMENTO', 
              y=['EB_NECESARIAS', 'EB_FALTANTES'], 
              title=f'Comparación de Estaciones Base Necesarias vs Faltantes en {departamento_seleccionado}',
              labels={'value': 'Cantidad de Estaciones Base', 'variable': 'Tipo'},
              barmode='group')  # Agrupar las barras por tipo (necesarias y faltantes)


    
    # Convertir la figura de Plotly a HTML para incrustarla en la página
    grafico_html1 = fig1.to_html(full_html=False)
    grafico_html2 = fig2.to_html(full_html=False)
    grafico_html3 = fig3.to_html(full_html=False)
    grafico_html4 = fig4.to_html(full_html=False)

    return render_template('RESULTADO_DEPARTAMENTO.html', propuesta=propuesta, info=info, grafico_html1=grafico_html1, grafico_html2=grafico_html2, grafico_html3=grafico_html3, grafico_html4=grafico_html4)


@app.route('/descargar_reporte/<departamento>')
def descargar_reporte(departamento):
    # Limpieza del nombre del departamento
    departamento_limpio = departamento.replace("\n", "").replace("\r", "")
    
    # Filtrar y limpiar columnas del DataFrame
    df_filtrado = df_1[df_1['DEPARTAMENTO'] == departamento_limpio]
    df_filtrado.columns = df_filtrado.columns.str.replace("\n", " ").str.replace("\r", " ")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Reporte')
        
        # Ajustar el ancho de las columnas
        worksheet = writer.sheets['Reporte']
        for col_idx, column in enumerate(df_filtrado.columns, 1):
            max_length = max(df_filtrado[column].astype(str).apply(len).max(), len(column))
            worksheet.column_dimensions[get_column_letter(col_idx)].width = max_length + 2  # Añade margen

    output.seek(0)
    
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, 
                     download_name=f'{departamento_limpio}_reporte.xlsx')
# Otras rutas
@app.route('/NIVEL_NACIONAL')
def nivel_nacional():
    return render_template('NIVEL_NACIONAL.html')

@app.route('/NIVEL_PROVINCIA.html', methods=['POST'])
def nivel_provincia():
    departamento_seleccionado = request.form['departamento']
    
    # Filtrar provincias del departamento seleccionado
    provincias = df_2[df_2['DEPARTAMENTO'] == departamento_seleccionado]['PROVINCIA'].unique()
    
    return render_template('NIVEL_PROVINCIA.html', 
                           departamento=departamento_seleccionado, 
                           provincias=provincias)

# Ruta para mostrar resultados de la provincia seleccionada
@app.route('/RESULTADO_PROVINCIA.html', methods=['POST'])
def resultado_provincia():
    departamento_seleccionado = request.form['departamento']
    provincia_seleccionada = request.form['provincia']
    
    # Filtrar los datos para la provincia seleccionada
    df_filtrado = df_2[(df_2['DEPARTAMENTO'] == departamento_seleccionado) & 
                       (df_2['PROVINCIA'] == provincia_seleccionada)]
    
    # Extraer las columnas necesarias
    info = df_filtrado[['PROVINCIA', 'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G', 
                        'CALIFICACION', 'CANT_EB_TOTAL', 'POBLACION_TOTAL', 'ALCANCE_EB']]
    
    # Crear gráficos (similar a lo que ya tienes en departamento)
    fig1 = px.bar(df_filtrado, 
                 x='PROVINCIA', 
                 y=['CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G'],
                 title=f'Estaciones Base por Tecnología en {provincia_seleccionada}',
                 barmode='group')
    
    grafico_html1 = fig1.to_html(full_html=False)
    
    return render_template('RESULTADO_PROVINCIA.html', 
                           info=info, 
                           grafico_html1=grafico_html1)

# Ruta para descargar el reporte de la provincia seleccionada
@app.route('/descargar_reporte_provincia/<departamento>/<provincia>')
def descargar_reporte_provincia(departamento, provincia):
    # Limpieza de los nombres
    departamento_limpio = departamento.strip()
    provincia_limpia = provincia.strip()
    
    # Filtrar los datos
    df_filtrado = df_2[(df_2['DEPARTAMENTO'] == departamento_limpio) & 
                       (df_2['PROVINCIA'] == provincia_limpia)]
    
    # Crear el archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Reporte')
        
    output.seek(0)
    
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, 
                     download_name=f'{provincia_limpia}_reporte.xlsx')

@app.route('/NIVEL_DISTRITO')
def nivel_distrito():
    return render_template('NIVEL_DISTRITO.html')

@app.route('/NIVEL_CENTRO_POBLADO')
def nivel_centro_poblado():
    return render_template('NIVEL_CENTRO_POBLADO.html')

@app.route('/article-details')
def article_details():
    return render_template('article-details.html')

@app.route('/log-in')
def log_in():
    return render_template('log-in.html')

@app.route('/privacy-policy')
def privacy_policy():
    return render_template('privacy-policy.html')

@app.route('/sign-up')
def sign_up():
    return render_template('sign-up.html')

@app.route('/terms-conditions')
def terms_conditions():
    return render_template('terms-conditions.html')

@app.route('/ANALIZADOR')
def analizador():
    return render_template('ANALIZADOR.html')

# Iniciar la aplicación
if __name__ == '__main__':
    app.run(debug=True)