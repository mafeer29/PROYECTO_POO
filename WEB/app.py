import pandas as pd
import plotly.express as px
from flask import Flask, render_template, request
from flask import send_file
from openpyxl.utils import get_column_letter
import io
import os
import numpy as np

# Configuración de Flask
app = Flask(__name__, static_url_path="", template_folder=".", static_folder='.')

# Cargar el archivo Excel y obtener la lista de departamentos
df_1 = pd.read_excel("../ANALIZADOR/REPORTE_DEPARTAMENTO.xlsx")

# Obtener los departamentos únicos
departamentos = df_1['DEPARTAMENTO'].dropna().unique()
# Cargar el archivo Excel y obtener la lista de departamentos

df_2 = pd.read_excel("../ANALIZADOR/REPORTE_PROVINCIA.xlsx")

# Obtener valores únicos de departamentos y provincias
departamentos_2 = df_2['DEPARTAMENTO'].unique()
df_3 = pd.read_excel("../ANALIZADOR/REPORTE_DISTRITO.xlsx")
# Obtener valores únicos de departamentos y provincias
departamentos_3 = df_3['DEPARTAMENTO'].unique()
df_4 = pd.read_excel("../ANALIZADOR/REPORTE_CENTRO_POBLADO.xlsx")

# Obtener valores únicos de departamentos y provincias
departamentos_4 = df_4['DEPARTAMENTO'].unique()
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
    fig3 = px.bar(poblacion_data, 
              x='Categoria', 
              y='Valor',
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
@app.route('/NIVEL_PROVINCIA.html', methods=['GET', 'POST'])
def nivel_provincia():
    departamento_seleccionado = None
    provincia_seleccionada = None
    provincias_2 = []

    # Si se recibe el formulario con el departamento seleccionado
    if request.method == 'POST':
        departamento_seleccionado = request.form['departamento']
        # Filtrar las provincias correspondientes al departamento
        provincias_2 = df_2[df_2['DEPARTAMENTO'] == departamento_seleccionado]['PROVINCIA'].unique()

    # Obtener todos los departamentos únicos
    departamentos_2 = df_2['DEPARTAMENTO'].unique()

    return render_template('NIVEL_PROVINCIA.html', 
                           departamentos_2=departamentos_2, 
                           departamento_seleccionado=departamento_seleccionado, 
                           provincias_2=provincias_2, 
                           provincia_seleccionada=provincia_seleccionada)

@app.route('/RESULTADO_PROVINCIA.html', methods=['GET'])
def resultado_provincia():
    # Obtener los datos del formulario
    departamento_seleccionado = request.args.get('departamento') 
    provincia_seleccionada = request.args.get('provincia')
    print(departamento_seleccionado, provincia_seleccionada)
    # Filtrar los datos para el departamento y la provincia seleccionados
    df_filtrado = df_2[
                       (df_2['PROVINCIA'] == provincia_seleccionada)]

    # Crear un archivo CSV en memoria
    output = io.BytesIO()
    df_filtrado.to_csv(output, index=False, encoding='utf-8')
    output.seek(0)  # Volver al inicio del archivo en memoria

    # Extraer las columnas que necesitas mostrar
    info = df_filtrado[['DEPARTAMENTO', 'PROVINCIA', 'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G',
                        'CALIFICACION', 'CANT_EB_TOTAL', 'POBLACION_TOTAL', 'ALCANCE_EB', 'POBLACION_CUBIERTA', 
                        'POBLACION_NO_CUBIERTA', 'EB_NECESARIAS', 'EB_FALTANTES', 'PROPUESTA_EB', "POBLACION_ACTIVA", 
                        "ALCANCE_EB_PEA", "EB_NECESARIAS_PEA", "PROPUESTAS_EB_PEA"]]
    print(info.head()) 
    propuesta = info["PROPUESTA_EB"].iloc[0]
    propuesta_pea = info["PROPUESTAS_EB_PEA"].iloc[0]

    # Extraer valores individuales de la fila única
    cant_eb_2g = df_filtrado['CANT_EB_2G'].iloc[0]
    cant_eb_3g = df_filtrado['CANT_EB_3G'].iloc[0]
    cant_eb_4g = df_filtrado['CANT_EB_4G'].iloc[0]
    cant_eb_5g = df_filtrado['CANT_EB_5G'].iloc[0]
    
    eb_necesarias = df_filtrado['EB_NECESARIAS'].iloc[0]
    eb_faltantes = df_filtrado['EB_FALTANTES'].iloc[0]
    eb_necesarias_pea = df_filtrado['EB_NECESARIAS_PEA'].iloc[0]
  
    
    poblacion_cubierta = df_filtrado['POBLACION_CUBIERTA'].iloc[0]
    poblacion_no_cubierta = df_filtrado['POBLACION_NO_CUBIERTA'].iloc[0]
    poblacion_total= df_filtrado["POBLACION_TOTAL"].iloc[0]
    poblacion_pea = df_filtrado["POBLACION_ACTIVA"].iloc[0]
    
 

    # Crear DataFrame para los gráficos
    eb_data = pd.DataFrame({
        'Tecnología': ["2G", "3G", "4G", "5G"],
        'Cantidad': [cant_eb_2g, cant_eb_3g, cant_eb_4g, cant_eb_5g]
    })
    comparacion_data = pd.DataFrame({
        'Tipo': ['EB Necesarias', 'EB Faltantes'],
        'Cantidad': [eb_necesarias, eb_faltantes]
    })

    comparacion_data_pea = pd.DataFrame({
        'Tipo': ['EB Necesarias Total', 'EB Necesarias PEA'],
        'Cantidad': [eb_necesarias, eb_necesarias_pea]
    })

    comparacion_poblacion_pea = pd.DataFrame({
        'Tipo': ['Población Total', 'Población PEA'],
        'Cantidad': [poblacion_total, poblacion_pea]
    })
    
    # Gráfico de barras para estaciones base por tecnología
    fig1 = px.bar(
        eb_data,
        x='Tecnología',
        y='Cantidad',
        title=f'Estaciones Base por Tecnología en {provincia_seleccionada}',
        labels={'Cantidad': 'Número de Estaciones'},
        color = "Tecnología",
        color_discrete_map={'2G': '#6379F2', '3G': '#F24C3D', '4G': '#04BF8A', '5G': '#9F63F2'}
    )
 
    # Gráfico circular para distribución de estaciones base por tecnología
    fig2 = px.pie(
        eb_data,
        names='Tecnología',
        values='Cantidad',
        title=f'Distribución de Estaciones Base por Tecnología en {provincia_seleccionada}',
        color='Tecnología',
        color_discrete_map={'2G': '#6379F2', '3G': '#F24C3D', '4G': '#04BF8A', '5G': '#9F63F2'}
    )
    
    # Gráfico de barras para comparación de población cubierta vs. no cubierta
    fig3 = px.bar(
        x=['Población Cubierta', 'Población No Cubierta'],
        y=[poblacion_cubierta, poblacion_no_cubierta],
        title=f'Comparación de Población Cubierta vs. No Cubierta en {provincia_seleccionada}',
        labels={'x': 'Categoría', 'y': 'Población'},
        color=['Población Cubierta', 'Población No Cubierta'],
        color_discrete_map={'Población Cubierta': '#6379F2', 'Población No Cubierta': '#F24C3D'}
    )
    
    # Gráfico de barras para estaciones base necesarias vs faltantes
    fig4 = px.bar(
        comparacion_data,
        x='Tipo',
        y='Cantidad',
        title=f'Estaciones Base Necesarias vs Faltantes en {provincia_seleccionada}',
        labels={'Cantidad': 'Número de Estaciones'},
        color='Tipo',
        color_discrete_map={'EB Necesarias': '#6379F2', 'EB Faltantes': '#F24C3D'}
    )

    # Gráfico de barras para poblacion PEA
    fig5 = px.bar(
        comparacion_poblacion_pea,
        x='Tipo',
        y='Cantidad',
        title=f'Población Total vs PEA en {provincia_seleccionada}',
        labels={'Cantidad': 'Número de Poblacion'},
        color='Tipo',
        color_discrete_map={'Población Total': '#6379F2', 'Población PEA': '#F24C3D'}
    )

    # Gráfico de barras para estaciones base necesarias vs faltantes
    fig6 = px.bar(
        comparacion_data_pea,
        x='Tipo',
        y='Cantidad',
        title=f'Estaciones Base Necesarias Total vs PEA en {provincia_seleccionada}',
        labels={'Cantidad': 'Número de Estaciones'},
        color='Tipo',
        color_discrete_map={'EB Necesarias': '#6379F2', 'EB Faltantes': '#F24C3D'}
    )
    # Convertir los gráficos a HTML para incrustarlos en la página
    grafico_html1 = fig1.to_html(full_html=False)
    grafico_html2 = fig2.to_html(full_html=False)
    grafico_html3 = fig3.to_html(full_html=False)
    grafico_html4 = fig4.to_html(full_html=False)
    grafico_html5 = fig5.to_html(full_html=False)
    grafico_html6 = fig6.to_html(full_html=False)

    return render_template('RESULTADO_PROVINCIA.html', 
                           propuesta=propuesta,
                           propuesta_pea = propuesta_pea,
                           info=info, 
                           departamento_seleccionado=departamento_seleccionado,
                           provincia_seleccionada=provincia_seleccionada,
                           grafico_html1=grafico_html1, 
                           grafico_html2=grafico_html2, 
                           grafico_html3=grafico_html3, 
                           grafico_html4=grafico_html4,
                           grafico_html5=grafico_html5,
                           grafico_html6=grafico_html6)

# Ruta para descargar el reporte de la provincia seleccionada
@app.route('/descargar_reporte_provincia/<provincia>')
def descargar_reporte_provincia(provincia):
    # Limpieza de los nombres
    
    provincia_limpia = provincia.strip()
    
    # Filtrar los datos
    df_filtrado = df_2[
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


@app.route('/NIVEL_DISTRITO.html', methods=['GET', 'POST'])
def nivel_distrito():
    # Obtener valores de selección previos
    departamento_seleccionado = request.form.get('departamento')
    provincia_seleccionada = request.form.get('provincia')
    
    # Obtener listas filtradas basadas en la selección previa
    provincias_3 = []
    distritos_3 = []

    if departamento_seleccionado:
        provincias_3 = df_3[df_3['DEPARTAMENTO'] == departamento_seleccionado]['PROVINCIA'].unique()

    if provincia_seleccionada:
        distritos_3 = df_3[(df_3['DEPARTAMENTO'] == departamento_seleccionado) & 
                           (df_3['PROVINCIA'] == provincia_seleccionada)]['DISTRITO'].unique()

    departamentos_3 = df_3['DEPARTAMENTO'].unique()

    return render_template('NIVEL_DISTRITO.html', 
                           departamentos_3=departamentos_3, 
                           departamento_seleccionado=departamento_seleccionado,
                           provincias_3=provincias_3, 
                           provincia_seleccionada=provincia_seleccionada,
                           distritos_3=distritos_3)


@app.route('/RESULTADO_DISTRITO.html', methods=["GET",'POST'])
def resultado_distrito():
    # Obtener los datos del formulario
    departamento_seleccionado = request.form.get('departamento')
    provincia_seleccionada = request.form.get('provincia')
    distrito_seleccionado = request.form.get('distrito')
    
    # Filtrar los datos para el centro poblado seleccionado
    df_filtrado = df_3[(df_3['DISTRITO'] == distrito_seleccionado) &
                       (df_3['PROVINCIA'] == provincia_seleccionada) &
                       (df_3['DEPARTAMENTO'] == departamento_seleccionado)]
    
    # Verifica si se obtuvo algún dato para evitar errores
    if df_filtrado.empty:
        return render_template('error.html', message="No se encontraron datos para el centro poblado seleccionado.")
    
    # Extraer las columnas necesarias para mostrar
    info = df_filtrado[['DEPARTAMENTO', 'PROVINCIA', 'DISTRITO',
                        'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G', 'CALIFICACION', 
                        'CANT_EB_TOTAL', 'POBLACION_TOTAL', 'ALCANCE_EB', 'POBLACION_CUBIERTA', 
                        'POBLACION_NO_CUBIERTA', 'EB_NECESARIAS', 'EB_FALTANTES', 'PROPUESTA_EB']]
    
    propuesta = info["PROPUESTA_EB"].iloc[0]
    
    # Extraer valores individuales de la fila única
    cant_eb_2g = df_filtrado['CANT_EB_2G'].iloc[0]
    cant_eb_3g = df_filtrado['CANT_EB_3G'].iloc[0]
    cant_eb_4g = df_filtrado['CANT_EB_4G'].iloc[0]
    cant_eb_5g = df_filtrado['CANT_EB_5G'].iloc[0]
    
    eb_necesarias = df_filtrado['EB_NECESARIAS'].iloc[0]
    eb_faltantes = df_filtrado['EB_FALTANTES'].iloc[0]
    
    poblacion_cubierta = df_filtrado['POBLACION_CUBIERTA'].iloc[0]
    poblacion_no_cubierta = df_filtrado['POBLACION_NO_CUBIERTA'].iloc[0]

    # Crear DataFrame para los gráficos
    eb_data = pd.DataFrame({
        'Tecnología': ["2G", "3G", "4G", "5G"],
        'Cantidad': [cant_eb_2g, cant_eb_3g, cant_eb_4g, cant_eb_5g]
    })
    
    comparacion_data = pd.DataFrame({
        'Tipo': ['EB Necesarias', 'EB Faltantes'],
        'Cantidad': [eb_necesarias, eb_faltantes]
    })
    
    # Gráfico de barras para estaciones base por tecnología
    fig1 = px.bar(
        eb_data,
        x='Tecnología',
        y='Cantidad',
        title=f'Estaciones Base por Tecnología en {distrito_seleccionado}',
        labels={'Cantidad': 'Número de Estaciones'},
        color = "Tecnología",
        color_discrete_map={'2G': '#6379F2', '3G': '#F24C3D', '4G': '#04BF8A', '5G': '#9F63F2'}
    )
 
    # Gráfico circular para distribución de estaciones base por tecnología
    fig2 = px.pie(
        eb_data,
        names='Tecnología',
        values='Cantidad',
        title=f'Distribución de Estaciones Base por Tecnología en {distrito_seleccionado}',
        color='Tecnología',
        color_discrete_map={'2G': '#6379F2', '3G': '#F24C3D', '4G': '#04BF8A', '5G': '#9F63F2'}
    )
    
    # Gráfico de barras para comparación de población cubierta vs. no cubierta
    fig3 = px.bar(
        x=['Población Cubierta', 'Población No Cubierta'],
        y=[poblacion_cubierta, poblacion_no_cubierta],
        title=f'Comparación de Población Cubierta vs. No Cubierta en {distrito_seleccionado}',
        labels={'x': 'Categoría', 'y': 'Población'},
        color=['Población Cubierta', 'Población No Cubierta'],
        color_discrete_map={'Población Cubierta': '#6379F2', 'Población No Cubierta': '#F24C3D'}
    )
    
    # Gráfico de barras para estaciones base necesarias vs faltantes
    fig4 = px.bar(
        comparacion_data,
        x='Tipo',
        y='Cantidad',
        title=f'Estaciones Base Necesarias vs Faltantes en {distrito_seleccionado}',
        labels={'Cantidad': 'Número de Estaciones'},
        color='Tipo',
        color_discrete_map={'EB Necesarias': '#6379F2', 'EB Faltantes': '#F24C3D'}
    )
    
    # Convertir los gráficos a HTML para incrustarlos en la página
    grafico_html1 = fig1.to_html(full_html=False)
    grafico_html2 = fig2.to_html(full_html=False)
    grafico_html3 = fig3.to_html(full_html=False)
    grafico_html4 = fig4.to_html(full_html=False)

    # Renderizar la plantilla con los datos y gráficos generados
    return render_template('RESULTADO_DISTRITO.html', 
                           propuesta=propuesta, 
                           info=info, 
                           departamento_seleccionado=departamento_seleccionado,
                           provincia_seleccionada=provincia_seleccionada,
                           distrito_seleccionado=distrito_seleccionado,
                           grafico_html1=grafico_html1, 
                           grafico_html2=grafico_html2, 
                           grafico_html3=grafico_html3, 
                           grafico_html4=grafico_html4)

    

# Ruta para descargar el reporte de la provincia seleccionada
@app.route('/descargar_reporte_distrito/<distrito>')
def descargar_reporte_distrito(distrito):
    # Limpieza de los nombres
    
    distrito_limpio = distrito.strip()
    
    # Filtrar los datos
    df_filtrado = df_2[
                       (df_2['DISTRITO'] == distrito_limpio)]
    
    # Crear el archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Reporte')
        
    output.seek(0)
    
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, 
                     download_name=f'{distrito_limpio}_reporte.xlsx')

@app.route('/NIVEL_CENTRO_POBLADO.html', methods=['GET', 'POST'])
def nivel_centro_poblado():
    # Obtener valores de selección previos
    departamento_seleccionado = request.form.get('departamento')
    provincia_seleccionada = request.form.get('provincia')
    distrito_seleccionado = request.form.get('distrito')

    # Filtrar los centros poblados si ya se seleccionaron los filtros de departamento, provincia y distrito
    centros_poblados = []
    if departamento_seleccionado and provincia_seleccionada and distrito_seleccionado:
        centros_poblados = df_4[(df_4['DEPARTAMENTO'] == departamento_seleccionado) & 
                                (df_4['PROVINCIA'] == provincia_seleccionada) & 
                                (df_4['DISTRITO'] == distrito_seleccionado)]['CENTRO_POBLADO'].unique()

    # Obtener los departamentos, provincias y distritos
    departamentos_4 = df_4['DEPARTAMENTO'].unique()
    provincias_4 = df_4[df_4['DEPARTAMENTO'] == departamento_seleccionado]['PROVINCIA'].unique() if departamento_seleccionado else []
    distritos_4 = df_4[(df_4['DEPARTAMENTO'] == departamento_seleccionado) & 
                       (df_4['PROVINCIA'] == provincia_seleccionada)]['DISTRITO'].unique() if provincia_seleccionada else []

    return render_template('NIVEL_CENTRO_POBLADO.html', 
                           departamentos_4=departamentos_4, 
                           departamento_seleccionado=departamento_seleccionado,
                           provincias_4=provincias_4, 
                           provincia_seleccionada=provincia_seleccionada,
                           distritos_4=distritos_4,
                           distrito_seleccionado=distrito_seleccionado,
                           centros_poblados=centros_poblados)

@app.route('/RESULTADO_CENTRO_POBLADO.html', methods=["GET", 'POST'])
def resultado_centro_poblado():
    # Obtener los datos del formulario
    departamento_seleccionado = request.form.get('departamento')
    provincia_seleccionada = request.form.get('provincia')
    distrito_seleccionado = request.form.get('distrito')
    centro_poblado_seleccionado = request.form.get('centro_poblado')
    
    # Filtrar los datos para el centro poblado seleccionado
    df_filtrado = df_4[(df_4['CENTRO_POBLADO'] == centro_poblado_seleccionado) &
                       (df_4['DISTRITO'] == distrito_seleccionado) &
                       (df_4['PROVINCIA'] == provincia_seleccionada) &
                       (df_4['DEPARTAMENTO'] == departamento_seleccionado)]
    
    # Verifica si se obtuvo algún dato para evitar errores
    if df_filtrado.empty:
        return render_template('error.html', message="No se encontraron datos para el centro poblado seleccionado.")
    
    # Extraer las columnas necesarias para mostrar
    info = df_filtrado[['DEPARTAMENTO', 'PROVINCIA', 'DISTRITO', 'CENTRO_POBLADO', 
                        'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G', 'CALIFICACION', 
                        'CANT_EB_TOTAL', 'POBLACION', 'ALCANCE_EB', 'POBLACION_CUBIERTA', 
                        'POBLACION_NO_CUBIERTA', 'EB_NECESARIAS', 'EB_FALTANTES', 'PROPUESTA_EB']]
    
    propuesta = info["PROPUESTA_EB"].iloc[0]
    
    # Extraer valores individuales de la fila única
    cant_eb_2g = df_filtrado['CANT_EB_2G'].iloc[0]
    cant_eb_3g = df_filtrado['CANT_EB_3G'].iloc[0]
    cant_eb_4g = df_filtrado['CANT_EB_4G'].iloc[0]
    cant_eb_5g = df_filtrado['CANT_EB_5G'].iloc[0]
    
    eb_necesarias = df_filtrado['EB_NECESARIAS'].iloc[0]
    eb_faltantes = df_filtrado['EB_FALTANTES'].iloc[0]
    
    poblacion_cubierta = df_filtrado['POBLACION_CUBIERTA'].iloc[0]
    poblacion_no_cubierta = df_filtrado['POBLACION_NO_CUBIERTA'].iloc[0]

    # Crear DataFrame para los gráficos
    eb_data = pd.DataFrame({
        'Tecnología': ["2G", "3G", "4G", "5G"],
        'Cantidad': [cant_eb_2g, cant_eb_3g, cant_eb_4g, cant_eb_5g]
    })
    
    comparacion_data = pd.DataFrame({
        'Tipo': ['EB Necesarias', 'EB Faltantes'],
        'Cantidad': [eb_necesarias, eb_faltantes]
    })
    
    # Gráfico de barras para estaciones base por tecnología
    fig1 = px.bar(
        eb_data,
        x='Tecnología',
        y='Cantidad',
        title=f'Estaciones Base por Tecnología en {centro_poblado_seleccionado}',
        labels={'Cantidad': 'Número de Estaciones'},
        color = "Tecnología",
        color_discrete_map={'2G': '#6379F2', '3G': '#F24C3D', '4G': '#04BF8A', '5G': '#9F63F2'}
    )
   
    # Gráfico circular para distribución de estaciones base por tecnología
    fig2 = px.pie(
        eb_data,
        names='Tecnología',
        values='Cantidad',
        title=f'Distribución de Estaciones Base por Tecnología en {centro_poblado_seleccionado}',
        color='Tecnología',
        color_discrete_map={'2G': '#6379F2', '3G': '##F24C3D', '4G': '##04BF8A', '5G': '#9F63F2'}
    )
    
    # Gráfico de barras para comparación de población cubierta vs. no cubierta
    fig3 = px.bar(
        x=['Población Cubierta', 'Población No Cubierta'],
        y=[poblacion_cubierta, poblacion_no_cubierta],
        title=f'Comparación de Población Cubierta vs. No Cubierta en {centro_poblado_seleccionado}',
        labels={'x': 'Categoría', 'y': 'Población'},
        color=['Población Cubierta', 'Población No Cubierta'],
        color_discrete_map={'Población Cubierta': '#6379F2', 'Población No Cubierta': '#F24C3D'}
    )
    
    # Gráfico de barras para estaciones base necesarias vs faltantes
    fig4 = px.bar(
        comparacion_data,
        x='Tipo',
        y='Cantidad',
        title=f'Estaciones Base Necesarias vs Faltantes en {centro_poblado_seleccionado}',
        labels={'Cantidad': 'Número de Estaciones'},
        color='Tipo',
        color_discrete_map={'EB Necesarias': '#6379F2', 'EB Faltantes': '#F24C3D'}
    )
    
    # Convertir los gráficos a HTML para incrustarlos en la página
    grafico_html1 = fig1.to_html(full_html=False)
    grafico_html2 = fig2.to_html(full_html=False)
    grafico_html3 = fig3.to_html(full_html=False)
    grafico_html4 = fig4.to_html(full_html=False)

    # Renderizar la plantilla con los datos y gráficos generados
    return render_template('RESULTADO_CENTRO_POBLADO.html', 
                           propuesta=propuesta, 
                           info=info, 
                           departamento_seleccionado=departamento_seleccionado,
                           provincia_seleccionada=provincia_seleccionada,
                           distrito_seleccionado=distrito_seleccionado,
                           centro_poblado_seleccionado=centro_poblado_seleccionado,
                           grafico_html1=grafico_html1, 
                           grafico_html2=grafico_html2, 
                           grafico_html3=grafico_html3, 
                           grafico_html4=grafico_html4)

# Ruta para descargar el reporte de la provincia seleccionada
@app.route('/descargar_reporte_centro_poblado/<centro_poblado>')
def descargar_reporte_centro_poblado(centro_poblado):
    # Limpieza de los nombres
    
    centro_poblado_limpio = centro_poblado.strip()
    
    # Filtrar los datos
    df_filtrado = df_4[
                       (df_4['CENTRO_POBLADO'] == centro_poblado_limpio)]
    
    # Crear el archivo Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, index=False, sheet_name='Reporte')
        
    output.seek(0)
    
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, 
                     download_name=f'{centro_poblado_limpio}_reporte.xlsx')



# Iniciar la aplicación
if __name__ == '__main__':
    app.run(debug=True, host= "0.0.0.0")

