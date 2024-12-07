import pytest
import pandas as pd
import openpyxl
from CODIGO_SEBAS import ActualizarDatos, GenerarGraficas, PropuestasSoluciones, Poblacion

@pytest.fixture
def setup_data():
    # Cargar datos de prueba 
    archivo_excel = "COBERTURA MOVIL.xlsx"  # Asegúrate de que este archivo exista
    archivo_poblacion = "CCPP_INEI.xlsx"  # Asegúrate de que este archivo exista
    
    return ActualizarDatos(archivo_excel),  PropuestasSoluciones(archivo_excel, archivo_poblacion)
def test_cuantificar_cobertura(setup_data): 
    actualizar_datos, _ = setup_data
    atributos_cuantificados = actualizar_datos.cuantificar_cobertura()
    assert 'CALIFICACION' in atributos_cuantificados.columns  # Verifica que la columna 'calificacion' exista
    assert not atributos_cuantificados.empty  # Verifica que el DataFrame no esté vacío

def test_generar_reporte(setup_data):
    actualizar_datos, _ = setup_data
    actualizar_datos.cuantificar_cobertura()  # Asegúrate de cuantificar primero
    actualizar_datos.generar_reporte()
    assert pd.read_excel("COBERTURA_MOVIL_CUANTIFICADA.xlsx").shape[0] > 0  # Verifica que el archivo no esté vacío

def test_generar_graficas(setup_data):
    actualizar_datos, _ = setup_data
    actualizar_datos.cuantificar_cobertura()  # Asegúrate de cuantificar primero
    graficador = GenerarGraficas(actualizar_datos)
    try:
        graficador.generar_histograma_calificacion()  # Prueba que el histograma se genere sin errores
    except Exception as e:
        pytest.fail(f"Generación de histograma falló: {e}")

    try:
        graficador.generar_grafico_por_departamento()  # Prueba que el gráfico por departamento se genere sin errores
    except Exception as e:
        pytest.fail(f"Generación de gráfico por departamento falló: {e}")

    try:
        graficador.generar_eb_por_departamento()  # Prueba que el histograma se genere sin errores
    except Exception as e:
        pytest.fail(f"Generación de histograma falló: {e}")
    
    try:
        graficador.generar_grafico_pastel_eb_total()  # Prueba que el histograma se genere sin errores
    except Exception as e:
        pytest.fail(f"Generación de histograma falló: {e}")

def test_procesamiento_poblacion():
    # Ruta del archivo Excel y nombre de la hoja
    # Instancia de la clase Poblacion
    excel_path = 'Cuadros Estadístico del Tomo II.xlsx'
    sheet_name = 'PET1'
    poblacion = Poblacion(excel_path ,  sheet_name)
    
    # Archivo de cobertura móvil
    cobertura_excel_path = "COBERTURA_MOVIL_CUANTIFICADA.xlsx"
    
    # Ejecutar el procesamiento
    poblacion.ejecutar_procesamiento()
    
    # Verificar que los DataFrames no estén vacíos
    assert poblacion.df_total is not None, "df_total es None"
    assert not poblacion.df_total.empty, "df_total está vacío"
    assert poblacion.df_reducido is not None, "df_reducido es None"
    assert not poblacion.df_reducido.empty, "df_reducido está vacío"
    
    if poblacion.df_estaciones is not None:
        assert not poblacion.df_estaciones.empty, "df_estaciones está vacío"
    
    # Cargar datos activa y agregar población activa
    poblacion.cargar_datos_activa()
    poblacion.agregar_poblacion_activa()
    
    # Validar que se agregó la columna de población activa
    assert 'POBLACION_ACTIVA' in poblacion.df_reporte.columns, "No se agregó la columna 'POBLACION_ACTIVA'"
    
    # Ajustar reporte con propuestas
    poblacion.agregar_propietas_eb_pea()
    
    # Validar que las columnas necesarias se generaron correctamente
    assert 'ALCANCE_EB_PEA' in poblacion.df_reporte.columns, "No se generó 'ALCANCE_EB_PEA'"
    assert 'EB_NECESARIAS_PEA' in poblacion.df_reporte.columns, "No se generó 'EB_NECESARIAS_PEA'"
    assert 'PROPUESTAS_EB_PEA' in poblacion.df_reporte.columns, "No se generó 'PROPUESTAS_EB_PEA'"
    
    # Verificar dimensiones del Excel ajustadas
    wb = openpyxl.load_workbook("REPORTE_PROVINCIA.xlsx")
    ws = wb.active
    for col in ws.columns:
        column = col[0].column_letter
        assert ws.column_dimensions[column].width > 0, f"El ancho de la columna {column} no fue ajustado"
    
    # Verificar que los archivos se guardaron correctamente
    assert pd.read_excel('archivo_poblacion_total.xlsx').shape[0] > 0
    assert pd.read_excel('archivo_poblacion_reducida.xlsx').shape[0] > 0
    if poblacion.df_estaciones is not None:
        assert pd.read_excel('estaciones_base_por_provincia.xlsx').shape[0] > 0
    
    print("Pruebas completadas con éxito")


if __name__ == "__main__":
    pytest.main()