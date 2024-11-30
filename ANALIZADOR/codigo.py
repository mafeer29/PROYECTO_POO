import pandas as pd # para hacer un DataFrame
import openpyxl # Para utilizar el excel deirectamente 
import unicodedata
from openpyxl import load_workbook
import matplotlib.pyplot as plt # para graficar
import os
import numpy as np # para arreglos
import tensorflow as tf # Para el aprendizaje automatico  yredes neuronales 
from tensorflow.keras.models import Sequential #importa al modelo Secuencial 
from tensorflow.keras.layers import Dense # Es para agregar capas de manera secuencial
from sklearn.model_selection import train_test_split # divide los datos en entrenamietno y prueba .-*
from sklearn.preprocessing import StandardScaler, LabelEncoder # las diferentes herramientas utilizadas

## Iniciamos nuestra clase para analizar la cobertura a partir de nuestra base de datos
class AnalizadorDatos:
    def __init__(self, archivo_excel):
        # Cargar el archivo excel en un DataFrame
        self.datos = pd.read_excel(archivo_excel)
        self.centros_poblados = []  # Lista para almacenar los resultados

    # Iniciamos un nuevo método para obtener los atributos que necesitamos
    def obtener_atributos_clave(self):
        # Extraer las columnas clave del DataFrame
        atributos_clave = self.datos[[ 'DEPARTAMENTO', "PROVINCIA", "DISTRITO", 'CENTRO_POBLADO', "UBIGEO_CCPP", 'EMPRESA_OPERADORA', '2G', '3G', '4G', '5G',
            'HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G'
        ]]
        return atributos_clave


# Creamos una clase que herede la anterior
class ActualizarDatos(AnalizadorDatos):
    def __init__(self, archivo_excel):
        # Inicializamos la clase base
        super().__init__(archivo_excel)
        self.atributos_cuantificados = None  # Atributo para almacenar los datos cuantificados

    # Método para cuantificar la cobertura de tecnologías 2G, 3G, 4G y 5G
    def cuantificar_cobertura(self):
        # Obtenemos los atributos clave
        self.atributos_cuantificados = self.obtener_atributos_clave().copy()

        # Sumamos los atributos que presenta cada centro poblado
        # Los valores son 1 si existe cobertura, 0 si no existe
        self.atributos_cuantificados.loc[:,'CALIFICACION'] = (
            self.atributos_cuantificados['2G']*2 + 
            self.atributos_cuantificados['3G']*3 + 
            self.atributos_cuantificados['4G']*4 + 
            self.atributos_cuantificados['5G']*5 +
            self.atributos_cuantificados["HASTA_1_MBPS"]*0.5 +
            self.atributos_cuantificados["MÁS_DE_1_MBPS"]*1 +
            self.atributos_cuantificados["CANT_EB_2G"] +
            self.atributos_cuantificados["CANT_EB_3G"] +
            self.atributos_cuantificados["CANT_EB_4G"] +
            self.atributos_cuantificados["CANT_EB_5G"]
        )
        self.atributos_cuantificados["CALIFICACION"] = (
        self.atributos_cuantificados["CALIFICACION"].apply(lambda x: int(round(x)))
        )
        # Hallamos la cantidad de eb total en cada CP
        self.atributos_cuantificados["CANT_EB_TOTAL"] = self.atributos_cuantificados["CANT_EB_2G"] + self.atributos_cuantificados["CANT_EB_3G"] + self.atributos_cuantificados["CANT_EB_4G"] + self.atributos_cuantificados["CANT_EB_5G"]
        # Normalizamos los nombres de los departamentos para evitar errores
        self.atributos_cuantificados['DEPARTAMENTO'] = (
        self.atributos_cuantificados['DEPARTAMENTO'].str.upper().str.strip()
        )
        self.atributos_cuantificados['PROVINCIA'] = (
        self.atributos_cuantificados['PROVINCIA'].str.upper().str.strip()
        )
        self.atributos_cuantificados['DISTRITO'] = (
        self.atributos_cuantificados['DISTRITO'].str.upper().str.strip()
        )
        self.atributos_cuantificados['CENTRO_POBLADO'] = (
        self.atributos_cuantificados['CENTRO_POBLADO'].str.upper().str.strip()
        )

        # Retornamos los datos cuantificados
        return self.atributos_cuantificados

    # Generar reporte de atributos_cuantificados
    def generar_reporte(self):
        reporte_cobertura = "COBERTURA_MOVIL_CUANTIFICADA.xlsx"
        # Método para convertir el DataFrame a un archivo Excel
        self.atributos_cuantificados.to_excel(reporte_cobertura, index=False)
        print(f"Reporte generado: {reporte_cobertura}")
    
    def generar_estadistica(self):
        # Obtener estadísticas descriptivas
        estadisticas = self.atributos_cuantificados['CALIFICACION'].describe()

        # Convertir las estadísticas en un DataFrame
        df_estadisticas = estadisticas.to_frame(name='Valor').reset_index()
        df_estadisticas.columns = ['Estadística', 'Valor']  # Renombrar columnas para mayor claridad

        # Guardar en un archivo Excel
        reporte_estadistica = "ESTADISTICA_DESCRIPTIVA.xlsx"
        df_estadisticas.to_excel(reporte_estadistica, index=False)
        print(f"Estadística generada: {reporte_estadistica}")

class GenerarGraficas:
    def __init__(self, actualizar_datos):
        self.actualizar_datos = actualizar_datos  # Instancia de ActualizarDatos

    def generar_histograma_calificacion(self):
        try:
            calificaciones = self.actualizar_datos.atributos_cuantificados['CALIFICACION']
            #Establece el tamaño de la figura
            plt.figure(figsize=(10, 6))
            #Genera histograma con 16 intervalos, color de fondo, borde y transparencia
            plt.hist(calificaciones, bins=16, color='skyblue', edgecolor='black', alpha=0.7)
            #Titulo y etiquetas para el eje x e y
            plt.title('Distribución de Calificaciones de Cobertura')
            plt.xlabel('Calificación')
            plt.ylabel('Frecuencia')
            #Cuadricula en el eje y con transparencia de 0.75
            plt.grid(axis='y', alpha=0.75)
            plt.savefig('calificaciones.png', bbox_inches='tight') 
            plt.show()
        # En caso de que no exista la columna mencionada
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        # Cualquier otro tipo de error
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")

    def generar_grafico_por_departamento(self):
        try:
            # Agrupamos por departamento y calculamos la media de las calificaciones
            promedio_departamentos = self.actualizar_datos.atributos_cuantificados.groupby('DEPARTAMENTO')['CALIFICACION'].mean()

            # Crear un gráfico de barras con los promedios de cada departamento
            plt.figure(figsize=(10, 6))  # Tamaño de la imagen
            barras = promedio_departamentos.plot(kind='bar', color='skyblue')  # Color de las barras 

            # Añadir títulos y etiquetas
            plt.title('Calificación Promedio por Departamento', fontsize=14)
            plt.xlabel('Departamento', fontsize=12)
            plt.ylabel('Calificación Promedio', fontsize=12)

            # Rotar etiquetas de los departamentos
            plt.xticks(rotation=90)  # Rota la imagen 90 grados por estética

            # Añadir las calificaciones sobre cada barra
            for i, valor in enumerate(promedio_departamentos):
                plt.text(i, valor + 0.05, f'{valor:.2f}', ha='center', va='bottom', fontsize=10, color='black')

            # Ajustamos y mostramos el gráfico
            plt.tight_layout()
            plt.savefig('calificacion_distrito.png', bbox_inches='tight') 
            plt.show() 
            print("Gráfico de calificación promedio por departamento generado y mostrado.")
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")

    def generar_eb_por_departamento(self):
        try:
            # Calculamos el total acumulado de estaciones base por departamento
            total_eb_dpt = self.actualizar_datos.atributos_cuantificados.groupby('DEPARTAMENTO')['CANT_EB_TOTAL'].sum()

            # Crear un gráfico de barras con los totales de estaciones base por departamento
            plt.figure(figsize=(10, 6))  # Tamaño de la imagen
            total_eb_dpt.plot(kind='bar', color='skyblue')  # Color de las barras

            # Añadir títulos y etiquetas
            plt.title('Total de Estaciones Base por Departamento', fontsize=14)
            plt.xlabel('Departamento', fontsize=12)
            plt.ylabel('Total de Estaciones Base', fontsize=12)

            # Añadir las cantidades totales de EB sobre las barras
            for i, v in enumerate(total_eb_dpt):
                plt.text(i, v + 50, str(v), ha='center', fontsize=10)  # Ajusta el valor 50 para posicionar el texto

            # Rotar etiquetas de los departamentos
            plt.xticks(rotation=90)  # Rota las etiquetas de los departamentos

            # Ajustamos y mostramos el gráfico
            plt.tight_layout()
            plt.savefig('total_eb_departamento.png', bbox_inches='tight')  # Guardar la imagen
            plt.show()  # Mostrar el gráfico
            print("Gráfico de total de estaciones base por departamento generado y mostrado.")
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")


    def generar_grafico_pastel_eb_total(self):
        try:
            # Calculamos el total de estaciones base para cada tecnología
            total_2g = self.actualizar_datos.atributos_cuantificados['CANT_EB_2G'].sum()
            total_3g = self.actualizar_datos.atributos_cuantificados['CANT_EB_3G'].sum()
            total_4g = self.actualizar_datos.atributos_cuantificados['CANT_EB_4G'].sum()
            total_5g = self.actualizar_datos.atributos_cuantificados['CANT_EB_5G'].sum()
            
            # Lista con los totales por tecnología
            totales_eb = [total_2g, total_3g, total_4g, total_5g]
            tecnologias = ['2G', '3G', '4G', '5G']
            
            # Crear el gráfico de pastel con ajustes visuales
            plt.figure(figsize=(8, 6))
            wedges, texts, autotexts = plt.pie(totales_eb, 
                                            labels=tecnologias, 
                                            autopct='%1.1f%%', 
                                            startangle=140, 
                                            colors=['#66b3ff', '#99ff99', '#ff9999', '#ffcc99'],  # Colores similares al gráfico de barras
                                            wedgeprops={'edgecolor': 'black', 'linewidth': 1.5},  # Bordes definidos
                                            textprops={'fontsize': 12, 'color': 'black'})
            
            # Estilo del porcentaje
            for autotext in autotexts:
                autotext.set_fontsize(14)
                autotext.set_fontweight('bold')
            
            # Título del gráfico
            plt.title('Distribución de Estaciones Base por Tecnología en el País', fontsize=16, fontweight='bold', color='navy')
            
            # Asegura que el gráfico sea circular
            plt.axis('equal')  
            plt.tight_layout()
            plt.savefig('eb_distribucion.png', bbox_inches='tight')
            plt.show()
            
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")

class PropuestasSoluciones(ActualizarDatos):
    def __init__(self, archivo_excel, archivo_poblacion):

        # Inicializamos la clase base
        super().__init__(archivo_excel)
        # Cargamos los datos de población
        self.datos_poblacion = pd.read_excel(archivo_poblacion)

        # Verificamos si los datos de población se han cargado correctamente
        if self.datos_poblacion is None or self.datos_poblacion.empty: # El empty nos dice que si el DataFrame esta vacio sera True
            raise ValueError("Error al cargar los datos de población: el DataFrame está vacío o es None.")
        
        # Seleccionamos las columnas que necesitamos
        self.datos_poblacion = self.datos_poblacion[["Departamento", "Provincia","Distrito","Centro Poblado", "Id Centro Poblado", "Población censada"]]
        
        # Renombrar columnas en el DataFrame de población
        self.datos_poblacion.rename(columns={
            "Departamento": "DEPARTAMENTO",
            "Provincia": "PROVINCIA",
            "Distrito": "DISTRITO",
            "Centro Poblado": "CENTRO_POBLADO",
            "Id Centro Poblado": "UBIGEO_CCPP",
            "Población censada": "POBLACION"
        }, inplace=True)

        # Asignar valor 0 si hay datos nulos en la columna "POBLACION"
        self.datos_poblacion["POBLACION"] = self.datos_poblacion["POBLACION"].fillna(0)
        self.atributos_cuantificados = pd.read_excel("COBERTURA_MOVIL_CUANTIFICADA.xlsx") # lo que hace read_excel es leer el excel y transformarlo a un DataFrame
        try:
            # Usamos la función merge para agregar la columna deseada
            self.atributos_cuantificados = pd.merge(
            self.atributos_cuantificados,
            self.datos_poblacion[["DEPARTAMENTO", "PROVINCIA", "DISTRITO", "CENTRO_POBLADO", "POBLACION"]],
            on=["DEPARTAMENTO", "PROVINCIA", "DISTRITO", "CENTRO_POBLADO"],  # Especificar varias columnas como lista
            how="inner"  # Esto indica que solo queremos filas donde haya coincidencias
            )
            reporte_cuantificado = 'REPORTE_CUANTIFICADO.xlsx'
            self.atributos_cuantificados.to_excel(reporte_cuantificado , index=False)
            print(f"Estadística generada: {reporte_cuantificado}")
            if self.atributos_cuantificados.empty:
                print("Error: La fusión resultó en un DataFrame vacío.")
                return
            
            # Verificar que CCPP no tiene información de población
            self.centros_sin_poblacion = self.atributos_cuantificados[self.atributos_cuantificados["POBLACION"]==0]
            self.centros_sin_poblacion.to_excel("CCPP_SIN_INFORMACION.xlsx", index=False)

            # Ahora filtramos los datos para trabajar con los que tienen poblacion
            self.atributos_cuantificados = self.atributos_cuantificados[self.atributos_cuantificados["POBLACION"] != 0]
        except KeyError as e:
            # Esto indica si hay columnas faltantes
            print(f"Error: {e}")
        except Exception as e:
            # Esto es para otro tipo de error
            print(f"Ocurrió un error inesperado: {e}")
    def capacidad_eb_cp(self):
        # Calcular habitantes por antena en cada centro poblado
        self.atributos_cuantificados["ALCANCE_EB"] = (
            self.atributos_cuantificados['POBLACION'] / self.atributos_cuantificados['CANT_EB_TOTAL']
        ).replace([float('inf'), float('nan')], 0)

        self.atributos_cuantificados["ALCANCE_EB"] = (
            self.atributos_cuantificados["ALCANCE_EB"].apply(lambda x: int(round(x)))
        )

        self.atributos_cuantificados["POBLACION_CUBIERTA"] = (
            self.atributos_cuantificados['CANT_EB_TOTAL'] * 150
        ).replace([float('inf'), float('nan')], 0)

        self.atributos_cuantificados["POBLACION_CUBIERTA"] = (
            self.atributos_cuantificados["POBLACION_CUBIERTA"].apply(lambda x: int(round(x)))
        )

        # Si la población cubierta es mayor o igual a la población, asignar 0 a POBLACION_NO_CUBIERTA
        self.atributos_cuantificados.loc[
            self.atributos_cuantificados["POBLACION_CUBIERTA"] >= self.atributos_cuantificados["POBLACION"],
            "POBLACION_NO_CUBIERTA"
        ] = 0

        # Si la población cubierta es menor que la población, calcular la diferencia y reemplazar valores inf y NaN por 0
        self.atributos_cuantificados.loc[
            self.atributos_cuantificados["POBLACION_CUBIERTA"] < self.atributos_cuantificados["POBLACION"],
            "POBLACION_NO_CUBIERTA"
        ] = (
            self.atributos_cuantificados['POBLACION'] - self.atributos_cuantificados["POBLACION_CUBIERTA"]
        ).replace([float('inf'), float('nan')], 0)

        self.atributos_cuantificados["EB_NECESARIAS"] = self.atributos_cuantificados["POBLACION"] / 150
        # Asignar 1 a valores menores que 1 y mayores que 0
        self.atributos_cuantificados.loc[
            (self.atributos_cuantificados["EB_NECESARIAS"] < 1) & (self.atributos_cuantificados["EB_NECESARIAS"] > 0),
            "EB_NECESARIAS"
        ] = 1

        # Redondeo de los valores y convertir a entero
        self.atributos_cuantificados["EB_NECESARIAS"] = np.round(self.atributos_cuantificados["EB_NECESARIAS"]).astype(int)

        self.atributos_cuantificados["EB_FALTANTES"] = (
            self.atributos_cuantificados["EB_NECESARIAS"] - self.atributos_cuantificados["CANT_EB_TOTAL"]
        ).replace([float('inf'), float('nan')], 0)

        # Iterar sobre cada fila para evaluar si el centro poblado necesita más estaciones base 4G
        for index, row in self.atributos_cuantificados.iterrows():
            if row["POBLACION"] != 0:
                if row["CANT_EB_TOTAL"] == 0:
                    self.atributos_cuantificados.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es nula, debemos aumentar las estaciones base a {row['EB_NECESARIAS']} para mejorar la capacidad en {row['CENTRO_POBLADO']}."
                elif row["ALCANCE_EB"] <= 150:
                    self.atributos_cuantificados.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es suficiente y óptima para la población en {row['CENTRO_POBLADO']}."
                elif row["ALCANCE_EB"] > 150:
                    self.atributos_cuantificados.at[index, "PROPUESTA_EB"] = f"El alcance de las estaciones base en {row['CENTRO_POBLADO']} es superior a 150, por lo que se recomienda aumentar la cantidad de EB a {row['EB_NECESARIAS']} para mejorar la cobertura."
            else:
                self.atributos_cuantificados.at[index, "PROPUESTA_EB"] = f"Actualmente no se encuentran registros de población en {row['CENTRO_POBLADO']}."

        # Eliminar las filas duplicadas antes de guardar el reporte
        self.atributos_cuantificados = self.atributos_cuantificados.drop_duplicates()

        # Imprimir para verificar
        print(self.atributos_cuantificados.head())

        # Guardar el archivo Excel
        self.atributos_cuantificados.to_excel("REPORTE_CENTRO_POBLADO.xlsx", index=False)

        # Abrir el archivo Excel previamente guardado con pandas utilizando openpyxl
        wb = openpyxl.load_workbook("REPORTE_CENTRO_POBLADO.xlsx")
        ws = wb.active  # Seleccionar la hoja activa del archivo

        # Iterar sobre todas las columnas del archivo para ajustar el ancho
        for col in ws.columns:
            max_length = 0  # Inicializar variable para el largo máximo de la celda en la columna
            column = col[0].column_letter  # Obtener la letra de la columna (A, B, C, etc.)
            
            # Iterar sobre las celdas de la columna para encontrar la longitud máxima de contenido
            for cell in col:
                try:
                    # Verificar si la longitud del contenido de la celda es mayor que el largo máximo actual
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)  # Actualizar el largo máximo
                except:
                    pass  # Ignorar celdas vacías o errores de tipo de datos
            
            # Ajustar el ancho de la columna basándose en el largo máximo encontrado
            adjusted_width = (max_length + 2)  # Se agrega un pequeño margen de 2 para mejor visibilidad
            ws.column_dimensions[column].width = adjusted_width  # Ajustar el tamaño de la columna

        # Guardar el archivo Excel con las columnas ajustadas
        wb.save("REPORTE_CENTRO_POBLADO.xlsx")
        wb.save("REPORTE.xlsx")

    def capacidad_eb_distrito(self):
        # Creamos un nuevo dataframe
        self.reporte_distrito = self.atributos_cuantificados.copy()

        # Calificación promedio por distrito
        self.reporte_distrito["CALIFICACION"] = self.atributos_cuantificados.groupby('DISTRITO')['CALIFICACION'].transform('mean')
        self.reporte_distrito["CALIFICACION"] = self.reporte_distrito["CALIFICACION"].apply(lambda x: int(round(x)))

        # Sumar la cantidad de estaciones base por tecnología
        self.reporte_distrito["CANT_EB_2G"] = self.atributos_cuantificados.groupby("DISTRITO")["CANT_EB_2G"].transform("sum")
        self.reporte_distrito["CANT_EB_3G"] = self.atributos_cuantificados.groupby("DISTRITO")["CANT_EB_3G"].transform("sum")
        self.reporte_distrito["CANT_EB_4G"] = self.atributos_cuantificados.groupby("DISTRITO")["CANT_EB_4G"].transform("sum")
        self.reporte_distrito["CANT_EB_5G"] = self.atributos_cuantificados.groupby("DISTRITO")["CANT_EB_5G"].transform("sum")
        self.reporte_distrito["CANT_EB_TOTAL"] = self.atributos_cuantificados.groupby("DISTRITO")["CANT_EB_TOTAL"].transform("sum")

        # Sumar la población total por distrito
        self.reporte_distrito["POBLACION_TOTAL"] = self.atributos_cuantificados.groupby("DISTRITO")["POBLACION"].transform("sum")

        # Calcular habitantes por antena en cada centro poblado
        self.reporte_distrito["ALCANCE_EB"] = (self.reporte_distrito['POBLACION_TOTAL'] / self.reporte_distrito['CANT_EB_TOTAL']).replace([float('inf'), float('nan')], 0)
        self.reporte_distrito["ALCANCE_EB"] = self.reporte_distrito["ALCANCE_EB"].apply(lambda x: int(round(x)))

        # Calcular población cubierta (asumimos 150 personas por cada antena base)
        self.reporte_distrito["POBLACION_CUBIERTA"] = (self.reporte_distrito['CANT_EB_TOTAL'] * 150).replace([float('inf'), float('nan')], 0)
        self.reporte_distrito["POBLACION_CUBIERTA"] = self.reporte_distrito["POBLACION_CUBIERTA"].apply(lambda x: int(round(x)))

        # Si la población cubierta es mayor o igual que la población total, asignar 0 a población no cubierta
        self.reporte_distrito.loc[self.reporte_distrito["POBLACION_CUBIERTA"] >= self.reporte_distrito["POBLACION_TOTAL"], "POBLACION_NO_CUBIERTA"] = 0

        # Si la población cubierta es menor que la población total, calcular la diferencia
        self.reporte_distrito.loc[self.reporte_distrito["POBLACION_CUBIERTA"] < self.reporte_distrito["POBLACION_TOTAL"], "POBLACION_NO_CUBIERTA"] = (self.reporte_distrito['POBLACION_TOTAL'] - self.reporte_distrito["POBLACION_CUBIERTA"]).replace([float('inf'), float('nan')], 0)

        # Calcular el número de estaciones base necesarias para cubrir la población total
        self.reporte_distrito["EB_NECESARIAS"] = self.reporte_distrito['POBLACION_TOTAL'] / 150

        # Asignar 1 a valores menores que 1 y mayores que 0
        self.reporte_distrito.loc[(self.reporte_distrito["EB_NECESARIAS"] < 1) & (self.reporte_distrito["EB_NECESARIAS"] > 0), "EB_NECESARIAS"] = 1

        # Redondeo de los valores y convertir a entero
        self.reporte_distrito["EB_NECESARIAS"] = np.round(self.reporte_distrito["EB_NECESARIAS"]).astype(int)

        # Calcular las estaciones base faltantes
        self.reporte_distrito["EB_FALTANTES"] = (self.reporte_distrito["EB_NECESARIAS"] - self.reporte_distrito["CANT_EB_TOTAL"]).replace([float('inf'), float('nan')], 0)

        # Eliminar las columnas innecesarias
        columnas_a_eliminar = ['CENTRO_POBLADO', 'UBIGEO_CCPP', 'EMPRESA_OPERADORA', '2G', '3G', '4G', '5G', 'HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'POBLACION', "PROPUESTA_EB"]
        self.reporte_distrito.drop(columnas_a_eliminar, axis=1, inplace=True)

        # Eliminar filas duplicadas
        self.reporte_distrito = self.reporte_distrito.drop_duplicates()

        # Inicializar la columna "PROPUESTA_EB"
        self.reporte_distrito["PROPUESTA_EB"] = None

        # Evaluar las propuestas de estaciones base
        for index, row in self.reporte_distrito.iterrows():
            if row["POBLACION_TOTAL"] != 0:
                if row["CANT_EB_TOTAL"] == 0:
                    self.reporte_distrito.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es nula, debemos aumentar las estaciones base a {row['EB_NECESARIAS']} para mejorar la capacidad en el distrito de {row['DISTRITO']}."
                elif row["ALCANCE_EB"] <= 150:
                    self.reporte_distrito.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es suficiente y óptima para la población en el distrito de {row['DISTRITO']}."
                elif row["ALCANCE_EB"] > 150:
                    self.reporte_distrito.at[index, "PROPUESTA_EB"] = f"El alcance de las estaciones base en el distrito de {row['DISTRITO']} es superior a 150, por lo que se recomienda aumentar la cantidad de EB a {row['EB_NECESARIAS']} para mejorar la cobertura."
            else:
                self.reporte_distrito.at[index, "PROPUESTA_EB"] = f"Actualmente no se encuentran registros de población en {row['DISTRITO']}."

        # Eliminar filas duplicadas
        self.reporte_distrito = self.reporte_distrito.drop_duplicates()

        # Imprimir para verificar
        print(self.reporte_distrito.head())

        # Guardar el reporte a Excel
        self.reporte_distrito.to_excel("REPORTE_DISTRITO.xlsx", index=False)

        # Ajustar el tamaño de las columnas del Excel
        wb = openpyxl.load_workbook("REPORTE_DISTRITO.xlsx")
        ws = wb.active
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column].width = adjusted_width

        # Guardar el archivo Excel con las columnas ajustadas
        wb.save("REPORTE_DISTRITO.xlsx")

    def capacidad_eb_provincia(self):
        # Creamos un nuevo dataframe
        self.reporte_provincia = self.atributos_cuantificados.copy()
        self.reporte_provincia["CALIFICACION"] = self.atributos_cuantificados.groupby('PROVINCIA')['CALIFICACION'].transform('mean')
        self.reporte_provincia["CALIFICACION"] = (
            self.reporte_provincia["CALIFICACION"].apply(lambda x: int(round(x)))
        )
        self.reporte_provincia["CANT_EB_2G"] = self.atributos_cuantificados.groupby("PROVINCIA")["CANT_EB_2G"].transform("sum")
        self.reporte_provincia["CANT_EB_3G"] = self.atributos_cuantificados.groupby("PROVINCIA")["CANT_EB_3G"].transform("sum")
        self.reporte_provincia["CANT_EB_4G"] = self.atributos_cuantificados.groupby("PROVINCIA")["CANT_EB_4G"].transform("sum")
        self.reporte_provincia["CANT_EB_5G"] = self.atributos_cuantificados.groupby("PROVINCIA")["CANT_EB_5G"].transform("sum")
        self.reporte_provincia["CANT_EB_TOTAL"] = self.atributos_cuantificados.groupby("PROVINCIA")["CANT_EB_TOTAL"].transform("sum")
        self.reporte_provincia["POBLACION_TOTAL"] = self.atributos_cuantificados.groupby("PROVINCIA")["POBLACION"].transform("sum")
        
        # Calcular habitantes por antena en cada centro poblado
        self.reporte_provincia["ALCANCE_EB"] = (
            self.reporte_provincia['POBLACION_TOTAL'] /
            self.reporte_provincia['CANT_EB_TOTAL']
        ).replace([float('inf'), float('nan')], 0)

        self.reporte_provincia["ALCANCE_EB"] = (
            self.reporte_provincia["ALCANCE_EB"].apply(lambda x: int(round(x)))
        )

        self.reporte_provincia["POBLACION_CUBIERTA"] = (
            self.reporte_provincia['CANT_EB_TOTAL'] * 150
        ).replace([float('inf'), float('nan')], 0)

        self.reporte_provincia["POBLACION_CUBIERTA"] = (
            self.reporte_provincia["POBLACION_CUBIERTA"].apply(lambda x: int(round(x)))
        )
        
        self.reporte_provincia.loc[
            self.reporte_provincia["POBLACION_CUBIERTA"] >= self.reporte_provincia["POBLACION_TOTAL"],
            "POBLACION_NO_CUBIERTA"
        ] = 0

        self.reporte_provincia.loc[
            self.reporte_provincia["POBLACION_CUBIERTA"] < self.reporte_provincia["POBLACION_TOTAL"],
            "POBLACION_NO_CUBIERTA"
        ] = (
            self.reporte_provincia['POBLACION_TOTAL'] - self.reporte_provincia["POBLACION_CUBIERTA"]
        ).replace([float('inf'), float('nan')], 0)
        
        self.reporte_provincia["EB_NECESARIAS"] = self.reporte_provincia['POBLACION_TOTAL'] / 150

        # Asignar 1 a valores menores que 1 y mayores que 0
        self.reporte_provincia.loc[
            (self.reporte_provincia["EB_NECESARIAS"] < 1) & (self.reporte_provincia["EB_NECESARIAS"] > 0),
            "EB_NECESARIAS"
        ] = 1

        self.reporte_provincia["EB_NECESARIAS"] = np.round(self.reporte_provincia["EB_NECESARIAS"]).astype(int)
        self.reporte_provincia["EB_FALTANTES"] = (
            self.reporte_provincia["EB_NECESARIAS"] - self.reporte_provincia["CANT_EB_TOTAL"]
        ).replace([float('inf'), float('nan')], 0)
        
        # Eliminar columnas innecesarias
        columnas_a_eliminar = [
            'DISTRITO', 'CENTRO_POBLADO', 'UBIGEO_CCPP', 
            'EMPRESA_OPERADORA', '2G', '3G', '4G', '5G', 'HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'POBLACION', "PROPUESTA_EB"
        ]
        self.reporte_provincia.drop(columnas_a_eliminar, axis=1, inplace=True)
        
        # Eliminar filas duplicadas
        self.reporte_provincia = self.reporte_provincia.drop_duplicates()
        self.reporte_provincia["PROPUESTA_EB"] = None

        for index, row in self.reporte_provincia.iterrows():
            if row["POBLACION_TOTAL"] != 0:
                if row["CANT_EB_TOTAL"] == 0:
                    self.reporte_provincia.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es nula, debemos aumentar las estaciones base a {row['EB_NECESARIAS']} para mejorar la capacidad en la provincia de {row['PROVINCIA']}."
                elif row["ALCANCE_EB"] <= 150:
                    # Si el alcance es menor o igual a 150, la cantidad de estaciones base es suficiente y se considera óptima
                    self.reporte_provincia.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es suficiente y óptima para la población en la provincia de {row['PROVINCIA']}."
                elif row["ALCANCE_EB"] > 150:
                    # Si el alcance es mayor a 150, se debe aumentar la cantidad de EB
                    self.reporte_provincia.at[index, "PROPUESTA_EB"] = f"El alcance de las estaciones base en la provincia de {row['PROVINCIA']} es superior a 150, por lo que se recomienda aumentar la cantidad de EB a {row['EB_NECESARIAS']} para mejorar la cobertura."
            else:
                self.reporte_provincia.at[index, "PROPUESTA_EB"] = f"Actualmente no se encuentran registros de población en {row['PROVINCIA']}."

        
        # Guardar el archivo Excel con el reporte
        self.reporte_provincia.to_excel("REPORTE_PROVINCIA.xlsx", index=False)
        self.reporte_provincia.drop_duplicates()

        # Ajustar el ancho de las columnas
        self.reporte_provincia.to_excel("REPORTE_PROVINCIA.xlsx", index=False)
        wb = load_workbook("REPORTE_PROVINCIA.xlsx")
        ws = wb.active
        for column in ws.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width
        wb.save("REPORTE_PROVINCIA.xlsx")


    def capacidad_eb_departamento(self):
        # Creamos un nuevo dataframe
        self.reporte_departamento = self.atributos_cuantificados.copy()
        self.reporte_departamento["CALIFICACION"] = self.atributos_cuantificados.groupby('DEPARTAMENTO')['CALIFICACION'].transform('mean')
        self.reporte_departamento["CALIFICACION"] = (
            self.reporte_departamento["CALIFICACION"].apply(lambda x: int(round(x)))
        )
        
        # Agregar sumas de estaciones base por departamento
        self.reporte_departamento["CANT_EB_2G"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["CANT_EB_2G"].transform("sum")
        self.reporte_departamento["CANT_EB_3G"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["CANT_EB_3G"].transform("sum")
        self.reporte_departamento["CANT_EB_4G"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["CANT_EB_4G"].transform("sum")
        self.reporte_departamento["CANT_EB_5G"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["CANT_EB_5G"].transform("sum")
        self.reporte_departamento["CANT_EB_TOTAL"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["CANT_EB_TOTAL"].transform("sum")
        self.reporte_departamento["POBLACION_TOTAL"] = self.atributos_cuantificados.groupby("DEPARTAMENTO")["POBLACION"].transform("sum")
        
        # Calcular habitantes por antena en cada centro poblado
        self.reporte_departamento["ALCANCE_EB"] = (
            self.reporte_departamento["POBLACION_TOTAL"] / self.reporte_departamento["CANT_EB_TOTAL"]
        ).replace([float('inf'), float('nan')], 0)

        self.reporte_departamento["ALCANCE_EB"] = (
            self.reporte_departamento["ALCANCE_EB"].apply(lambda x: int(round(x)))
        )

        # Calcular población cubierta
        self.reporte_departamento["POBLACION_CUBIERTA"] = (
            self.reporte_departamento['CANT_EB_TOTAL'] * 150
        ).replace([float('inf'), float('nan')], 0)

        self.reporte_departamento["POBLACION_CUBIERTA"] = (
            self.reporte_departamento["POBLACION_CUBIERTA"].apply(lambda x: int(round(x)))
        )
        
        # Calcular población no cubierta
        self.reporte_departamento.loc[
            self.reporte_departamento["POBLACION_CUBIERTA"] >= self.reporte_departamento["POBLACION_TOTAL"],
            "POBLACION_NO_CUBIERTA"
        ] = 0

        self.reporte_departamento.loc[
            self.reporte_departamento["POBLACION_CUBIERTA"] < self.reporte_departamento["POBLACION_TOTAL"],
            "POBLACION_NO_CUBIERTA"
        ] = (
            self.reporte_departamento['POBLACION_TOTAL'] - self.reporte_departamento["POBLACION_CUBIERTA"]
        ).replace([float('inf'), float('nan')], 0)
        
        self.reporte_departamento["EB_NECESARIAS"] = self.reporte_departamento['POBLACION_TOTAL'] / 150
        self.reporte_departamento.loc[
            (self.reporte_departamento["EB_NECESARIAS"] < 1) & (self.reporte_departamento["EB_NECESARIAS"] > 0),
            "EB_NECESARIAS"
        ] = 1

        # Redondeo de los valores y convertir a entero
        self.reporte_departamento["EB_NECESARIAS"] = np.round(self.reporte_departamento["EB_NECESARIAS"]).astype(int)

        # Calcular faltantes de EB
        self.reporte_departamento["EB_FALTANTES"] = (
            self.reporte_departamento["EB_NECESARIAS"] - self.reporte_departamento["CANT_EB_TOTAL"]
        ).replace([float('inf'), float('nan')], 0)
        
        # Eliminar columnas innecesarias
        columnas_a_eliminar = [
            'PROVINCIA', 'DISTRITO', 'CENTRO_POBLADO', 'UBIGEO_CCPP', 
            'EMPRESA_OPERADORA', '2G', '3G', '4G', '5G', 'HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'POBLACION', "PROPUESTA_EB"
        ]
        self.reporte_departamento.drop(columnas_a_eliminar, axis=1, inplace=True)
        
        # **Eliminar duplicados**: Eliminamos duplicados solo basados en la columna 'DEPARTAMENTO', conservando la última aparición
        self.reporte_departamento.drop_duplicates(subset=['DEPARTAMENTO'], keep='last', inplace=True)

        # Asignar propuestas para estaciones base necesarias
        for index, row in self.reporte_departamento.iterrows():
            if row["POBLACION_TOTAL"] != 0:
                if row["CANT_EB_TOTAL"] == 0:
                    self.reporte_departamento.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es nula, debemos aumentar las estaciones base a {row['EB_NECESARIAS']} para mejorar la capacidad en el departamento de {row['DEPARTAMENTO']}."
                elif row["ALCANCE_EB"] <= 150:
                    self.reporte_departamento.at[index, "PROPUESTA_EB"] = f"La cantidad de EB ({row['CANT_EB_TOTAL']}) es suficiente y óptima para la población en el departamento de {row['DEPARTAMENTO']}."
                elif row["ALCANCE_EB"] > 150:
                    self.reporte_departamento.at[index, "PROPUESTA_EB"] = f"El alcance de las estaciones base en el departamento de {row['DEPARTAMENTO']} es superior a 150, por lo que se recomienda aumentar la cantidad de EB a {row['EB_NECESARIAS']} para mejorar la cobertura."
            else:
                self.reporte_departamento.at[index, "PROPUESTA_EB"] = f"Actualmente no se encuentran registros de población en {row['DEPARTAMENTO']}."

        # Eliminar duplicados después de asignar propuestas (en caso de que se haya agregado alguna nueva fila)
        self.reporte_departamento.drop_duplicates(subset=['DEPARTAMENTO'], keep='last', inplace=True)
        
        # Guardar el archivo Excel
        self.reporte_departamento.to_excel("REPORTE_DEPARTAMENTO.xlsx", index=False)

        # Ajustar el ancho de las columnas en el archivo Excel
        wb = openpyxl.load_workbook("REPORTE_DEPARTAMENTO.xlsx")
        ws = wb.active  # Seleccionar la hoja activa del archivo

        for col in ws.columns:
            max_length = 0  # Inicializar variable para el largo máximo de la celda en la columna
            column = col[0].column_letter  # Obtener la letra de la columna (A, B, C, etc.)
            
            # Iterar sobre las celdas de la columna para encontrar la longitud máxima de contenido
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)  # Actualizar el largo máximo
                except:
                    pass  # Ignorar celdas vacías o errores de tipo de datos
            
            # Ajustar el ancho de la columna basándose en el largo máximo encontrado
            adjusted_width = (max_length + 2)  # Se agrega un pequeño margen de 2 para mejor visibilidad
            ws.column_dimensions[column].width = adjusted_width  # Ajustar el tamaño de la columna

        # Guardar el archivo Excel con las columnas ajustadas
        wb.save("REPORTE_DEPARTAMENTO.xlsx")

# Función para quitar tildes
def quitar_tildes(texto):
    return ''.join((c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn'))

class Poblacion:
    def __init__(self, excel_path, sheet_name):
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.data = None
        self.df_total = None
        self.df_reducido = None
        self.df_estaciones = None
        self.df_reporte = None  # Agregar un atributo para el reporte

    def cargar_datos(self):
        # Cargar el archivo Excel
        self.data = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=0)

        # Renombrar columnas si la primera aparece como NaN
        if pd.isna(self.data.columns[0]):
            self.data.rename(columns={self.data.columns[0]: 'Departamento, provincia, area urbana y rural; y sexo'}, inplace=True)

        # Limpiar nombres de columnas
        self.data.columns = self.data.columns.str.strip().str.replace('\n', '')

        # Eliminar tildes en los nombres de las columnas
        self.data.columns = [quitar_tildes(col) for col in self.data.columns]

        # Agregar columnas para DEPARTAMENTO y PROVINCIA
        self.data['DEPARTAMENTO'] = None
        self.data['PROVINCIA'] = None

        # Extraer departamentos y provincias
        provincia_actual = None
        departamento_actual = None

        for index, row in self.data.iterrows():
            columna_principal = row['Departamento, provincia, area urbana y rural; y sexo']

            if isinstance(columna_principal, str):
                columna_principal = columna_principal.strip()

                # Identificar departamentos y provincias
                if columna_principal.isupper() and "URBANA" not in columna_principal and "RURAL" not in columna_principal:
                    if "DEPARTAMENTO" in columna_principal:
                        departamento_actual = columna_principal.replace("DEPARTAMENTO ", "").strip()
                        provincia_actual = None  # Reiniciar la provincia cuando se detecta un departamento
                    else:
                        provincia_actual = columna_principal.strip()
                else:
                    # Asignar provincia y departamento
                    self.data.at[index, 'PROVINCIA'] = provincia_actual
                    self.data.at[index, 'DEPARTAMENTO'] = departamento_actual

        # Eliminar la palabra 'PROVINCIA' en los valores de la columna 'PROVINCIA'
        self.data['PROVINCIA'] = self.data['PROVINCIA'].str.replace('PROVINCIA ', '', regex=False)

        # Eliminar tildes en los valores de las celdas
        self.data = self.data.applymap(lambda x: quitar_tildes(x) if isinstance(x, str) else x)

    def procesar_total(self):
        """Procesar los datos de población total."""
        df_filtered = self.data[self.data['Departamento, provincia, area urbana y rural; y sexo']
                                 .str.contains('URBANA|RURAL', na=False)]

        self.df_total = df_filtered.groupby(['DEPARTAMENTO', 'PROVINCIA'], as_index=False).sum()
        self.df_total = self.df_total[['DEPARTAMENTO', 'PROVINCIA', 'Total', '14 a 29', '30 a 44', '45 a 64', '65 y mas']]

    def ajustar_reducida(self):
        """Ajustar la columna '65 y mas' y recalcular los totales."""
        df_filtered = self.data[self.data['Departamento, provincia, area urbana y rural; y sexo']
                                 .str.contains('URBANA|RURAL', na=False)].copy()

        def ajustar_poblacion(row):
            if 'URBANA' in row['Departamento, provincia, area urbana y rural; y sexo']:
                return round(row['65 y mas'] * 0.35)  # 35% para zonas urbanas
            elif 'RURAL' in row['Departamento, provincia, area urbana y rural; y sexo']:
                return round(row['65 y mas'] * 0.20)  # 20% para zonas rurales
            return row['65 y mas']

        df_filtered['65 y mas'] = df_filtered.apply(ajustar_poblacion, axis=1)
        df_filtered['Total'] = df_filtered[['14 a 29', '30 a 44', '45 a 64', '65 y mas']].sum(axis=1)

        self.df_reducido = df_filtered.groupby(['DEPARTAMENTO', 'PROVINCIA'], as_index=False).sum()
        self.df_reducido = self.df_reducido[['DEPARTAMENTO', 'PROVINCIA', 'Total', '14 a 29', '30 a 44', '45 a 64', '65 y mas']]

    def guardar_archivos(self):
        """Guardar los datos procesados en archivos Excel."""
        if self.df_total is not None:
            self.df_total.to_excel('archivo_poblacion_total.xlsx', index=False)
            print("Archivo 'archivo_poblacion_total.xlsx' guardado.")
        if self.df_reducido is not None:
            self.df_reducido.to_excel('archivo_poblacion_reducida.xlsx', index=False)
            print("Archivo 'archivo_poblacion_reducida.xlsx' guardado.")
        
    def ejecutar_procesamiento(self, cobertura_excel_path=None):
        """Ejecutar todo el proceso: cargar, procesar y guardar."""
        print("Cargando datos...")
        self.cargar_datos()

        print("Procesando población total...")
        self.procesar_total()

        print("Ajustando población reducida...")
        self.ajustar_reducida()

        print("Guardando archivos...")
        self.guardar_archivos()

    def cargar_datos_activa(self):
        """Cargar los archivos necesarios para agregar la población activa."""
        # Cargar el archivo de población reducida
        self.df_reducido = pd.read_excel('archivo_poblacion_reducida.xlsx')

        # Cargar el archivo de reporte por provincia
        self.df_reporte = pd.read_excel('REPORTE_PROVINCIA.xlsx')

    def agregar_poblacion_activa(self):
        """Agregar la columna 'Total' de archivo_poblacion_reducida a REPORTE_PROVINCIA."""
        # Realizar el merge entre ambos DataFrames usando las columnas 'DEPARTAMENTO' y 'PROVINCIA'
        df_merge = pd.merge(self.df_reporte, self.df_reducido[['DEPARTAMENTO', 'PROVINCIA', 'Total']], on=['DEPARTAMENTO', 'PROVINCIA'], how='left')

        # Renombrar la columna 'Total' como 'POBLACION_ACTIVA'
        df_merge['POBLACION_ACTIVA'] = df_merge['Total']
        df_merge.drop(columns=['Total'], inplace=True)  # Eliminar la columna 'Total' original

        # Asignar el DataFrame con la nueva columna al atributo df_reporte
        self.df_reporte = df_merge

        # Sobrescribir el archivo 'REPORTE_PROVINCIA.xlsx' con la nueva columna
        self.df_reporte.to_excel('REPORTE_PROVINCIA.xlsx', index=False)
        print("Archivo 'REPORTE_PROVINCIA.xlsx' actualizado con la columna POBLACION_ACTIVA.")


    def agregar_propietas_eb_pea(self):
        """Agregar la columna 'PROPUESTAS_EB_PEA' y asegurar que las columnas ALCANCE_EB_PEA y EB_NECESARIAS_PEA sean enteros."""
        # Calcular ALCANCE_EB_PEA y EB_NECESARIAS_PEA
        self.df_reporte['ALCANCE_EB_PEA'] = self.df_reporte['POBLACION_ACTIVA'] / self.df_reporte['CANT_EB_TOTAL']
        self.df_reporte['EB_NECESARIAS_PEA'] = self.df_reporte['POBLACION_ACTIVA'] / 150

        # Reemplazar los valores NaN o inf con 0 antes de la conversión a entero
        self.df_reporte['ALCANCE_EB_PEA'] = self.df_reporte['ALCANCE_EB_PEA'].replace([float('inf'), -float('inf')], 0).fillna(0)
        self.df_reporte['EB_NECESARIAS_PEA'] = self.df_reporte['EB_NECESARIAS_PEA'].replace([float('inf'), -float('inf')], 0).fillna(0)

        # Convertir a enteros (truncar decimales)
        self.df_reporte['ALCANCE_EB_PEA'] = self.df_reporte['ALCANCE_EB_PEA'].astype(int)
        self.df_reporte['EB_NECESARIAS_PEA'] = self.df_reporte['EB_NECESARIAS_PEA'].astype(int)

        # Función para definir la propuesta en 'PROPUESTAS_EB_PEA'
        def propuesta(row):
            if row['ALCANCE_EB_PEA'] < 150:
                return f"La cantidad de estaciones base ({row['CANT_EB_TOTAL']}) es muy baja para la población {row['PROVINCIA']}, debemos aumentar la cantidad de EB a {row['EB_NECESARIAS_PEA']} en total para una cobertura óptima."
            else:
                return f"La cantidad de estaciones base ({row['CANT_EB_TOTAL']}) es suficiente pero podemos aumentar las estaciones base para mejorar la capacidad de la población en {row['PROVINCIA']}."

        # Aplicar la función a cada fila
        self.df_reporte['PROPUESTAS_EB_PEA'] = self.df_reporte.apply(propuesta, axis=1)

        # Guardar el archivo actualizado
        self.df_reporte.to_excel('REPORTE_PROVINCIA.xlsx', index=False)
        print("Archivo 'REPORTE_PROVINCIA.xlsx' actualizado con la columna PROPUESTAS_EB_PEA.")


        # Imprimimos para verificar
        print(self.df_reporte.head())
        self.df_reporte.to_excel("REPORTE_PROVINCIA.xlsx", index=False)

        # Abrir el archivo Excel previamente guardado con pandas utilizando openpyxl
        wb = openpyxl.load_workbook("REPORTE_PROVINCIA.xlsx")
        ws = wb.active  # Seleccionar la hoja activa del archivo

        # Iterar sobre todas las columnas del archivo
        for col in ws.columns:
            max_length = 0  # Inicializar variable para el largo máximo de la celda en la columna
            column = col[0].column_letter  # Obtener la letra de la columna (A, B, C, etc.)
            
            # Iterar sobre las celdas de la columna para encontrar la longitud máxima de contenido
            for cell in col:
                try:
                    # Verificar si la longitud del contenido de la celda es mayor que el largo máximo actual
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)  # Actualizar el largo máximo
                except:
                    pass  # Ignorar celdas vacías o errores de tipo de datos
            
            # Ajustar el ancho de la columna basándose en el largo máximo encontrado
            adjusted_width = (max_length + 2)  # Se agrega un pequeño margen de 2 para mejor visibilidad
            ws.column_dimensions[column].width = adjusted_width  # Ajustar el tamaño de la columna

        # Guardar el archivo Excel con las columnas ajustadas
        wb.save("REPORTE_PROVINCIA.xlsx")
        
        wb.save("REPORTE_PROVINCIA.xlsx")

class PrediccionSituacionFutura:
    def __init__(self, archivo_datos, archivo_modelo='modelo_entrenado.h5'):
        # Cargar los datos
        self.df = pd.read_excel(archivo_datos)

        # Las columnas a considerar para el modelo
        self.columnas_categoricas = ['DEPARTAMENTO', 'PROVINCIA', 'DISTRITO', 'CENTRO_POBLADO', 'UBIGEO_CCPP', 'EMPRESA_OPERADORA']
        self.columnas_numericas = ['2G','3G','4G','5G','HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G', 'POBLACION']

        self.archivo_modelo = archivo_modelo
        # Inicializar el LabelEncoder para cada columna categórica
        self.label_encoders = {col: LabelEncoder() for col in self.columnas_categoricas}

        # Asegurarse que las columnas categóricas sean transformadas adecuadamente
        for col in self.columnas_categoricas:
            self.df[col] = self.df[col].astype(str)  # Asegurarse de que se entrenen todas las categorías

        # Codificar las columnas categóricas
        for col in self.columnas_categoricas:
            self.df[col] = self.label_encoders[col].fit_transform(self.df[col])

        # Insertar la columna 'SITUACION_FUTURA' con una lógica en base a la conectividad y la población
        # lo que hace el apply es insertar una funcion a cada columna y fila
        # el axis=1 indica que se aplicara solo a las filas
        self.df['SITUACION_FUTURA'] = self.df.apply(self.determinar_situacion_futura, axis=1)

        # Codificar la columna 'SITUACION_FUTURA'
        self.label_encoder_situacion = LabelEncoder()
        self.df['SITUACION_FUTURA'] = self.label_encoder_situacion.fit_transform(self.df['SITUACION_FUTURA'])

        # Rellenar valores vacíos con 0 para las columnas numéricas
        self.df[self.columnas_numericas] = self.df[self.columnas_numericas].fillna(0)

        # Asegurarse de que X_train y y_train sean del tipo adecuado
        self.X = self.df.drop('SITUACION_FUTURA', axis=1).values.astype(np.float32)
        self.y = self.df['SITUACION_FUTURA'].values.astype(np.float32)
        self.archivo_modelo = archivo_modelo

        if os.path.exists(self.archivo_modelo):
            print("Modelo entrenado encontrado. Cargando modelo...")
            self.cargar_modelo_y_encoders()
        else:
            print("No se encontró un modelo entrenado. Es necesario entrenarlo primero.")
            self.model = None

    def determinar_situacion_futura(self, row):
      # Puntajes basados en las condiciones
      puntaje = 0

      # Evaluar la cantidad de estaciones base 5G
      if row['CANT_EB_4G'] > 2:
         puntaje += 1  # Mejorará si hay más de 10 estaciones base 5G
      elif row['CANT_EB_4G'] > 0:
         puntaje += 0  # Se mantiene igual

      # Evaluar el número de usuarios con más de 1 Mbps
      if row['MÁS_DE_1_MBPS'] > 1:
         puntaje += 1  # Mejorará si hay más de 1000 usuarios con más de 1 Mbps
      elif row['MÁS_DE_1_MBPS'] > 0:
         puntaje += 0  # Se mantiene igual

      # Evaluar la población
      if row['POBLACION'] > 300:
         puntaje -= 1  # Puede empeorar si la población es muy alta y hay pocos recursos
      else:
         puntaje += 1  # Mejorará si la población es pequeña y las estaciones base son suficientes
      # Evaluar la cobertura de 5G en relación con las demás tecnologías
      if row['4G'] > 1:  # Supongamos que un 75% de la cobertura 5G es ideal
         puntaje += 1  # Mejorará si la cobertura de 5G es alta
      elif row['4G'] > 0:
         puntaje += 0  # Se mantiene igual

      # Decidir la situación futura en función del puntaje total
      if puntaje >= 3:
         return 'mejorará'
      elif puntaje == 2:
         return 'igual'
      else:
         return 'empeorará'
    def entrenar_modelo(self):
        # Dividir los datos en entrenamiento y prueba (80% - 20%)
        X_train, X_test, y_train, y_test = train_test_split(self.X, self.y, test_size=0.2, random_state=42)

        # Crear el modelo de Deep Learning (Ejemplo sencillo de red neuronal)
        self.model = tf.keras.Sequential([ 
            tf.keras.layers.Input(shape=(self.X.shape[1],)),
            tf.keras.layers.Dense(64, activation='relu'),
            tf.keras.layers.Dense(64, activation='relu'),
            tf.keras.layers.Dense(3, activation='softmax')  # 3 salidas para 'mejorará', 'igual', 'empeorará'
        ])

        # Compilar y entrenar el modelo
        self.model.compile(optimizer='adam', loss='sparse_categorical_crossentropy', metrics=['accuracy'])
        self.model.fit(X_train, y_train, epochs=10, batch_size=32)

        # Evaluar el modelo con los datos de prueba
        test_loss, test_acc = self.model.evaluate(X_test, y_test)
        print(f"Precisión en los datos de prueba: {test_acc * 100:.2f}%")
        self.guardar_modelo_y_encoders()
    def guardar_modelo_y_encoders(self):
        # Guardar el modelo
        self.model.save(self.archivo_modelo)
        print(f"Modelo guardado en {self.archivo_modelo}.")

        # Guardar los LabelEncoders para las columnas categóricas
        for columna, encoder in self.label_encoders.items():
            np.save(f'{columna}_encoder.npy', encoder.classes_)
        np.save('situacion_encoder.npy', self.label_encoder_situacion.classes_)
        print("Codificadores guardados correctamente.")
    def cargar_modelo_y_encoders(self):
        # Cargar el modelo
        self.model = tf.keras.models.load_model(self.archivo_modelo)
        print(f"Modelo cargado desde {self.archivo_modelo}.")

        # Cargar los LabelEncoders
        for columna in self.columnas_categoricas:
            clases = np.load(f'{columna}_encoder.npy', allow_pickle=True)
            self.label_encoders[columna].classes_ = clases
        self.label_encoder_situacion.classes_ = np.load('situacion_encoder.npy', allow_pickle=True)
        print("Codificadores cargados correctamente.")

    def predecir_situacion(self, distrito, centro_poblado, provincia):
        # Convertir los valores categóricos a su codificación numérica
        distrito_codificado = self.transformar_categoria('DISTRITO', distrito)
        centro_poblado_codificado = self.transformar_categoria('CENTRO_POBLADO', centro_poblado)
        provincia_codificada = self.transformar_categoria('PROVINCIA', provincia)

        # Verificar si se pudo codificar correctamente
        if any(val == -1 for val in [ distrito_codificado, centro_poblado_codificado, provincia_codificada]):
            print("No se encontraron datos para la localidad ingresada.")
            return

        # Filtrar el DataFrame con los valores proporcionados por el usuario
        datos_filtrados = self.df[(self.df['DISTRITO'] == distrito_codificado) & 
                                  (self.df['CENTRO_POBLADO'] == centro_poblado_codificado) &
                                  (self.df['PROVINCIA'] == provincia_codificada)]

        # Verificar si se encontró algún registro
        if not datos_filtrados.empty:
            # Obtener las demás columnas que el modelo espera
            cant_2g = datos_filtrados['CANT_EB_2G'].values[0]
            cant_3g = datos_filtrados['CANT_EB_3G'].values[0]
            cant_4g = datos_filtrados['CANT_EB_4G'].values[0]
            cant_5g = datos_filtrados['CANT_EB_5G'].values[0]
            hasta_1mbps = datos_filtrados['HASTA_1_MBPS'].values[0]
            mas_de_1mbps = datos_filtrados['MÁS_DE_1_MBPS'].values[0]
            poblacion = datos_filtrados['POBLACION'].values[0]
            dosG = datos_filtrados['2G'].values[0]
            tresG = datos_filtrados['3G'].values[0]
            cuatroG = datos_filtrados['4G'].values[0]
            cincoG = datos_filtrados['5G'].values[0]
            Situacion = datos_filtrados['SITUACION_FUTURA'].values[0]
            departamento = datos_filtrados['DEPARTAMENTO'].values[0]
            ubigeo = datos_filtrados['UBIGEO_CCPP'].values[0]
            empresa_operadora = datos_filtrados['EMPRESA_OPERADORA'].values[0]
            # Crear el array con los datos para predecir y darle la forma correcta
            datos_para_predecir = np.array([[departamento, distrito_codificado, centro_poblado_codificado, 
                                             provincia_codificada, ubigeo, empresa_operadora,dosG,tresG,cuatroG,cincoG,
                                             cant_2g, cant_3g, cant_4g, cant_5g, hasta_1mbps, mas_de_1mbps, poblacion,Situacion]])

            # Hacer la predicción con el modelo
            prediccion = self.model.predict(datos_para_predecir)

            # Obtener la clase predicha
            prediccion_clase = np.argmax(prediccion, axis=1)[0]

            # Mapear la clase predicha a su valor correspondiente
            resultado = self.label_encoder_situacion.inverse_transform([prediccion_clase])[0]
            ubigeo_codificado = self.label_encoders['UBIGEO_CCPP'].inverse_transform([ubigeo])[0]  
            empresa_codificado = self.label_encoders['EMPRESA_OPERADORA'].inverse_transform([empresa_operadora])[0]
            # Mostrar el resultado
            print(f"La situación futura para el centro poblado {centro_poblado} en {distrito}, {provincia}, con UBIGEO {ubigeo_codificado} y Empresa {empresa_codificado} será: {resultado}")
        else:
            print("No se encontraron datos para la localidad ingresada.")

    def transformar_categoria(self, columna, valor):
        """Función auxiliar para manejar la transformación de categorías con LabelEncoder"""
        # Verificar si el valor existe en la columna original antes de codificar
        if valor in self.label_encoders[columna].classes_:
            return self.label_encoders[columna].transform([valor])[0]
        else:
            print(f"El valor '{valor}' no se encuentra en la columna '{columna}'.")
            return -1  # Valor por defecto si no se encuentra

# Ejemplo de uso:
# Primero, se debe proporcionar el nombre del archivo Excel que contiene los datos de cobertura
analizador = ActualizarDatos("COBERTURA MOVIL.xlsx")

# Cuantificamos la cobertura
atributos_cuantificados = analizador.cuantificar_cobertura()

# Generamos el reporte
analizador.generar_reporte()
analizador.generar_estadistica()

# Generamos las graficas
graficador = GenerarGraficas(analizador)
graficador.generar_histograma_calificacion()
graficador.generar_grafico_por_departamento()
graficador.generar_eb_por_departamento()
graficador.generar_grafico_pastel_eb_total()

# Proporcionamos el archivo necesario para las soluciones
solucionador = PropuestasSoluciones(
    "COBERTURA MOVIL.xlsx",
    "CCPP_INEI.xlsx"
)
solucionador.capacidad_eb_cp()
solucionador.capacidad_eb_departamento()
solucionador.capacidad_eb_distrito()
solucionador.capacidad_eb_provincia()
def main():
    # Ruta de los archivos Excel
    excel_path = 'Cuadros Estadístico del Tomo II.xlsx'  
    sheet_name = 'PET1'  

    # Crear una instancia de la clase Poblacion
    poblacion = Poblacion(excel_path, sheet_name)

    # Ejecutar el procesamiento
    poblacion.ejecutar_procesamiento()

    # Cargar los datos de población activa y agregar la columna
    print("Cargando y agregando población activa...")
    poblacion.cargar_datos_activa()
    poblacion.agregar_poblacion_activa()

    # Agregar la columna de propuestas
    poblacion.agregar_propietas_eb_pea()

# Asegurarse de que este bloque solo se ejecute si el script es ejecutado directamente
if __name__ == '_main_':
    main()
# Crear una instancia de la clase y pasar el archivo Excel con los datos
    prediccion = PrediccionSituacionFutura('REPORTE_CUANTIFICADO.xlsx')
# Crear una instancia de la clase y pasar el archivo Excel con los datos
prediccion = PrediccionSituacionFutura('REPORTE_CUANTIFICADO.xlsx')
# Entrenar el modelo
if prediccion.model is None:
    prediccion.entrenar_modelo()

# Hacer una predicción
prediccion.predecir_situacion('CHACHAPOYAS', 'CHACHAPOYAS', 'CHACHAPOYAS')
