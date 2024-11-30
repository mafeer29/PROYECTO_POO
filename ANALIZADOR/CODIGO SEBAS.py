import pandas as pd # para hacer un DataFrame
import openpyxl # Para utilizar el excel deirectamente 
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
        self.datos = pd.read_excel(archivo_excel) # Con esto lo convierte en un DataFrame
        self.centros_poblados = []  # Lista para almacenar los resultados (Pendiente Borrar)

    # Iniciamos un nuevo método para obtener los atributos que necesitamos
    def obtener_atributos_clave(self):
        # Extraer las columnas clave del DataFrame
        atributos_clave = self.datos[[ 'DEPARTAMENTO', "PROVINCIA", "DISTRITO", 'CENTRO_POBLADO', "UBIGEO_CCPP", 'EMPRESA_OPERADORA', '2G', '3G', '4G', '5G',
            'HASTA_1_MBPS', 'MÁS_DE_1_MBPS', 'CANT_EB_2G', 'CANT_EB_3G', 'CANT_EB_4G', 'CANT_EB_5G'
        ]] # Como habia mas datos, vamos a extraer Departamento, etc (Hace otro DataFrame)
        return atributos_clave # Aqui nos lo retorna


# Creamos una clase que herede la anterior
class ActualizarDatos(AnalizadorDatos): # Herencia
    def __init__(self, archivo_excel):
        # Inicializamos la clase base
        super().__init__(archivo_excel) #Cargamos el archivo excel 
        self.atributos_cuantificados = None  # Aqui vamos a almacenar el nuevo DataFrame con la cuantificacion realizada

    # Método para cuantificar la cobertura de tecnologías 2G, 3G, 4G y 5G
    def cuantificar_cobertura(self):
        # Obtenemos los atributos clave
        # Con el la herramineta copy hace una copia para que no tome la misma posicion de memoria  (Para que no halla errores se hace esto)
        self.atributos_cuantificados = self.obtener_atributos_clave().copy()

        # Sumamos los atributos que presenta cada centro poblado
        # Los valores son 1 si existe cobertura, 0 si no existe
        # el loc se usa para usar solo una columna, el :, es para decir que se usaran todas las filas y solo la columna clasificacion
        self.atributos_cuantificados.loc[:,'calificacion'] = (
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
        ) # Aqui esta creando una nueva columna llamada calificacion
        # Hallamos la cantidad de eb total en cada CP
        self.atributos_cuantificados["CANT_EB_TOTAL"] = self.atributos_cuantificados["CANT_EB_2G"] + self.atributos_cuantificados["CANT_EB_3G"] + self.atributos_cuantificados["CANT_EB_4G"] + self.atributos_cuantificados["CANT_EB_5G"]
        # Normalizamos los nombres de los departamentos para evitar errores
        self.atributos_cuantificados['DEPARTAMENTO'] = (
        self.atributos_cuantificados['DEPARTAMENTO'].str.upper().str.strip() # esto sirve para que todos los datos se transformen a mayuscula
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
        reporte_cobertura = "COBERTURA_MOVIL_CUANTIFICADA.xlsx" #Primer Reporte
        # Método para convertir el DataFrame a un archivo Excel
        # por defecto la libreria pandas añade columnas que indican el numero de filas
        # Poniendo Index=False le dices que no añada nada
        self.atributos_cuantificados.to_excel(reporte_cobertura, index=False)
        print(f"Reporte generado: {reporte_cobertura}")
    
    def generar_estadistica(self):
        # Con el describe() esta funcion lo que hace es generar un resumen estadistico del a columna 
        # Osea calcula su media, deviacion estandar ,etc
        estadisticas = self.atributos_cuantificados['calificacion'].describe()

        # Convertir las estadísticas en un DataFrame
        # Investigar
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
        try: # Manejo de Errores
            # aca solo llama a la columna de calificaciones del dataframe
            calificaciones = self.actualizar_datos.atributos_cuantificados['calificacion']
            #Establece el tamaño de la figura
            plt.figure(figsize=(10, 6))
            #Genera histograma con 16 intervalos, color de fondo, borde y transparencia
            # El color define el color del as barras del histograma
            # el edgecolor , estable los bordes de cada barra
            # El alpha define la transparencia de las barras varia entre 0 a 1 
            # El bins define el numero de intervalor que se dividira los datos del histograma
            # el plt.hist es una funcion para visualizar el histograma en la distribuccion de los datos de clasificacion
            plt.hist(calificaciones, bins=16, color='skyblue', edgecolor='black', alpha=0.7)
            #Titulo y etiquetas para el eje x e y
            plt.title('Distribución de Calificaciones de Cobertura')
            plt.xlabel('Calificación')
            plt.ylabel('Frecuencia')
            #Cuadricula en el eje y con transparencia de 0.75
            plt.grid(axis='y', alpha=0.75) #  No Entender
            plt.savefig('calificaciones.png', bbox_inches='tight') # Guarda el grafico como una imagen en un archivo en formato png
            plt.show() # Muestra la imagen
        # En caso de que no exista la columna mencionada
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        # Cualquier otro tipo de error
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")

    def generar_grafico_por_departamento(self):
        try:
            # Agrupamos por departamento y calculamos la media de las calificaciones
            # Agrupar los elementos todos los con se iguales y calculara su media
            promedio_departamentos = self.actualizar_datos.atributos_cuantificados.groupby('DEPARTAMENTO')['calificacion'].mean()

            # Crear un gráfico de barras con los promedios de cada departamento
            plt.figure(figsize=(10, 6)) # el tamaño del aimagen
            promedio_departamentos.plot(kind='bar', color='skyblue') # el color de la barras 

            # Añadir títulos y etiquetas
            plt.title('Calificación Promedio por Departamento', fontsize=14)
            plt.xlabel('Departamento', fontsize=12)
            plt.ylabel('Calificación Promedio', fontsize=12)

            # Rotar etiquetas de los departamentos
            plt.xticks(rotation=90) # Rota la imagen 90 grados por estetica

            # Ajustamos y mostramos el gráfico
            plt.tight_layout()
            plt.savefig('calificacion_distrito.png', bbox_inches='tight') 
            plt.show() 
            print("Gráfico de calificación promedio por departamento generado y mostrado.")
        except KeyError:
            print("Error: No se han cuantificado los datos. Por favor, llama a 'cuantificar_cobertura()' primero.")
        except Exception as e:
            print(f"Ocurrió un error inesperado: {e}")
    def generar_grafico_eb(self):
        try:
            promedio_eb_dpt = self.actualizar_datos.atributos_cuantificados.groupby('DEPARTAMENTO')['CANT_EB_TOTAL'].mean()
            # Crear un gráfico de barras con los promedios de cada departamento
            plt.figure(figsize=(10, 6)) # el tamaño del aimagen
            promedio_eb_dpt.plot(kind='bar', color='skyblue') # el color de la barras 

            # Añadir títulos y etiquetas
            plt.title('Estaciones Base Promedio por Departamento', fontsize=14)
            plt.xlabel('Departamento', fontsize=12)
            plt.ylabel('Estaciones Base', fontsize=12)

            # Rotar etiquetas de los departamentos
            plt.xticks(rotation=90) # Rota la imagen 90 grados por estetica

            # Ajustamos y mostramos el gráfico
            plt.tight_layout()
            plt.savefig('eb_departamento.png', bbox_inches='tight') 
            plt.show() 
            print("Gráfico de eb por departamento generado y mostrado.")
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
    def capacidad_eb(self):
        # Calcular habitantes por antena 2G en cada centro poblado
        self.atributos_cuantificados["ALCANCE_EB"] = (
            self.atributos_cuantificados['POBLACION'] / self.atributos_cuantificados['CANT_EB_TOTAL']
        ).replace([float('inf'), float('nan')], 0)

        # Iterar sobre cada fila para evaluar si el centro poblado necesita más estaciones base 4G
        for index, row in self.atributos_cuantificados.iterrows():
            if row["ALCANCE_EB"] <= 150:
                self.atributos_cuantificados.at[index, "PROPUESTA_EB"]= f"La cantidad de EB ({row["CANT_EB_TOTAL"]})es suficiente pero podemos aumentar las estaciones base 2G para mejorar la capacidad en el centro poblado {row['CENTRO_POBLADO']}."
            else:
                self.atributos_cuantificados.at[index, "PROPUESTA_EB"] = f"La cantidad de estaciones base ({row['CANT_EB_TOTAL']}) es muy baja para la población en {row['CENTRO_POBLADO']}, debemos aumentar la cantidad de EB."
            if row["CANT_EB_TOTAL"] == 0 and row["POBLACION"] != 0:
                self.atributos_cuantificados.at[index, "PROPUESTA_EB"]= f"La cantidad de EB ({row["CANT_EB_TOTAL"]}) es nula, debemos aumentar las estaciones base para mejorar la capacidad en el centro poblado {row['CENTRO_POBLADO']}."
            if row["CANT_EB_TOTAL"] == 0 and row["POBLACION"] == 0:
                self.atributos_cuantificados.at[index, "PROPUESTA_EB"]= f"Actualmente no se ecuentran registros de población en {row['CENTRO_POBLADO']}."
        # Imprimimos para verificar
        print(self.atributos_cuantificados.head())
        self.atributos_cuantificados.to_excel("REPORTE.xlsx", index=False)

        # Abrir el archivo Excel previamente guardado con pandas utilizando openpyxl
        wb = openpyxl.load_workbook("REPORTE.xlsx")
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
        wb.save("REPORTE.xlsx")
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
graficador.generar_grafico_eb()

# Proporcionamos el archivo necesario para las soluciones
solucionador = PropuestasSoluciones(
    "COBERTURA MOVIL.xlsx",
    "CCPP_INEI.xlsx"
)
solucionador.capacidad_eb()
# Crear una instancia de la clase y pasar el archivo Excel con los datos
prediccion = PrediccionSituacionFutura('REPORTE_CUANTIFICADO.xlsx')
# Entrenar el modelo
if prediccion.model is None:
    prediccion.entrenar_modelo()

# Hacer una predicción
prediccion.predecir_situacion('CHACHAPOYAS', 'CHACHAPOYAS', 'CHACHAPOYAS')
