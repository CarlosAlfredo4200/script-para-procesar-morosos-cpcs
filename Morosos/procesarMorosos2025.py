import pandas as pd
from datetime import datetime

# entrada de datos
print("")
dia_cita = input("Ingrese el dia de la cita (eje. 9): ").strip()
mes_cita = input("Ingrese el mes de la cita (eje. Agosto): ").strip()
hora = input("Ingrese la hora de la cita (eje. 9 am): ").strip()
lugar = input("Ingrese el lugar de la cita : ").strip()


input_columns =[0,2,5,7,22]
archivo = pd.read_excel('./data.xls', usecols=input_columns)
df = pd.DataFrame(archivo)

# Cortar filas desde la 16 en adelante
df = df.iloc[16:].reset_index(drop=True)

# Eliminar filas con NaN en la columna 'Unnamed: 7'
df = df.dropna(subset=['Unnamed: 7']).reset_index(drop=True)

# llenar hacia abajo
pd.set_option('future.no_silent_downcasting', True)
df = df.ffill().infer_objects(copy=False)

 


# limpiar columas y extraer solo el nombre 
df['Nombre estudiante'] = df['Unnamed: 0'].astype(str).str.split("\n").str[0]
df['acudiente'] = df['Unnamed: 2'].astype(str).str.split("\n").str[0]
df['grado'] = df['Unnamed: 5'].astype(str).str.split(" ").str[0]

# Renombrar columna
df['deuda total'] = df['Unnamed: 22']
df['pensiones pendientes'] = df['Unnamed: 7']

# configurar la fecha actual 
meses = {
    1: "enero",2: "febrero",3: "marzo",4: "abril",5: "mayo",6: "junio",
    7: "julio",8: "agosto",9: "septiembre",10: "octubre",11: "noviembre",12: "diciembre"
}

hoy = datetime.now()
fecha_actual = f"{hoy.day} {meses[hoy.month]} de {hoy.year}"

# columnas adicionales 
df['dia_cita'] = dia_cita
df['mes_cita'] = mes_cita
df['año_cita'] = 2025
df['administrador'] = "Geovanny Callejas Acevedo"
df['fecha actual'] = fecha_actual
df['lugar'] = lugar



df = df[['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','Nombre estudiante','acudiente','grado','pensiones pendientes','deuda total']]

# ====================================================================================================
# 🔥 Agrupar por estudiante y concatenar las pensiones pendientes
# Esto agrupa el DataFrame df en función de varias columnas:

# ['Nombre estudiante', 'acudiente', 'grado', 'deuda total']
#     Significa que todas las filas que tengan el mismo estudiante, con el mismo acudiente, grado y deuda total, 
#     serán agrupadas en un solo grupo.
# as_index=False evita que las columnas de agrupación pasen a ser el índice del DataFrame.

# 2. .agg({...})

# La función .agg() (de aggregate) nos deja especificar qué operación aplicar a cada columna del grupo.

# En este caso, solo nos interesa qué hacer con la columna pensiones pendientes.

# {'pensiones pendientes': lambda x: ... }


# Aquí decimos:

# "Para cada grupo, toma la columna pensiones pendientes y aplícale la función lambda que definimos".

# 🔹 3. lambda x: " - ".join(...)

# El parámetro x es la serie de valores de la columna pensiones pendientes dentro de cada grupo.

# Ejemplo: para "Sayago Acevedo Isabella", x sería:

# ["Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio"]


# Queremos unirlos en una sola celda.

# Para eso usamos:

# " - ".join([...])


# 👉 Une los elementos de la lista con " - " como separador.
# Resultado:

# "Febrero - Marzo - Abril - Mayo - Junio - Julio"

# 🔹 4. sorted(set(x), key=list(x).index)

# Esto es un pequeño truco 👇

# set(x) → elimina duplicados si había pensiones repetidas.

# ["Marzo", "Marzo", "Abril"] → {"Marzo", "Abril"}

# list(x).index como key → mantiene el orden original en el que aparecían en el DataFrame.

# sorted(set(x), key=list(x).index) devuelve la lista sin duplicados y en el mismo orden.

# Ejemplo:

# x = ["Marzo", "Febrero", "Marzo", "Abril"]
# sorted(set(x), key=list(x).index)
# # ['Marzo', 'Febrero', 'Abril']

# ====================================================================================================================

df_final = df.groupby(
    ['dia_cita', 'mes_cita','año_cita','administrador','fecha actual','lugar','Nombre estudiante', 'acudiente', 'grado', 'deuda total'], as_index=False
).agg({
    'pensiones pendientes': lambda x: " - ".join(sorted(set(x), key=list(x).index))
})
 

# Exportar a Excel
ruta_salida = "./morosos.xlsx"
df_final.to_excel(ruta_salida, index=False)
print(f"Archivo Excel creado en: {ruta_salida}")

 