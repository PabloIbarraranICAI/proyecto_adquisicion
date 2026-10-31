"""
PROYECTO COLABORATIVO - Adquisición y Visualización de Datos (Grado IMAT, Comillas ICAI, 2026/27)
BLOQUE 1: Recopilación, limpieza y tratamiento de datasets del Mundial 2026

Autores: Antonio Díaz Ballesteros y Pablo Ibarrarán Olivares
Genera: partidosWC26.xlsx, eventosWC26.xlsx y jugadoresWC2026.xlsx (carpeta RUTA_SALIDA)

Uso: ajustar RUTA_DATOS (carpeta con los .csv originales) y ejecutar:  python bloque1_limpieza_datos.py
"""


# ============================================================================
# APARTADO 1: LIMPIEZA DE LOS DATOS Y PREPARACAÓN DE LOS DF.
# ============================================================================


# ---- Importación de los paquetes -------------------------------------------

import pandas as pd
import numpy as np
import re
from sklearn.linear_model import LinearRegression

RUTA_DATOS = "./datasetsWC26/"     #carpeta con los ficheros csv originales
RUTA_SALIDA = "./datasets_limpios/"             #carpeta donde se exportan los .xlsx finales limpios


# ---- Función auxiliar para detectar el separador de cada csv (detecta si cada csv usa ';' o ',') ----

def separator(file_path: str) -> str:
    """Devuelve el separador (";" o ",") que delimita el archivo csv, leyendo su cabecera.
    Cada fichero del paquete usa un separador distinto, por eso se detecta automáticamente.
    Por defecto devuelve ",".
    """
    with open(file_path, mode="r", encoding="utf-8-sig") as f:
        header = f.readline()
    if ";" in header:
        return ";" 
    else:
        return ","

#definimos una función que realice un analisis básido inicial de los df

def analisis_inicial(df: pd.DataFrame):
    print(f'\n -> La forma del df es : {df.shape}')
    print(f'\n -> Las columnas son {df.columns.to_list()}')
    print(f'\n -> los formatos de las columnas son:')
    print(df.dtypes)
    print(f'\n -> Número de valores nulos: ')
    print(df.isna().sum())
    print(f'\n -> Número de lineas duplicadas {df.duplicated().sum()}')

    print(f' \n-> Análisis general con info: ')
    print(df.info())

    print(f' \n-> Métricas generales con describe {df.describe(include=["number"])}')

    print(f'\n-> Vista previa del df cabecera:')
    print(df.head())

    print(f'\n-> Vista previa del df final:')
    print(df.tail())


# ---- PARTIDOS --------------------------------------------------------------
#
# empezamos importando matches_mundial2026.csv que servirá como base para el df partidos

partidos_raw = pd.read_csv(fr"{RUTA_DATOS}matches_mundial2026.csv", sep = separator(fr"{RUTA_DATOS}matches_mundial2026.csv"), na_values=["N/A"])

print(partidos_raw)

# observamos que hay 110 lineas, mientras que solo hubo 104 partidos en el mundial, por lo que habrá que evr si hay duplicados o lineas erroneas

print(f'ANÁLISIS INICIAL DEL DF PARTIDOS:')
analisis_inicial(partidos_raw)

# eliminamos directamente las columnas que no vamos a necesitar.
# Pero conservamos result_type que nos permitirá crear más adelante la columna prorroga, la eliminaremos después


partidos = partidos_raw.copy()
partidos = partidos.drop(columns=["kickoff_time_utc", "status", "home_xg", "away_xg", "player_of_the_match_id", "capacidad_estadio_pct", "clima" ])

mapeo_a_español_1_partidos={"match_id": "index_partido", "stage_id": "ronda", "venue_id": "estadio", "home_team_id": "seleccion_local", "away_team_id": "seleccion_visitante", "home_score": "marcador_local", "away_score": "marcador_visitante", "home_penalty_score": "penalty_local", "away_penalty_score": "penalty_visitante", "referee_id": "arbitro_partido"}
partidos = partidos.rename(columns = mapeo_a_español_1_partidos)
partidos.set_index("index_partido")
print(partidos.columns)

# Gestión de duplicados


print(partidos[partidos.duplicated()])

# Gestión de duplicados: hay 6 lineas duplicadas que eliminar para que ya solo queden las 104 lineas de los 104 partidos

print(partidos.drop_duplicates(inplace=True))

#comprobar que todos los partidos del 1 al 104 estén en el df
indices_partidos = partidos["index_partido"].to_list()
algun_error = False

for i in range(1, 105):
    if i not in indices_partidos:
        print(f'El partido {i} no está en el df')
        algun_error = True

if not algun_error:
    print(f'Todos, los 104 los partidos, aparecen en el df partidos. ')

# Gestión de los formatos erroneos en las fechas (expresiones regulares regex)

### bloque para la captura y conversion de las fechas
patron1 = re.compile(r"(\d{2})[/\.](\d{2})[/\.](2026)")               # DD/MM/YYYY  o  DD.MM.YYYY
patron_us = re.compile(r"(\d{2})-(\d{2})-(2026)")                     # MM-DD-YYYY
patron2 = re.compile(r"(\d{2})[\s-](\w{3,4})[\s-](\d{2,4})")          # 13-jun-26  /  12 June 2026
patron3 = re.compile(r"\s*(\w{3,4})\s+(\d{2}),?\s+(2026)\s*")         # June 13, 2026  /  Jun 14 2026
patron4 = re.compile(r"(\d{1,2})\sde\s(\w+)\sde\s(2026)")             # 14 de junio de 2026
patron5 = re.compile(r"(2026)(\d{2})(\d{2})")                         # 20260614
patron6 = re.compile(r"(2026)-(\d{2})-(\d{2})T19:00:00Z")             # 2026-06-13T19:00:00Z

def convertir_mes(mes_sucio):
    
    mapeo_de_meses = {"june": "06", "junio": "06", "jun": "06", "jul": "07", "july": "07", "julio": "07"}
    return mapeo_de_meses.get(mes_sucio.lower())

def aplicar_formato(row):
    fecha = row["date"]

    match1 = re.fullmatch(patron1, fecha)
    match_us = re.fullmatch(patron_us, fecha)
    match2 = re.fullmatch(patron2, fecha)
    match3 = re.fullmatch(patron3, fecha)
    match4 = re.fullmatch(patron4, fecha)
    match5 = re.fullmatch(patron5, fecha)
    match6 = re.fullmatch(patron6, fecha)


    if match1:
        dia, mes, año = match1.groups()
        fecha_limpia = f"2026-{mes}-{dia}"

    if match_us:
        mes, dia, año = match_us.groups()
        fecha_limpia = f"2026-{mes}-{dia}"

    elif match2:
        dia, mes, año = match2.groups()
        mes = convertir_mes(mes)
        fecha_limpia = f"2026-{mes}-{dia}"

    elif match3:
        mes, dia, año = match3.groups()
        mes = convertir_mes(mes)
        fecha_limpia = f"2026-{mes}-{dia}"

    elif match4:
        dia, mes, año = match4.groups()
        mes = convertir_mes(mes)
        fecha_limpia = f"2026-{mes}-{dia}"

    elif match5:
        año, mes, dia = match5.groups()
        fecha_limpia = f"2026-{mes}-{dia}"

    elif match6:
        año, mes, dia = match6.groups()
        fecha_limpia = f"2026-{mes}-{dia}"

    return fecha_limpia

partidos["fecha_partido"] = partidos.apply(aplicar_formato, axis = 1)
partidos["fecha_partido"] = pd.to_datetime(partidos["fecha_partido"])
partidos.drop(columns=["date"], inplace = True)
print(partidos)

# Hay fechas que quedan fuera del periodo del mundial (11 jun - 19 jul 2026): son fechas escritas como MM/DD/YYYY que el patrón DD/MM ha leído con día y mes intercambiados. Como son solo 3, las podemos corregir a mano.

#eliminar los registros que no corresponden a las fechas del mundial
fecha_inicio = pd.to_datetime('2026-06-11')
fecha_fin = pd.to_datetime('2026-07-19')
partidos['fecha_partido'] = partidos['fecha_partido']

fuera_de_rango = partidos.loc[(partidos["fecha_partido"] < fecha_inicio) | (partidos["fecha_partido"] > fecha_fin)] 

print(fuera_de_rango)

# revisamos y corregimos esas 3 fechas

#arreglo a mano los 3 dias que se han leido mal y se han intercambiado día y mes

partidos.loc[2,"fecha_partido"] = pd.to_datetime("2026-06-12")
partidos.loc[86,"fecha_partido"] = pd.to_datetime("2026-07-04")
partidos.loc[98,"fecha_partido"] = pd.to_datetime("2026-07-11")

fuera_de_rango = partidos.loc[(partidos["fecha_partido"] < fecha_inicio) | (partidos["fecha_partido"] > fecha_fin)] 

print(fuera_de_rango)

print(partidos)

# FORMATEAR



partidos["penalty_local"] = partidos["penalty_local"].astype("Int64")
partidos["arbitro_partido"] = partidos["arbitro_partido"].astype("Int64")
partidos["penalty_visitante"] = partidos["penalty_visitante"].astype("Int64")

print(partidos.dtypes)


# ---- importamos ahora los arbitros de index_referees.csv para asociarlos a los partidos ----

arbitros_raw = pd.read_csv(fr"{RUTA_DATOS}index_referees.csv", sep = separator(fr"{RUTA_DATOS}index_referees.csv"))
print(arbitros_raw.head(5))

analisis_inicial(arbitros_raw)

print(f'En este df hay {arbitros_raw.duplicated().sum()} duplicados. Los eliminamos.')
print(arbitros_raw[arbitros_raw.duplicated()])
arbitros = arbitros_raw.drop_duplicates()

# hacemos un merge para cruzar los nombres de los arbitros con los partidos

partidos = pd.merge(partidos.reset_index(), arbitros[["referee_id", "name"]], how='left', left_on="arbitro_partido", right_on="referee_id").set_index("index_partido")
partidos.drop(columns = ["index", "referee_id", "arbitro_partido"], inplace = True)
partidos = partidos.rename(columns = {"name": "arbitro_partido"})
print(f'Tras el merge, comprobamos que se ha realizado bien porque sigue habiendo {partidos.shape[0]} lineas.')

# el árbitro de ese partido no figura en los datos originales (referee_id vacío) y no hay forma de recuperarlo, así que lo marcamos explícitamente como `Desconocido` en lugar de dejar un nulo.

print(partidos[partidos["arbitro_partido"].isna()])

partidos["arbitro_partido"] = partidos["arbitro_partido"].fillna("Desconocido")

print(partidos)


# ---- importamos ahora las fases de tournament_stages.csv -------------------

tournament_stages_raw = pd.read_csv(fr"{RUTA_DATOS}tournament_stages.csv", sep = separator(fr"{RUTA_DATOS}tournament_stages.csv"))
print(tournament_stages_raw.head(9))

analisis_inicial(tournament_stages_raw)

print(f'En este df hay {tournament_stages_raw.duplicated().sum()} duplicados. Los eliminamos.')
print(tournament_stages_raw[tournament_stages_raw.duplicated()])
tournament_stages = tournament_stages_raw.drop_duplicates()
tournament_stages["stage_name"] = tournament_stages["stage_name"].str.strip()
print(tournament_stages)

#sustituimos el id de la fase por su nombre
partidos = pd.merge(partidos.reset_index(), tournament_stages[["stage_id", "stage_name"]], how='left', left_on="ronda", right_on="stage_id").set_index("index_partido")
partidos.drop(columns = ["stage_id", "ronda"], inplace = True)
partidos = partidos.rename(columns = {"stage_name": "ronda"})
print(f'Tras el merge, comprobamos que se ha realizado bien porque sigue habiendo {partidos.shape[0]} lineas.')
print(partidos)


# ---- importamos ahora los estadios de worldcup_venues.csv ------------------

worldcup_venues_raw = pd.read_csv(fr"{RUTA_DATOS}worldcup_venues.csv", sep = separator(fr"{RUTA_DATOS}worldcup_venues.csv"))
print(worldcup_venues_raw)

analisis_inicial(worldcup_venues_raw)

print(f'En este df hay {worldcup_venues_raw.duplicated().sum()} duplicados. Los eliminamos.')
print(worldcup_venues_raw[worldcup_venues_raw.duplicated()])
worldcup_venues = worldcup_venues_raw.drop_duplicates()
print(worldcup_venues)

#observamos que USA está escrito con mayusculas y con minusculas
print(worldcup_venues["country"].unique())
print(f'corregimos escritura')
worldcup_venues["country"] = worldcup_venues["country"].replace("usa", "USA")
print(worldcup_venues["country"].unique())

partidos = pd.merge(partidos.reset_index(), worldcup_venues[["venue_id", "stadium_name", "city", "country"]], how='left', left_on="estadio", right_on="venue_id").set_index("index_partido")
partidos.drop(columns = ["venue_id", "estadio"], inplace = True)
partidos = partidos.rename(columns = {"city": "ciudad_partido", "country": "pais_partido", "stadium_name": "estadio"})
print(f'Tras el merge, comprobamos que se ha realizado bien porque sigue habiendo {partidos.shape[0]} lineas.')
print(partidos)

# Columna `prorroga`: usamos `result_type` (`AET` = ganó en la prórroga, `Penalties` = llegó a penaltis, que implica prórroga previa). Hay un partido que sorprende, el partdo 84 entre protugal y croacia, que sin tener prorroga tiene un evento en el 103, esto se debe a un tiempo añadido de 20 minutos (hubo muchas polemicas en el partido), pero es correcto.

partidos["prorroga"] = ((partidos["result_type"] == "AET") |(partidos["result_type"] == "Penalties"))
partidos["prorroga"] = partidos["prorroga"].map({True: "SI", False: "NO"})
partidos.drop(columns=["result_type"], inplace = True)
print(f'\nPartidos con prórroga: {(partidos["prorroga"] == "SI").sum()}')


# ---- JUGADORES -------------------------------------------------------------

# importamos los datos de los jugadores 2026player_stats.csv que servirá de base para el df jugadores

jugadores_raw = pd.read_csv(fr"{RUTA_DATOS}2026player_stats.csv", sep = separator(fr"{RUTA_DATOS}2026player_stats.csv"))
print(jugadores_raw)

analisis_inicial(jugadores_raw)
#no hay duplicados

print(jugadores_raw.columns)

# eliminamos directamente las columnas que no vamos a necesitar
# quitamos tambien `player_name`porque vienen los nombrs con una codificacion mal escrita (`JosŽ Raœl Rangel`), preferiremos quedarnos con la del csv 2026squads_and_players.csv (`José Raúl Rangel`)



jugadores = jugadores_raw.copy()

jugadores = jugadores.drop(columns= ["player_name", "shots_on_target","matches_started","penalty_goals","own_goals","clean_sheets", "Unnamed: 18"])
mapeo_a_español_2_jugadores={"player_name": "nombre_jugador", "position": "posicion", "matches_played": "partidos_jugados", "minutes_played": "minutos_jugados", "goals": "goles", "assists": "asistencias", "shots": "tiros", "yellow_cards": "tarjetas_amarillas", "red_cards": "tarjetas_rojas", "saves": "paradas", "goals_conceded":"goles_concedidos", "average_rating":"calificacion_media"}
jugadores = jugadores.rename(columns = mapeo_a_español_2_jugadores)

print(jugadores)


#ninguno de los player_id es nulo lo cual es util para futuros merge
print(jugadores.isna().sum())


# ---- importamos los datos de los jugadores 2026squads_and_players.csv ------

squads_players_raw = pd.read_csv(fr"{RUTA_DATOS}2026squads_and_players.csv", sep = separator(fr"{RUTA_DATOS}2026squads_and_players.csv"))
print(squads_players_raw)

analisis_inicial(squads_players_raw)
#no hay duplicados ni nulos en los valores identificativos y coincide el tamaño con el número de jugadores que hobo en el mundial

print(jugadores)


jugadores = pd.merge(jugadores, squads_players_raw[["player_id", "player_name", "club_team", "market_value_eur", "caps"]], how='left', on="player_id")


jugadores = jugadores.rename(columns = {"club_team":"club_jugador", "market_value_eur": "valor_mercado", "player_name": "nombre_jugador", "caps": "partidos_seleccion"})
print(f'Tras el merge, comprobamos que se ha realizado bien porque sigue habiendo {jugadores.shape[0]} lineas.')

print(jugadores)


# ---- importamos ahora las selecciones de 2026nationalteams.csv -------------

selecciones_raw = pd.read_csv(fr"{RUTA_DATOS}2026nationalteams.csv", sep = separator(fr"{RUTA_DATOS}2026nationalteams.csv"))
print(selecciones_raw)

analisis_inicial(selecciones_raw)
#no hay duplicados

print(selecciones_raw[selecciones_raw.duplicated()])

#eliminamos filas duplicadas
selecciones = selecciones_raw[["team_id", "team_name"]].drop_duplicates()
print(selecciones.shape)

#hay 7 selecciones cuyo team id es nulo NaN
print(selecciones[selecciones["team_id"].isna()])

#eliminamos con una mascara aquellos equipos que no tienen id
mask = ~selecciones["team_id"].isna()
selecciones = selecciones.loc[mask, ["team_id", "team_name"]]
print(selecciones.shape)

print(selecciones.head(10))

#quedan ya solo 48 lineas validad sin NaN en sus id y sin duplicados, coincide con el numero real de selecciones en el mundial, y esta listo apra el merge
print(selecciones.shape)

jugadores = pd.merge(jugadores, selecciones, how='left', on="team_id")

jugadores.drop(columns = ["team_id",], inplace = True)
jugadores.rename(columns = {"team_name": "seleccion_nacional"}, inplace=True)

print(f'Tras el merge, comprobamos que se ha realizado bien porque sigue habiendo {jugadores.shape[0]} lineas.')
jugadores.set_index("nombre_jugador", inplace = True )
print(jugadores)

print(partidos)

#añadir la columna de seleccion local a partidos
partidos = pd.merge(partidos.reset_index(), selecciones, how = "left", left_on="seleccion_local", right_on="team_id").set_index("index_partido")
partidos.drop(columns = ["seleccion_local", "team_id"], inplace=True)
partidos.rename(columns = {"team_name":"seleccion_local"}, inplace=True)

print(partidos)

#añadir la columna de seleccion visitante a partidos
partidos = pd.merge(partidos.reset_index(), selecciones, how = "left", left_on="seleccion_visitante", right_on="team_id").set_index("index_partido")
partidos.drop(columns = ["seleccion_visitante", "team_id"], inplace=True)
partidos.rename(columns = {"team_name":"seleccion_visitante"}, inplace=True)

print(partidos)

#ya están todos los datos que necesitamos en el df jugadores, ahora tenemos que formatearlo
print(jugadores.dtypes)

for columna in ["goles", "asistencias", "tiros", "tarjetas_amarillas", "tarjetas_rojas", "paradas", "goles_concedidos"]:
    jugadores[columna] = jugadores[columna].astype('Int64')

print(jugadores.dtypes)


# ---- Revisión de valores anómalos y de nulos en jugadores ------------------

#chequeo de la validez de los valores (que sean positivos)
for columna in ["partidos_jugados", "minutos_jugados", "goles", "asistencias", "tiros", "tarjetas_amarillas", "tarjetas_rojas", "paradas", "goles_concedidos"]:
    print(f'\n-> {columna}')
    print(f"\tValores_nulos: {jugadores[columna].isna().sum()}")

    if jugadores[jugadores[columna]<0].shape[0] > 0:
        print(f'\tValores negativos : {jugadores[jugadores[columna]<0]}')
    else:
        print(f'\tValores negativos : 0')

#checkeo de validez de las calificacion
print(jugadores["calificacion_media"].isna().sum())
mascara=(jugadores["calificacion_media"]<0) | (jugadores["calificacion_media"]>10)
print(jugadores[mascara])

# no hay ningún jugador con notas fuera del rango 0-10 pero sí bastantes que no tienen calificacion

#chequeo de la validez de las posiciones: todas son o portero o defensa. medio o delamntero
print(jugadores["posicion"].unique())

selecciones_posibles = selecciones["team_name"].to_list()
error_seleccion = False

for seleccion in jugadores["seleccion_nacional"]:
    if seleccion not in selecciones_posibles:
        print(f'La seleccion {seleccion} no está entre las permitidas')
        error_seleccion = True

if not error_seleccion:
    print(f'No hay selecciones prohibidas')

# Penaltis de las tandas: los valores vienen multiplicados por 10 (20, 30, 40 en lugar de 2, 3, 4), lo comprobamos y los dividimos. Además, el enunciado pide `0 si no aplica`, así que los partidos sin tanda de penaltis (NaN) pasan a 0.

print(partidos["penalty_local"].unique())

print(partidos["penalty_visitante"].unique())

partidos["penalty_visitante"] = (partidos["penalty_visitante"] / 10).astype("Int64")
partidos["penalty_local"] = (partidos["penalty_local"] / 10).astype("Int64")
partidos["penalty_visitante"] = partidos["penalty_visitante"].fillna(0)
partidos["penalty_local"] = partidos["penalty_local"].fillna(0)

print(partidos["penalty_visitante"].unique())

print(partidos["penalty_local"].unique())

# hemos comprobado: el formato de todas las columnas, que los valores de las estadisticas sean factibles, no sean negativos, para las calificaciones, que estén entre 0-10, y que las posiciones de los jugadores estén entre las permitidas, y que las selecciones de los jugadores estén también entre las permitidas
# hemos limpiado duplicados, hecho los merge pertinentes, convertido a los formatos que queriamos etc


# ---- curiosidad, los 20 futbolistas mejor valorados del mundial y los 20 que mas goles han metido ----

print(jugadores["calificacion_media"].sort_values(ascending=False).head(20))

print(jugadores["goles"].sort_values(ascending=False).head(20))


# ---- importamos ahora 2026match_register.csv para procesar los eventos en los partidos ----

eventos_raw = pd.read_csv(fr"{RUTA_DATOS}2026match_register.csv", sep = separator(fr"{RUTA_DATOS}2026match_register.csv"))
print(eventos_raw)

print(f'ANÁLISIS INICIAL DEL DF PARTIDOS:')
analisis_inicial(eventos_raw)

eventos = eventos_raw.copy()
eventos = eventos.drop(columns = ["team_id"])
eventos.drop_duplicates(inplace=True)
eventos.rename(columns={"event_id": "index_evento", "match_id": "index_partido", "minute": "minuto_evento", "event_type": "tipo_evento"}, inplace=True)
print(eventos)

eventos["index_partido"] = eventos["index_partido"].astype("Int64")
eventos["minuto_evento"] = eventos["minuto_evento"].astype("Int64")
eventos["player_id"] = eventos["player_id"].astype("Int64")

#eliminar los eventos asociados a partidos que no existen
mascara = (eventos["index_partido"] >= 1) & (eventos["index_partido"] <= 104)
eventos = eventos.loc[mascara]
print(eventos)

print(eventos[(eventos["minuto_evento"] < 0)])

print(eventos[(eventos["minuto_evento"] > 120)])

# observamos que en el minuto 121 se apuntan las tandas de penaltis, por lo que tenemos quue ocnservar ese minuto también!!!

#eliminar los eventos asociados a minutos más allá del 121 (no existen tampoco) OJO QUE LOS DEL 121 son los penaltis
print(eventos.shape)
mascara = (
    (eventos["minuto_evento"] >= 0) & (eventos["minuto_evento"] <= 121)
) | eventos["minuto_evento"].isna()
eventos = eventos.loc[mascara]
print(eventos.shape)

# se han eliminado los 10 eventos que correspondian a minutos negativos y más alla del 121 (300, 500, 999 centinela etc)

print(eventos.dtypes)
print(f'Convertimos los tipos a int')
eventos["index_partido"] = eventos["index_partido"].astype('Int64')
eventos["minuto_evento"] = eventos["minuto_evento"].astype('Int64')
eventos["player_id"] = eventos["player_id"].astype('Int64')
print(eventos.dtypes)

#asegurarse de que los jugadores presentes en la tabla existen, es decir, que su index esté entre 1 y 1248  -> mo hay ninguno: es correcto.
print(eventos[eventos["player_id"]<1])

print(eventos[eventos["player_id"]>1248])

#nos damos cuenta que las revisiones del VAR vienen dadas con los str diferentes, y las unificamos en 1
print(eventos["tipo_evento"].value_counts())
print(f'\n Vamos a unificar VAR Review y Review VAR en una única.')
eventos["tipo_evento"] = eventos["tipo_evento"].str.strip().replace("Review VAR", "VAR Review")
print(eventos["tipo_evento"].value_counts())

print(f'Ya tenemos limpio y validado el df de match_register')
print(eventos)

#hacemos un merge para rellenar con los datos de nombre del jugador y seleccion
eventos = pd.merge(eventos, jugadores.reset_index()[["player_id", "nombre_jugador", "seleccion_nacional"]])
eventos.drop(columns=["player_id"], inplace=True)
eventos.rename(columns={"seleccion_nacional":"seleccion_evento"}, inplace=True)
print(eventos)

print(jugadores)

# verificacion de los NA de columnas de estadisticas individuales, en muchas se PODRÍA poner los nan a 0, pero no lo vamos a hacer de momento porque no queremos perder la información (por ejemplo un jugador de campo va a tener siempre nan paradas porque solo puede hacer paradas un portero, realmente no es que su valor de la estadistica de paradas sea 0, porque nopuede hacerlas, sino nan; lo mismo pasa para los jugadores que no estñan jugando, no tienen un valor 0, porque estan en el banquillo, sino nan)

print(jugadores["tarjetas_amarillas"].value_counts(dropna=False))
print(f'Tras comprobar las bases oficiales del mundial, y como es lógico, encaja que aquellos registros de tarjetas amarillas que aparecen como NA sean 0')

print(jugadores["tarjetas_rojas"].value_counts(dropna=False))
print(f'Tras comprobar las bases oficiales del mundial, y como es lógico, encaja que aquellos registros de tarjetas rojas que aparecen como NA sean 0.')

print(jugadores["asistencias"].value_counts(dropna=False))
print(jugadores["paradas"].value_counts(dropna=False))
print(jugadores["goles_concedidos"].value_counts(dropna=False))
print(f'Todos ellos encajan con los datos reales de la web de la FIFA y podríamos poner a 0 los nan cuando sea necesario para la imputacion por regresion')


# ---- Imputación de los datos que están como NAN ----------------------------

print(jugadores.isna().sum())

print(eventos.isna().sum())


# ---- vamos a imputar el valor_mercado por la media de la seleccion del jugador, calificacion_media por regresión, minuto_evento por interpolación ----

# imputamos los valores nulos del valos_mercado de los jugadores por la media de los valores de mercado de sus compañeros de selección

# un valor de mercado 0.0 no es real, es un dato ausente: lo pasamos a NaN para que también se impute
jugadores["valor_mercado"] = jugadores["valor_mercado"].replace(0.0, np.nan)

jugadores["valor_mercado"] = jugadores["valor_mercado"].fillna(jugadores.groupby("seleccion_nacional")["valor_mercado"].transform("mean"))
print(jugadores["valor_mercado"].isna().sum())

# Calificaciones iguales a 0.0: hay 18 jugadores que solo han jugado 1 o 2 minutos. Una nota de 0 sobre 10 no es una valoración real
# (porque la media del torneo es de aprox 6.6) sino un valor centinela por no tener minutos suficientes, y distorsionaría tanto la regresión como la
# calificación media de cada selección. Por eso se consideran anómalos: los pasamos a NaN para que se imputen por regresión como el resto.
print(f'Jugadores con calificacion_media = 0.0 (anómalos): {(jugadores["calificacion_media"] == 0).sum()}')
jugadores.loc[jugadores["calificacion_media"] == 0, "calificacion_media"] = np.nan
print(f'Nulos en calificacion_media antes de imputar: {jugadores["calificacion_media"].isna().sum()}')

# Calificacion_media: imputamos los datos faltantes con Imputación por Regresión basada en la correlación con las estadísticas de goles, asistencias, tarjetas
# amarillas y rojas, paradas o goles concedidos (Distiguimos por posición del jugador: no es lo mismo un delantero que un portero)

# PORTEROS

### hago la regresion de la calificacion media para los porteros GK

#generamos una copia de los datos y SOLO EN ESA COPIA sustotuyo los nan de las estadisticas personales por 0, que ya hemos visto que se podia en este caso pero luego en el df principal de jugadores seguiremos manteniedo es dato nan
datos_modelo = jugadores.copy()
columnas_de_entrenamiento = ["goles", "asistencias", "tarjetas_amarillas", "tarjetas_rojas", "paradas", "goles_concedidos"]
datos_modelo[columnas_de_entrenamiento] = (datos_modelo[columnas_de_entrenamiento].fillna(0)) #si no se rellenan los valores de las columnas de enyrenamiento por nulos (que en este caso sí se puede como vimos antes, explotara el modelo de regresión)

#crear los dartos de entrenamiento y entrenar
mascara_portero = (jugadores["posicion"]=="GK")

train = datos_modelo.loc[mascara_portero].copy()
train = train.dropna(subset=["calificacion_media"])

model = LinearRegression().fit(train[columnas_de_entrenamiento], train[["calificacion_media"]])

#imputar las calificaciones ausentes
mask_imputar = (mascara_portero & jugadores["calificacion_media"].isna())

predicciones = model.predict(datos_modelo.loc[mask_imputar, columnas_de_entrenamiento])
predicciones = np.clip(predicciones, 0, 10).round(2)
jugadores.loc[mask_imputar, "calificacion_media"] = predicciones

# RESTO DE POSICIONES (no les aplica ya ni paradas ni goles concedidios que es solo para los porteros)

print(jugadores["posicion"].unique())

### regresion de la calificacion media para el resto de posiciones (DEF, MID, FWD)

for posicion in ["DEF", "MID", "FWD"]:
    #generamos una copia de los datos y SOLO EN ESA COPIA sustotuyo los nan de las estadisticas personales por 0, que ya hemos visto que se podia en este caso pero luego en el df principal de jugadores seguiremos manteniedo es dato nan
    datos_modelo = jugadores.copy()
    columnas_de_entrenamiento = ["goles", "asistencias", "tarjetas_amarillas", "tarjetas_rojas"]
    datos_modelo[columnas_de_entrenamiento] = (datos_modelo[columnas_de_entrenamiento].fillna(0)) #si no se rellenan los valores de las columnas de enyrenamiento por nulos (que en este caso sí se puede como vimos antes, explotara el modelo de regresión)

    #crear los dartos de entrenamiento y entrenar
    mascara_posicion = (jugadores["posicion"]==posicion)

    train = datos_modelo.loc[mascara_posicion].copy()
    train = train.dropna(subset=["calificacion_media"])

    model = LinearRegression().fit(train[columnas_de_entrenamiento], train[["calificacion_media"]])

    #imputar las calificaciones ausentes
    mask_imputar = (mascara_posicion & jugadores["calificacion_media"].isna())

    predicciones = model.predict(datos_modelo.loc[mask_imputar, columnas_de_entrenamiento])
    predicciones = np.clip(predicciones, 0, 10).round(2)
    jugadores.loc[mask_imputar, "calificacion_media"] = predicciones

mascara_portero = (jugadores["posicion"]=="GK")
print(jugadores[mascara_portero])

for seleccion in jugadores["seleccion_nacional"].unique():
    num_porteros = len(jugadores.loc[((jugadores["posicion"]=="GK")&(jugadores["seleccion_nacional"]==seleccion))])
    print(seleccion, num_porteros)

print(jugadores.loc[((jugadores["posicion"]=="GK")&(jugadores["seleccion_nacional"]=="Egypt"))])

# como curiosidad, hemos observado que hay 48 selecciones y en teoria cada una debería llevar 3 porteros lo cual da un total de 144 porteros. Pero como observabamos que salían seleccionados 145 porteros, uno más de lo que debería, hemos buscado y es Egipto la unica seleccion que llevó 4 porteros en lugar de 3. Tras verificarlo con las fuentes oficiales del mundial, este dato es correcto.

print(jugadores)

# observamos que ya no queda ingun valor de calificacion_media nulo

print(jugadores["calificacion_media"].isna().sum())


# ---- Imputacion por interpolacion de los valores nulos de minuto_evento ----

print(eventos["minuto_evento"].isna().sum())

# vamos a interpolar los ausentes, pero agrupando segun el partido (que no se interpole un valor con los minutos de partido de dos partidos diferentes...)
# Para ello necesitamos que estén ordenados de manera creciente por index_partido y minuto_partido, que ya lo está

print(eventos)

print(eventos.head(20))

eventos["minuto_evento"] = eventos.groupby("index_partido")["minuto_evento"].transform(lambda x: x.interpolate())

print(eventos.head(20))

print(eventos["minuto_evento"].isna().sum())

# ATENCIÓN observamos que tras el interpolate solo queda 1 valor de minuto como nan, esto se debe a que es el primer evento del partido con indice 3, por lo que no tiene un valor previo en el mismo partido e interpolate no puede inventarse un valor de minuto previo.
# Aquí podemos decidir poner un bfill además del interpolate para rellenar ese último dato. Lo hacemos para que no quede ningún nulo, como pide el enunciado, PERO sabiendo que es una imputación muy poco precisa!!

eventos["minuto_evento"] = eventos.groupby("index_partido")["minuto_evento"].transform(lambda x: x.bfill())
eventos["minuto_evento"] = eventos["minuto_evento"].round().astype("Int64")   #el minuto es un entero, la interpolacion no puede dejar minutos decimales...
print(eventos["minuto_evento"].isna().sum())

print(partidos)


# ---- ya tenemos los 3 df pedidos: partidos, jugadores y eventos, listos para comenzar el proyecto! ----

jugadores.drop(columns=["player_id"], inplace=True)
#reordeno las columnas para que salgan igual que en el enunciado
jugadores = jugadores[["seleccion_nacional", "club_jugador", "partidos_seleccion", "posicion", "valor_mercado",
                       "partidos_jugados", "minutos_jugados", "goles", "asistencias", "tiros",
                       "tarjetas_amarillas", "tarjetas_rojas", "paradas", "goles_concedidos", "calificacion_media"]]
print(jugadores)

#reordeno las columnas para que salgan igual que en el enunciado
eventos = eventos.set_index("index_evento")[["minuto_evento", "index_partido", "nombre_jugador", "tipo_evento", "seleccion_evento"]]
print(eventos)


# ---- Guardo los dfs ya limpios en excel (.xlsx) para trabajar con ellos proximamente ----

#permite crear una carpeta donde guardar los df limpios
import os
os.makedirs(RUTA_SALIDA, exist_ok=True)

partidos.to_excel(fr"{RUTA_SALIDA}partidosWC26.xlsx", index=True)
jugadores.to_excel(fr"{RUTA_SALIDA}jugadoresWC2026.xlsx", index=True)
print(eventos.to_excel(fr"{RUTA_SALIDA}eventosWC26.xlsx", index=True))
