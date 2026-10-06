import numpy as np
import os
import seaborn as sns # type: ignore
import pandas as pd # type: ignore
import matplotlib.pyplot as plt # type: ignore
from sklearn.linear_model import LinearRegression # type: ignore



def separator(file_path:str) -> str:

    """Devuelve el tipo de separador que delimita el archivo csv.
    Esta función ha sido creada porque depende del fichero, tienen 
    un tipo de separador u otro.

    Args:
        file_path (str): dirección válida del archivo a explorar

    Returns:
        str: Tipo de separador, será "," ó ";" por lo general. Por defecto devuelve ","
    """

    with(open(file_path, mode="r", encoding="utf-8-sig")) as f:
        header = f.readline()

    return ";" if ";" in header else ","

def create_compare_linear_interpolation(df:pd.DataFrame, col:str) -> pd.Series:
    """Realiza una interpolación lineal y a continuación muestra los nuevos 
    puntos que se han creado frente a los originales.

    Args:
        df (pd.DataFrame): dataframe sobre el que operar
        col (str): columna que se quiere imputar

    Returns:
        pd.Series: valores con imputación actualizados
    """

    df_imputed = df[col].interpolate("linear")

    new_data = df_imputed.loc[df[col].isna()]

    fig = plt.figure(1, figsize=(8, 6))
    plt.scatter(x = df.index, y = df[col], label="original data")
    plt.scatter(x = new_data.index, y = new_data, label="new_points", color="red")
    plt.grid(), plt.legend()
    plt.show()

    return df_imputed