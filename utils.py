import pandas as pd


def convert_string_to_float(numero_string):
    numero_string = numero_string.replace(".", "").replace(",", ".")
    return float(numero_string)