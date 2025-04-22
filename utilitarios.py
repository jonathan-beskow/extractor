import re
import pandas as pd
from datetime import timedelta, time

def normalizar_semana(valor):
    if not isinstance(valor, str):
        valor = str(valor)
    valor = valor.strip()
    valor = valor.replace("–", " a ").replace("-", " a ")
    valor = re.sub(r"\s+", " ", valor)
    valor = re.sub(r"/\d{4}", "", valor)
    return valor.strip()

def converter_para_timedelta(valor):
    if isinstance(valor, time):
        return timedelta(hours=valor.hour, minutes=valor.minute, seconds=valor.second)
    try:
        return pd.to_timedelta(valor)
    except:
        return pd.NaT

def formatar_horas(td):
    if pd.isnull(td):
        return ""
    total_segundos = int(td.total_seconds())
    horas = total_segundos // 3600
    minutos = (total_segundos % 3600) // 60
    segundos = total_segundos % 60
    return f"{horas:02}:{minutos:02}:{segundos:02}"
