import pandas as pd
import os
from datetime import datetime, timedelta, time
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Caminhos base
CAMINHO_BASE = os.path.dirname(__file__)
CAMINHO_PASTA_APLICACOES = os.path.join(CAMINHO_BASE, "Aplicações")

relatorio = []
semana_extraida = None  # Será preenchida com a última semana identificada

# Funções auxiliares
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

# Percorrer todas as aplicações
for nome_aplicacao in os.listdir(CAMINHO_PASTA_APLICACOES):
    pasta = os.path.join(CAMINHO_PASTA_APLICACOES, nome_aplicacao)
    if not os.path.isdir(pasta):
        continue

    arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]
    if not arquivos:
        continue

    caminho_arquivo = os.path.join(pasta, arquivos[0])

    try:
        df = pd.read_excel(caminho_arquivo)

        if "Semana" not in df.columns or "Horas" not in df.columns:
            print(f"⚠️ Planilha de '{nome_aplicacao}' não tem colunas 'Semana' e 'Horas'")
            continue

        ultima_semana = sorted(df["Semana"].dropna().unique())[-1]
        if not semana_extraida:
            semana_extraida = ultima_semana  # Armazena para nome do arquivo

        df_filtrado = df[df["Semana"] == ultima_semana]
        df_filtrado["Horas"] = df_filtrado["Horas"].apply(converter_para_timedelta)
        horas_total = df_filtrado["Horas"].sum()

        relatorio.append({
            "Aplicação": nome_aplicacao,
            "Semana": ultima_semana,
            "Horas na semana": formatar_horas(horas_total)
        })

        print(f"✅ {nome_aplicacao}: {formatar_horas(horas_total)} na semana {ultima_semana}")

    except Exception as e:
        print(f"❌ Erro em '{nome_aplicacao}': {e}")

# Gerar nome do arquivo com base na semana
if semana_extraida:
    nome_semana_formatado = semana_extraida.replace("/", "-").replace(" ", "")
    nome_arquivo = f"Relatorio_Horas_{nome_semana_formatado}.xlsx"
else:
    nome_arquivo = "Relatorio_Horas_Semana.xlsx"  # fallback

CAMINHO_RELATORIO = os.path.join(CAMINHO_BASE, nome_arquivo)

# Criar e salvar DataFrame
df_resultado = pd.DataFrame(relatorio)
df_resultado.to_excel(CAMINHO_RELATORIO, index=False)

# Autoajuste e centralização com openpyxl
wb = load_workbook(CAMINHO_RELATORIO)
ws = wb.active

for coluna in ws.columns:
    max_largura = 0
    for celula in coluna:
        celula.alignment = Alignment(horizontal='center', vertical='center')
        if celula.value:
            max_largura = max(max_largura, len(str(celula.value)))
    letra_coluna = celula.column_letter
    ws.column_dimensions[letra_coluna].width = max_largura + 2

wb.save(CAMINHO_RELATORIO)
print(f"\n📄 Relatório final salvo com nome: {nome_arquivo}")
