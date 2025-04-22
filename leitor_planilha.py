import pandas as pd

class LeitorPlanilha:
    @staticmethod
    def encontrar_linha_cabecalho(caminho_arquivo):
        df_preview = pd.read_excel(caminho_arquivo, sheet_name="Horas-Trabalhada", header=None, nrows=10)
        for idx, row in df_preview.iterrows():
            valores = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            if "semana" in valores and "horas" in valores:
                return idx
        return None

    @staticmethod
    def carregar_planilha(caminho_arquivo, linha_header):
        return pd.read_excel(caminho_arquivo, sheet_name="Horas-Trabalhada", header=linha_header)
