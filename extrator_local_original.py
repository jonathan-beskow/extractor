import pandas as pd
import os
from datetime import datetime, timedelta, time
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re

# Caminhos base
CAMINHO_BASE = os.path.dirname(__file__)
CAMINHO_PASTA_APLICACOES = os.path.join(CAMINHO_BASE, "Aplicações")

relatorio = []
nao_atualizadas = []

# Semana passada - cálculo com base na data atual
hoje = datetime.today()
segunda_semana_passada = hoje - timedelta(days=hoje.weekday() + 7)
sexta_semana_passada = segunda_semana_passada + timedelta(days=4)
semana_passada_formatada = f"{segunda_semana_passada.strftime('%d/%m')} a {sexta_semana_passada.strftime('%d/%m')}"

print("\n📅 Semana passada esperada (normalizada):", f"'{semana_passada_formatada}'\n")

# Funções auxiliares
def normalizar_semana(valor):
    if not isinstance(valor, str):
        valor = str(valor)
    valor = valor.strip()
    valor = valor.replace("–", " a ").replace("-", " a ")
    valor = re.sub(r"\s+", " ", valor)
    valor = re.sub(r"/\d{4}", "", valor)  # Remove anos como 2025
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

def encontrar_linha_cabecalho(caminho_arquivo):
    df_preview = pd.read_excel(caminho_arquivo, sheet_name="Horas-Trabalhada", header=None, nrows=10)
    for idx, row in df_preview.iterrows():
        valores = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if "semana" in valores and "horas" in valores:
            return idx
    return None

# Processar cada aplicação
# Processar cada aplicação
for nome_aplicacao in os.listdir(CAMINHO_PASTA_APLICACOES):
    pasta = os.path.join(CAMINHO_PASTA_APLICACOES, nome_aplicacao)
    if not os.path.isdir(pasta):
        continue

    arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]
    if not arquivos:
        continue

    print(f"\n📂 Processando aplicação: {nome_aplicacao}")
    arquivo_valido = None
    caminho_arquivo_valido = None
    linha_header_valida = None

    for arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta, arquivo)
        try:
            linha_header = encontrar_linha_cabecalho(caminho_arquivo)
            if linha_header is not None:
                df_teste = pd.read_excel(caminho_arquivo, sheet_name="Horas-Trabalhada", header=linha_header, nrows=1)
                if "Semana" in df_teste.columns and "Horas" in df_teste.columns:
                    arquivo_valido = arquivo
                    caminho_arquivo_valido = caminho_arquivo
                    linha_header_valida = linha_header
                    break
        except Exception:
            continue

    if not caminho_arquivo_valido:
        print(f"⚠️ Nenhum arquivo com aba 'Horas-Trabalhada' válida encontrado em '{nome_aplicacao}'.")
        nao_atualizadas.append({"Aplicação": nome_aplicacao, "Motivo": "Nenhum arquivo com aba válida encontrado"})
        continue

    print(f"📄 Arquivo encontrado: {arquivo_valido}")
    try:
        df = pd.read_excel(caminho_arquivo_valido, sheet_name="Horas-Trabalhada", header=linha_header_valida)
        print("✅ Aba 'Horas-Trabalhada' carregada com sucesso.")
        print("🔎 Colunas encontradas:", df.columns.tolist())

        if "Semana" not in df.columns or "Horas" not in df.columns:
            print(f"⚠️ Colunas obrigatórias não encontradas em '{nome_aplicacao}'.")
            nao_atualizadas.append({"Aplicação": nome_aplicacao, "Motivo": "Colunas ausentes"})
            continue

        # Normalização da coluna "Semana"
        df["Semana"] = df["Semana"].astype(str).apply(normalizar_semana)
        semanas_validas = df["Semana"].dropna().unique()
        print("📌 Semanas encontradas (normalizadas):", list(semanas_validas))

        df_filtrado = df[df["Semana"] == semana_passada_formatada].copy()
        if df_filtrado.empty:
            print("⚠️ Nenhum registro encontrado com a semana esperada.")

            # Capturar a última semana disponível
            semanas_validas = sorted(df["Semana"].dropna().unique())
            ultima_semana = semanas_validas[-1] if semanas_validas else "Nenhuma semana registrada"

            nao_atualizadas.append({
                "Aplicação": nome_aplicacao,
                "Motivo": "Semana não encontrada",
                "Última semana encontrada": ultima_semana
            })
            continue


        print(f"🔎 Registros encontrados para a semana passada: {len(df_filtrado)}")

        df_filtrado.loc[:, "Horas"] = df_filtrado["Horas"].apply(converter_para_timedelta)
        horas_total = df_filtrado["Horas"].sum()
        print("⏱️ Total de horas somadas:", formatar_horas(horas_total))

        relatorio.append({
            "Aplicação": nome_aplicacao,
            "Semana": semana_passada_formatada,
            "Horas na semana": formatar_horas(horas_total)
        })

        print("✅ Aplicação adicionada ao relatório.\n")

    except Exception as e:
        print(f"❌ Erro na aplicação '{nome_aplicacao}': {e}")
        nao_atualizadas.append({"Aplicação": nome_aplicacao, "Motivo": str(e)})


# Gerar nome do arquivo com base na semana passada
nome_semana_formatado = semana_passada_formatada.replace("/", "-").replace(" ", "")
nome_arquivo = f"Relatorio_Horas_{nome_semana_formatado}.xlsx"
CAMINHO_RELATORIO = os.path.join(CAMINHO_BASE, nome_arquivo)

# Salvar relatório principal
df_resultado = pd.DataFrame(relatorio)
df_resultado.to_excel(CAMINHO_RELATORIO, index=False)

# Autoajuste e centralização
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

print(f"\n📄 Relatório final salvo como: {nome_arquivo}")

# Salvar planilha de aplicações não atualizadas
if nao_atualizadas:
    caminho_nao_atualizadas = os.path.join(CAMINHO_BASE, "Aplicacoes_Nao_Atualizadas.xlsx")
    df_nao = pd.DataFrame(nao_atualizadas)
    df_nao.to_excel(caminho_nao_atualizadas, index=False)
    print("⚠️ Arquivo gerado com aplicações não atualizadas: Aplicacoes_Nao_Atualizadas.xlsx")
