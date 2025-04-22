import os
import pandas as pd
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Alignment

from leitor_planilha import LeitorPlanilha
from utilitarios import normalizar_semana, converter_para_timedelta, formatar_horas

class RelatorioHoras:
    def __init__(self, caminho_base):
        self.caminho_base = caminho_base

        # Caminho correto: mesma pasta onde está o .bat, pasta "Aplicações"
        self.caminho_aplicacoes = os.path.join(self.caminho_base, "Aplicações")

        print(f"📁 Caminho base (projeto raiz): {self.caminho_base}")
        print(f"📁 Caminho das aplicações: {self.caminho_aplicacoes}")

        self.relatorio = []
        self.nao_atualizadas = []
        self.data_referencia = self._definir_semana_passada()


    def _definir_semana_passada(self):
        hoje = datetime.today()
        segunda = hoje - timedelta(days=hoje.weekday() + 7)
        sexta = segunda + timedelta(days=4)
        semana_formatada = f"{segunda.strftime('%d/%m')} a {sexta.strftime('%d/%m')}"
        print(f"\n📅 Semana passada esperada (normalizada): '{semana_formatada}'\n")
        return semana_formatada

    def processar(self):
        print("🔍 Iniciando processamento das aplicações...\n")

        for nome_aplicacao in os.listdir(self.caminho_aplicacoes):
            pasta = os.path.join(self.caminho_aplicacoes, nome_aplicacao)
            if not os.path.isdir(pasta):
                continue

            arquivos = [f for f in os.listdir(pasta) if f.endswith(".xlsx")]
            if not arquivos:
                continue

            print(f"\n📂 Processando aplicação: {nome_aplicacao}")
            print(f"  ➤ Verificando arquivos em: {pasta}")
            self._processar_aplicacao(nome_aplicacao, pasta, arquivos)

        print(f"\n📋 Total de aplicações com dados de horas: {len(self.relatorio)}")
        print(f"📋 Total de aplicações não atualizadas: {len(self.nao_atualizadas)}")

    def _processar_aplicacao(self, nome_aplicacao, pasta, arquivos):
        for arquivo in arquivos:
            caminho_arquivo = os.path.join(pasta, arquivo)
            try:
                linha_header = LeitorPlanilha.encontrar_linha_cabecalho(caminho_arquivo)
                if linha_header is None:
                    print(f"  ⚠️ Cabeçalho não encontrado em '{arquivo}'")
                    continue

                df = LeitorPlanilha.carregar_planilha(caminho_arquivo, linha_header)
                if "Semana" not in df.columns or "Horas" not in df.columns:
                    print(f"  ⚠️ Colunas obrigatórias ausentes no arquivo '{arquivo}'")
                    continue

                print(f"  ✅ Arquivo válido: {arquivo} - Linhas lidas: {len(df)}")

                df["Semana"] = df["Semana"].astype(str).apply(normalizar_semana)
                df_filtrado = df[df["Semana"] == self.data_referencia].copy()

                if df_filtrado.empty:
                    print(f"  ⚠️ Nenhum dado encontrado para a semana '{self.data_referencia}' nesse arquivo.")
                    continue

                df_filtrado["Horas"] = df_filtrado["Horas"].apply(converter_para_timedelta)
                horas_total = df_filtrado["Horas"].sum()

                self.relatorio.append({
                    "Aplicação": nome_aplicacao,
                    "Semana": self.data_referencia,
                    "Horas na semana": formatar_horas(horas_total)
                })

                print(f"  ✅ {nome_aplicacao} - Horas: {formatar_horas(horas_total)}")
                return  # Sai do loop após encontrar o primeiro arquivo válido

            except Exception as e:
                print(f"  ❌ Erro ao processar '{arquivo}': {e}")
                continue

        # Nenhum dado válido encontrado
        ultima_semana = "Nenhuma semana registrada"
        try:
            df_todas = LeitorPlanilha.carregar_planilha(caminho_arquivo, linha_header)
            if "Semana" in df_todas.columns:
                df_todas["Semana"] = df_todas["Semana"].astype(str).apply(normalizar_semana)
                semanas_validas = sorted(df_todas["Semana"].dropna().unique())
                if semanas_validas:
                    ultima_semana = semanas_validas[-1]
        except Exception as e:
            print(f"  ⚠️ Falha ao tentar identificar a última semana de '{nome_aplicacao}': {e}")

        self.nao_atualizadas.append({
            "Aplicação": nome_aplicacao,
            "Motivo": "Semana não encontrada",
            "Última semana encontrada": ultima_semana
        })

    # def salvar_relatorios(self):
    #     nome_semana = self.data_referencia.replace("/", "-").replace(" ", "")
    #     caminho_saida = os.path.join(self.caminho_base, f"Relatorio_Horas_{nome_semana}.xlsx")

    #     df = pd.DataFrame(self.relatorio)
    #     df.to_excel(caminho_saida, index=False)
    #     self._ajustar_excel(caminho_saida)
    #     print(f"\n💾 Relatório principal salvo em: {caminho_saida}")

    #     if self.nao_atualizadas:
    #         caminho_nao = os.path.join(self.caminho_base, "Aplicacoes_Nao_Atualizadas.xlsx")
    #         pd.DataFrame(self.nao_atualizadas).to_excel(caminho_nao, index=False)
    #         print(f"⚠️ Relatório de aplicações não atualizadas salvo em: {caminho_nao}")

    def salvar_relatorios(self):
        nome_semana = self.data_referencia.replace("/", "-").replace(" ", "")
        caminho_saida = os.path.join(self.caminho_base, f"Relatorio_Horas_{nome_semana}.xlsx")

        df = pd.DataFrame(self.relatorio)
        df.to_excel(caminho_saida, index=False)
        self._ajustar_excel(caminho_saida)
        print(f"\n💾 Relatório principal salvo em: {caminho_saida}")

        if self.nao_atualizadas:
            # Gerar apenas com as duas colunas: Aplicação e Última semana encontrada
            dados_nao_atualizadas = [
                {
                    "Aplicação": item["Aplicação"],
                    "Última semana encontrada": item.get("Última semana encontrada", "Nenhuma semana registrada")
                }
                for item in self.nao_atualizadas
            ]

            caminho_nao = os.path.join(self.caminho_base, "Aplicacoes_Nao_Atualizadas.xlsx")
            pd.DataFrame(dados_nao_atualizadas).to_excel(caminho_nao, index=False)
            print(f"⚠️ Relatório de aplicações não atualizadas salvo em: {caminho_nao}")


    def _ajustar_excel(self, caminho_excel):
        wb = load_workbook(caminho_excel)
        ws = wb.active
        for coluna in ws.columns:
            max_largura = max((len(str(cell.value)) for cell in coluna if cell.value), default=0)
            letra = coluna[0].column_letter
            for celula in coluna:
                celula.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[letra].width = max_largura + 2
        wb.save(caminho_excel)
