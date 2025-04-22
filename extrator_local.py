from relatorio_horas import RelatorioHoras
import os

def main():
    print("🚀 Iniciando execução do extrator_local.py")
    pasta_atual = os.path.dirname(__file__)
    caminho_projeto = os.path.abspath(os.path.join(pasta_atual, ".."))
    print(f"🗂️ Caminho raiz do projeto: {caminho_projeto}")

    relatorio = RelatorioHoras(caminho_projeto)
    relatorio.processar()
    relatorio.salvar_relatorios()
    print("✅ Finalizou execução de extrator_local.py\n")

# Apenas se rodar diretamente (não pelo main.exe)
if __name__ == "__main__":
    main()
