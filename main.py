import os
import sys

def main():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(__file__)

    script_dir = os.path.join(base_dir, "extrator-files")
    print(f"📁 Caminho base detectado: {base_dir}")
    print(f"📁 Esperado 'extrator-files': {script_dir}")

    if not os.path.exists(script_dir):
        print("❌ A pasta 'extrator-files' não foi encontrada.")
        input("Pressione qualquer tecla para sair...")
        return

    sys.path.insert(0, script_dir)  # 💡 Adiciona o caminho ao sys.path

    # Só importa agora que o path está configurado
    try:
        from extrator_local import main as executar_extrator
    except Exception as import_error:
        print(f"❌ Erro ao importar o extrator: {import_error}")
        input("Pressione qualquer tecla para sair...")
        return

    os.chdir(script_dir)
    print("=============================================")
    print(f"Diretório atual: {os.getcwd()}")
    print("Executando extracao de horas por aplicacao...")
    print("=============================================")

    try:
        executar_extrator()
        print("✅ Execução concluída com sucesso.")
    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")

    input("Pressione qualquer tecla para continuar...")

if __name__ == "__main__":
    main()
