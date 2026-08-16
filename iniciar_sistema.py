import sys
import subprocess
import time
import webbrowser
from pathlib import Path


def caminho_recurso(nome):
    """
    Localiza arquivos tanto durante o desenvolvimento
    quanto dentro do aplicativo gerado pelo PyInstaller.
    """

    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / nome

    return Path(__file__).parent / nome


def main():

    app = caminho_recurso("app.py")

    comando = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(app),
        "--server.headless=true",
        "--server.port=8501",
        "--browser.gatherUsageStats=false",
    ]

    processo = subprocess.Popen(comando)

    # Aguarda o Streamlit iniciar
    time.sleep(3)

    # Abre a interface no navegador
    webbrowser.open("http://localhost:8501")

    try:
        processo.wait()

    except KeyboardInterrupt:
        processo.terminate()


if __name__ == "__main__":
    main()
