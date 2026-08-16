import sys
import threading
import time
import webbrowser
import socket
import traceback
from pathlib import Path

from streamlit.web import cli as stcli


def caminho_recurso(nome):
    """
    Localiza arquivos tanto durante o desenvolvimento
    quanto dentro do aplicativo gerado pelo PyInstaller.
    """

    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / nome

    return Path(__file__).parent / nome


def caminho_log():
    """
    Cria o log na pasta pessoal do usuário.
    """

    return Path.home() / "SistemaCobranca.log"


def porta_esta_aberta(
    host="127.0.0.1",
    porta=8501
):
    """
    Verifica se o servidor Streamlit já está respondendo.
    """

    try:
        with socket.create_connection(
            (host, porta),
            timeout=1
        ):
            return True

    except OSError:
        return False


def abrir_navegador_quando_pronto():
    """
    Aguarda o Streamlit realmente iniciar antes
    de abrir o navegador.
    """

    for _ in range(60):

        if porta_esta_aberta():

            webbrowser.open(
                "http://localhost:8501"
            )

            return

        time.sleep(0.5)


def main():

    arquivo_log = caminho_log()

    try:

        app = caminho_recurso(
            "app.py"
        )

        # Thread responsável apenas por esperar
        # o servidor ficar disponível e então
        # abrir o navegador.
        thread_navegador = threading.Thread(
            target=abrir_navegador_quando_pronto,
            daemon=True
        )

        thread_navegador.start()

        # Simula os argumentos que normalmente seriam:
        #
        # streamlit run app.py
        #
        sys.argv = [
            "streamlit",
            "run",
            str(app),
            "--global.developmentMode=false",
            "--server.headless=true",
            "--server.port=8501",
            "--browser.gatherUsageStats=false",
        ]

        # Inicia Streamlit dentro do próprio processo.
        stcli.main()

    except Exception:

        with open(
            arquivo_log,
            "w",
            encoding="utf-8"
        ) as log:

            log.write(
                traceback.format_exc()
            )

        raise


if __name__ == "__main__":
    main()
