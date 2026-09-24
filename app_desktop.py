# =============================================================================
# SISTEMA DE COBRANÇA - MÓDULO 1
# Interface Desktop
#
# Esta interface utiliza o parser estável:
# parser_pdf_base_mestre_v1_3_estavel.py
#
# Não altera nenhuma regra do parser.
# =============================================================================

from pathlib import Path
import shutil
import sys
import tkinter as tk
from tkinter import (
    filedialog,
    messagebox,
    ttk
)

import pandas as pd

from banco_dados import (
    caminho_banco, caminho_banco_local, importar_telefones, inicializar_banco,
    migrar_banco_local_para_portatil, modo_portatil_ativo, sincronizar_devedores,
)

from openpyxl import load_workbook
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment,
    Border,
    Side
)
from openpyxl.utils import get_column_letter

from parser_pdf_base_mestre_v1_3_estavel import (
    processar_pdf,
    gerar_resumo,
    validar_estrutura
)
from gerador_comunicados import (
    caminho_modelo_padrao,
    gerar_pacote_comunicados_zip,
)
from modulo_comunicacoes import (
    NOME_MODELO,
    caminho_modelo_importacao,
    ler_planilha_contatos,
)
from modulo_consultas import JanelaConsultas
from cargas_dados import carregar_base_atual, listar_condominios_atuais, registrar_carga


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def resultado_esta_ok(resultado):
    """
    Retorna True somente se todas as validações
    financeiras do PDF estiverem OK.
    """

    validacao = resultado.get(
        "validacao",
        {}
    )

    if not validacao:
        return False

    return all(
        item.get("ok", False)
        for item in validacao.values()
    )


def formatar_excel(arquivo_excel):
    """
    Aplica a formatação padrão ao Excel gerado.
    """

    wb = load_workbook(
        arquivo_excel
    )

    cor_cabecalho = "4F81BD"

    borda_fina = Border(
        left=Side(
            style="thin",
            color="B7B7B7"
        ),
        right=Side(
            style="thin",
            color="B7B7B7"
        ),
        top=Side(
            style="thin",
            color="B7B7B7"
        ),
        bottom=Side(
            style="thin",
            color="B7B7B7"
        )
    )

    for nome_aba in [
        "Dados",
        "Resumo",
        "Validação"
    ]:

        ws = wb[nome_aba]

        # -----------------------------------------------------
        # Bordas
        # -----------------------------------------------------

        for linha in ws.iter_rows():
            for cell in linha:
                cell.border = borda_fina

        # -----------------------------------------------------
        # Cabeçalho
        # -----------------------------------------------------

        for cell in ws[1]:

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor=cor_cabecalho
            )

            cell.font = Font(
                color="FFFFFF",
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # -----------------------------------------------------
        # Congelar cabeçalho
        # -----------------------------------------------------

        ws.freeze_panes = "A2"

        # -----------------------------------------------------
        # Filtro
        # -----------------------------------------------------

        ws.auto_filter.ref = ws.dimensions

        # -----------------------------------------------------
        # Ajustar largura
        # -----------------------------------------------------

        for coluna in ws.columns:

            letra = get_column_letter(
                coluna[0].column
            )

            maior = 0

            for cell in coluna:

                valor = (
                    ""
                    if cell.value is None
                    else str(cell.value)
                )

                maior = max(
                    maior,
                    len(valor)
                )

            ws.column_dimensions[
                letra
            ].width = min(
                maior + 2,
                45
            )

    # =========================================================================
    # Formatação monetária - aba Dados
    # =========================================================================

    ws_dados = wb["Dados"]

    cabecalhos_dados = {
        cell.value: cell.column
        for cell in ws_dados[1]
    }

    for nome_coluna in [
        "vlr_base",
        "multa",
        "juros",
        "correcao",
        "vlr_corrigido"
    ]:

        if nome_coluna in cabecalhos_dados:

            coluna = cabecalhos_dados[
                nome_coluna
            ]

            for linha in range(
                2,
                ws_dados.max_row + 1
            ):

                ws_dados.cell(
                    row=linha,
                    column=coluna
                ).number_format = '#,##0.00'

    # =========================================================================
    # Formatação monetária - aba Resumo
    # =========================================================================

    ws_resumo = wb["Resumo"]

    cabecalhos_resumo = {
        cell.value: cell.column
        for cell in ws_resumo[1]
    }

    for nome_coluna in [
        "valor_total_devido_base",
        "valor_total_corrigido"
    ]:

        if nome_coluna in cabecalhos_resumo:

            coluna = cabecalhos_resumo[
                nome_coluna
            ]

            for linha in range(
                2,
                ws_resumo.max_row + 1
            ):

                ws_resumo.cell(
                    row=linha,
                    column=coluna
                ).number_format = '#,##0.00'

    wb.save(
        arquivo_excel
    )


# =============================================================================
# APLICAÇÃO
# =============================================================================

class SistemaCobrancaApp:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Sistema de Cobrança"
        )

        # A janela se adapta à área útil do monitor. No Windows, a escala
        # de exibição (125% ou 150%) reduz essa área mesmo em tela cheia.
        largura_tela = self.root.winfo_screenwidth()
        altura_tela = self.root.winfo_screenheight()

        largura_inicial = min(1250, max(900, largura_tela - 60))
        altura_inicial = min(850, max(620, altura_tela - 100))

        self.root.geometry(
            f"{largura_inicial}x{altura_inicial}"
        )

        self.root.minsize(
            min(950, largura_tela - 60),
            min(640, altura_tela - 100)
        )
        # No Mac, abre em tela cheia; no Windows, maximiza mantendo a barra
        # do sistema. Escape sai da tela cheia do Mac.
        self.root.bind("<Escape>", self.sair_tela_cheia)
        self.root.after_idle(self.abrir_em_tela_cheia)

        # Somente a linha da tabela cresce ou diminui. As linhas dos botões
        # de salvar e do rodapé ficam sempre reservadas e visíveis.
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        # =====================================================
        # ESTILO
        # =====================================================

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Titulo.TLabel",
            font=("Arial", 22, "bold")
        )

        style.configure(
            "Subtitulo.TLabel",
            font=("Arial", 11)
        )

        style.configure(
            "Indicador.TLabel",
            font=("Arial", 11, "bold")
        )

        style.configure(
            "Status.TLabel",
            font=("Arial", 10, "bold")
        )

        style.configure(
            "Acao.TButton",
            font=("Arial", 11, "bold"),
            padding=10
        )

        style.configure(
            "Secundario.TButton",
            font=("Arial", 10),
            padding=8
        )

        # =====================================================
        # DADOS DA EXECUÇÃO
        # =====================================================

        self.arquivos_pdf = []

        self.base_mestre = None

        self.resumo = None

        self.validacao_estrutura = None

        self.relatorio = None

        self.devedores_com_contatos = None

        self.resumo_contatos = None

        # A primeira aba reúne importação e saídas; a consulta permanece
        # acessível sem abrir janelas intermediárias.
        barra = ttk.Frame(root, padding=(16, 8))
        barra.grid(row=0, column=0, sticky="ew")
        linha_titulo = ttk.Frame(barra)
        linha_titulo.pack(fill="x")
        ttk.Label(
            linha_titulo, text="Sistema de Cobrança — versão de teste",
            font=("Arial", 16, "bold"),
        ).pack(side="left")
        self.botao_encerrar = ttk.Button(
            linha_titulo, text="Sair", command=self.encerrar_sistema
        )
        self.botao_encerrar.pack(side="right")

        self.abas_principais = ttk.Notebook(root)
        self.abas_principais.grid(row=1, column=0, sticky="nsew")
        self.tela_consulta = ttk.Frame(self.abas_principais)
        self.tela_importacao = ttk.Frame(self.abas_principais)
        self.abas_principais.add(self.tela_importacao, text="Importar / salvar")
        self.abas_principais.add(self.tela_consulta, text="Condomínios / devedores")
        self.tela_importacao.columnconfigure(0, weight=1)
        self.tela_importacao.rowconfigure(3, weight=1, minsize=90)

        # =====================================================
        # CABEÇALHO
        # =====================================================

        frame_cabecalho = ttk.Frame(
            self.tela_importacao,
            padding=(20, 18, 20, 8)
        )

        frame_cabecalho.grid(
            row=0,
            column=0,
            sticky="ew"
        )

        self.label_carga_atual = ttk.Label(
            frame_cabecalho,
            text="Sem carga salva",
        )
        self.label_carga_atual.pack(anchor="e")

        titulo = ttk.Label(
            frame_cabecalho, text="Importação de PDFs", style="Titulo.TLabel"
        )
        titulo.pack()

        subtitulo = ttk.Label(
            frame_cabecalho,
            text=(
                "Módulo 1 — Conversão de relatórios "
                "de inadimplência em Base Mestre"
            ),
            style="Subtitulo.TLabel"
        )

        subtitulo.pack(
            pady=(4, 0)
        )

        # =====================================================
        # SELEÇÃO DE ARQUIVOS
        # =====================================================

        frame_selecao = ttk.LabelFrame(
            self.tela_importacao,
            text="1. Seleção dos relatórios",
            padding=15
        )

        frame_selecao.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=20,
            pady=(5, 10)
        )

        frame_botoes = ttk.Frame(
            frame_selecao
        )

        frame_botoes.pack()

        self.botao_selecionar = ttk.Button(
            frame_botoes,
            text="Selecionar PDFs",
            command=self.selecionar_pdfs,
            style="Acao.TButton"
        )

        self.botao_selecionar.grid(
            row=0,
            column=0,
            padx=6
        )

        self.botao_processar = ttk.Button(
            frame_botoes,
            text="Processar arquivos",
            command=self.processar_arquivos,
            state="disabled",
            style="Acao.TButton"
        )

        self.botao_processar.grid(
            row=0,
            column=1,
            padx=6
        )

        self.label_selecionados = ttk.Label(
            frame_selecao,
            text="Nenhum PDF selecionado."
        )

        self.label_selecionados.pack(
            pady=(10, 4)
        )

        # =====================================================
        # LISTA DOS ARQUIVOS SELECIONADOS
        # =====================================================

        frame_lista = ttk.Frame(
            frame_selecao
        )

        frame_lista.pack(
            fill="both",
            expand=True,
            padx=5,
            pady=(5, 0)
        )

        self.lista_arquivos = ttk.Treeview(
            frame_lista,
            columns=("arquivo",),
            show="headings",
            height=4,
            selectmode="browse"
        )

        self.lista_arquivos.heading(
            "arquivo",
            text="Arquivos selecionados"
        )

        self.lista_arquivos.column(
            "arquivo",
            anchor="w",
            width=700
        )

        scroll_lista = ttk.Scrollbar(
            frame_lista,
            orient="vertical",
            command=self.lista_arquivos.yview
        )

        self.lista_arquivos.configure(
            yscrollcommand=scroll_lista.set
        )

        self.lista_arquivos.pack(
            side="left",
            fill="both",
            expand=True
        )

        scroll_lista.pack(
            side="right",
            fill="y"
        )

        # =====================================================
        # INDICADORES
        # =====================================================

        frame_indicadores = ttk.LabelFrame(
            self.tela_importacao,
            text="2. Resultado do processamento",
            padding=12
        )

        frame_indicadores.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=20,
            pady=(0, 10)
        )

        indicadores = ttk.Frame(
            frame_indicadores
        )

        indicadores.pack()

        self.label_processados = ttk.Label(
            indicadores,
            text="PDFs processados: 0",
            style="Indicador.TLabel"
        )

        self.label_processados.grid(
            row=0,
            column=0,
            padx=20
        )

        self.label_validos = ttk.Label(
            indicadores,
            text="PDFs validados: 0",
            style="Indicador.TLabel"
        )

        self.label_validos.grid(
            row=0,
            column=1,
            padx=20
        )

        self.label_divergentes = ttk.Label(
            indicadores,
            text="Arquivos com problema: 0",
            style="Indicador.TLabel"
        )

        self.label_divergentes.grid(
            row=0,
            column=2,
            padx=20
        )

        self.label_registros = ttk.Label(
            indicadores,
            text="Registros consolidados: 0",
            style="Indicador.TLabel"
        )

        self.label_registros.grid(
            row=0,
            column=3,
            padx=20
        )

        # =====================================================
        # TABELA DE RESULTADOS
        # =====================================================

        frame_tabela = ttk.Frame(
            self.tela_importacao
        )

        frame_tabela.grid(
            row=3,
            column=0,
            sticky="nsew",
            padx=20,
            pady=(0, 10)
        )

        colunas = (
            "arquivo",
            "codigo",
            "condominio",
            "registros",
            "status"
        )

        self.tabela = ttk.Treeview(
            frame_tabela,
            columns=colunas,
            show="headings",
            height=10
        )

        self.tabela.heading(
            "arquivo",
            text="Arquivo"
        )

        self.tabela.heading(
            "codigo",
            text="Código"
        )

        self.tabela.heading(
            "condominio",
            text="Condomínio"
        )

        self.tabela.heading(
            "registros",
            text="Registros"
        )

        self.tabela.heading(
            "status",
            text="Situação"
        )

        self.tabela.column(
            "arquivo",
            width=160
        )

        self.tabela.column(
            "codigo",
            width=90,
            anchor="center"
        )

        self.tabela.column(
            "condominio",
            width=380
        )

        self.tabela.column(
            "registros",
            width=100,
            anchor="center"
        )

        self.tabela.column(
            "status",
            width=160,
            anchor="center"
        )

        scrollbar = ttk.Scrollbar(
            frame_tabela,
            orient="vertical",
            command=self.tabela.yview
        )

        self.tabela.configure(
            yscrollcommand=scrollbar.set
        )

        self.tabela.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        # =====================================================
        # STATUS
        # =====================================================

        self.label_status = ttk.Label(
            self.tela_importacao,
            text="Pronto.",
            anchor="center",
            style="Status.TLabel"
        )

        self.label_status.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(0, 8)
        )

        # =====================================================
        # SAÍDAS
        # =====================================================

        frame_saida = ttk.Frame(self.tela_importacao, padding=(0, 4))

        frame_saida.grid(
            row=5,
            column=0,
            sticky="ew",
            padx=20,
            pady=(0, 10)
        )

        for coluna in range(3):
            frame_saida.columnconfigure(coluna, weight=1, uniform="saidas")

        area_mestre = ttk.LabelFrame(
            frame_saida, text="Exportar Base Mestre", padding=10
        )
        area_comunicados = ttk.LabelFrame(
            frame_saida, text="Exportar comunicados", padding=10
        )
        area_telefones = ttk.LabelFrame(
            frame_saida, text="Importar telefones", padding=10
        )
        for coluna, area in enumerate(
            (area_mestre, area_comunicados, area_telefones)
        ):
            area.grid(row=0, column=coluna, padx=5, sticky="nsew")

        self.botao_excel = ttk.Button(
            area_mestre,
            text="Salvar Base Mestre Excel",
            command=self.salvar_excel,
            state="disabled",
            style="Secundario.TButton",
        )
        self.botao_excel.pack(fill="x", pady=(0, 6))

        self.botao_comunicados = ttk.Button(
            area_comunicados,
            text="Salvar Comunicados por Condomínio",
            command=self.salvar_comunicados,
            state="disabled",
            style="Secundario.TButton",
        )
        self.botao_comunicados.pack(fill="x")

        self.botao_csv = ttk.Button(
            area_mestre,
            text="Salvar Base Mestre CSV",
            command=self.salvar_csv,
            state="disabled",
            style="Secundario.TButton"
        )

        self.botao_csv.pack(fill="x")

        self.botao_importar_telefones = ttk.Button(
            area_telefones,
            text="Selecionar planilha",
            command=self.importar_planilha_telefones,
            style="Secundario.TButton"
        )
        self.botao_importar_telefones.pack(fill="x", pady=(0, 6))
        ttk.Button(
            area_telefones, text="Salvar modelo de telefones",
            command=self.salvar_modelo_telefones,
            style="Secundario.TButton",
        ).pack(fill="x")
        self.consulta = JanelaConsultas(self.tela_consulta, incorporada=True)
        self.consulta.janela.pack(fill="both", expand=True)

        # Fecha pela bolinha vermelha no macOS
        # ou pelo X no Windows usando a mesma confirmação.
        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.encerrar_sistema
        )

        self.carregar_base_persistida(mostrar_condominios=True)


    def abrir_em_tela_cheia(self) -> None:
        try:
            if sys.platform == "darwin":
                self.root.attributes("-fullscreen", True)
            elif sys.platform == "win32":
                self.root.state("zoomed")
            else:
                self.root.attributes("-zoomed", True)
        except tk.TclError:
            # Alguns Tk não oferecem a opção de maximização. Nesse caso,
            # usamos a área da tela como aproximação.
            self.root.geometry(
                f"{self.root.winfo_screenwidth()}x"
                f"{self.root.winfo_screenheight()}+0+0"
            )

    def sair_tela_cheia(self, _evento=None) -> None:
        if sys.platform == "darwin":
            try:
                self.root.attributes("-fullscreen", False)
            except tk.TclError:
                pass


    def carregar_base_persistida(self, *, mostrar_condominios=False):
        """Restaura a visao atual sem exigir nova importacao dos PDFs."""

        base = carregar_base_atual()
        self.base_mestre = base if not base.empty else None

        if self.base_mestre is None:
            self.resumo = None
            self.validacao_estrutura = None
            self.devedores_com_contatos = None
            self.resumo_contatos = None
        else:
            self.resumo = gerar_resumo(self.base_mestre)
            self.validacao_estrutura = validar_estrutura(self.base_mestre)
            contatos = sincronizar_devedores(self.resumo.to_dict(orient="records"))
            self.devedores_com_contatos = pd.DataFrame(contatos)
            com_telefone = sum(bool(contato["telefone"]) for contato in contatos)
            self.resumo_contatos = {
                "total": len(contatos),
                "com_telefone": com_telefone,
                "sem_telefone": len(contatos) - com_telefone,
            }

        estado = "normal" if self.base_mestre is not None else "disabled"
        for botao in (self.botao_excel, self.botao_csv, self.botao_comunicados):
            botao.config(state=estado)

        self.label_registros.config(
            text=f"Registros consolidados: {len(base)}"
        )

        if mostrar_condominios:
            for item in self.tabela.get_children():
                self.tabela.delete(item)
            condominios = listar_condominios_atuais()
            for condominio in condominios:
                arquivos = ", ".join(
                    item["nome"] for item in condominio["arquivos"]
                )
                self.tabela.insert(
                    "",
                    "end",
                    values=(
                        arquivos,
                        condominio["codigo_condominio"],
                        condominio["condominio"],
                        condominio["quantidade_lancamentos"],
                        condominio["criada_em"][:16].replace("T", " "),
                    ),
                )
            if condominios:
                datas = [item["criada_em"] for item in condominios]
                self.label_status.config(
                    text=(
                        "Dados recuperados do banco. "
                        f"Última atualização: {max(datas)[:16].replace('T', ' ')}."
                    )
                )
        condominios_atuais = listar_condominios_atuais()
        if condominios_atuais:
            data_mais_recente = max(item["criada_em"] for item in condominios_atuais)
            self.label_carga_atual.config(
                text=f"Última carga: {data_mais_recente[:16].replace('T', ' ')}"
            )
        else:
            self.label_carga_atual.config(text="Sem carga salva")
        self.consulta.atualizar_condominios()

    def importar_planilha_telefones(self) -> None:
        arquivo = filedialog.askopenfilename(
            title="Selecionar planilha de telefones",
            filetypes=[
                ("Planilhas aceitas", "*.xlsx *.csv"),
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
            ],
            parent=self.root,
        )
        if not arquivo:
            return
        try:
            leitura = ler_planilha_contatos(arquivo)
            resultado = importar_telefones(leitura["registros"])
            erros = leitura["erros"] + resultado["erros"]
        except Exception as erro:
            messagebox.showerror(
                "Erro na importação", f"Não foi possível importar a planilha.\n\n{erro}",
                parent=self.root,
            )
            return
        self.consulta.atualizar_condominios()
        mensagem = f"Telefones importados: {resultado['atualizados']}."
        if resultado["ignorados"]:
            mensagem += f" Repetidos ignorados: {resultado['ignorados']}."
        if erros:
            detalhes = "\n".join(
                f"Linha {item['linha']}: {item['erro']}" for item in erros[:10]
            )
            if len(erros) > 10:
                detalhes += f"\n... e mais {len(erros) - 10} erro(s)."
            mensagem += f"\n\nLinhas não importadas: {len(erros)}.\n{detalhes}"
        messagebox.showinfo("Importação concluída", mensagem, parent=self.root)

    def salvar_modelo_telefones(self) -> None:
        destino = filedialog.asksaveasfilename(
            title="Salvar modelo de importação",
            defaultextension=".xlsx", initialfile=NOME_MODELO,
            filetypes=[("Arquivo Excel", "*.xlsx")], parent=self.root,
        )
        if not destino:
            return
        try:
            shutil.copyfile(caminho_modelo_importacao(), destino)
        except Exception as erro:
            messagebox.showerror(
                "Erro", f"Não foi possível salvar o modelo.\n\n{erro}",
                parent=self.root,
            )
            return
        messagebox.showinfo("Modelo salvo", "O modelo foi salvo.", parent=self.root)


    # =========================================================================
    # SELEÇÃO DOS PDFs
    # =========================================================================

    def selecionar_pdfs(self):

        arquivos = filedialog.askopenfilenames(
            title="Selecione os relatórios PDF",
            filetypes=[
                (
                    "Arquivos PDF",
                    "*.pdf"
                )
            ]
        )

        if not arquivos:
            return

        # A seleção não altera os dados atuais nem desabilita as saídas.

        self.arquivos_pdf = [
            Path(arquivo)
            for arquivo in arquivos
        ]

        quantidade = len(
            self.arquivos_pdf
        )

        self.label_selecionados.config(
            text=(
                f"{quantidade} "
                f"PDF(s) selecionado(s)."
            )
        )

        for item in self.lista_arquivos.get_children():
            self.lista_arquivos.delete(item)

        for arquivo in self.arquivos_pdf:

            self.lista_arquivos.insert(
                "",
                "end",
                values=(
                    arquivo.name,
                )
            )
        self.botao_processar.config(
            state="normal"
        )

        self.label_status.config(
            text=(
                "Nova seleção carregada. "
                "Clique em Processar arquivos."
            )
        )


    # =========================================================================
    # PROCESSAMENTO
    # =========================================================================

    def processar_arquivos(self):

        if not self.arquivos_pdf:
            return

        # -----------------------------------------------------
        # Limpar execução anterior
        # -----------------------------------------------------

        for item in self.tabela.get_children():

            self.tabela.delete(
                item
            )

        self.botao_processar.config(
            state="disabled"
        )

        self.botao_selecionar.config(
            state="disabled"
        )

        self.botao_excel.config(
            state="disabled"
        )

        self.botao_csv.config(
            state="disabled"
        )

        self.botao_comunicados.config(
            state="disabled"
        )

        self.label_status.config(
            text="Processando..."
        )

        self.root.update()

        bases_validas = []

        fontes_validas = []

        registros_relatorio = []

        # -----------------------------------------------------
        # Processar PDF por PDF
        # -----------------------------------------------------

        for caminho_pdf in self.arquivos_pdf:

            self.label_status.config(
                text=(
                    f"Processando "
                    f"{caminho_pdf.name}..."
                )
            )

            self.root.update()

            try:

                resultado = processar_pdf(
                    caminho_pdf
                )

                valido = resultado_esta_ok(
                    resultado
                )

                dados = resultado[
                    "dados"
                ]

                cabecalho = resultado[
                    "cabecalho"
                ]

                status = (
                    "OK"
                    if valido
                    else "DIVERGENTE"
                )

                quantidade = len(
                    dados
                )

                registro_relatorio = {
                    "arquivo":
                        caminho_pdf.name,

                    "codigo_condominio":
                        cabecalho.get(
                            "codigo_condominio",
                            ""
                        ),

                    "condominio":
                        cabecalho.get(
                            "condominio",
                            ""
                        ),

                    "registros":
                        quantidade,

                    "status":
                        status
                }

                registros_relatorio.append(
                    registro_relatorio
                )

                if valido:

                    bases_validas.append(
                        dados
                    )

                    fontes_validas.append({
                        "arquivo": caminho_pdf,
                        "codigo_condominio": cabecalho.get("codigo_condominio", ""),
                        "condominio": cabecalho.get("condominio", ""),
                        "dados": dados,
                    })

                self.tabela.insert(
                    "",
                    "end",
                    values=(
                        caminho_pdf.name,
                        registro_relatorio[
                            "codigo_condominio"
                        ],
                        registro_relatorio[
                            "condominio"
                        ],
                        quantidade,
                        status
                    )
                )

            except Exception as erro:

                registros_relatorio.append({
                    "arquivo":
                        caminho_pdf.name,

                    "codigo_condominio":
                        "",

                    "condominio":
                        "",

                    "registros":
                        0,

                    "status":
                        "ERRO"
                })

                self.tabela.insert(
                    "",
                    "end",
                    values=(
                        caminho_pdf.name,
                        "",
                        "",
                        0,
                        "ERRO"
                    )
                )

                print(
                    f"Erro em "
                    f"{caminho_pdf.name}: "
                    f"{erro}"
                )

        # -----------------------------------------------------
        # Relatório
        # -----------------------------------------------------

        self.relatorio = pd.DataFrame(
            registros_relatorio
        )

        # -----------------------------------------------------
        # Consolidar PDFs válidos
        # -----------------------------------------------------

        falha_recarga = False

        if bases_validas:
            try:
                resultado_carga = registrar_carga(fontes_validas)
            except Exception as erro:
                resultado_carga = None
                estado = "normal" if self.base_mestre is not None else "disabled"
                for botao in (
                    self.botao_excel, self.botao_csv, self.botao_comunicados
                ):
                    botao.config(state=estado)
                messagebox.showerror(
                    "Banco de dados",
                    (
                        "Os PDFs foram lidos, mas não foi possível salvar "
                        "a nova carga. Os dados anteriores continuam disponíveis.\n\n"
                        f"Detalhes: {erro}"
                    ),
                    parent=self.root,
                )
            else:
                try:
                    self.carregar_base_persistida()
                except Exception as erro:
                    falha_recarga = True
                    estado = "normal" if self.base_mestre is not None else "disabled"
                    for botao in (
                        self.botao_excel, self.botao_csv, self.botao_comunicados
                    ):
                        botao.config(state=estado)
                    messagebox.showerror(
                        "Banco de dados",
                        (
                            "A carga foi salva, mas não foi possível atualizar "
                            "a tela. Feche e abra o aplicativo novamente.\n\n"
                            f"Detalhes: {erro}"
                        ),
                        parent=self.root,
                    )

        else:
            resultado_carga = None
            self.carregar_base_persistida(mostrar_condominios=True)

        # -----------------------------------------------------
        # Indicadores
        # -----------------------------------------------------

        total_processados = len(
            self.relatorio
        )

        total_validos = int(
            (
                self.relatorio[
                    "status"
                ]
                == "OK"
            ).sum()
        )

        total_divergentes = (
            total_processados
            -
            total_validos
        )

        total_registros = (
            len(self.base_mestre)
            if self.base_mestre
            is not None
            else 0
        )

        self.label_processados.config(
            text=(
                f"PDFs processados: "
                f"{total_processados}"
            )
        )

        self.label_validos.config(
            text=(
                f"PDFs validados: "
                f"{total_validos}"
            )
        )

        self.label_divergentes.config(
            text=(
                f"Arquivos com problema: "
                f"{total_divergentes}"
            )
        )

        self.label_registros.config(
            text=(
                f"Registros consolidados: "
                f"{total_registros}"
            )
        )

        # -----------------------------------------------------
        # Reativar controles
        # -----------------------------------------------------

        self.botao_selecionar.config(
            state="normal"
        )

        self.botao_processar.config(
            state="normal"
        )

        # -----------------------------------------------------
        # Status final
        # -----------------------------------------------------

        if falha_recarga:

            self.label_status.config(
                text="Carga salva. Reabra o aplicativo para atualizar a tela."
            )

        elif resultado_carga is None:

            self.label_status.config(
                text=(
                    "Nenhuma carga nova foi salva. "
                    "Os dados anteriores foram mantidos."
                )
            )

        elif resultado_carga["atualizados"] == 0:

            self.label_status.config(
                text=(
                    "PDFs já processados anteriormente. "
                    "Os dados atuais foram mantidos."
                )
            )

        elif total_validos == total_processados:

            self.label_status.config(
                text=(
                    "Processamento concluído. "
                    f"{resultado_carga['atualizados']} condomínio(s) atualizado(s)."
                )
            )

        else:

            self.label_status.config(
                text=(
                    f"{resultado_carga['atualizados']} condomínio(s) atualizado(s). "
                    "Os PDFs com erro ou divergência não alteraram o banco."
                )
            )

        if self.resumo_contatos is not None:

            texto_atual = self.label_status.cget(
                "text"
            )

            self.label_status.config(
                text=(
                    f"{texto_atual} "
                    "Contatos: "
                    f"{self.resumo_contatos['com_telefone']} com telefone; "
                    f"{self.resumo_contatos['sem_telefone']} sem telefone."
                )
            )


    # =========================================================================
    # SALVAR EXCEL
    # =========================================================================

    def salvar_excel(self):

        if self.base_mestre is None:
            return

        arquivo = filedialog.asksaveasfilename(
            title="Salvar Base Mestre",
            defaultextension=".xlsx",
            initialfile="base_mestre.xlsx",
            filetypes=[
                (
                    "Arquivo Excel",
                    "*.xlsx"
                )
            ]
        )

        if not arquivo:
            return

        arquivo = Path(
            arquivo
        )

        try:

            with pd.ExcelWriter(
                arquivo,
                engine="openpyxl"
            ) as writer:

                self.base_mestre.to_excel(
                    writer,
                    sheet_name="Dados",
                    index=False
                )

                self.resumo.to_excel(
                    writer,
                    sheet_name="Resumo",
                    index=False
                )

                self.validacao_estrutura.to_excel(
                    writer,
                    sheet_name="Validação",
                    index=False
                )

            formatar_excel(
                arquivo
            )

            messagebox.showinfo(
                "Sistema de Cobrança",
                "Base Mestre Excel gerada com sucesso."
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro",
                (
                    "Não foi possível salvar "
                    "o arquivo Excel.\n\n"
                    f"{erro}"
                )
            )


    # =========================================================================
    # SALVAR COMUNICADOS DE COBRANÇA
    # =========================================================================

    def salvar_comunicados(self):

        if self.base_mestre is None:
            return

        arquivo = filedialog.asksaveasfilename(
            title="Salvar Comunicados por Condomínio",
            defaultextension=".zip",
            initialfile="comunicados_por_condominio.zip",
            filetypes=[
                (
                    "Arquivo ZIP",
                    "*.zip"
                )
            ]
        )

        if not arquivo:
            return

        try:

            self.label_status.config(
                text="Gerando comunicados de cobrança..."
            )
            self.root.update()

            conteudo = gerar_pacote_comunicados_zip(
                self.base_mestre,
                caminho_modelo_padrao(),
            )

            Path(arquivo).write_bytes(conteudo)

            self.label_status.config(
                text="Comunicados gerados com sucesso."
            )

            messagebox.showinfo(
                "Sistema de Cobrança",
                "Arquivos Excel por condomínio gerados com sucesso."
            )

        except Exception as erro:

            self.label_status.config(
                text="Erro ao gerar os comunicados."
            )

            messagebox.showerror(
                "Erro",
                (
                    "Não foi possível gerar os comunicados.\n\n"
                    f"{erro}"
                )
            )


    # =========================================================================
    # SALVAR CSV
    # =========================================================================

    def salvar_csv(self):

        if self.base_mestre is None:
            return

        arquivo = filedialog.asksaveasfilename(
            title="Salvar Base Mestre CSV",
            defaultextension=".csv",
            initialfile="base_mestre.csv",
            filetypes=[
                (
                    "Arquivo CSV",
                    "*.csv"
                )
            ]
        )

        if not arquivo:
            return

        try:

            self.base_mestre.to_csv(
                arquivo,
                index=False,
                sep=";",
                encoding="utf-8-sig",
                decimal=","
            )

            messagebox.showinfo(
                "Sistema de Cobrança",
                "Base Mestre CSV gerada com sucesso."
            )

        except Exception as erro:

            messagebox.showerror(
                "Erro",
                (
                    "Não foi possível salvar "
                    "o arquivo CSV.\n\n"
                    f"{erro}"
                )
            )


    # =========================================================================
    # ENCERRAR SISTEMA
    # =========================================================================

    def encerrar_sistema(self):

        resposta = messagebox.askyesno(
            "Encerrar sistema",
            (
                "Deseja realmente encerrar "
                "o Sistema de Cobrança?"
            )
        )

        if resposta:
            self.root.destroy()


# =============================================================================
# INICIAR APLICAÇÃO
# =============================================================================

def main():

    root = tk.Tk()

    try:
        if (
            modo_portatil_ativo()
            and not caminho_banco().exists()
            and caminho_banco_local().is_file()
        ):
            resposta = messagebox.askyesnocancel(
                "Banco de dados no pendrive",
                (
                    "Encontrei dados deste sistema salvos neste computador. "
                    "Deseja copiá-los para o pendrive?\n\n"
                    "Sim: copiar os dados existentes para o pendrive.\n"
                    "Não: começar com um banco vazio no pendrive.\n"
                    "Cancelar: fechar o programa sem alterar os dados.\n\n"
                    "O banco original continuará neste computador."
                ),
                parent=root,
            )
            if resposta is None:
                root.destroy()
                return
            if resposta:
                migrar_banco_local_para_portatil()
        inicializar_banco()
    except Exception as erro:
        messagebox.showerror(
            "Sistema de Cobrança",
            (
                "Não foi possível inicializar o banco de dados local.\n\n"
                "O sistema será encerrado para proteger os dados.\n\n"
                f"Detalhes: {erro}"
            ),
            parent=root,
        )
        root.destroy()
        return

    try:
        SistemaCobrancaApp(root)
    except Exception as erro:
        messagebox.showerror(
            "Sistema de Cobrança",
            (
                "Não foi possível recuperar os dados salvos.\n\n"
                f"Detalhes: {erro}"
            ),
            parent=root,
        )
        root.destroy()
        return

    root.mainloop()


if __name__ == "__main__":
    main()
