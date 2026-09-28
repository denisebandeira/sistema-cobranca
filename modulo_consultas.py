"""Consulta de condominios, devedores e historico das cargas no Mac/desktop."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from decimal import Decimal, InvalidOperation
import re

import pandas as pd

from banco_dados import (
    desativar_meio_contato,
    definir_unidade_ajuizada, definir_unidades_ajuizadas,
    listar_contatos, listar_unidades_ajuizadas,
    listar_historico_comunicacoes,
    listar_meios_contato,
    registrar_comunicacao,
    salvar_contato,
    salvar_meio_contato,
    unidade_ajuizada,
)
from gerador_comunicados import (
    caminho_modelo_padrao,
    gerar_comunicado_individual_excel,
)
from cargas_dados import (
    carregar_base_condominio,
    listar_cargas_condominio,
    listar_condominios_atuais,
    listar_historico_debitos,
)


def _moeda(valor: object) -> str:
    if valor is None or valor == "":
        return "—"
    numero = float(valor)
    return f"R$ {numero:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")


def _data_carga(valor: str) -> str:
    return valor[:16].replace("T", " ")


def _ler_criterios_ajuizamento(
    debitos: str, valor: str,
) -> tuple[int | None, Decimal | None]:
    """Interpreta os limites informados em formato brasileiro."""

    debitos = debitos.strip()
    valor = valor.strip().removeprefix("R$").strip()
    if not debitos and not valor:
        raise ValueError("Informe ao menos um critério para ver as candidatas.")
    if debitos and not re.fullmatch(r"\d+", debitos):
        raise ValueError("Informe um número inteiro de débitos, sem sinal.")
    limite_debitos = int(debitos) if debitos else None
    limite_valor = None
    if valor:
        numero_br = re.fullmatch(
            r"(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d{1,2})?", valor
        )
        decimal_ponto = re.fullmatch(r"\d+\.\d{1,2}", valor)
        if not (numero_br or decimal_ponto):
            raise ValueError("Informe o valor em reais, como 1000,00 ou 1.000,00.")
        normalizado = (
            valor.replace(".", "").replace(",", ".")
            if numero_br else valor
        )
        try:
            limite_valor = Decimal(normalizado)
        except InvalidOperation as erro:
            raise ValueError("O valor corrigido informado é inválido.") from erro
    return limite_debitos, limite_valor


def _candidatas_por_criterios(
    resumo: pd.DataFrame,
    *,
    limite_debitos: int | None,
    limite_valor: Decimal | None,
    ajuizadas: set[str],
) -> set[str]:
    """Une os critérios (OU) e devolve economias únicas ainda não ajuizadas."""

    candidatas = set()
    for _, linha in resumo.iterrows():
        economia = str(linha["economia"])
        if economia in ajuizadas:
            continue
        atende_debitos = limite_debitos is not None and int(linha["count"]) > limite_debitos
        atende_valor = (
            limite_valor is not None
            and Decimal(str(linha["sum"])) > limite_valor
        )
        if atende_debitos or atende_valor:
            candidatas.add(economia)
    return candidatas


def _criar_tabela(parent, colunas, titulos, larguras, *, altura=8):
    quadro = ttk.Frame(parent)
    tabela = ttk.Treeview(
        quadro, columns=colunas, show="headings", height=altura,
        selectmode="browse",
    )
    for coluna in colunas:
        tabela.heading(coluna, text=titulos[coluna])
        tabela.column(
            coluna, width=larguras[coluna], minwidth=65, stretch=False,
            anchor="e" if coluna in {"valor", "total"} else "w",
        )
    barra = ttk.Scrollbar(quadro, orient="vertical", command=tabela.yview)
    barra_horizontal = ttk.Scrollbar(
        quadro, orient="horizontal", command=tabela.xview
    )
    tabela.configure(
        yscrollcommand=barra.set, xscrollcommand=barra_horizontal.set
    )
    quadro.columnconfigure(0, weight=1)
    quadro.rowconfigure(0, weight=1)
    tabela.grid(row=0, column=0, sticky="nsew")
    barra.grid(row=0, column=1, sticky="ns")
    barra_horizontal.grid(row=1, column=0, sticky="ew")
    return quadro, tabela


class JanelaConsultas:
    def __init__(self, parent: tk.Misc, *, incorporada: bool = False):
        self.janela = ttk.Frame(parent) if incorporada else tk.Toplevel(parent)
        if not incorporada:
            self.janela.title("Sistema de Cobrança — Condomínios e Devedores")
            largura = min(1200, self.janela.winfo_screenwidth() - 80)
            altura = min(780, self.janela.winfo_screenheight() - 120)
            esquerda_janela = max(0, (self.janela.winfo_screenwidth() - largura) // 2)
            topo_janela = max(0, (self.janela.winfo_screenheight() - altura) // 2)
            self.janela.geometry(
                f"{largura}x{altura}+{esquerda_janela}+{topo_janela}"
            )
            self.janela.minsize(min(920, largura), min(640, altura))
            self.janela.transient(parent)

        self.codigo_condominio = None
        self.cargas = []
        self.base_carga = pd.DataFrame()
        self.devedores = {}
        self.devedor_selecionado = None
        self.contato_selecionado = None
        self.meio_edicao_id = None
        self.meios_por_id = {}
        self.mostrar_apenas_nao_ajuizadas = tk.BooleanVar(value=False)
        self.ajuizada_sim = tk.BooleanVar(value=False)
        self.ajuizada_nao = tk.BooleanVar(value=True)
        self.limite_debitos = tk.StringVar()
        self.limite_valor = tk.StringVar()
        self.candidatas_ajuizamento: set[str] = set()
        self.previa_ajuizamento_ativa = False
        self.resumo_devedores = pd.DataFrame()

        topo = ttk.Frame(self.janela, padding=(12, 6, 12, 5))
        topo.pack(fill="x")
        ttk.Label(
            topo, text="Consulta", font=("Arial", 12, "bold"),
        ).pack(side="left")
        ttk.Button(
            topo, text="Atualizar consulta", command=self.atualizar_condominios,
        ).pack(side="right")

        # As duas colunas compartilham as mesmas linhas. Assim, contatos e
        # ficha do devedor sempre começam exatamente na mesma altura.
        corpo = ttk.Frame(self.janela)
        corpo.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        corpo.columnconfigure(0, weight=2, minsize=340)
        corpo.columnconfigure(1, weight=3, minsize=520)
        # A lista de devedores fica compacta; o espaço extra vai para a ficha
        # e para os contatos, mantendo os dois painéis inferiores alinhados.
        corpo.rowconfigure(1, weight=0)
        corpo.rowconfigure(2, weight=1)

        quadro_condominios = ttk.LabelFrame(
            corpo, text="1. Condomínios", padding=8
        )
        quadro_contatos = ttk.LabelFrame(
            corpo, text="Contatos do devedor", padding=8
        )
        quadro_condominios.grid(
            row=0, column=0, rowspan=2, sticky="nsew",
            padx=(0, 8), pady=(0, 8),
        )
        quadro_contatos.grid(
            row=2, column=0, sticky="nsew", padx=(0, 8)
        )

        quadro, self.tabela_condominios = _criar_tabela(
            quadro_condominios,
            ("codigo", "nome", "data"),
            {"codigo": "Código", "nome": "Condomínio", "data": "Última carga"},
            {"codigo": 70, "nome": 210, "data": 135},
            altura=7,
        )
        quadro.pack(fill="both", expand=True)
        self.tabela_condominios.bind(
            "<<TreeviewSelect>>", self.ao_selecionar_condominio
        )

        self.label_contatos = ttk.Label(
            quadro_contatos, text="Selecione um devedor para ver os contatos.",
            wraplength=310,
        )
        self.label_contatos.pack(fill="x", pady=(0, 5))
        quadro, self.tabela_meios = _criar_tabela(
            quadro_contatos,
            ("tipo", "valor", "observacoes"),
            {"tipo": "Tipo", "valor": "Telefone / e-mail", "observacoes": "Obs."},
            {"tipo": 75, "valor": 155, "observacoes": 390},
            altura=4,
        )
        quadro.pack(fill="both", expand=True)
        self.tabela_meios.bind("<<TreeviewSelect>>", self.ao_selecionar_meio)

        formulario_meio = ttk.Frame(quadro_contatos)
        formulario_meio.pack(fill="x", pady=(6, 0))
        formulario_meio.columnconfigure(1, weight=1)
        ttk.Label(formulario_meio, text="Tipo:").grid(row=0, column=0, sticky="w")
        self.tipo_meio = ttk.Combobox(
            formulario_meio, values=("Telefone", "E-mail"), state="readonly",
            width=11,
        )
        self.tipo_meio.current(0)
        self.tipo_meio.grid(row=0, column=1, sticky="w", padx=(5, 0))
        ttk.Label(formulario_meio, text="Valor:").grid(row=1, column=0, sticky="w")
        self.valor_meio = ttk.Entry(formulario_meio)
        self.valor_meio.grid(row=1, column=1, sticky="ew", padx=(5, 0))
        ttk.Label(formulario_meio, text="Obs.:").grid(row=2, column=0, sticky="w")
        self.obs_meio = ttk.Entry(formulario_meio)
        self.obs_meio.grid(row=2, column=1, sticky="ew", padx=(5, 0))
        botoes_meio = ttk.Frame(formulario_meio)
        botoes_meio.grid(row=3, column=0, columnspan=2, sticky="e", pady=(5, 0))
        self.botao_novo_meio = ttk.Button(
            botoes_meio, text="Novo", command=self.novo_meio, state="disabled"
        )
        self.botao_novo_meio.pack(side="left", padx=(0, 5))
        self.botao_salvar_meio = ttk.Button(
            botoes_meio, text="Salvar contato", command=self.salvar_meio,
            state="disabled",
        )
        self.botao_salvar_meio.pack(side="left")
        self.botao_excluir_meio = ttk.Button(
            botoes_meio, text="Excluir selecionado", command=self.excluir_meio,
            state="disabled",
        )
        self.botao_excluir_meio.pack(side="left", padx=(5, 0))

        escolha = ttk.LabelFrame(corpo, text="2. Carga do condomínio", padding=8)
        escolha.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        escolha.columnconfigure(0, weight=1)
        self.combo_cargas = ttk.Combobox(escolha, state="readonly")
        self.combo_cargas.grid(row=0, column=0, sticky="ew")
        self.combo_cargas.bind("<<ComboboxSelected>>", self.ao_selecionar_carga)
        self.label_arquivos = ttk.Label(
            escolha, text="Selecione um condomínio.", wraplength=700,
        )
        self.label_arquivos.grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Checkbutton(
            escolha, text="Mostrar somente unidades não ajuizadas",
            variable=self.mostrar_apenas_nao_ajuizadas,
            command=self.atualizar_lista_devedores,
        ).grid(row=2, column=0, sticky="w", pady=(5, 0))

        quadro_devedores = ttk.LabelFrame(
            corpo, text="3. Devedores da carga selecionada", padding=8
        )
        quadro_devedores.grid(row=1, column=1, sticky="nsew", pady=(0, 8))
        ttk.Label(
            quadro_devedores,
            text="Ajuizada: clique na caixa da linha para alternar NÃO/SIM. Salva na hora.",
            wraplength=600,
        ).pack(anchor="w", pady=(0, 4))
        self.botao_exibir_criterios = ttk.Button(
            quadro_devedores, text="Marcação em lote ▸",
            command=self.alternar_painel_lote,
        )
        self.botao_exibir_criterios.pack(anchor="w", pady=(0, 4))
        self.painel_lote = ttk.Frame(quadro_devedores)
        self.label_previa_ajuizamento = ttk.Label(
            self.painel_lote,
            text="Informe um critério para ver unidades ainda não ajuizadas.",
            wraplength=600,
        )
        self.label_previa_ajuizamento.pack(anchor="w", pady=(0, 4))
        criterios = ttk.Frame(self.painel_lote)
        criterios.pack(fill="x", pady=(0, 4))
        ttk.Label(criterios, text="Débitos >").grid(row=0, column=0, sticky="w")
        ttk.Entry(criterios, textvariable=self.limite_debitos, width=6).grid(
            row=0, column=1, sticky="w", padx=(4, 10)
        )
        ttk.Label(criterios, text="ou total corrigido > R$").grid(
            row=0, column=2, sticky="w"
        )
        ttk.Entry(criterios, textvariable=self.limite_valor, width=12).grid(
            row=0, column=3, sticky="w", padx=(4, 8)
        )
        ttk.Button(
            criterios, text="Ver candidatas", command=self.ver_candidatas_ajuizamento,
        ).grid(row=0, column=4, sticky="w")
        previa = ttk.Frame(self.painel_lote)
        previa.pack(fill="x", pady=(0, 4))
        self.botao_marcar_lote = ttk.Button(
            previa, text="Marcar candidatas como SIM",
            command=self.marcar_candidatas_ajuizadas, state="disabled",
        )
        self.botao_marcar_lote.pack(side="right")
        self.botao_limpar_selecao = ttk.Button(
            previa, text="Limpar seleção",
            command=self.limpar_selecao_ajuizamento, state="disabled",
        )
        self.botao_limpar_selecao.pack(side="right", padx=(0, 6))
        quadro, self.tabela_devedores = _criar_tabela(
            quadro_devedores,
            ("economia", "nome", "parcelas", "total", "ajuizada"),
            {
                "economia": "Economia", "nome": "Condômino",
                "parcelas": "Débitos", "total": "Total corrigido",
                "ajuizada": "Ajuizada",
            },
            {
                "economia": 95, "nome": 250, "parcelas": 65,
                "total": 115, "ajuizada": 90,
            },
            altura=7,
        )
        self.quadro_tabela_devedores = quadro
        quadro.pack(fill="both", expand=True)
        self.limite_debitos.trace_add("write", self._limpar_previa_ajuizamento)
        self.limite_valor.trace_add("write", self._limpar_previa_ajuizamento)
        self.tabela_devedores.bind(
            "<<TreeviewSelect>>", self.ao_selecionar_devedor
        )
        self.tabela_devedores.bind("<ButtonRelease-1>", self.alternar_ajuizada_na_lista)

        detalhes = ttk.LabelFrame(corpo, text="4. Ficha do devedor", padding=8)
        detalhes.grid(row=2, column=1, sticky="nsew")
        detalhes.columnconfigure(0, weight=1)

        self.label_devedor = ttk.Label(
            detalhes, text="Selecione um devedor.", font=("Arial", 11, "bold"),
        )
        self.label_devedor.grid(row=0, column=0, sticky="w", pady=(0, 5))

        situacao_unidade = ttk.Frame(detalhes)
        situacao_unidade.grid(row=1, column=0, sticky="w", pady=(0, 5))
        ttk.Label(situacao_unidade, text="Unidade ajuizada:").pack(side="left")
        self.check_ajuizada_sim = ttk.Checkbutton(
            situacao_unidade, text="SIM", variable=self.ajuizada_sim,
            command=lambda: self._selecionar_ajuizada_ficha(True),
        )
        self.check_ajuizada_sim.pack(side="left", padx=(8, 0))
        self.check_ajuizada_nao = ttk.Checkbutton(
            situacao_unidade, text="NÃO", variable=self.ajuizada_nao,
            command=lambda: self._selecionar_ajuizada_ficha(False),
        )
        self.check_ajuizada_nao.pack(side="left", padx=(8, 0))
        self.botao_salvar_ajuizada = ttk.Button(
            situacao_unidade, text="Salvar situação",
            command=self.salvar_ajuizada_ficha, state="disabled",
        )
        self.botao_salvar_ajuizada.pack(side="left", padx=(12, 0))

        acoes = ttk.Frame(detalhes)
        acoes.grid(row=2, column=0, sticky="w")
        self.botao_comunicado = ttk.Button(
            acoes, text="Salvar comunicado individual (Excel)",
            command=self.salvar_comunicado_individual, state="disabled",
        )
        self.botao_comunicado.pack(side="left", padx=(0, 8))
        self.botao_registrar = ttk.Button(
            acoes, text="Registrar contato com devedor",
            command=self.registrar_contato,
            state="disabled",
        )
        self.botao_registrar.pack(side="left")

        abas = ttk.Notebook(detalhes)
        abas.grid(row=3, column=0, sticky="nsew", pady=(6, 0))
        detalhes.rowconfigure(3, weight=1)

        aba_atual = ttk.Frame(abas, padding=5)
        aba_historico = ttk.Frame(abas, padding=5)
        aba_contatos = ttk.Frame(abas, padding=5)
        abas.add(aba_atual, text="Débitos nesta carga")
        abas.add(aba_historico, text="Histórico de débitos")
        abas.add(aba_contatos, text="Histórico de contatos")

        quadro, self.tabela_debitos = _criar_tabela(
            aba_atual,
            ("competencia", "vencimento", "valor"),
            {
                "competencia": "Competência", "vencimento": "Vencimento",
                "valor": "Valor corrigido",
            },
            {"competencia": 110, "vencimento": 120, "valor": 140},
        )
        quadro.pack(fill="both", expand=True)

        quadro, self.tabela_historico = _criar_tabela(
            aba_historico,
            ("competencia", "vencimento", "valor", "carga"),
            {
                "carga": "Data da carga", "competencia": "Competência",
                "vencimento": "Vencimento",
                "valor": "Valor corrigido",
            },
            {
                "carga": 145, "competencia": 100, "vencimento": 110,
                "valor": 120,
            },
        )
        quadro.pack(fill="both", expand=True)

        quadro, self.tabela_contatos = _criar_tabela(
            aba_contatos,
            ("data", "canal", "valor_canal", "acao", "resultado"),
            {
                "data": "Data", "canal": "Canal",
                "valor_canal": "Número / e-mail", "acao": "Ação",
                "resultado": "Resultado",
            },
            {"data": 145, "canal": 95, "valor_canal": 170,
             "acao": 160, "resultado": 180},
        )
        quadro.pack(fill="both", expand=True)

        self.atualizar_condominios()

    def novo_meio(self, *, focar: bool = True) -> None:
        self.meio_edicao_id = None
        self.botao_excluir_meio.config(state="disabled")
        selecionados = self.tabela_meios.selection()
        if selecionados:
            self.tabela_meios.selection_remove(*selecionados)
        self.tipo_meio.current(0)
        self.valor_meio.delete(0, "end")
        self.obs_meio.delete(0, "end")
        if focar:
            self.valor_meio.focus_set()

    def ao_selecionar_meio(self, _evento=None) -> None:
        selecao = self.tabela_meios.selection()
        if not selecao:
            return
        meio = self.meios_por_id.get(selecao[0])
        if meio is None:
            return
        self.meio_edicao_id = meio["id"]
        self.botao_excluir_meio.config(state="normal")
        self.tipo_meio.set(meio["tipo"])
        self.valor_meio.delete(0, "end")
        self.valor_meio.insert(0, meio["valor"])
        self.obs_meio.delete(0, "end")
        self.obs_meio.insert(0, meio["observacoes"] or "")

    def atualizar_meios(self) -> None:
        self._limpar_tabela(self.tabela_meios)
        self.botao_excluir_meio.config(state="disabled")
        self.meios_por_id = {}
        if not self.contato_selecionado:
            self.label_contatos.config(text="Selecione um devedor para ver os contatos.")
            return
        meios = listar_meios_contato(self.contato_selecionado["id"])
        self.label_contatos.config(
            text=(f"{len(meios)} contato(s) de {self.devedor_selecionado[1]}. "
                  "Selecione para editar ou clique em Novo.")
        )
        for meio in meios:
            iid = str(meio["id"])
            self.meios_por_id[iid] = meio
            self.tabela_meios.insert(
                "", "end", iid=iid,
                values=(meio["tipo"], meio["valor"], meio["observacoes"] or ""),
            )

    def salvar_meio(self) -> None:
        if self.devedor_selecionado is None:
            return
        try:
            if self.contato_selecionado is None:
                economia, nome = self.devedor_selecionado
                contato_id = salvar_contato(
                    self.codigo_condominio, economia, nome,
                    condominio=self.cargas[self.combo_cargas.current()]["condominio"],
                )
            else:
                contato_id = self.contato_selecionado["id"]
            salvar_meio_contato(
                contato_id, self.tipo_meio.get(), self.valor_meio.get(),
                observacoes=self.obs_meio.get(), meio_id=self.meio_edicao_id,
            )
        except (OSError, ValueError) as erro:
            messagebox.showerror("Contato", str(erro), parent=self.janela)
            return
        self.ao_selecionar_devedor()
        self.novo_meio(focar=False)

    def excluir_meio(self) -> None:
        if self.contato_selecionado is None or self.meio_edicao_id is None:
            return
        meio = self.meios_por_id.get(str(self.meio_edicao_id))
        if meio is None:
            return
        if not messagebox.askyesno(
            "Excluir contato",
            f"Excluir {meio['tipo'].lower()} {meio['valor']} da lista deste devedor?",
            parent=self.janela,
        ):
            return
        try:
            desativar_meio_contato(
                self.contato_selecionado["id"], self.meio_edicao_id
            )
        except (OSError, ValueError) as erro:
            messagebox.showerror("Contato", str(erro), parent=self.janela)
            return
        self.novo_meio(focar=False)
        self.atualizar_meios()

    def alternar_painel_lote(self) -> None:
        if self.painel_lote.winfo_manager():
            if self.previa_ajuizamento_ativa:
                self.limpar_selecao_ajuizamento()
            self.painel_lote.pack_forget()
            self.botao_exibir_criterios.config(text="Marcação em lote ▸")
        else:
            self.painel_lote.pack(
                before=self.quadro_tabela_devedores, fill="x", pady=(0, 4)
            )
            self.botao_exibir_criterios.config(text="Marcação em lote ▾")

    def _limpar_tabela(self, tabela: ttk.Treeview) -> None:
        for item in tabela.get_children():
            tabela.delete(item)

    def _limpar_previa_ajuizamento(self, *_args) -> None:
        if self.previa_ajuizamento_ativa:
            for iid in self.devedores:
                if self.tabela_devedores.exists(iid):
                    self.tabela_devedores.move(iid, "", "end")
        self.previa_ajuizamento_ativa = False
        self.candidatas_ajuizamento = set()
        self.botao_marcar_lote.config(state="disabled")
        self.botao_limpar_selecao.config(state="disabled")
        self.label_previa_ajuizamento.config(
            text="Informe um critério para ver unidades ainda não ajuizadas."
        )

    def limpar_selecao_ajuizamento(self) -> None:
        self.limite_debitos.set("")
        self.limite_valor.set("")
        self._limpar_previa_ajuizamento()
        # As linhas ocultas mantêm os valores antigos do Treeview; reconstrói
        # a lista lendo a situação atual de cada unidade no banco.
        self.atualizar_lista_devedores()

    def ver_candidatas_ajuizamento(self) -> None:
        self._limpar_previa_ajuizamento()
        if self.base_carga.empty or not self.codigo_condominio:
            self.label_previa_ajuizamento.config(text="Selecione primeiro uma carga.")
            return
        try:
            limite_debitos, limite_valor = _ler_criterios_ajuizamento(
                self.limite_debitos.get(), self.limite_valor.get()
            )
        except ValueError as erro:
            messagebox.showerror("Critérios", str(erro), parent=self.janela)
            return
        ajuizadas = {
            economia for codigo, economia in listar_unidades_ajuizadas()
            if codigo == self.codigo_condominio
        }
        self.candidatas_ajuizamento = _candidatas_por_criterios(
            self.resumo_devedores,
            limite_debitos=limite_debitos,
            limite_valor=limite_valor,
            ajuizadas=ajuizadas,
        )
        self.previa_ajuizamento_ativa = True
        for iid, (economia, _nome) in self.devedores.items():
            if economia not in self.candidatas_ajuizamento:
                self.tabela_devedores.detach(iid)
        self._limpar_devedor()
        quantidade = len(self.candidatas_ajuizamento)
        self.label_previa_ajuizamento.config(
            text=f"Exibindo somente {quantidade} unidade(s) candidata(s). Confira antes de marcar."
        )
        self.botao_marcar_lote.config(state="normal" if quantidade else "disabled")
        self.botao_limpar_selecao.config(state="normal")

    def marcar_candidatas_ajuizadas(self) -> None:
        if not self.candidatas_ajuizamento or not self.codigo_condominio:
            return
        try:
            quantidade = definir_unidades_ajuizadas(
                self.codigo_condominio, self.candidatas_ajuizamento, True
            )
        except (OSError, ValueError) as erro:
            messagebox.showerror("Situação das unidades", str(erro), parent=self.janela)
            return
        self.limpar_selecao_ajuizamento()
        self.label_previa_ajuizamento.config(
            text=f"{quantidade} unidade(s) marcada(s) como ajuizada(s)."
        )

    def _selecionar_ajuizada_ficha(self, ajuizada: bool) -> None:
        self.ajuizada_sim.set(ajuizada)
        self.ajuizada_nao.set(not ajuizada)

    def _limpar_devedor(self) -> None:
        self.devedor_selecionado = None
        self.contato_selecionado = None
        self.label_devedor.config(text="Selecione um devedor.")
        self.novo_meio(focar=False)
        self.botao_novo_meio.config(state="disabled")
        self.botao_salvar_meio.config(state="disabled")
        self.botao_excluir_meio.config(state="disabled")
        self.atualizar_meios()
        self.botao_comunicado.config(state="disabled")
        self.botao_registrar.config(state="disabled")
        self.botao_salvar_ajuizada.config(state="disabled")
        self._selecionar_ajuizada_ficha(False)
        for tabela in (
            self.tabela_debitos, self.tabela_historico, self.tabela_contatos
        ):
            self._limpar_tabela(tabela)

    def atualizar_condominios(self) -> None:
        selecionado = self.tabela_condominios.selection()
        codigo_anterior = selecionado[0] if selecionado else None
        self._limpar_tabela(self.tabela_condominios)
        condominios = listar_condominios_atuais()
        for item in condominios:
            self.tabela_condominios.insert(
                "", "end", iid=item["codigo_condominio"],
                values=(
                    item["codigo_condominio"], item["condominio"],
                    _data_carga(item["criada_em"]),
                ),
            )
        if codigo_anterior and self.tabela_condominios.exists(codigo_anterior):
            self.tabela_condominios.selection_set(codigo_anterior)
        elif condominios:
            self.tabela_condominios.selection_set(condominios[0]["codigo_condominio"])
        else:
            self.codigo_condominio = None
            self.combo_cargas["values"] = ()
            self.combo_cargas.set("")
            self.label_arquivos.config(
                text="Nenhuma carga salva. Processe os PDFs na janela principal."
            )
            self._limpar_tabela(self.tabela_devedores)
            self._limpar_devedor()
            return

        # selection_set não emite <<TreeviewSelect>> em todas as versões do Tk.
        # Recarrega explicitamente a carga e os devedores do condomínio ativo.
        self.ao_selecionar_condominio()

    def ao_selecionar_condominio(self, _evento=None) -> None:
        selecao = self.tabela_condominios.selection()
        if not selecao:
            return
        codigo = selecao[0]
        self.codigo_condominio = codigo
        self.cargas = listar_cargas_condominio(codigo)
        descricoes = [
            f"{_data_carga(item['criada_em'])} — "
            f"{item['quantidade_lancamentos']} lançamento(s)"
            for item in self.cargas
        ]
        self.combo_cargas["values"] = descricoes
        if self.cargas:
            self.combo_cargas.current(0)
            self.ao_selecionar_carga()

    def ao_selecionar_carga(self, _evento=None) -> None:
        indice = self.combo_cargas.current()
        if indice < 0 or indice >= len(self.cargas):
            return
        carga = self.cargas[indice]
        self.base_carga = carregar_base_condominio(carga["id"])
        arquivos = ", ".join(item["nome"] for item in carga["arquivos"])
        prefixo = "Carga atual" if indice == 0 else "Carga anterior"
        self.label_arquivos.config(text=f"{prefixo}. PDF(s): {arquivos}")
        self.atualizar_lista_devedores()

    def atualizar_lista_devedores(self) -> None:
        self._limpar_previa_ajuizamento()
        self._limpar_tabela(self.tabela_devedores)
        self._limpar_devedor()
        self.devedores = {}
        self.resumo_devedores = pd.DataFrame()
        if self.base_carga.empty:
            return
        ajuizadas = listar_unidades_ajuizadas()
        resumo = (
            self.base_carga.groupby(
                ["economia", "nome_condomino"], sort=True, dropna=False
            )["vlr_corrigido"]
            .agg(["sum", "count"])
            .reset_index()
        )
        self.resumo_devedores = resumo
        for numero, linha in resumo.iterrows():
            economia = str(linha["economia"])
            nome = str(linha["nome_condomino"])
            ajuizada = (self.codigo_condominio, economia) in ajuizadas
            if ajuizada and self.mostrar_apenas_nao_ajuizadas.get():
                continue
            iid = str(numero)
            self.devedores[iid] = (economia, nome)
            self.tabela_devedores.insert(
                "", "end", iid=iid,
                values=(
                    economia, nome, int(linha["count"]),
                    _moeda(linha["sum"]), "☑ SIM" if ajuizada else "☐ NÃO",
                ),
            )

    def alternar_ajuizada_na_lista(self, evento) -> None:
        if self.tabela_devedores.identify_column(evento.x) != "#5":
            return
        iid = self.tabela_devedores.identify_row(evento.y)
        if iid not in self.devedores:
            return
        economia, _nome = self.devedores[iid]
        novo_estado = not unidade_ajuizada(self.codigo_condominio, economia)
        try:
            definir_unidade_ajuizada(self.codigo_condominio, economia, novo_estado)
        except (OSError, ValueError) as erro:
            messagebox.showerror("Situação da unidade", str(erro), parent=self.janela)
            return
        if self.previa_ajuizamento_ativa and novo_estado:
            self.candidatas_ajuizamento.discard(economia)
            for linha_id, (unidade, _nome) in self.devedores.items():
                if unidade == economia:
                    self.tabela_devedores.detach(linha_id)
            self._limpar_devedor()
            quantidade = len(self.candidatas_ajuizamento)
            if quantidade == 0:
                self.limpar_selecao_ajuizamento()
                self.label_previa_ajuizamento.config(
                    text="Todas as candidatas foram marcadas. Lista completa atualizada."
                )
                return
            self.label_previa_ajuizamento.config(
                text=f"Exibindo somente {quantidade} unidade(s) candidata(s). Confira antes de marcar."
            )
            self.botao_marcar_lote.config(state="normal" if quantidade else "disabled")
            return
        selecionado = self.devedores[iid]
        self.atualizar_lista_devedores()
        for linha_id, devedor in self.devedores.items():
            if devedor == selecionado:
                self.tabela_devedores.selection_set(linha_id)
                self.ao_selecionar_devedor()
                break

    def salvar_ajuizada_ficha(self) -> None:
        if self.devedor_selecionado is None:
            return
        economia, _nome = self.devedor_selecionado
        try:
            definir_unidade_ajuizada(
                self.codigo_condominio, economia, self.ajuizada_sim.get()
            )
        except (OSError, ValueError) as erro:
            messagebox.showerror("Situação da unidade", str(erro), parent=self.janela)
            return
        selecionado = self.devedor_selecionado
        self.atualizar_lista_devedores()
        for iid, devedor in self.devedores.items():
            if devedor == selecionado:
                self.tabela_devedores.selection_set(iid)
                self.ao_selecionar_devedor()
                break

    def ao_selecionar_devedor(self, _evento=None) -> None:
        selecao = self.tabela_devedores.selection()
        if not selecao or selecao[0] not in self.devedores:
            return
        economia, nome = self.devedores[selecao[0]]
        self.devedor_selecionado = (economia, nome)
        self.contato_selecionado = next(
            (
                contato for contato in listar_contatos()
                if contato["codigo_condominio"] == self.codigo_condominio
                and contato["economia"] == economia
                and contato["condomino"].casefold() == nome.casefold()
            ),
            None,
        )

        situacao = ""
        if self.combo_cargas.current() > 0:
            atual = carregar_base_condominio(self.cargas[0]["id"])
            presente = (
                (atual["economia"] == economia)
                & (atual["nome_condomino"].str.casefold() == nome.casefold())
            ).any()
            if not presente:
                situacao = " — Não consta na carga mais recente"
        self.label_devedor.config(
            text=f"{nome} — Economia {economia}{situacao}"
        )
        self.novo_meio(focar=False)
        self.botao_novo_meio.config(state="normal")
        self.botao_salvar_meio.config(state="normal")
        self.atualizar_meios()
        ajuizada = unidade_ajuizada(self.codigo_condominio, economia)
        self._selecionar_ajuizada_ficha(ajuizada)
        self.botao_salvar_ajuizada.config(state="normal")
        self.botao_comunicado.config(state="disabled" if ajuizada else "normal")
        self.botao_registrar.config(state="normal")

        for tabela in (
            self.tabela_debitos, self.tabela_historico, self.tabela_contatos
        ):
            self._limpar_tabela(tabela)

        selecionados = self.base_carga.loc[
            (self.base_carga["economia"] == economia)
            & (self.base_carga["nome_condomino"] == nome)
        ]
        for _, linha in selecionados.iterrows():
            self.tabela_debitos.insert(
                "", "end",
                values=(
                    linha["competencia"], linha["data_vencimento"],
                    _moeda(linha["vlr_corrigido"]),
                ),
            )

        for lancamento in listar_historico_debitos(
            self.codigo_condominio, economia, nome
        ):
            self.tabela_historico.insert(
                "", "end",
                values=(
                    lancamento["competencia"],
                    lancamento["data_vencimento"],
                    _moeda(lancamento["vlr_corrigido"]),
                    _data_carga(lancamento["criada_em"]),
                ),
            )

        for comunicacao in listar_historico_comunicacoes(
            codigo_condominio=self.codigo_condominio,
            economia=economia,
            condomino=nome,
        ):
            self.tabela_contatos.insert(
                "", "end",
                values=(
                    _data_carga(comunicacao["realizado_em"]),
                    comunicacao["canal"], comunicacao["valor_canal"] or "",
                    comunicacao["acao"],
                    comunicacao["resultado"] or "",
                ),
            )

    def salvar_comunicado_individual(self) -> None:
        if self.devedor_selecionado is None:
            return
        economia, nome = self.devedor_selecionado
        if unidade_ajuizada(self.codigo_condominio, economia):
            messagebox.showerror(
                "Comunicado", "A unidade está ajuizada e não pode receber comunicado.",
                parent=self.janela,
            )
            return
        nome_seguro = re.sub(r"[^\w-]+", "_", nome, flags=re.UNICODE).strip("_")[:45]
        economia_segura = re.sub(r"[^\w-]+", "_", economia, flags=re.UNICODE).strip("_")[:20]
        destino = filedialog.asksaveasfilename(
            parent=self.janela,
            title="Salvar comunicado individual",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
            initialfile=f"comunicado_{self.codigo_condominio}_{economia_segura}_{nome_seguro}.xlsx",
        )
        if not destino:
            return
        try:
            conteudo = gerar_comunicado_individual_excel(
                self.base_carga, caminho_modelo_padrao(),
                codigo_condominio=self.codigo_condominio,
                economia=economia, nome_condomino=nome,
                unidades_ajuizadas=listar_unidades_ajuizadas(),
            )
            Path(destino).write_bytes(conteudo)
        except (OSError, ValueError) as erro:
            messagebox.showerror("Comunicado", str(erro), parent=self.janela)
            return
        messagebox.showinfo(
            "Comunicado salvo",
            "O Excel individual foi salvo. Confira os valores antes de usar.\n"
            "O salvamento não registra um envio ao devedor.",
            parent=self.janela,
        )

    def registrar_contato(self) -> None:
        if self.devedor_selecionado is None:
            return
        dialogo = tk.Toplevel(self.janela)
        dialogo.title("Registrar contato manualmente")
        dialogo.transient(self.janela)
        dialogo.resizable(False, False)
        quadro = ttk.Frame(dialogo, padding=16)
        quadro.pack(fill="both", expand=True)
        ttk.Label(quadro, text="Canal:").grid(row=0, column=0, sticky="w")
        canal = ttk.Combobox(
            quadro, values=("Telefone", "E-mail", "SMS", "Carta", "Outro"),
            state="readonly", width=25,
        )
        canal.current(0)
        canal.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(quadro, text="Número / e-mail usado:").grid(
            row=1, column=0, sticky="w"
        )
        valor_canal = ttk.Combobox(quadro, width=38)
        valor_canal.grid(row=1, column=1, sticky="ew", pady=4)

        def atualizar_opcoes(_evento=None) -> None:
            tipo = "E-mail" if canal.get() == "E-mail" else "Telefone"
            meios = (
                listar_meios_contato(self.contato_selecionado["id"])
                if self.contato_selecionado else []
            )
            valores = [
                meio["valor"] for meio in meios
                if meio["tipo"] == tipo
            ] if canal.get() in {"Telefone", "E-mail", "SMS"} else []
            valor_canal["values"] = valores
            valor_canal.set(valores[0] if valores else "")

        canal.bind("<<ComboboxSelected>>", atualizar_opcoes)
        atualizar_opcoes()

        ttk.Label(quadro, text="Ocorrência:").grid(row=2, column=0, sticky="w")
        acao = ttk.Combobox(
            quadro,
            values=(
                "Tentativa sem resposta", "Contato realizado",
                "Cliente solicitou boleto", "Boleto enviado ao cliente",
                "Cliente informou pagamento", "Sem retorno do cliente",
                "Mensagem enviada", "Ligação realizada", "Carta enviada", "Outro",
            ),
            state="readonly", width=25,
        )
        acao.current(0)
        acao.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(quadro, text="Resultado/observação:").grid(row=3, column=0, sticky="w")
        resultado = ttk.Entry(quadro, width=42)
        resultado.grid(row=3, column=1, sticky="ew", pady=4)

        def confirmar() -> None:
            economia, nome = self.devedor_selecionado
            if canal.get() in {"Telefone", "E-mail", "SMS"} and not valor_canal.get().strip():
                messagebox.showerror(
                    "Registro", "Informe o número ou e-mail utilizado.",
                    parent=dialogo,
                )
                return
            try:
                registrar_comunicacao(
                    self.codigo_condominio, economia, nome,
                    canal.get(), acao.get(),
                    contato_id=(
                        self.contato_selecionado["id"]
                        if self.contato_selecionado else None
                    ),
                    valor_canal=valor_canal.get(),
                    resultado=resultado.get(),
                )
            except (OSError, ValueError) as erro:
                messagebox.showerror("Registro", str(erro), parent=dialogo)
                return
            dialogo.destroy()
            self.ao_selecionar_devedor()

        botoes = ttk.Frame(quadro)
        botoes.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botoes, text="Cancelar", command=dialogo.destroy).pack(side="right")
        ttk.Button(botoes, text="Registrar", command=confirmar).pack(
            side="right", padx=(0, 8)
        )
        # Aproxima a janela da ficha e da tabela de débitos, sem sair da tela.
        dialogo.update_idletasks()
        ancora = self.tabela_debitos
        largura = dialogo.winfo_reqwidth()
        altura = dialogo.winfo_reqheight()
        x = max(0, min(
            ancora.winfo_rootx() + 20,
            dialogo.winfo_screenwidth() - largura - 20,
        ))
        y = max(0, min(
            ancora.winfo_rooty() + 20,
            dialogo.winfo_screenheight() - altura - 50,
        ))
        dialogo.geometry(f"+{x}+{y}")
        resultado.focus_set()
        dialogo.grab_set()


def abrir_janela_consultas(parent: tk.Misc) -> JanelaConsultas:
    return JanelaConsultas(parent)
