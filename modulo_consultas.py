"""Consulta de condominios, devedores e historico das cargas no Mac/desktop."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import re

import pandas as pd

from banco_dados import (
    listar_contatos,
    listar_historico_comunicacoes,
    listar_meios_contato,
    registrar_comunicacao,
    salvar_contato,
    salvar_meio_contato,
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

        topo = ttk.Frame(self.janela, padding=(16, 14, 16, 10))
        topo.pack(fill="x")
        ttk.Label(
            topo, text="Condomínios e devedores", font=("Arial", 18, "bold"),
        ).pack(side="left")
        ttk.Button(
            topo, text="Atualizar consulta", command=self.atualizar_condominios,
        ).pack(side="right")

        # As duas colunas compartilham as mesmas linhas. Assim, contatos e
        # ficha do devedor sempre começam exatamente na mesma altura.
        corpo = ttk.Frame(self.janela)
        corpo.pack(fill="both", expand=True, padx=16, pady=(0, 12))
        corpo.columnconfigure(0, weight=2, minsize=340)
        corpo.columnconfigure(1, weight=3, minsize=520)
        corpo.rowconfigure(1, weight=1)
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
            altura=8,
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
            {"tipo": 75, "valor": 155, "observacoes": 105},
            altura=5,
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

        quadro_devedores = ttk.LabelFrame(
            corpo, text="3. Devedores da carga selecionada", padding=8
        )
        quadro_devedores.grid(row=1, column=1, sticky="nsew", pady=(0, 8))
        quadro, self.tabela_devedores = _criar_tabela(
            quadro_devedores,
            ("economia", "nome", "parcelas", "total"),
            {
                "economia": "Economia", "nome": "Condômino",
                "parcelas": "Débitos", "total": "Total corrigido",
            },
            {
                "economia": 95, "nome": 250, "parcelas": 65,
                "total": 115,
            },
            altura=9,
        )
        quadro.pack(fill="both", expand=True)
        self.tabela_devedores.bind(
            "<<TreeviewSelect>>", self.ao_selecionar_devedor
        )

        detalhes = ttk.LabelFrame(corpo, text="4. Ficha do devedor", padding=8)
        detalhes.grid(row=2, column=1, sticky="nsew")
        detalhes.columnconfigure(0, weight=1)

        self.label_devedor = ttk.Label(
            detalhes, text="Selecione um devedor.", font=("Arial", 11, "bold"),
        )
        self.label_devedor.grid(row=0, column=0, sticky="w", pady=(0, 5))

        acoes = ttk.Frame(detalhes)
        acoes.grid(row=1, column=0, sticky="w")
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
        abas.grid(row=2, column=0, sticky="nsew", pady=(6, 0))
        detalhes.rowconfigure(2, weight=1)

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
        self.tipo_meio.set(meio["tipo"])
        self.valor_meio.delete(0, "end")
        self.valor_meio.insert(0, meio["valor"])
        self.obs_meio.delete(0, "end")
        self.obs_meio.insert(0, meio["observacoes"] or "")

    def atualizar_meios(self) -> None:
        self._limpar_tabela(self.tabela_meios)
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

    def _limpar_tabela(self, tabela: ttk.Treeview) -> None:
        for item in tabela.get_children():
            tabela.delete(item)

    def _limpar_devedor(self) -> None:
        self.devedor_selecionado = None
        self.contato_selecionado = None
        self.label_devedor.config(text="Selecione um devedor.")
        self.novo_meio(focar=False)
        self.botao_novo_meio.config(state="disabled")
        self.botao_salvar_meio.config(state="disabled")
        self.atualizar_meios()
        self.botao_comunicado.config(state="disabled")
        self.botao_registrar.config(state="disabled")
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
        self._limpar_tabela(self.tabela_devedores)
        self._limpar_devedor()
        self.devedores = {}

        if self.base_carga.empty:
            return
        resumo = (
            self.base_carga.groupby(
                ["economia", "nome_condomino"], sort=True, dropna=False
            )["vlr_corrigido"]
            .agg(["sum", "count"])
            .reset_index()
        )
        for numero, linha in resumo.iterrows():
            economia = str(linha["economia"])
            nome = str(linha["nome_condomino"])
            iid = str(numero)
            self.devedores[iid] = (economia, nome)
            self.tabela_devedores.insert(
                "", "end", iid=iid,
                values=(
                    economia, nome, int(linha["count"]),
                    _moeda(linha["sum"]),
                ),
            )

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
        self.botao_comunicado.config(state="normal")
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

        ttk.Label(quadro, text="Ação realizada:").grid(row=2, column=0, sticky="w")
        acao = ttk.Combobox(
            quadro,
            values=(
                "Tentativa de contato", "Mensagem enviada",
                "Ligação realizada", "Carta enviada", "Resposta recebida", "Outro",
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
        resultado.focus_set()
        dialogo.grab_set()


def abrir_janela_consultas(parent: tk.Misc) -> JanelaConsultas:
    return JanelaConsultas(parent)
