"""Tela de contatos do Modulo 3 e importacao de telefones."""

from __future__ import annotations

from pathlib import Path
import re
import shutil
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import unicodedata

import pandas as pd
from modulo_consultas import abrir_janela_consultas

from banco_dados import (
    atualizar_contato,
    importar_telefones,
    listar_contatos,
    normalizar_telefone,
)


NOME_MODELO = "modelo_importacao_contatos.xlsx"

ALIASES_COLUNAS = {
    "codigo_condominio": {
        "codigo_condominio", "codigo_do_condominio", "cod_condominio",
        "condominio_codigo", "codigo",
    },
    "condominio": {"condominio", "nome_condominio"},
    "economia": {"economia", "unidade", "apartamento", "apto"},
    "nome_condomino": {
        "nome_condomino", "condomino", "nome_do_condomino", "nome",
    },
    "telefone": {"telefone", "celular", "whatsapp", "fone"},
    "observacoes": {"observacoes", "observacao", "obs"},
}


def _normalizar_cabecalho(valor: object) -> str:
    texto = str(valor).strip()
    texto = "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caractere)
    )
    return re.sub(r"[^a-z0-9]+", "_", texto.casefold()).strip("_")


def caminho_modelo_importacao() -> Path:
    if hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / NOME_MODELO
    return Path(__file__).resolve().parent / NOME_MODELO


def ler_planilha_contatos(caminho: str | Path) -> dict:
    """Le XLSX ou CSV e devolve registros com cabecalhos padronizados."""

    arquivo = Path(caminho)
    extensao = arquivo.suffix.casefold()
    if extensao == ".xlsx":
        tabela = pd.read_excel(
            arquivo,
            dtype=str,
            keep_default_na=False,
        )
    elif extensao == ".csv":
        try:
            tabela = pd.read_csv(
                arquivo,
                dtype=str,
                keep_default_na=False,
                sep=None,
                engine="python",
                encoding="utf-8-sig",
            )
        except UnicodeDecodeError:
            tabela = pd.read_csv(
                arquivo,
                dtype=str,
                keep_default_na=False,
                sep=None,
                engine="python",
                encoding="latin-1",
            )
    else:
        raise ValueError("Selecione uma planilha .xlsx ou .csv.")

    aliases_invertidos = {
        alias: destino
        for destino, aliases in ALIASES_COLUNAS.items()
        for alias in aliases
    }
    renomeadas = {}
    destinos_encontrados = set()
    for coluna in tabela.columns:
        normalizada = _normalizar_cabecalho(coluna)
        destino = aliases_invertidos.get(normalizada)
        if destino:
            if destino in destinos_encontrados:
                raise ValueError(
                    f"A planilha possui mais de uma coluna para {destino}."
                )
            destinos_encontrados.add(destino)
            renomeadas[coluna] = destino

    tabela = tabela.rename(columns=renomeadas)
    obrigatorias = {"codigo_condominio", "economia", "telefone"}
    ausentes = obrigatorias - set(tabela.columns)
    if ausentes:
        raise ValueError(
            "Colunas obrigatorias ausentes: " + ", ".join(sorted(ausentes))
        )

    registros = []
    erros = []
    for indice, linha in tabela.iterrows():
        numero_linha = int(indice) + 2
        registro = {
            coluna: str(linha.get(coluna, "")).strip()
            for coluna in ALIASES_COLUNAS
        }
        if not any(registro.values()):
            continue
        registro["_linha"] = numero_linha

        faltantes = [
            coluna
            for coluna in obrigatorias
            if not registro[coluna]
        ]
        if faltantes:
            erros.append(
                {
                    "linha": numero_linha,
                    "erro": "campos vazios: " + ", ".join(sorted(faltantes)),
                }
            )
            continue

        try:
            registro["telefone"] = normalizar_telefone(registro["telefone"])
        except ValueError as erro:
            erros.append({"linha": numero_linha, "erro": str(erro)})
            continue

        registros.append(registro)

    return {"registros": registros, "erros": erros}


class JanelaComunicacoes:
    def __init__(self, parent: tk.Misc):
        self.janela = tk.Toplevel(parent)
        self.janela.title("Sistema de Cobrança — Comunicações")
        largura = min(1050, self.janela.winfo_screenwidth() - 80)
        altura = min(680, self.janela.winfo_screenheight() - 120)
        esquerda_janela = max(0, (self.janela.winfo_screenwidth() - largura) // 2)
        topo_janela = max(0, (self.janela.winfo_screenheight() - altura) // 2)
        self.janela.geometry(
            f"{largura}x{altura}+{esquerda_janela}+{topo_janela}"
        )
        self.janela.minsize(min(820, largura), min(560, altura))
        self.janela.transient(parent)

        self.contatos_por_id = {}

        topo = ttk.Frame(self.janela, padding=(18, 16, 18, 8))
        topo.pack(fill="x")

        ttk.Label(
            topo,
            text="Módulo 3 — Comunicações",
            font=("Arial", 18, "bold"),
        ).pack(side="left")

        ttk.Button(
            topo,
            text="Salvar modelo de planilha",
            command=self.salvar_modelo,
        ).pack(side="right", padx=(8, 0))

        ttk.Button(
            topo,
            text="Importar telefones",
            command=self.importar_planilha,
        ).pack(side="right")

        ttk.Button(
            topo,
            text="Condomínios e devedores",
            command=lambda: abrir_janela_consultas(self.janela),
        ).pack(side="right", padx=(0, 8))

        ttk.Label(
            self.janela,
            text=(
                "Selecione um contato para cadastrar ou corrigir o telefone. "
                "A importação aceita arquivos XLSX e CSV."
            ),
            padding=(18, 0, 18, 8),
        ).pack(fill="x")

        quadro_tabela = ttk.Frame(self.janela, padding=(18, 0, 18, 10))
        quadro_tabela.pack(fill="both", expand=True)

        colunas = (
            "codigo", "condominio", "economia", "condomino",
            "telefone", "observacoes",
        )
        self.tabela = ttk.Treeview(
            quadro_tabela,
            columns=colunas,
            show="headings",
            selectmode="browse",
        )
        titulos = {
            "codigo": "Código",
            "condominio": "Condomínio",
            "economia": "Economia",
            "condomino": "Condômino",
            "telefone": "Telefone",
            "observacoes": "Observações",
        }
        larguras = {
            "codigo": 75,
            "condominio": 210,
            "economia": 85,
            "condomino": 260,
            "telefone": 130,
            "observacoes": 190,
        }
        for coluna in colunas:
            self.tabela.heading(coluna, text=titulos[coluna])
            self.tabela.column(
                coluna,
                width=larguras[coluna],
                anchor="center" if coluna in {"codigo", "economia"} else "w",
            )

        barra = ttk.Scrollbar(
            quadro_tabela,
            orient="vertical",
            command=self.tabela.yview,
        )
        self.tabela.configure(yscrollcommand=barra.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        self.tabela.bind("<<TreeviewSelect>>", self.ao_selecionar)

        formulario = ttk.LabelFrame(
            self.janela,
            text="Contato selecionado",
            padding=12,
        )
        formulario.pack(fill="x", padx=18, pady=(0, 12))
        formulario.columnconfigure(3, weight=1)

        self.label_selecionado = ttk.Label(formulario, text="Nenhum contato selecionado.")
        self.label_selecionado.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))

        ttk.Label(formulario, text="Telefone:").grid(row=1, column=0, sticky="w")
        self.campo_telefone = ttk.Entry(formulario, width=22)
        self.campo_telefone.grid(row=1, column=1, sticky="w", padx=(6, 18))

        ttk.Label(formulario, text="Observações:").grid(row=1, column=2, sticky="w")
        self.campo_observacoes = ttk.Entry(formulario)
        self.campo_observacoes.grid(row=1, column=3, sticky="ew", padx=(6, 12))

        self.botao_salvar = ttk.Button(
            formulario,
            text="Salvar contato",
            command=self.salvar_contato_selecionado,
            state="disabled",
        )
        self.botao_salvar.grid(row=1, column=4, sticky="e")

        self.label_status = ttk.Label(
            self.janela,
            text="",
            anchor="center",
            padding=(12, 0, 12, 12),
        )
        self.label_status.pack(fill="x")

        self.atualizar_tabela()

    def atualizar_tabela(self) -> None:
        selecionado = self.tabela.selection()
        selecionado_id = selecionado[0] if selecionado else None
        for item in self.tabela.get_children():
            self.tabela.delete(item)

        contatos = listar_contatos()
        self.contatos_por_id = {str(contato["id"]): contato for contato in contatos}
        com_telefone = 0
        for contato in contatos:
            telefone = contato["telefone"] or ""
            if telefone:
                com_telefone += 1
            identificador = str(contato["id"])
            self.tabela.insert(
                "",
                "end",
                iid=identificador,
                values=(
                    contato["codigo_condominio"],
                    contato["condominio"] or "",
                    contato["economia"],
                    contato["condomino"],
                    telefone,
                    contato["observacoes"] or "",
                ),
            )

        if selecionado_id in self.contatos_por_id:
            self.tabela.selection_set(selecionado_id)
            self.tabela.see(selecionado_id)

        self.label_status.config(
            text=(
                f"{len(contatos)} contato(s): {com_telefone} com telefone e "
                f"{len(contatos) - com_telefone} sem telefone."
            )
        )

    def ao_selecionar(self, _evento=None) -> None:
        selecionados = self.tabela.selection()
        if not selecionados:
            return
        contato = self.contatos_por_id[selecionados[0]]
        self.label_selecionado.config(
            text=(
                f"{contato['condomino']} — Condomínio "
                f"{contato['codigo_condominio']}, economia {contato['economia']}"
            )
        )
        self.campo_telefone.delete(0, "end")
        self.campo_telefone.insert(0, contato["telefone"] or "")
        self.campo_observacoes.delete(0, "end")
        self.campo_observacoes.insert(0, contato["observacoes"] or "")
        self.botao_salvar.config(state="normal")

    def salvar_contato_selecionado(self) -> None:
        selecionados = self.tabela.selection()
        if not selecionados:
            return
        try:
            atualizar_contato(
                int(selecionados[0]),
                self.campo_telefone.get(),
                observacoes=self.campo_observacoes.get(),
            )
        except ValueError as erro:
            messagebox.showerror("Telefone inválido", str(erro), parent=self.janela)
            return
        self.atualizar_tabela()
        messagebox.showinfo(
            "Contato salvo",
            "O contato foi atualizado com sucesso.",
            parent=self.janela,
        )

    def importar_planilha(self) -> None:
        arquivo = filedialog.askopenfilename(
            title="Selecionar planilha de telefones",
            filetypes=[
                ("Planilhas aceitas", "*.xlsx *.csv"),
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
            ],
            parent=self.janela,
        )
        if not arquivo:
            return
        try:
            leitura = ler_planilha_contatos(arquivo)
            resultado = importar_telefones(leitura["registros"])
            erros = leitura["erros"] + resultado["erros"]
        except Exception as erro:
            messagebox.showerror(
                "Erro na importação",
                f"Não foi possível importar a planilha.\n\n{erro}",
                parent=self.janela,
            )
            return

        self.atualizar_tabela()
        mensagem = f"Telefones atualizados: {resultado['atualizados']}."
        if erros:
            detalhes = "\n".join(
                f"Linha {item['linha']}: {item['erro']}"
                for item in erros[:10]
            )
            if len(erros) > 10:
                detalhes += f"\n... e mais {len(erros) - 10} erro(s)."
            mensagem += f"\n\nLinhas não importadas: {len(erros)}.\n{detalhes}"

        messagebox.showinfo("Importação concluída", mensagem, parent=self.janela)

    def salvar_modelo(self) -> None:
        destino = filedialog.asksaveasfilename(
            title="Salvar modelo de importação",
            defaultextension=".xlsx",
            initialfile=NOME_MODELO,
            filetypes=[("Arquivo Excel", "*.xlsx")],
            parent=self.janela,
        )
        if not destino:
            return
        try:
            shutil.copyfile(caminho_modelo_importacao(), destino)
        except Exception as erro:
            messagebox.showerror(
                "Erro",
                f"Não foi possível salvar o modelo.\n\n{erro}",
                parent=self.janela,
            )
            return
        messagebox.showinfo(
            "Modelo salvo",
            "O modelo de importação foi salvo com sucesso.",
            parent=self.janela,
        )


def abrir_janela_comunicacoes(parent: tk.Misc) -> JanelaComunicacoes:
    return JanelaComunicacoes(parent)
