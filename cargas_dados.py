"""Cargas historicas dos PDFs e Base Mestre atual por condominio.

O parser continua sendo a origem dos dados. Este modulo recebe apenas seus
DataFrames ja validados e preserva cada versao no SQLite.
"""

from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import hashlib
import json
from pathlib import Path
from typing import Iterable, Mapping

import pandas as pd

from banco_dados import conectar


COLUNAS_BASE = (
    "codigo_condominio", "condominio", "economia", "nome_condomino",
    "competencia", "data_vencimento", "tipo", "vlr_base", "multa",
    "juros", "correcao", "vlr_corrigido", "assessor", "imobiliaria",
    "endereco", "nosso_numero", "advogado",
)
COLUNAS_LANCAMENTO = COLUNAS_BASE[2:]
COLUNAS_VALORES = {"vlr_base", "multa", "juros", "correcao", "vlr_corrigido"}


def _texto(valor: object) -> str:
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (date, datetime)):
        return valor.isoformat()
    return str(valor).strip()


def _valor(valor: object) -> str | None:
    texto = _texto(valor)
    if not texto:
        return None
    try:
        numero = Decimal(texto)
    except InvalidOperation as erro:
        raise ValueError(f"Valor financeiro inválido: {texto}") from erro
    if not numero.is_finite():
        raise ValueError(f"Valor financeiro inválido: {texto}")
    return str(numero)


def _preparar_fontes(fontes: Iterable[Mapping]) -> dict[str, dict]:
    agrupadas: dict[str, dict] = {}
    hashes_por_condominio: dict[str, set[str]] = defaultdict(set)

    for fonte in fontes:
        dados = fonte["dados"]
        if not isinstance(dados, pd.DataFrame):
            raise TypeError("A fonte deve conter um DataFrame validado pelo parser.")
        ausentes = set(COLUNAS_BASE) - set(dados.columns)
        if ausentes:
            raise ValueError(
                "Colunas ausentes na Base Mestre: " + ", ".join(sorted(ausentes))
            )
        if dados.empty:
            raise ValueError("Um PDF validado não pode ter zero lançamentos.")

        codigo = _texto(fonte["codigo_condominio"])
        nome = _texto(fonte["condominio"])
        arquivo = Path(fonte["arquivo"])
        if not codigo or not nome:
            raise ValueError("Condomínio sem código ou nome.")
        digest = hashlib.sha256(arquivo.read_bytes()).hexdigest()
        if digest in hashes_por_condominio[codigo]:
            continue
        hashes_por_condominio[codigo].add(digest)

        grupo = agrupadas.setdefault(
            codigo,
            {"condominio": nome, "arquivos": [], "linhas": []},
        )
        if grupo["condominio"] != nome:
            raise ValueError(f"Nomes diferentes para o condomínio {codigo}.")
        grupo["arquivos"].append({"nome": arquivo.name, "sha256": digest})

        for registro in dados.to_dict(orient="records"):
            if _texto(registro["codigo_condominio"]) != codigo:
                raise ValueError(
                    f"O PDF {arquivo.name} contém outro código de condomínio."
                )
            if not _texto(registro["economia"]) or not _texto(registro["nome_condomino"]):
                raise ValueError(
                    f"O PDF {arquivo.name} contém devedor sem economia ou nome."
                )
            linha = tuple(
                _valor(registro[coluna]) if coluna in COLUNAS_VALORES
                else _texto(registro[coluna])
                for coluna in COLUNAS_LANCAMENTO
            )
            grupo["linhas"].append(linha)

    for grupo in agrupadas.values():
        hashes = sorted(arquivo["sha256"] for arquivo in grupo["arquivos"])
        grupo["conteudo_sha256"] = hashlib.sha256(
            "\n".join(hashes).encode("ascii")
        ).hexdigest()
    return agrupadas


def registrar_carga(
    fontes: Iterable[Mapping],
    *,
    caminho: str | Path | None = None,
) -> dict:
    """Grava versoes novas e ignora PDFs identicos a ultima carga do condominio.

    Toda a gravacao ocorre em uma transacao. Um erro nao deixa carga parcial.
    """

    grupos = _preparar_fontes(fontes)
    if not grupos:
        return {"carga_id": None, "atualizados": 0, "repetidos": 0}

    with conectar(caminho) as conexao:
        novos = []
        repetidos = 0
        for codigo, grupo in grupos.items():
            anterior = conexao.execute(
                """
                SELECT conteudo_sha256 FROM cargas_condominio
                WHERE codigo_condominio = ?
                ORDER BY id DESC LIMIT 1
                """,
                (codigo,),
            ).fetchone()
            if anterior and anterior["conteudo_sha256"] == grupo["conteudo_sha256"]:
                repetidos += 1
            else:
                novos.append((codigo, grupo))

        if not novos:
            return {"carga_id": None, "atualizados": 0, "repetidos": repetidos}

        criada_em = datetime.now().astimezone().isoformat(timespec="seconds")
        carga_id = conexao.execute(
            "INSERT INTO cargas (criada_em) VALUES (?)", (criada_em,)
        ).lastrowid
        colunas_sql = ", ".join(COLUNAS_LANCAMENTO)
        marcadores = ", ".join("?" for _ in COLUNAS_LANCAMENTO)

        for codigo, grupo in novos:
            carga_condominio_id = conexao.execute(
                """
                INSERT INTO cargas_condominio (
                    carga_id, codigo_condominio, condominio, arquivos_json,
                    conteudo_sha256, quantidade_lancamentos
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    carga_id, codigo, grupo["condominio"],
                    json.dumps(grupo["arquivos"], ensure_ascii=False),
                    grupo["conteudo_sha256"], len(grupo["linhas"]),
                ),
            ).lastrowid

            conexao.executemany(
                f"""
                INSERT INTO lancamentos_carga (
                    carga_condominio_id, ordem, {colunas_sql}
                ) VALUES (?, ?, {marcadores})
                """,
                (
                    (carga_condominio_id, ordem, *linha)
                    for ordem, linha in enumerate(grupo["linhas"], start=1)
                ),
            )

    return {
        "carga_id": int(carga_id),
        "atualizados": len(novos),
        "repetidos": repetidos,
    }


def listar_condominios_atuais(*, caminho: str | Path | None = None) -> list[dict]:
    """Mostra somente a versao mais recente de cada condominio."""

    with conectar(caminho) as conexao:
        linhas = conexao.execute(
            """
            SELECT cc.id, cc.codigo_condominio, cc.condominio,
                   cc.quantidade_lancamentos, c.criada_em,
                   cc.arquivos_json
            FROM cargas_condominio cc
            JOIN cargas c ON c.id = cc.carga_id
            WHERE cc.id = (
                SELECT MAX(mais_recente.id)
                FROM cargas_condominio mais_recente
                WHERE mais_recente.codigo_condominio = cc.codigo_condominio
            )
            ORDER BY cc.codigo_condominio
            """
        ).fetchall()
    return [
        {**dict(linha), "arquivos": json.loads(linha["arquivos_json"])}
        for linha in linhas
    ]


def carregar_base_atual(*, caminho: str | Path | None = None) -> pd.DataFrame:
    """Reconstrói a Base Mestre com a ultima carga de cada condominio."""

    colunas_l = ", ".join(f"l.{coluna}" for coluna in COLUNAS_LANCAMENTO)
    with conectar(caminho) as conexao:
        linhas = conexao.execute(
            f"""
            SELECT cc.codigo_condominio, cc.condominio, {colunas_l}
            FROM cargas_condominio cc
            JOIN lancamentos_carga l ON l.carga_condominio_id = cc.id
            WHERE cc.id = (
                SELECT MAX(mais_recente.id)
                FROM cargas_condominio mais_recente
                WHERE mais_recente.codigo_condominio = cc.codigo_condominio
            )
            ORDER BY cc.codigo_condominio, l.ordem
            """
        ).fetchall()

    registros = []
    for linha in linhas:
        registro = dict(linha)
        for coluna in COLUNAS_VALORES:
            registro[coluna] = (
                float(registro[coluna]) if registro[coluna] is not None else None
            )
        registros.append(registro)
    return pd.DataFrame(registros, columns=COLUNAS_BASE)


def carregar_base_condominio(
    carga_condominio_id: int,
    *,
    caminho: str | Path | None = None,
) -> pd.DataFrame:
    """Reconstrói a Base Mestre de uma versão específica do condomínio."""

    colunas_l = ", ".join(f"l.{coluna}" for coluna in COLUNAS_LANCAMENTO)
    with conectar(caminho) as conexao:
        linhas = conexao.execute(
            f"""
            SELECT cc.codigo_condominio, cc.condominio, {colunas_l}
            FROM cargas_condominio cc
            JOIN lancamentos_carga l ON l.carga_condominio_id = cc.id
            WHERE cc.id = ?
            ORDER BY l.ordem
            """,
            (carga_condominio_id,),
        ).fetchall()

    registros = []
    for linha in linhas:
        registro = dict(linha)
        for coluna in COLUNAS_VALORES:
            registro[coluna] = (
                float(registro[coluna]) if registro[coluna] is not None else None
            )
        registros.append(registro)
    return pd.DataFrame(registros, columns=COLUNAS_BASE)


def listar_cargas_condominio(
    codigo_condominio: str,
    *,
    caminho: str | Path | None = None,
) -> list[dict]:
    """Lista as versoes de um condominio, inclusive as anteriores."""

    with conectar(caminho) as conexao:
        linhas = conexao.execute(
            """
            SELECT cc.id, cc.condominio, cc.quantidade_lancamentos,
                   cc.arquivos_json, c.criada_em
            FROM cargas_condominio cc
            JOIN cargas c ON c.id = cc.carga_id
            WHERE cc.codigo_condominio = ?
            ORDER BY cc.id DESC
            """,
            (codigo_condominio,),
        ).fetchall()
    return [
        {**dict(linha), "arquivos": json.loads(linha["arquivos_json"])}
        for linha in linhas
    ]


def listar_historico_debitos(
    codigo_condominio: str,
    economia: str,
    nome_condomino: str,
    *,
    caminho: str | Path | None = None,
) -> list[dict]:
    """Lista lancamentos do devedor em todas as cargas, com data da carga."""

    with conectar(caminho) as conexao:
        linhas = conexao.execute(
            """
            SELECT c.criada_em, cc.id AS carga_condominio_id,
                   l.competencia, l.data_vencimento, l.tipo,
                   l.vlr_base, l.vlr_corrigido, l.nosso_numero
            FROM lancamentos_carga l
            JOIN cargas_condominio cc ON cc.id = l.carga_condominio_id
            JOIN cargas c ON c.id = cc.carga_id
            WHERE cc.codigo_condominio = ?
              AND l.economia = ?
              AND l.nome_condomino = ? COLLATE NOCASE
            ORDER BY cc.id DESC, l.ordem
            """,
            (codigo_condominio, economia, nome_condomino),
        ).fetchall()
    return [dict(linha) for linha in linhas]
