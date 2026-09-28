"""Adapta o parser estável aos telefones prefixados ao nome no PDF.

As regras de leitura dos débitos e validação financeira continuam na versão
1.3. Este módulo apenas separa um prefixo que tenha formato de telefone;
números que façam parte do nome permanecem intocados.
"""

from __future__ import annotations

import re

from parser_pdf_base_mestre_v1_3_estavel import (
    gerar_resumo,
    processar_pdf as _processar_pdf_estavel,
    validar_estrutura,
)


_PREFIXO_TELEFONE = re.compile(
    r"^\s*(?P<telefone>[+\d().\s-]+\d)\s*-\s*"
    r"(?P<nome>[A-Za-zÀ-ÿ][^\n]*)$"
)
_MARCADOR_UNIDADE = re.compile(r"^(?P<marcador>\(\d{2}\)\*)\s+(?P<resto>.+)$")
_DDD_NA_ECONOMIA = re.compile(
    r"^(?P<economia>.+\(\d{2}\)\*)\s+\((?P<ddd>\d{2,3})\)$"
)
_REPRESENTANTE = re.compile(
    r"^(?P<pessoa>[^:]+):\s*(?P<papel>Procurador(?:a)?)\s+(?P<resto>.+)$",
    re.IGNORECASE,
)

# DDDs brasileiros: evita classificar um identificador numérico arbitrário
# de 10 ou 11 dígitos como telefone apenas porque termina em hífen.
_DDDS = {
    11, 12, 13, 14, 15, 16, 17, 18, 19,
    21, 22, 24, 27, 28,
    31, 32, 33, 34, 35, 37, 38,
    41, 42, 43, 44, 45, 46, 47, 48, 49,
    51, 53, 54, 55, 61, 62, 63, 64, 65, 66, 67, 68, 69,
    71, 73, 74, 75, 77, 79,
    81, 82, 83, 84, 85, 86, 87, 88, 89,
    91, 92, 93, 94, 95, 96, 97, 98, 99,
}


def analisar_nome_com_telefone(
    valor: object,
) -> tuple[str, str | None, bool, bool, str | None]:
    """Devolve nome, número, DDD pendente, formato pendente e marcador da unidade.

    É deliberadamente conservador: aceita apenas prefixo numérico terminado
    em hífen, seguido de nome iniciado por letra. Um telefone local de 8 ou
    9 dígitos é reconhecido, mas não recebe DDD presumido.
    """

    original = str(valor).strip()
    marcador = _MARCADOR_UNIDADE.fullmatch(original)
    texto = marcador.group("resto") if marcador else original
    encontrado = _PREFIXO_TELEFONE.fullmatch(texto)
    if not encontrado:
        return original, None, False, False, None

    telefone = re.sub(r"\D", "", encontrado.group("telefone"))
    nome = encontrado.group("nome").strip()
    if not nome:
        return original, None, False, False, None

    unidade = marcador.group("marcador") if marcador else None

    if len(telefone) in (8, 9) and telefone[0] in "23456789":
        return nome, telefone, True, False, unidade
    if len(telefone) in (10, 11) and int(telefone[:2]) in _DDDS:
        return nome, telefone, False, False, unidade
    if (
        len(telefone) in (12, 13)
        and telefone.startswith("55")
        and int(telefone[2:4]) in _DDDS
    ):
        return nome, telefone, False, False, unidade
    # Alguns PDFs trazem DDD mais dez dígitos locais. É plausível que haja um
    # dígito duplicado, mas não escolhemos qual remover automaticamente.
    if (
        len(telefone) == 12
        and int(telefone[:2]) in _DDDS
        and re.match(r"^\(?\d{2}\)?[\s-]+", encontrado.group("telefone"))
    ):
        return nome, telefone, False, True, unidade
    return original, None, False, False, None


def separar_telefone_do_nome(valor: object) -> tuple[str, str | None, bool]:
    """Interface simples para nome, número e ausência de DDD."""

    nome, numero, sem_ddd, _, _ = analisar_nome_com_telefone(valor)
    return nome, numero, sem_ddd


def analisar_registro(economia: object, valor_nome: object) -> dict:
    """Corrige a identidade da unidade e separa um telefone do campo do nome."""

    economia_original = str(economia).strip()
    nome_original = str(valor_nome).strip()
    economia_limpa = economia_original
    texto = nome_original
    ddd_recuperado = None

    ddd_na_economia = _DDD_NA_ECONOMIA.fullmatch(economia_original)
    if ddd_na_economia:
        economia_limpa = ddd_na_economia.group("economia")
        ddd_recuperado = ddd_na_economia.group("ddd")
        # Na leitura de um PDF novo, a parte local ainda estará no começo do
        # nome. Em um banco já processado, ela poderá estar só nos contatos.
        if _PREFIXO_TELEFONE.fullmatch(texto):
            texto = f"({ddd_recuperado}) {texto}"

    representante = _REPRESENTANTE.fullmatch(texto)
    observacao = None
    if representante:
        candidato = analisar_nome_com_telefone(representante.group("resto"))
        if candidato[1]:
            pessoa = representante.group("pessoa").strip()
            papel = representante.group("papel")
            # O texto extraído pode conter um nome após o número, mesmo que
            # essa parte não esteja visível na linha do PDF. A identidade
            # mostrada é a representante; não atribuímos a dívida ao sufixo.
            nome = f"{pessoa}: {papel}"
            numero, sem_ddd, formato = candidato[1:4]
            marcador = None
            observacao = f"{pessoa} ({papel.lower()}); contato informado no PDF."
        else:
            nome, numero, sem_ddd, formato, marcador = analisar_nome_com_telefone(texto)
    else:
        nome, numero, sem_ddd, formato, marcador = analisar_nome_com_telefone(texto)
    if marcador:
        economia_limpa = f"{economia_limpa} {marcador}"
    elif not numero:
        # O relatório também pode ter sido gravado antes da importação dos
        # telefones: nesse caso o marcador ficou sozinho no campo do nome.
        marcador_sozinho = _MARCADOR_UNIDADE.fullmatch(nome)
        if marcador_sozinho and re.match(r"^[A-Za-zÀ-ÿ]", marcador_sozinho.group("resto")):
            economia_limpa = f"{economia_limpa} {marcador_sozinho.group('marcador')}"
            nome = marcador_sozinho.group("resto").strip()

    return {
        "economia": economia_limpa,
        "nome_condomino": nome,
        "telefone": numero,
        "ddd_pendente": sem_ddd,
        "formato_pendente": formato,
        "observacoes": observacao,
        "ddd_recuperado": ddd_recuperado,
    }


def processar_pdf(caminho_pdf):
    """Processa PDF sem alterar os cálculos da versão estável."""

    resultado = _processar_pdf_estavel(caminho_pdf)
    dados = resultado["dados"].copy()
    registros = {
        (economia, nome): analisar_registro(economia, nome)
        for economia, nome in zip(
            dados["economia"], dados["nome_condomino"]
        )
    }
    dados["economia"] = [
        registros[(economia, nome)]["economia"]
        for economia, nome in zip(
            resultado["dados"]["economia"], resultado["dados"]["nome_condomino"]
        )
    ]
    dados["nome_condomino"] = [
        registros[(economia, nome)]["nome_condomino"]
        for economia, nome in zip(
            resultado["dados"]["economia"], resultado["dados"]["nome_condomino"]
        )
    ]

    codigo = resultado["cabecalho"]["codigo_condominio"]
    telefones = {}
    for _, linha in resultado["dados"].iterrows():
        registro = registros[(linha["economia"], linha["nome_condomino"])]
        numero = registro["telefone"]
        if numero:
            chave = (str(codigo), registro["economia"],
                     registro["nome_condomino"], numero)
            telefones[chave] = {
                "codigo_condominio": str(codigo),
                "economia": registro["economia"],
                "nome_condomino": registro["nome_condomino"],
                "telefone": numero,
                "ddd_pendente": registro["ddd_pendente"],
                "formato_pendente": registro["formato_pendente"],
                "observacoes": registro["observacoes"],
            }

    resultado["dados"] = dados
    resultado["telefones_pdf"] = list(telefones.values())
    resultado["parser_revisao"] = (
        "telefones-v3"
        if any(item["observacoes"] or item["ddd_recuperado"] for item in registros.values())
        else "telefones-v2"
        if any(
            item["formato_pendente"] or item["economia"] != economia
            for (economia, _), item in registros.items()
        )
        else "telefones-v1" if telefones else ""
    )
    return resultado
