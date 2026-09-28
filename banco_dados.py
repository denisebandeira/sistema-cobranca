"""Persistencia local do Sistema de Cobranca.

O banco fica em um diretorio de dados do usuario, fora do executavel e da
pasta temporaria usada pelo PyInstaller. Assim, os dados sobrevivem a novas
versoes do programa.
"""

from __future__ import annotations

from contextlib import closing, contextmanager
from datetime import date, datetime
import hashlib
import json
import os
from pathlib import Path
import re
import sqlite3
import sys
import tempfile
from typing import Iterable, Iterator, Mapping
import unicodedata


NOME_APLICATIVO = "SistemaCobranca"
NOME_BANCO = "sistema_cobranca.db"
VARIAVEL_DIRETORIO_DADOS = "SISTEMA_COBRANCA_DATA_DIR"
VERSAO_ESQUEMA = 5
MARCADOR_PORTATIL = "MODO_PORTATIL.txt"


def diretorio_programa() -> Path:
    """Pasta externa ao executavel, nunca a pasta temporaria do PyInstaller."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def modo_portatil_ativo() -> bool:
    return (
        not os.environ.get(VARIAVEL_DIRETORIO_DADOS)
        and (diretorio_programa() / MARCADOR_PORTATIL).is_file()
    )


def caminho_diretorio_dados_local() -> Path:
    """Local anterior, usado tambem para oferecer migracao sem perda de dados."""

    if sys.platform == "win32":
        base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        if base:
            return Path(base) / NOME_APLICATIVO
        return Path.home() / "AppData" / "Local" / NOME_APLICATIVO

    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / NOME_APLICATIVO

    base = os.environ.get("XDG_DATA_HOME")
    if base:
        return Path(base) / NOME_APLICATIVO
    return Path.home() / ".local" / "share" / NOME_APLICATIVO


def caminho_diretorio_dados() -> Path:
    """Retorna o diretorio persistente de dados da aplicacao.

    ``SISTEMA_COBRANCA_DATA_DIR`` pode ser usado por suporte e testes para
    escolher outro local sem alterar o codigo.
    """

    diretorio_configurado = os.environ.get(VARIAVEL_DIRETORIO_DADOS)
    if diretorio_configurado:
        return Path(diretorio_configurado).expanduser().resolve()
    if modo_portatil_ativo():
        return diretorio_programa() / "dados"
    return caminho_diretorio_dados_local()


def caminho_banco() -> Path:
    """Retorna o caminho padrao do arquivo SQLite."""

    return caminho_diretorio_dados() / NOME_BANCO


def caminho_banco_local() -> Path:
    return caminho_diretorio_dados_local() / NOME_BANCO


def migrar_banco_local_para_portatil() -> Path:
    """Copia consistente do banco antigo, inclusive se ele usava WAL."""

    if not modo_portatil_ativo() or os.environ.get(VARIAVEL_DIRETORIO_DADOS):
        raise RuntimeError("O modo portatil nao esta ativo.")
    origem = caminho_banco_local()
    destino = caminho_banco()
    if not origem.is_file():
        raise FileNotFoundError(f"Banco local nao encontrado: {origem}")
    if destino.exists():
        raise FileExistsError(f"O banco portatil ja existe: {destino}")

    destino.parent.mkdir(parents=True, exist_ok=True)
    descritor, temporario = tempfile.mkstemp(
        prefix=".copia_banco_", suffix=".db", dir=destino.parent
    )
    os.close(descritor)
    try:
        with closing(sqlite3.connect(origem)) as fonte:
            with closing(sqlite3.connect(temporario)) as copia:
                fonte.backup(copia)
                if copia.execute("PRAGMA quick_check").fetchone()[0] != "ok":
                    raise RuntimeError("A copia do banco nao passou na verificacao.")
        if destino.exists():
            raise FileExistsError(f"O banco portatil ja existe: {destino}")
        os.replace(temporario, destino)
    finally:
        Path(temporario).unlink(missing_ok=True)
    return destino


def _resolver_caminho(caminho: str | Path | None) -> Path:
    return Path(caminho).expanduser().resolve() if caminho else caminho_banco()


def _texto_obrigatorio(valor: object, campo: str) -> str:
    texto = str(valor).strip() if valor is not None else ""
    if not texto:
        raise ValueError(f"{campo} e obrigatorio.")
    return texto


def _texto_opcional(valor: object | None) -> str | None:
    if valor is None:
        return None
    texto = str(valor).strip()
    return texto or None


def _data_iso(valor: date | datetime | str | None) -> str | None:
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.isoformat(timespec="seconds")
    if isinstance(valor, date):
        return valor.isoformat()
    return _texto_opcional(valor)


def normalizar_telefone(valor: object | None, *, permitir_vazio: bool = False) -> str | None:
    """Remove formatacao e valida um telefone brasileiro com ou sem DDI 55."""

    texto = _texto_opcional(valor)
    if texto is None:
        if permitir_vazio:
            return None
        raise ValueError("telefone e obrigatorio.")

    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]

    telefone = re.sub(r"\D", "", texto)
    tamanho_valido = len(telefone) in (10, 11)
    tamanho_com_ddd = telefone.startswith("55") and len(telefone) in (12, 13)
    if not (tamanho_valido or tamanho_com_ddd):
        raise ValueError(
            "telefone deve ter DDD e 10 ou 11 digitos, com DDI 55 opcional."
        )
    return telefone


def normalizar_email(valor: object) -> str:
    email = _texto_obrigatorio(valor, "e-mail")
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
        raise ValueError("Informe um endereço de e-mail válido.")
    return email


def _chave_identificador(valor: object) -> str:
    texto = _texto_obrigatorio(valor, "identificador")
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto.casefold()


def _chave_nome(valor: object) -> str:
    texto = _texto_obrigatorio(valor, "nome_condomino")
    sem_acentos = "".join(
        caractere
        for caractere in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(caractere)
    )
    return " ".join(sem_acentos.casefold().split())


@contextmanager
def conectar(caminho: str | Path | None = None) -> Iterator[sqlite3.Connection]:
    """Abre uma conexao transacional com chaves estrangeiras habilitadas."""

    arquivo = _resolver_caminho(caminho)
    arquivo.parent.mkdir(parents=True, exist_ok=True)

    conexao = sqlite3.connect(arquivo, timeout=30)
    conexao.row_factory = sqlite3.Row
    try:
        conexao.execute("PRAGMA foreign_keys = ON")
        conexao.execute("PRAGMA busy_timeout = 30000")
        yield conexao
        conexao.commit()
    except Exception:
        conexao.rollback()
        raise
    finally:
        conexao.close()


def _migrar_versao_1_para_2(conexao: sqlite3.Connection) -> None:
    """Inclui o condomino na chave unica sem perder os dados existentes."""

    conexao.executescript(
        """
        PRAGMA foreign_keys = OFF;
        BEGIN;

        CREATE TABLE contatos_novos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_condominio TEXT NOT NULL,
            economia TEXT NOT NULL,
            condominio TEXT,
            condomino TEXT NOT NULL COLLATE NOCASE,
            telefone TEXT,
            observacoes TEXT,
            ativo INTEGER NOT NULL DEFAULT 1 CHECK (ativo IN (0, 1)),
            criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (codigo_condominio, economia, condomino)
        );

        INSERT INTO contatos_novos (
            id, codigo_condominio, economia, condominio, condomino,
            telefone, observacoes, ativo, criado_em, atualizado_em
        )
        SELECT
            id, codigo_condominio, economia, condominio, condomino,
            telefone, observacoes, ativo, criado_em, atualizado_em
        FROM contatos;

        DROP TABLE contatos;
        ALTER TABLE contatos_novos RENAME TO contatos;

        CREATE INDEX idx_contatos_condomino
            ON contatos (condomino COLLATE NOCASE);
        CREATE INDEX idx_contatos_telefone
            ON contatos (telefone);

        PRAGMA user_version = 2;
        COMMIT;
        PRAGMA foreign_keys = ON;
        """
    )


def inicializar_banco(caminho: str | Path | None = None) -> Path:
    """Cria o arquivo, as tabelas e os indices caso ainda nao existam."""

    arquivo = _resolver_caminho(caminho)

    with conectar(arquivo) as conexao:
        versao_atual = conexao.execute("PRAGMA user_version").fetchone()[0]
        if versao_atual > VERSAO_ESQUEMA:
            raise RuntimeError(
                "O banco de dados foi criado por uma versao mais nova "
                "do Sistema de Cobranca."
            )

        if 0 < versao_atual < VERSAO_ESQUEMA:
            copia = arquivo.with_name(f"{arquivo.stem}.antes_v{VERSAO_ESQUEMA}.db")
            if not copia.exists():
                with closing(sqlite3.connect(copia)) as destino:
                    conexao.backup(destino)

        # No pendrive, o modo DELETE evita arquivos -wal/-shm deixados ao
        # lado do banco quando a unidade e removida apos fechar o programa.
        portatil = (
            modo_portatil_ativo()
            and not os.environ.get(VARIAVEL_DIRETORIO_DADOS)
            and arquivo == caminho_banco()
        )
        conexao.execute(
            "PRAGMA journal_mode = DELETE" if portatil
            else "PRAGMA journal_mode = WAL"
        )

        if versao_atual == 1:
            _migrar_versao_1_para_2(conexao)

        conexao.executescript(
            """
            CREATE TABLE IF NOT EXISTS contatos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo_condominio TEXT NOT NULL,
                economia TEXT NOT NULL,
                condominio TEXT,
                condomino TEXT NOT NULL COLLATE NOCASE,
                telefone TEXT,
                observacoes TEXT,
                ativo INTEGER NOT NULL DEFAULT 1 CHECK (ativo IN (0, 1)),
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (codigo_condominio, economia, condomino)
            );

            CREATE INDEX IF NOT EXISTS idx_contatos_condomino
                ON contatos (condomino COLLATE NOCASE);

            CREATE INDEX IF NOT EXISTS idx_contatos_telefone
                ON contatos (telefone);

            CREATE TABLE IF NOT EXISTS historico_comunicacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contato_id INTEGER,
                codigo_condominio TEXT NOT NULL,
                economia TEXT NOT NULL,
                condomino TEXT NOT NULL,
                competencia TEXT,
                vencimento TEXT,
                valor NUMERIC,
                canal TEXT NOT NULL,
                valor_canal TEXT,
                acao TEXT NOT NULL,
                resultado TEXT,
                observacoes TEXT,
                realizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (contato_id) REFERENCES contatos (id)
                    ON UPDATE CASCADE ON DELETE SET NULL
            );

            CREATE INDEX IF NOT EXISTS idx_historico_contato_data
                ON historico_comunicacoes (contato_id, realizado_em DESC);

            CREATE INDEX IF NOT EXISTS idx_historico_debito
                ON historico_comunicacoes
                    (codigo_condominio, economia, vencimento);

            CREATE INDEX IF NOT EXISTS idx_historico_realizado_em
                ON historico_comunicacoes (realizado_em DESC);

            CREATE TABLE IF NOT EXISTS meios_contato (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contato_id INTEGER NOT NULL,
                tipo TEXT NOT NULL CHECK (tipo IN ('Telefone', 'E-mail')),
                valor TEXT NOT NULL COLLATE NOCASE,
                observacoes TEXT,
                ativo INTEGER NOT NULL DEFAULT 1 CHECK (ativo IN (0, 1)),
                criado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (contato_id) REFERENCES contatos (id),
                UNIQUE (contato_id, tipo, valor)
            );

            CREATE INDEX IF NOT EXISTS idx_meios_contato_pessoa
                ON meios_contato (contato_id, ativo, tipo, id);

            CREATE TABLE IF NOT EXISTS cargas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                criada_em TEXT NOT NULL DEFAULT
                    (strftime('%Y-%m-%dT%H:%M:%f', 'now'))
            );

            CREATE TABLE IF NOT EXISTS cargas_condominio (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                carga_id INTEGER NOT NULL,
                codigo_condominio TEXT NOT NULL,
                condominio TEXT NOT NULL,
                arquivos_json TEXT NOT NULL,
                conteudo_sha256 TEXT NOT NULL,
                quantidade_lancamentos INTEGER NOT NULL,
                FOREIGN KEY (carga_id) REFERENCES cargas (id),
                UNIQUE (carga_id, codigo_condominio)
            );

            CREATE INDEX IF NOT EXISTS idx_cargas_condominio_codigo_id
                ON cargas_condominio (codigo_condominio, id DESC);

            CREATE TABLE IF NOT EXISTS lancamentos_carga (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                carga_condominio_id INTEGER NOT NULL,
                ordem INTEGER NOT NULL,
                economia TEXT NOT NULL,
                nome_condomino TEXT NOT NULL,
                competencia TEXT,
                data_vencimento TEXT,
                tipo TEXT,
                vlr_base TEXT,
                multa TEXT,
                juros TEXT,
                correcao TEXT,
                vlr_corrigido TEXT,
                assessor TEXT,
                imobiliaria TEXT,
                endereco TEXT,
                nosso_numero TEXT,
                advogado TEXT,
                FOREIGN KEY (carga_condominio_id)
                    REFERENCES cargas_condominio (id),
                UNIQUE (carga_condominio_id, ordem)
            );

            CREATE INDEX IF NOT EXISTS idx_lancamentos_carga_devedor
                ON lancamentos_carga
                    (carga_condominio_id, economia, nome_condomino);

            CREATE INDEX IF NOT EXISTS idx_lancamentos_carga_vencimento
                ON lancamentos_carga (data_vencimento);

            CREATE TABLE IF NOT EXISTS situacao_unidades (
                codigo_condominio TEXT NOT NULL,
                economia TEXT NOT NULL,
                ajuizada INTEGER NOT NULL DEFAULT 0 CHECK (ajuizada IN (0, 1)),
                atualizado_em TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (codigo_condominio, economia)
            );

            CREATE INDEX IF NOT EXISTS idx_situacao_unidades_ajuizada
                ON situacao_unidades (ajuizada, codigo_condominio);
            """
        )
        if versao_atual < 4:
            colunas_historico = {
                linha["name"] for linha in conexao.execute(
                    "PRAGMA table_info(historico_comunicacoes)"
                )
            }
            if "valor_canal" not in colunas_historico:
                conexao.execute(
                    "ALTER TABLE historico_comunicacoes ADD COLUMN valor_canal TEXT"
                )
            conexao.execute(
                """
                INSERT OR IGNORE INTO meios_contato
                    (contato_id, tipo, valor, observacoes)
                SELECT id, 'Telefone', telefone, observacoes FROM contatos
                WHERE telefone IS NOT NULL AND TRIM(telefone) <> ''
                """
            )
        if versao_atual < VERSAO_ESQUEMA:
            conexao.execute(f"PRAGMA user_version = {VERSAO_ESQUEMA}")

    return arquivo


def unidade_ajuizada(
    codigo_condominio: object, economia: object, *,
    caminho: str | Path | None = None,
) -> bool:
    """Estado persistente da unidade, independente de qualquer carga."""

    codigo = _texto_obrigatorio(codigo_condominio, "codigo_condominio")
    unidade = _texto_obrigatorio(economia, "economia")
    with conectar(caminho) as conexao:
        registro = conexao.execute(
            "SELECT ajuizada FROM situacao_unidades "
            "WHERE codigo_condominio = ? AND economia = ?",
            (codigo, unidade),
        ).fetchone()
    return bool(registro["ajuizada"]) if registro else False


def listar_unidades_ajuizadas(*, caminho: str | Path | None = None) -> set[tuple[str, str]]:
    """Chaves das unidades que não podem receber comunicado."""

    with conectar(caminho) as conexao:
        registros = conexao.execute(
            "SELECT codigo_condominio, economia FROM situacao_unidades WHERE ajuizada = 1"
        ).fetchall()
    return {(linha["codigo_condominio"], linha["economia"]) for linha in registros}


def definir_unidade_ajuizada(
    codigo_condominio: object, economia: object, ajuizada: bool, *,
    caminho: str | Path | None = None,
) -> None:
    """Altera o estado da unidade imediatamente, sem modificar lançamentos."""

    codigo = _texto_obrigatorio(codigo_condominio, "codigo_condominio")
    unidade = _texto_obrigatorio(economia, "economia")
    if not isinstance(ajuizada, bool):
        raise ValueError("A situação ajuizada deve ser SIM ou NÃO.")
    with conectar(caminho) as conexao:
        conexao.execute(
            """INSERT INTO situacao_unidades (codigo_condominio, economia, ajuizada)
               VALUES (?, ?, ?)
               ON CONFLICT (codigo_condominio, economia) DO UPDATE SET
                   ajuizada = excluded.ajuizada,
                   atualizado_em = CURRENT_TIMESTAMP""",
            (codigo, unidade, int(ajuizada)),
        )


def definir_unidades_ajuizadas(
    codigo_condominio: object, economias: Iterable[object], ajuizada: bool, *,
    caminho: str | Path | None = None,
) -> int:
    """Atualiza várias unidades do mesmo condomínio numa única transação."""

    codigo = _texto_obrigatorio(codigo_condominio, "codigo_condominio")
    if not isinstance(ajuizada, bool):
        raise ValueError("A situação ajuizada deve ser SIM ou NÃO.")
    unidades = sorted({_texto_obrigatorio(item, "economia") for item in economias})
    if not unidades:
        return 0
    with conectar(caminho) as conexao:
        conexao.executemany(
            """INSERT INTO situacao_unidades (codigo_condominio, economia, ajuizada)
               VALUES (?, ?, ?)
               ON CONFLICT (codigo_condominio, economia) DO UPDATE SET
                   ajuizada = excluded.ajuizada,
                   atualizado_em = CURRENT_TIMESTAMP""",
            [(codigo, unidade, int(ajuizada)) for unidade in unidades],
        )
    return len(unidades)


def salvar_contato(
    codigo_condominio: object,
    economia: object,
    condomino: object,
    telefone: object | None = None,
    *,
    condominio: object | None = None,
    observacoes: object | None = None,
    ativo: bool = True,
    caminho: str | Path | None = None,
) -> int:
    """Inclui ou atualiza um contato e retorna seu identificador."""

    codigo = _texto_obrigatorio(codigo_condominio, "codigo_condominio")
    unidade = _texto_obrigatorio(economia, "economia")
    nome = _texto_obrigatorio(condomino, "condomino")

    with conectar(caminho) as conexao:
        conexao.execute(
            """
            INSERT INTO contatos (
                codigo_condominio, economia, condominio, condomino,
                telefone, observacoes, ativo
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT (codigo_condominio, economia, condomino) DO UPDATE SET
                condominio = excluded.condominio,
                telefone = excluded.telefone,
                observacoes = excluded.observacoes,
                ativo = excluded.ativo,
                atualizado_em = CURRENT_TIMESTAMP
            """,
            (
                codigo,
                unidade,
                _texto_opcional(condominio),
                nome,
                _texto_opcional(telefone),
                _texto_opcional(observacoes),
                int(ativo),
            ),
        )
        registro = conexao.execute(
            """
            SELECT id FROM contatos
            WHERE codigo_condominio = ? AND economia = ?
              AND condomino = ? COLLATE NOCASE
            """,
            (codigo, unidade, nome),
        ).fetchone()
        if _texto_opcional(telefone):
            conexao.execute(
                """
                INSERT INTO meios_contato
                    (contato_id, tipo, valor, observacoes)
                VALUES (?, 'Telefone', ?, ?)
                ON CONFLICT (contato_id, tipo, valor) DO UPDATE SET
                    observacoes = COALESCE(
                        excluded.observacoes, meios_contato.observacoes
                    ),
                    ativo = 1,
                    atualizado_em = CURRENT_TIMESTAMP
                """,
                (registro["id"], _texto_opcional(telefone),
                 _texto_opcional(observacoes)),
            )

    return int(registro["id"])


def obter_contato(
    codigo_condominio: object,
    economia: object,
    condomino: object | None = None,
    *,
    caminho: str | Path | None = None,
) -> dict | None:
    """Busca um contato pela chave de negocio condominio + economia."""

    consulta = """
        SELECT * FROM contatos
        WHERE codigo_condominio = ? AND economia = ?
    """
    parametros = [
        _texto_obrigatorio(codigo_condominio, "codigo_condominio"),
        _texto_obrigatorio(economia, "economia"),
    ]
    if condomino is not None:
        consulta += " AND condomino = ? COLLATE NOCASE"
        parametros.append(_texto_obrigatorio(condomino, "condomino"))
    consulta += " ORDER BY ativo DESC, atualizado_em DESC, id DESC LIMIT 1"

    with conectar(caminho) as conexao:
        registro = conexao.execute(
            consulta,
            parametros,
        ).fetchone()
    return dict(registro) if registro else None


def sincronizar_devedores(
    devedores: Iterable[Mapping[str, object]],
    *,
    caminho: str | Path | None = None,
) -> list[dict]:
    """Cadastra devedores novos e associa os contatos ja existentes.

    Cada item deve conter ``codigo_condominio``, ``economia`` e
    ``nome_condomino``. O telefone existente nunca e sobrescrito por esta
    sincronizacao.
    """

    normalizados = []
    for devedor in devedores:
        normalizados.append(
            {
                **dict(devedor),
                "codigo_condominio": _texto_obrigatorio(
                    devedor.get("codigo_condominio"),
                    "codigo_condominio",
                ),
                "economia": _texto_obrigatorio(
                    devedor.get("economia"),
                    "economia",
                ),
                "nome_condomino": _texto_obrigatorio(
                    devedor.get("nome_condomino"),
                    "nome_condomino",
                ),
            }
        )

    with conectar(caminho) as conexao:
        for devedor in normalizados:
            conexao.execute(
                """
                INSERT INTO contatos (
                    codigo_condominio, economia, condominio, condomino
                ) VALUES (?, ?, ?, ?)
                ON CONFLICT (
                    codigo_condominio, economia, condomino
                ) DO UPDATE SET
                    condominio = excluded.condominio,
                    ativo = 1,
                    atualizado_em = CURRENT_TIMESTAMP
                WHERE COALESCE(contatos.condominio, '') <>
                      COALESCE(excluded.condominio, '')
                   OR contatos.ativo <> 1
                """,
                (
                    devedor["codigo_condominio"],
                    devedor["economia"],
                    _texto_opcional(devedor.get("condominio")),
                    devedor["nome_condomino"],
                ),
            )

        contatos = conexao.execute(
            """
            SELECT id, codigo_condominio, economia, condomino, telefone
            FROM contatos
            WHERE ativo = 1
            """
        ).fetchall()

    por_chave = {
        (
            contato["codigo_condominio"],
            contato["economia"],
            contato["condomino"].casefold(),
        ): contato
        for contato in contatos
    }

    resultado = []
    for devedor in normalizados:
        chave = (
            devedor["codigo_condominio"],
            devedor["economia"],
            devedor["nome_condomino"].casefold(),
        )
        contato = por_chave[chave]
        telefone = contato["telefone"]
        resultado.append(
            {
                **devedor,
                "contato_id": contato["id"],
                "telefone": telefone,
                "situacao_contato": (
                    "Com telefone" if telefone else "Sem telefone"
                ),
            }
        )

    return resultado


def listar_contatos(
    *,
    somente_ativos: bool = True,
    caminho: str | Path | None = None,
) -> list[dict]:
    """Lista contatos em ordem de condominio, economia e condomino."""

    consulta = "SELECT * FROM contatos"
    parametros: tuple[object, ...] = ()
    if somente_ativos:
        consulta += " WHERE ativo = ?"
        parametros = (1,)
    consulta += " ORDER BY codigo_condominio, economia, condomino COLLATE NOCASE"

    with conectar(caminho) as conexao:
        registros = conexao.execute(consulta, parametros).fetchall()
    return [dict(registro) for registro in registros]


def atualizar_contato(
    contato_id: int,
    telefone: object | None,
    *,
    observacoes: object | None = None,
    caminho: str | Path | None = None,
) -> None:
    """Atualiza telefone e, quando informado, observacoes de um contato."""

    telefone_normalizado = normalizar_telefone(
        telefone,
        permitir_vazio=True,
    )
    observacoes_normalizadas = _texto_opcional(observacoes)

    with conectar(caminho) as conexao:
        cursor = conexao.execute(
            """
            UPDATE contatos
            SET telefone = ?,
                observacoes = COALESCE(?, observacoes),
                atualizado_em = CURRENT_TIMESTAMP
            WHERE id = ?
            """,
            (telefone_normalizado, observacoes_normalizadas, contato_id),
        )
        if cursor.rowcount != 1:
            raise ValueError("Contato nao encontrado.")
        if telefone_normalizado:
            conexao.execute(
                """
                INSERT INTO meios_contato
                    (contato_id, tipo, valor, observacoes)
                VALUES (?, 'Telefone', ?, ?)
                ON CONFLICT (contato_id, tipo, valor) DO UPDATE SET
                    observacoes = COALESCE(
                        excluded.observacoes, meios_contato.observacoes
                    ),
                    ativo = 1,
                    atualizado_em = CURRENT_TIMESTAMP
                """,
                (contato_id, telefone_normalizado, observacoes_normalizadas),
            )


def _sincronizar_telefone_principal(
    conexao: sqlite3.Connection, contato_id: int
) -> None:
    candidatos = conexao.execute(
        """
        SELECT valor FROM meios_contato
        WHERE contato_id = ? AND tipo = 'Telefone' AND ativo = 1
        ORDER BY id DESC
        """,
        (contato_id,),
    ).fetchall()
    principal = None
    for candidato in candidatos:
        try:
            principal = normalizar_telefone(candidato["valor"])
            break
        except ValueError:
            # Números locais vindos do PDF não têm DDD confirmado.
            continue
    conexao.execute(
        """
        UPDATE contatos SET telefone = ?, atualizado_em = CURRENT_TIMESTAMP
        WHERE id = ?
        """,
        (principal, contato_id),
    )


def listar_meios_contato(
    contato_id: int, *, caminho: str | Path | None = None
) -> list[dict]:
    """Lista telefones e e-mails ativos de um devedor."""

    with conectar(caminho) as conexao:
        registros = conexao.execute(
            """
            SELECT * FROM meios_contato
            WHERE contato_id = ? AND ativo = 1
            ORDER BY CASE tipo WHEN 'Telefone' THEN 0 ELSE 1 END, id DESC
            """,
            (contato_id,),
        ).fetchall()
    return [dict(registro) for registro in registros]


def salvar_meio_contato(
    contato_id: int,
    tipo: str,
    valor: object,
    *,
    observacoes: object | None = None,
    meio_id: int | None = None,
    caminho: str | Path | None = None,
) -> int:
    """Inclui ou edita um telefone/e-mail, preservando os demais."""

    if tipo not in {"Telefone", "E-mail"}:
        raise ValueError("Selecione Telefone ou E-mail.")
    valor_normalizado = (
        normalizar_telefone(valor) if tipo == "Telefone" else normalizar_email(valor)
    )
    nota = _texto_opcional(observacoes)
    with conectar(caminho) as conexao:
        existente = conexao.execute(
            "SELECT id FROM contatos WHERE id = ?", (contato_id,)
        ).fetchone()
        if not existente:
            raise ValueError("Devedor não encontrado.")
        if meio_id is None:
            try:
                cursor = conexao.execute(
                    """
                    INSERT INTO meios_contato (contato_id, tipo, valor, observacoes)
                    VALUES (?, ?, ?, ?)
                    """,
                    (contato_id, tipo, valor_normalizado, nota),
                )
            except sqlite3.IntegrityError as erro:
                raise ValueError("Este contato já está cadastrado para o devedor.") from erro
            identificador = int(cursor.lastrowid)
        else:
            try:
                cursor = conexao.execute(
                    """
                    UPDATE meios_contato
                    SET tipo = ?, valor = ?, observacoes = ?,
                        atualizado_em = CURRENT_TIMESTAMP
                    WHERE id = ? AND contato_id = ? AND ativo = 1
                    """,
                    (tipo, valor_normalizado, nota, meio_id, contato_id),
                )
            except sqlite3.IntegrityError as erro:
                raise ValueError("Este contato já está cadastrado para o devedor.") from erro
            if cursor.rowcount != 1:
                raise ValueError("Contato não encontrado.")
            identificador = meio_id
        _sincronizar_telefone_principal(conexao, contato_id)
    return identificador


def desativar_meio_contato(
    contato_id: int, meio_id: int, *, caminho: str | Path | None = None
) -> None:
    """Retira um meio da lista ativa sem apagar o histórico de comunicações."""

    with conectar(caminho) as conexao:
        cursor = conexao.execute(
            """
            UPDATE meios_contato SET ativo = 0,
                atualizado_em = CURRENT_TIMESTAMP
            WHERE id = ? AND contato_id = ? AND ativo = 1
            """,
            (meio_id, contato_id),
        )
        if cursor.rowcount != 1:
            raise ValueError("Contato não encontrado.")
        _sincronizar_telefone_principal(conexao, contato_id)


def importar_telefones(
    registros: Iterable[Mapping[str, object]],
    *,
    caminho: str | Path | None = None,
) -> dict:
    """Associa telefones importados apenas a contatos ja cadastrados."""

    with conectar(caminho) as conexao:
        contatos = [
            dict(registro)
            for registro in conexao.execute(
                "SELECT * FROM contatos WHERE ativo = 1"
            ).fetchall()
        ]

        por_unidade: dict[tuple[str, str], list[dict]] = {}
        for contato in contatos:
            chave = (
                _chave_identificador(contato["codigo_condominio"]),
                _chave_identificador(contato["economia"]),
            )
            por_unidade.setdefault(chave, []).append(contato)

        atualizados = 0
        ignorados = 0
        erros = []
        for numero_linha, registro in enumerate(registros, start=2):
            linha = int(registro.get("_linha", numero_linha))
            try:
                chave_unidade = (
                    _chave_identificador(registro.get("codigo_condominio")),
                    _chave_identificador(registro.get("economia")),
                )
                candidatos = por_unidade.get(chave_unidade, [])

                nome = _texto_opcional(registro.get("nome_condomino"))
                if nome:
                    chave_nome = _chave_nome(nome)
                    candidatos = [
                        contato
                        for contato in candidatos
                        if _chave_nome(contato["condomino"]) == chave_nome
                    ]

                if not candidatos:
                    raise ValueError(
                        "nenhum contato cadastrado corresponde a esta linha."
                    )
                if len(candidatos) > 1:
                    raise ValueError(
                        "mais de um condomino corresponde a esta unidade; "
                        "informe nome_condomino."
                    )

                contato = candidatos[0]
                telefone = normalizar_telefone(registro.get("telefone"))
                observacoes = _texto_opcional(registro.get("observacoes"))
                duplicado = conexao.execute(
                    """
                    SELECT id FROM meios_contato
                    WHERE contato_id = ? AND tipo = 'Telefone'
                      AND valor = ? COLLATE NOCASE AND ativo = 1
                    """,
                    (contato["id"], telefone),
                ).fetchone()
                if duplicado:
                    ignorados += 1
                    continue
                conexao.execute(
                    """
                    UPDATE contatos
                    SET telefone = ?,
                        observacoes = COALESCE(?, observacoes),
                        atualizado_em = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (telefone, observacoes, contato["id"]),
                )
                conexao.execute(
                    """
                    INSERT INTO meios_contato
                        (contato_id, tipo, valor, observacoes)
                    VALUES (?, 'Telefone', ?, ?)
                    """,
                    (contato["id"], telefone, observacoes),
                )
                atualizados += 1

            except (TypeError, ValueError) as erro:
                erros.append({"linha": linha, "erro": str(erro)})

    return {
        "atualizados": atualizados,
        "ignorados": ignorados,
        "erros": erros,
    }


def registrar_telefones_pdf(
    registros: Iterable[Mapping[str, object]],
    *,
    caminho: str | Path | None = None,
) -> dict:
    """Acrescenta telefones lidos do PDF sem substituir dados manuais.

    Números locais são guardados para conferência, mas não se tornam o
    telefone principal: sem DDD não estão prontos para uso.
    """

    novos = repetidos = sem_ddd = formatos_pendentes = 0
    erros = []
    with conectar(caminho) as conexao:
        for registro in registros:
            codigo = _texto_opcional(registro.get("codigo_condominio"))
            economia = _texto_opcional(registro.get("economia"))
            nome = _texto_opcional(registro.get("nome_condomino"))
            numero = re.sub(r"\D", "", str(registro.get("telefone") or ""))
            pendente = bool(registro.get("ddd_pendente"))
            formato_pendente = bool(registro.get("formato_pendente"))
            if not codigo or not economia or not nome:
                erros.append("Telefone sem identificação completa do devedor.")
                continue
            if pendente:
                if len(numero) not in (8, 9) or numero[0] not in "23456789":
                    erros.append(f"Telefone local inválido: condomínio {codigo}, economia {economia}.")
                    continue
                sem_ddd += 1
            elif formato_pendente:
                if len(numero) != 12:
                    erros.append(f"Telefone para conferência inválido: condomínio {codigo}, economia {economia}.")
                    continue
                formatos_pendentes += 1
            else:
                try:
                    numero = normalizar_telefone(numero)
                except ValueError:
                    erros.append(f"Telefone inválido: condomínio {codigo}, economia {economia}.")
                    continue

            contato = conexao.execute(
                """SELECT id, telefone FROM contatos
                   WHERE codigo_condominio = ? AND economia = ?
                     AND condomino = ? COLLATE NOCASE AND ativo = 1""",
                (codigo, economia, nome),
            ).fetchone()
            if contato is None:
                erros.append(f"Devedor não encontrado: condomínio {codigo}, economia {economia}.")
                continue

            existente = conexao.execute(
                """SELECT id FROM meios_contato
                   WHERE contato_id = ? AND tipo = 'Telefone' AND valor = ?""",
                (contato["id"], numero),
            ).fetchone()
            if existente:
                repetidos += 1
                continue

            observacao = (
                "DDD não informado. Conferir antes de usar."
                if pendente else
                "Formato fora do padrão. Conferir antes de usar."
                if formato_pendente else "Importado do PDF."
            )
            if registro.get("observacoes"):
                observacao = f"{observacao} {str(registro['observacoes']).strip()}"
            conexao.execute(
                """INSERT INTO meios_contato
                       (contato_id, tipo, valor, observacoes)
                   VALUES (?, 'Telefone', ?, ?)""",
                (contato["id"], numero, observacao),
            )
            if not pendente and not formato_pendente and not contato["telefone"]:
                conexao.execute(
                    """UPDATE contatos SET telefone = ?,
                           atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
                    (numero, contato["id"]),
                )
            novos += 1

    return {"novos": novos, "repetidos": repetidos,
            "sem_ddd": sem_ddd, "formato_pendente": formatos_pendentes,
            "erros": erros}


def recuperar_telefones_de_cargas_antigas(
    *, caminho: str | Path | None = None
) -> dict:
    """Corrige cargas lidas antes do parser com telefones, sem exigir os PDFs.

    A alteração é atômica e precedida de uma cópia do banco. Corrige somente
    nomes e economias reconhecidos pelas regras de leitura; os valores dos
    débitos não mudam.
    """

    from parser_pdf_base_mestre_v1_4_telefones import analisar_registro

    arquivo = _resolver_caminho(caminho)
    with conectar(arquivo) as conexao:
        linhas = conexao.execute(
            """SELECT lc.id, lc.economia, lc.nome_condomino,
                      cc.id AS carga_condominio_id, cc.codigo_condominio,
                      cc.condominio, cc.arquivos_json
               FROM lancamentos_carga lc
               JOIN cargas_condominio cc ON cc.id = lc.carga_condominio_id
               ORDER BY lc.id"""
        ).fetchall()
        correcoes = []
        for linha in linhas:
            registro = analisar_registro(linha["economia"], linha["nome_condomino"])
            if (
                registro["telefone"]
                or registro["economia"] != linha["economia"]
                or registro["nome_condomino"] != linha["nome_condomino"]
            ):
                correcoes.append((linha, registro))
        if not correcoes:
            return {"nomes_corrigidos": 0, "telefones_novos": 0,
                    "sem_ddd": 0, "formato_pendente": 0, "copia": None}

        copia = arquivo.with_name(f"{arquivo.stem}.antes_parser_v5.db")
        if not copia.exists():
            with closing(sqlite3.connect(copia)) as destino:
                conexao.backup(destino)

        grupos_afetados = {}
        devedores = {}
        for linha, registro in correcoes:
            economia_nova = registro["economia"]
            nome = registro["nome_condomino"]
            conexao.execute(
                """UPDATE lancamentos_carga
                   SET economia = ?, nome_condomino = ? WHERE id = ?""",
                (economia_nova, nome, linha["id"]),
            )
            grupo_id = linha["carga_condominio_id"]
            revisao = (
                "telefones-v3"
                if registro["observacoes"] or registro["ddd_recuperado"]
                else "telefones-v2"
                if registro["formato_pendente"] or economia_nova != linha["economia"]
                else "telefones-v1"
            )
            anterior = grupos_afetados.get(grupo_id)
            ordem = {"telefones-v1": 1, "telefones-v2": 2, "telefones-v3": 3}
            grupos_afetados[grupo_id] = (
                linha["arquivos_json"],
                anterior[1] if anterior and ordem[anterior[1]] > ordem[revisao]
                else revisao,
            )
            chave = (linha["codigo_condominio"], linha["economia"],
                     linha["nome_condomino"], economia_nova, nome)
            devedores.setdefault(chave, {
                "condominio": linha["condominio"], "telefones": set(),
                "ddd_recuperado": registro["ddd_recuperado"],
            })
            if registro["telefone"]:
                devedores[chave]["telefones"].add((
                    registro["telefone"], registro["ddd_pendente"],
                    registro["formato_pendente"], registro["observacoes"],
                ))

        novos = 0
        locais = 0
        formatos = 0
        for (codigo, economia_antiga, nome_antigo, economia, nome), dados in devedores.items():
            antigo = conexao.execute(
                """SELECT * FROM contatos WHERE codigo_condominio = ?
                   AND economia = ? AND condomino = ? COLLATE NOCASE""",
                (codigo, economia_antiga, nome_antigo),
            ).fetchone()
            atual = conexao.execute(
                """SELECT * FROM contatos WHERE codigo_condominio = ?
                   AND economia = ? AND condomino = ? COLLATE NOCASE""",
                (codigo, economia, nome),
            ).fetchone()
            if antigo and not atual:
                conexao.execute(
                    """UPDATE contatos SET economia = ?, condomino = ?, ativo = 1,
                       atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
                    (economia, nome, antigo["id"]),
                )
                contato_id = antigo["id"]
            elif antigo and atual and antigo["id"] != atual["id"]:
                contato_id = atual["id"]
                conexao.execute(
                    """INSERT OR IGNORE INTO meios_contato
                       (contato_id, tipo, valor, observacoes, ativo)
                       SELECT ?, tipo, valor, observacoes, ativo
                       FROM meios_contato WHERE contato_id = ?""",
                    (contato_id, antigo["id"]),
                )
                conexao.execute(
                    """UPDATE contatos SET telefone = COALESCE(telefone, ?),
                       observacoes = COALESCE(observacoes, ?), ativo = 1,
                       atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
                    (antigo["telefone"], antigo["observacoes"], contato_id),
                )
                conexao.execute(
                    "UPDATE historico_comunicacoes SET contato_id = ? WHERE contato_id = ?",
                    (contato_id, antigo["id"]),
                )
                conexao.execute(
                    "UPDATE contatos SET ativo = 0 WHERE id = ?", (antigo["id"],)
                )
            elif atual:
                contato_id = atual["id"]
                conexao.execute(
                    "UPDATE contatos SET ativo = 1 WHERE id = ?", (contato_id,)
                )
            else:
                contato_id = conexao.execute(
                    """INSERT INTO contatos
                       (codigo_condominio, economia, condominio, condomino)
                       VALUES (?, ?, ?, ?)""",
                    (codigo, economia, dados["condominio"], nome),
                ).lastrowid

            conexao.execute(
                """UPDATE historico_comunicacoes
                   SET economia = ?, condomino = ?
                   WHERE codigo_condominio = ? AND economia = ?
                     AND condomino = ? COLLATE NOCASE""",
                (economia, nome, codigo, economia_antiga, nome_antigo),
            )
            if economia_antiga != economia:
                situacao = conexao.execute(
                    """SELECT ajuizada FROM situacao_unidades
                       WHERE codigo_condominio = ? AND economia = ?""",
                    (codigo, economia_antiga),
                ).fetchone()
                if situacao:
                    conexao.execute(
                        """INSERT INTO situacao_unidades
                           (codigo_condominio, economia, ajuizada)
                           VALUES (?, ?, ?)
                           ON CONFLICT (codigo_condominio, economia)
                           DO UPDATE SET ajuizada = MAX(ajuizada, excluded.ajuizada),
                               atualizado_em = CURRENT_TIMESTAMP""",
                        (codigo, economia, situacao["ajuizada"]),
                    )
            if dados["ddd_recuperado"]:
                locais_antigos = conexao.execute(
                    """SELECT valor FROM meios_contato
                       WHERE contato_id = ? AND tipo = 'Telefone' AND ativo = 1
                         AND (observacoes LIKE 'Importado do PDF%'
                              OR observacoes LIKE 'DDD não informado.%'
                              OR observacoes LIKE 'Formato fora do padrão.%')""",
                    (contato_id,),
                ).fetchall()
                for meio_local in locais_antigos:
                    local = meio_local["valor"]
                    if len(local) in (8, 9):
                        completo = f"{dados['ddd_recuperado']}{local}"
                        try:
                            normalizar_telefone(completo)
                        except ValueError:
                            continue
                        dados["telefones"].add((completo, False, False, None))

            for numero, pendente, formato, nota_extra in dados["telefones"]:
                if pendente:
                    locais += 1
                if formato:
                    formatos += 1
                existente = conexao.execute(
                    """SELECT id FROM meios_contato WHERE contato_id = ?
                       AND tipo = 'Telefone' AND valor = ?""",
                    (contato_id, numero),
                ).fetchone()
                if existente:
                    if nota_extra:
                        conexao.execute(
                            """UPDATE meios_contato SET observacoes = ?,
                               atualizado_em = CURRENT_TIMESTAMP
                               WHERE id = ? AND (observacoes IS NULL
                                   OR TRIM(observacoes) = ''
                                   OR observacoes = 'Importado do PDF.')""",
                            (f"Importado do PDF. {nota_extra}", existente["id"]),
                        )
                    continue
                observacao = (
                    "DDD não informado. Conferir antes de usar."
                    if pendente else
                    "Formato fora do padrão. Conferir antes de usar."
                    if formato else "Importado do PDF."
                )
                if nota_extra:
                    observacao = f"{observacao} {nota_extra}"
                local_anterior = None
                if dados["ddd_recuperado"] and numero.startswith(dados["ddd_recuperado"]):
                    local_anterior = conexao.execute(
                        """SELECT id FROM meios_contato
                           WHERE contato_id = ? AND tipo = 'Telefone'
                             AND valor = ? AND ativo = 1
                             AND (observacoes LIKE 'Importado do PDF%'
                                  OR observacoes LIKE 'DDD não informado.%'
                                  OR observacoes LIKE 'Formato fora do padrão.%')""",
                        (contato_id, numero[len(dados["ddd_recuperado"]):]),
                    ).fetchone()
                if local_anterior:
                    conexao.execute(
                        """UPDATE meios_contato SET valor = ?, observacoes = ?,
                           atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
                        (numero, observacao, local_anterior["id"]),
                    )
                else:
                    conexao.execute(
                        """INSERT INTO meios_contato
                           (contato_id, tipo, valor, observacoes)
                           VALUES (?, 'Telefone', ?, ?)""",
                        (contato_id, numero, observacao),
                    )
                if not pendente and not formato:
                    conexao.execute(
                        """UPDATE contatos SET telefone = ?,
                           atualizado_em = CURRENT_TIMESTAMP
                           WHERE id = ? AND telefone IS NULL""",
                        (numero, contato_id),
                    )
                novos += 1

        # Evita que a mesma imagem PDF gere outra carga apenas por ter sido
        # lida anteriormente pelo parser antigo.
        for grupo_id, (arquivos_json, revisao) in grupos_afetados.items():
            hashes = sorted(item["sha256"] for item in json.loads(arquivos_json))
            hashes.append(revisao)
            digest = hashlib.sha256("\n".join(hashes).encode("ascii")).hexdigest()
            conexao.execute(
                """UPDATE cargas_condominio SET conteudo_sha256 = ?
                   WHERE id = ?""",
                (digest, grupo_id),
            )

    return {"nomes_corrigidos": len(correcoes), "telefones_novos": novos,
            "sem_ddd": locais, "formato_pendente": formatos,
            "copia": str(copia)}


def corrigir_identidade_terezinha_v5(
    *, caminho: str | Path | None = None
) -> dict:
    """Desfaz a atribuição indevida da carga GARDEN 2A (1) ao José.

    A carga anterior, na qual José já constava, permanece intocada. O reparo
    é específico para o PDF e os dois títulos afetados pelo parser v5.
    """

    arquivo = _resolver_caminho(caminho)
    sha_pdf = "442ce2fd05cfedb32aee6e17bf479dd76a0d25fcb8317e07a14582add0427c7a"
    numeros = ("0000026477829", "0000026489930")
    with conectar(arquivo) as conexao:
        linhas = conexao.execute(
            """SELECT lc.id, cc.condominio, cc.arquivos_json
               FROM lancamentos_carga lc
               JOIN cargas_condominio cc ON cc.id = lc.carga_condominio_id
               WHERE cc.codigo_condominio = '00147'
                 AND lc.economia = '108 (01)'
                 AND lc.nome_condomino = 'JOSE DE OLIVEIRA'
                 AND lc.nosso_numero IN (?, ?)""",
            numeros,
        ).fetchall()
        afetadas = [
            linha for linha in linhas
            if any(item.get("sha256") == sha_pdf
                   for item in json.loads(linha["arquivos_json"]))
        ]
        if not afetadas:
            return {"nomes_corrigidos": 0, "copia": None}

        copia = arquivo.with_name(f"{arquivo.stem}.antes_parser_v6.db")
        if not copia.exists():
            with closing(sqlite3.connect(copia)) as destino:
                conexao.backup(destino)

        nome = "Terezinha: Procuradora"
        telefone = "5192287193"
        for linha in afetadas:
            conexao.execute(
                "UPDATE lancamentos_carga SET nome_condomino = ? WHERE id = ?",
                (nome, linha["id"]),
            )

        representante = conexao.execute(
            """SELECT id FROM contatos WHERE codigo_condominio = '00147'
               AND economia = '108 (01)' AND condomino = ? COLLATE NOCASE""",
            (nome,),
        ).fetchone()
        bruto = conexao.execute(
            """SELECT id FROM contatos WHERE codigo_condominio = '00147'
               AND economia = '108 (01)'
               AND condomino = ? COLLATE NOCASE""",
            ("Terezinha: Procuradora (51)9228-7193-JOSE DE OLIVEIRA",),
        ).fetchone()
        if representante:
            contato_id = representante["id"]
        elif bruto:
            contato_id = bruto["id"]
            conexao.execute(
                """UPDATE contatos SET condomino = ?, ativo = 1,
                   telefone = COALESCE(telefone, ?),
                   atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
                (nome, telefone, contato_id),
            )
        else:
            contato_id = conexao.execute(
                """INSERT INTO contatos
                   (codigo_condominio, economia, condominio, condomino, telefone)
                   VALUES ('00147', '108 (01)', ?, ?, ?)""",
                (afetadas[0]["condominio"], nome, telefone),
            ).lastrowid
        conexao.execute(
            """UPDATE contatos SET ativo = 1, telefone = COALESCE(telefone, ?),
               atualizado_em = CURRENT_TIMESTAMP WHERE id = ?""",
            (telefone, contato_id),
        )
        conexao.execute(
            """INSERT OR IGNORE INTO meios_contato
               (contato_id, tipo, valor, observacoes)
               VALUES (?, 'Telefone', ?, 'Importado do PDF; Terezinha (procuradora).')""",
            (contato_id, telefone),
        )

        # No v5 este telefone foi copiado para o José. Removemos somente a
        # cópia identificada como importação da procuradora, nunca os meios
        # cadastrados manualmente nem o contato histórico de José.
        jose = conexao.execute(
            """SELECT id, telefone FROM contatos WHERE codigo_condominio = '00147'
               AND economia = '108 (01)' AND condomino = 'JOSE DE OLIVEIRA'"""
        ).fetchone()
        if jose:
            meios_v5 = conexao.execute(
                """SELECT id FROM meios_contato WHERE contato_id = ?
                   AND tipo = 'Telefone' AND valor = ?
                   AND observacoes LIKE 'Importado do PDF.%Terezinha (procuradora)%'""",
                (jose["id"], telefone),
            ).fetchall()
            for meio in meios_v5:
                conexao.execute(
                    "UPDATE meios_contato SET ativo = 0 WHERE id = ?", (meio["id"],)
                )
            if meios_v5 and jose["telefone"] == telefone:
                conexao.execute(
                    "UPDATE contatos SET telefone = NULL WHERE id = ?", (jose["id"],)
                )

    return {"nomes_corrigidos": len(afetadas), "copia": str(copia)}


def registrar_comunicacao(
    codigo_condominio: object,
    economia: object,
    condomino: object,
    canal: object,
    acao: object,
    *,
    contato_id: int | None = None,
    competencia: object | None = None,
    vencimento: date | datetime | str | None = None,
    valor: int | float | str | None = None,
    resultado: object | None = None,
    valor_canal: object | None = None,
    observacoes: object | None = None,
    realizado_em: datetime | str | None = None,
    caminho: str | Path | None = None,
) -> int:
    """Registra uma acao de comunicacao e retorna seu identificador."""

    with conectar(caminho) as conexao:
        cursor = conexao.execute(
            """
            INSERT INTO historico_comunicacoes (
                contato_id, codigo_condominio, economia, condomino,
                competencia, vencimento, valor, canal, valor_canal, acao,
                resultado, observacoes, realizado_em
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, COALESCE(?, CURRENT_TIMESTAMP))
            """,
            (
                contato_id,
                _texto_obrigatorio(codigo_condominio, "codigo_condominio"),
                _texto_obrigatorio(economia, "economia"),
                _texto_obrigatorio(condomino, "condomino"),
                _texto_opcional(competencia),
                _data_iso(vencimento),
                valor,
                _texto_obrigatorio(canal, "canal"),
                _texto_opcional(valor_canal),
                _texto_obrigatorio(acao, "acao"),
                _texto_opcional(resultado),
                _texto_opcional(observacoes),
                _data_iso(realizado_em),
            ),
        )
        identificador = cursor.lastrowid

    return int(identificador)


def listar_historico_comunicacoes(
    *,
    contato_id: int | None = None,
    codigo_condominio: object | None = None,
    economia: object | None = None,
    condomino: object | None = None,
    limite: int = 200,
    caminho: str | Path | None = None,
) -> list[dict]:
    """Lista o historico, permitindo filtrar contato ou unidade."""

    if limite < 1:
        raise ValueError("limite deve ser maior que zero.")

    filtros = []
    parametros: list[object] = []
    if contato_id is not None:
        filtros.append("contato_id = ?")
        parametros.append(contato_id)
    if codigo_condominio is not None:
        filtros.append("codigo_condominio = ?")
        parametros.append(
            _texto_obrigatorio(codigo_condominio, "codigo_condominio")
        )
    if economia is not None:
        filtros.append("economia = ?")
        parametros.append(_texto_obrigatorio(economia, "economia"))
    if condomino is not None:
        filtros.append("condomino = ? COLLATE NOCASE")
        parametros.append(_texto_obrigatorio(condomino, "condomino"))

    consulta = "SELECT * FROM historico_comunicacoes"
    if filtros:
        consulta += " WHERE " + " AND ".join(filtros)
    consulta += " ORDER BY realizado_em DESC, id DESC LIMIT ?"
    parametros.append(limite)

    with conectar(caminho) as conexao:
        registros = conexao.execute(consulta, parametros).fetchall()
    return [dict(registro) for registro in registros]
