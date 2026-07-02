#!/usr/bin/env python3
"""Baixa os dados brutos de fiscalização do abastecimento da ANP (xlsx),
valida, converte para CSV no formato consumido pelo app octano e grava
data/base_anp.csv + data/metadata.json.

Sai com código != 0 em qualquer problema (download, header inesperado,
contagem de linhas suspeita), o que faz o workflow falhar e disparar a
notificação de erro.
"""

import csv
import hashlib
import io
import json
import sys
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import xlrd
from openpyxl import load_workbook

SOURCE_URL = (
    "https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/"
    "arquivos-dados-brutos-do-painel-dinamico-da-fiscalizacao-do-abastecimento-da-sfi/"
    "dados-fisc-a-partir-2019.xlsx"
)

# API pública de revendedores da ANP: cadastro oficial com CNPJ, razão social,
# bandeira e COORDENADAS de cada posto — usada pelo app para validar
# geograficamente o cruzamento Google ⇄ ANP.
CADASTRO_API = "https://revendedoresapi.anp.gov.br/v1/combustivel?numeropagina={page}"
CADASTRO_MIN_ROWS = 30_000  # a base tem ~46 mil postos; menos que isso é problema

EXPECTED_HEADER = [
    "UF",
    "Município",
    "Bairro",
    "ENDEREÇO",
    "CNPJ/CPF",
    "Agente Econômico",
    "Segmento Fiscalizado",
    "DATA DO DF",
    "Número do Documento",
    "Procedimento de Fiscalização",
    "Resultado",
]

# A base de abr/2025 tinha ~225 mil linhas e só cresce (dados de 2019 em diante).
# Se a ANP publicar um arquivo muito menor, algo mudou no formato: melhor falhar.
MIN_ROWS = 200_000

# --- Preços de DISTRIBUIÇÃO (semanal, desde ago/2020): Brasil e por UF. ---
PRECOS_DIST_URLS = [
    ("BRASIL", "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/"
               "precos/pdc/semanal/combustiveis-liquidos-brasil.xlsx"),
    ("ESTADOS", "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/"
                "precos/pdc/semanal/combustiveis-liquidos-estados.xlsx"),
]
PRECOS_DIST_MIN_ROWS = 20_000  # jul/2026: ~42 mil (Brasil + 27 UFs × 5 produtos)

# --- Preços de PRODUTORES/importadores (semanal, desde 2013): por região. ---
PRECOS_PRODUTORES_URL = (
    "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/"
    "precos/ppidp/precos-medios-ponderados-semanais-2013.xls"
)
# A planilha tem 21 produtos (asfalto, nafta, QAV...); publicamos só os
# relevantes para o consumidor do app.
PRODUTORES_PRODUTOS = (
    "Gasolina A Comum",
    "Gasolina A Premium",
    "Óleo Diesel S-10",
    "Óleo Diesel S-500",
    "Biodiesel B-100",
    "Gás Liquefeito de Petróleo",
)
PRECOS_PROD_MIN_ROWS = 3_000

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def format_cell(value, is_date_column: bool) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        # Mesmo formato do CSV histórico do app: M/D/YYYY sem zero à esquerda.
        return f"{value.month}/{value.day}/{value.year}"
    text = str(value)
    if is_date_column and text and "/" not in text:
        # Data que veio como texto/serial inesperado: falha em vez de corromper.
        raise ValueError(f"Valor de data não reconhecido: {text!r}")
    # O parser de CSV do app é linha-a-linha e não trata aspas escapadas:
    # remove quebras de linha e aspas duplas de dentro das células.
    return text.replace("\r", " ").replace("\n", " ").replace('"', "'").strip()


def fetch_cadastro() -> bytes:
    """Baixa o cadastro completo de postos da API de revendedores (paginada,
    5.000 registros por página) e devolve um CSV `;`-separado com CNPJ,
    razão social, bandeira, UF, município, latitude e longitude."""
    lines = ["CNPJ;RAZAOSOCIAL;BANDEIRA;UF;MUNICIPIO;LATITUDE;LONGITUDE"]
    count = 0
    page = 1
    while page <= 30:  # trava de segurança: hoje são ~10 páginas
        url = CADASTRO_API.format(page=page)
        print(f"Cadastro: página {page}...")
        request = urllib.request.Request(url, headers={
            "User-Agent": "octano-anp-data/1.0", "Accept": "application/json",
        })
        with urllib.request.urlopen(request, timeout=300) as response:
            payload = json.loads(response.read())
        rows = payload.get("data") or []
        if not rows:
            break
        for row in rows:
            def clean(value):
                return str(value or "").replace(";", ",").replace("\n", " ").strip()
            lines.append(";".join([
                clean(row.get("cnpj")),
                clean(row.get("razaoSocial")),
                clean(row.get("distribuidora")),
                clean(row.get("uf")),
                clean(row.get("municipio")),
                clean(row.get("latitude")),
                clean(row.get("longitude")),
            ]))
            count += 1
        page += 1

    print(f"Cadastro: {count} postos")
    if count < CADASTRO_MIN_ROWS:
        raise RuntimeError(
            f"Cadastro com apenas {count} postos (mínimo {CADASTRO_MIN_ROWS}) — abortando"
        )
    return ("\n".join(lines) + "\n").encode("utf-8")


def download(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "octano-anp-data/1.0"})
    with urllib.request.urlopen(request, timeout=300) as response:
        return response.read()


def fetch_precos_distribuicao() -> bytes:
    """Preços semanais de DISTRIBUIÇÃO (Brasil + por UF) num único CSV:
    INICIO;FIM;LOCAL;PRODUTO;UNIDADE;PRECO — LOCAL é "BRASIL" ou o nome da UF.
    Fonte: relatórios semanais de combustíveis líquidos da SDL/ANP."""
    lines = ["INICIO;FIM;LOCAL;PRODUTO;UNIDADE;PRECO"]
    count = 0

    for scope, url in PRECOS_DIST_URLS:
        print(f"Preços distribuição ({scope}): baixando...")
        workbook = load_workbook(io.BytesIO(download(url)), read_only=True, data_only=True)
        sheet = workbook.worksheets[0]

        col: dict[str, int] = {}
        price_index = -1
        for row in sheet.iter_rows(values_only=True):
            # As primeiras linhas são um cabeçalho institucional; a tabela
            # começa na linha cujo primeiro valor é "DATA INICIAL".
            if not col:
                if row and str(row[0]).strip().upper() == "DATA INICIAL":
                    names = [str(c).strip().upper() if c else "" for c in row]
                    col = {name: i for i, name in enumerate(names)}
                    # "PREÇO MÉDIO DISTRIBUIÇÃO" no arquivo de estados,
                    # "PREÇO MÉDIO DE DISTRIBUIÇÃO" no do Brasil.
                    price_index = next(
                        i for i, n in enumerate(names) if n.startswith("PREÇO MÉDIO")
                    )
                continue

            inicio = row[col["DATA INICIAL"]]
            fim = row[col["DATA FINAL"]]
            preco = row[price_index] if price_index < len(row) else None
            if not isinstance(inicio, datetime) or not isinstance(preco, (int, float)):
                continue

            local = str(row[col["ESTADO"]]).strip().upper() if "ESTADO" in col else "BRASIL"
            produto = str(row[col["PRODUTO"]]).strip().upper()
            unidade = (
                str(row[col["UNIDADE DE MEDIDA"]]).strip()
                if "UNIDADE DE MEDIDA" in col else ""
            )
            lines.append(";".join([
                inicio.strftime("%Y-%m-%d"),
                fim.strftime("%Y-%m-%d") if isinstance(fim, datetime) else "",
                local,
                produto,
                unidade,
                f"{float(preco):.3f}",
            ]))
            count += 1

        if not col:
            raise RuntimeError(f"Cabeçalho de preços não encontrado em {url}")

    print(f"Preços distribuição: {count} linhas")
    if count < PRECOS_DIST_MIN_ROWS:
        raise RuntimeError(
            f"Preços de distribuição com {count} linhas (mínimo {PRECOS_DIST_MIN_ROWS}) — abortando"
        )
    return ("\n".join(lines) + "\n").encode("utf-8")


def fetch_precos_produtores() -> bytes:
    """Preços semanais de PRODUTORES/importadores por região (desde 2013):
    INICIO;FIM;REGIAO;PRODUTO;UNIDADE;PRECO. Publica apenas os produtos de
    interesse do consumidor (PRODUTORES_PRODUTOS); valores sigilosos vêm
    como "***" na planilha e são descartados."""
    print("Preços produtores: baixando...")
    book = xlrd.open_workbook(file_contents=download(PRECOS_PRODUTORES_URL))
    sheet = book.sheet_by_index(0)

    # Colunas fixas da planilha: produto, início, fim, 5 regiões e Brasil.
    regioes = ["NORTE", "NORDESTE", "CENTRO-OESTE", "SUL", "SUDESTE", "BRASIL"]
    lines = ["INICIO;FIM;REGIAO;PRODUTO;UNIDADE;PRECO"]
    count = 0

    for r in range(sheet.nrows):
        row = sheet.row_values(r)
        if len(row) < 9:
            continue
        produto_raw = str(row[0]).strip()
        if not produto_raw.startswith(PRODUTORES_PRODUTOS):
            continue
        try:
            inicio = xlrd.xldate_as_datetime(row[1], book.datemode)
            fim = xlrd.xldate_as_datetime(row[2], book.datemode)
        except (TypeError, ValueError, xlrd.xldate.XLDateError):
            continue

        # "Gasolina A Comum (R$/litro)" → nome limpo + unidade separada.
        nome, unidade = produto_raw, ""
        if "(" in produto_raw and produto_raw.endswith(")"):
            nome, _, resto = produto_raw.partition("(")
            nome = nome.strip().upper()
            unidade = resto[:-1].strip()

        for regiao, valor in zip(regioes, row[3:9]):
            if isinstance(valor, (int, float)) and valor > 0:
                lines.append(";".join([
                    inicio.strftime("%Y-%m-%d"),
                    fim.strftime("%Y-%m-%d"),
                    regiao,
                    nome,
                    unidade,
                    f"{float(valor):.4f}",
                ]))
                count += 1

    print(f"Preços produtores: {count} linhas")
    if count < PRECOS_PROD_MIN_ROWS:
        raise RuntimeError(
            f"Preços de produtores com {count} linhas (mínimo {PRECOS_PROD_MIN_ROWS}) — abortando"
        )
    return ("\n".join(lines) + "\n").encode("utf-8")


def main() -> None:
    print(f"Baixando {SOURCE_URL}")
    payload = download(SOURCE_URL)
    print(f"Download concluído: {len(payload)/1e6:.1f} MB")
    if len(payload) < 1_000_000:
        raise RuntimeError(f"Arquivo suspeito de {len(payload)} bytes — esperado > 1 MB")

    workbook = load_workbook(io.BytesIO(payload), read_only=True, data_only=True)
    sheet = workbook.worksheets[0]

    rows_iter = sheet.iter_rows(values_only=True)
    header = [str(cell).strip() if cell is not None else "" for cell in next(rows_iter)]
    if header != EXPECTED_HEADER:
        raise RuntimeError(
            "Header do xlsx mudou — atualizar o app antes de publicar.\n"
            f"Esperado: {EXPECTED_HEADER}\nRecebido: {header}"
        )

    date_index = EXPECTED_HEADER.index("DATA DO DF")
    buffer = io.StringIO()
    writer = csv.writer(buffer, quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerow(EXPECTED_HEADER)

    count = 0
    for row in rows_iter:
        if row is None or all(cell is None for cell in row):
            continue
        cells = list(row[: len(EXPECTED_HEADER)])
        cells += [None] * (len(EXPECTED_HEADER) - len(cells))
        writer.writerow(
            format_cell(cell, index == date_index) for index, cell in enumerate(cells)
        )
        count += 1

    print(f"Convertidas {count} linhas")
    if count < MIN_ROWS:
        raise RuntimeError(f"Apenas {count} linhas (mínimo esperado {MIN_ROWS}) — abortando")

    csv_bytes = buffer.getvalue().encode("utf-8")
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "base_anp.csv").write_bytes(csv_bytes)

    cadastro_bytes = fetch_cadastro()
    (DATA_DIR / "cadastro_postos.csv").write_bytes(cadastro_bytes)

    precos_dist_bytes = fetch_precos_distribuicao()
    (DATA_DIR / "precos_distribuicao.csv").write_bytes(precos_dist_bytes)

    precos_prod_bytes = fetch_precos_produtores()
    (DATA_DIR / "precos_produtores.csv").write_bytes(precos_prod_bytes)

    now = datetime.now(timezone.utc)
    metadata = {
        "updatedAt": now.strftime("%Y-%m-%d"),
        "updatedAtISO": now.isoformat(timespec="seconds"),
        "rows": count,
        "sha256": hashlib.sha256(csv_bytes).hexdigest(),
        "source": SOURCE_URL,
        "cadastroRows": cadastro_bytes.count(b"\n") - 1,
        "cadastroSha256": hashlib.sha256(cadastro_bytes).hexdigest(),
        "cadastroSource": "https://revendedoresapi.anp.gov.br/v1/combustivel",
        "precosDistRows": precos_dist_bytes.count(b"\n") - 1,
        "precosDistSha256": hashlib.sha256(precos_dist_bytes).hexdigest(),
        "precosProdRows": precos_prod_bytes.count(b"\n") - 1,
        "precosProdSha256": hashlib.sha256(precos_prod_bytes).hexdigest(),
        "precosSource": "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos",
    }
    (DATA_DIR / "metadata.json").write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(
        f"OK: base_anp.csv ({len(csv_bytes)/1e6:.1f} MB), "
        f"cadastro_postos.csv ({len(cadastro_bytes)/1e6:.1f} MB), "
        f"precos_distribuicao.csv ({len(precos_dist_bytes)/1e6:.1f} MB), "
        f"precos_produtores.csv ({len(precos_prod_bytes)/1e6:.1f} MB) "
        f"e metadata.json gravados"
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as error:  # noqa: BLE001 - queremos falhar o workflow com mensagem clara
        print(f"ERRO: {error}", file=sys.stderr)
        sys.exit(1)
