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

from openpyxl import load_workbook

SOURCE_URL = (
    "https://www.gov.br/anp/pt-br/centrais-de-conteudo/paineis-dinamicos-da-anp/"
    "arquivos-dados-brutos-do-painel-dinamico-da-fiscalizacao-do-abastecimento-da-sfi/"
    "dados-fisc-a-partir-2019.xlsx"
)

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


def main() -> None:
    print(f"Baixando {SOURCE_URL}")
    request = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "octano-anp-data/1.0"})
    with urllib.request.urlopen(request, timeout=300) as response:
        payload = response.read()
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

    now = datetime.now(timezone.utc)
    metadata = {
        "updatedAt": now.strftime("%Y-%m-%d"),
        "updatedAtISO": now.isoformat(timespec="seconds"),
        "rows": count,
        "sha256": hashlib.sha256(csv_bytes).hexdigest(),
        "source": SOURCE_URL,
    }
    (DATA_DIR / "metadata.json").write_text(
        json.dumps(metadata, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(f"OK: data/base_anp.csv ({len(csv_bytes)/1e6:.1f} MB) e data/metadata.json gravados")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:  # noqa: BLE001 - queremos falhar o workflow com mensagem clara
        print(f"ERRO: {error}", file=sys.stderr)
        sys.exit(1)
