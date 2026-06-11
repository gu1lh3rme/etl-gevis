"""
Módulo responsável pela escrita dos novos registros na planilha VISTORIAS.

Utiliza openpyxl para:
- Ler a planilha modelo (preservando formatação)
- Identificar duplicatas pela coluna PLACA/CHASSI
- Adicionar apenas registros novos ao final da planilha
- Salvar o resultado em data/output/VISTORIAS_ATUALIZADA.xlsx

Futuras melhorias:
- Preservar estilos de células (cores, bordas) ao inserir novas linhas
- Suportar múltiplas abas na planilha modelo
- Gerar relatório de auditoria com os registros inseridos/ignorados
"""

import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import OUTPUT_PATH, UNIQUE_KEY_COLUMN, VISTORIAS_TEMPLATE_PATH
from logger import get_logger

logger = get_logger()


def write_to_vistorias(df_new: pd.DataFrame) -> None:
    """
    Escreve os novos registros na planilha VISTORIAS, evitando duplicatas.

    O arquivo modelo (templates/VISTORIAS.xlsx) nunca é alterado.
    O resultado é salvo em data/output/VISTORIAS_ATUALIZADA.xlsx.

    Args:
        df_new: DataFrame com os registros mapeados prontos para inserção.

    Raises:
        FileNotFoundError: Se a planilha modelo não for encontrada.
    """
    _ensure_template_exists()

    output_path = _prepare_output_file()
    workbook = load_workbook(output_path)
    sheet = workbook.active

    header_map = _build_header_map(sheet)
    existing_keys = _collect_existing_keys(sheet, header_map)

    added, skipped = _append_new_rows(sheet, df_new, header_map, existing_keys)

    workbook.save(output_path)

    logger.info(f"{skipped} registros ignorados por duplicidade")
    logger.info(f"{added} registros adicionados")
    logger.info(f"Arquivo salvo com sucesso: {output_path}")


# ---------------------------------------------------------------------------
# Helpers privados
# ---------------------------------------------------------------------------


def _ensure_template_exists() -> None:
    """Verifica se a planilha modelo existe."""
    if not VISTORIAS_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Planilha modelo não encontrada: {VISTORIAS_TEMPLATE_PATH}\n"
            "Coloque o arquivo VISTORIAS.xlsx na pasta templates/"
        )


def _prepare_output_file() -> Path:
    """
    Garante que o arquivo de saída existe.

    Se ainda não existir, cria uma cópia a partir do template.
    Cria o diretório de saída se necessário.

    Returns:
        Caminho do arquivo de saída.
    """
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not OUTPUT_PATH.exists():
        shutil.copy2(VISTORIAS_TEMPLATE_PATH, OUTPUT_PATH)
        logger.info(f"Cópia do template criada em: {OUTPUT_PATH}")

    return OUTPUT_PATH


def _build_header_map(sheet: Worksheet) -> dict[str, int]:
    """
    Lê a primeira linha da planilha e retorna um dicionário
    {nome_coluna: índice_coluna (1-based)}.

    Args:
        sheet: Aba ativa da planilha.

    Returns:
        Dicionário mapeando nome da coluna para seu índice.
    """
    return {
        cell.value: cell.column
        for cell in sheet[1]
        if cell.value is not None
    }


def _collect_existing_keys(
    sheet: Worksheet, header_map: dict[str, int]
) -> set[str]:
    """
    Coleta os valores já existentes na coluna de chave única (PLACA/CHASSI).

    Args:
        sheet: Aba ativa da planilha.
        header_map: Mapa de nomes de coluna para índices.

    Returns:
        Conjunto de chaves já presentes na planilha.
    """
    if UNIQUE_KEY_COLUMN not in header_map:
        return set()

    key_col_idx = header_map[UNIQUE_KEY_COLUMN]
    existing_keys: set[str] = set()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        key_value = row[key_col_idx - 1]
        if key_value is not None:
            existing_keys.add(str(key_value).strip().upper())

    return existing_keys


def _append_new_rows(
    sheet: Worksheet,
    df_new: pd.DataFrame,
    header_map: dict[str, int],
    existing_keys: set[str],
) -> tuple[int, int]:
    """
    Itera sobre o DataFrame e insere apenas registros novos na planilha.

    Args:
        sheet: Aba ativa da planilha.
        df_new: DataFrame com os dados mapeados.
        header_map: Mapa de nomes de coluna para índices.
        existing_keys: Chaves já presentes na planilha.

    Returns:
        Tupla (quantidade_adicionados, quantidade_ignorados).
    """
    added = 0
    skipped = 0
    next_row = sheet.max_row + 1

    key_col = UNIQUE_KEY_COLUMN

    for _, record in df_new.iterrows():
        plate = str(record.get(key_col, "")).strip().upper()

        if plate and plate in existing_keys:
            skipped += 1
            continue

        for col_name, col_idx in header_map.items():
            if col_name in record.index:
                sheet.cell(row=next_row, column=col_idx, value=record[col_name])

        if plate:
            existing_keys.add(plate)

        next_row += 1
        added += 1

    return added, skipped
