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
from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import OUTPUT_PATH, UNIQUE_KEY_COLUMN, VISTORIAS_TEMPLATE_PATH
from logger import get_logger

logger = get_logger()
TARGET_SHEET_NAME = "VISTORIAS"


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
    sheet = _select_target_sheet(workbook)

    header_row = _find_header_row(sheet)
    header_map = _build_header_map(sheet, header_row)
    existing_keys = _collect_existing_keys(sheet, header_map, header_row)

    added, skipped = _append_new_rows(
        sheet, df_new, header_map, existing_keys, header_row
    )

    saved_path = _save_workbook(workbook, output_path)

    logger.info(f"{skipped} registros ignorados por duplicidade")
    logger.info(f"{added} registros adicionados")
    logger.info(f"Registros escritos na aba: {sheet.title}")
    logger.info(f"Arquivo salvo com sucesso: {saved_path}")


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


def _save_workbook(workbook, output_path: Path) -> Path:
    """
    Salva o workbook no caminho principal ou em um arquivo alternativo se o
    destino estiver bloqueado por outro programa.
    """
    try:
        workbook.save(output_path)
        return output_path
    except PermissionError:
        fallback_path = output_path.with_name(
            f"{output_path.stem}_gerada{output_path.suffix}"
        )
        workbook.save(fallback_path)
        logger.warning(
            f"Não foi possível sobrescrever {output_path}. Arquivo salvo em: {fallback_path}"
        )
        return fallback_path


def _build_header_map(sheet: Worksheet, header_row: int) -> dict[str, int]:
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
        for cell in sheet[header_row]
        if cell.value is not None
    }


def _select_target_sheet(workbook) -> Worksheet:
    """
    Seleciona explicitamente a aba de dados VISTORIAS.
    """
    if TARGET_SHEET_NAME not in workbook.sheetnames:
        raise ValueError(
            f"Aba '{TARGET_SHEET_NAME}' não encontrada na planilha de saída."
        )
    return workbook[TARGET_SHEET_NAME]


def _find_header_row(sheet: Worksheet) -> int:
    """Encontra a linha de cabeçalho pela coluna-chave PLACA/CHASSI."""
    target = UNIQUE_KEY_COLUMN.strip().upper()
    for row_idx in range(1, min(sheet.max_row, 50) + 1):
        row_values = [
            str(sheet.cell(row_idx, c).value).strip().upper()
            for c in range(1, sheet.max_column + 1)
            if sheet.cell(row_idx, c).value is not None
        ]
        if target in row_values:
            return row_idx

    raise ValueError(
        f"Não foi possível localizar o cabeçalho '{UNIQUE_KEY_COLUMN}' na aba {sheet.title}."
    )


def _collect_existing_keys(
    sheet: Worksheet, header_map: dict[str, int], header_row: int
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

    for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        key_value = row[key_col_idx - 1]
        if key_value is not None:
            existing_keys.add(str(key_value).strip().upper())

    return existing_keys


def _append_new_rows(
    sheet: Worksheet,
    df_new: pd.DataFrame,
    header_map: dict[str, int],
    existing_keys: set[str],
    header_row: int,
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
    next_row = _next_data_row(sheet, header_map, header_row)

    key_col = UNIQUE_KEY_COLUMN

    for _, record in df_new.iterrows():
        plate = str(record.get(key_col, "")).strip().upper()
        wrote_cell = False

        if plate and plate in existing_keys:
            skipped += 1
            continue

        for col_name, col_idx in header_map.items():
            if col_name in record.index:
                sheet.cell(row=next_row, column=col_idx, value=record[col_name])
                wrote_cell = True

        if not wrote_cell:
            raise ValueError(
                "Nenhuma coluna do DataFrame corresponde aos cabeçalhos da aba de destino."
            )

        if plate:
            existing_keys.add(plate)

        _copy_row_style(sheet, next_row - 1, next_row)

        next_row += 1
        added += 1

    return added, skipped


def _next_data_row(sheet: Worksheet, header_map: dict[str, int], header_row: int) -> int:
    """Calcula a próxima linha livre com base na coluna-chave da tabela."""
    key_col_idx = header_map.get(UNIQUE_KEY_COLUMN)
    if key_col_idx is None:
        return header_row + 1

    last_data_row = header_row
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        key_value = sheet.cell(row=row_idx, column=key_col_idx).value
        if key_value not in (None, ""):
            last_data_row = row_idx

    return last_data_row + 1


def _copy_row_style(sheet: Worksheet, source_row: int, target_row: int) -> None:
    """Replica o estilo da linha anterior para manter o visual do template."""
    if source_row <= 0:
        return

    for col_idx in range(1, sheet.max_column + 1):
        source_cell = sheet.cell(row=source_row, column=col_idx)
        target_cell = sheet.cell(row=target_row, column=col_idx)

        target_cell.font = copy(source_cell.font)
        target_cell.fill = copy(source_cell.fill)
        target_cell.border = copy(source_cell.border)
        target_cell.alignment = copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
