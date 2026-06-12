"""
Módulo responsável pela leitura do arquivo exportado pelo Gevis.

Utiliza pandas para carregar o arquivo .xlsx e retornar um DataFrame.

Futuras melhorias:
- Suportar arquivos .csv e outros formatos exportados pelo Gevis
- Validar o schema do arquivo de entrada automaticamente
"""

from pathlib import Path

import pandas as pd

from config import COLUMN_MAPPING
from logger import get_logger

logger = get_logger()


def read_gevis(file_path: Path) -> pd.DataFrame:
    """
    Lê o arquivo Excel exportado do Gevis.

    Args:
        file_path: Caminho para o arquivo gevis.xlsx.

    Returns:
        DataFrame com os dados do Gevis.

    Raises:
        FileNotFoundError: Se o arquivo não for encontrado.
        ValueError: Se o arquivo estiver vazio ou corrompido.
    """
    if not file_path.exists():
        raise FileNotFoundError(
            f"Arquivo Gevis não encontrado: {file_path}\n"
            "Exporte a tabela do sistema Gevis e salve-a em data/input/gevis.xlsx"
        )

    header_row = _detect_header_row(file_path)
    df = pd.read_excel(file_path, engine="openpyxl", header=header_row)

    # Normaliza nomes das colunas e remove colunas vazias/auxiliares do export.
    df.columns = [str(col).strip() for col in df.columns]
    df = df.loc[
        :,
        [
            col
            for col in df.columns
            if col and col != "nan" and not col.lower().startswith("unnamed:")
        ],
    ]

    if df.empty:
        raise ValueError(f"O arquivo '{file_path}' está vazio.")

    logger.info("Arquivo Gevis carregado")
    logger.info(f"{len(df)} registros encontrados")

    return df


def _detect_header_row(file_path: Path) -> int:
    """
    Detecta a linha de cabeçalho do arquivo Gevis.

    O export padrão do Gevis costuma trazer uma linha de título antes do cabeçalho
    real. Esta função procura, nas primeiras linhas, aquela com maior aderência
    às colunas esperadas no mapeamento.
    """
    preview = pd.read_excel(file_path, engine="openpyxl", header=None, nrows=10)
    expected = {str(col).strip().lower() for col in COLUMN_MAPPING.keys()}

    best_index = 0
    best_score = -1

    for idx, row in preview.iterrows():
        normalized = {str(value).strip().lower() for value in row.tolist() if pd.notna(value)}
        score = len(expected.intersection(normalized))
        if score > best_score:
            best_score = score
            best_index = int(idx)

    return best_index
