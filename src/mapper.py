"""
Módulo responsável pelo mapeamento dos dados do Gevis para o formato VISTORIAS.

Aplica o dicionário COLUMN_MAPPING definido em config.py e injeta os valores
fixos definidos em FIXED_VALUES.

Futuras melhorias:
- Suportar transformações de dados por coluna (ex.: formatar datas, normalizar placas)
- Validar os valores antes da inserção (ex.: formato de placa, valor monetário)
"""

import re

import pandas as pd

from config import COLUMN_MAPPING, FIXED_VALUES
from logger import get_logger

logger = get_logger()


def map_gevis_to_vistorias(df_gevis: pd.DataFrame) -> pd.DataFrame:
    """
    Mapeia as colunas do DataFrame do Gevis para as colunas da planilha VISTORIAS.

    Também injeta os valores fixos configurados em FIXED_VALUES.

    Args:
        df_gevis: DataFrame lido do arquivo gevis.xlsx.

    Returns:
        DataFrame com as colunas renomeadas e valores fixos preenchidos.

    Raises:
        KeyError: Se alguma coluna de origem não existir no arquivo do Gevis.
    """
    resolved_mapping = _resolve_source_mapping(df_gevis)
    df_mapped = df_gevis.rename(columns=resolved_mapping)

    # Mantém apenas as colunas de destino mapeadas
    target_columns = list(COLUMN_MAPPING.values())
    df_result = df_mapped[target_columns].copy()

    # Injeta valores fixos
    for column, value in FIXED_VALUES.items():
        df_result[column] = value

    return df_result


def _resolve_source_mapping(df: pd.DataFrame) -> dict[str, str]:
    """
    Resolve as colunas de origem do Gevis para o mapeamento de destino.

    A resolução usa comparação normalizada (sem diferença de caixa,
    espaços extras e pontuação simples) para suportar pequenas variações
    no cabeçalho exportado.

    Args:
        df: DataFrame a ser validado.

    Returns:
        Dicionário no formato {coluna_real_no_df: coluna_destino}.

    Raises:
        KeyError: Se uma ou mais colunas esperadas estiverem ausentes.
    """
    normalized_to_real: dict[str, str] = {
        _normalize_column_name(col): col for col in df.columns
    }

    missing: list[str] = []
    resolved: dict[str, str] = {}

    for source_col, target_col in COLUMN_MAPPING.items():
        normalized_source = _normalize_column_name(source_col)
        real_col = normalized_to_real.get(normalized_source)

        if real_col is None:
            missing.append(source_col)
            continue

        resolved[real_col] = target_col

    if missing:
        raise KeyError(
            f"Coluna(s) não encontrada(s) no arquivo Gevis: {missing}\n"
            "Verifique o dicionário COLUMN_MAPPING em src/config.py."
        )

    return resolved


def _normalize_column_name(name: str) -> str:
    """Normaliza nome de coluna para comparação resiliente."""
    text = str(name).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w ]", "", text)
    return text
