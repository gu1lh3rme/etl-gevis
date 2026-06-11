"""
Módulo responsável pelo mapeamento dos dados do Gevis para o formato VISTORIAS.

Aplica o dicionário COLUMN_MAPPING definido em config.py e injeta os valores
fixos definidos em FIXED_VALUES.

Futuras melhorias:
- Suportar transformações de dados por coluna (ex.: formatar datas, normalizar placas)
- Validar os valores antes da inserção (ex.: formato de placa, valor monetário)
"""

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
    _validate_source_columns(df_gevis)

    df_mapped = df_gevis.rename(columns=COLUMN_MAPPING)

    # Mantém apenas as colunas de destino mapeadas
    target_columns = list(COLUMN_MAPPING.values())
    df_result = df_mapped[target_columns].copy()

    # Injeta valores fixos
    for column, value in FIXED_VALUES.items():
        df_result[column] = value

    return df_result


def _validate_source_columns(df: pd.DataFrame) -> None:
    """
    Verifica se todas as colunas de origem definidas em COLUMN_MAPPING
    existem no DataFrame fornecido.

    Args:
        df: DataFrame a ser validado.

    Raises:
        KeyError: Se uma ou mais colunas estiverem ausentes.
    """
    missing = [col for col in COLUMN_MAPPING if col not in df.columns]
    if missing:
        raise KeyError(
            f"Coluna(s) não encontrada(s) no arquivo Gevis: {missing}\n"
            "Verifique o dicionário COLUMN_MAPPING em src/config.py."
        )
