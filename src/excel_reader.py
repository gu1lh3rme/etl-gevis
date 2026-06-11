"""
Módulo responsável pela leitura do arquivo exportado pelo Gevis.

Utiliza pandas para carregar o arquivo .xlsx e retornar um DataFrame.

Futuras melhorias:
- Suportar arquivos .csv e outros formatos exportados pelo Gevis
- Validar o schema do arquivo de entrada automaticamente
"""

from pathlib import Path

import pandas as pd

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

    df = pd.read_excel(file_path, engine="openpyxl")

    if df.empty:
        raise ValueError(f"O arquivo '{file_path}' está vazio.")

    logger.info("Arquivo Gevis carregado")
    logger.info(f"{len(df)} registros encontrados")

    return df
