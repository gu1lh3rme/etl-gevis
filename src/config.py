"""
Módulo de configuração central da aplicação.

Centraliza mapeamentos de colunas e valores fixos usados no processo ETL.
Lê variáveis de ambiente do arquivo .env via python-dotenv.

Futuras melhorias:
- Carregar o mapeamento de um arquivo JSON/YAML externo para facilitar configuração sem editar código
- Suportar múltiplos perfis de mapeamento (ex.: diferentes versões do Gevis)
"""

import os
from pathlib import Path

from dotenv import load_dotenv

# Carrega variáveis do arquivo .env (se existir)
load_dotenv()

# ---------------------------------------------------------------------------
# Caminhos
# ---------------------------------------------------------------------------

BASE_DIR: Path = Path(__file__).resolve().parent.parent

GEVIS_INPUT_PATH: Path = Path(
    os.getenv("GEVIS_INPUT_PATH", str(BASE_DIR / "data" / "input" / "gevis.xlsx"))
)

VISTORIAS_TEMPLATE_PATH: Path = Path(
    os.getenv(
        "VISTORIAS_TEMPLATE_PATH",
        str(BASE_DIR / "templates" / "VISTORIAS.xlsx"),
    )
)

OUTPUT_PATH: Path = Path(
    os.getenv(
        "OUTPUT_PATH",
        str(BASE_DIR / "data" / "output" / "VISTORIAS_ATUALIZADA.xlsx"),
    )
)

# ---------------------------------------------------------------------------
# Mapeamento de colunas
# Chave   = nome da coluna no arquivo exportado do Gevis
# Valor   = nome da coluna correspondente na planilha VISTORIAS
#
# Ajuste este dicionário caso o layout do Gevis seja diferente.
# ---------------------------------------------------------------------------

COLUMN_MAPPING: dict[str, str] = {
    "Data": "DATA",
    "Placa": "PLACA/CHASSI",
    "Marca": "VEICULO",
    "Nome Contratante OS": "CLIENTE",
    "Valor": "VLR RECEBIDO",
}

# Coluna usada como chave única para evitar duplicatas
UNIQUE_KEY_COLUMN: str = "PLACA/CHASSI"

# ---------------------------------------------------------------------------
# Valores fixos inseridos automaticamente em cada novo registro
# ---------------------------------------------------------------------------

FIXED_VALUES: dict[str, str] = {
    "TIPO DE VISTORIA": "NOVA",
    "FORMA DE PGTO": "PIX",
    "RESP. PELO PGTO": "CLIENTE",
    "STATUS DO PGTO": "PAGO",
}
