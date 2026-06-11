"""
Ponto de entrada da aplicação ETL Gevis → VISTORIAS.

Orquestra as etapas:
  1. Leitura do arquivo gevis.xlsx
  2. Leitura da planilha modelo VISTORIAS.xlsx
  3. Mapeamento das colunas
  4. Preenchimento dos valores fixos
  5. Deduplicação por PLACA/CHASSI
  6. Inserção dos registros novos
  7. Salvamento do resultado

Futuras melhorias:
- Interface gráfica com Tkinter para arrastar e soltar arquivos
- Integração direta com a API do Gevis para eliminar a etapa de exportação manual
- Agendamento automático via APScheduler ou cron
- Exportação dos dados para Power BI via conector REST
"""

import sys

from config import GEVIS_INPUT_PATH
from excel_reader import read_gevis
from excel_writer import write_to_vistorias
from logger import get_logger
from mapper import map_gevis_to_vistorias

logger = get_logger()


def run() -> None:
    """Executa o fluxo completo de ETL."""
    try:
        # Etapa 1 – Leitura do arquivo Gevis
        df_gevis = read_gevis(GEVIS_INPUT_PATH)

        # Etapas 3 e 4 – Mapeamento de colunas e valores fixos
        df_mapped = map_gevis_to_vistorias(df_gevis)

        # Etapas 2, 5, 6 e 7 – Escrita na planilha VISTORIAS
        write_to_vistorias(df_mapped)

    except FileNotFoundError as exc:
        logger.error(str(exc))
        sys.exit(1)
    except KeyError as exc:
        logger.error(str(exc))
        sys.exit(1)
    except ValueError as exc:
        logger.error(str(exc))
        sys.exit(1)


if __name__ == "__main__":
    run()
