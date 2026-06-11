"""
Módulo de configuração de logging para a aplicação.

Futuras melhorias:
- Adicionar handler para gravação em arquivo com rotação (RotatingFileHandler)
- Enviar logs críticos por e-mail ou notificação
"""

import logging
import sys


def get_logger(name: str = "etl-gevis") -> logging.Logger:
    """
    Retorna um logger configurado com formatação padronizada.

    Args:
        name: Nome do logger (padrão: 'etl-gevis').

    Returns:
        Instância de logging.Logger configurada.
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)

    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)

    logger.addHandler(handler)
    return logger
