# -*- coding: utf-8 -*-
"""Migration et gestion particulière de bases de données."""

# Bibliothèque standard
from typing import Callable

# Bibliothèque PIPy
from sqlalchemy import MetaData

# Improts relatifs
from ..database import BaseDeDonnées


def reset(adresse: str, schema: MetaData):
    """Réinitialiser une base de données."""
    db = BaseDeDonnées(adresse, schema)
    db.réinitialiser()


def init(adresse: str, schema: MetaData):
    """Initialiser une base de données."""
    db = BaseDeDonnées(adresse, schema)

    for t in schema.tables:
        if t not in db.tables:
            pass


def migrer(a: BaseDeDonnées,
           b: BaseDeDonnées,
           clé: dict[str, str],
           conv: dict[str, Callable]):
    """Migrer d'une structure à une autre."""
    pass
