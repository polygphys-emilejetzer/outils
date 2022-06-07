# -*- coding: utf-8 -*-
"""Exemple de programme de base de données."""

# Bibliothèques standard
import argparse

from pathlib import Path

# Bibliothèques PIPy
from sqlalchemy import MetaData

# Imports relatifs
from . import BaseDeDonnéesConfig, BaseDeDonnées, BaseTableau
from .modeles import créer_dbs

parseur_darguments = argparse.ArgumentParser()
parseur_darguments.add_argument('-f',
                                dest='fichier',
                                type=str,
                                help='fichier à lire',
                                required=False,
                                default='demo.cfg')
arguments = parseur_darguments.parse_args()

# Fichier de configuration
fichier_config = Path(arguments.fichier)
config = BaseDeDonnéesConfig(arguments.fichier)

try:
    md = créer_dbs(MetaData())

    protocole = config['bd']['protocole']
    fichier_bd = config.getpath('bd', 'adresse')

    if not fichier_bd.exists():
        fichier_bd.touch()

    try:
        adresse = f'{protocole}:///{fichier_bd}'

        # Exemple de l'objet BaseDeDonnées
        base = BaseDeDonnées(adresse, md)
        base.réinitialiser()

        # Exemple de l'objet BaseTableau
        for nom_tab in config.getlist('bd', 'tables'):
            tab = BaseTableau(base, nom_tab)
    finally:
        fichier_bd.unlink()
finally:
    fichier_config.unlink()
