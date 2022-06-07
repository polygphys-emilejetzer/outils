# -*- coding: utf-8 -*-
"""Démonstration de l'utilisation des classes de fichiers de configuration."""

# Bibliothèque standard
import argparse

from pathlib import Path

# Imports relatifs
from . import FichierConfig

# Configuration pour l'application de ligne de commande
parseur_darguments = argparse.ArgumentParser('Démonstration de FichierConfig.')
parseur_darguments.add_argument('-f',
                                dest='fichier',
                                type=str,
                                help='fichier à lire',
                                required=False,
                                default='demo.cfg')
arguments = parseur_darguments.parse_args()

# Ouvrir un fichier de configuration
fichier = Path(arguments.fichier)
try:
    config = FichierConfig(fichier)

    print(f'Fichier de configuration {fichier!r}')
    print(config)

    # Modifier le contenu du fichier
    config.add_section('test')
    config['test']['sous-test'] = 'valeur test'
    print(config)
finally:
    # Détruire le fichier de démonstration
    fichier.unlink()
