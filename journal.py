# -*- coding: utf-8 -*-
"""
Journalisation synchrone avec différents modules.

- le module logging
- un répertoire git
- une base de données.
"""

# Bibliothèque standard
from pathlib import Path
from logging import Handler, LogRecord
from subprocess import run
from dataclasses import dataclass
from functools import wraps
from typing import Callable

# Bibliothèque PIPy
import pandas as pd

# Imports relatifs
from .base_de_donnees import BaseTableau, BaseDeDonnées


@dataclass
class Formats:
    """Chaîne de format pour la journalisation."""

    default: str = '[%(asctime)s]\t%(levelname)s\t%(name)s\t%(message)s'
    labo: str = '[%(asctime)s]\t%(message)s'
    détails: str = '[%(asctime)s]\t%(levelname)s\t%(name)s\n\tFichier: \
%(filename)s\n\tFonction: %(funcName)s\n\tLigne: %(lineno)s\n\n\t%(message)s'


class Repository:
    """Répertoire git."""

    def __init__(self, path: Path):
        """
        Répertoire git.

        Parameters
        ----------
        path : Path
            DESCRIPTION.

        Returns
        -------
        None.

        """
        self.path = path

    def init(self):
        """
        Initialiser un répertoire.

        Returns
        -------
        None.

        """
        run(['git', 'init'], cwd=self.path)

    def clone(self, other: str):
        """
        Cloner un répertoire.

        Parameters
        ----------
        other : str
            DESCRIPTION.

        Returns
        -------
        None.

        """
        run(['git', 'clone', other], cwd=self.path)

    def add(self, *args):
        """
        Ajouter un fichier à commettre.

        Parameters
        ----------
        *args : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        run(['git', 'add'] + list(args), cwd=self.path)

    def rm(self, *args):
        """
        Retirer un fichier.

        Parameters
        ----------
        *args : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        run(['git', 'rm'] + list(args), cwd=self.path)

    def commit(self, msg: str, *args):
        """
        Commettre les changements.

        Parameters
        ----------
        msg : str
            DESCRIPTION.
        *args : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        run(['git', 'commit', '-m', msg] + list(args), cwd=self.path)

    def pull(self):
        """
        Télécharger les changements lointains.

        Returns
        -------
        None.

        """
        run(['git', 'pull'], cwd=self.path)

    def push(self):
        """
        Pousser les changements locaux.

        Returns
        -------
        None.

        """
        run(['git', 'push'], cwd=self.path)

    def status(self):
        """
        Évaluer l'état du répertoire.

        Returns
        -------
        None.

        """
        run(['git', 'status'], cwd=self.path)

    def log(self):
        """
        Afficher l'historique.

        Returns
        -------
        None.

        """
        run(['git', 'log'], cwd=self.path)

    def branch(self, b: str = ''):
        """
        Passer à une nouvelle branche.

        Parameters
        ----------
        b : str, optional
            DESCRIPTION. The default is ''.

        Returns
        -------
        None.

        """
        run(['git', 'branch', b], cwd=self.path)

    @property
    def head(self):
        head = (Path(self.path) / '.git' / 'head').open('r', encoding='utf-8')\
            .read()\
            .split(':', 1)[1]\
            .strip()
        commit = (Path(self.path) / '.git' / Path(head)).open('r', encoding='utf-8')\
            .read()\
            .strip()
        return commit


class JournalBD:

    def __init__(self,
                 db: str,
                 table: str):
        self._db = db
        self._table = table
        self._index_col = 'index'
        self._à_partir_de = pd.DataFrame({'créé': [],
                                         'niveau': [],
                                          'logger': [],
                                          'msg': [],
                                          'head': []})

    def __getattr__(self, attr):
        # Ceci devrait faire en sorte que le module fonctionne
        # avec différentes threads.
        # NE FONCTIONNE PAS
        # Voir engine.dispose()?
        tableau = BaseTableau(self._db,
                              self._table,
                              self._index_col,
                              self._à_partir_de)

        résultat = getattr(tableau, attr)
        if isinstance(résultat, Callable):

            @wraps(résultat)
            def f(*args, **kargs):
                rés = résultat(*args, **kargs)
                return rés

            return f
        else:
            del tableau
            return résultat


class Journal(Handler):
    """Journal compatible avec le module logging.

    Maintiens une base de données des changements apportés,
    par un programme ou manuellement. Les changements sont
    aussi sauvegardés dans un répertoire git.
    """

    def __init__(self,
                 level: float,
                 dossier: Path,
                 tableau: JournalBD):
        """Journal compatible avec le module logging.

        Maintiens une base de données des changements apportés,
        par un programme ou manuellement. Les changements sont
        aussi sauvegardés dans un répertoire git.

        Parameters
        ----------
        level : float
            Niveau des messages envoyés.
        dossier : Path
            Chemin vers le répertoire git.
        tableau : BaseTableau
            Objet de base de données.

        Returns
        -------
        None.

        """
        self.repo: Repository = Repository(dossier)
        self.tableau: JournalBD = tableau

        super().__init__(level)

    @property
    def fichier(self):
        """Fichier de base de données (pour SQLite)."""
        return self.tableau.adresse

    # Interface avec le répertoire git

    def init(self):
        """Initialise le répertoire git et la base de données."""
        self.repo.init()
        # self.tableau.initialiser()

    # Fonctions de logging.Handler

    def flush(self):
        """Ne fais rien."""
        pass

    def emit(self, record: LogRecord):
        """
        Enregistre une nouvelle entrée.

        Cette méthode ne devrait pas être appelée directement.

        Parameters
        ----------
        record : LogRecord
            L'entrée à enregistrer.

        Returns
        -------
        None.

        """
        msg = record.getMessage()

        message = pd.DataFrame({'créé': [record.created],
                                'niveau': [record.levelno],
                                'logger': [record.name],
                                'msg': [msg],
                                'head': [self.repo.head]})

        # self.tableau.append(message)
        csv = Path(self.repo.path) / 'résumé.csv'
        if csv.exists():
            en_têtes = False
        else:
            en_têtes = True
            csv.touch()

        message.to_csv(csv, mode='a', header=en_têtes, index=False)
        self.repo.add(csv.name)

        self.repo.commit(msg, '-a')

# TODO Modèle de base de données pour journal
