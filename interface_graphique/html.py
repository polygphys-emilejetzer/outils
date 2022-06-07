# -*- coding: utf-8 -*-
"""Éléments d'interface HTML."""

# Bibliothèque standard
from typing import Callable


class HTMLÉlémentInterface:
    """Interface HTML."""

    def __init__(self,
                 master,
                 tag: str,
                 attrs: dict[str, str],
                 contenu: list = None):
        """
        Interface HTML.

        Parameters
        ----------
        master : TYPE
            DESCRIPTION.
        tag : str
            DESCRIPTION.
        attrs : dict[str, str]
            DESCRIPTION.
        contenu : list, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        self.master = master
        self.tag = tag
        self.attrs = attrs
        self.contenu = contenu

    def grid(self, row: int, column: int):
        """
        Afficher.

        Parameters
        ----------
        row : int
            DESCRIPTION.
        column : int
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return str(self)

    def __repr__(self):
        """
        Description.

        Returns
        -------
        str
            DESCRIPTION.

        """
        return f'<Élément {self.tag}>'

    def __str__(self):
        """
        Afficher.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        attributs = ' '.join(f'{a}="{b}"' for a, b in self.attrs.items())
        if self.contenu is None:
            return f'<{self.tag} {attributs} />'
        elif isinstance(self.contenu, list):
            return f'<{self.tag} {attributs}>\n' + '\n'.join(str(e) for e in
                                                             self.contenu) +\
                f'</{self.tag}>'


class HTMLTable(HTMLÉlémentInterface):
    """Tableau HTML."""

    def __init__(self, master=None):
        """
        Tableau HTML.

        Parameters
        ----------
        master : TYPE, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        super().__init__(master, 'table')
        self.grille = [[]]

    def grid(self, row: int, column: int):
        """
        Afficher.

        Parameters
        ----------
        row : int
            DESCRIPTION.
        column : int
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return str(self)


class HTMLCellule(HTMLÉlémentInterface):
    """Cellule de tableau HTML."""

    def __init__(self, master: HTMLTable = None):
        """
        Cellule de tableau HTML.

        Parameters
        ----------
        master : HTMLTable, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        super().__init__(master, 'td')

    def grid(self, row: int, column: int):
        """
        Afficher.

        Parameters
        ----------
        row : int
            DESCRIPTION.
        column : int
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        while row >= len(self.master.grille):
            self.master.grille.append([])
        while column >= len(self.master.grille[row]):
            self.master.grille[row].append(None)
        self.master.grille[row][column] = self

        return super().grid(row, column)


class HTMLEntrée(HTMLCellule):
    """Entrée de données."""

    def __init__(self, master: HTMLTable, texte: str, commande: Callable):
        """
        Entrée de données.

        Parameters
        ----------
        master : HTMLTable
            DESCRIPTION.
        texte : str
            DESCRIPTION.
        commande : Callable
            DESCRIPTION.

        Returns
        -------
        None.

        """
        pass


class HTMLTexte(HTMLCellule):
    """Texte."""

    def __init__(self, master: HTMLTable, texte: str):
        """
        Texte.

        Parameters
        ----------
        master : HTMLTable
            DESCRIPTION.
        texte : str
            DESCRIPTION.

        Returns
        -------
        None.

        """
        pass


class HTMLBouton(HTMLCellule):
    """Bouton."""

    def __init__(self, master: HTMLTable, texte: str, commande: Callable):
        """
        Bouton.

        Parameters
        ----------
        master : HTMLTable
            DESCRIPTION.
        texte : str
            DESCRIPTION.
        commande : Callable
            DESCRIPTION.

        Returns
        -------
        None.

        """
        pass
