# -*- coding: utf-8 -*-
"""Manipulation et affichage de base de données."""

# Bibliothèque standard
import itertools as it
import tkinter as tk

from typing import Callable

# Biblothèque PIPy
import pandas as pd

# Imposts relatifs
from ..base_de_donnees import BaseDeDonnées, BaseTableau
from ..base_de_donnees.dtypes import default
from . import InterfaceHandler


class Tableau(BaseTableau):
    """Encapsulation de InterfaceHandler, avec héritage de BaseTableau."""

    def __init__(self,
                 handler: InterfaceHandler,
                 db: BaseDeDonnées,
                 table: str):
        """
        Encapsule InterfaceHandler & BaseDeDonnées.

        Parameters
        ----------
        handler : InterfaceHandler
            Instance de InterfaceHandler.
        db : BaseDeDonnées
            Base de Données à gérer.
        table : str
            Tableau à gérer.

        Returns
        -------
        None.

        """
        super().__init__(db, table)
        self.widgets = pd.DataFrame()
        self.commandes = []
        self.handler = handler

    def oublie_pas_la_màj(self, f: Callable, *args) -> Callable:
        """
        Force la mise à jour de la grille.

        À utiliser après un changement à la base de données.

        Parameters
        ----------
        f : Callable
            Fonction à envelopper.
        *args : TYPE
            Arguments transmis à f.

        Returns
        -------
        F : Callable
            Fonction enveloppée.

        """

        def F():
            f(*args)
            self.update_grid()

        return F

    def build_commandes(self, rangée: int) -> tuple:
        """
        Construire les widgets de boutons.

        Eg: soummettre des données, effacer, etc.

        Parameters
        ----------
        rangée : int
            Rangée des widgets.

        Returns
        -------
        tuple
            Les widgets.

        """
        a = self.handler.bouton('+', self.oublie_pas_la_màj(self.append))
        b = self.handler.bouton('-', self.oublie_pas_la_màj(self.delete,
                                                            rangée))

        return a, b

    def build(self):
        """
        Construire les widgets.

        Returns
        -------
        None.

        """
        self.widgets = self.df.copy()

        colonnes = filter(lambda x: x != 'index', self.columns)
        colonnes = list(map(self.handler.texte, colonnes))
        self.widgets.columns = colonnes

        index = list(map(self.handler.texte, self.index))
        self.widgets.index = index

        I, C = self.widgets.shape

        for i, c in it.product(range(I), range(C)):
            df = self.iloc()[[i], [c]]
            dtype = self.dtype(self.columns[c])
            _ = self.handler.entrée(df, self.màj, dtype)
            self.widgets.iloc[i, c] = _

        self.commandes = list(map(self.build_commandes, self.index))

    @property
    def rowspan(self):
        """Retourne le nombre de rangées + 1 (pour l'index)."""
        return self.shape[0] + 2

    @property
    def columnspan(self):
        """Retourne le nombre de colonnes + 1 (pour l'en-tête)."""
        return self.shape[1] + 2

    def grid(self, row: int, column: int):
        """
        Display the DataFrame.

        Parameters
        ----------
        row : int
            Rangée initiale.
        column : int
            Colonne initiale.

        Returns
        -------
        None.

        """
        self.__grid_params = {'row': row, 'column': column}
        self.build()

        for i, c in enumerate(self.widgets.columns):
            c.grid(row=row, column=column + i + 3)

        for i, ((idx, rang),
                (plus, moins)) in enumerate(zip(self.widgets.iterrows(),
                                                self.commandes)):
            for k, w in enumerate((plus, moins, idx)):
                w.grid(row=row + i + 1, column=column + k)

            for j, col in enumerate(rang):
                col.grid(row=row + i + 1, column=column + k + j + 1)

    def pack(self, *args, **kargs):
        pass

    @property
    def children(self):
        """
        Retourne tous les widgets de l'affichage.

        Returns
        -------
        itertools.chain
            Itérateur de tous les widgets.

        """
        return it.chain(self.widgets.columns,
                        self.widgets.index,
                        *self.widgets.values,
                        *self.commandes)

    def destroy_children(self):
        """
        Détruit les widgets.

        Returns
        -------
        None.

        """
        for widget in self.children:
            widget.destroy()

    def destroy(self):
        """
        Assure la destruction des enfants avec la notre.

        Returns
        -------
        None.

        """
        self.destroy_children()
        super().destroy()

    def update_grid(self):
        """
        Met l'affichage à jour.

        Returns
        -------
        None.

        """
        self.destroy_children()
        self.grid(**self.__grid_params)


class Formulaire(BaseTableau):
    """Formulaire d'entrée de données."""

    def __init__(self,
                 handler: InterfaceHandler,
                 db: BaseDeDonnées,
                 table: str):
        """
        Crée un formulaire d'entrée de données.

        Parameters
        ----------
        handler : InterfaceHandler
            Gestionnaire d'interface.
        db : BaseDeDonnées
            Base de donnée.
        table : str
            Tableau.

        Returns
        -------
        None.

        """
        super().__init__(db, table)
        self.widgets = pd.DataFrame()
        self.commandes = []
        self.handler = handler

    def oublie_pas_la_màj(self, f: Callable, *args):
        """
        Force la mise à jour de la grille.

        À utiliser après un changement à la base de données.
        """

        def F():
            f(*args)
            self.update_grid()

        return F

    def effacer(self):
        """
        Effacer les champs du formulaire.

        Returns
        -------
        None.

        """
        self.update_grid()

    def soumettre(self):
        """
        Rentre les données dans la base de données.

        Returns
        -------
        None.

        """
        _ = {}
        for c, v in self.widgets.loc[0, :].items():
            if hasattr(v, 'get'):
                _[c.cget('text')] = [v.get()]
            elif isinstance(v, tk.Checkbutton):
                _[c.cget('text')] = [v.instate(['selected'])]

        _ = pd.DataFrame(_)
        self.append(_)
        self.effacer()

    def build_commandes(self) -> tuple:
        """
        Construit les widgets de commandes.

        Eg: boutons.

        Returns
        -------
        tuple
            Boutons créés.

        """
        a = self.handler.bouton('Effacer', self.effacer)
        b = self.handler.bouton('Soumettre', self.soumettre)
        return a, b

    def build(self):
        """
        Construire tous les widgets.

        Returns
        -------
        None.

        """
        self.widgets = pd.DataFrame(None, columns=self.columns, index=[0])

        colonnes = filter(lambda x: x != 'index', self.columns)
        colonnes = list(map(self.handler.texte, colonnes))
        self.widgets.columns = colonnes

        for n, col in zip(self.columns, colonnes):
            dtype = self.dtype(n)

            df = pd.DataFrame(default(dtype),
                              columns=[col],
                              index=[max(self.index, default=0) + 1])

            _ = self.handler.entrée(df, lambda x: None, dtype)
            self.widgets.loc[0, col] = _

        self.commandes = self.build_commandes()

    @ property
    def rowspan(self):
        """Retourne le nombre de rangées + 1 (pour l'index)."""
        return self.shape[1] + 2

    @ property
    def columnspan(self):
        """Retourne le nombre de colonnes + 1 (pour l'en-tête)."""
        return 2

    def grid(self, row: int, column: int):
        """
        Affiche le formulaire.

        Parameters
        ----------
        row : int
            Rangée initiale.
        column : int
            Colonne initiale.

        Returns
        -------
        None.

        """
        self.__grid_params = {'row': row, 'column': column}

        self.build()

        for j, (c, v) in enumerate(zip(self.widgets.columns,
                                       self.widgets.loc[0, :])):
            c.grid(row=row + j, column=column)
            v.grid(row=row + j, column=column + 1)

        for i, c in enumerate(self.commandes):
            c.grid(row=row + j + 1, column=column + i)

    def pack(self, *args, **kargs):
        pass

    @ property
    def children(self):
        """
        Liste les widgets.

        Returns
        -------
        itertools.chain
            Widgets.

        """
        return it.chain(self.widgets.columns,
                        *self.widgets.values,
                        self.commandes)

    def destroy_children(self):
        """
        Détruire les enfants.

        Returns
        -------
        None.

        """
        for widget in self.children:
            widget.destroy()

    def destroy(self):
        """
        Détruire les enfants, puis nous.

        Returns
        -------
        None.

        """
        self.destroy_children()
        super().destroy()

    def update_grid(self):
        """
        Update the grid after a change to the DataFrame.

        Returns
        -------
        None.

        """
        self.destroy_children()
        self.grid(**self.__grid_params)


class Graphe(Tableau):

    def build(self):
        pass

    def grid(self):
        pass

    def pack(self):
        pass

    def update_grid(self):
        pass
