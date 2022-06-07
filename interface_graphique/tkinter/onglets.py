# -*- coding: utf-8 -*-
"""Afficher différentes bases de données dans différents onglets."""

# Bibliothèque standard
import pathlib
import logging

import tkinter as tk

from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from pathlib import Path

# Bibliothèque PIPy
import sqlalchemy as sqla

# Imports relatifs
from . import tkHandler
from ..tableau import Tableau, Formulaire
from ...base_de_donnees import BaseDeDonnées
from ...config import FichierConfig


class OngletConfig(tk.Frame):
    """Onglet de configuration."""

    def __init__(self, master: tk.Frame, config: FichierConfig):
        """
        Crée un onglet de configuration.

        Parameters
        ----------
        master : tk.Frame
            Maître dans tk.
        config : FichierConfig
            Configuration.

        Returns
        -------
        None.

        """
        self.config = config

        super().__init__(master)
        self.build()

    def __repr__(self):
        """Affiche les informations de base sur l'objet."""
        return f'<{type(self)} at {hex(id(self))}>'

    @property
    def chemin(self) -> pathlib.Path:
        """
        Retourne le chemin vers le fichier de configuration.

        Returns
        -------
        pathlib.Path
            Chemin vers le fichier de configuration.

        """
        return self.config.chemin

    def build_champ(self, sec, champ, valeur):
        """
        Construire un champ.

        Parameters
        ----------
        sec : TYPE
            DESCRIPTION.
        champ : TYPE
            DESCRIPTION.
        valeur : TYPE
            DESCRIPTION.

        Returns
        -------
        champ_entrée : TYPE
            DESCRIPTION.
        valeur_entrée : TYPE
            DESCRIPTION.
        boutons : TYPE
            DESCRIPTION.

        """
        champ_var = tk.StringVar(self, value=champ)
        valeur_var = tk.StringVar(self, value=valeur)

        champ_var.trace_add(
            'write',
            lambda x, i, m, v=champ_var: self.update_config())
        valeur_var.trace_add(
            'write',
            lambda x, i, m, v=valeur_var: self.update_config())

        champ_entrée = ttk.Entry(self, textvariable=champ_var)
        valeur_entrée = ttk.Entry(self, textvariable=valeur_var)

        boutons = (ttk.Button(self, text='+',
                              command=lambda: self.ajouter_champ(sec)),
                   ttk.Button(self, text='-',
                              command=lambda: self.retirer_champ(sec, champ)))

        return champ_entrée, valeur_entrée, boutons

    def retirer_champ(self, sec, champ):
        """
        Retirer un champ.

        Parameters
        ----------
        sec : TYPE
            DESCRIPTION.
        champ : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        self.champs[sec][champ].destroy()
        self.valeurs[sec][champ].destroy()
        self.boutons[sec][1][champ][0].destroy()
        self.boutons[sec][1][champ][1].destroy()
        del self.champs[sec][champ]
        del self.valeurs[sec][champ]
        del self.boutons[sec][1][champ]
        self.update_config()
        self.update_grid()

    def ajouter_champ(self, sec, champ='Nouveau champ', valeur=None):
        """
        Ajouter un champ.

        Parameters
        ----------
        sec : TYPE
            DESCRIPTION.
        champ : TYPE, optional
            DESCRIPTION. The default is 'Nouveau champ'.
        valeur : TYPE, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        c, v, b = self.build_champ(sec, champ, valeur)
        self.champs[sec][champ] = c
        self.valeurs[sec][champ] = v
        self.boutons[sec][1][champ] = b
        self.update_config()
        self.update_grid()

    def build_section(self, sec=None):
        """
        Construire une section.

        Parameters
        ----------
        sec : TYPE, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        titre : TYPE
            DESCRIPTION.
        champs : TYPE
            DESCRIPTION.
        valeurs : TYPE
            DESCRIPTION.
        boutons : TYPE
            DESCRIPTION.
        bouton : TYPE
            DESCRIPTION.

        """
        section = self.config[sec]

        titre_var = tk.StringVar(self, value=sec)
        titre_var.trace_add(
            'write',
            lambda x, i, m, v=titre_var: self.update_config())
        titre = ttk.Entry(self, textvariable=titre_var)

        bouton = ttk.Button(self, text='-', command=lambda: 1)

        champs, valeurs, boutons = {}, {}, {}
        for champ, valeur in section.items():
            c, v, b = self.build_champ(sec, champ, valeur)
            champs[champ] = c
            valeurs[champ] = v
            boutons[champ] = b

        return titre, champs, valeurs, boutons, bouton

    def retirer_section(self, sec):
        """
        Retirer une section.

        Parameters
        ----------
        sec : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        del self.titres[sec]
        del self.champs[sec]
        del self.valeurs[sec]
        del self.boutons[sec]
        self.update_config()
        self.update_grid()

    def build(self):
        """
        Construire les widgets.

        Returns
        -------
        None.

        """
        self.titres, self.champs, self.valeurs, self.boutons = {}, {}, {}, {}

        for sec in self.config.sections():
            t, cs, vs, bs, b = self.build_section(sec)
            self.titres[sec] = t
            self.champs[sec] = cs
            self.valeurs[sec] = vs
            self.boutons[sec] = [b, bs]

    def update_config(self):
        """
        Mettre la configuration à jour.

        Returns
        -------
        None.

        """
        # Effacer les sections non présentes
        for sec in self.config.sections():
            if sec not in map(lambda x: x.get(), self.titres.values()):
                self.config.remove_section(sec)

        # Vérifier que les sections présentes existent
        for sec in map(lambda x: x.get(), self.titres.values()):
            if sec not in self.config.sections():
                self.config.add_section(sec)

        # Pour chaque section présente
        for sec in map(lambda x: x.get(), self.titres.values()):
            # effacer les champs non-existants
            for champ in map(lambda x: x.get(), self.champs[sec].values()):
                if champ not in self.config.options(sec):
                    self.config.set(sec, champ, '')

        # vérifier les valeurs des champs
        for section in self.champs:
            for clé in list(self.champs[section].keys()):
                nouvelle_clé = self.champs[section][clé].get()
                valeur = self.valeurs[section][clé].get()
                self.config[section][nouvelle_clé] = valeur

                self.champs[section][nouvelle_clé] = self.champs[section][clé]
                self.valeurs[section][nouvelle_clé] = self.valeurs[section][clé]

        self.config.write()

    def subgrid(self):
        """
        Affichage des widgets.

        Returns
        -------
        None.

        """
        self.build()

        colonne = 0
        for titre, étiquette in self.titres.items():
            étiquette.grid(row=0, column=colonne,
                           columnspan=3, sticky=tk.W + tk.E)
            self.boutons[titre][0].grid(
                row=0, column=colonne + 3)

            rangée = 1
            for étiquette, entrée in self.champs[titre].items():
                entrée.grid(row=rangée, column=colonne)
                self.valeurs[titre][étiquette].grid(
                    row=rangée, column=colonne + 1)
                self.boutons[titre][1][étiquette][0].grid(
                    row=rangée, column=colonne + 2)
                self.boutons[titre][1][étiquette][1].grid(
                    row=rangée, column=colonne + 3)
                rangée += 1

            colonne += 4

    def grid(self, *args, **kargs):
        """
        Affichage de l'onglet.

        Parameters
        ----------
        *args : TYPE
            Arguments transmis à tk.Frame.grid.
        **kargs : TYPE
            Arguments transmis à tk.Frame.grid..

        Returns
        -------
        None.

        """
        self.subgrid()
        super().grid(*args, **kargs)

    def update_grid(self):
        """
        Mettre l'affichage à jour.

        Returns
        -------
        None.

        """
        self.destroy_children()
        self.subgrid()

    def destroy_children(self):
        """
        Effacer l'affichage.

        Returns
        -------
        None.

        """
        del self.titres
        del self.champs
        del self.valeurs
        del self.boutons


class OngletBaseDeDonnées(tk.Frame):
    """Onglet de base de données (affichage tableur)."""

    def __init__(self,
                 master: tk.Tk,
                 db: BaseDeDonnées,
                 table: str,
                 *args,
                 config: FichierConfig = None,
                 **kargs):
        """
        Crée un onglet de base de données.

        Parameters
        ----------
        master : tk.Tk
            Maître tk pour l'affichage.
        db : BaseDeDonnées
            Base de données à afficher.
        table : str
            Tableau à afficher.
        *args : TYPE
            Arguments transmis au parent tk.Frame.
        config : FichierConfig, optional
            Configuration externe. The default is None.
        **kargs : TYPE
            Arguments transmis au parent tk.Frame.

        Returns
        -------
        None.

        """
        self.config = config
        self.table = table
        self.db = db

        super().__init__(master, *args, **kargs)
        self.build()

    def __repr__(self):
        """Affiche les informations de base sur l'objet."""
        return f'<{type(self)} at {hex(id(self))}>'

    @property
    def adresse(self):
        """Adresse de la base de données."""
        res = self.config.get('bd', 'adresse', fallback='test.db')
        return res

    def importer(self):
        """
        Importer les données.

        Returns
        -------
        None.

        """
        chemin = Path(askopenfilename())
        self.tableau.read_file(chemin)

    def exporter(self):
        """
        Exporter les données.

        Returns
        -------
        None.

        """
        chemin = asksaveasfilename()
        self.tableau.to_excel(chemin, self.table)

    def exporter_modèle(self):
        """
        Exporter un modèle pour l'entrée de données.

        Returns
        -------
        None.

        """
        chemin = asksaveasfilename()
        self.tableau.loc()[[], :].to_excel(chemin, self.table)

    def build(self):
        """Construit les widgets."""
        self.canevas = tk.Canvas(self, width='50c', height='15c')
        défiler_horizontalement = tk.Scrollbar(
            self, orient='horizontal', command=self.canevas.xview)
        défiler_verticalement = tk.Scrollbar(
            self, orient='vertical', command=self.canevas.yview)
        self.canevas.configure(xscrollcommand=défiler_horizontalement.set,
                               yscrollcommand=défiler_verticalement.set)

        self.contenant = tk.Frame(self.canevas)

        self.contenant.bind('<Configure>', lambda x: self.canevas.configure(
            scrollregion=self.canevas.bbox('all')))

        self.tableau = Tableau(tkHandler(self.contenant), self.db, self.table)

        màj = tk.Button(self, text='Màj',
                        command=lambda: self.tableau.update_grid())

        importer = tk.Button(self, text='Importer',
                             command=self.importer)

        exporter = tk.Button(self, text='Exporter',
                             command=self.exporter)

        modèle = tk.Button(self, text='Modèle',
                           command=self.exporter_modèle)

        self.défiler = [défiler_horizontalement, défiler_verticalement]

        self.boutons = [màj, importer, exporter, modèle]

    def subgrid(self):
        """Afficher les widgets."""
        self.défiler[0].grid(row=16, column=1, columnspan=1, sticky='we')
        self.défiler[1].grid(row=1, column=2, rowspan=15, sticky='ns')
        self.canevas.grid(row=1, column=1, rowspan=15, sticky='news')
        self.canevas.create_window((30, 15), window=self.contenant)
        self.tableau.grid(0, 0)

        for i, b in enumerate(self.boutons):
            b.grid(row=i, column=0)

    def grid(self, *args, **kargs):
        """Afficher le tableau."""

        self.subgrid()
        super().grid(*args, **kargs)


class OngletFormulaire(tk.Frame):
    """Afficher un formulaire d'entrée de données."""

    def __init__(self,
                 master: tk.Tk,
                 db: BaseDeDonnées,
                 table: str,
                 *args,
                 config: FichierConfig = None,
                 **kargs):
        """
        Crée un formulaire d'entrée de données.

        Parameters
        ----------
        master : tk.Tk
            Maître d'interface tk.
        db : BaseDeDonnées
            Base de données.
        table : str
            Tableau où on veut entrer des données.
        *args : TYPE
            Arguments transmis au parent tk.Frame.
        config : FichierConfig, optional
            Fichier de configuration externe. The default is None.
        **kargs : TYPE
            Arguments transmis au parent tk.Frame.

        Returns
        -------
        None.

        """
        self.config = config
        self.table = table
        self.db = db

        super().__init__(master, *args, **kargs)
        self.build()

    def __repr__(self):
        """Affiche les informations de base sur l'objet."""
        return f'<{type(self)} at {hex(id(self))}>'

    @property
    def adresse(self):
        """Adresse de la base de données."""
        res = self.config.get('bd', 'adresse', fallback='test.db')
        return res

    def build(self):
        """
        Construire les widgets.

        Returns
        -------
        None.

        """
        self.canevas = tk.Canvas(self, width='50c', height='15c')

        défiler_horizontalement = tk.Scrollbar(
            self, orient='horizontal', command=self.canevas.xview)

        défiler_verticalement = tk.Scrollbar(
            self, orient='vertical', command=self.canevas.yview)

        self.canevas.configure(xscrollcommand=défiler_horizontalement.set,
                               yscrollcommand=défiler_verticalement.set)

        self.contenant = tk.Frame(self.canevas)
        self.contenant.bind('<Configure>', lambda x: self.canevas.configure(
            scrollregion=self.canevas.bbox('all')))

        self.formulaire = Formulaire(
            tkHandler(self.contenant), self.db, self.table)

        self.défiler = [défiler_horizontalement, défiler_verticalement]

    def subgrid(self):
        """Affiche les widgets."""
        self.défiler[0].grid(row=16, column=1, columnspan=1, sticky='we')
        self.défiler[1].grid(row=1, column=2, rowspan=15, sticky='ns')
        self.canevas.grid(row=1, column=1, rowspan=15, sticky='news')
        self.canevas.create_window((30, 15), window=self.contenant)
        self.formulaire.grid(0, 0)

    def grid(self, *args, **kargs):
        """Affiche le formulaire."""
        self.subgrid()
        super().grid(*args, **kargs)


class Onglets(ttk.Notebook):
    """Groupe d'onglets."""

    def __init__(self,
                 master: tk.Frame,
                 config: FichierConfig,
                 schema: sqla.MetaData,
                 dialect: str = 'sqlite'):
        """
        Crée un groupe d'onglets.

        Parameters
        ----------
        master : tkinter.Frame
            Maître dans l'interface tkinter.
        config : FichierConfig
            Configuration externe.
        schema : sqlalchemy.MetaData
            Structure de base de données.

        Returns
        -------
        None.

        """

        super().__init__(master)
        self.onglets = []

        onglet = OngletConfig(self, config)
        self.add(onglet, text=Path(onglet.chemin).name)

        db = BaseDeDonnées(config.get('bd', 'adresse'), schema)

        tables = config.getlist('bd', 'tables')
        logging.debug('tables = %r', tables)
        for nom_table in tables:
            onglet = OngletBaseDeDonnées(
                self, db, nom_table, config=config)
            self.add(onglet, text=nom_table)

        formulaires = config.getlist('bd', 'formulaires')
        for nom_formulaire in formulaires:
            onglet = OngletFormulaire(self, db, nom_formulaire)
            self.add(onglet, text=f'[F] {nom_formulaire}')

    def __repr__(self):
        """Affiche les informations de base sur l'objet."""
        return f'<{type(self)} at {hex(id(self))}>'

    def add(self, obj: tk.Frame, *args, **kargs):
        """
        Ajouter un onglet.

        Parameters
        ----------
        obj : tk.Frame
            Onglet à ajouter.
        *args : TYPE
            Arguments transmis à la méthode add du parent tk.Frame.
        **kargs : TYPE
            Arguments transmis à la méthode add du parent tk.Frame.

        Returns
        -------
        None.

        """
        self.onglets.append(obj)
        super().add(obj, *args, **kargs)

    def grid(self, *args, **kargs):
        """
        Afficher les onglets.

        Parameters
        ----------
        *args : TYPE
            Arguments transmis à la méthode grid du parent tk.Frame.
        **kargs : TYPE
            Arguments transmis à la méthode grid du parent tk.Frame.

        Returns
        -------
        None.

        """
        for onglet in self.children.values():
            logging.debug('onglet = %r', onglet)
            onglet.subgrid()

        super().grid(*args, **kargs)
