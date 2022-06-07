# -*- coding: utf-8 -*-
"""Modèles de base de données."""

# Bibliothèque PIPy
from sqlalchemy import MetaData, Table, Column, ForeignKey

# Imports relatifs
from .dtypes import column


def col_index() -> Column:
    """
    Retourne une colonne d'index.

    :return: Colonne d'index
    :rtype: Column

    """
    return column('index', int, primary_key=True, autoincrement=True)


def personnes(metadata: MetaData) -> Table:
    """
    Retourne le tableau du personnel.

    :param metadata: Schema
    :type metadata: MetaData
    :return: Une table décrivant des personnes.
    :rtype: Table

    """
    cols = [col_index(),  # Index
            column('matricule', str),  # Matricule institutionnel
            column('nom', str),  # Nom
            column('prénom', str),  # Prénom
            column('courriel', str),  # Courriel institutionnel
            column('role', str)  # Rôle comme employé
            ]

    return Table('personnes', metadata, *cols)


def locaux(metadata: MetaData) -> Table:
    """
    Retourne le tableau des locaux.

    :param metadata: Schema
    :type metadata: MetaData
    :return: Une table décrivant des locaux
    :rtype: Table

    """
    # Pour référer à des personnes uniques
    matricule = metadata.tables['personnes'].columns['index']
    cols = [col_index(),  # Index
            column('porte principale', str),  # N  de orte principale du local
            column('responsable', int, ForeignKey(matricule), default=1),
            column('description', str),  # Description du local
            column('utilisation', str)  # Résumé de l'utilisation du local
            ]

    return Table('locaux', metadata, *cols)


def portes(metadata: MetaData) -> Table:
    """
    Retourne le tableau des portes.

    :param metadata: Schema
    :type metadata: MetaData
    :return: Une table décrivant des portes de locaux
    :rtype: Table

    """
    # Identifier un local précis
    local = metadata.tables['locaux'].columns['index']
    cols = [col_index(),  # Index
            column('numéro', str),  # N  de porte
            column('local', int, ForeignKey(local), default=1)
            ]

    return Table('portes', metadata, *cols)


def etageres(metadata: MetaData) -> Table:
    """
    Retourne le tableau des étagères.

    :param metadata: Schema
    :type metadata: MetaData
    :return: Une table décrivant des étagères.
    :rtype: Table

    """
    # Identifier le local précis
    local = metadata.tables['locaux'].columns['index']

    # Identifier la personne responsable
    matricule = metadata.tables['personnes'].columns['index']

    cols = [col_index(),  # Index
            column('local', int, ForeignKey(local), default=1),
            column('responsable', int, ForeignKey(matricule), default=1),
            column('numéro', str),  # Numéro d'étagère dans la pièce
            column('tablette', str),  # N  de tablette
            column('sous-division', str),  # Au besoin
            column('designation', str),  # Nom court
            column('description', str)  # Description plus détaillée
            ]

    return Table('etageres', metadata, *cols)


def créer_dbs(metadata: MetaData) -> MetaData:
    """
    Créer les bases de données.

    :param metadata: Schema
    :type metadata: MetaData
    :return: Le schémas à jour.
    :rtype: MetaData

    """
    personnes(metadata)
    locaux(metadata)
    portes(metadata)
    etageres(metadata)

    return metadata
