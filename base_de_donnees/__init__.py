# -*- coding: utf-8 -*-
"""Construire une base de donnée selon un fichier de configuration simple."""

# Bibliothèques standards
import pathlib  # Manipulation de chemins

# Description de signatures de fonctions
from typing import Union, Callable, Any
from functools import partial, wraps  # Manipuler des fonctions
from inspect import signature  # Utiliser les signatures de fonctions

# Bibliothèques via PIPy
import sqlalchemy as sqla  # Fonctions et objets de bases de données
import pandas as pd  # Manipulations de données en Python

# Imports relatifs
# Conversion en types internes de différents modules
from ..config import FichierConfig
from .dtypes import get_type, default, column

# Certains types de fichiers, pour deviner quelle fonction de lecture
# utiliser quand on importe un fichier dans une base de données.
TYPES_FICHIERS: dict[str, Callable] = {'.xlsx': pd.read_excel,
                                       '.xls': pd.read_excel,
                                       '.csv': pd.read_csv,
                                       '.pickle': pd.read_pickle,
                                       '.txt': pd.read_table}


class BaseDeDonnéesConfig(FichierConfig):
    """Configuration de base de données."""

    def default(self) -> str:
        """
        Retourne le contenu par défaut de la configuration.

        :return: Contenu par défaut de la configuration.
        :rtype: str

        """
        return (pathlib.Path(__file__).parent / 'default.cfg').open().read()


class BaseDeDonnées:
    """Lien avec une base de données spécifique."""

    def __init__(self, adresse: str, metadata: sqla.MetaData, index_col='index'):
        """
        Lien avec la base de donnée se trouvant à adresse.

        Utilise le schema metadata.

        :param adresse: Adresse vers la base de données.
            Voir https://docs.sqlalchemy.org/en/14/core/engines.html#database-urls
        :type adresse: str
        :param metadata: Structure de la base de données.
            Voir https://docs.sqlalchemy.org/en/14/core/schema.html
        :type metadata: sqla.MetaData
        :return: DESCRIPTION
        :rtype: TYPE

        """
        # Adresse de la base de données
        # https://docs.sqlalchemy.org/en/14/core/engines.html#database-urls
        self.adresse = adresse

        # Structure de la base de données
        # https://docs.sqlalchemy.org/en/14/core/schema.html
        self.metadata = metadata

        self.index_col = index_col

    # Interface de sqlalchemy

    @property
    def tables(self) -> dict[str, sqla.Table]:
        """Liste des tables contenues dans la base de données."""
        res = self.metadata.tables
        return res

    def table(self, table: str) -> sqla.Table:
        """Retourne une table de la base de données."""
        res = self.tables[table]
        return res

    def execute(self, requête, *args, **kargs):
        """Exécute la requête SQL donnée et retourne le résultat."""
        with self.begin() as con:
            res = con.execute(requête, *args, **kargs)
            return res

    def select(self,
               table: str,
               columns: tuple[str] = tuple(),
               where: tuple = tuple(),
               errors: str = 'ignore') -> pd.DataFrame:
        """
        Sélectionne des colonnes et items de la base de données.

        Selon les  critères fournis.

        :param table: Tableau d'où extraire les données.
        :type table: str
        :param columns: Colonnes à extraire. The default is tuple().
            Un tuple vide sélectionne toutes les colonnes,
            defaults to tuple()
        :type columns: tuple[str], optional
        :param where: Critères supplémentaires. The default is tuple(),
            defaults to tuple()
        :type where: tuple, optional
        :param errors: Comportement des erreurs., defaults to 'ignore'
        :type errors: str, optional
        :return: Retourne un DataFrame contenant les items et colonnes
        sélectionnées.
        :rtype: pandas.DataFrame

        """
        # Si aucune colonne n'est spécifiée, on les prends toutes.
        if not len(columns):
            columns = self.columns(table)

        # Si une liste de colonnes est fournie, on vérifie qu'elles sont
        # toutes présentes dans le tableau.
        # On utilise aussi les objets Column du tableau
        columns = [self.table(table).columns[self.index_col]] + list(
            filter(lambda x: x.name in columns, self.table(table).columns))

        requête = sqla.select(columns).select_from(self.table(table))

        for clause in where:
            requête = requête.where(clause)

        with self.begin() as con:
            df = pd.read_sql(requête, con, index_col=self.index_col)

        return df

    def update(self, table: str, values: pd.DataFrame):
        """
        Mets à jour des items déjà présents dans la base de données.

        :param table: Tableau où se trouvent les données.
        :type table: str
        :param values: DataFrame contenant les valeurs à modifier.
        L'index est le critère de sélection.
        :type values: pd.DataFrame
        :return: None
        :rtype: NoneType

        """
        with self.begin() as con:
            for index, rangée in values.iterrows():
                requête = self.table(table).update()\
                                           .where(index == self.table(table).columns[self.index_col])\
                                           .values(rangée)
                con.execute(requête)

    def insert(self, table: str, values: pd.DataFrame):
        """
        Insère un nouvel élément dans la base de données.

        :param table: Tableau utilisé.
        :type table: str
        :param values: Valeurs à insérer.
        :type values: pd.DataFrame
        :return: None
        :rtype: NoneType

        """
        with self.begin() as con:
            values.to_sql(table, con, if_exists='append')

    def append(self, table: str, values: pd.DataFrame):
        """
        Ajoute un item à la fin de la base de données.

        :param table: Table où ajouter les données.
        :type table: str
        :param values: Valuers à ajouter.
        :type values: pd.DataFrame
        :return: None
        :rtype: NoneType

        """
        with self.begin() as con:
            values.to_sql(table, con, if_exists='append', index=False)

    def delete(self, table: str, values: pd.DataFrame):
        """
        Retire une entrée de la base de données.

        :param table: Tableau d'où retirer l'entrée.
        :type table: str
        :param values: Valeurs à retirer.
        :type values: pd.DataFrame
        :return: None
        :rtype: NoneType

        """
        requête = self.table(table).delete()

        # Réparation temporaire
        if isinstance(values, pd.DataFrame):
            index = values.index.name
            idx = values.index
        else:
            index = self.index_col
            idx = pd.Index([values], name=index)
        for i in idx:
            clause = self.table(table).columns[index] == 1
            r = requête.where(clause)
            self.execute(r)

    def màj(self, table: str, values: pd.DataFrame):
        """
        Met à jour des entrées de la base de données.

        Utilise update ou insert selon la préexistence de l'élément.
        L'index de values est utilisé comme critère.

        :param table: Tableau à mettre à jour.
        :type table: str
        :param values: Valeurs à mettre à jour.
        :type values: pd.DataFrame
        :return: None
        :rtype: NoneType

        """
        index = self.index(table)
        existe = values.index.isin(index)

        if existe.any():
            self.update(table, values.loc[existe, :])

        if not existe.all():
            self.insert(table, values.loc[~existe, :])

    def create_engine(self) -> sqla.engine:
        """
        Créer le moteur de base de données.

        :return: Moteur de base de données.
        :rtype: sqlalchemy.engine

        """
        return sqla.create_engine(str(self.adresse), future=True)

    def begin(self):
        """
        Retourne une connection active.

        Eg:
            with instance_BdD.begin() as con:
                ...

        :return: Connection active
        :rtype: Connection SQLAlchemy

        """
        return self.create_engine().begin()

    def initialiser(self, checkfirst: bool = True):
        """
        Créer les tableaux d'une base de données.

        :param checkfirst: Vérfier ou non l'existence des tableaux et champs,
            defaults to True
        :type checkfirst: bool, optional
        :return: None
        :rtype: NoneType

        """
        with self.begin() as con:
            self.metadata.create_all(con, checkfirst=checkfirst)

    def réinitialiser(self, checkfirst: bool = True):
        """
        Effacer puis créer les tableaux d'une base de données.

        :param checkfirst: Vérifier ou non l'existence des tableaux et champs
            , defaults to True
        :type checkfirst: bool, optional
        :return: None
        :rtype: NoneType

        """
        with self.begin() as con:
            self.metadata.drop_all(con, checkfirst=checkfirst)
            self.metadata.create_all(con)

    # Interface ajoutée, pour faciliter les migrations et modifications
    def alter(self,
              table: str,
              champ: str,
              dtype: type = str,
              opération: str = 'add'):
        dtype = str(get_type('python', dtype, 'sqlalchemy'))
        instruction = sqla.text(
            f'alter table {table} {opération} {champ} {dtype}')

        with self.begin() as connexion:
            connexion.execute(instruction)

    # Interface de pandas.DataFrame

    def dtype(self, table: str, champ: str) -> str:
        """
        Retourne le type de données d'un champ dans un tableau.

        :param table: Tableau.
        :type table: str
        :param champ: Champ dont on veut le type.
        :type champ: str
        :return: dtype selon pandas.
        :rtype: str

        """
        type_champ = self.table(table).columns[champ].type
        type_champ: str = get_type('sqlalchemy', type_champ, 'pandas')
        return type_champ

    def dtypes(self, table: str) -> pd.Series:
        """
        Retourne les types des colonnes d'un tableau.

        :param table: Tableau dont on veut les types.
        :type table: str
        :return: Series avec les colonnes comme index, les types comme valeurs.
        :rtype: pandas.Series

        """
        cols = self.columns(table)
        dtypes = map(lambda x: self.dtype(table, x), self.columns(table))
        dtypes = pd.Series(dtypes, index=cols)

        return dtypes

    def columns(self, table: str) -> pd.Index:
        """

        :param table: Retourne un index des colonnes présentes dans le tableau.

        :type table: str
        :return: DESCRIPTION
        :rtype: TYPE

        """
        """
        Retourne un index des colonnes présentes dans le tableau.

        Parameters
        ----------
        table : str
            Tableau dont on veut les colonnes.

        Returns
        -------
        res: pandas.Index
            Index des colonnes du tableau..

        """
        res = pd.Index(c.name for c in self.table(
            table).columns if c.name != self.index_col)

        return res

    def index(self, table: str) -> pd.Index:
        """
        Retourne l'index d'un tableau (colonne `index`).

        :param table: Tableau dont on veut l'index.
        :type table: str
        :return: Index du tableau.
        :rtype: pandas.Index

        """
        requête = sqla.select([self.table(
            table).columns[self.index_col]]).select_from(self.table(table))

        with self.begin() as con:
            résultat = con.execute(requête)
            res = pd.Index(r[self.index_col] for r in résultat)
            return res

    def loc(self,
            table: str,
            columns: tuple[str] = None,
            where: tuple = tuple(),
            errors: str = 'ignore'):
        """
        Retourne un objet de sélection pandas.

        :param table: Tableau à extraire.
        :type table: str
        :param columns: Colonnes à sélectionner, defaults to None
        :type columns: tuple[str], optional
        :param where: Contraintes supplémentaires, defaults to tuple()
        :type where: tuple, optional
        :param errors: Traitement des erreurs, defaults to 'ignore'
        :type errors: str, optional
        :return: Objet de sélection.
        :rtype: pandas.DataFrame.loc

        """
        if columns is None:
            columns = self.columns(table)

        res = self.select(table, columns, where, errors).loc

        return res

    def iloc(self,
             table: str,
             columns: tuple[str] = tuple(),
             where: tuple = tuple(),
             errors: str = 'ignore'):
        """
        Retourne un objet de sélection numérique pandas.

        :param table: Tableau à extraire.
        :type table: str
        :param columns: Colonnes à sélectionner., defaults to tuple()
        :type columns: tuple[str], optional
        :param where: Contraintes supplémentaires, defaults to tuple()
        :type where: tuple, optional
        :param errors: Traitement des erreurs, defaults to 'ignore'
        :type errors: str, optional
        :return: Objet de sélection numérique.
        :rtype: pandas.DataFrame.iloc

        """
        if columns is None:
            columns = self.columns(table)

        res = self.select(table, columns, where, errors).iloc

        return res

    def deviner_type_fichier(self, chemin: pathlib.Path) -> Callable:
        """
        Retourne la fonction pandas à utiliser pour importer un fichier.

        :param chemin: Chemin du fichier qu'on veut importer.
        :type chemin: pathlib.Path
        :return: Fonction du module pandas pour importer un fichier.
        :rtype: Callable

        """
        return TYPES_FICHIERS[chemin.suffix]

    def read_file(self,
                  table: str,
                  chemin: pathlib.Path,
                  type_fichier: Union[str, Callable] = None):
        """
        Importer un fichier dans la base de données.

        :param table: Tableau dans lequel importer les données.
        :type table: str
        :param chemin: Fichier à importer.
        :type chemin: pathlib.Path
        :param type_fichier: Type de fichier.
            Si non spécifié, on devine avec l'extension, defaults to None
        :type type_fichier: Union[str, Callable], optional
        :return: None
        :rtype: NoneType

        """
        if type_fichier is None:
            type_fichier = self.deviner_type_fichier(chemin)
        elif isinstance(type_fichier, str):
            type_fichier = TYPES_FICHIERS[type_fichier]

        df = type_fichier(chemin, index_col=self.index_col)

        self.màj(table, df)


class BaseTableau:
    """Encapsulation de la classe BaseDeDonnées."""

    def __init__(self,
                 db: BaseDeDonnées,
                 table: str,
                 index_col: str = 'index',
                 à_partir_de: pd.DataFrame = None):
        """
        Encapsule de la classe BaseDeDonnées.

        Avec accès seulement à la table table.

        :param db: Une interface à une base de données.
        :type db: BaseDeDonnées
        :param table: Le nom d'un tableau dans db.
        :type table: str
        :return: None
        :rtype: NoneType

        """
        self.nom_table: str = table
        self.index_col = index_col

        # Obtenir la structure déjà présente du tableau
        # Ou utiliser l'objet base de données transmis en argument.
        if isinstance(db, str):
            self.db: BaseDeDonnées = BaseDeDonnées(
                db, sqla.MetaData(), index_col)
            with self.db.begin() as connexion:
                self.db.metadata.reflect(connexion)
        else:
            self.db: BaseDeDonnées = db

        # Forcer la présence de certaines colonnes à l'initialisation
        # Utile pour charger des formulaires qui pourraient changer
        # (eg ajout de champs)
        if à_partir_de is not None:
            for champ, dtype in à_partir_de.dtypes.items():
                if champ not in self.columns:
                    self.alter(champ, get_type('pandas', dtype, 'python'))

            with self.db.begin() as connexion:
                self.db.metadata.reflect(connexion)

    def __getattr__(self, attr: str) -> Any:
        """
        Obtiens un attribut de self.db ou self.df.

        :param attr: Attribut à obtenir.
        :type attr: str
        :return: L'attribut demandé.
        :rtype: Any

        :raises AttributeError: Si l'attribut ne peut pas être trouvé.

        """
        résultat = None
        try:
            # Pour utiliser des colonnes d'index différentes, avec une même base de données.
            self.db.index_col, vieux_index_col = self.index_col, self.db.index_col
            if hasattr(BaseDeDonnées, attr):
                obj = getattr(self.db, attr)
                if isinstance(obj, Callable):
                    sig = signature(obj)

                    if len(sig.parameters) == 1 and 'table' in sig.parameters:
                        résultat = partial(obj, self.nom_table)()
                    elif 'table' in sig.parameters:
                        partielle = partial(obj, self.nom_table)

                        @wraps(partielle)
                        def résultat(*args, **kargs):
                            try:
                                self.db.index_col, vieux_index_col = self.index_col, self.db.index_col
                                res = partielle(*args, **kargs)
                            finally:
                                self.db.index_col = vieux_index_col

                            return res
                    else:
                        résultat = obj
                else:
                    résultat = obj
            elif hasattr(pd.DataFrame, attr):
                résultat = getattr(self.df, attr)
            else:
                msg = f'{self!r} de type {type(self)} n\'a pas d\'attribut {attr}\
    , ni (self.__db: BaseDeDonnées, self.df: pandas.DataFrame).'
                raise AttributeError(msg)
        finally:
            self.db.index_col = vieux_index_col

        return résultat

    @ property
    def df(self) -> pd.DataFrame:
        """Le tableau comme pandas.DataFrame."""
        return self.select()
