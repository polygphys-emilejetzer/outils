# -*- coding: utf-8 -*-
"""Description et manipulation de types de données."""

# Bibliothèque standard
import pathlib
import datetime
import logging

import tkinter as tk

from typing import Union, Any

# Bibliothèque PIPy
import sqlalchemy as sqla

# Correspondances de types entre différents modules standards.
# Pour tkinter, voir la section «variables» de
# https://tkdocs.com/pyref/index.html
# Pour pandas, voir
# https://pandas.pydata.org/pandas-docs/stable/reference/arrays.html
# Pour SQLAlchemy voir
# https://docs.sqlalchemy.org/en/14/core/type_basics.html
TYPES: tuple[dict[str, Union[str, type]]] = (
    {  # Type générique
        'config': None,
        'python': str,
        'pandas': 'object',
        'sqlalchemy': sqla.PickleType(),
        'tk': tk.StringVar
    },
    {  # Dates
        'config': 'datetime.date',
        'python': datetime.date,
        'pandas': 'datetime64[D]',
        'sqlalchemy': sqla.Date(),
        'tk': tk.StringVar
    },
    {  # Date & heure
        'config': 'datetime.datetime',
        'python': datetime.datetime,
        'pandas': 'datetime64[ns]',
        'sqlalchemy': sqla.DateTime(),
        'tk': tk.StringVar
    },
    {  # Heure
        'config': 'datetime.time',
        'python': datetime.time,
        'pandas': 'datetime64[ns]',
        'sqlalchemy': sqla.Time(),
        'tk': tk.StringVar
    },
    {  # Différence de temps
        'config': 'datetime.timedelta',
        'python': datetime.timedelta,
        'pandas': 'period[ns]',
        'sqlalchemy': sqla.Interval(),
        'tk': tk.StringVar
    },
    {  # Chaîne de caractères
        'config': 'str',
        'python': str,
        'pandas': 'string',
        'sqlalchemy': sqla.UnicodeText(),
        'tk': tk.StringVar
    },
    {  # Nombres entiers
        'config': 'int',
        'python': int,
        'pandas': 'int64',
        'sqlalchemy': sqla.BigInteger(),
        'tk': tk.IntVar
    },
    {  # Nombres à virgule flottante
        'config': 'float',
        'python': float,
        'pandas': 'float64',
        'sqlalchemy': sqla.Float(),
        'tk': tk.DoubleVar
    },
    {  # Valeurs booléennes
        'config': 'bool',
        'python': bool,
        'pandas': 'boolean',
        'sqlalchemy': sqla.Boolean(),
        'tk': tk.BooleanVar
    },
    {  # Chemins de fichiers
        'config': 'pathlib.Path',
        'python': pathlib.Path,
        'pandas': 'object',
        'sqlalchemy': sqla.PickleType(),
        'tk': tk.StringVar
    }
)


def get_type(de: str, t: Union[Any, type, str], à: str) -> Union[type, str]:
    """
    Retourne un type ou description de type dans le bon format.

    :param de: Format de départ.
    :type de: str
    :param t: Type ou description de type dans le format de départ.
    :type t: Union[Any, type, str]
    :param à: Format final.
    :type à: str
    :return: Type ou description de type.
    :rtype: Union[type, str]

    """

    def comp(x):
        return x[de] == t

    for s in filter(comp, TYPES):
        logging.debug('s[à] = %r', s[à])
        return s[à]

    return next(filter(lambda x: x['config'] is None, TYPES))[à]


def default(dtype: str) -> Any:
    """
    Retourne la valeur par défaut pour un type.

    :param dtype: Type de données Pandas.
    :type dtype: str
    :return: Valeur par défaut du type.
    :rtype: Any

    """
    if 'period' in dtype:
        return datetime.timedelta(0)
    elif 'date' in dtype or 'time' in dtype:
        return datetime.datetime.now()
    else:
        return get_type('pandas', dtype, 'python')()


def column(name: str, dtype: type = str, *args, **kargs) -> sqla.Column:
    """
    Retourne une description de colonne du bon type et nom.

    :param name: Nom de la colonne.
    :type name: str
    :param dtype: Type de la colonne, defaults to str
    :type dtype: type, optional
    :param *args: Arguments supplémentaires transmis au constructeur de colonne.
    :param **kargs: Arguments supplémentaires transmis au constructeur de colonne.
    :return: Description de colonne.
    :rtype: sqlalchemy.Column

    """
    def_val = default(get_type('python', dtype, 'pandas'))
    dtype = get_type('python', dtype, 'sqlalchemy')

    if 'default' not in kargs:
        kargs['default'] = def_val

    return sqla.Column(name, dtype, *args, **kargs)
