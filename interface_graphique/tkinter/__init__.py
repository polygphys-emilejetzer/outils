# -*- coding: utf-8 -*-
"""Interface graphique avec tkinter."""

# Bibliothèque standard
import tkinter as tk

from tkinter import ttk
from tkinter.simpledialog import askstring, askinteger, askfloat
from typing import Callable

# Bibliothèque PIPy
import pandas as pd

# Imports relatifs
from ...base_de_donnees.dtypes import get_type
from .. import InterfaceHandler


def tkHandler(master: tk.Tk, editable: bool = True) -> InterfaceHandler:
    """Retourne une instance InterfaceHandler pour tk."""

    def demander(question: str = '', dtype: type = str):
        """Demander une entrée."""
        if dtype == str:
            return askstring('?', question)
        elif dtype == int:
            return askinteger('?', question)
        elif dtype == float:
            return askfloat('?', question)

    def entrée(value: pd.DataFrame,
               commande: Callable,
               dtype: str = 'object',
               editable: bool = editable,
               référence: bool = False,
               valeurs_référencées: list[str] = tuple()) -> tk.Entry:
        conversion = get_type('pandas', dtype, 'python')
        if dtype == 'boolean':
            val = conversion(value.iloc[0, 0])
        else:
            val = value.iloc[0, 0]

        variable = get_type('pandas', dtype, 'tk')(master, val)

        def F(x, i, m, v=variable):
            res = v.get()
            res = conversion(res)

            arg = pd.DataFrame(res,
                               index=value.index,
                               columns=value.columns,
                               dtype=dtype)

            return commande(arg)

        variable.trace_add('write', F)

        if not editable:
            widget = ttk.Label(master, textvariable=variable)
        elif dtype == 'boolean':
            widget = ttk.Checkbutton(master,
                                     variable=variable)
        elif dtype == 'int64' and référence:
            widget = ttk.OptionMenu(
                master, textvariable=variable, **valeurs_référencées)
        elif dtype in ('int64', 'float64'):
            widget = ttk.Spinbox(master, textvariable=variable)
        elif any(i in variable.get() for i in ('\n', '\r', '\t', '  ')):
            # TODO Permettre d'afficher du texte sur plusieurs lignes
            widget = ttk.Entry(master, textvariable=variable)
        else:
            widget = ttk.Entry(master, textvariable=variable)

        return widget

    def texte(s):
        return ttk.Label(master, text=s)

    def bouton(s, c):
        return ttk.Button(master, text=s, command=c)

    def fenetre():
        return tk.Toplevel(master)

    return InterfaceHandler(entrée,
                            texte,
                            bouton,
                            demander,
                            fenetre,
                            tkHandler)
