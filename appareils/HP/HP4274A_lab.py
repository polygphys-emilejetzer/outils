#!/usr/bin/env python3.10
# -*- coding: utf-8 -*-
"""
Programme d'acquisition de données.

Created on Thu Jan 20 09:29:12 2022

@author: Jeremy Lafortune, Émile Jetzer
"""

import decimal  # Manipulation et mise en forme de nombres
import logging  # Noter les résultats des commandes
import time
import tkinter as tk  # Interface graphique
import traceback
import datetime

from pathlib import Path  # Manipulation facile de fichiers
from tkinter import ttk  # Interface graphique avec thème
from tkinter import messagebox as mb  # Alertes
from collections import namedtuple

import pyvisa  # Communication avec l'appareil de mesure

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,\
    NavigationToolbar2Tk

from numpy import linspace  # Domaine de balayage
from pandas import DataFrame  # Stockage et manipulation des données
from matplotlib import pyplot as plt  # Affichage des données


class HP4274AException(Exception):
    """Exception générique pour l'appareil HP4274A."""

    pass


class HP4274AOverflowException(HP4274AException):
    """Étendue de mesure trop basse."""

    pass


class HP4274AUnderflowException(HP4274AException):
    """Précision de mesure trop basse."""

    pass


class HP4274AWrongMeasuringFunctionException(HP4274AException):
    """Fonction de mesure innapropriée."""

    pass


# État de la machine
HP4274AState = namedtuple('HP4274A', ('circuit_mode',
                                      'measuring_frequency',
                                      'data_status_A',
                                      'function_A',
                                      'value_A',
                                      'data_status_B',
                                      'function_B',
                                      'value_B'))

# Fréquences utilisées par l'appareil
fs: list[float] = [100, 120, 200, 400, 1e3, 2e3, 4e3, 1e4, 2e4, 4e4, 1e5]

# Listes des ressources VISA utilisables
rm = pyvisa.ResourceManager()
ressources: list[str] = rm.list_resources()
nom_ressource: str = 'GPIB0::1::INSTR'  # Adresse de l'appareil de mesure


def query(appareil, cmd: str, delai=0.01) -> HP4274AState:
    """
    Envoyer une requête et lire la réponse.

    Parameters
    ----------
    appareil : TYPE
        DESCRIPTION.
    cmd : str
        DESCRIPTION.
    delai : TYPE, optional
        DESCRIPTION. The default is 0.01.

    Returns
    -------
    HP4274AState
        DESCRIPTION.

    """
    logging.debug('>%s', cmd)
    res = appareil.query(cmd)
    time.sleep(delai)
    logging.debug('%s', res)
    return check_status(res)


def regler_resolution(appareil, res: int = None, delai=0.01) -> HP4274AState:
    """
    Régler la résolution de l'appareil.

    Parameters
    ----------
    appareil : TYPE
        DESCRIPTION.
    res : int, optional
        DESCRIPTION. The default is None.
    delai : TYPE, optional
        DESCRIPTION. The default is 0.01.

    Returns
    -------
    HP4274AState
        DESCRIPTION.

    """
    lcrz_ranges = {-i: f'R{23-i}' for i in range(0, 12 + 1)}
    cmd = lcrz_ranges.get(res, 'R31')
    logging.debug('>%s', cmd)
    res = query(appareil, cmd, delai)
    logging.debug('%s', res)
    return res


def mesure_moyenne(appareil, cmd, N: int = 10, delai=0.01) -> float:
    """
    Prendre la moyenne d'une série de mesures.

    Parameters
    ----------
    appareil : TYPE
        DESCRIPTION.
    cmd : TYPE
        DESCRIPTION.
    N : int, optional
        DESCRIPTION. The default is 10.
    delai : TYPE, optional
        DESCRIPTION. The default is 0.01.

    Returns
    -------
    float
        DESCRIPTION.

    """
    res = sum([query(appareil, cmd, delai).value_A for i in range(N)]) / N
    return res


def regler_biais(appareil, v: float, delai=0.01) -> HP4274AState:
    """
    Régler le potentiel de biais.

    Parameters
    ----------
    appareil : TYPE
        DESCRIPTION.
    v : float
        DESCRIPTION.
    delai : TYPE, optional
        DESCRIPTION. The default is 0.01.

    Returns
    -------
    HP4274AState
        DESCRIPTION.

    """
    with decimal.localcontext() as context:
        context.prec = 3
        v = context.create_decimal(v)
        signe, chiffres, exposant = v.as_tuple()

        chiffres: list[str] = list(chiffres)
        while len(chiffres) < 3:
            chiffres.append('0')
            exposant -= 1

        mantisse: str = ''.join(str(j) for j in chiffres[:3])
        signe: str = {0: '+', 1: '-'}[signe]

        res: str = f'BI{signe}{mantisse}E{exposant:+03d}V'
    return query(appareil, res, delai)


def regler_freq(appareil, f: float, delai=0.01) -> HP4274AState:
    """
    Régler la fréquence du signal AC.

    Parameters
    ----------
    appareil : TYPE
        DESCRIPTION.
    f : float
        DESCRIPTION.
    delai : TYPE, optional
        DESCRIPTION. The default is 0.01.

    Returns
    -------
    HP4274AState
        DESCRIPTION.

    """
    i: int = fs.index(f)
    return query(appareil, f'F{11+i}', delai)


def check_status(res: str) -> HP4274AState:
    """
    Retourner l'état de la machine.

    Parameters
    ----------
    res : str
        DESCRIPTION.

    Returns
    -------
    HP4274AState
        DESCRIPTION.

    """
    A, B = res.strip().split(',')
    circuit_mode, measuring_freq, data_status_A, function_A, *value_A = A
    data_status_B, function_B, *value_B = B
    value_A = float(''.join(value_A))
    value_B = float(''.join(value_B))
    res = HP4274AState(circuit_mode,
                       measuring_freq,
                       data_status_A,
                       function_A,
                       value_A,
                       data_status_B,
                       function_B,
                       value_B)
    logging.debug('%s', res)
    return res


def enregistrer(df: DataFrame, sauvegarde: Path, comp: str):
    """
    Enregistrer les données et graphique.

    Parameters
    ----------
    df : DataFrame
        DESCRIPTION.
    sauvegarde : Path
        DESCRIPTION.
    comp : str
        DESCRIPTION.

    Returns
    -------
    None.

    """
    # Enregistrement en fichier Excel
    df.to_excel(sauvegarde / 'resultats.xlsx')

    # Graphique
    fig = plt.figure()
    ax = plt.gca()
    df.plot(ax=ax)
    plt.title(f'Courbes C-V pour composant {comp}')
    plt.legend()
    plt.savefig(sauvegarde / 'fig.png')
    plt.close(fig)


def exe(root,
        matricule_var,
        ressource_var,
        max_vb_var,
        min_vb_var,
        composante_var,
        progres,
        canvas,
        ax,
        biais_var,
        freq_var,
        delai_var,
        nbr_var,
        res_var,
        stop_var,
        min_freq_var,
        max_freq_var):
    """
    Prendre une série de mesures.

    Parameters
    ----------
    root : TYPE
        DESCRIPTION.
    matricule_var : TYPE
        DESCRIPTION.
    ressource_var : TYPE
        DESCRIPTION.
    max_vb_var : TYPE
        DESCRIPTION.
    min_vb_var : TYPE
        DESCRIPTION.
    composante_var : TYPE
        DESCRIPTION.
    progres : TYPE
        DESCRIPTION.
    canvas : TYPE
        DESCRIPTION.
    ax : TYPE
        DESCRIPTION.
    biais_var : TYPE
        DESCRIPTION.
    freq_var : TYPE
        DESCRIPTION.
    delai_var : TYPE
        DESCRIPTION.
    nbr_var : TYPE
        DESCRIPTION.
    res_var : TYPE
        DESCRIPTION.
    stop_var : TYPE
        DESCRIPTION.
    min_freq_var : TYPE
        DESCRIPTION.
    max_freq_var : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    df: DataFrame = DataFrame()

    nom: str = matricule_var.get()
    comp: str = composante_var.get()
    sauvegarde: Path = Path(f'{nom}/{comp}')
    nom_ressource: str = ressource_var.get()

    max_vb: float = float(max_vb_var.get())
    min_vb: float = float(min_vb_var.get())
    nbr: int = int(nbr_var.get())
    vb: list[float] = list(linspace(min_vb, max_vb, nbr)
                           )  # En potentiel de biais
    delai: float = float(delai_var.get())
    min_f = float(min_freq_var.get())
    max_f = float(max_freq_var.get())
    freqs = [f for f in fs if min_f <= f <= max_f]

    progres['maximum'] = len(freqs) * len(vb)

    # Vérifier l'existence du répertoire de sauvegarde
    if not sauvegarde.exists():
        if not sauvegarde.parent.exists():
            sauvegarde.parent.mkdir()
        sauvegarde.mkdir()

    _ = datetime.datetime.now()
    _ = _.isoformat()
    _ = _.replace(':', '_')
    _ = _.replace('.', '_')
    sauvegarde /= _
    sauvegarde.mkdir()

    # Enregistrer les paramètres de mesure
    with (sauvegarde / 'details.txt').open('w') as F:
        msg = f'''{nom=}
{comp=}
{freqs=}
{vb=}
{nom_ressource=}'''
        print(msg, file=F)
        logging.info(msg)

    try:
        # Ouvrir la connection avec l'appareil
        hp4274a = rm.open_resource(nom_ressource)

        res = query(hp4274a, 'A2', delai)
        if res.function_A != 'C':
            raise HP4274AWrongMeasuringFunctionException

        regler_resolution(hp4274a)
        query(hp4274a, 'C1', delai)

        # Balayage en fréquence
        for f in freqs:
            regler_freq(hp4274a, f, delai=delai)
            freq_var.set(f'{f:.2e}')
            root.update()

            # Balayage en potentiel
            for v in vb:
                if stop_var.get():
                    stop_var.set(0)
                    raise StopIteration

                res = regler_biais(hp4274a, v, delai=delai)
                biais_var.set(f'{v:.2e}')
                root.update()
                time.sleep(delai)
                if res.data_status_A == 'N' or res.value_A < 0:
                    res = mesure_moyenne(hp4274a, 'A2', 10, delai=delai)
                    df.loc[v, f] = res
                    res_var.set(res)

                    # Clear axis
                    ax.clear()
                    ax.set_axis_on()
                    df.plot(ax=ax, style='.', legend=True)
                    ax.set_xticks(vb, [f'{v:.2e}' for v in vb], rotation=45)
                    cs = linspace(min(df.min()), max(df.max()), 10)
                    ax.set_yticks(cs[::int(len(cs) / 10)],
                                  [f'{c:.2e}' for c in
                                   cs[::int(len(cs) / 10)]])
                    ax.set_xlabel('Potentiel de biais (V)')
                    ax.set_ylabel('Mesure de capacité (F)')
                    canvas.draw()

                progres.step()
                root.update()

        enregistrer(df, sauvegarde, comp)
    except StopIteration:
        pass
    except Exception as e:
        msg = f'Une erreur {e!r} est survenue, voir console ou journal.'
        mb.showerror(f'{type(e)}', msg)
        msg += f'\n{traceback.format_exc()}'
        logging.error(msg)
    finally:
        # Toujours fermer la connection à l'appareil
        hp4274a.close()


def exe1(root,
         matricule_var,
         ressource_var,
         max_vb_var,
         min_vb_var,
         composante_var,
         progres,
         canvas,
         ax,
         biais_var,
         freq_var,
         delai_var,
         nbr_var,
         res_var,
         stop_var,
         min_freq_var,
         max_freq_var):
    """
    Ajuster les paramètres et lire une mesure.

    Parameters
    ----------
    root : TYPE
        DESCRIPTION.
    matricule_var : TYPE
        DESCRIPTION.
    ressource_var : TYPE
        DESCRIPTION.
    max_vb_var : TYPE
        DESCRIPTION.
    min_vb_var : TYPE
        DESCRIPTION.
    composante_var : TYPE
        DESCRIPTION.
    progres : TYPE
        DESCRIPTION.
    canvas : TYPE
        DESCRIPTION.
    ax : TYPE
        DESCRIPTION.
    biais_var : TYPE
        DESCRIPTION.
    freq_var : TYPE
        DESCRIPTION.
    delai_var : TYPE
        DESCRIPTION.
    nbr_var : TYPE
        DESCRIPTION.
    res_var : TYPE
        DESCRIPTION.
    stop_var : TYPE
        DESCRIPTION.
    min_freq_var : TYPE
        DESCRIPTION.
    max_freq_var : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    """
    nom: str = matricule_var.get()
    comp: str = composante_var.get()
    sauvegarde: Path = Path(f'{nom}/{comp}')
    nom_ressource: str = ressource_var.get()

    vb: float = float(biais_var.get())  # En potentiel de biais

    # Vérifier l'existence du répertoire de sauvegarde
    if not sauvegarde.exists():
        if not sauvegarde.parent.exists():
            sauvegarde.parent.mkdir()
        sauvegarde.mkdir()

    _ = datetime.datetime.now()
    _ = _.isoformat()
    _ = _.replace(':', '_')
    _ = _.replace('.', '_')
    sauvegarde /= _
    sauvegarde.mkdir()

    # Enregistrer les paramètres de mesure
    with (sauvegarde / 'details.txt').open('w') as F:
        msg = f'''{nom=}
{comp=}
{fs=}
{vb=}
{nom_ressource=}'''
        print(msg, file=F)
        logging.info(msg)

    try:
        # Ouvrir la connection avec l'appareil
        hp4274a = rm.open_resource(nom_ressource)

        delai: float = float(delai_var.get())

        res = query(hp4274a, 'A2', delai)
        if res.function_A != 'C':
            raise HP4274AWrongMeasuringFunctionException

        regler_resolution(hp4274a)
        query(hp4274a, 'C1', delai)

        freq = float(freq_var.get())
        regler_freq(hp4274a, freq, delai=delai)
        res = regler_biais(hp4274a, vb, delai=delai)
        time.sleep(delai)
        if res.data_status_A == 'N':
            res = mesure_moyenne(hp4274a, 'A2', 10, delai=delai)
            res_var.set(res)
    except Exception as e:
        msg = f'Une erreur {e!r} est survenue, voir console ou journal.'
        mb.showerror(f'{type(e)}', msg)
        msg += f'\n{traceback.format_exc()}'
        logging.error(msg)
    finally:
        # Toujours fermer la connection à l'appareil
        hp4274a.close()


def main():
    """
    Lancer le programme de prise de mesures.

    Returns
    -------
    None.

    """
    # logging.basicConfig(filename='log.txt',
    #                    encoding='utf-8',
    #                    level=logging.DEBUG)
    # logging.basicConfig(level=logging.DEBUG)
    root = tk.Tk()
    root.title('HP4274A LCR-mètre')

    # Commande d'exécution de la fonction exe
    exe_bouton = ttk.Button(root,
                            text='n mesures',
                            command=lambda: exe(root=root,
                                                matricule_var=matricule_var,
                                                ressource_var=ressource_var,
                                                max_vb_var=max_vb_var,
                                                min_vb_var=min_vb_var,
                                                composante_var=composante_var,
                                                progres=progres,
                                                canvas=canvas,
                                                ax=ax,
                                                biais_var=biais_var,
                                                freq_var=freq_var,
                                                delai_var=delai_var,
                                                nbr_var=nbr_var,
                                                res_var=res_var,
                                                stop_var=stop_var,
                                                min_freq_var=min_freq_var,
                                                max_freq_var=max_freq_var))
    exe1_bouton = ttk.Button(root,
                             text='1 mesure',
                             command=lambda: exe1(root=root,
                                                  matricule_var=matricule_var,
                                                  ressource_var=ressource_var,
                                                  max_vb_var=max_vb_var,
                                                  min_vb_var=min_vb_var,
                                                  composante_var=composante_var,
                                                  progres=progres,
                                                  canvas=canvas,
                                                  ax=ax,
                                                  biais_var=biais_var,
                                                  freq_var=freq_var,
                                                  delai_var=delai_var,
                                                  nbr_var=nbr_var,
                                                  res_var=res_var,
                                                  stop_var=stop_var,
                                                  min_freq_var=min_freq_var,
                                                  max_freq_var=max_freq_var))
    stop_var = tk.IntVar(root, 0)
    stop_bouton = ttk.Button(root,
                             text='Arrêt',
                             command=lambda: stop_var.set(1))

    # Sélection de l'appareil avec lequel communiquer
    ressource_var = tk.StringVar(root, value=nom_ressource)
    combo_adresse = ttk.Combobox(root,
                                 values=ressources,
                                 textvariable=ressource_var,
                                 exportselection=True)
    label_adresse = ttk.Label(root, text='Adresse VISA')

    # Paramètres de la prise de mesures
    max_vb_var = tk.StringVar(root, value=6)
    min_vb_var = tk.StringVar(root, value=-6)
    max_vb_entry = ttk.Spinbox(
        root, from_=-35, to=35, increment=0.1, textvariable=max_vb_var)
    min_vb_entry = ttk.Spinbox(
        root, from_=-35, to=35, increment=0.1, textvariable=min_vb_var)
    max_vb_label = ttk.Label(root, text='Potentiel de biais maximal')
    min_vb_label = ttk.Label(root, text='Potentiel de biais minimal')

    max_freq_var = tk.StringVar(root, value=1e5)
    min_freq_var = tk.StringVar(root, value=1e3)
    max_freq_entry = ttk.Spinbox(
        root, from_=100, to=1e5, increment=100, textvariable=max_freq_var)
    min_freq_entry = ttk.Spinbox(
        root, from_=100, to=1e5, increment=100, textvariable=min_freq_var)
    max_freq_label = ttk.Label(root, text='Fréquence maximale')
    min_freq_label = ttk.Label(root, text='Fréquence minimale')

    # Identification de la série de mesures
    matricule_var = tk.StringVar(root)
    composante_var = tk.StringVar(root)
    matricule_entry = ttk.Entry(root, textvariable=matricule_var)
    matricule_label = ttk.Label(root, text="Matricule")
    composante_entry = ttk.Entry(root, textvariable=composante_var)
    composante_label = ttk.Label(root, text="Composante")

    # Barre de progrès
    progres = ttk.Progressbar(root,
                              orient=tk.HORIZONTAL,
                              maximum=220,
                              mode='determinate')

    # Graphique
    fig = Figure()
    ax = fig.add_subplot()
    canvas = FigureCanvasTkAgg(fig, master=root)
    barre = NavigationToolbar2Tk(canvas, root, pack_toolbar=False)
    barre.update()

    # Paramètres de prise de données
    biais_var = tk.StringVar(root, value=0)
    biais_entry = ttk.Spinbox(root, increment=0.1, textvariable=biais_var)
    biais_label = ttk.Label(root, text='Biais (V)')
    freq_var = tk.StringVar(root, value=1000)
    freq_entry = ttk.Spinbox(root, increment=100, textvariable=freq_var)
    freq_label = ttk.Label(root, text='Fréquence (Hz)')
    delai_var = tk.StringVar(root, value=0.01)
    delai_entry = tk.Spinbox(root, increment=0.01, textvariable=delai_var)
    delai_label = ttk.Label(root, text='Délai (s)')
    nbr_var = tk.StringVar(root, value=10)
    nbr_entry = ttk.Spinbox(root, increment=1, textvariable=nbr_var)
    nbr_label = ttk.Label(root, text='Nombre de points')
    res_var = tk.StringVar(root)
    res_entry = ttk.Entry(root, textvariable=res_var, state=tk.DISABLED)
    res_label = ttk.Label(root, text='Mesure (F)')

    # Positionnement
    pad = 5
    label_adresse.grid(row=0, column=0, sticky=tk.E, padx=pad, pady=pad)
    combo_adresse.grid(row=0, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    matricule_label.grid(row=1, column=0, sticky=tk.E, padx=pad, pady=pad)
    matricule_entry.grid(row=1, column=1,
                         sticky=tk.E + tk.W,
                         padx=pad, pady=pad)
    composante_label.grid(row=2, column=0, sticky=tk.E, padx=pad, pady=pad)
    composante_entry.grid(row=2, column=1,
                          sticky=tk.E + tk.W,
                          padx=pad, pady=pad)
    min_vb_label.grid(row=3, column=0, sticky=tk.E, padx=pad, pady=pad)
    min_vb_entry.grid(row=3, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    max_vb_label.grid(row=4, column=0, sticky=tk.E, padx=pad, pady=pad)
    max_vb_entry.grid(row=4, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    min_freq_label.grid(row=5, column=0, sticky=tk.E, padx=pad, pady=pad)
    min_freq_entry.grid(row=5, column=1,
                        sticky=tk.E + tk.W,
                        padx=pad, pady=pad)
    max_freq_label.grid(row=6, column=0, sticky=tk.E, padx=pad, pady=pad)
    max_freq_entry.grid(row=6, column=1,
                        sticky=tk.E + tk.W,
                        padx=pad, pady=pad)
    biais_label.grid(row=7, column=0, sticky=tk.E, padx=pad, pady=pad)
    biais_entry.grid(row=7, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    freq_label.grid(row=8, column=0, sticky=tk.E, padx=pad, pady=pad)
    freq_entry.grid(row=8, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    delai_label.grid(row=9, column=0, sticky=tk.E, padx=pad, pady=pad)
    delai_entry.grid(row=9, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    nbr_label.grid(row=10, column=0, sticky=tk.E, padx=pad, pady=pad)
    nbr_entry.grid(row=10, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    res_label.grid(row=11, column=0, sticky=tk.E, padx=pad, pady=pad)
    res_entry.grid(row=11, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    exe_bouton.grid(row=12, column=0, sticky=tk.E + tk.W, padx=pad, pady=pad)
    exe1_bouton.grid(row=12, column=1, sticky=tk.E + tk.W, padx=pad, pady=pad)
    stop_bouton.grid(row=13, column=0, sticky=tk.E + tk.W, padx=pad, pady=pad)
    progres.grid(row=14, column=0, columnspan=2,
                 sticky=tk.E + tk.W, padx=pad, pady=pad)
    canvas.get_tk_widget().grid(row=0, column=2,
                                rowspan=14,
                                padx=pad, pady=pad)
    barre.grid(row=14, column=2, padx=pad, pady=pad)

    root.mainloop()
    logging.shutdown()


if __name__ == '__main__':
    main()
