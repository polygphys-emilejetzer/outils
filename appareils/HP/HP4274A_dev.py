#!python
# -*- coding: utf-8 -*-
"""
Programme ou module pour ...

Créé le Thu Jan 13 09:45:19 2022

@author: ejetzer
"""

from collections import namedtuple
from decimal import Decimal
from pathlib import Path

import numpy as np
import pandas as pd

from matplotlib import pyplot as plt

from ...appareils import Appareil, Expérience


class HP4274A(Appareil):
    """Puissancemètre."""

    def bias(self, potential: float) -> str:
        """
        Régler le biais.

        Parameters
        ----------
        potential : float
            DESCRIPTION.

        Returns
        -------
        str
            DESCRIPTION.

        """
        signe, chiffres, exposant = Decimal(potential).as_tuple()
        mantisse = ''.join(chiffres[:3])
        signe = {0: '+', 1: '-'}[signe]
        return self.query(f'BI {signe}{mantisse} E {exposant:+02d} V')

    def fréquence(self, fréquence: float) -> str:
        """
        Régler la fréquence.

        Parameters
        ----------
        fréquence : float
            DESCRIPTION.

        Returns
        -------
        str
            DESCRIPTION.

        """
        mode = {100: 'F11',
                120: 'F12',
                200: 'F13',
                400: 'F14',
                1e3: 'F15',
                2e3: 'F16',
                4e3: 'F17',
                1e4: 'F18',
                2e4: 'F19',
                4e4: 'F20',
                1e5: 'F21'
                }[fréquence]
        return self.query(mode)

    def C(self):
        """
        Mesurer la capacité.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        res = self.query('A2')
        return float(res[5:17])

    def set(display_a_function: int = None,
            display_b_function: int = None,
            circuit_mode: int = None,
            deviation_measurement: int = None,
            frequency_step: int = None,
            high_resolution: int = None,
            data_ready: int = None,
            key_status_output: bool = None,
            level_monitor: str = None,
            multiplier: int = None,
            LCRZ_range: int = None,
            self_test: bool = None,
            trigger: int = None,
            zero_open: bool = None):
        """
        Régler l'état de la machine.

        Parameters
        ----------
        display_a_function : int, optional
            DESCRIPTION. The default is None.
        display_b_function : int, optional
            DESCRIPTION. The default is None.
        circuit_mode : int, optional
            DESCRIPTION. The default is None.
        deviation_measurement : int, optional
            DESCRIPTION. The default is None.
        frequency_step : int, optional
            DESCRIPTION. The default is None.
        high_resolution : int, optional
            DESCRIPTION. The default is None.
        data_ready : int, optional
            DESCRIPTION. The default is None.
        key_status_output : bool, optional
            DESCRIPTION. The default is None.
        level_monitor : str, optional
            DESCRIPTION. The default is None.
        multiplier : int, optional
            DESCRIPTION. The default is None.
        LCRZ_range : int, optional
            DESCRIPTION. The default is None.
        self_test : bool, optional
            DESCRIPTION. The default is None.
        trigger : int, optional
            DESCRIPTION. The default is None.
        zero_open : bool, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        pass

    def get(self, param: str = None) -> namedtuple:
        """
        Obtenir l'état de la machine.

        Parameters
        ----------
        param : str, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        namedtuple
            DESCRIPTION.

        """
        pass


class PHS8302_CV(Expérience):
    """Caractérisation CV."""

    @property
    def adresse(self) -> str:
        """
        Adresse de l'appareil.

        Returns
        -------
        str
            DESCRIPTION.

        """
        return self.config.get('gpib', 'adresse')

    @property
    def domaine_biais(self) -> np.linspace:
        """
        Retourne les potentiels de biais à utiliser.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        début = self.config.getfloat('biais', 'début')
        fin = self.config.getfloat('biais', 'fin')
        nombre_de_pas = self.config.getint('biais', 'nombre de pas')
        return np.linspace(début, fin, nombre_de_pas)

    @property
    def domaine_fréquence(self):
        """
        Retourne les fréquences à utiliser.

        Returns
        -------
        list
            DESCRIPTION.

        """
        return [100, 120, 200, 400, 1e3, 2e3, 4e3, 1e4, 2e4, 4e4, 1e5]

    def run(self):
        """
        Prendre une série de mesures.

        Returns
        -------
        None.

        """
        cs = {}

        impedance_metre = HP4274A(self.adresse)
        vs = self.domaine_biais
        fs = self.domaine_fréquence

        for f in fs:
            impedance_metre.fréquence(f)

            cs[f] = []

            for v in vs:
                impedance_metre.bias(v)
                cs[f].append(impedance_metre.C())

        df = pd.DataFrame(cs, index=vs)
        df.to_csv('capacités.xlsx')
        df.plot(kind='scatter')

        plt.legend()
        plt.show()


if __name__ == '__main__':
    config = Path(__file__).parent / 'phs8302.cfg'
    exp = PHS8302_CV(config)
    exp.run()
