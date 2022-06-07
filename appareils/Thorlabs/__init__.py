# -*- coding: utf-8 -*-
"""
Created on Tue Aug  3 10:36:48 2021.

@author: Émile Jetzer, Vincent Perreault
"""

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,\
    NavigationToolbar2Tk
import time
import smtplib
import pathlib
import tkinter as tk

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Download https://sourceforge.net/projects/libusb-win32/files/latest/download
import usbtmc
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from pylablib.devices import Thorlabs as Tl
# from ThorlabsPM100 import ThorlabsPM100


# Configuration spéciale de matplotlib pour afficher des graphiques
# Tiré de
# https://pythonprogramming.net/how-to-embed-matplotlib-graph-tkinter-gui/
mpl.use("TkAgg")


def trouver_proche(liste: list, valeur: float):
    """
    Trouver la valeur la plus près d'une valeur de référence.

    Parameters
    ----------
    liste : list
        DESCRIPTION.
    valeur : float
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    tableau: np.array = np.asarray(liste)
    indice: int = (np.abs(tableau - valeur)).argmin()
    return tableau[indice]


class Données:
    """Classe *mutable* pour les opérations entre éléments d'interface."""

    def __init__(self):
        """
        Classe *mutable* pour les opérations entre éléments d'interface.

        Returns
        -------
        None.

        """
        self.réinitialiser()

    def réinitialiser(self):
        """
        Réinitialiser les données.

        Returns
        -------
        None.

        """
        self.position: list[float] = []
        self.puissance: list[float] = []

    @property
    def sommet(self):
        """
        Trouver le sommet de la courbe.

        Returns
        -------
        x_max : TYPE
            DESCRIPTION.
        y_max : TYPE
            DESCRIPTION.

        """
        y_max = max(self.puissance)
        x_max = self.position[self.puissance.index(y_max)]
        return x_max, y_max

    @property
    def epsilon(self):
        """
        Estimer la largeur de la courbe.

        Returns
        -------
        y_epsilon : TYPE
            DESCRIPTION.
        epsilon : TYPE
            DESCRIPTION.

        """
        x_max, y_max = self.sommet
        y_epsilon = y_max / 2
        y_epsilon = trouver_proche(self.puissance, y_epsilon)
        i_1 = self.puissance.index(y_epsilon)
        epsilon = abs(2 * (x_max - self.position[i_1]))
        return y_epsilon, epsilon

    def graphique(self,
                  fig: plt.Figure,
                  ylabel: str = '',
                  xlabel: str = '',
                  title: str = ''):
        """
        Afficher le graphique.

        Parameters
        ----------
        fig : plt.Figure
            DESCRIPTION.
        ylabel : str, optional
            DESCRIPTION. The default is ''.
        xlabel : str, optional
            DESCRIPTION. The default is ''.
        title : str, optional
            DESCRIPTION. The default is ''.

        Returns
        -------
        fig : TYPE
            DESCRIPTION.
        ax : TYPE
            DESCRIPTION.

        """
        ax = fig.gca()
        ax.clear()
        ax.plot(self.position, self.puissance)
        ax.set_ylabel(ylabel)
        ax.set_xlabel(xlabel)
        ax.set_title(title)
        y_epsilon, epsilon = self.epsilon
        x_max, y_max = self.sommet
        ax.annotate(f'$\\epsilon = {epsilon:.2f}$ mm', (x_max, y_epsilon))

        return fig, ax

    def exporter(self, fig: plt.Figure):
        """
        Exporter les données.

        Parameters
        ----------
        fig : plt.Figure
            DESCRIPTION.

        Returns
        -------
        nom_tableur : TYPE
            DESCRIPTION.
        nom_image : TYPE
            DESCRIPTION.

        """
        cadre = pd.DataFrame(
            {'Position': self.position, 'Puissance': self.puissance})

        temps = time.ctime().replace(':', '_')
        nom_dossier = pathlib.Path(f"~/Desktop/Résultats {temps}").expanduser()
        nom_dossier.mkdir()
        nom_tableur = nom_dossier / f'Données {temps}.csv'
        nom_image = nom_dossier / f'Données {temps}.png'

        cadre.to_csv(nom_tableur)
        fig.savefig(nom_image)

        return nom_tableur, nom_image

    def courriel(self, pièces_jointes: list[str], émmetteur: str):
        """
        Partager les données par courriel.

        Parameters
        ----------
        pièces_jointes : list[str]
            DESCRIPTION.
        émmetteur : str
            DESCRIPTION.

        Returns
        -------
        None.

        """
        contenu = "Voici les données du labo laser:"

        récepteur = 'emile.jetzer@polymtl.ca'

        message = MIMEMultipart()
        message['From'] = émmetteur
        message['To'] = récepteur
        message['Cc'] = émmetteur
        message['Subject'] = 'Données du labo laser'

        message.attach(MIMEText(contenu, 'plain'))

        for nom, type_mime in zip(pièces_jointes,
                                  (('text', 'csv'),
                                   ('image', 'png'))):
            with open(nom, 'rb') as pièce_jointe:
                payload = MIMEBase(*type_mime)
                payload.set_payload(pièce_jointe.read())
                encoders.encode_base64(payload)
                payload.add_header('Content-Disposition',
                                   f'attachment; filename={nom.name}')
                message.attach(payload)

        serveur = smtplib.SMTP('smtp.polymtl.ca', 25)
        texte = message.as_string()
        serveur.sendmail(émmetteur, [récepteur, émmetteur], texte)


class DummyPuissancemètre:
    """Pour les tests."""

    def __init__(self, id_détecteur: tuple[int, int] = None):
        """
        Pour les tests.

        Parameters
        ----------
        id_détecteur : tuple[int, int], optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        # connection au puissance mètre
        self.__détecteur: usbtmc.Instrument = None
        self.open()

    def open(self):
        """Rien."""
        pass

    def close(self):
        """Rien."""
        pass

    def read(self):
        """Retourne les données au hazard."""
        from random import random
        return random()


class Puissancemètre(DummyPuissancemètre):
    """Vrai puissancemètre."""

    def __init__(self, id_détecteur: tuple[int, int] = None):
        """
        Connecter au vrai puissancemètre.

        Parameters
        ----------
        id_détecteur : tuple[int, int], optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        if id_détecteur is None:
            # liste les appareil connectée
            info = str(usbtmc.list_devices()[0])

            # trouvé les IDs
            posven = info.find('idVendor')
            id_vendeur = info[posven + 25:posven + 31]

            pospro = info.find('idProduct')
            id_produit = info[pospro + 25:pospro + 31]

            id_vendeur, id_produit = int(id_vendeur, 16), int(id_produit, 16)
        else:
            id_vendeur, id_produit = id_détecteur

        self.__détecteur: usbtmc.Instrument = usbtmc.Instrument(
            id_vendeur, id_produit)  # connection au puissance mètre
        self.open()

    def open(self):
        """
        Ouvrir la connection.

        Returns
        -------
        None.

        """
        self.__détecteur.open()

    def close(self):
        """
        Fermer la conenction.

        Returns
        -------
        None.

        """
        self.__détecteur.close()

    def read(self):
        """
        Lire les données.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return float(self.__détecteur.ask('READ?'))


class DummyMoteur:
    """Moteur pour les tests."""

    # constante trouver sur
    # https://www.thorlabs.com/Software/Motion%20Control/APT_Communications_Protocol.pdf
    T = 2048 / 6e6
    V = 65536
    pas_par_mm = 34304  # Scaling factor en fonction du moteur

    def __init__(self, id_moteur=None):
        """
        Moteur pour les tests.

        Parameters
        ----------
        id_moteur : TYPE, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        self.__moteur = None  # connection du stage
        self.__position = 0
        self.aller(0)
        self.attendre()

    @property
    def position(self):
        """
        Retourne la position fictive.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return self.__position

    def mesurer(self,
                détecteur: Puissancemètre,
                données: Données,
                dx: float = 6.0):
        """
        Effectue une mesure fictive.

        Parameters
        ----------
        détecteur : Puissancemètre
            DESCRIPTION.
        données : Données
            DESCRIPTION.
        dx : float, optional
            DESCRIPTION. The default is 6.0.

        Returns
        -------
        None.

        """
        nombre_de_pas = dx * self.pas_par_mm  # distance a parcourir

        self.aller(0)
        self.attendre()

        données.réinitialiser()
        for i in range(int(nombre_de_pas)):
            self.aller(i)
            x, P = self.position, détecteur.read()
            données.position.append(x)
            données.puissance.append(P)

        self.aller(0)
        self.attendre()

    def aller(self, position_finale: float = 0):
        """
        Se déplacer jusqu'à une position.

        Parameters
        ----------
        position_finale : float, optional
            DESCRIPTION. The default is 0.

        Returns
        -------
        None.

        """
        self.__position = position_finale

    def attendre(self):
        """Rien."""
        pass

    def close(self):
        """Rien."""
        pass


class Moteur(DummyMoteur):
    """Moteur Thorlabs."""

    # constante trouver sur
    # https://www.thorlabs.com/Software/Motion%20Control/APT_Communications_Protocol.pdf
    T = 2048 / 6e6
    V = 65536
    pas_par_mm = 34304  # Scaling factor en fonction du moteur

    def __init__(self, id_moteur=None):
        """
        Moteur Thorlabs.

        Parameters
        ----------
        id_moteur : TYPE, optional
            DESCRIPTION. The default is None.

        Returns
        -------
        None.

        """
        if id_moteur is None:
            id_moteur = Tl.list_kinesis_devices()[0][0]  # list l'ID du stage

        self.__moteur = Tl.KinesisMotor(id_moteur)  # connection du stage
        self.aller(0)
        self.attendre()

    @property
    def position(self):
        """
        Position du moteur.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return float(self.__moteur.get_position() / self.pas_par_mm)

    def mesurer(self,
                détecteur: Puissancemètre,
                données: Données,
                dx: float = 6.0):
        """
        Prendre des mesures.

        Parameters
        ----------
        détecteur : Puissancemètre
            DESCRIPTION.
        données : Données
            DESCRIPTION.
        dx : float, optional
            DESCRIPTION. The default is 6.0.

        Returns
        -------
        None.

        """
        nombre_de_pas = dx * self.pas_par_mm  # distance a parcourir
        self.__moteur.setup_velocity(0, 250, 100005)

        self.aller(0)
        self.attendre()

        self.__moteur.move_by(nombre_de_pas)  # mouvement

        données.réinitialiser()
        while self.__moteur.is_moving():  # capture de position et puissance
            try:
                x, P = self.position, détecteur.read()
                données.position.append(x)
                données.puissance.append(P)
            except Exception as ex:
                print(ex)

        self.aller(0)
        self.attendre()

    def aller(self, position_finale: float = 0):
        """
        Aller jusqu'à une position.

        Parameters
        ----------
        position_finale : float, optional
            DESCRIPTION. The default is 0.

        Returns
        -------
        None.

        """
        self.__moteur.setup_velocity(0, 250, 700005)
        self.__moteur.move_to(position_finale)

    def attendre(self):
        """
        Attendre que le moteur ait fini de bouger.

        Returns
        -------
        None.

        """
        self.__moteur.wait_move()

    def close(self):
        """
        Fermer la connection au moteur.

        Returns
        -------
        None.

        """
        self.__moteur.close()


class LabGui(tk.Frame):
    """Interface de prise de mesures."""

    def __init__(self,
                 maître: tk.Tk,
                 étage_de_translation: Moteur,
                 puissancemètre: Puissancemètre,
                 données: Données,
                 *args, **kargs):
        """
        Interface de prise de mesures.

        Parameters
        ----------
        maître : tk.Tk
            DESCRIPTION.
        étage_de_translation : Moteur
            DESCRIPTION.
        puissancemètre : Puissancemètre
            DESCRIPTION.
        données : Données
            DESCRIPTION.
        *args : TYPE
            DESCRIPTION.
        **kargs : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        super().__init__(master=maître, *args, **kargs)

        self.étage_de_translation = étage_de_translation
        self.puissancemètre = puissancemètre
        self.données = données

        self.cadre_formulaire = tk.Frame(self)
        self.bouton_exécuter = tk.Button(self.cadre_formulaire,
                                         text="Exécuter",
                                         command=lambda: self.exécuter())
        self.cadre_entrée = tk.Frame(self.cadre_formulaire)
        self.variable_courriel = tk.StringVar()
        self.étiquette_courriel = tk.Label(self.cadre_entrée, text='Courriel:')
        self.champ_courriel = tk.Entry(
            self.cadre_entrée, textvariable=self.variable_courriel)

        self.cadre_paramètres = tk.Frame(self)
        self.cadre_xaxis = tk.Frame(self.cadre_paramètres)
        self.variable_xaxis = tk.StringVar()
        self.étiquette_xaxis = tk.Label(
            self.cadre_xaxis, text='Titre de l\'abscisse:')
        self.champ_xaxis = tk.Entry(
            self.cadre_xaxis, textvariable=self.variable_xaxis)

        self.cadre_yaxis = tk.Frame(self.cadre_paramètres)
        self.variable_yaxis = tk.StringVar()
        self.étiquette_yaxis = tk.Label(
            self.cadre_yaxis, text='Titre de l\'ordonnée:')
        self.champ_yaxis = tk.Entry(
            self.cadre_yaxis, textvariable=self.variable_yaxis)

        self.cadre_titre = tk.Frame(self.cadre_paramètres)
        self.variable_titre = tk.StringVar()
        self.étiquette_titre = tk.Label(
            self.cadre_titre, text='Titre du graphique:')
        self.champ_titre = tk.Entry(
            self.cadre_titre, textvariable=self.variable_titre)

        self.figure = plt.Figure(figsize=(10, 10))
        self.canevas = FigureCanvasTkAgg(self.figure, self)
        self.outils = NavigationToolbar2Tk(self.canevas, self)

    def pack(self, *args, **kargs):
        """
        Afficher l'interface.

        Parameters
        ----------
        *args : TYPE
            DESCRIPTION.
        **kargs : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        self.cadre_formulaire.pack(side=tk.TOP, fill=tk.X)
        self.bouton_exécuter.pack(side=tk.RIGHT)
        self.cadre_entrée.pack(side=tk.LEFT, fill=tk.X)
        self.étiquette_courriel.pack(side=tk.LEFT)
        self.champ_courriel.pack(fill=tk.X)

        self.cadre_paramètres.pack(side=tk.TOP, fill=tk.X)
        self.cadre_xaxis.pack(fill=tk.X)
        self.étiquette_titre.pack(side=tk.LEFT)
        self.champ_titre.pack(fill=tk.X)
        self.cadre_yaxis.pack(fill=tk.X)
        self.étiquette_xaxis.pack(side=tk.LEFT)
        self.champ_xaxis.pack(fill=tk.X)
        self.cadre_titre.pack(fill=tk.X)
        self.étiquette_yaxis.pack(side=tk.LEFT)
        self.champ_yaxis.pack(fill=tk.X)

        self.outils.update()
        self.canevas.get_tk_widget().pack(side=tk.BOTTOM, fill=tk.BOTH)

        super().pack(*args, **kargs)

    def exécuter(self):
        """
        Prendre une série de mesures.

        Returns
        -------
        None.

        """
        self.données.réinitialiser()
        self.étage_de_translation.mesurer(self.puissancemètre, self.données)

        fig, ax = self.données.graphique(self.figure,
                                         self.variable_yaxis.get(),
                                         self.variable_xaxis.get(),
                                         self.variable_titre.get())
        self.canevas.draw()

        pièces_jointes = self.données.exporter(self.figure)
        try:
            self.données.courriel(pièces_jointes, self.variable_courriel.get())
        except Exception:
            pass

    def destroy(self):
        """
        Effacer le widget.

        Returns
        -------
        None.

        """
        self.étage_de_translation.close()
        self.puissancemètre.close()
        super().destroy()


test = False
if test:
    ClasseMoteur, ClassePuissancemètre = DummyMoteur, DummyPuissancemètre
else:
    ClasseMoteur, ClassePuissancemètre = Moteur, Puissancemètre

if __name__ == '__main__':
    fenêtre = tk.Tk()
    fenêtre.title('Labo Laser')
    interface = LabGui(fenêtre, ClasseMoteur(),
                       ClassePuissancemètre(), Données())

    interface.pack()
    fenêtre.mainloop()
