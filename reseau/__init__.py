# -*- coding: utf-8 -*-
"""
Gestion de connexions réseau.

Facilite les connexions à des disques réseau ou à des VPNs.
"""

# Bibliothèque standard
import time
import datetime
import os
import platform
import io
import urllib.request
import requests
import getpass

from subprocess import run
from pathlib import Path


class ExceptionDisqueReseau(Exception):
    """Exception générique avec les disques réseau."""

    pass


class LePointDeMontageExiste(ExceptionDisqueReseau):
    """Le point de montage existe déjà."""

    pass


class LeVolumeNEstPasMonte(ExceptionDisqueReseau):
    """Le volume n,est pas monté."""

    pass


class LePointDeMontageNExistePas(ExceptionDisqueReseau):
    """Le point de montage n'existe pas."""

    pass


class ErreurDeMontage(ExceptionDisqueReseau):
    """Le montage n'a pas réussi."""

    pass


class DisqueRéseau:
    """Disque réseau."""

    @staticmethod
    def mount_cmd(nom: str, mdp: str, url: str, mode: str, chemin: Path):
        """Commande de montage de disque."""
        return ['mount', '-t', mode, f'//{nom}:{mdp}@{url}', str(chemin)]

    @staticmethod
    def umount_cmd(chemin: Path):
        """Commande de démontage de disque."""
        return ['umount', str(chemin)]

    @staticmethod
    def net_use_cmd(nom: str, mdp: str, url: str, drive: str):
        """Commande de montage Windows."""
        return ['net', 'use', drive, url, f'/u:{nom}', mdp]

    @staticmethod
    def net_use_delete_cmd(drive: str, url: str):
        """Commande de démontage Windows."""
        return ['net', 'use', drive, url, '/delete']

    def __init__(self,
                 adresse: str,
                 chemin: Path,
                 drive: str,
                 nom: str,
                 mdp: str,
                 mode: str = 'smbfs',
                 timeout: int = 1):
        """
        Disque réseau.

        Parameters
        ----------
        adresse : str
            DESCRIPTION.
        chemin : Path
            DESCRIPTION.
        nom : str
            DESCRIPTION.
        mdp : str
            DESCRIPTION.
        mode : str, optional
            DESCRIPTION. The default is 'smbfs'.
        timeout : int, optional
            DESCRIPTION. The default is 10.

        Returns
        -------
        None.

        """
        self.adresse = adresse
        self.chemin = chemin if isinstance(chemin, Path) else Path(chemin)
        self.drive = drive
        self.nom = nom
        self.mdp = mdp
        self.mode = mode
        self.timeout = timeout

    def mount(self):
        """
        Monter le disque.

        Raises
        ------
        ErreurDeMontage
            DESCRIPTION.
        LePointDeMontageExiste
            DESCRIPTION.

        Returns
        -------
        None.

        """
        if not self.exists():
            self.chemin.mkdir()
            if platform.system() == 'Windows':
                res = run(self.mount_cmd(self.nom,
                                         self.mdp,
                                         self.adresse,
                                         self.drive))
            else:
                res = run(self.mount_cmd(self.nom,
                                         self.mdp,
                                         self.adresse,
                                         self.mode,
                                         self.chemin))
            for i in range(self.timeout * 1000):
                if self.is_mount():
                    break
                else:
                    time.sleep(0.001)
            if not self.is_mount():
                self.chemin.rmdir()
                raise ErreurDeMontage(f'Valeur retournée de {res}')
        else:
            raise LePointDeMontageExiste(f'{self.chemin!r} existe déjà.')

    def umount(self):
        """
        Démonter le disque.

        Raises
        ------
        LeVolumeNEstPasMonte
            DESCRIPTION.
        LePointDeMontageNExistePas
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        if self.exists():
            if self.is_mount():
                if platform.system() == 'Windows':
                    return run(self.net_use_delete_cmd(self.drive,
                                                       self.adresse))
                else:
                    res = run(self.umount_cmd(self.chemin))
                    self.chemin.rmdir()
                    return res
            else:
                raise LeVolumeNEstPasMonte(
                    f'{self.url!r} n\'est pas monté au point {self.chemin!r}.')
        else:
            raise LePointDeMontageNExistePas(
                f'Le point de montage {self.chemin!r} n\'existe pas.')

    def __enter__(self):
        """
        Monter sécuritairement.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        if not self:
            self.mount()

        with (self / f'.{os.getpid()}.lock').open('w') as lockfile:
            lockfile.write(f'{datetime.datetime.now()}')

        return self

    def __exit__(self, exc_type, exc_value, traceback):
        """
        Démonter en fin d'utilisation.

        Parameters
        ----------
        exc_type : TYPE
            DESCRIPTION.
        exc_value : TYPE
            DESCRIPTION.
        traceback : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        (self / f'.{os.getpid()}.lock').unlink()
        loquets = list((self / '.').glob('.*.lock'))

        if not len(loquets):
            self.umount()

    def is_mount(self):
        """
        Vérifie le montage.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        if platform.system() == 'Windows':
            return Path(f'{self.drive}:').is_mount()
        else:
            return self.chemin.is_mount()

    def exists(self):
        """
        Vérifie l'existence.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        if platform.system() == 'Windows':
            return Path(f'{self.drive}:').exists()
        else:
            return self.chemin.exists()

    def __bool__(self):
        """
        Vérifie l'existence et le montage.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        return self.exists() and self.is_mount()

    def __truediv__(self, other):
        """
        Naviguer dans le disque.

        Parameters
        ----------
        other : TYPE
            DESCRIPTION.

        Returns
        -------
        TYPE
            DESCRIPTION.

        """
        if platform.system() == 'Windows':
            return Path(f'{self.drive}:') / other
        else:
            return self.chemin / other

    def __rtruediv__(self, other):
        """Rien."""
        return NotImplemented


class FichierLointain:

    def __init__(self, chemin_distant, chemin_local):
        self.chemin_distant = chemin_distant
        self.chemin_local = Path(chemin_local)

    def update(self):
        self.open()
        self.close()

    def open(self, mode='rb'):
        req = requests.get(self.chemin_distant)

        if req.status_code == 403:
            nom = input('nom: ')
            mdp = getpass.getpass('mdp: ')
            req = requests.get(self.chemin_distant, auth=(nom, mdp))

        if req.status_code == 200:
            with self.chemin_local.open('wb') as g:
                g.write(req.content)

            self.fichier = self.chemin_local.open(mode)
            return self.fichier
        else:
            raise Exception(
                f'Fichier non disponible, code d\'état {req.status_code}')

    def close(self):
        self.fichier.close()

    def read(self, *args, **kargs):
        return self.fichier.read(*args, **kargs)

    def readlines(self, *args, **kargs):
        return self.fichier.read(*args, **kargs)

    def seek(self, *args, **kargs):
        return self.fichier.seek(*args, **kargs)

    def tell(self, *args, **kargs):
        return self.fichier.tell(*args, **kargs)

    def __enter__(self):
        return self.open()

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()


class OneDrive:

    def __init__(self, utilisateur, organisation='', *args, partagé=False):
        self.utilisateur = utilisateur
        self.organisation = organisation
        self.sous_dossiers = args
        self.partagé = partagé

    @property
    def dossier_onedrive(self):
        if self.partagé:
            return self.organisation
        elif self.organisation:
            return f'OneDrive -{self.organisation}'
        else:
            return 'OneDrive'

    @property
    def chemin(self):
        chemin = Path(f'~{self.utilisateur}').expanduser() / \
            self.dossier_onedrive

        for sd in self.sous_dossiers:
            chemin /= sd

        return chemin

    def __bool__(self):
        return self.chemin.exists()

    def __truediv__(self, other):
        return self.chemin / other

    def __rtruediv__(self, other):
        return NotImplemented
