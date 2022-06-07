# -*- coding: utf-8 -*-
"""
Programme ou module pour gérer des appareils de mesure.

Ce module contient les classes génériques pour communiquer avec et contrôler
des appareils de mesures, et donc automatiser certaines tâches de laboratoire.

Le module utilise pyVISA et l'interface VISA plus généralement pour les
communications. VISA fonctionne sur Windows, MacOS et Linux une fois les
drivers adéquats installés.
"""

# Bibliothèques standard
from pathlib import Path  # Manipulation de chemins et fichiers

# Bibliothèques via PIPy
import pyvisa as visa  # Communication série avec des appareils externes

# Imports relatifs
from ..config import FichierConfig  # Fichiers de configuration


class Gestionnaire:
    """Gestionnaire d'appareils."""

    def __init__(self):
        self.rm = visa.ResourceManager()

    def list_resources(self):
        return self.rm.list_resources()

    def open(self, nom: str) -> Appareil:
        return Appareil(self, nom).open()

    def grid(self):
        pass

    def pack(self):
        pass


class Appareil:
    """Appareil."""

    def __init__(self, gestionnaire: Gestionnaire, nom: str):
        self.nom = nom
        self.gestionnaire: Gestionnaire = gestionnaire

    def open(self):
        rm = self.gestionnaire.rm
        self.resource = rm.open_resource(self.nom)
        return self

    def close(self):
        self.resource.close()

    def read(self) -> str:
        return self.resource.read()

    def write(self, m: str):
        self.resource.write(m)

    def query(self, q: str) -> str:
        return self.resource.query(q)

    def get(self):
        pass

    def set(self):
        pass

    def grid(self):
        pass

    def pack(self):
        pass

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def export(self):
        pass


class Expérience:
    """Montage et prise de mesures."""

    def __init__(self, fichier_config: Path):
        self.config = FichierConfig(fichier_config)

    def build(self):
        pass

    def run(self):
        pass
