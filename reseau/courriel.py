# -*- coding: utf-8 -*-
"""Envoi facile de courriels."""

# Bibliothèque standard
import os
import sys
import itertools
import time
import logging
import mimetypes
import smtplib

import email.message
import email.parser
import email.policy

from email.message import EmailMessage
from pathlib import Path
from functools import partial
from datetime import datetime
from imaplib import IMAP4_SSL
from typing import Any

import getpass
import quopri
import dateutil.parser
import chardet
import schedule
import pandas
import keyring

from polygphys.outils.config import FichierConfig
from polygphys.outils.base_de_donnees import BaseTableau
from polygphys.outils.reseau import DisqueRéseau

CONFIGURATION_PAR_DÉFAUT: str = '''
[messagerie]
adresse = 
nom = 

[db]
adresse = 
'''


class CourrielsConfig(FichierConfig):

    def default(self) -> str:
        return CONFIGURATION_PAR_DÉFAUT


class Courriel:
    équivalences_attributs = {'destinataire': 'To',
                              'expéditeur': 'From',
                              'objet': 'Subject'}

    def __init__(self,
                 destinataire=None,
                 expéditeur=None,
                 objet=None,
                 contenu=None,
                 html=None,
                 pièces_jointes=tuple(),
                 message: EmailMessage = None):
        if message:
            self.message = message
        else:
            self.message = EmailMessage()

        self.destinataire = destinataire
        self.expéditeur = expéditeur
        self.objet = objet
        self.contenu = contenu
        self.html = html

        self.pièces_jointes = set()
        self.joindre(*pièces_jointes)

    def __getitem__(self, clé: Any) -> Any:
        return self.message[clé]

    def __setitem__(self, clé: Any, val: Any) -> Any:
        self.message[clé] = val

    def __getattr__(self, clé: str) -> Any:
        if clé in self.équivalences_attributs:
            clé = self.équivalences_attributs[clé]
            return self[clé]
        else:
            return getattr(self.message, clé)

    def __setattr__(self, clé: str, val: Any) -> Any:
        if clé == 'contenu':
            self.message.set_content(val)
        elif clé == 'html':
            self.message.add_alternative(val, subtype='html')
        elif clé in self.équivalences_attributs:
            clé = self.équivalences_attributs[clé]
            self[clé] = val
        else:
            setattr(self.message, clé, val)

    def joindre(self, *chemins):
        for chemin in chemins:
            chemin = Path(chemin)

            type_mime = mimetypes.guess_type(chemin.name)
            if None in type_mime:
                type_mime = ('application', 'octet-stream')

            with chemin.open('rb') as f:
                self.message.add_attachment(f.read(),
                                            maintype=type_mime[0],
                                            subtype=type_mime[1],
                                            filename=chemin.name)

    def envoyer(self, adresse, port=25):
        self.construire()
        serveur = smtplib.SMTP(adresse, port)
        serveur.send_message(self.message)
        serveur.quit()

    @property
    def date(self) -> datetime:
        return dateutil.parser.parse(self['Date'], ignoretz=True)

    @property
    def contenu(self):
        contenu = self.message.get_body(('plain', 'html', 'related'))

        if contenu is None:
            contenu = f'{contenu!r}'
        else:
            contenu = contenu.get_payload()

        if contenu.isascii():
            contenu = quopri.decodestring(contenu.encode('utf-8'))

            encodings = ('utf-8',
                         'cp1252',
                         chardet.detect(contenu)['encoding'])
            for encoding in encodings:
                try:
                    contenu = contenu.decode(encoding)
                except UnicodeDecodeError:
                    logging.exception(f'{encoding} ne fonctionne pas.')
                else:
                    break

        if isinstance(contenu, bytes):
            contenu = str(contenu, encoding='utf-8')

        contenu = contenu.replace('\r', '\n')
        while '\n\n\n' in contenu:
            contenu = contenu.replace('\n\n', '\n')

        return contenu

    @staticmethod
    def nettoyer_nom(nom: str) -> str:
        for c in ':, )(.?![]{}#/\\':
            nom = nom.replace(c, '_')
        while '__' in nom:
            nom = nom.replace('__', '_')

        nom = nom.strip('_')

        return nom

    @property
    def name(self) -> str:
        sujet: str = self['Subject']\
            .encode('ascii', 'ignore')\
            .decode('utf-8')\
            .strip()

        sujet = self.nettoyer_nom(sujet)

        return sujet + '.md'

    @property
    def parent(self) -> Path:
        nom = self.message.get('Thread-Topic', self.name[:-3])
        nom = self.nettoyer_nom(nom)

        for prefix, f in itertools.product(('fwd', 're', 'tr', 'ré'),
                                           (str.upper,
                                            str.lower,
                                            str.capitalize,
                                            str.title,
                                            str)):
            nom = nom.replace(f(prefix), '')

        return Path(nom)

    @property
    def path(self):
        return self.parent / self.name

    def __str__(self):
        return f'''- - - 
Date: {self.date.isoformat()}
De: {self['From']}
À: {self['To']}
Sujet: {self['Subject']}

{self.contenu}
'''

    def sauver(self, dossier: Path):
        chemin = dossier / self.path

        if not chemin.parent.exists():
            chemin.parent.mkdir()
        if not chemin.exists():
            chemin.touch()

        with chemin.open('w', encoding='utf-8') as f:
            f.write(str(self))


class Messagerie:

    def __init__(self, config: CourrielsConfig):
        if isinstance(config, (str, Path)):
            self.config = CourrielsConfig(config)
        else:
            self.config = config

        self._mdp = None

    @property
    def adresse(self):
        return self.config.get('messagerie', 'adresse')

    @property
    def nom(self):
        return self.config.get('messagerie', 'nom')

    @property
    def mdp(self):
        mdp_sys = keyring.get_password('system',
                                       nom := f'courriels.{self.nom}')
        if self._mdp is None and mdp_sys is None:
            self._mdp = getpass.getpass('mdp>')
            keyring.set_password('system', nom, self._mdp)
        elif self._mdp is None:
            self._mdp = keyring.get_password('system', nom)

        return self._mdp

    def message(self, serveur: IMAP4_SSL, numéro: str) -> Courriel:
        typ, data = serveur.fetch(numéro, '(RFC822)')
        message = email.parser.BytesParser(policy=email.policy.default)\
            .parsebytes(bytes(data[0][1]))

        return Courriel(message)

    def messages(self) -> Courriel:
        with IMAP4_SSL(self.adresse) as serveur:
            serveur.noop()
            serveur.login(self.nom, self.mdp)
            serveur.enable('UTF-8=ACCEPT')
            serveur.select()  # TODO: permettre de sélectionner d'autres boîtes
            typ, data = serveur.search(None, 'ALL')
            messages: list[str] = data[0].split()
            f = partial(self.message, serveur)

            yield from map(f, messages)

    def __iter__(self):
        return self.messages()

    @property
    def df(self) -> pandas.DataFrame:
        return pandas.DataFrame([[c.date,
                                  c['Subject'],
                                  c['From'],
                                  c['To'],
                                  c.parent.name,
                                  c.contenu] for c in self],
                                columns=('date',
                                         'sujet',
                                         'de',
                                         'a',
                                         'chaine',
                                         'contenu'))


class CourrielsTableau(BaseTableau):

    def __init__(self, config: CourrielsConfig):
        if isinstance(config, (str, Path)):
            self.config = CourrielsConfig(config)
        else:
            self.config = config

        db = self.config.get('db', 'adresse')
        table = 'courriels'

        super().__init__(db, table)

    def ajouter_messagerie(self, messagerie: Messagerie):
        courriels_actuels = self.df
        nouveaux_courriels = messagerie.df.fillna('')

        lim_db = 1000
        nouveaux_courriels.a = nouveaux_courriels.a.map(
            lambda x: x[:lim_db])
        nouveaux_courriels.sujet = nouveaux_courriels.sujet.map(
            lambda x: x[:lim_db])
        nouveaux_courriels.chaine = nouveaux_courriels.chaine.map(
            lambda x: x[:lim_db])

        tous_courriels = pandas.concat([courriels_actuels,
                                        nouveaux_courriels])
        nouveaux_courriels = tous_courriels.drop_duplicates(keep=False)

        self.append(nouveaux_courriels)
