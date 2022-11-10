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
import re

import email.message
import email.parser
import email.policy

from email.message import EmailMessage
from pathlib import Path
from functools import partial
from datetime import datetime
from imaplib import IMAP4_SSL
from typing import Any
from collections import namedtuple

import getpass
import quopri
import dateutil.parser
import chardet
import schedule
import pandas
import keyring

import numpy as np

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

# https://stackoverflow.com/questions/73941386/escaped-characters-of-form-aok-for-%c3%a9-in-imap4-list-response-using-imap4-s/73941387#73941387


def modified_base64(s: str) -> str:
    s_utf7 = s.encode('utf-7')  #
    aaa = s_utf7[1:-1].decode().replace('/', ',')  # rfc2060
    return aaa


def modified_unbase64(s: str) -> str:
    s_utf7 = '+' + s.replace(',', '/') + '-'
    return s_utf7.encode().decode('utf-7')


def encode_imap4_utf7(s: str, errors=None) -> tuple[str, int]:
    r = list()
    _in = list()
    for c in s:
        if ord(c) in range(0x20, 0x25) or ord(c) in range(0x27, 0x7e):
            if _in:
                r.extend(['&', modified_base64(''.join(_in)), '-'])
                del _in[:]
            r.append(str(c))
        elif ord(c) == 0x26:
            if _in:
                r.extend(['&', modified_base64(''.join(_in)), '-'])
                del _in[:]
            r.append('&-')
        else:
            _in.append(c)
    if _in:
        r.extend(['&', modified_base64(''.join(_in)), '-'])
    return ''.join(r), len(s)


def decode_imap4_utf7(s: str) -> str:
    r = list()
    if s.find('&-') != -1:
        s = s.split('&-')
        i = len(s)
        for subs in s:
            i -= 1
            r.append(decode_imap4_utf7(subs))
            if i != 0:
                r.append('&')
    else:
        regex = re.compile(r'[&]\S+?[-]')
        sym = re.split(regex, s)
        if len(regex.findall(s)) > 1:
            i = 0
            r.append(sym[i])
            for subs in regex.findall(s):
                r.append(decode_imap4_utf7(subs))
                i += 1
                r.append(sym[i])
        elif len(regex.findall(s)) == 1:
            r.append(sym[0])
            r.append(modified_unbase64(regex.findall(s)[0][1:-1]))
            r.append(sym[1])
        else:
            r.append(s)
    return ''.join(r)


class CourrielsConfig(FichierConfig):

    def default(self) -> str:
        return CONFIGURATION_PAR_DÉFAUT


PièceJointe = namedtuple('PièceJointe', ['nom', 'type_MIME', 'contenu'])


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

        if destinataire is not None:
            self.destinataire = destinataire
        if expéditeur is not None:
            self.expéditeur = expéditeur
        if objet is not None:
            self.objet = objet
        if contenu is not None:
            self.contenu = contenu
        if html is not None:
            self.html = html

        if pièces_jointes:
            self.joindre(*pièces_jointes)

    def __getitem__(self, clé: Any) -> Any:
        return self.message[clé]

    def __setitem__(self, clé: Any, val: Any) -> Any:
        self.message[clé] = val

    def __getattr__(self, clé: str) -> Any:
        if clé in self.équivalences_attributs:
            clé = self.équivalences_attributs[clé]
            return self[clé]
        elif hasattr(EmailMessage, clé):
            return getattr(self.message, clé)

    def __setattr__(self, clé: str, val: Any) -> Any:
        if clé == 'contenu':
            self.message.set_content(val)
        elif clé == 'html':
            self.message.add_alternative(val, subtype='html')
        elif clé in self.équivalences_attributs:
            clé = self.équivalences_attributs[clé]
            self[clé] = val
        elif hasattr(EmailMessage, clé):
            setattr(self.message, clé, val)
        else:
            super().__setattr__(clé, val)

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

    @property
    def pièces_jointes(self):
        if self.is_multipart() and len(self.get_payload()) > 1:
            for pj in self.get_payload():
                if pj.get('Content-Disposition', '').startswith('attachment'):
                    nom = pj.get_filename()
                    content_type = pj.get_content_type()
                    pj.set_charset('utf-8')
                    contenu = pj.get_payload(decode=True)
                    yield PièceJointe(nom, content_type, contenu)

    def envoyer(self, adresse, port=25):
        serveur = smtplib.SMTP(adresse, port)
        serveur.send_message(self.message)
        serveur.quit()

    @property
    def date(self) -> datetime:
        if isinstance(self['Date'], str):
            return dateutil.parser.parse(self['Date'], ignoretz=True)
        else:
            return None

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


BoîteAuxLettres = namedtuple('BoîteAuxLettres', ['est_parent',
                                                 'sep',
                                                 'nom'])


class Messagerie:

    def __init__(self, config: CourrielsConfig):
        if isinstance(config, (str, Path)):
            self.config = CourrielsConfig(config)
        else:
            self.config = config

        self._mdp = None
        self.sélection = 'INBOX'

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

        return Courriel(message=message)

    def messages(self, *recherche) -> Courriel:
        if not recherche:
            recherche = ('ALL',)
        with self.connecter() as serveur:
            serveur.select(self.sélection)
            typ, data = serveur.search(None, *recherche)
            messages: list[str] = data[0].split()
            f = partial(self.message, serveur)

            yield from map(f, messages)

    def __iter__(self):
        return self.messages()

    def boîtes(self):
        with self.connecter() as serveur:
            état, boîtes = serveur.list()

        yield from (BoîteAuxLettres(b.split(')', 1)[0].strip('('),
                                    b.split(') "', 1)[1].split('" "', 1)[0],
                                    decode_imap4_utf7(
                                        b.split('" "', 1)[1].strip('"')))
                    for b in
                    (decode_imap4_utf7(s) for s in
                     (str(b, encoding='utf-7') for b in boîtes)))

    def select(self, boîte: BoîteAuxLettres):
        if isinstance(boîte, str):
            for b in self.boîtes():
                if b.nom == boîte:
                    boîte = b
            if isinstance(boîte, str):
                raise ValueError('Cette boîte aux lettres n\'existe pas.')

        nom, l = encode_imap4_utf7(boîte.nom)
        self.sélection = '"{}"'.format(nom)

    @ property
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

    def connecter(self):
        serveur = IMAP4_SSL(self.adresse)
        serveur.login(self.nom, self.mdp)
        serveur.enable('UTF-8=ACCEPT')
        return serveur


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
        nouveaux_courriels.contenu = nouveaux_courriels.contenu.map(
            partial(bytes, encoding='utf-8'))

        tous_courriels = pandas.concat([courriels_actuels,
                                        nouveaux_courriels])

        nouveaux_courriels = tous_courriels.drop_duplicates(('date',
                                                             'de',
                                                            'a',
                                                             'sujet'),
                                                            keep=False)\
            .replace({np.nan: None})

        self.màj(nouveaux_courriels)
