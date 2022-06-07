# -*- coding: utf-8 -*-
"""Envoi facile de courriels."""

# Bibliothèque standard
import mimetypes
import smtplib

from email.message import EmailMessage
from pathlib import Path


class Courriel:

    def __init__(self, destinataire, expéditeur, objet, contenu, html=None, pièces_jointes=tuple()):
        self.destinataire = destinataire
        self.expéditeur = expéditeur
        self.objet = objet
        self.contenu = contenu
        self.html = html

        self.pièces_jointes = set()
        self.joindre(*pièces_jointes)

    def construire(self):
        self.message = EmailMessage()
        self.message['From'] = self.expéditeur
        self.message['To'] = self.destinataire
        self.message['Subject'] = self.objet

        self.message.set_content(self.contenu)

        if self.html is not None:
            self.message.add_alternative(self.html, subtype='html')

        for pj in self.pièces_jointes:
            type_mime = mimetypes.guess_type(pj.name)

            if None in type_mime:
                type_mime = ('application', 'octet-stream')

            with pj.open('rb') as f:
                self.message.add_attachment(f.read(),
                                            maintype=type_mime[0],
                                            subtype=type_mime[1],
                                            filename=pj.name)

    def joindre(self, *chemins):
        for chemin in chemins:
            chemin = Path(chemin)
            self.pièces_jointes.add(chemin)

    def envoyer(self, adresse, port=25):
        self.construire()
        serveur = smtplib.SMTP(adresse, port)
        serveur.send_message(self.message)
        serveur.quit()
