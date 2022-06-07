# -*- coding: utf-8 -*-
"""
Enveloppe pour les fichiers de configuration.

La classe FichierConfig permet de garder un fichier de configuration
synchronisé quand des modifications y sont faites dans le programme.
"""

# Bibliothèque standard
from io import StringIO  # Pour FichierConfig.__str__
from pathlib import Path  # Manipulation de chemins
# Pour parsage d'urls, utilisé dans FichierConfig.__init__
from urllib.parse import urlparse
from typing import Any, Callable
from functools import partial

# configparser contient la classe ConfigParser qu'on surclasse ici.
# Certaines constantes de configuration sont aussi importées.
from configparser import ConfigParser, _UNSET, DEFAULTSECT, _default_dict


class FichierConfig(ConfigParser):
    """Garde ConfigParser synchronisé avec un fichier."""

    def __init__(self,
                 chemin: Path,
                 defaults: dict = None,
                 dict_type: type = _default_dict,
                 allow_no_value: bool = False,
                 *,
                 fichier_defaut: Path = Path(__file__).parent / 'default.cfg',
                 delimiters: tuple[str] = ('=', ':'),
                 comment_prefixes: tuple[str] = ('#', ';'),
                 inline_comment_prefixes: tuple[str] = None,
                 strict: bool = True,
                 empty_lines_in_values: bool = True,
                 default_section: str = DEFAULTSECT,
                 interpolation: type = _UNSET,
                 converters: dict[str, Callable] = _UNSET):
        """
        Garde un fichier synchronisé avec ConfigParser.

        FichierConfig peut être utilisé exactement comme ConfigParser, mais
        est spécialement destinée à des cas où des réglages doivent être
        sauvegardés et sont modifiés souvent.

        Dans le cadre d'une utilisation simple, on peut surclasser
        FichierConfig et redéfinir la méthode default ou forcer la valeur de
        l'argument fichier_defaut.

        Parameters
        ----------
        chemin: str
            Le chemin du fichier de configuration.
        defaults: dict, optionel
            Les valeurs par défaut. Correspond à la section DEFAULT du
            fichier de configuration. La valeur par défaut est None.
        dict_type: type, optionel
            Le type de dictionnaire à utiliser comme type sous-jacent à
            ConfigParser. La valeur par défaut est dict.
        allow_no_value: bool, optionel
            Détermine si une option peut être laissée sans valeur.
            La valeur par défaut est False.
        fichier_defaut: Path, optionel
            Le fichier contenant le modèle de fichier par défaut.
            La valeur par défaut désigne le fichier `default.cfg` inclut
            dans le module.
        delimiters: tuple[str], optionel
            Les caractères servant à délimiter les options et leur valeurs.
            Par défaut, les caractères ':' et '=' sont utilisés.
        comment_prefixes: tuple[str], optionel
            Les caractères permettant d'insérer des blocs de commentaires.
            Par défaut, les caractères '#' et ';' sont utilisés.
        inline_comment_prefixes: tuple[str], optionel
            Les caractères permettant d'insérer des commentaires en fin
            de ligne. Par défaut, ils ne sont pas permis.
        strict: bool, optionel
            Détermine le comportement du parseur.
            La valeur par défaut est True.
        empty_lines_in_values: bool, optionel
            Détermine si les lignes vides sont gardées dans les valeurs
            multiligne. La valeur par défaut est True.
        default_section : str, optionel
            Le nom de la section par défaut, utilisée pour les valeurs
            par défaut. La valeur par défaut est 'DEFAULT'.
        interpolation : type, optionel
            Classe d'interpolation à utiliser.
            La valeur par défaut correspond à BasicInterpolation.
        converters : dict[str, Callable], optionel
            Les convertisseurs à utiliser. Un dictionnaire en argument ayant
            la clé 'test' associée à la fonction `test_func` permettra
            d'utiliser la méthode `gettest` qui retournera le résultat de
            `test_func`. La valeur par défaut inclut 'int', 'float', 'boolean',
            'list', 'path' et 'url'.

        Returns
        -------
        None.

        """
        # Principale différence avec ConfigParser:
        # l'attribut chemin réfère au fichier de configuration
        # lu et écrit dans le programme.
        self.chemin: Path = Path(chemin)

        # Si le fichier n'existe pas, il est créé,
        # et on lui donne la valeur par défaut,
        # telle que définie par la méthode défaut
        self.fichier_defaut: Path = fichier_defaut
        if not self.chemin.exists():
            self.chemin.touch()
            with self.chemin.open('w') as f:
                f.write(self.default())

        # Certaines conversions sont utiles dans beaucoup de cas
        # donc on s'assure de pouvoir convertir les listes, chemins et urls.
        # On garde quand même ça simple pour quelqu'un de remplacer
        # les convertisseurs par défaut.
        if converters is _UNSET:
            converters = {}
        if isinstance(converters, dict):
            if 'list' not in converters:
                converters['list'] = lambda x: list(
                    map(
                        str.strip,
                        x.strip().split('\n')
                    )
                )
            if 'path' not in converters:
                converters['path'] = lambda x: Path(x).expanduser().resolve()
            if 'url' not in converters:
                converters['url'] = urlparse
            if 'bytes' not in converters:
                converters['bytes'] = partial(bytes, encoding='utf-8')

        # On crée l'objet ConfigParser sous-jacent.
        super().__init__(defaults,
                         dict_type,
                         allow_no_value,
                         delimiters=delimiters,
                         comment_prefixes=comment_prefixes,
                         inline_comment_prefixes=inline_comment_prefixes,
                         strict=strict,
                         empty_lines_in_values=empty_lines_in_values,
                         default_section=default_section,
                         interpolation=interpolation,
                         converters=converters)

        # Et on synchronise une première fois.
        self.read()

    def optionxform(self, option: str) -> str:
        """
        Formater une option.

        Pour assurer l'uniformité et la facilité
        d'écriture et d'utilisation.

        Parameters
        ----------
        option : str
            Nom à formater.

        Returns
        -------
        str
            Le nom, formaté. Par défaut, rien n'est changé.

        """
        return str(option)

    def default(self) -> str:
        """
        Retourner le contenu du fichier de configuration par défaut.

        Un minimum de formatage peut y être apporté via la méthode
        `str.format`, par défaut seul `self` est passée en argument
        pour le formatage.

        Returns
        -------
        str
            Contenu par défaut.

        """
        return self.fichier_defaut.open().read().format(self=self)

    def read(self, encoding: str = 'utf-8'):
        """
        Lire le contenu du fichier.

        Parameters
        ----------
        encoding : str, optional
            Encodage de lecture. The default is 'utf-8'.

        Returns
        -------
        None.

        """
        super().read(self.chemin, encoding=encoding)

    def write(self, space_around_delimiters: bool = True):
        """
        Écrire le fichier de configuration à self.chemin.

        Parameters
        ----------
        space_around_delimiters : bool, optional
            Détermine si des espaces sont placés autour des séparateurs.
            The default is True.

        Returns
        -------
        None.

        """
        with self.chemin.open('w') as f:
            super().write(f, space_around_delimiters)

    def __delitem__(self, section: str):
        """
        Retirer une section.

        Parameters
        ----------
        section : str
            Section à retirer.

        Returns
        -------
        None.

        """
        super().__delitem__(section)
        self.write()

    def __setitem__(self, section: str, value: Any):
        """
        Modifier une section.

        Parameters
        ----------
        section : str
            Section à redéfinir.
        value : Any
            Valeur.

        Returns
        -------
        None.

        """
        super().__setitem__(section, value)
        self.write()

    def set(self, section: str, option: str, value: Any = None):
        """
        Modifier une option.

        Parameters
        ----------
        section : str
            Section contenant l'option.
        option : str
            Option à modifier.
        value : Any, optional
            Valeur. The default is None.

        Returns
        -------
        None.

        """
        super().set(section, option, value)
        self.write()

    def add_section(self, section: str):
        """
        Ajouter une section.

        Parameters
        ----------
        section : str
            Nom de la section à ajouter.

        Returns
        -------
        None.

        """
        super().add_section(section)
        self.write()

    def remove_section(self, section: str):
        """
        Retirer une section.

        Parameters
        ----------
        section : str
            Section à retirer.

        Returns
        -------
        None.

        """
        super().remove_section(section)
        self.write()

    def remove_option(self, section: str, option: str):
        """
        Retirer une option.

        Parameters
        ----------
        section : str
            Section contenant l'option.
        option : str
            Option à retirer.

        Returns
        -------
        None.

        """
        super().remove_option(section, option)
        self.write()

    def __str__(self) -> str:
        """
        Retourner le fichier de configuration comme une chaîne.

        Returns
        -------
        str
            Contenu du fichier de configuration.

        """
        # Simuler un fichier avec StringIO
        with StringIO() as fp:
            super().write(fp)
            rés: str = fp.getvalue()

        return rés
