#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - version et depot.

Un seul endroit ou la version est ecrite : l'interface, le verificateur de
mise a jour et l'etiquette de publication la lisent tous ici.

Convention de numerotation : sauf demande contraire, chaque mise a jour
incremente le dernier chiffre (0.2 -> 0.2.1 -> 0.2.2). Les changements de
comportement notables passent au chiffre du milieu.
"""

VERSION = "0.2"

DEPOT = "Trano89/Ollook"
API_PUBLICATIONS = "https://api.github.com/repos/%s/releases/latest" % DEPOT
PAGE_PUBLICATIONS = "https://github.com/%s/releases" % DEPOT

# Nom attendu du fichier joint a une publication.
NOM_EXECUTABLE = "Ollook.exe"


def tuple_version(texte):
    """'v0.2.1' -> (0, 2, 1). Les parties non numeriques valent 0.

    Permet de comparer deux versions sans dependre du nombre de segments :
    0.2 est bien anterieur a 0.2.1."""
    morceaux = (texte or "").strip().lstrip("vV").split(".")
    nombres = []
    for morceau in morceaux:
        chiffres = ""
        for caractere in morceau:
            if not caractere.isdigit():
                break
            chiffres += caractere
        nombres.append(int(chiffres) if chiffres else 0)
    while len(nombres) < 3:
        nombres.append(0)
    return tuple(nombres[:3])


def plus_recente(candidate, reference=VERSION):
    """Vrai si `candidate` est strictement posterieure a `reference`."""
    return tuple_version(candidate) > tuple_version(reference)
