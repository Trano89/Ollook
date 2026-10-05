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

VERSION = "0.2.7"

DEPOT = "Trano89/Ollook"
import os as _os

# Surchargeable pour eprouver le mecanisme de mise a jour sans publier :
# le verificateur interroge alors un serveur local au lieu de GitHub.
API_PUBLICATIONS = _os.environ.get(
    "OLLOOK_API_PUBLICATIONS",
    "https://api.github.com/repos/%s/releases/latest" % DEPOT)
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


def plus_recente(candidate, reference=None):
    """Vrai si `candidate` est strictement posterieure a `reference`.

    La reference est lue a l'appel, pas figee a l'import : un defaut evalue
    une seule fois rendrait la fonction insensible a toute modification de
    VERSION, et donc intestable."""
    if reference is None:
        reference = VERSION
    return tuple_version(candidate) > tuple_version(reference)
