#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - mise a jour depuis GitHub.

Interroge la derniere publication du depot, et sait remplacer l'executable
la ou il se trouve : chacun range Ollook ou il veut, la mise a jour suit.

Un programme en cours d'execution ne peut pas s'ecraser lui-meme sous
Windows. Le nouvel executable est donc telecharge a cote, puis un court
script prend le relais : il attend la fin du processus, echange les fichiers,
relance Ollook et s'efface.

Aucune dependance pip.
"""

import json
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request

import ollook_version as version

WINDOWS = sys.platform.startswith("win")

DELAI = 8          # secondes : la verification ne doit jamais retarder l'ouverture


def _dossier_application():
    """Dossier de l'executable -- celui que la mise a jour doit remplacer."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def verifier():
    """Renvoie les informations de la publication si une version plus recente
    existe, sinon None. Toute erreur -- reseau coupe, depot injoignable, quota
    d'API -- se traduit par None : une mise a jour n'est jamais bloquante."""
    requete = urllib.request.Request(
        version.API_PUBLICATIONS,
        headers={"Accept": "application/vnd.github+json",
                 "User-Agent": "Ollook/%s" % version.VERSION})
    try:
        with urllib.request.urlopen(requete, timeout=DELAI) as reponse:
            publication = json.loads(reponse.read().decode("utf-8"))
    except (urllib.error.URLError, ValueError, OSError):
        return None

    etiquette = publication.get("tag_name") or publication.get("name") or ""
    if not version.plus_recente(etiquette):
        return None

    lien = None
    for piece in publication.get("assets") or []:
        if (piece.get("name") or "").lower() == version.NOM_EXECUTABLE.lower():
            lien = piece.get("browser_download_url")
            break

    return {
        "version": etiquette.lstrip("vV"),
        "lien": lien,
        "notes": (publication.get("body") or "").strip(),
        "page": publication.get("html_url") or version.PAGE_PUBLICATIONS,
    }


def telechargeable():
    """Vrai si Ollook peut se remplacer lui-meme.

    Lance depuis les sources, il n'y a pas d'executable a echanger : on se
    contente alors d'ouvrir la page des publications."""
    return WINDOWS and getattr(sys, "frozen", False)


def telecharger(lien, on_progress=None):
    """Telecharge le nouvel executable a cote de l'ancien. Renvoie son chemin."""
    cible = os.path.join(_dossier_application(), "Ollook-nouveau.exe")
    requete = urllib.request.Request(
        lien, headers={"User-Agent": "Ollook/%s" % version.VERSION})
    with urllib.request.urlopen(requete, timeout=60) as reponse:
        total = int(reponse.headers.get("Content-Length") or 0)
        recu = 0
        with open(cible, "wb") as sortie:
            while True:
                morceau = reponse.read(262144)
                if not morceau:
                    break
                sortie.write(morceau)
                recu += len(morceau)
                if on_progress:
                    on_progress(recu, total)

    # Un executable PyInstaller commence par l'en-tete MZ : un fichier
    # tronque ou une page d'erreur HTML seraient sinon installes tels quels.
    with open(cible, "rb") as f:
        entete = f.read(2)
    # Suppression APRES fermeture : Windows refuse d'effacer un fichier ouvert,
    # et l'erreur masquerait le vrai motif du rejet.
    if entete != b"MZ":
        try:
            os.remove(cible)
        except OSError:
            pass
        raise ValueError("le fichier telecharge n'est pas un executable")
    return cible


_RELAIS = r'''
$ErrorActionPreference = 'SilentlyContinue'

# Attendre la fermeture d'Ollook : un executable en cours ne peut pas etre
# remplace tant qu'il tient le fichier.
$fin = (Get-Date).AddSeconds(30)
while ((Get-Date) -lt $fin) {
    if (-not (Get-Process -Id $env:OLLOOK_PID)) { break }
    Start-Sleep -Milliseconds 300
}
Start-Sleep -Milliseconds 400

Move-Item -LiteralPath $env:OLLOOK_NOUVEAU -Destination $env:OLLOOK_CIBLE -Force
if ($?) {
    Start-Process -FilePath $env:OLLOOK_CIBLE -WorkingDirectory (Split-Path $env:OLLOOK_CIBLE)
}
Remove-Item -LiteralPath $PSCommandPath -Force
'''


def installer(nouveau):
    """Confie l'echange a un script externe, puis rend la main.

    L'appelant doit fermer l'application immediatement apres : le script
    attend justement sa disparition."""
    cible = os.path.abspath(sys.executable)
    script = os.path.join(tempfile.gettempdir(), "ollook-maj.ps1")
    with open(script, "w", encoding="utf-8") as f:
        f.write(_RELAIS)

    env = dict(os.environ)
    env["OLLOOK_PID"] = str(os.getpid())
    env["OLLOOK_NOUVEAU"] = nouveau
    env["OLLOOK_CIBLE"] = cible

    subprocess.Popen(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-WindowStyle", "Hidden", "-File", script],
        env=env, creationflags=0x08000000 if WINDOWS else 0)
    return True


def ouvrir_page(page=None):
    """Ouvre la page des publications dans le navigateur."""
    import webbrowser
    webbrowser.open(page or version.PAGE_PUBLICATIONS)
