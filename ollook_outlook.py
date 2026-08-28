#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - acces direct au corps du message via Outlook.

Le pilotage au clavier (Ctrl+A, Ctrl+V) est approximatif : il depend du
comportement de l'editeur, et un collage integral detruit la mise en forme de
la signature et du fil cite. Ce module passe par le modele objet d'Outlook et
son editeur Word, ce qui permet de remplacer une PLAGE DE CARACTERES exacte :
tout ce qui est en dehors -- signature, fil de discussion, mise en forme --
reste rigoureusement intact.

Le dialogue se fait par PowerShell, seul moyen d'atteindre COM sans dependance
externe. Les textes transitent par des fichiers UTF-8, ce qui evite toute
question d'echappement et de page de code.

Windows et Outlook classique uniquement ; l'appelant retombe sur le
presse-papiers quand ce module n'est pas disponible.
"""

import os
import subprocess
import sys
import tempfile

WINDOWS = sys.platform.startswith("win")

# Outlook represente une fin de paragraphe par un seul retour chariot.
MARQUE_PARAGRAPHE = "\r"

# Word n'utilise pas que \r. Un <br> -- ce dont sont faites la plupart des
# signatures -- devient une rupture de ligne \x0b, invisible pour un simple
# decoupage sur \r : sans cette table, les lignes d'une signature se retrouvent
# collees et la signature n'est pas reconnue.
# Toutes ces substitutions sont d'un caractere pour un caractere : les
# positions restent donc valables cote Outlook.
RUPTURES = {
    "\x0b": "\n",   # rupture de ligne (Maj+Entree, <br>)
    "\x07": "\n",   # fin de cellule ou de ligne de tableau
    "\x0c": "\n",   # saut de page
    "\r": "\n",     # fin de paragraphe
}


def normaliser(texte):
    """Ramene toutes les ruptures de Word a '\\n', sans changer les longueurs."""
    for marque, remplacement in RUPTURES.items():
        texte = texte.replace(marque, remplacement)
    return texte

_PREAMBULE = r'''
$ErrorActionPreference = 'Stop'

function Get-OllookEditor {
    try {
        $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
    } catch {
        return $null
    }

    # Fenetre de redaction ouverte.
    try {
        $insp = $ol.ActiveInspector()
        if ($insp) {
            $doc = $insp.WordEditor
            if ($doc) { return $doc }
        }
    } catch { }

    # Reponse en ligne, redigee dans le volet de lecture.
    try {
        $exp = $ol.ActiveExplorer()
        if ($exp) {
            $doc = $exp.ActiveInlineResponseWordEditor
            if ($doc) { return $doc }
        }
    } catch { }

    return $null
}
'''

_LIRE = _PREAMBULE + r'''
$doc = Get-OllookEditor
if (-not $doc) { exit 2 }
$texte = $doc.Range().Text
if ($null -eq $texte) { $texte = '' }
# UTF8Encoding($false) : SANS marque d'ordre des octets. Le BOM ecrit par
# defaut ajouterait un caractere au debut du texte et decalerait d'autant
# toutes les positions -- donc le remplacement.
$sansBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($env:OLLOOK_FICHIER, $texte, $sansBom)
exit 0
'''

_REMPLACER = _PREAMBULE + r'''
$doc = Get-OllookEditor
if (-not $doc) { exit 2 }

$nouveau = [System.IO.File]::ReadAllText($env:OLLOOK_FICHIER, [System.Text.Encoding]::UTF8)
$attendu = [System.IO.File]::ReadAllText($env:OLLOOK_TEMOIN, [System.Text.Encoding]::UTF8)
$fin = [int]$env:OLLOOK_FIN

$total = $doc.Range().End
if ($fin -lt 0 -or $fin -gt $total) { exit 3 }

$plage = $doc.Range(0, $fin)

# L'utilisateur a pu continuer a ecrire entre la lecture et le remplacement :
# on ne remplace que si la plage contient toujours ce qu'on croit.
$actuel = $plage.Text
if ($null -eq $actuel) { $actuel = '' }
$sansEspace = { param($t) ($t -replace '\s', '') }
if ((& $sansEspace $actuel) -ne (& $sansEspace $attendu)) { exit 4 }

# Seuls les caracteres 0..$fin sont remplaces. Tout ce qui suit -- signature,
# fil cite, mise en forme -- n'est pas touche.
$plage.Text = $nouveau
exit 0
'''


def disponible():
    return WINDOWS


def _executer(script, fichier, variables=None):
    """Execute un script PowerShell, sans fenetre de console."""
    dossier = tempfile.mkdtemp(prefix="ollook-")
    chemin = os.path.join(dossier, "op.ps1")
    try:
        with open(chemin, "w", encoding="utf-8") as f:
            f.write(script)

        env = dict(os.environ)
        env["OLLOOK_FICHIER"] = fichier
        for cle, valeur in (variables or {}).items():
            env[cle] = str(valeur)

        creation = 0x08000000 if WINDOWS else 0          # CREATE_NO_WINDOW
        resultat = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive",
             "-ExecutionPolicy", "Bypass", "-File", chemin],
            env=env, creationflags=creation,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
        return resultat.returncode
    except (OSError, subprocess.SubprocessError):
        return -1
    finally:
        for reste in (chemin, dossier):
            try:
                os.remove(reste) if os.path.isfile(reste) else os.rmdir(reste)
            except OSError:
                pass


def lire_corps():
    """Corps du message en cours de redaction, ou None.

    Les fins de paragraphe sont normalisees en '\\n' : un caractere pour un
    caractere, de sorte que les positions restent valables cote Outlook."""
    if not WINDOWS:
        return None

    poignee, fichier = tempfile.mkstemp(prefix="ollook-", suffix=".txt")
    os.close(poignee)
    try:
        if _executer(_LIRE, fichier) != 0:
            return None
        # utf-8-sig : retire une eventuelle marque d'ordre des octets, qui
        # n'existe pas cote Outlook et fausserait les positions.
        with open(fichier, encoding="utf-8-sig") as f:
            texte = f.read()
    except OSError:
        return None
    finally:
        try:
            os.remove(fichier)
        except OSError:
            pass

    return normaliser(texte)


def remplacer_debut(fin, texte, temoin):
    """Remplace les `fin` premiers caracteres du corps par `texte`.

    `fin` est une position dans le texte renvoye par lire_corps().
    `temoin` est ce que cette plage doit contenir : le remplacement est
    abandonne si elle a change entre-temps."""
    if not WINDOWS:
        return False

    poignee, fichier = tempfile.mkstemp(prefix="ollook-", suffix=".txt")
    os.close(poignee)
    poignee2, fichier_temoin = tempfile.mkstemp(prefix="ollook-", suffix=".txt")
    os.close(poignee2)
    try:
        with open(fichier, "w", encoding="utf-8") as f:
            f.write(texte.replace("\n", MARQUE_PARAGRAPHE))
        with open(fichier_temoin, "w", encoding="utf-8") as f:
            f.write(temoin)
        return _executer(_REMPLACER, fichier,
                         {"OLLOOK_FIN": int(fin),
                          "OLLOOK_TEMOIN": fichier_temoin}) == 0
    except OSError:
        return False
    finally:
        for reste in (fichier, fichier_temoin):
            try:
                os.remove(reste)
            except OSError:
                pass


# --------------------------------------------------------------------------
# Signatures configurees dans Outlook classique
#
# Outlook y depose trois formats par signature ; le .txt est la version texte
# brut, celle qui se compare au corps du message. Connaitre la vraie signature
# rend la frontiere du brouillon EXACTE, la ou l'heuristique doit deviner.
#
# Le nouvel Outlook, lui, ne depose rien : ses signatures vivent dans la boite
# aux lettres. Le dossier est alors vide et l'appelant retombe sur l'heuristique.
# --------------------------------------------------------------------------

def dossier_signatures():
    base = os.environ.get("APPDATA")
    return os.path.join(base, "Microsoft", "Signatures") if base else None


def _decoder(brut):
    """Les fichiers sont en UTF-16 avec marque d'ordre des octets."""
    for encodage in ("utf-16", "utf-8-sig", "cp1252"):
        try:
            texte = brut.decode(encodage)
        except (UnicodeDecodeError, LookupError):
            continue
        if "\x00" not in texte:
            return texte
    return ""


def signatures_connues():
    """Textes des signatures deposees par Outlook, la plus longue d'abord.

    On les essaie toutes plutot que de consulter le registre pour savoir
    laquelle appartient a quel compte : c'est plus court, et cela couvre les
    profils a plusieurs comptes sans risque de se tromper de correspondance."""
    dossier = dossier_signatures()
    if not dossier or not os.path.isdir(dossier):
        return []

    textes = []
    try:
        noms = os.listdir(dossier)
    except OSError:
        return []

    for nom in noms:
        if not nom.lower().endswith(".txt"):
            continue
        try:
            with open(os.path.join(dossier, nom), "rb") as f:
                brut = f.read()
        except OSError:
            continue
        texte = _decoder(brut).strip()
        if texte:
            textes.append(texte)

    # La plus longue d'abord : entre deux signatures dont l'une est le prefixe
    # de l'autre, on veut reconnaitre la plus complete.
    textes.sort(key=len, reverse=True)
    return textes
