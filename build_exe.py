#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - fabrication de l'executable autonome.

Produit dist/Ollook.exe : un seul fichier, sans Python a installer, sans
dependance externe. L'executable embarque l'interpreteur, tkinter et le code
d'Ollook, et sait installer Ollama et un modele a son premier lancement.

    python build_exe.py

Seule dependance de fabrication : PyInstaller, installe automatiquement s'il
manque. Elle n'est requise que pour construire, jamais pour executer.
"""

import os
import struct
import subprocess
import sys

RACINE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(RACINE, "assets")      # PNG statiques de l'application
ICONE = os.path.join(RACINE, "build", "ollook.ico")
NOM = "Ollook"


def dire(message, gras=False):
    print(("\n== %s" % message) if gras else ("   %s" % message))


# --------------------------------------------------------------------------
# Icone Windows : un .ico est un conteneur, on y empile les PNG existants.
# --------------------------------------------------------------------------

def fabriquer_ico(sources, destination):
    entrees, donnees = [], b""
    decalage = 6 + 16 * len(sources)

    for chemin in sources:
        with open(chemin, "rb") as f:
            blob = f.read()
        # Dimensions lues dans le bloc IHDR du PNG.
        largeur, hauteur = struct.unpack(">II", blob[16:24])
        entrees.append(struct.pack(
            "<BBBBHHII",
            0 if largeur >= 256 else largeur,    # 0 signifie 256
            0 if hauteur >= 256 else hauteur,
            0, 0, 1, 32, len(blob), decalage))
        decalage += len(blob)
        donnees += blob

    os.makedirs(os.path.dirname(destination), exist_ok=True)
    with open(destination, "wb") as f:
        f.write(struct.pack("<HHH", 0, 1, len(sources)))   # ICONDIR
        for entree in entrees:
            f.write(entree)
        f.write(donnees)
    return destination


def preparer_icone():
    sources = [os.path.join(ASSETS, "icon-%d.png" % t)
               for t in (16, 32, 64, 128, 256)]
    sources = [s for s in sources if os.path.exists(s)]
    if not sources:
        dire("Aucune icône dans assets/, l'exécutable en sera dépourvu.")
        return None
    fabriquer_ico(sources, ICONE)
    dire("Icône : %s (%d tailles)" % (ICONE, len(sources)))
    return ICONE


# --------------------------------------------------------------------------
# PyInstaller
# --------------------------------------------------------------------------

def assurer_pyinstaller():
    try:
        import PyInstaller  # noqa: F401
        dire("PyInstaller déjà présent.")
        return True
    except ImportError:
        pass

    dire("Installation de PyInstaller…")
    code = subprocess.call([sys.executable, "-m", "pip", "install",
                            "--disable-pip-version-check", "pyinstaller"])
    if code != 0:
        dire("Échec de l'installation de PyInstaller.")
        return False
    return True


def construire(icone):
    commande = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",              # un seul fichier, rien a deployer
        "--windowed",             # application graphique, pas de console
        "--name", NOM,
        "--distpath", os.path.join(RACINE, "dist"),
        "--workpath", os.path.join(RACINE, "build", "travail"),
        "--specpath", os.path.join(RACINE, "build"),
        # Modules locaux : importes normalement, mais on les declare pour que
        # l'analyse statique ne puisse pas les manquer.
        "--hidden-import", "ollook_core",
        "--hidden-import", "ollook_setup",
        "--hidden-import", "ollook_icon",
        "--hidden-import", "ollook_outlook",
        "--hidden-import", "ollook_version",
        "--hidden-import", "ollook_update",
        "--paths", RACINE,
    ]
    if icone:
        commande += ["--icon", icone]
    commande.append(os.path.join(RACINE, "ollook_app.py"))

    dire("Compilation en cours, comptez une à deux minutes…")
    return subprocess.call(commande) == 0


def main():
    dire("Ollook — fabrication de l'exécutable", gras=True)

    if not assurer_pyinstaller():
        return 1

    icone = preparer_icone()
    if not construire(icone):
        dire("La compilation a échoué.")
        return 1

    exe = os.path.join(RACINE, "dist", NOM + (".exe" if os.name == "nt" else ""))
    if not os.path.exists(exe):
        dire("Exécutable introuvable après compilation.")
        return 1

    # Les intermediaires pesent une quinzaine de mega-octets et se
    # reconstruisent en une minute : inutile de les garder.
    import shutil
    for reste in (os.path.join(RACINE, "build", "travail"),
                  os.path.join(RACINE, "__pycache__")):
        shutil.rmtree(reste, ignore_errors=True)

    taille = os.path.getsize(exe) / (1024.0 * 1024.0)
    dire("Terminé", gras=True)
    dire("%s  (%.1f Mo)" % (exe, taille))
    dire("Copiable et exécutable tel quel, sans Python installé.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
