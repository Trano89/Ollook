#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - mise en route.

Verifie, et installe au besoin, ce dont Ollook a besoin pour fonctionner :

  1. Ollama present sur la machine ;
  2. Ollama demarre ;
  3. au moins un modele installe -- a defaut, gemma4:12b.

Rien n'est telecharge sans accord explicite : l'interface demande confirmation
avant d'installer Ollama comme avant de tirer un modele, en annoncant la taille.

Aucune dependance pip.
"""

import os
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request

# Une fois embarque dans un executable, le dossier est deja sur le chemin.
if not getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ollook_core as core  # noqa: E402

WINDOWS = sys.platform.startswith("win")
MACOS = sys.platform == "darwin"

# Modele par defaut. Le 12b plutot qu'un e2b/e4b : en dessous d'environ quatre
# milliards de parametres, les modeles corrigent l'orthographe mais ignorent la
# consigne de registre -- le tutoiement en pro interne ne passe pas -- ce qui
# vide l'outil de son interet.
DEFAULT_MODEL = "gemma4:12b"
DEFAULT_MODEL_SIZE = "7,6 Go"

OLLAMA_WINDOWS_URL = "https://ollama.com/download/OllamaSetup.exe"
OLLAMA_MACOS_URL = "https://ollama.com/download/Ollama.dmg"
OLLAMA_PAGE = "https://ollama.com/download"


# --------------------------------------------------------------------------
# Diagnostic
# --------------------------------------------------------------------------

def find_ollama():
    """Chemin de l'executable ollama, ou None."""
    found = shutil.which("ollama")
    if found:
        return found
    candidates = [
        os.path.expandvars(r"%LOCALAPPDATA%\Programs\Ollama\ollama.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Ollama\ollama.exe"),
        "/usr/local/bin/ollama",
        "/opt/homebrew/bin/ollama",
        "/Applications/Ollama.app/Contents/Resources/ollama",
    ]
    for path in candidates:
        if path and os.path.exists(path):
            return path
    return None


def ollama_alive():
    try:
        core.ollama_version()
        return True
    except Exception:
        return False


def installed_models():
    try:
        return core.list_models()
    except Exception:
        return []


def diagnose():
    """Etat courant : (ollama_installe, ollama_demarre, nombre_de_modeles)."""
    binaire = find_ollama()
    vivant = ollama_alive()
    modeles = installed_models() if vivant else []
    return bool(binaire), vivant, len(modeles)


def ready():
    """Vrai si Ollook peut fonctionner immediatement."""
    _, vivant, nb = diagnose()
    return vivant and nb > 0


# --------------------------------------------------------------------------
# Actions
# --------------------------------------------------------------------------

def start_ollama(log=None, timeout=40):
    """Demarre le service Ollama et attend qu'il reponde."""
    binaire = find_ollama()
    if not binaire:
        return False
    if log:
        log("Démarrage d'Ollama…")

    try:
        if WINDOWS:
            # CREATE_NO_WINDOW : pas de fenetre de console qui traine.
            subprocess.Popen([binaire, "serve"], creationflags=0x08000000,
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            subprocess.Popen([binaire, "serve"],
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                             start_new_session=True)
    except OSError as e:
        if log:
            log("Échec du démarrage : %s" % e)
        return False

    fin = time.time() + timeout
    while time.time() < fin:
        if ollama_alive():
            if log:
                log("Ollama répond.")
            return True
        time.sleep(0.6)
    if log:
        log("Ollama n'a pas répondu dans le délai imparti.")
    return False


def download(url, destination, on_progress=None, log=None):
    """Telecharge un fichier en signalant l'avancement (octets, total)."""
    if log:
        log("Téléchargement : %s" % url)
    requete = urllib.request.Request(url, headers={"User-Agent": "Ollook"})
    with urllib.request.urlopen(requete, timeout=60) as reponse:
        total = int(reponse.headers.get("Content-Length") or 0)
        recu = 0
        with open(destination, "wb") as sortie:
            while True:
                morceau = reponse.read(262144)
                if not morceau:
                    break
                sortie.write(morceau)
                recu += len(morceau)
                if on_progress:
                    on_progress(recu, total)
    return destination


def install_ollama(on_progress=None, log=None):
    """Telecharge et lance l'installateur officiel d'Ollama.

    Windows : installation silencieuse. macOS : le fichier est revele dans le
    Finder, l'installation d'un .dmg restant une operation manuelle."""
    if not (WINDOWS or MACOS):
        if log:
            log("Installation automatique non prise en charge. Voir %s" % OLLAMA_PAGE)
        return False

    url = OLLAMA_WINDOWS_URL if WINDOWS else OLLAMA_MACOS_URL
    dossier = os.environ.get("TEMP") or "/tmp"
    fichier = os.path.join(dossier, os.path.basename(url))

    try:
        download(url, fichier, on_progress, log)
    except (urllib.error.URLError, OSError) as e:
        if log:
            log("Téléchargement impossible : %s" % e)
        return False

    if MACOS:
        if log:
            log("Ouvrez le fichier téléchargé pour terminer l'installation.")
        subprocess.Popen(["open", "-R", fichier])
        return False

    if log:
        log("Installation en cours… (Windows peut demander votre accord)")
    try:
        # Installateur Inno Setup : /SILENT affiche une barre, pas de questions.
        code = subprocess.call([fichier, "/SILENT", "/NORESTART"])
    except OSError as e:
        if log:
            log("Lancement de l'installateur impossible : %s" % e)
        return False

    if code != 0:
        if log:
            log("L'installateur s'est terminé avec le code %d." % code)
        return False

    # Le service demarre parfois seul apres l'installation.
    for _ in range(20):
        if find_ollama():
            break
        time.sleep(0.5)
    if log:
        log("Ollama installé.")
    return bool(find_ollama())


def pull_model(nom=DEFAULT_MODEL, on_progress=None, log=None):
    """Telecharge un modele via l'API d'Ollama, avec avancement."""
    if log:
        log("Téléchargement du modèle %s…" % nom)
    try:
        flux = core.ollama_post_stream("/api/pull", {"model": nom, "stream": True},
                                       timeout=900)
    except urllib.error.HTTPError as e:
        if log:
            log("Ollama a refusé : %s" % e.read().decode("utf-8", "replace")[:200])
        return False
    except urllib.error.URLError as e:
        if log:
            log("Ollama injoignable : %s" % e)
        return False

    import json
    dernier = ""
    try:
        for ligne in flux:
            ligne = ligne.strip()
            if not ligne:
                continue
            try:
                obj = json.loads(ligne.decode("utf-8"))
            except ValueError:
                continue
            if obj.get("error"):
                if log:
                    log("Erreur : %s" % obj["error"])
                return False
            etat = obj.get("status", "")
            if etat and etat != dernier and log:
                log(etat)
                dernier = etat
            fait, total = obj.get("completed"), obj.get("total")
            if on_progress and total:
                on_progress(int(fait or 0), int(total))
    finally:
        try:
            flux.close()
        except Exception:
            pass

    reussi = any(m["name"] == nom or m["name"].startswith(nom + ":")
                 for m in installed_models())
    if log:
        log("Modèle installé." if reussi else "Le modèle n'apparaît pas dans la liste.")
    return reussi


def human(octets):
    for unite in ("o", "Ko", "Mo", "Go"):
        if octets < 1024 or unite == "Go":
            return "%.1f %s" % (octets, unite)
        octets /= 1024.0
    return "%.1f Go" % octets


# --------------------------------------------------------------------------
# Assistant de mise en route
# --------------------------------------------------------------------------

def run_setup_ui():
    """Affiche l'assistant. Renvoie True si Ollook peut demarrer ensuite.

    N'agit qu'apres accord explicite : les tailles sont annoncees avant tout
    telechargement."""
    import queue
    import threading
    import tkinter as tk
    from tkinter import ttk

    evenements = queue.Queue()
    resultat = {"pret": False, "en_cours": False}

    racine = tk.Tk()
    racine.title("Ollook — mise en route")
    racine.resizable(False, False)
    try:
        import ollook_app
        ollook_app.appliquer_icone(racine)
    except Exception:
        pass
    try:
        racine.call("ttk::style", "theme", "use", "vista" if WINDOWS else "clam")
    except tk.TclError:
        pass

    cadre = ttk.Frame(racine, padding=16)
    cadre.pack(fill="both", expand=True)

    ttk.Label(cadre, text="Ollook a besoin d'Ollama",
              font=("Segoe UI", 12, "bold")).pack(anchor="w")

    resume = ttk.Label(cadre, text="", wraplength=430, justify="left")
    resume.pack(anchor="w", pady=(8, 0))

    etat = ttk.Label(cadre, text="", foreground="#555", wraplength=430, justify="left")
    etat.pack(anchor="w", pady=(12, 4))

    barre = ttk.Progressbar(cadre, length=430, mode="determinate")
    barre.pack(fill="x")

    journal = tk.Text(cadre, height=7, width=56, relief="flat", background="#f4f5f7",
                      font=("Consolas", 8), state="disabled", wrap="word")
    journal.pack(fill="x", pady=(10, 0))

    boutons = ttk.Frame(cadre)
    boutons.pack(fill="x", pady=(14, 0))
    quitter = ttk.Button(boutons, text="Quitter", command=racine.destroy)
    quitter.pack(side="left")
    lancer = ttk.Button(boutons, text="Installer")
    lancer.pack(side="right")

    # ---- rendu ----------------------------------------------------------

    def ecrire(ligne):
        journal.configure(state="normal")
        journal.insert("end", ligne + "\n")
        journal.see("end")
        journal.configure(state="disabled")

    def log(ligne):
        evenements.put(("log", ligne))

    def progres(fait, total):
        evenements.put(("progres", (fait, total)))

    besoins = {"ollama": False, "demarrage": False, "modele": False}

    def rafraichir():
        installe, vivant, nb = diagnose()
        besoins["ollama"] = not installe
        besoins["demarrage"] = installe and not vivant
        besoins["modele"] = vivant and nb == 0

        lignes = []
        lignes.append(("✓" if installe else "✗") + "  Ollama installé")
        lignes.append(("✓" if vivant else "✗") + "  Ollama démarré")
        lignes.append(("✓" if nb else "✗") + "  Modèle disponible"
                      + (" (%d)" % nb if nb else ""))
        etat.configure(text="\n".join(lignes))

        a_faire = []
        if besoins["ollama"]:
            a_faire.append("télécharger et installer Ollama depuis ollama.com")
        if besoins["modele"] or (besoins["ollama"] and True):
            a_faire.append("télécharger le modèle %s (%s)"
                           % (DEFAULT_MODEL, DEFAULT_MODEL_SIZE))
        if besoins["demarrage"]:
            a_faire.insert(0, "démarrer Ollama")

        if not a_faire:
            resume.configure(text="Tout est prêt.")
            lancer.configure(text="Continuer")
            return True

        resume.configure(
            text="Avec votre accord, Ollook va :\n  · "
                 + "\n  · ".join(a_faire)
                 + "\n\nRien n'est téléchargé avant que vous ne cliquiez.")
        lancer.configure(text="Installer")
        return False

    # ---- traitement ------------------------------------------------------

    def travail():
        try:
            if besoins["ollama"]:
                if not install_ollama(progres, log):
                    log("Installation d'Ollama non aboutie.")
                    evenements.put(("fini", False))
                    return
            if not ollama_alive():
                if not start_ollama(log):
                    log("Ollama ne démarre pas. Lancez-le à la main, puis relancez.")
                    evenements.put(("fini", False))
                    return
            if not installed_models():
                if not pull_model(DEFAULT_MODEL, progres, log):
                    evenements.put(("fini", False))
                    return
            evenements.put(("fini", True))
        except Exception as e:                      # noqa: BLE001
            log("Erreur : %s" % e)
            evenements.put(("fini", False))

    def demarrer():
        if resultat["en_cours"]:
            return
        if rafraichir():                            # deja pret
            resultat["pret"] = True
            racine.destroy()
            return
        resultat["en_cours"] = True
        lancer.state(["disabled"])
        quitter.state(["disabled"])
        threading.Thread(target=travail, daemon=True).start()

    lancer.configure(command=demarrer)

    def pomper():
        try:
            while True:
                genre, valeur = evenements.get_nowait()
                if genre == "log":
                    ecrire(valeur)
                elif genre == "progres":
                    fait, total = valeur
                    if total:
                        barre.configure(maximum=total, value=fait)
                        etat.configure(text="%s / %s" % (human(fait), human(total)))
                elif genre == "fini":
                    resultat["en_cours"] = False
                    resultat["pret"] = bool(valeur)
                    quitter.state(["!disabled"])
                    lancer.state(["!disabled"])
                    barre.configure(value=barre["maximum"] if valeur else 0)
                    if valeur:
                        ecrire("Terminé.")
                        racine.after(700, racine.destroy)
                    else:
                        rafraichir()
        except queue.Empty:
            pass
        racine.after(80, pomper)

    rafraichir()
    pomper()

    racine.update_idletasks()
    x = (racine.winfo_screenwidth() - racine.winfo_width()) // 2
    y = (racine.winfo_screenheight() - racine.winfo_height()) // 3
    racine.geometry("+%d+%d" % (x, y))
    racine.mainloop()

    return resultat["pret"] or ready()
