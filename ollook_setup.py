#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - mise en route.

Verifie, et installe au besoin, ce dont Ollook a besoin pour fonctionner :

  1. Ollama present sur la machine ;
  2. Ollama demarre ;
  3. au moins un modele installe, choisi selon la carte graphique.

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

# --------------------------------------------------------------------------
# Catalogue Qwen 3.5
#
# `besoin` est la memoire graphique a avoir pour faire tourner le modele :
# la taille du fichier majoree d'environ 15 % pour le cache d'attention.
# La liste est classee par QUALITE decroissante : le premier modele qui tient
# dans la carte est le bon.
#
# Sous environ quatre milliards de parametres, les modeles corrigent
# l'orthographe mais ignorent la consigne de registre -- le tutoiement en pro
# interne ne passe pas. Les petites variantes portent donc un avertissement.
# --------------------------------------------------------------------------

MODELES = [
    {"tag": "qwen3.5:35b", "taille": 24.0, "besoin": 27.6, "titre": "35 milliards",
     "note": "le plus capable ; carte de 28 Go ou plus"},
    {"tag": "qwen3.5:27b", "taille": 17.0, "besoin": 19.6, "titre": "27 milliards",
     "note": "excellent respect du registre et de la traduction"},
    {"tag": "qwen3.5:9b", "taille": 6.6, "besoin": 7.6, "titre": "9 milliards",
     "note": "bon compromis, convient a la plupart des cartes"},
    {"tag": "qwen3.5:4b", "taille": 3.4, "besoin": 3.9, "titre": "4 milliards",
     "note": "leger ; le registre passe mal"},
    {"tag": "qwen3.5:2b", "taille": 2.7, "besoin": 3.1, "titre": "2 milliards",
     "note": "dernier recours ; corrige l'orthographe, ignore le registre"},
]

# Repli quand la memoire graphique ne peut pas etre mesuree.
DEFAULT_MODEL = "qwen3.5:9b"
DEFAULT_MODEL_SIZE = "6,6 Go"


_PS_CARTES = r"""
$presentes = @(Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
                ForEach-Object { $_.Name })
Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}' -ErrorAction SilentlyContinue |
  ForEach-Object {
    $p = Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue
    $octets = $p.'HardwareInformation.qwMemorySize'
    if ($octets -and ($presentes -contains $p.DriverDesc)) {
      '{0}|{1}' -f $p.DriverDesc, $octets
    }
  }
"""


def _sortie_commande(commande, entree=None):
    try:
        return subprocess.check_output(
            commande, stderr=subprocess.DEVNULL, input=entree,
            creationflags=0x08000000 if WINDOWS else 0,
            timeout=20).decode("utf-8", "replace")
    except (OSError, subprocess.SubprocessError):
        return ""


def detecter_carte():
    """Renvoie (nom, memoire_en_Go) de la carte graphique la plus capable.

    Deux methodes, dans cet ordre :
      1. nvidia-smi, exact pour toute carte NVIDIA ;
      2. le registre Windows, qui expose la vraie taille sur 64 bits pour
         n'importe quel fabricant -- AMD et Intel compris.

    On ne se sert pas du champ AdapterRAM de WMI : c'est un entier 32 bits qui
    plafonne a 4 Go et annoncerait 4 Go pour une carte de 20. Les entrees de
    registre sont recoupees avec les cartes reellement presentes, sinon un
    pilote desinstalle fausserait la mesure."""
    chemins = ["nvidia-smi"]
    for base in (os.environ.get("PROGRAMFILES", ""), os.environ.get("SYSTEMROOT", "")):
        if base:
            chemins.append(os.path.join(base, "NVIDIA Corporation", "NVSMI",
                                        "nvidia-smi.exe"))
            chemins.append(os.path.join(base, "System32", "nvidia-smi.exe"))

    for chemin in chemins:
        sortie = _sortie_commande(
            [chemin, "--query-gpu=name,memory.total",
             "--format=csv,noheader,nounits"])
        cartes = []
        for ligne in sortie.splitlines():
            if "," not in ligne:
                continue
            nom, _, memoire = ligne.rpartition(",")
            memoire = memoire.strip()
            if memoire.isdigit():
                cartes.append((nom.strip(), int(memoire) / 1024.0))
        if cartes:
            return max(cartes, key=lambda c: c[1])

    if not WINDOWS:
        return None, None

    sortie = _sortie_commande(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", _PS_CARTES])
    cartes = []
    for ligne in sortie.splitlines():
        if "|" not in ligne:
            continue
        nom, _, octets = ligne.rpartition("|")
        octets = octets.strip()
        if octets.isdigit():
            cartes.append((nom.strip(), int(octets) / (1024.0 ** 3)))
    if cartes:
        return max(cartes, key=lambda c: c[1])

    return None, None


def detecter_vram():
    """Memoire graphique en gigaoctets, ou None."""
    return detecter_carte()[1]


def modele_recommande(vram=None):
    """Meilleur modele tenant dans la memoire graphique.

    Le catalogue etant classe par qualite, le premier qui tient est le bon."""
    if vram is None:
        vram = detecter_vram()
    if vram is None:
        return DEFAULT_MODEL
    for modele in MODELES:
        if modele["besoin"] <= vram:
            return modele["tag"]
    # Aucune carte assez grande : le plus leger, quitte a tourner sur le
    # processeur -- Ollama bascule tout seul, plus lentement.
    return MODELES[-1]["tag"]


def modele_par_tag(tag):
    for modele in MODELES:
        if modele["tag"] == tag:
            return modele
    return None


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
    choisi = [DEFAULT_MODEL]        # fige au moment du clic

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

    # ---- choix du modele ------------------------------------------------
    choix = ttk.Frame(cadre)
    choix.pack(fill="x", pady=(12, 0))
    ttk.Label(choix, text="MODÈLE À INSTALLER", font=("Segoe UI", 8),
              foreground="#777").pack(anchor="w")

    vram = detecter_vram()
    conseille = modele_recommande(vram)
    libelles = ["%s — %s, %.1f Go%s"
                % (m["tag"], m["titre"], m["taille"],
                   "   ← conseillé" if m["tag"] == conseille else "")
                for m in MODELES]
    liste = ttk.Combobox(choix, state="readonly", width=52, values=libelles)
    liste.current([m["tag"] for m in MODELES].index(conseille))
    liste.pack(fill="x", pady=(2, 0))

    detail = ttk.Label(choix, text="", foreground="#777", font=("Segoe UI", 8),
                       wraplength=430, justify="left")
    detail.pack(anchor="w", pady=(3, 0))

    def decrire(_e=None):
        modele = MODELES[liste.current()]
        tient = vram is None or modele["besoin"] <= vram
        detail.configure(
            text="%s. Il faut environ %.1f Go de mémoire graphique%s."
                 % (modele["note"], modele["besoin"],
                    "" if tient else " — au-delà de votre carte, il tournera lentement"),
            foreground="#777" if tient else "#a86a00")

    liste.bind("<<ComboboxSelected>>", decrire)
    decrire()

    ttk.Label(choix, foreground="#777", font=("Segoe UI", 8),
              text=("Carte graphique : %.1f Go de mémoire" % vram) if vram
                   else "Mémoire graphique non détectée — choix par défaut."
              ).pack(anchor="w", pady=(2, 0))

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
        if besoins["modele"] or besoins["ollama"]:
            modele = MODELES[liste.current()]
            a_faire.append("télécharger le modèle %s (%.1f Go)"
                           % (modele["tag"], modele["taille"]))
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
                if not pull_model(choisi[0], progres, log):
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
        # Fige le choix : le fil d'execution ne doit pas lire un widget.
        choisi[0] = MODELES[liste.current()]["tag"]
        lancer.state(["disabled"])
        quitter.state(["disabled"])
        liste.state(["disabled"])
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

# --------------------------------------------------------------------------
# Accelerateur neuronal
#
# Windows expose les NPU dans la classe de peripheriques ComputeAccelerator.
# Le filtre par nom couvre les pilotes qui s'en ecartent.
#
# Ollama ne sait PAS s'en servir : chaque fabricant impose sa propre pile
# (OpenVINO chez Intel, QNN chez Qualcomm, Ryzen AI chez AMD) et des modeles
# convertis en ONNX, compiles a l'avance pour cette pile. On detecte donc le
# NPU pour pouvoir le dire, pas pour s'en servir.
# --------------------------------------------------------------------------

_PS_NPU = r"""
$trouves = @()
foreach ($d in (Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue)) {
    $nom = $d.FriendlyName
    if (-not $nom) { continue }
    if ($d.Class -eq 'ComputeAccelerator' -or
        $nom -match 'AI Boost|Hexagon|XDNA|Ryzen AI|Neural Processor|NPU Compute') {
        $trouves += $nom
    }
}
($trouves | Select-Object -Unique) -join '|'
"""

# Du plus specifique au plus general : "Ryzen AI" avant un simple "AMD".
_FABRICANTS = (
    ("intel", ("ai boost", "intel(r) ai", "npu compute accelerator")),
    ("qualcomm", ("hexagon", "qualcomm", "snapdragon")),
    ("amd", ("xdna", "ryzen ai")),
)

# Pile logicielle a employer pour exploiter reellement chaque NPU.
PILES_NPU = {
    "intel": "OpenVINO",
    "qualcomm": "QNN (Hexagon)",
    "amd": "Ryzen AI",
}


def detecter_npu():
    """Renvoie (nom, fabricant) du NPU present, ou (None, None).

    `fabricant` vaut 'intel', 'qualcomm', 'amd', ou None si le nom ne permet
    pas de trancher."""
    if not WINDOWS:
        return None, None

    sortie = _sortie_commande(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", _PS_NPU])
    noms = [n.strip() for n in (sortie or "").split("|") if n.strip()]
    if not noms:
        return None, None

    nom = noms[0]
    minuscule = nom.lower()
    for fabricant, marqueurs in _FABRICANTS:
        if any(m in minuscule for m in marqueurs):
            return nom, fabricant
    return nom, None
