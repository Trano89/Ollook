#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - utilitaire autonome de relecture.

Independant d'Outlook : fonctionne dans n'importe quelle application, sur
n'importe quel type de compte de messagerie (Exchange, IMAP, Gmail...), sans
serveur, sans certificat et sans autorisation d'administrateur.

Deroulement (Windows) :
  1. ouvrez votre fenetre de redaction et ecrivez ;
  2. Ctrl+Alt+R, ou basculez sur Ollook ;
  3. choisissez la FENETRE CIBLE dans la liste, le type de courriel, le modele ;
  4. Reecrire : Ollook prend le texte de cette fenetre et l'y remplace.

macOS et repli : copiez le texte (Cmd+C), lancez Ollook, Reecrire, puis collez.

Aucune dependance pip : tkinter (fourni avec Python) et ollook_core.
"""

import os
import queue
import sys
import threading
import time
import urllib.error

try:
    import tkinter as tk
    from tkinter import ttk
except ImportError:
    sys.exit(
        "tkinter est introuvable.\n"
        "  Windows : reinstallez Python en cochant 'tcl/tk and IDLE'.\n"
        "  macOS   : brew install python-tk, ou utilisez le Python de python.org.\n"
        "  Linux   : sudo apt install python3-tk"
    )

# Une fois embarque dans un executable, le dossier est deja sur le chemin.
if not getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ollook_core as core  # noqa: E402

import ollook_version as version  # noqa: E402

try:
    import ollook_outlook as outlook  # voie exacte, Windows + Outlook
except ImportError:
    outlook = None

try:
    import ollook_update as maj
except ImportError:
    maj = None

WINDOWS = sys.platform.startswith("win")
MACOS = sys.platform == "darwin"

APP_NAME = "Ollook"
PREFS_NAME = "ollook.conf"


def app_dir():
    """Dossier de l'application : celui de l'executable une fois compile."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _writable(directory):
    probe = os.path.join(directory, ".ollook-write-test")
    try:
        with open(probe, "w"):
            pass
        os.remove(probe)
        return True
    except OSError:
        return False


def resolve_prefs_path():
    """Emplacement des preferences.

    Mode PORTABLE d'abord : le fichier vit a cote de l'executable, de sorte
    qu'une copie sur cle USB emporte ses reglages et ne laisse rien derriere
    elle. Repli sur %APPDATA% quand le dossier est en lecture seule -- cle
    protegee, installation sous Program Files."""
    local = os.path.join(app_dir(), PREFS_NAME)
    roaming = os.path.join(
        os.environ.get("APPDATA") or os.path.expanduser("~/.config"), PREFS_NAME)

    if os.path.exists(local):
        return local
    if not _writable(app_dir()):
        return roaming

    # Premier demarrage en portable : on recupere les reglages deja pris.
    if os.path.exists(roaming):
        try:
            with open(roaming, encoding="utf-8") as source:
                contenu = source.read()
            with open(local, "w", encoding="utf-8") as cible:
                cible.write(contenu)
        except OSError:
            return roaming
    return local


PREFS_PATH = resolve_prefs_path()

HOTKEY_LABEL = "Ctrl+Alt+R"
ACTIVE_WINDOW = "— fenêtre active au moment du raccourci —"

# outlook.exe : Outlook classique (fenetre principale et fenetres de redaction).
# olk.exe     : le nouvel Outlook pour Windows.
OUTLOOK_PROCESSES = frozenset(("outlook.exe", "olk.exe"))


# --------------------------------------------------------------------------
# Preferences persistantes (fichier plat, sans dependance)
# --------------------------------------------------------------------------

def load_prefs():
    prefs = {"register": "pro_externe", "model": "", "translate": "0", "scope": "all",
             "develop": core.DEFAULT_DEVELOPMENT, "maj_auto": "1"}
    try:
        with open(PREFS_PATH, encoding="utf-8") as f:
            for line in f:
                if "=" in line:
                    k, v = line.split("=", 1)
                    prefs[k.strip()] = v.strip()
    except OSError:
        pass
    if prefs.get("register") not in core.REGISTERS:
        prefs["register"] = "pro_externe"
    if prefs.get("scope") not in ("all", "selection"):
        prefs["scope"] = "all"
    if prefs.get("develop") not in core.DEVELOPMENTS:
        prefs["develop"] = core.DEFAULT_DEVELOPMENT
    return prefs


def save_prefs(prefs):
    try:
        os.makedirs(os.path.dirname(PREFS_PATH), exist_ok=True)
        with open(PREFS_PATH, "w", encoding="utf-8") as f:
            for k, v in prefs.items():
                f.write("%s=%s\n" % (k, v))
    except OSError:
        pass


# --------------------------------------------------------------------------
# Integration Windows : fenetres, raccourci global, clavier
# --------------------------------------------------------------------------

class WindowsBridge(object):
    """Enumeration des fenetres, changement de premier plan, copier/coller
    et raccourci clavier global. Uniquement ctypes, aucune dependance."""

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_NOREPEAT = 0x4000
    WM_HOTKEY = 0x0312
    VK_CONTROL = 0x11
    VK_SHIFT = 0x10
    VK_HOME = 0x24
    VK_DOWN = 0x28
    KEYEVENTF_KEYUP = 0x0002
    SW_RESTORE = 9
    GWL_EXSTYLE = -20
    WS_EX_TOOLWINDOW = 0x00000080
    DWMWA_CLOAKED = 14

    def __init__(self):
        import ctypes
        from ctypes import wintypes
        self.ctypes = ctypes
        self.wintypes = wintypes
        self.user32 = ctypes.windll.user32
        self.kernel32 = ctypes.windll.kernel32
        try:
            self.dwmapi = ctypes.windll.dwmapi
        except Exception:
            self.dwmapi = None
        self.own_hwnds = set()

        # Les handles doivent transiter en pointeur, jamais en c_int : la valeur
        # serait tronquee sur les sessions ou les HWND depassent 32 bits.
        self.user32.GetWindowTextLengthW.restype = ctypes.c_int
        self.user32.GetWindowLongW.restype = ctypes.c_long
        self.user32.GetForegroundWindow.restype = ctypes.c_void_p
        self.user32.GetParent.restype = ctypes.c_void_p
        self.user32.GetParent.argtypes = [wintypes.HWND]
        for name in ("IsWindow", "IsWindowVisible", "IsIconic", "BringWindowToTop",
                     "SetForegroundWindow"):
            getattr(self.user32, name).argtypes = [wintypes.HWND]

    # -- enumeration ----------------------------------------------------

    def _is_cloaked(self, hwnd):
        """Les fenetres UWP inactives restent 'visibles' mais sont masquees
        par le gestionnaire de composition : sans ce filtre la liste est
        polluee de fenetres fantomes."""
        if not self.dwmapi:
            return False
        ctypes = self.ctypes
        value = ctypes.c_int(0)
        try:
            self.dwmapi.DwmGetWindowAttribute(
                self.wintypes.HWND(hwnd), ctypes.c_int(self.DWMWA_CLOAKED),
                ctypes.byref(value), ctypes.sizeof(value))
        except Exception:
            return False
        return bool(value.value)

    def process_name(self, hwnd):
        """Nom de l'executable proprietaire d'une fenetre, en minuscules."""
        ctypes = self.ctypes
        wintypes = self.wintypes
        pid = wintypes.DWORD(0)
        self.user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
        if not pid.value:
            return ""

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        self.kernel32.OpenProcess.restype = ctypes.c_void_p
        handle = self.kernel32.OpenProcess(
            PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not handle:
            return ""
        try:
            size = wintypes.DWORD(512)
            buf = ctypes.create_unicode_buffer(size.value)
            ok = self.kernel32.QueryFullProcessImageNameW(
                ctypes.c_void_p(handle), 0, buf, ctypes.byref(size))
            if not ok:
                return ""
            return os.path.basename(buf.value or "").lower()
        finally:
            self.kernel32.CloseHandle(ctypes.c_void_p(handle))

    def list_windows(self, processes=None):
        """Fenetres de premier niveau visibles, dans l'ordre de profondeur :
        la plus recemment utilisee en tete.

        `processes` : ensemble de noms d'executables auquel se limiter."""
        ctypes = self.ctypes
        wintypes = self.wintypes
        found = []

        proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

        def callback(hwnd, _lparam):
            if hwnd in self.own_hwnds:
                return True
            if not self.user32.IsWindowVisible(hwnd):
                return True
            if self.user32.GetWindowLongW(hwnd, self.GWL_EXSTYLE) & self.WS_EX_TOOLWINDOW:
                return True
            if self._is_cloaked(hwnd):
                return True
            length = self.user32.GetWindowTextLengthW(hwnd)
            if length <= 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            self.user32.GetWindowTextW(hwnd, buf, length + 1)
            title = (buf.value or "").strip()
            if not title:
                return True
            if processes and self.process_name(hwnd) not in processes:
                return True
            found.append((int(hwnd), title))
            return True

        self.user32.EnumWindows(proc(callback), 0)
        return found

    def window_title(self, hwnd):
        ctypes = self.ctypes
        length = self.user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(length + 1)
        self.user32.GetWindowTextW(hwnd, buf, length + 1)
        return (buf.value or "").strip()

    def is_window(self, hwnd):
        return bool(hwnd) and bool(self.user32.IsWindow(hwnd))

    # -- premier plan ---------------------------------------------------

    def foreground(self):
        return int(self.user32.GetForegroundWindow() or 0)

    def focus(self, hwnd):
        """Met une fenetre au premier plan. Windows refuse SetForegroundWindow
        a un processus qui n'a pas le focus ; on s'attache brievement a la file
        d'entree de la fenetre visee, ce qui leve la restriction."""
        if not self.is_window(hwnd):
            return False
        if self.user32.IsIconic(hwnd):
            self.user32.ShowWindow(hwnd, self.SW_RESTORE)
            time.sleep(0.08)
        if self.foreground() == hwnd:
            return True

        cur = self.kernel32.GetCurrentThreadId()
        target = self.user32.GetWindowThreadProcessId(hwnd, None)
        attached = False
        if target and target != cur:
            attached = bool(self.user32.AttachThreadInput(cur, target, True))
        try:
            self.user32.BringWindowToTop(hwnd)
            self.user32.SetForegroundWindow(hwnd)
        finally:
            if attached:
                self.user32.AttachThreadInput(cur, target, False)

        for _ in range(10):
            time.sleep(0.03)
            if self.foreground() == hwnd:
                return True
        return False

    # -- clavier --------------------------------------------------------

    def _key(self, vk, ctrl=False, shift=False):
        if ctrl:
            self.user32.keybd_event(self.VK_CONTROL, 0, 0, 0)
        if shift:
            self.user32.keybd_event(self.VK_SHIFT, 0, 0, 0)
        self.user32.keybd_event(vk, 0, 0, 0)
        self.user32.keybd_event(vk, 0, self.KEYEVENTF_KEYUP, 0)
        if shift:
            self.user32.keybd_event(self.VK_SHIFT, 0, self.KEYEVENTF_KEYUP, 0)
        if ctrl:
            self.user32.keybd_event(self.VK_CONTROL, 0, self.KEYEVENTF_KEYUP, 0)

    def _tap(self, vk):
        self._key(vk, ctrl=True)

    def select_all(self):
        self._tap(0x41)   # A
        time.sleep(0.06)

    def select_top_paragraphs(self, count):
        """Selectionne les `count` premiers paragraphes du corps.

        Ctrl+Maj+Bas etend la selection PARAGRAPHE par paragraphe : le resultat
        ne depend ni de la largeur de la fenetre ni de la taille de police,
        contrairement a une selection ligne a ligne.

        Les frappes sont espacees : envoyees en rafale, Outlook en perd, et une
        selection incomplete signifie un texte colle au mauvais endroit."""
        time.sleep(0.12)                        # laisse la fenetre accepter les frappes
        self._key(self.VK_HOME, ctrl=True)      # debut du corps
        time.sleep(0.10)
        for _ in range(count):
            self._key(self.VK_DOWN, ctrl=True, shift=True)
            time.sleep(0.05)
        time.sleep(0.10)

    def copy(self):
        self._tap(0x43)   # C
        time.sleep(0.22)  # laisse l'application remplir le presse-papiers

    def paste(self):
        self._tap(0x56)   # V
        time.sleep(0.06)

    # -- raccourci global -----------------------------------------------

    def listen(self, on_hotkey):
        """Boucle de messages Win32, a lancer dans un thread dedie."""
        ctypes = self.ctypes
        if not self.user32.RegisterHotKey(
                None, 1, self.MOD_CONTROL | self.MOD_ALT | self.MOD_NOREPEAT, 0x52):  # R
            return False
        msg = self.wintypes.MSG()
        while self.user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == self.WM_HOTKEY:
                on_hotkey()
        return True


# --------------------------------------------------------------------------
# Presse-papiers
# --------------------------------------------------------------------------

def appliquer_icone(fenetre):
    """Pose l'icone d'Ollook sur une fenetre Tk.

    Sans cela, Tk affiche sa plume par defaut -- ce qui donne a l'application
    l'air de tourner sous Python. Les images sont embarquees en base64 pour que
    l'executable portable reste un fichier unique."""
    try:
        import ollook_icon
    except ImportError:
        return

    images = []
    for taille in (256, 64, 32, 16):
        donnees = ollook_icon.ICONES.get(taille)
        if not donnees:
            continue
        try:
            images.append(tk.PhotoImage(data=donnees))
        except tk.TclError:
            pass
    if not images:
        return
    try:
        fenetre.iconphoto(True, *images)
        # Reference conservee : Tk ne retient pas les PhotoImage lui-meme.
        fenetre._ollook_icones = images
    except tk.TclError:
        pass


def declarer_application():
    """Identifie Ollook aupres du shell Windows.

    Sans identifiant propre, la barre des taches regroupe la fenetre sous
    l'icone de l'interpreteur qui l'heberge."""
    if not WINDOWS:
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Ollook.Relecture")
    except Exception:
        pass


def paragraphes(texte):
    """Nombre de paragraphes d'un texte, borne par securite.

    Chaque saut de ligne du texte capture correspond a un paragraphe dans
    l'editeur d'Outlook : c'est ce qui permet de reselectionner exactement la
    zone du brouillon."""
    return max(1, min(300, len((texte or "").split("\n"))))


def memes_mots(a, b):
    """Compare deux textes en ignorant toute l'espace.

    Sert a verifier qu'une selection correspond bien au texte attendu : les
    marques de paragraphe et les fins de ligne different entre l'editeur et le
    presse-papiers, les mots non."""
    return "".join((a or "").split()) == "".join((b or "").split())


def clipboard_get(root):
    try:
        return root.clipboard_get()
    except tk.TclError:
        return ""


def clipboard_set(root, text):
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()      # publie le contenu aupres du systeme


# --------------------------------------------------------------------------
# Application
# --------------------------------------------------------------------------

class App(object):

    def __init__(self, initial_target=None):
        self.prefs = load_prefs()
        self.bridge = WindowsBridge() if WINDOWS else None
        self.events = queue.Queue()

        self.windows = []          # [(hwnd, titre)] dans l'ordre de la liste
        # Fenetre active relevee AVANT la creation de l'interface : Tk passe
        # brievement au premier plan, apres quoi Ollook se lirait lui-meme.
        self.target_hwnd = initial_target
        self.captured = ""         # texte brut capture, tel quel
        # "outlook" : lu par le modele objet, positions exactes.
        # "presse-papiers" : lu au clavier, remplacement approximatif.
        self.source = "presse-papiers"
        self.draft = ""            # seule partie soumise au modele
        self.signature = ""
        self.context = ""
        self.result = ""
        self.busy = False
        self.model_names = []
        self._placed = False
        # Vrai quand le dernier remplacement a pu cibler les seuls paragraphes
        # du brouillon, faux quand il a fallu se rabattre sur le message entier.
        self.derniere_precision = False
        self.maj_info = None       # publication detectee
        self.maj_chemin = None     # executable telecharge, pret a remplacer
        self.maj_avis = None       # fenetre d'avis, si affichee

        self.root = tk.Tk()
        self.root.title(APP_NAME)
        appliquer_icone(self.root)
        self.root.withdraw()
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.resizable(False, False)

        self._build()
        self._on_develop()          # affiche l'indication du niveau
        self._register_own_window()
        self._poll_events()
        self._refresh_models_async()
        self._chercher_maj_async()

    def _register_own_window(self):
        """Memorise nos propres HWND pour ne jamais nous proposer comme cible."""
        if not self.bridge:
            return
        try:
            self.root.update_idletasks()
            hwnd = self.root.winfo_id()
            self.bridge.own_hwnds.add(int(hwnd))
            parent = self.bridge.user32.GetParent(hwnd)
            if parent:
                self.bridge.own_hwnds.add(int(parent))
        except Exception:
            pass

    # ---- construction de l'interface ---------------------------------

    def _build(self):
        style = ttk.Style()
        try:
            style.theme_use("vista" if WINDOWS else "aqua" if MACOS else "clam")
        except tk.TclError:
            pass

        pad = {"padx": 12}
        frame = ttk.Frame(self.root, padding=(0, 10, 0, 10))
        frame.pack(fill="both", expand=True)

        # En-tete
        head = ttk.Frame(frame)
        head.pack(fill="x", **pad)
        ttk.Label(head, text=APP_NAME, font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Label(head, text="v" + version.VERSION, foreground="#999",
                  font=("Segoe UI", 8)).pack(side="left", padx=(5, 0))
        self.status = ttk.Label(head, text="connexion…", foreground="#777")
        self.status.pack(side="right")

        # ---- fenetre cible (Windows uniquement) ----
        if self.bridge:
            ttk.Label(frame, text="FENÊTRE OUTLOOK À RELIRE", font=("Segoe UI", 8),
                      foreground="#777").pack(anchor="w", pady=(12, 2), **pad)

            row = ttk.Frame(frame)
            row.pack(fill="x", **pad)
            self.window_box = ttk.Combobox(row, state="readonly", width=40)
            self.window_box.pack(side="left", fill="x", expand=True)
            self.window_box.bind("<<ComboboxSelected>>", self._on_window_chosen)
            ttk.Button(row, text="⟳", width=3,
                       command=self.refresh_windows).pack(side="left", padx=(4, 0))

            scope_row = ttk.Frame(frame)
            scope_row.pack(fill="x", pady=(6, 0), **pad)
            self.scope = tk.StringVar(value=self.prefs.get("scope", "all"))
            ttk.Radiobutton(scope_row, text="Tout le message", value="all",
                            variable=self.scope, command=self._save).pack(side="left")
            ttk.Radiobutton(scope_row, text="Ma sélection", value="selection",
                            variable=self.scope, command=self._save).pack(side="left",
                                                                          padx=(12, 0))
            ttk.Button(scope_row, text="Capturer",
                       command=self.capture_now).pack(side="right")

        # Ce qui a ete capture
        self.target = ttk.Label(frame, text="", foreground="#555", wraplength=360,
                                justify="left")
        self.target.pack(fill="x", pady=(8, 0), **pad)

        ttk.Separator(frame).pack(fill="x", pady=10)

        # Registre
        ttk.Label(frame, text="TYPE DE COURRIEL", font=("Segoe UI", 8),
                  foreground="#777").pack(anchor="w", **pad)
        self.register = tk.StringVar(value=self.prefs["register"])
        for key in core.REGISTER_ORDER:
            reg = core.REGISTERS[key]
            ttk.Radiobutton(frame, variable=self.register, value=key,
                            text="%s   —   %s" % (reg["label"], reg["hint"]),
                            command=self._save).pack(anchor="w", pady=1, **pad)

        # Developpement
        ttk.Label(frame, text="DÉVELOPPEMENT", font=("Segoe UI", 8),
                  foreground="#777").pack(anchor="w", pady=(12, 2), **pad)
        dev_row = ttk.Frame(frame)
        dev_row.pack(fill="x", **pad)
        self.develop = tk.StringVar(value=self.prefs["develop"])
        for key in core.DEVELOPMENT_ORDER:
            ttk.Radiobutton(dev_row, text=core.DEVELOPMENTS[key]["label"], value=key,
                            variable=self.develop,
                            command=self._on_develop).pack(side="left", padx=(0, 12))
        self.develop_hint = ttk.Label(frame, text="", foreground="#777",
                                      font=("Segoe UI", 8))
        self.develop_hint.pack(anchor="w", **pad)

        # Traduction
        self.translate = tk.BooleanVar(value=self.prefs.get("translate") == "1")
        ttk.Checkbutton(frame, text="Traduire en anglais", variable=self.translate,
                        command=self._save).pack(anchor="w", pady=(10, 0), **pad)

        # Modele
        ttk.Label(frame, text="MODÈLE", font=("Segoe UI", 8),
                  foreground="#777").pack(anchor="w", pady=(12, 2), **pad)
        self.model = ttk.Combobox(frame, state="readonly", width=40)
        self.model.pack(fill="x", **pad)
        self.model.bind("<<ComboboxSelected>>", lambda e: self._save())

        # Message d'etat
        self.message = ttk.Label(frame, text="", wraplength=360, justify="left")
        self.message.pack(fill="x", pady=(10, 0), **pad)

        # Boutons
        buttons = ttk.Frame(frame)
        buttons.pack(fill="x", pady=(14, 0), **pad)
        ttk.Button(buttons, text="Fermer", command=self.hide).pack(side="left")
        self.undo_btn = ttk.Button(buttons, text="Restaurer", command=self.undo)
        self.undo_btn.pack(side="left", padx=(6, 0))
        self.undo_btn.state(["disabled"])
        self.go_btn = ttk.Button(buttons, text="Réécrire", command=self.run)
        self.go_btn.pack(side="right")

        self.root.bind("<Escape>", lambda e: self.hide())
        self.root.bind("<Return>", lambda e: self.run())

    def _on_develop(self):
        self.develop_hint.configure(
            text=core.DEVELOPMENTS[self.develop.get()]["hint"])
        self._save()

    def _save(self):
        self.prefs["register"] = self.register.get()
        self.prefs["develop"] = self.develop.get()
        self.prefs["translate"] = "1" if self.translate.get() else "0"
        if self.bridge:
            self.prefs["scope"] = self.scope.get()
        # Le nom reel du modele, jamais le libelle affiche : celui-ci porte la
        # taille entre parentheses et Ollama le rejetterait.
        name = self.selected_model()
        if name:
            self.prefs["model"] = name
        save_prefs(self.prefs)

    # ---- liste des fenetres --------------------------------------------

    def refresh_windows(self, keep=True):
        """Recharge la liste. `keep` : tenter de conserver la cible choisie."""
        if not self.bridge:
            return
        previous = self.target_hwnd if keep else None
        if not keep:
            self.target_hwnd = None

        self.windows = self.bridge.list_windows(processes=OUTLOOK_PROCESSES)

        # Une cible valide absente de l'enumeration — autre application
        # designee par le raccourci, fenetre outil — resterait introuvable :
        # on la remet en tete plutot que de la perdre.
        if previous and not any(h == previous for h, _t in self.windows):
            if self.bridge.is_window(previous):
                self.windows.insert(0, (previous, self.bridge.window_title(previous)
                                        or "(fenêtre sans titre)"))
            else:
                previous = None
                self.target_hwnd = None

        labels = [ACTIVE_WINDOW]
        for _hwnd, title in self.windows:
            labels.append(title if len(title) <= 58 else title[:57] + "…")
        self.window_box["values"] = labels

        index = 0
        if previous:
            for i, (hwnd, _t) in enumerate(self.windows):
                if hwnd == previous:
                    index = i + 1
                    break
        self.window_box.current(index)

        if not self.windows:
            self.set_message("Aucune fenêtre Outlook ouverte. Ouvrez votre message, "
                             "puis cliquez sur ⟳.", error=True)

    def _on_window_chosen(self, _event=None):
        index = self.window_box.current()
        self.target_hwnd = None if index <= 0 else self.windows[index - 1][0]
        self.captured = ""
        self.set_message("")
        self.target.configure(
            text="Cliquez sur « Capturer » pour voir le texte, "
                 "ou directement sur « Réécrire »." if self.target_hwnd
                 else "La fenêtre active au moment du raccourci sera utilisée.")

    # ---- etat d'Ollama --------------------------------------------------

    def _refresh_models_async(self):
        threading.Thread(target=self._refresh_models, daemon=True).start()

    def _refresh_models(self):
        try:
            version = core.ollama_version()
            models = core.list_models()
        except Exception as e:
            self.events.put(("ollama_down", str(e)))
            return
        self.events.put(("ollama_up", (version, models)))

    def _apply_models(self, version, models):
        self.status.configure(text="Ollama %s" % version, foreground="#1a7f4b")
        if not models:
            self.model["values"] = []
            self.set_message("Aucun modèle installé. Par exemple : ollama pull mistral-small",
                             error=True)
            self.go_btn.state(["disabled"])
            return

        labels = ["%s  (%s)" % (m["name"], m["parameters"]) if m["parameters"] else m["name"]
                  for m in models]
        self.model_names = [m["name"] for m in models]
        self.model["values"] = labels

        want = self.prefs.get("model")
        self.model.current(self.model_names.index(want) if want in self.model_names else 0)
        self.go_btn.state(["!disabled"])

    def selected_model(self):
        index = self.model.current()
        if index < 0 or not self.model_names:
            return ""
        return self.model_names[index]

    # ---- capture ---------------------------------------------------------

    def capture_now(self):
        """Bouton « Capturer » : va chercher le texte dans la fenetre cible."""
        if self.busy:
            return
        ok, why = self._grab()
        if not ok:
            self.set_message(why, error=True)
        self._show_capture()

    def _grab(self):
        """Recupere le texte de la fenetre cible. Renvoie (succes, raison)."""
        if not self.bridge:
            # macOS et repli : on se contente du presse-papiers.
            self._absorb(clipboard_get(self.root))
            return (bool(self.draft.strip()),
                    "Le presse-papiers est vide. Copiez votre texte (Cmd+C).")

        hwnd = self.target_hwnd
        if not hwnd:
            self._absorb(clipboard_get(self.root))
            return (bool(self.draft.strip()),
                    "Aucune fenêtre désignée. Choisissez-la dans la liste ci-dessus.")

        if not self.bridge.is_window(hwnd):
            self.refresh_windows(keep=False)
            return False, "Cette fenêtre n'existe plus. La liste a été rafraîchie."

        # Voie exacte : lire le corps par le modele objet d'Outlook. Elle donne
        # des positions de caracteres utilisables pour un remplacement
        # chirurgical, la ou le presse-papiers ne donne que du texte.
        if self.scope.get() == "all" and outlook is not None:
            self.root.update()
            if self.bridge.focus(hwnd):          # ActiveInspector = cette fenetre
                corps = outlook.lire_corps()
                if corps and corps.strip():
                    self._absorb(corps)
                    self._stay_on_top()
                    if not self.draft.strip():
                        return False, ("Ce message ne contient qu'une signature "
                                       "ou du texte cité : rien à relire.")
                    self.source = "outlook"
                    return True, ""
            self._stay_on_top()

        # Ollook reste visible et au-dessus : seul le FOCUS CLAVIER doit passer
        # a la fenetre cible pour qu'elle recoive les frappes. La masquer etait
        # inutile, et c'est ce qui la faisait disparaitre apres chaque relecture.
        self.root.update()
        if not self.bridge.focus(hwnd):
            return False, ("Windows a refusé de basculer sur « %s ». "
                           "Cliquez dans cette fenêtre, puis %s."
                           % (self.bridge.window_title(hwnd)[:40], HOTKEY_LABEL))
        if self.scope.get() == "all":
            self.bridge.select_all()
        self.bridge.copy()
        self._absorb(clipboard_get(self.root))
        self.source = "presse-papiers"
        self._stay_on_top()

        if not self.draft.strip():
            return False, ("Rien n'a été capturé. Placez le curseur dans le corps du "
                           "message, ou passez en « Ma sélection » après avoir sélectionné.")
        return True, ""

    def _absorb(self, text):
        self.captured = text or ""
        self.draft, self.signature, self.context = core.split_message(self.captured)
        self.result = ""
        # Remise a zero systematique : une capture au presse-papiers qui
        # heriterait de "outlook" ferait remplacer a des positions fausses.
        self.source = "presse-papiers"
        self.zone_ecrite = ""

    def _show_capture(self):
        words = len(self.draft.split())
        if not words:
            self.target.configure(text="Aucun texte capturé.")
            return
        garde = []
        if self.signature:
            garde.append("signature")
        if self.context:
            garde.append("fil cité")
        self.target.configure(text="%d mot%s à relire%s" % (
            words, "s" if words > 1 else "",
            (" · %s conservé%s tel%s quel%s" % (
                " et ".join(garde), "s" if len(garde) > 1 else "",
                "s" if len(garde) > 1 else "", "s" if len(garde) > 1 else ""))
            if garde else ""))

    # ---- reecriture ------------------------------------------------------

    def run(self):
        if self.busy:
            return
        model = self.selected_model()
        if not model:
            self.set_message("Aucun modèle sélectionné.", error=True)
            return

        # Capture systematique juste avant la reecriture : le message a pu
        # changer depuis l'ouverture de la fenetre ou depuis « Capturer ».
        ok, why = self._grab()
        self._show_capture()
        if not ok:
            self.set_message(why, error=True)
            return

        self.busy = True
        self._save()
        self.go_btn.state(["disabled"])
        self.go_btn.configure(text="Relecture…")
        self.set_message("")

        threading.Thread(target=self._work, args=(
            model, self.draft, self.context, self.register.get(),
            self.translate.get(), self.develop.get()), daemon=True).start()

    def _work(self, model, draft, context, register, translate, develop):
        started = time.time()
        try:
            for kind, value in core.rewrite_stream(model, draft, context, register,
                                                   translate, develop):
                if kind == "thinking":
                    self.events.put(("thinking", value))
                elif kind == "done":
                    self.events.put(("done", (value, time.time() - started)))
        except ValueError as e:
            self.events.put(("failed", str(e)))
        except urllib.error.URLError:
            self.events.put(("failed", "Ollama ne répond pas. Lancez-le :  ollama serve"))
        except Exception as e:
            self.events.put(("failed", str(e)))

    def _finish(self, text, seconds):
        self.busy = False
        self.go_btn.state(["!disabled"])
        self.go_btn.configure(text="Réécrire")

        if not text or text == "(rien a reecrire)":
            self.set_message("Le modèle n'a rien renvoyé.", error=True)
            return

        self.result = text                       # le brouillon reecrit, seul

        if self._remplacer_brouillon(self.result):
            # « Restaurer » n'a de sens que si quelque chose a ete ecrit.
            self.undo_btn.state(["!disabled"])
            garde = []
            if self.signature:
                garde.append("signature")
            if self.context:
                garde.append("fil cité")
            if self.derniere_precision and garde:
                detail = " — %s intact%s" % (" et ".join(garde),
                                             "s" if len(garde) > 1 else "")
            elif garde:
                detail = " — message entier réécrit, la mise en forme du %s a pu être aplatie" \
                         % (" et du ".join(garde))
            else:
                detail = ""
            self.set_message("Réécrit en %.1f s et remplacé%s. "
                             "« Restaurer » remet le texte d'origine."
                             % (seconds, detail))
        elif self.source == "outlook":
            self.set_message("Réécrit en %.1f s, mais le message a changé depuis la "
                             "capture — rien n'a été remplacé. Le texte est copié : "
                             "collez-le avec %s, ou relancez la relecture."
                             % (seconds, "Cmd+V" if MACOS else "Ctrl+V"))
        else:
            self.set_message("Réécrit en %.1f s et copié. Collez avec %s — "
                             "Windows a refusé de rendre le focus."
                             % (seconds, "Cmd+V" if MACOS else "Ctrl+V"))

    # ---- ecriture dans la fenetre cible ----------------------------------
    #
    # Trois mecanismes, du plus sur au plus grossier. Chacun note dans
    # `self.zone_ecrite` ce qu'il a reellement inscrit : c'est ce temoin, et
    # non une reconstruction, qui permet a « Restaurer » de viser juste.

    def _ecrire_par_outlook(self, texte, temoin):
        """Remplace une plage de caracteres via le modele objet d'Outlook.

        Voie exacte : rien d'autre que la plage n'est touche, donc la mise en
        forme de la signature et du fil cite est integralement preservee."""
        if outlook is None or self.source != "outlook":
            return False
        if not outlook.remplacer_debut(len(temoin), texte, temoin):
            return False
        self.zone_ecrite = texte
        self.derniere_precision = True
        return True

    def _ecrire_par_paragraphes(self, texte, temoin):
        """Selectionne les paragraphes du temoin au clavier, puis colle.

        La selection est RELUE et comparee avant tout collage : sans cette
        verification, une selection ratee ferait ajouter le texte a la suite
        au lieu de le remplacer."""
        if not self.bridge:
            return False
        self.bridge.select_top_paragraphs(paragraphes(temoin))
        self.bridge.copy()
        if not memes_mots(clipboard_get(self.root), temoin):
            return False
        clipboard_set(self.root, texte)
        self.bridge.paste()
        self.zone_ecrite = texte
        self.derniere_precision = True
        return True

    def _ecrire_tout(self, texte):
        """Dernier recours : tout selectionner et tout remplacer.

        La mise en forme du fil cite est aplatie ; on ne l'emploie que faute
        de mieux, et l'utilisateur en est averti."""
        if not self.bridge:
            return False
        self.bridge.select_all()
        clipboard_set(self.root, texte)
        self.bridge.paste()
        self.zone_ecrite = texte
        self.derniere_precision = False
        return True

    def _cible_utilisable(self):
        return bool(self.bridge and self.target_hwnd
                    and self.bridge.is_window(self.target_hwnd))

    def _remplacer_brouillon(self, texte):
        """Ecrit le brouillon reecrit dans la fenetre cible.

        Renvoie False si rien n'a pu etre ecrit ; le texte est alors laisse
        dans le presse-papiers pour que l'utilisateur colle lui-meme."""
        if not self._cible_utilisable():
            clipboard_set(self.root, texte)
            return False

        if self.source == "outlook":
            if self._ecrire_par_outlook(texte, self.draft):
                self._stay_on_top()
                return True
            # On ne se rabat PAS sur un collage integral : cela abimerait la
            # mise en forme de la signature, ce que l'utilisateur refuse.
            self.derniere_precision = False
            clipboard_set(self.root, texte)
            self._stay_on_top()
            return False

        self.root.update()
        if not self.bridge.focus(self.target_hwnd):
            clipboard_set(self.root, texte)
            self._stay_on_top()
            return False

        if self.scope.get() == "all":
            if not self._ecrire_par_paragraphes(texte + self._separateur_signature(),
                                                self.draft):
                self._ecrire_tout(core.join_message(texte, self.signature,
                                                    self.context))
        else:
            # La selection de l'utilisateur delimite la zone : on lui restitue
            # tout ce qu'elle contenait, sans quoi le reste serait perdu.
            clipboard_set(self.root,
                          core.join_message(texte, self.signature, self.context))
            self.bridge.paste()
            self.zone_ecrite = core.join_message(texte, self.signature, self.context)
            self.derniere_precision = True

        self._stay_on_top()
        return True

    def _separateur_signature(self):
        """Ligne vide a conserver entre le message et la signature.

        Quand seul le brouillon est remplace, la ligne vide qui le separait de
        la signature reste en place et il n'y a rien a faire. Si le message
        d'origine n'en avait pas, on l'ajoute : une signature collee au texte
        se lit mal."""
        if not self.signature or not self.captured:
            return ""
        debut = self.captured.find(self.signature[:40])
        if debut <= 0:
            return ""
        return "" if self.captured[:debut].endswith("\n\n") else "\n"

    def undo(self):
        """Remet le texte d'origine exactement la ou la reecriture a ecrit."""
        if not self.captured or not self.zone_ecrite:
            return
        if not self._cible_utilisable():
            clipboard_set(self.root, self.draft)
            self.set_message("Fenêtre cible introuvable. Texte d'origine copié.",
                             error=True)
            return

        if self.derniere_precision:
            # Seule la zone du brouillon avait ete touchee.
            if self.source == "outlook":
                ok = self._ecrire_par_outlook(self.draft, self.zone_ecrite)
            else:
                self.root.update()
                ok = (self.bridge.focus(self.target_hwnd)
                      and self._ecrire_par_paragraphes(self.draft, self.zone_ecrite))
        else:
            # Le message entier avait ete remplace : on le remet entier.
            self.root.update()
            ok = (self.bridge.focus(self.target_hwnd)
                  and self._ecrire_tout(self.captured))

        self._stay_on_top()
        if ok:
            self.set_message("Texte d'origine restauré.")
            self.undo_btn.state(["disabled"])
        else:
            clipboard_set(self.root, self.draft)
            self.set_message("Restauration impossible — le message a changé. "
                             "Texte d'origine copié, collez avec %s."
                             % ("Cmd+V" if MACOS else "Ctrl+V"), error=True)

    # ---- fenetre ---------------------------------------------------------

    def _stay_on_top(self):
        """Reaffirme la position au-dessus des autres fenetres. Donner le focus
        a une autre fenetre peut faire perdre le rang superieur ; on le remet
        sans voler le focus clavier."""
        try:
            self.root.attributes("-topmost", True)
            self.root.lift()
        except tk.TclError:
            pass

    def show(self, focus=True):
        self.set_message("")
        self.go_btn.configure(text="Réécrire")
        self.root.deiconify()
        self.root.update_idletasks()
        self._place_once()
        self._stay_on_top()
        if focus:
            self.root.focus_force()

    def _place_once(self):
        """Positionne la fenetre a la premiere ouverture seulement : ensuite
        elle reste ou l'utilisateur l'a mise."""
        if self._placed:
            return
        self._placed = True

        w = self.root.winfo_width()
        try:
            x = int(self.prefs.get("x", ""))
            y = int(self.prefs.get("y", ""))
        except (TypeError, ValueError):
            # Par defaut en haut a droite : moins genant qu'au centre pour une
            # fenetre qui reste au-dessus de tout.
            x = self.root.winfo_screenwidth() - w - 40
            y = 60

        # Ramene la fenetre dans l'ecran si l'affichage a change depuis.
        x = max(0, min(x, self.root.winfo_screenwidth() - 200))
        y = max(0, min(y, self.root.winfo_screenheight() - 200))
        self.root.geometry("+%d+%d" % (x, y))

    def _remember_position(self):
        try:
            self.prefs["x"] = str(self.root.winfo_x())
            self.prefs["y"] = str(self.root.winfo_y())
            save_prefs(self.prefs)
        except tk.TclError:
            pass

    def hide(self):
        if self.busy:
            return
        self._remember_position()
        self.root.attributes("-topmost", False)
        self.root.withdraw()

    def set_message(self, text, error=False):
        self.message.configure(text=text, foreground="#b4232b" if error else "#555")


    # ---- mise a jour ---------------------------------------------------
    #
    # Deroulement automatique : on interroge GitHub au demarrage, on telecharge
    # en arriere-plan sans rien demander, puis on installe des que l'instant
    # s'y prete. Ce qui n'est jamais automatique, c'est le MOMENT : jamais
    # pendant une relecture, et jamais sans un court avis a l'ecran, sous peine
    # de faire disparaitre l'application sous les doigts de l'utilisateur.

    DELAI_AVIS = 10          # secondes d'avis avant redemarrage

    def _chercher_maj_async(self):
        """Interroge GitHub en arriere-plan. Silencieux en cas d'echec :
        une mise a jour ne doit jamais retarder ni empecher l'usage."""
        if maj is None:
            return

        def travail():
            info = maj.verifier()
            if info:
                self.events.put(("maj", info))

        threading.Thread(target=travail, daemon=True).start()

    def _maj_trouvee(self, info):
        """Une version plus recente existe : on la telecharge sans rien demander."""
        self.maj_info = info

        # Sans executable a echanger -- lance depuis les sources -- ou si
        # l'utilisateur a coupe l'automatisme, on se contente de le signaler.
        if not (maj.telechargeable() and info.get("lien")
                and self.prefs.get("maj_auto", "1") == "1"):
            self.set_message("Ollook %s est disponible sur la page des publications."
                             % info["version"])
            return

        def travail():
            try:
                chemin = maj.telecharger(info["lien"])
            except Exception as e:                       # noqa: BLE001
                self.events.put(("maj_echec", str(e)))
                return
            self.events.put(("maj_prete", chemin))

        threading.Thread(target=travail, daemon=True).start()

    def _maj_prete(self, chemin):
        """Le nouvel executable est en place a cote de l'ancien."""
        self.maj_chemin = chemin
        self._tenter_installation()

    def _tenter_installation(self):
        """Installe des que l'application est disponible.

        Une relecture en cours n'est jamais interrompue : on repasse plus tard."""
        if not self.maj_chemin or self.maj_avis is not None:
            return
        if self.busy:
            self.root.after(4000, self._tenter_installation)
            return
        self._afficher_avis_maj()

    def _afficher_avis_maj(self):
        """Court avis, puis redemarrage. L'utilisateur peut differer."""
        info = self.maj_info or {}
        fenetre = tk.Toplevel(self.root)
        self.maj_avis = fenetre
        fenetre.title("Mise a jour d'Ollook")
        fenetre.resizable(False, False)
        appliquer_icone(fenetre)
        fenetre.protocol("WM_DELETE_WINDOW", self._differer_maj)

        cadre = ttk.Frame(fenetre, padding=16)
        cadre.pack(fill="both", expand=True)
        ttk.Label(cadre, text="Ollook %s est prete" % info.get("version", ""),
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")

        compte = ttk.Label(cadre, foreground="#555", wraplength=340, justify="left")
        compte.pack(anchor="w", pady=(6, 0))

        auto = tk.BooleanVar(value=self.prefs.get("maj_auto", "1") == "1")

        def basculer():
            self.prefs["maj_auto"] = "1" if auto.get() else "0"
            save_prefs(self.prefs)

        ttk.Checkbutton(cadre, text="Installer les mises a jour automatiquement",
                        variable=auto, command=basculer).pack(anchor="w", pady=(10, 0))

        boutons = ttk.Frame(cadre)
        boutons.pack(fill="x", pady=(14, 0))
        ttk.Button(boutons, text="Plus tard",
                   command=self._differer_maj).pack(side="left")
        ttk.Button(boutons, text="Redemarrer maintenant",
                   command=self._installer_maj).pack(side="right")

        fenetre.update_idletasks()
        fenetre.geometry("+%d+%d" % (max(0, self.root.winfo_x() + 40),
                                     max(0, self.root.winfo_y() + 60)))
        fenetre.attributes("-topmost", True)

        restant = [self.DELAI_AVIS]

        def tictac():
            if self.maj_avis is not fenetre:
                return
            if self.busy:                      # relecture en cours : on patiente
                compte.configure(text="Installation des la fin de la relecture.")
                fenetre.after(1000, tictac)
                return
            if restant[0] <= 0:
                self._installer_maj()
                return
            compte.configure(
                text="Ollook redemarre dans %d seconde%s pour terminer "
                     "l'installation." % (restant[0], "s" if restant[0] > 1 else ""))
            restant[0] -= 1
            fenetre.after(1000, tictac)

        tictac()

    def _differer_maj(self):
        """Reporte l'installation a la prochaine ouverture."""
        if self.maj_avis is not None:
            try:
                self.maj_avis.destroy()
            except tk.TclError:
                pass
            self.maj_avis = None
        self.set_message("Mise a jour reportee : elle s'installera au prochain "
                         "demarrage d'Ollook.")

    def _installer_maj(self):
        """Confie l'echange au script de relais, puis ferme Ollook.

        Le relais attend justement la disparition du processus pour pouvoir
        remplacer le fichier, puis relance l'application."""
        if not self.maj_chemin:
            return
        try:
            maj.installer(self.maj_chemin)
        except Exception as e:                       # noqa: BLE001
            self.set_message("Installation impossible : %s" % e, error=True)
            self._differer_maj()
            return
        self.root.after(300, self.root.destroy)

    # ---- boucle d'evenements ---------------------------------------------

    def _poll_events(self):
        # La fenetre peut disparaitre pendant qu'un rappel est en attente.
        try:
            if not self.root.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "ollama_up":
                    self._apply_models(*value)
                elif kind == "ollama_down":
                    self.status.configure(text="Ollama arrêté", foreground="#b4232b")
                    self.set_message("Ollama ne répond pas. Lancez-le :  ollama serve",
                                     error=True)
                    self.go_btn.state(["disabled"])
                elif kind == "thinking":
                    self.go_btn.configure(text="Réflexion…" if value else "Relecture…")
                elif kind == "done":
                    self._finish(*value)
                elif kind == "failed":
                    self.busy = False
                    self.go_btn.state(["!disabled"])
                    self.go_btn.configure(text="Réécrire")
                    self.set_message(value, error=True)
                elif kind == "maj":
                    self._maj_trouvee(value)
                elif kind == "maj_prete":
                    self._maj_prete(value)
                elif kind == "maj_echec":
                    pass       # un telechargement rate ne gene pas l usage
                elif kind == "hotkey":
                    self._on_hotkey()
        except queue.Empty:
            pass
        self.root.after(60, self._poll_events)

    def _select_active_window(self):
        """Presélectionne la fenetre actuellement au premier plan.

        Si c'est Ollook lui-meme — l'utilisateur vient d'y cliquer — on garde la
        cible precedente, sinon on perdrait la fenetre qu'il visait."""
        if not self.bridge:
            return
        hwnd = self.bridge.foreground()
        if hwnd and hwnd not in self.bridge.own_hwnds:
            self.target_hwnd = hwnd

        # La liste se limitant a Outlook, une cible etrangere y serait une
        # anomalie : on l'ecarte plutot que de l'y reinjecter.
        if (self.target_hwnd
                and self.bridge.process_name(self.target_hwnd) not in OUTLOOK_PROCESSES):
            self.target_hwnd = None

        self.refresh_windows(keep=True)

        # A defaut de fenetre active pertinente, la fenetre Outlook la plus
        # recemment utilisee est le choix le plus probable.
        if self.target_hwnd is None and self.windows:
            self.target_hwnd = self.windows[0][0]
            self.window_box.current(1)

        if self.target_hwnd:
            titre = self.bridge.window_title(self.target_hwnd) or "(sans titre)"
            self.target.configure(text="Cible : %s" % titre[:60])
        else:
            self.target.configure(text="Aucune fenêtre Outlook ouverte.")

    def _on_hotkey(self):
        """Le raccourci designe la fenetre de travail : celle qui etait au
        premier plan quand l'utilisateur l'a presse."""
        if self.busy:
            return
        self._refresh_models_async()
        self._select_active_window()

        ok, why = self._grab()
        self._show_capture()
        self.show()
        if not ok:
            self.set_message(why, error=True)

    # ---- demarrage --------------------------------------------------------

    def start(self):
        if self.bridge:
            def listener():
                if not self.bridge.listen(lambda: self.events.put(("hotkey", None))):
                    self.events.put(("failed",
                                     "Raccourci %s déjà pris par une autre application. "
                                     "Utilisez la liste des fenêtres ci-dessus."
                                     % HOTKEY_LABEL))
            threading.Thread(target=listener, daemon=True).start()

            # La fenetre active AVANT l'apparition d'Ollook est la cible la plus
            # probable : c'est celle que l'utilisateur regardait.
            self._select_active_window()
            self.show()
            self.set_message("Fenêtre cible présélectionnée. Vérifiez-la, puis Réécrire. "
                             "%s la redésigne à tout moment." % HOTKEY_LABEL)
        else:
            self._absorb(clipboard_get(self.root))
            self._show_capture()
            self.show()
            if not self.draft.strip():
                self.set_message("Le presse-papiers est vide. Copiez votre texte (Cmd+C), "
                                 "puis cliquez sur Réécrire.", error=True)

        self.root.mainloop()


# --------------------------------------------------------------------------
# Mode filtre : lit stdin, ecrit le resultat sur stdout.
# Moteur des Actions rapides macOS (Automator), qui remplacent directement
# le texte selectionne par la sortie du script.
# --------------------------------------------------------------------------

def run_check():
    """Auto-diagnostic. Sans console — cas de l'executable compile — le rapport
    part dans un fichier plutot que d'etre perdu."""
    import json
    import ollook_setup as setup

    installe, vivant, modeles = setup.diagnose()
    rapport = {
        "version": version.VERSION,
        "depot": version.DEPOT,
        "gele": bool(getattr(sys, "frozen", False)),
        "python": sys.version.split()[0],
        "plateforme": sys.platform,
        "tkinter": True,
        "ollama_installe": installe,
        "ollama_chemin": setup.find_ollama() or "",
        "ollama_demarre": vivant,
        "modeles": [m["name"] for m in setup.installed_models()],
        "modele_par_defaut": setup.DEFAULT_MODEL,
        "pret": setup.ready(),
        "preferences": PREFS_PATH,
        "portable": os.path.dirname(PREFS_PATH) == app_dir(),
        "dossier_application": app_dir(),
    }
    texte = json.dumps(rapport, ensure_ascii=False, indent=1)

    if sys.stdout is not None:
        sys.stdout.write(texte + "\n")
    else:
        chemin = os.path.join(os.environ.get("TEMP") or ".", "ollook-check.json")
        with open(chemin, "w", encoding="utf-8") as f:
            f.write(texte)
    return 0


def run_filter():
    # Un executable compile en mode fenetre n'a pas de flux standard.
    if sys.stdin is None or sys.stdout is None:
        return 1

    # Sans cela, Windows encode la sortie dans la page de code de la console et
    # echoue sur tout caractere qui n'y figure pas. Automator attend de l'UTF-8.
    # En sortie, newline="" empeche aussi la traduction \n -> \r\n : le texte
    # rendu doit avoir exactement les fins de ligne du texte recu. En entree on
    # garde la normalisation, pour que \r\n devienne \n.
    if hasattr(sys.stdin, "reconfigure"):
        try:
            sys.stdin.reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass
    for flux in (sys.stdout, sys.stderr):
        if hasattr(flux, "reconfigure"):
            try:
                flux.reconfigure(encoding="utf-8", errors="replace", newline="")
            except (ValueError, OSError):
                pass

    text = sys.stdin.read()

    # En Action rapide, la sortie REMPLACE le texte selectionne : toute erreur
    # non rattrapee le detruirait. Quoi qu'il arrive, on renvoie au moins
    # l'original.
    try:
        draft, signature, context = core.split_message(text)
        if not draft.strip():
            sys.stdout.write(text)
            return 0

        prefs = load_prefs()
        installed = [m["name"] for m in core.list_models()]
        if not installed:
            sys.stderr.write("Aucun modele Ollama installe.\n")
            sys.stdout.write(text)
            return 1

        # Une preference obsolete ne doit pas faire echouer la relecture.
        model = prefs.get("model")
        if model not in installed:
            model = installed[0]

        out = ""
        for kind, value in core.rewrite_stream(
                model, draft, context, prefs["register"],
                prefs.get("translate") == "1", prefs["develop"]):
            if kind == "done":
                out = value

        if not out or out == "(rien a reecrire)":
            sys.stdout.write(text)
            return 1

        sys.stdout.write(core.join_message(out, signature, context))
        return 0

    except Exception as e:
        sys.stderr.write("Ollook : %s\n" % e)
        sys.stdout.write(text)
        return 1


# --------------------------------------------------------------------------
# Instance unique
#
# Relancer Ollook ne doit pas empiler les instances : seule la premiere obtient
# le raccourci global, les suivantes echouent en silence. Une deuxieme
# execution se contente donc de declencher le raccourci de celle qui tourne,
# puis se retire. C'est ce qui permet a un bouton Outlook de se resumer a
# « lancer ollook_app.py », qu'il tourne deja ou non.
# --------------------------------------------------------------------------

MUTEX_NAME = "Local\\OllookSingleInstance"
ERROR_ALREADY_EXISTS = 183


def acquire_single_instance():
    """Renvoie le mutex si nous sommes la premiere instance, None sinon
    (auquel cas le raccourci de l'instance en place a ete declenche)."""
    if not WINDOWS:
        return True

    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.CreateMutexW.restype = ctypes.c_void_p
    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        return True                      # en cas de doute, on demarre
    if kernel32.GetLastError() != ERROR_ALREADY_EXISTS:
        return handle

    # Une instance repond deja : on lui envoie Ctrl+Alt+R.
    user32 = ctypes.windll.user32
    VK_CONTROL, VK_MENU, VK_R, KEYUP = 0x11, 0x12, 0x52, 0x0002
    for vk, flag in ((VK_CONTROL, 0), (VK_MENU, 0), (VK_R, 0),
                     (VK_R, KEYUP), (VK_MENU, KEYUP), (VK_CONTROL, KEYUP)):
        user32.keybd_event(vk, 0, flag, 0)
    return None


def foreground_before_startup():
    """Fenetre au premier plan avant que tkinter n'existe. A relever ici :
    des que la fenetre Tk est creee, c'est elle qui passe devant."""
    if not WINDOWS:
        return None
    import ctypes
    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = ctypes.c_void_p
    return int(user32.GetForegroundWindow() or 0) or None


def main():
    if "--check" in sys.argv:
        return run_check()
    if "--filter" in sys.argv:
        return run_filter()

    guard = acquire_single_instance()
    if guard is None:
        return 0

    declarer_application()
    # La fenetre active doit etre relevee avant toute fenetre Tk, assistant compris.
    cible = foreground_before_startup()

    import ollook_setup as setup
    if not setup.ready():
        if not setup.run_setup_ui():
            return 1

    App(initial_target=cible).start()
    return 0


if __name__ == "__main__":
    sys.exit(main())
