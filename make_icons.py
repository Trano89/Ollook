#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ollook - generateur des icones de l'application.

Dessine une fenetre de message Outlook contenant "#@!" -- les jurons masques
de la bande dessinee : ce qu'Ollook transforme en courriel presentable.

Les formes sont tracees en Python et suréchantillonnées 4x ; le texte est rendu
par GDI, avec une vraie police, puis compose comme un masque de couverture.
Rendu Windows uniquement -- c'est un outil de fabrication, pas d'execution :
les PNG produits dans assets/ sont des ressources statiques versionnees.

    python make_icons.py
"""

import ctypes
import os
import struct
import sys
import zlib
from ctypes import wintypes

RACINE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(RACINE, "assets")
TAILLES = (16, 32, 64, 128, 256)
S = 4                                    # facteur de suréchantillonnage

FOND = (31, 58, 95)                      # bleu nuit, fond de l'icone
FENETRE = (255, 255, 255)                # corps du message
BARRE = (15, 108, 189)                   # bandeau de titre, bleu Outlook
TEXTE = (26, 42, 66)                     # "#@!"
TEXTE_PETIT = (15, 108, 189)             # plus contraste aux petites tailles

GLYPHES = "#@!"
POLICE = "Segoe UI"


# --------------------------------------------------------------------------
# Ecriture PNG (aucune dependance)
# --------------------------------------------------------------------------

def ecrire_png(chemin, taille, lignes):
    def bloc(tag, data):
        c = tag + data
        return (struct.pack(">I", len(data)) + c
                + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF))

    brut = b"".join(b"\x00" + lignes[y] for y in range(taille))
    with open(chemin, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(bloc(b"IHDR", struct.pack(">IIBBBBB", taille, taille, 8, 6, 0, 0, 0)))
        f.write(bloc(b"IDAT", zlib.compress(brut, 9)))
        f.write(bloc(b"IEND", b""))


# --------------------------------------------------------------------------
# Rendu du texte par GDI : couverture 0..255 sur une grille n x n
# --------------------------------------------------------------------------

class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG),
                ("biHeight", wintypes.LONG), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", wintypes.LONG),
                ("biYPelsPerMeter", wintypes.LONG), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


def masque_texte(n, texte, rect, hauteur_police):
    """Dessine `texte` centre dans `rect` et renvoie un masque de couverture.

    Le texte est trace en noir sur blanc : la couverture d'un pixel est le
    complement de sa luminance. GDI se charge de l'anticrenelage."""
    gdi32 = ctypes.windll.gdi32
    user32 = ctypes.windll.user32

    gdi32.CreateDIBSection.restype = ctypes.c_void_p
    gdi32.CreateCompatibleDC.restype = ctypes.c_void_p
    gdi32.CreateFontW.restype = ctypes.c_void_p
    gdi32.SelectObject.restype = ctypes.c_void_p

    info = BITMAPINFO()
    info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    info.bmiHeader.biWidth = n
    info.bmiHeader.biHeight = -n              # negatif : origine en haut
    info.bmiHeader.biPlanes = 1
    info.bmiHeader.biBitCount = 32
    info.bmiHeader.biCompression = 0          # BI_RGB

    dc = gdi32.CreateCompatibleDC(None)
    bits = ctypes.c_void_p()
    bitmap = gdi32.CreateDIBSection(ctypes.c_void_p(dc), ctypes.byref(info), 0,
                                   ctypes.byref(bits), None, 0)
    if not bitmap:
        raise RuntimeError("CreateDIBSection a echoue")
    gdi32.SelectObject(ctypes.c_void_p(dc), ctypes.c_void_p(bitmap))

    # Fond blanc.
    ctypes.memset(bits, 0xFF, n * n * 4)

    ANTIALIASED_QUALITY, FW_BOLD, DEFAULT_CHARSET = 4, 700, 1
    police = gdi32.CreateFontW(-int(hauteur_police), 0, 0, 0, FW_BOLD, 0, 0, 0,
                               DEFAULT_CHARSET, 0, 0, ANTIALIASED_QUALITY, 0,
                               POLICE)
    gdi32.SelectObject(ctypes.c_void_p(dc), ctypes.c_void_p(police))
    gdi32.SetBkMode(ctypes.c_void_p(dc), 1)               # TRANSPARENT
    gdi32.SetTextColor(ctypes.c_void_p(dc), 0x000000)

    zone = wintypes.RECT(int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3]))
    DT_CENTER, DT_VCENTER, DT_SINGLELINE, DT_NOCLIP = 0x1, 0x4, 0x20, 0x100
    user32.DrawTextW(ctypes.c_void_p(dc), texte, -1, ctypes.byref(zone),
                     DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOCLIP)

    donnees = ctypes.string_at(bits, n * n * 4)
    gdi32.DeleteObject(ctypes.c_void_p(police))
    gdi32.DeleteObject(ctypes.c_void_p(bitmap))
    gdi32.DeleteDC(ctypes.c_void_p(dc))

    # Canal bleu de chaque pixel BGRA : 255 = fond, 0 = plein trait.
    return bytes(255 - donnees[i * 4] for i in range(n * n))


# --------------------------------------------------------------------------
# Composition
# --------------------------------------------------------------------------

def dans_rect_arrondi(x, y, x0, y0, x1, y1, r):
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    dx = max(x0 + r - x, x - (x1 - r), 0.0)
    dy = max(y0 + r - y, y - (y1 - r), 0.0)
    return dx * dx + dy * dy <= r * r


def fabriquer(taille):
    n = taille * S
    petit = taille <= 32
    couleur_texte = TEXTE_PETIT if petit else TEXTE

    # Geometrie, en fraction du cote. Sous 32 px les marges mangent la
    # lisibilite : la fenetre occupe alors presque toute l'icone.
    r_fond = n * (0.20 if petit else 0.22)
    if petit:
        fx0, fy0, fx1, fy1 = n * 0.085, n * 0.165, n * 0.915, n * 0.835
        ratio_barre, marge, ratio_police = 0.21, n * 0.02, 1.02
    else:
        fx0, fy0, fx1, fy1 = n * 0.145, n * 0.205, n * 0.855, n * 0.755
        ratio_barre, marge, ratio_police = 0.24, n * 0.045, 0.86

    r_fen = n * (0.035 if petit else 0.05)
    h_barre = (fy1 - fy0) * ratio_barre

    zone_texte = (fx0 + marge, fy0 + h_barre, fx1 - marge, fy1 - marge * 0.4)
    hauteur_police = (zone_texte[3] - zone_texte[1]) * ratio_police
    masque = masque_texte(n, GLYPHES, zone_texte, hauteur_police)

    lignes = []
    for py in range(taille):
        ligne = bytearray()
        for px in range(taille):
            sr = sg = sb = sa = 0.0
            for sy in range(S):
                y = py * S + sy + 0.5
                for sx in range(S):
                    x = px * S + sx + 0.5

                    if not dans_rect_arrondi(x, y, 0, 0, n - 1, n - 1, r_fond):
                        continue                       # hors de l'icone
                    sa += 1.0

                    if dans_rect_arrondi(x, y, fx0, fy0, fx1, fy1, r_fen):
                        base = BARRE if y < fy0 + h_barre else FENETRE
                    else:
                        base = FOND

                    # Le texte ne s'imprime que dans le corps du message.
                    couverture = 0.0
                    if y >= fy0 + h_barre:
                        couverture = masque[int(y) * n + int(x)] / 255.0

                    sr += base[0] * (1 - couverture) + couleur_texte[0] * couverture
                    sg += base[1] * (1 - couverture) + couleur_texte[1] * couverture
                    sb += base[2] * (1 - couverture) + couleur_texte[2] * couverture

            total = float(S * S)
            alpha = sa / total
            if alpha <= 0:
                ligne += bytes((0, 0, 0, 0))
                continue
            ligne += bytes((int(round(sr / sa)), int(round(sg / sa)),
                            int(round(sb / sa)), int(round(alpha * 255))))
        lignes.append(bytes(ligne))
    return lignes


def ecrire_module_python():
    """Ecrit ollook_icon.py : les PNG encodes en base64.

    C'est ce qui permet a la fenetre d'avoir son icone dans l'executable
    portable, sans fichier a deployer a cote."""
    import base64

    lignes = [
        "# -*- coding: utf-8 -*-",
        '"""Icones de l\'application, encodees en base64.',
        "",
        "Genere par make_icons.py -- ne pas modifier a la main.",
        "Embarquer les images dans le code evite tout fichier annexe : la",
        "fenetre garde son icone meme quand l'executable est copie seul.",
        '"""',
        "",
        "ICONES = {",
    ]
    for taille in TAILLES:
        with open(os.path.join(ASSETS, "icon-%d.png" % taille), "rb") as f:
            encode = base64.b64encode(f.read()).decode("ascii")
        lignes.append("    %d: (" % taille)
        for i in range(0, len(encode), 76):
            lignes.append('        "%s"' % encode[i:i + 76])
        lignes.append("    ),")
    lignes.append("}")
    lignes.append("")

    chemin = os.path.join(RACINE, "ollook_icon.py")
    with open(chemin, "w", encoding="utf-8") as f:
        f.write("\n".join(lignes))
    return chemin


def main():
    if not sys.platform.startswith("win"):
        sys.exit("Le rendu du texte utilise GDI : ce generateur ne tourne que sur Windows.\n"
                 "Les PNG d'assets/ sont versionnes, il n'est pas necessaire de les refaire.")

    os.makedirs(ASSETS, exist_ok=True)
    for taille in TAILLES:
        chemin = os.path.join(ASSETS, "icon-%d.png" % taille)
        ecrire_png(chemin, taille, fabriquer(taille))
        print("  %-24s %6d octets" % (os.path.basename(chemin),
                                      os.path.getsize(chemin)))
    module = ecrire_module_python()
    print("  %-24s %6d octets" % (os.path.basename(module),
                                  os.path.getsize(module)))
    print("\nIcônes régénérées. Relancez  python build_exe.py  pour les intégrer.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
