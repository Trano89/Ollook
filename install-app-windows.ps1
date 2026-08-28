<#
.SYNOPSIS
    Lance l'utilitaire autonome Ollook a chaque ouverture de session.

.DESCRIPTION
    L'utilitaire autonome ne demande aucune installation : il suffit de lancer
    Ollook-App.cmd. Ce script ne fait qu'ajouter (ou retirer) un raccourci dans
    le dossier Demarrage, pour que le raccourci Ctrl+Alt+R soit toujours actif.

    Aucun droit administrateur, aucun certificat, aucun serveur, aucun Exchange.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install-app-windows.ps1
    powershell -ExecutionPolicy Bypass -File .\install-app-windows.ps1 -Uninstall
#>

[CmdletBinding()]
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Lnk  = Join-Path ([Environment]::GetFolderPath('Startup')) 'Ollook.lnk'

function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c }

Say ''
Say '  Ollook - utilitaire autonome' 'Cyan'
Say '  ----------------------------' 'Cyan'

if ($Uninstall) {
    if (Test-Path $Lnk) { Remove-Item $Lnk -Force; Say '  [ok]   Demarrage automatique retire.' 'Green' }
    else                { Say '  [!]    Aucun demarrage automatique installe.' 'Yellow' }
    Say ''
    return
}

# --- Cible du raccourci ---------------------------------------------------
#
# L'executable compile d'abord : il porte l'icone de l'application et ne
# depend d'aucune installation de Python. On ne retombe sur le script que
# s'il n'a pas encore ete construit.

$Exe = Join-Path $Root 'dist\Ollook.exe'
$cible = $null
$arguments = ''

if (Test-Path $Exe) {
    $cible = (Get-Item $Exe).FullName
    Say "  [ok]   Executable : $cible" 'Green'
} else {
    Say ''
    Say '  [!]    dist\Ollook.exe est absent. Construisez-le :  python build_exe.py' 'Yellow'
    Say '         En attendant, le raccourci utilisera le script Python.' 'Yellow'
    Say ''

    $py = $null
    foreach ($c in @('python', 'py', 'python3')) {
        $found = Get-Command $c -ErrorAction SilentlyContinue
        if ($found) { $py = $found.Source; break }
    }
    if (-not $py) {
        throw "Ni dist\Ollook.exe ni Python. Lancez d'abord : python build_exe.py"
    }

    # Pas de '2>&1' ici : en PowerShell 5.1, rediriger la sortie d'erreur d'un
    # executable natif enveloppe chaque ligne dans un ErrorRecord et fausse l'etat.
    # Et pas de pythonw.exe : sans console, il ne renvoie rien d'exploitable.
    $tkVersion = ''
    try { $tkVersion = (& $py -c "import tkinter; print(tkinter.TkVersion)") } catch { $tkVersion = '' }
    if ($LASTEXITCODE -ne 0 -or -not $tkVersion) {
        Say '  [!]    tkinter est absent de cette installation de Python.' 'Yellow'
        Say '         Reinstallez Python en cochant "tcl/tk and IDLE".' 'Yellow'
        Say ''
        return
    }
    Say "  [ok]   Python $py, tkinter $tkVersion" 'Green'

    # pythonw.exe (ou pyw.exe) : pas de fenetre de console noire en arriere-plan.
    $pyDir = Split-Path $py
    $cible = Join-Path $pyDir 'pythonw.exe'
    if (-not (Test-Path $cible)) { $cible = Join-Path $pyDir 'pyw.exe' }
    if (-not (Test-Path $cible)) { $cible = $py }
    $arguments = '"' + (Join-Path $Root 'ollook_app.py') + '"'
}

$shell = New-Object -ComObject WScript.Shell
$lnkObj = $shell.CreateShortcut($Lnk)
$lnkObj.TargetPath       = $cible
$lnkObj.Arguments        = $arguments
$lnkObj.WorkingDirectory = Split-Path $cible
$lnkObj.Description      = 'Ollook - relecture de courriels (Ctrl+Alt+R)'
$lnkObj.Save()

Say "  [ok]   Demarrage automatique installe : $Lnk" 'Green'
Say ''
Say '  Utilisation :' 'White'
Say '    1. Assurez-vous qu''Ollama tourne   ->  ollama serve'
Say '    2. Lancez maintenant                ->  .\Ollook-App.cmd'
Say '    3. Selectionnez du texte, n''importe ou, puis Ctrl+Alt+R'
Say ''
Say '  Fonctionne dans Outlook classique, Outlook web, Word, un navigateur,' 'DarkGray'
Say '  sur le compte Exchange comme sur le compte Gmail.' 'DarkGray'
Say ''
