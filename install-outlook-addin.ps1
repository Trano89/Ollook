<#
.SYNOPSIS
    Compile et installe le complement COM Ollook dans Outlook classique.

.DESCRIPTION
    Ajoute un bouton "Relire" au ruban d'Outlook, dans la fenetre principale
    comme dans les fenetres de redaction.

    Tout est LOCAL : l'enregistrement se fait dans HKEY_CURRENT_USER, sans
    passer par Exchange ni Microsoft 365. Le complement fonctionne donc sur
    tous les comptes, y compris IMAP et Gmail, et n'exige aucun droit
    administrateur.

    Le compilateur C# utilise est celui livre avec Windows
    (C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe) : rien a
    installer.

.PARAMETER Uninstall
    Retire le complement et toutes ses cles de registre.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\install-outlook-addin.ps1
    powershell -ExecutionPolicy Bypass -File .\install-outlook-addin.ps1 -Uninstall

.NOTES
    Le nouvel Outlook (olk.exe) ne charge aucun complement COM : ce bouton
    n'apparait que dans Outlook classique. Il n'apparait pas non plus dans
    "Toutes les applications", surface reservee aux complements web
    installes via la boite aux lettres.
#>

[CmdletBinding()]
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'

$Root      = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Source    = Join-Path $Root 'outlook-addin\OllookAddin.cs'
$OutDir    = Join-Path $Root 'outlook-addin'
$Dll       = Join-Path $OutDir 'OllookAddin.dll'
$Exe       = Join-Path $Root 'dist\Ollook.exe'

$ProgId    = 'Ollook.Connect'
$Clsid     = '{7C4A3F1E-9B2D-4E58-A0C6-3D5B8E1F2A94}'
$AsmName   = 'OllookAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$Class     = 'Ollook.Connect'
$Runtime   = 'v4.0.30319'

$KeyClasses = "HKCU:\Software\Classes"
$KeyAddin   = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$ProgId"
$KeyOllook  = "HKCU:\Software\Ollook"

function Say ([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c }
function Ok  ([string]$m) { Say "  [ok]   $m" 'Green' }
function Warn([string]$m) { Say "  [!]    $m" 'Yellow' }

Say ''
Say '  Ollook - complement COM pour Outlook classique' 'Cyan'
Say '  ---------------------------------------------' 'Cyan'

# --------------------------------------------------------------------------
# Outlook doit etre ferme : il verrouille la DLL une fois chargee.
# --------------------------------------------------------------------------

if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
    Say ''
    Warn 'Outlook est ouvert. Fermez-le completement, puis relancez ce script.'
    Warn '(il verrouille la DLL du complement tant qu''il tourne)'
    Say ''
    return
}

# --------------------------------------------------------------------------
# Desinstallation
# --------------------------------------------------------------------------

if ($Uninstall) {
    foreach ($k in @($KeyAddin,
                     "$KeyClasses\$ProgId",
                     "$KeyClasses\CLSID\$Clsid")) {
        if (Test-Path $k) { Remove-Item $k -Recurse -Force; Ok "Retire : $k" }
    }
    if (Test-Path $KeyOllook) {
        Remove-ItemProperty -Path $KeyOllook -Name 'ExePath' -ErrorAction SilentlyContinue
        Ok 'Chemin de l''executable retire.'
    }
    if (Test-Path $Dll) { Remove-Item $Dll -Force; Ok 'DLL supprimee.' }
    Say ''
    Say '  Termine. Rouvrez Outlook.' 'Cyan'
    Say ''
    return
}

# --------------------------------------------------------------------------
# 1. Compilation
# --------------------------------------------------------------------------

$Csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $Csc)) {
    $Csc = 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path $Csc)) {
    throw "Compilateur C# introuvable. Le .NET Framework 4 doit etre present sur Windows."
}
Ok "Compilateur : $Csc"

if (-not (Test-Path $Source)) { throw "Source introuvable : $Source" }

# /platform:anycpu : la DLL se charge dans Outlook 32 comme 64 bits.
& $Csc /nologo /target:library /platform:anycpu /optimize+ `
       /out:"$Dll" `
       /reference:System.dll `
       /reference:System.Windows.Forms.dll `
       "$Source"

if ($LASTEXITCODE -ne 0 -or -not (Test-Path $Dll)) {
    throw "La compilation a echoue."
}
Ok "DLL compilee : $Dll ($([math]::Round((Get-Item $Dll).Length / 1KB, 1)) Ko)"

# --------------------------------------------------------------------------
# 2. Enregistrement COM dans HKCU
#
#    Equivalent de "RegAsm /codebase", mais ecrit sous HKCU\Software\Classes
#    au lieu de HKCR : aucun droit administrateur n'est necessaire.
# --------------------------------------------------------------------------

function Set-Key([string]$path) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}
function Set-Val([string]$path, [string]$name, $value, [string]$type = 'String') {
    Set-Key $path
    New-ItemProperty -Path $path -Name $name -Value $value -PropertyType $type -Force | Out-Null
}

$CodeBase = ([Uri](Get-Item $Dll).FullName).AbsoluteUri

Set-Val "$KeyClasses\$ProgId"        '(default)' 'Ollook'
Set-Val "$KeyClasses\$ProgId\CLSID"  '(default)' $Clsid

Set-Val "$KeyClasses\CLSID\$Clsid"        '(default)' 'Ollook'
Set-Val "$KeyClasses\CLSID\$Clsid\ProgId" '(default)' $ProgId

# Categorie "composant gere", attendue par le moteur COM pour un objet .NET.
Set-Key "$KeyClasses\CLSID\$Clsid\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

foreach ($server in @("$KeyClasses\CLSID\$Clsid\InprocServer32",
                      "$KeyClasses\CLSID\$Clsid\InprocServer32\1.0.0.0")) {
    Set-Val $server '(default)'      "$env:SystemRoot\System32\mscoree.dll"
    Set-Val $server 'ThreadingModel' 'Both'
    Set-Val $server 'Class'          $Class
    Set-Val $server 'Assembly'       $AsmName
    Set-Val $server 'RuntimeVersion' $Runtime
    Set-Val $server 'CodeBase'       $CodeBase
}
Ok 'Objet COM enregistre pour cet utilisateur (sans droits administrateur).'

# --------------------------------------------------------------------------
# 3. Declaration a Outlook
# --------------------------------------------------------------------------

Set-Val $KeyAddin 'FriendlyName' 'Ollook'
Set-Val $KeyAddin 'Description'  'Relecture de courriels par un modele local (Ollama).'
Set-Val $KeyAddin 'LoadBehavior' 3 'DWord'      # 3 = charge au demarrage
Set-Val $KeyAddin 'CommandLineSafe' 0 'DWord'
Ok "Complement declare : $KeyAddin"

# --------------------------------------------------------------------------
# 4. Chemin de l'executable, lu par le complement
# --------------------------------------------------------------------------

if (Test-Path $Exe) {
    Set-Val $KeyOllook 'ExePath' (Get-Item $Exe).FullName
    Ok "Executable : $Exe"
} else {
    Warn "dist\Ollook.exe est absent. Compilez-le d'abord :  python build_exe.py"
    Warn "Le bouton apparaitra, mais signalera que l'executable est introuvable."
}

# --------------------------------------------------------------------------

Say ''
Say '  Installation terminee.' 'Cyan'
Say ''
Say '  Ouvrez Outlook : le groupe "Ollook" et son bouton "Relire" apparaissent'
Say '  dans l''onglet Accueil, et dans l''onglet Message des fenetres de redaction.'
Say ''
Say '  Si le bouton n''apparait pas :' 'DarkGray'
Say '    Fichier > Options > Complements > Gerer : Complements COM > Atteindre' 'DarkGray'
Say '    Ollook doit y figurer et etre coche. S''il est dans "Elements desactives",' 'DarkGray'
Say '    reactivez-le depuis la meme boite de dialogue.' 'DarkGray'
Say ''
Say '  Rappel : le nouvel Outlook ne charge aucun complement COM.' 'DarkGray'
Say ''
