' =====================================================================
'  Ollook - bouton de lancement depuis Outlook
' =====================================================================
'
'  Pourquoi une macro plutot qu'un complement : les complements web
'  Outlook s'installent sur la boite aux lettres Exchange et sont
'  souvent bloques en entreprise. Une macro VBA est locale au poste,
'  ne demande aucun droit administrateur, et fonctionne aussi bien sur
'  un compte Exchange que sur un compte IMAP ou Gmail.
'
'  La macro se contente de lancer Ollook.exe. Si Ollook tourne deja,
'  la nouvelle execution reveille celle en place puis se retire : il
'  n'y a jamais deux fenetres.
'
'  Dans les deux cas, Ollook presélectionne la fenetre Outlook depuis
'  laquelle vous avez clique -- fenetre principale ou fenetre de
'  redaction, indifferemment.
'
' ---------------------------------------------------------------------
'  INSTALLATION
' ---------------------------------------------------------------------
'
'  1. Autoriser les macros
'     Outlook > Fichier > Options > Centre de gestion de la confidentialite
'     > Parametres du Centre de gestion de la confidentialite
'     > Parametres des macros > "Notifications pour toutes les macros"
'     Redemarrez Outlook. A chaque demarrage, acceptez l'invite.
'
'  2. Coller cette macro
'     Alt+F11 ouvre l'editeur Visual Basic.
'     Menu Insertion > Module, puis collez tout ce fichier.
'     Ctrl+S pour enregistrer, puis fermez l'editeur.
'
'  3. Verifier le chemin ci-dessous (constante OLLOOK).
'     Executez la macro une fois depuis l'editeur (touche F5) : la
'     fenetre Ollook doit apparaitre.
'
'  4. Ajouter le bouton -- A FAIRE DEUX FOIS, une par type de fenetre
'
'     a) Fenetre principale : Fichier > Options > Personnaliser le ruban
'     b) Fenetre de redaction : ouvrez un nouveau message, puis dans
'        CETTE fenetre Fichier > Options > Personnaliser le ruban
'
'     Dans les deux cas :
'       - "Choisir les commandes dans les categories suivantes" : Macros
'       - selectionnez Ollook.Relire
'       - a droite, creez un groupe (Nouveau groupe) dans l'onglet voulu
'         -- Accueil pour la fenetre principale, Message pour la
'         redaction -- puis Ajouter.
'
'     Variante plus rapide, disponible dans les deux fenetres :
'     clic droit sur le ruban > "Personnaliser la barre d'outils Acces
'     rapide" > categorie Macros > Ajouter.
'
' =====================================================================

Option Explicit

' --- A ADAPTER SI VOTRE INSTALLATION DIFFERE ------------------------
' Chemin de l'executable autonome. C'est la seule ligne a verifier.
Private Const OLLOOK As String = "C:\Users\atrottet\Ollook\dist\Ollook.exe"
' --------------------------------------------------------------------


' Point d'entree : c'est cette macro qu'il faut placer sur le ruban.
Public Sub Relire()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(OLLOOK) Then
        MsgBox "Ollook est introuvable :" & vbCrLf & OLLOOK & vbCrLf & vbCrLf & _
               "Corrigez la constante OLLOOK dans la macro (Alt+F11).", _
               vbExclamation, "Ollook"
        Exit Sub
    End If

    On Error Resume Next
    Shell """" & OLLOOK & """", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "Lancement impossible : " & Err.Description, vbExclamation, "Ollook"
    End If
    On Error GoTo 0
End Sub
