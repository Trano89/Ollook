// =====================================================================
//  Ollook - complement COM pour Outlook classique
// =====================================================================
//
//  Ajoute un bouton "Relire" dans le ruban d'Outlook, a la fois dans la
//  fenetre principale et dans les fenetres de redaction. Le bouton lance
//  Ollook.exe ; si Ollook tourne deja, l'instance en place est reveillee.
//
//  Contrairement a un add-in web, ce complement s'enregistre dans le
//  registre local de l'utilisateur : il ne passe ni par Exchange ni par
//  Microsoft 365, ne demande aucun droit administrateur, et fonctionne
//  donc sur tous les comptes -- y compris IMAP et Gmail.
//
//  Les interfaces d'Office sont redeclarees ici plutot qu'importees d'une
//  bibliotheque d'interoperabilite primaire (PIA) : cela evite toute
//  dependance et permet de compiler avec le csc.exe livre avec Windows.
//
//  Compilation et installation : voir install-outlook-addin.ps1
// =====================================================================

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Ollook - complement Outlook")]
[assembly: AssemblyProduct("Ollook")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace Ollook
{
    /// <summary>Interface que tout complement COM Office doit implementer.</summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection(object application, int connectMode, object addInInst, ref Array custom);
        void OnDisconnection(int removeMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    /// <summary>Permet au complement de fournir sa propre portion de ruban.</summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        string GetCustomUI(string ribbonID);
    }

    /// <summary>
    /// Point d'entree du complement.
    /// ClassInterfaceType.AutoDual est necessaire : Office appelle les
    /// rappels du ruban (OnRelire) par liaison tardive via IDispatch, ce qui
    /// suppose que ces methodes soient exposees sur l'interface de classe.
    /// </summary>
    [ComVisible(true)]
    [Guid("7C4A3F1E-9B2D-4E58-A0C6-3D5B8E1F2A94")]
    [ProgId("Ollook.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private const string LABEL = "Relire";
        private const string TIP = "Faire relire et reecrire ce message par Ollook, "
                                 + "avec un modele tournant sur cet ordinateur.";

        // ---- IDTExtensibility2 : rien a faire, le bouton suffit ----------

        public void OnConnection(object application, int connectMode,
                                 object addInInst, ref Array custom) { }
        public void OnDisconnection(int removeMode, ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ---- Ruban -------------------------------------------------------

        public string GetCustomUI(string ribbonID)
        {
            // TabMail : onglet Accueil de la fenetre principale.
            // TabNewMailMessage : onglet Message d'une fenetre de redaction.
            // La fenetre de LECTURE est volontairement exclue : Ollook reecrit
            // un brouillon, pas un message recu. Son onglet porte d'ailleurs un
            // autre identifiant (TabReadMessage).
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    return Ribbon("TabMail");
                case "Microsoft.Outlook.Mail.Compose":
                    return Ribbon("TabNewMailMessage");
                default:
                    return string.Empty;
            }
        }

        private static string Ribbon(string tabIdMso)
        {
            return
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                  "<ribbon>" +
                    "<tabs>" +
                      "<tab idMso=\"" + tabIdMso + "\">" +
                        "<group id=\"OllookGroup\" label=\"Ollook\">" +
                          "<button id=\"OllookRelire\" label=\"" + LABEL + "\" size=\"large\"" +
                                 " imageMso=\"SpellingAndGrammar\" onAction=\"OnRelire\"" +
                                 " screentip=\"Ollook\" supertip=\"" + TIP + "\"/>" +
                        "</group>" +
                      "</tab>" +
                    "</tabs>" +
                  "</ribbon>" +
                "</customUI>";
        }

        /// <summary>Rappel du bouton. Office le trouve par son nom, via IDispatch.</summary>
        public void OnRelire(object control)
        {
            string exe = FindExecutable();
            if (exe == null)
            {
                MessageBox.Show(
                    "Ollook.exe est introuvable.\n\n" +
                    "Relancez install-outlook-addin.ps1 depuis le dossier d'Ollook " +
                    "pour reenregistrer le chemin.",
                    "Ollook", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Ollook n'accepte qu'une instance : ce lancement reveille
                // celle qui tourne deja, le cas echeant.
                Process.Start(new ProcessStartInfo(exe)
                {
                    UseShellExecute = true,
                    WorkingDirectory = Path.GetDirectoryName(exe)
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lancement impossible :\n" + ex.Message,
                                "Ollook", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ---- Localisation de l'executable --------------------------------

        private static string FindExecutable()
        {
            // 1. Chemin ecrit par l'installateur : la source de verite.
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Ollook"))
                {
                    if (key != null)
                    {
                        string path = key.GetValue("ExePath") as string;
                        if (!string.IsNullOrEmpty(path) && File.Exists(path))
                            return path;
                    }
                }
            }
            catch { /* registre illisible : on tente les emplacements connus */ }

            // 2. Emplacements habituels, relatifs a la DLL du complement.
            try
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string[] candidates =
                {
                    "Ollook.exe",
                    Path.Combine("dist", "Ollook.exe"),
                    Path.Combine("..", "dist", "Ollook.exe"),
                };
                foreach (string relative in candidates)
                {
                    string full = Path.GetFullPath(Path.Combine(dir, relative));
                    if (File.Exists(full))
                        return full;
                }
            }
            catch { /* rien de plus a tenter */ }

            return null;
        }
    }
}
