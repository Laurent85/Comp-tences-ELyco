using Compétences.Properties;
using Microsoft.Office.Interop.Excel;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using ListBox = System.Windows.Forms.ListBox;

namespace Compétences
{
    public partial class Principal : Form
    {
        public Message Frm2 = new Message();

        public Principal()
        {
            InitializeComponent();
        }

        private void Principal_Load(object sender, EventArgs e)
        {
            Directory.CreateDirectory("C:\\ELyco");
            File.WriteAllText("C:\\ELyco\\ELyco_classes.txt", String.Empty);
            File.WriteAllText("C:\\ELyco\\ELyco_classes_annee.txt", String.Empty);
            File.WriteAllText("C:\\ELyco\\ELyco_classes_dnb.txt", String.Empty);
            BtnLancerTraitement.Enabled = false;
            RafraichirListbox();
            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(), LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText("C:\\ELyco\\ELyco_classes.txt", Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
            }

            try
            {
                ComboAnnéeScolaire.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(2).Take(3).First();
                ComboNiveau6.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(3).Take(4).First();
                ComboNiveau5.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(4).Take(5).First();
                ComboNiveau4.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(5).Take(6).First();
                ComboNiveau3.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(6).Take(7).First();
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void Dossier_travail_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string folder = dlg.SelectedPath + "\\ELyco_CSV\\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierCsv.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_in.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_in.txt"))
                    {
                    }
                }
                if (!File.Exists("C:\\ELyco\\ELyco_classes.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_classes.txt"))
                    {
                    }
                }
                if (!File.Exists("C:\\ELyco\\ELyco_classes_annee.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_classes_annee.txt"))
                    {
                    }
                }
                if (!File.Exists("C:\\ELyco\\ELyco_classes_dnb.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_classes_dnb.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco\\ELyco_in.txt"))
                {
                    sw.WriteLine(LblCheminDossierCsv.Text);
                    sw.WriteLine(dlg.SelectedPath + "\\ELyco_CSV" + "\n");
                }
            }
        }

        private void Dossier_destination_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string folder = dlg.SelectedPath + "\\ELyco_Competences\\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierXlsx.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_out.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_out.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco\\ELyco_out.txt"))
                {
                    sw.WriteLine(LblCheminDossierXlsx.Text);
                    sw.WriteLine(dlg.SelectedPath + "\\ELyco_Competences");
                }
            }
        }

        private void Créer_arborescence_Click(object sender, EventArgs e)
        {
            Création_arborescence("6");
            Création_arborescence("5");
            Création_arborescence("4");
            Création_arborescence("3");

            Directory.CreateDirectory(LblCheminDossierXlsx.Text + "\\" + "1ère période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + "\\" + "2ème période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + "\\" + "3ème période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + "\\" + "Année");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + "\\" + "DNB");

            Changer_ligne_fichier_txt(ComboAnnéeScolaire.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 3);
            Changer_ligne_fichier_txt(ComboNiveau6.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 4);
            Changer_ligne_fichier_txt(ComboNiveau5.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 5);
            Changer_ligne_fichier_txt(ComboNiveau4.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 6);
            Changer_ligne_fichier_txt(ComboNiveau3.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 7);
        }

        public void LancerTraitementCsv_Click(object sender, EventArgs e)
        {
            VérifierDoublonClasseCsv(DétectionPériode());

            BackgroundWorker traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += TraitementMacroLancement;
            traitementMacro.RunWorkerCompleted += TraitementMacroFini;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            if (LblFichiersCsvATraiter.Text == "")
            {
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text = @"Traitement des fichiers...Veuillez patienter...";
            }
            else
            {
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text = LblFichiersCsvATraiter.Text + @"...Veuillez patienter...";
            }

            Frm2.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = false;
            Frm2.ShowDialog();
        }

        private void TraitementMacroLancement(object sender, DoWorkEventArgs e)
        {
            if (RadioBtnPériode1.Checked)
            {
                Executer_macro("Deplacer_P1.Deplacer_P1");
                Executer_macro("Compétences_par_lot_P1.Compétences_par_lot_P1");
            }
            if (RadioBtnPériode2.Checked)
            {
                Executer_macro("Deplacer_P2.Deplacer_P2");
                Executer_macro("Compétences_par_lot_P2.Compétences_par_lot_P2");
            }
            if (RadioBtnPériode3.Checked)
            {
                Executer_macro("Deplacer_P3.Deplacer_P3");
                Executer_macro("Compétences_par_lot_P3.Compétences_par_lot_P3");
            }
            if (RadioBtnAnnée.Checked)
            {
                Executer_macro("Fusionner.Fusionner");
                Executer_macro("Compétences_par_lot_Année.Compétences_par_lot_Année");
            }
        }

        private void TraitementMacroFini(object sender, RunWorkerCompletedEventArgs e)
        {
            RafraichirListbox();

            string[] files = Directory.GetFiles(LblCheminDossierCsv.Text, "*.*");
            foreach (string file in files)
            {
                File.Delete(file);
            }

            Frm2.Controls.Find("LblMessageTraitement", true).First().Text = @"Traitement des fichiers terminé !";
            Frm2.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void LancerTraitementDnb_Click(object sender, EventArgs e)
        {
            BackgroundWorker traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += TraitementMacroDnbLancement;
            traitementMacro.RunWorkerCompleted += TraitementMacroDnbFini;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            if (LblFichiersCsvATraiter.Text == "")
            {
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text = @"Traitement des fichiers...Veuillez patienter...";
            }
            else
            {
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text = LblFichiersCsvATraiter.Text + @"...Veuillez patienter...";
            }

            Frm2.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = false;
            Frm2.ShowDialog();
        }

        private void TraitementMacroDnbLancement(object sender, DoWorkEventArgs e)
        {
            Executer_macro("Publipostage.Publipostage");
        }

        private void TraitementMacroDnbFini(object sender, RunWorkerCompletedEventArgs e)
        {
            RafraichirListbox();
            Frm2.Controls.Find("LblMessageTraitement", true).First().Text = @"Traitement des fichiers terminé !";
            Frm2.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void bouton_periode1_CheckedChanged(object sender, EventArgs e)
        {
            BtnLancerTraitement.Enabled = true;
        }

        private void bouton_periode2_CheckedChanged(object sender, EventArgs e)
        {
            BtnLancerTraitement.Enabled = true;
        }

        private void bouton_periode3_CheckedChanged(object sender, EventArgs e)
        {
            BtnLancerTraitement.Enabled = true;
        }

        private void bouton_annee_CheckedChanged(object sender, EventArgs e)
        {
            BtnLancerTraitement.Enabled = true;
        }

        private void BtnGénérerfichiersExcelDnb_Click(object sender, EventArgs e)
        {
            GénérerFichiersXlsxDnb();
        }

        private void SuppressionFichierCsv_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListeBoxCsvPrésents, SearchOption.AllDirectories, ListeBoxCsvPrésents.SelectedItem.ToString());
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListeBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListeBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void SuppressionFichierXlsx_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierXlsx.Text, ListeBoxXlsxPrésents, SearchOption.AllDirectories, ListeBoxXlsxPrésents.SelectedItem.ToString());
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListeBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListeBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void SuppressionBases_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show(@"Etes-vous sûr de vouloir tout supprimer ?", @"Attention !", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListeBoxCsvPrésents, SearchOption.AllDirectories, ListeBoxCsvPrésents.SelectedItem.ToString());
                }
                catch
                {
                    // ignored
                }
                try
                {
                    SuppressionFichiersIndividuels(LblCheminDossierXlsx.Text, ListeBoxXlsxPrésents, SearchOption.AllDirectories, ListeBoxXlsxPrésents.SelectedItem.ToString());
                }
                catch
                {
                    // ignored
                }
                try
                {
                    Directory.Delete(File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(1).Take(1).First(), true);
                }
                catch
                {
                    // ignored
                }
                try
                {
                    Directory.Delete(File.ReadLines("C:\\ELyco\\ELyco_out.txt").Skip(1).Take(1).First(), true);
                }
                catch (Exception)
                {
                    // ignored
                }
                try
                {
                    Directory.Delete(@"C:\ELyco", true);
                }
                catch (Exception)
                {
                    // ignored
                }

                LblCheminDossierCsv.Text = "";
                LblCheminDossierXlsx.Text = "";
                ComboAnnéeScolaire.Text = "";
                ComboNiveau6.Text = "";
                ComboNiveau5.Text = "";
                ComboNiveau4.Text = "";
                ComboNiveau3.Text = "";

                EffacerListbox(ListBoxCsvATraiter);
                EffacerListbox(ListeBoxCsvPrésents);
                EffacerListbox(ListeBoxXlsxPrésents);
                RemplirListeCsvPrésents();
                RemplirListeXlsxPrésents();
                ListBoxCsvATraiter.Refresh();
                ListeBoxCsvPrésents.Refresh();
                ListeBoxXlsxPrésents.Refresh();
                LblFichiersCsvATraiter.Text = "";
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void OuvrirFichierXlsx_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string fichierP1 = LblCheminDossierXlsx.Text + "1ère période\\" + ListeBoxXlsxPrésents.SelectedItem;
            string fichierP2 = LblCheminDossierXlsx.Text + "2ème période\\" + ListeBoxXlsxPrésents.SelectedItem;
            string fichierP3 = LblCheminDossierXlsx.Text + "3ème période\\" + ListeBoxXlsxPrésents.SelectedItem;
            string fichierAnnee = LblCheminDossierXlsx.Text + "Année\\" + ListeBoxXlsxPrésents.SelectedItem;
            string fichierDnb = LblCheminDossierXlsx.Text + "DNB\\" + ListeBoxXlsxPrésents.SelectedItem;
            if (File.Exists(fichierP1))
            {
                Process.Start(fichierP1);
            }
            if (File.Exists(fichierP2))
            {
                Process.Start(fichierP2);
            }
            if (File.Exists(fichierP3))
            {
                Process.Start(fichierP3);
            }
            if (File.Exists(fichierAnnee))
            {
                Process.Start(fichierAnnee);
            }
            if (File.Exists(fichierDnb))
            {
                Process.Start(fichierDnb);
            }
        }

        private void Création_arborescence(string niveau)
        {
            char classe = 'A';

            ComboBox combo = (ComboBox)Controls.Find(string.Format("ComboNiveau" + niveau), false).FirstOrDefault();
            if (combo != null)
                for (int i = 1; i <= int.Parse(combo.Items[combo.SelectedIndex].ToString()); i++)
                {
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "1ère période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "2ème période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "3ème période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "Année" + "\\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + niveau + classe);
                    classe++; // c1 is 'B' now
                }
        }

        private void Vérifier_chemins_dossiers()
        {
            if (File.Exists("C:\\ELyco\\ELyco_in.txt"))
            {
                using (TextReader tr = new StreamReader("C:\\ELyco\\ELyco_in.txt"))
                {
                    LblCheminDossierCsv.Text = tr.ReadLine() + @"\";
                }
            }
            if (File.Exists("C:\\ELyco\\ELyco_out.txt"))
            {
                using (TextReader tr1 = new StreamReader("C:\\ELyco\\ELyco_out.txt"))
                {
                    LblCheminDossierXlsx.Text = tr1.ReadLine() + @"\";
                }
            }
        }

        private static void Changer_ligne_fichier_txt(string newText, string fileName, int lineToEdit)
        {
            string[] arrLine = File.ReadAllLines(fileName);
            arrLine[lineToEdit - 1] = newText;
            File.WriteAllLines(fileName, arrLine);
        }

        private void Liste_fichiers_présents(string directoryPath, string periode, ListBox liste)

        {
            DirectoryInfo directoryInfo = new DirectoryInfo(directoryPath + periode);

            if (directoryInfo.Exists)

            {
                FileInfo[] fileInfo = directoryInfo.GetFiles();

                DirectoryInfo[] subdirectoryInfo = directoryInfo.GetDirectories();

                if (liste != ListBoxCsvATraiter)
                {
                    foreach (DirectoryInfo subDirectory in subdirectoryInfo)

                    {
                        Liste_fichiers_présents(subDirectory.FullName, "", liste);
                    }
                }

                foreach (FileInfo file in fileInfo)

                {
                    if (file.Length > 2000)
                        liste.Items.Add(file.Name);
                    else
                        file.Delete();
                    if (file.Name.Contains("Type"))
                    {
                        liste.Items.Remove(file.Name);
                    }
                }
            }
        }

        private int CompterFichiersXlsx(ListBox listbox)
        {
            int countXlsx = 0;
            foreach (var item in listbox.Items)
            {
                if (item.ToString().Contains("competence")) countXlsx++;
            }
            return countXlsx;
        }

        private int CompterFichiersDnb(ListBox listbox)
        {
            int countDocx = 0;
            foreach (var item in listbox.Items)
            {
                if (item.ToString().Contains("DNB-")) countDocx++;
            }
            return countDocx;
        }

        private void RemplirListeCsvPrésents()
        {
            ListeBoxCsvPrésents.Items.Add("1ère période");
            ListeBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "1ère période" + "\\", ListeBoxCsvPrésents);
            ListeBoxCsvPrésents.Items.Add("");
            ListeBoxCsvPrésents.Items.Add("2ème période");
            ListeBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "2ème période" + "\\", ListeBoxCsvPrésents);
            ListeBoxCsvPrésents.Items.Add("");
            ListeBoxCsvPrésents.Items.Add("3ème période");
            ListeBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "3ème période" + "\\", ListeBoxCsvPrésents);
            ListeBoxCsvPrésents.Items.Add("");
            ListeBoxCsvPrésents.Items.Add("Année");
            ListeBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "Année" + "\\", ListeBoxCsvPrésents);
        }

        private void RemplirListeXlsxPrésents()
        {
            ListeBoxXlsxPrésents.Items.Add("1ère période");
            ListeBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "1ère période" + "\\", ListeBoxXlsxPrésents);
            ListeBoxXlsxPrésents.Items.Add("");
            ListeBoxXlsxPrésents.Items.Add("2ème période");
            ListeBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "2ème période" + "\\", ListeBoxXlsxPrésents);
            ListeBoxXlsxPrésents.Items.Add("");
            ListeBoxXlsxPrésents.Items.Add("3ème période");
            ListeBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "3ème période" + "\\", ListeBoxXlsxPrésents);
            ListeBoxXlsxPrésents.Items.Add("");
            ListeBoxXlsxPrésents.Items.Add("Année");
            ListeBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "Année" + "\\", ListeBoxXlsxPrésents);
            ListeBoxXlsxPrésents.Items.Add("");
            ListeBoxXlsxPrésents.Items.Add("DNB");
            ListeBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "DNB" + "\\", ListeBoxXlsxPrésents);
        }

        private void RemplirListeCsvATraiter()
        {
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "", ListBoxCsvATraiter);
        }

        private void RafraichirListbox()
        {
            EffacerListbox(ListBoxCsvATraiter);
            EffacerListbox(ListeBoxCsvPrésents);
            EffacerListbox(ListeBoxXlsxPrésents);
            Vérifier_chemins_dossiers();
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            RemplirListeCsvATraiter();
            LblFichiersCsvATraiter.Text = ListBoxCsvATraiter.Items.Count + @" classes à traiter";
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListeBoxCsvPrésents) + @" fichiers CSV";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListeBoxXlsxPrésents) + @" fichiers XLSX et " + CompterFichiersDnb(ListeBoxXlsxPrésents) + @" fichiers DNB";
        }

        private void SelectionClasseTraitementAnnée(object sender, EventArgs e)
        {
            File.WriteAllText("C:\\ELyco\\ELyco_classes_annee.txt", String.Empty);
            foreach (var listBoxItem in ListeBoxCsvPrésents.SelectedItems)
            {
                if (listBoxItem.ToString().Contains("competence"))
                {
                    File.AppendAllText("C:\\ELyco\\ELyco_classes_annee.txt", Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
                }
            }
        }

        private void SelectionFichierDnb(object sender, EventArgs e)
        {
            File.WriteAllText("C:\\ELyco\\ELyco_classes_dnb.txt", String.Empty);
            foreach (var listBoxItem in ListeBoxXlsxPrésents.SelectedItems)
            {
                if (listBoxItem.ToString().Contains("DNB-"))
                {
                    File.AppendAllText("C:\\ELyco\\ELyco_classes_dnb.txt", Path.GetFileName(listBoxItem.ToString()).Substring(0, 17) + Environment.NewLine);
                }
            }
        }

        private void SuppressionFichiersIndividuels(string chemin, ListBox liste, SearchOption chercher, string fichier)
        {
            string[] files = Directory.GetFiles(chemin, "*.*", chercher);

            foreach (string file in files)
            {
                // ReSharper disable once UnusedVariable
                foreach (var selecteditem in liste.SelectedItems)
                {
                    if (file.Contains(fichier) && (file.Contains("competence") || file.Contains("DNB-")))
                        File.Delete(file);
                }
            }
            var selectedItems = liste.SelectedItems;

            if (liste.SelectedIndex != -1)
            {
                for (int i = selectedItems.Count - 1; i >= 0; i--)

                    if (selectedItems.Contains("competence") || selectedItems.Contains("DNB-"))
                    {
                        liste.Items.Remove(selectedItems[i]);
                    }
            }
            RafraichirListbox();
        }

        private void EffacerListbox(ListBox liste)
        {
            for (int i = liste.Items.Count - 1; i >= 0; i--)
            {
                liste.Items.RemoveAt(i);
            }
        }

        private void Supprimer_objets(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                // ignored
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Executer_macro(string macro)
        {
            //~~> Define your Excel Objects
            Application xlApp = new Application();

            Workbook xlWorkBook;

            string sPath = Path.GetTempFileName();
            File.WriteAllBytes(sPath, Resources.Compétences);

            //~~> Start Excel and open the workbook.
            xlWorkBook = xlApp.Workbooks.Open(sPath);

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run(macro);

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();

            //~~> Clean Up
            Supprimer_objets(xlApp);
            Supprimer_objets(xlWorkBook);
        }

        private void CopieFichierTypeDnb(Stream input, Stream output)
        {
            byte[] buffer = new byte[32768];
            while (true)
            {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        private void GénérerFichiersXlsxDnb()
        {
            string[] fichiers = Directory.GetFiles(LblCheminDossierXlsx.Text + "Année\\");

            foreach (var file in fichiers)
            {
                string classe = Path.GetFileNameWithoutExtension(file).Substring(17);
                string fichier = Path.GetFileName(file);

                string strPath = LblCheminDossierXlsx.Text + "DNB\\" + "Type_dnb.xlsx";
                if (File.Exists(strPath)) File.Delete(strPath);
                Assembly assembly = Assembly.GetExecutingAssembly();
                //In the next line you should provide      NameSpace.FileName.Extension that you have embedded
                var input = assembly.GetManifestResourceStream("Compétences.Resources.Type_dnb.xlsx");
                var output = File.Open(strPath, FileMode.CreateNew);
                CopieFichierTypeDnb(input, output);
                input?.Dispose();
                output.Dispose();

                string strPath1 = LblCheminDossierXlsx.Text + "DNB\\" + "Type_dnb.docx";
                if (File.Exists(strPath1)) File.Delete(strPath1);
                Assembly assembly1 = Assembly.GetExecutingAssembly();
                var input1 = assembly1.GetManifestResourceStream("Compétences.Resources.Type_dnb.docx");
                var output1 = File.Open(strPath1, FileMode.CreateNew);
                CopieFichierTypeDnb(input1, output1);
                input1?.Dispose();
                output1.Dispose();

                Application excelApplication = new Application();

                string srcPath = LblCheminDossierXlsx.Text + "Année\\" + fichier;
                Workbook srcworkBook = excelApplication.Workbooks.Open(srcPath);
                Worksheet srcworkSheet = (Worksheet)srcworkBook.Sheets.Item[1];

                string destPath = strPath;
                Workbook destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
                Worksheet destworkSheet = (Worksheet)destworkBook.Sheets.Item[1];
                Worksheet destworkSheet2 = (Worksheet)destworkBook.Sheets.Item[2];

                Range range = srcworkSheet.Range["A2:A50"];
                int cnt = -3;

                foreach (Range element in range.Cells)

                {
                    if (element.Value2 != null)
                    {
                        cnt = cnt + 1;
                    }
                }

                Range from = srcworkSheet.Range["B1:J" + (cnt + 1)];
                Range to = destworkSheet.Range["AI1"];

                Range from1 = srcworkSheet.Range["A2:A" + (cnt + 1)];
                Range to1 = destworkSheet.Range["A2"];

                Range from2 = srcworkSheet.Range["A2:A" + (cnt + 1)];
                Range to2 = destworkSheet2.Range["A2"];

                from.Copy(to);
                from1.Copy(to1);
                from2.Copy(to2);

                Range cells1 = destworkSheet.Range["B2:B" + (cnt + 1)];
                cells1.Value = classe;

                Range cells = destworkSheet.Range["A" + (cnt + 2) + ":AG50"];

                Range del = cells.EntireRow;

                del.Delete();

                destworkBook.SaveAs(LblCheminDossierXlsx.Text + "DNB\\DNB-" + classe + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                srcworkBook.Close();
                destworkBook.Close();
            }
            RafraichirListbox();
        }

        private void Drag(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string file in fileList)
            {
                string filename = Path.GetFullPath(file);
                ListBoxCsvATraiter.Items.Add(filename);
            }

            File.WriteAllText("C:\\ELyco\\ELyco_classes.txt", String.Empty);

            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(), LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText("C:\\ELyco\\ELyco_classes.txt", Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
            }

            RafraichirListbox();
        }

        private void Drag_Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void VérifierDoublonClasseCsv(string période)
        {
            ListBox templist = new ListBox();
            foreach (var t in ListBoxCsvATraiter.Items)

            {
                string classe = Path.GetFileName(t.ToString()).Substring(25, 2);

                foreach (string s in Directory.GetFiles(LblCheminDossierCsv.Text + "\\" + période + "\\" + classe + "\\", "*.csv").Select(Path.GetFileName))
                {
                    string classe1 = s.Substring(25, 2);

                    if (classe == classe1)
                    {
                        MessageBox.Show(@"Il existe déjà un fichier pour la classe " + classe1 + @" dans le dossier '" + période + @"'");
                        templist.Items.Add(classe);
                        File.Delete(LblCheminDossierCsv.Text + "\\" + s);
                    }
                }
            }

            foreach (var v in templist.Items)
            {
                ListBoxCsvATraiter.Items.Remove(v);
            }

            RafraichirListbox();
        }

        public void NettoyageFichiersCsvATraiter()
        {
            foreach (string s in Directory.GetFiles(LblCheminDossierCsv.Text + "\\", "*.csv").Select(Path.GetFileName))
            {
                File.Delete(LblCheminDossierCsv.Text + "\\" + s);
            }
        }

        private string DétectionPériode()
        {
            string périodeSelect = null;
            foreach (RadioButton période in PanelTrimestre.Controls)
            {
                if (période.Checked)
                {
                    périodeSelect = période.Text;
                }
            }
            return périodeSelect;
        }

        private void BtnSuppressionFichierCsvATraiter_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvATraiter, SearchOption.TopDirectoryOnly, ListBoxCsvATraiter.SelectedItem.ToString().Substring(ListBoxCsvATraiter.SelectedItem.ToString().LastIndexOf('\\') + 1));
        }
    }
}