using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Compétences.Properties;
using Microsoft.Office.Interop.Excel;
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
            File.WriteAllText("C:\\ELyco\\ELyco_classes.txt", string.Empty);
            File.WriteAllText("C:\\ELyco\\ELyco_classes_annee.txt", string.Empty);
            File.WriteAllText("C:\\ELyco\\ELyco_classes_dnb.txt", string.Empty);
            BtnLancerTraitement.Enabled = false;
            RafraichirListbox();
            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(),
                        LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText("C:\\ELyco\\ELyco_classes.txt",
                    Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
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
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var folder = dlg.SelectedPath + "\\ELyco_CSV\\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierCsv.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_in.txt"))
                    using (File.Create("C:\\ELyco\\ELyco_in.txt"))
                    {
                    }
                if (!File.Exists("C:\\ELyco\\ELyco_classes.txt"))
                    using (File.Create("C:\\ELyco\\ELyco_classes.txt"))
                    {
                    }
                if (!File.Exists("C:\\ELyco\\ELyco_classes_annee.txt"))
                    using (File.Create("C:\\ELyco\\ELyco_classes_annee.txt"))
                    {
                    }
                if (!File.Exists("C:\\ELyco\\ELyco_classes_dnb.txt"))
                    using (File.Create("C:\\ELyco\\ELyco_classes_dnb.txt"))
                    {
                    }
                using (var sw = new StreamWriter("C:\\ELyco\\ELyco_in.txt"))
                {
                    sw.WriteLine(LblCheminDossierCsv.Text);
                    sw.WriteLine(dlg.SelectedPath + "\\ELyco_CSV" + "\n");
                }
            }
        }

        private void Dossier_destination_Click(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var folder = dlg.SelectedPath + "\\ELyco_Competences\\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierXlsx.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_out.txt"))
                    using (File.Create("C:\\ELyco\\ELyco_out.txt"))
                    {
                    }
                using (var sw = new StreamWriter("C:\\ELyco\\ELyco_out.txt"))
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

            var traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += TraitementMacroLancement;
            traitementMacro.RunWorkerCompleted += TraitementMacroFini;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            if (LblFichiersCsvATraiter.Text == "")
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text =
                    @"Traitement des fichiers...Veuillez patienter...";
            else
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text =
                    LblFichiersCsvATraiter.Text + @"...Veuillez patienter...";

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

            var files = Directory.GetFiles(LblCheminDossierCsv.Text, "*.*");
            foreach (var file in files)
                File.Delete(file);

            Frm2.Controls.Find("LblMessageTraitement", true).First().Text = @"Traitement des fichiers terminé !";
            Frm2.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void LancerTraitementDnb_Click(object sender, EventArgs e)
        {
            var traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += TraitementMacroDnbLancement;
            traitementMacro.RunWorkerCompleted += TraitementMacroDnbFini;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            if (LblFichiersCsvATraiter.Text == "")
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text =
                    @"Traitement des fichiers...Veuillez patienter...";
            else
                Frm2.Controls.Find("LblMessageTraitement", true).First().Text =
                    LblFichiersCsvATraiter.Text + @"...Veuillez patienter...";

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

        private void BtnSuppressionFichierCsvATraiter_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvATraiter, SearchOption.TopDirectoryOnly);
        }

        private void SuppressionFichierCsv_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvPrésents, SearchOption.AllDirectories);
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void SuppressionFichierXlsx_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierXlsx.Text, ListBoxXlsxPrésents, SearchOption.AllDirectories);
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void SuppressionBases_Click(object sender, EventArgs e)
        {
            var dialogResult = MessageBox.Show(@"Etes-vous sûr de vouloir tout supprimer ?", @"Attention !",
                MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvPrésents,
                        SearchOption.AllDirectories);
                }
                catch
                {
                    // ignored
                }
                try
                {
                    SuppressionFichiersIndividuels(LblCheminDossierXlsx.Text, ListBoxXlsxPrésents,
                        SearchOption.AllDirectories);
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
                EffacerListbox(ListBoxCsvPrésents);
                EffacerListbox(ListBoxXlsxPrésents);
                RemplirListeCsvPrésents();
                RemplirListeXlsxPrésents();
                ListBoxCsvATraiter.Refresh();
                ListBoxCsvPrésents.Refresh();
                ListBoxXlsxPrésents.Refresh();
                LblFichiersCsvATraiter.Text = "";
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void SuppressionFichiersIndividuels(string chemin, ListBox liste, SearchOption chercher)
        {
            var files = Directory.GetFiles(chemin, "*.*", chercher);

            foreach (var file in files)
            foreach (var item in liste.SelectedItems)
                if (file.Contains(item.ToString()) && (file.Contains("competence") || file.Contains("DNB-")))
                    File.Delete(file);
            var selectedItems = liste.SelectedItems;

            if (liste.SelectedIndex != -1)
                for (var i = selectedItems.Count - 1; i >= 0; i--)

                    if (selectedItems.Contains("competence") || selectedItems.Contains("DNB-"))
                        liste.Items.Remove(selectedItems[i]);
            RafraichirListbox();
        }

        public void NettoyageFichiersCsvATraiter()
        {
            foreach (var s in Directory.GetFiles(LblCheminDossierCsv.Text + "\\", "*.csv").Select(Path.GetFileName))
                File.Delete(LblCheminDossierCsv.Text + "\\" + s);
        }

        private void EffacerListbox(ListBox liste)
        {
            for (var i = liste.Items.Count - 1; i >= 0; i--)
                liste.Items.RemoveAt(i);
        }

        private void OuvrirFichierXlsx_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            var fichierP1 = LblCheminDossierXlsx.Text + "1ère période\\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierP2 = LblCheminDossierXlsx.Text + "2ème période\\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierP3 = LblCheminDossierXlsx.Text + "3ème période\\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierAnnee = LblCheminDossierXlsx.Text + "Année\\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierDnb = LblCheminDossierXlsx.Text + "DNB\\" + ListBoxXlsxPrésents.SelectedItem;
            if (File.Exists(fichierP1))
                Process.Start(fichierP1);
            if (File.Exists(fichierP2))
                Process.Start(fichierP2);
            if (File.Exists(fichierP3))
                Process.Start(fichierP3);
            if (File.Exists(fichierAnnee))
                Process.Start(fichierAnnee);
            if (File.Exists(fichierDnb))
                Process.Start(fichierDnb);
        }

        private void Création_arborescence(string niveau)
        {
            var classe = 'A';

            var combo = (ComboBox) Controls.Find(string.Format("ComboNiveau" + niveau), false).FirstOrDefault();
            if (combo != null)
                for (var i = 1; i <= int.Parse(combo.Items[combo.SelectedIndex].ToString()); i++)
                {
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "1ère période" + "\\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "2ème période" + "\\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "3ème période" + "\\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + "Année" + "\\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + "\\" + niveau + classe);
                    classe++; // c1 is 'B' now
                }
        }

        private void Vérifier_chemins_dossiers()
        {
            if (File.Exists("C:\\ELyco\\ELyco_in.txt"))
                using (TextReader tr = new StreamReader("C:\\ELyco\\ELyco_in.txt"))
                {
                    LblCheminDossierCsv.Text = tr.ReadLine() + @"\";
                }
            if (File.Exists("C:\\ELyco\\ELyco_out.txt"))
                using (TextReader tr1 = new StreamReader("C:\\ELyco\\ELyco_out.txt"))
                {
                    LblCheminDossierXlsx.Text = tr1.ReadLine() + @"\";
                }
        }

        private void VérifierDoublonClasseCsv(string période)
        {
            var templist = new ListBox();
            foreach (var t in ListBoxCsvATraiter.Items)

            {
                var classe = Path.GetFileName(t.ToString()).Substring(25, 2);

                foreach (var s in Directory
                    .GetFiles(LblCheminDossierCsv.Text + "\\" + période + "\\" + classe + "\\", "*.csv")
                    .Select(Path.GetFileName))
                {
                    var classe1 = s.Substring(25, 2);

                    if (classe == classe1)
                    {
                        MessageBox.Show(@"Il existe déjà un fichier pour la classe " + classe1 + @" dans le dossier '" +
                                        période + @"'");
                        templist.Items.Add(classe);
                        File.Delete(LblCheminDossierCsv.Text + "\\" + s);
                    }
                }
            }

            foreach (var v in templist.Items)
                ListBoxCsvATraiter.Items.Remove(v);

            RafraichirListbox();
        }

        private static void Changer_ligne_fichier_txt(string newText, string fileName, int lineToEdit)
        {
            var arrLine = File.ReadAllLines(fileName);
            arrLine[lineToEdit - 1] = newText;
            File.WriteAllLines(fileName, arrLine);
        }

        private void Liste_fichiers_présents(string directoryPath, string periode, ListBox liste)

        {
            var directoryInfo = new DirectoryInfo(directoryPath + periode);

            if (directoryInfo.Exists)

            {
                var fileInfo = directoryInfo.GetFiles();

                var subdirectoryInfo = directoryInfo.GetDirectories();

                if (liste != ListBoxCsvATraiter)
                    foreach (var subDirectory in subdirectoryInfo)

                        Liste_fichiers_présents(subDirectory.FullName, "", liste);

                foreach (var file in fileInfo)

                {
                    if (file.Length > 2000)
                        liste.Items.Add(file.Name);
                    else
                        file.Delete();
                    if (file.Name.Contains("Type"))
                        liste.Items.Remove(file.Name);
                }
            }
        }

        private int CompterFichiersXlsx(ListBox listbox)
        {
            var countXlsx = 0;
            foreach (var item in listbox.Items)
                if (item.ToString().Contains("competence")) countXlsx++;
            return countXlsx;
        }

        private int CompterFichiersDnb(ListBox listbox)
        {
            var countDocx = 0;
            foreach (var item in listbox.Items)
                if (item.ToString().Contains("DNB-")) countDocx++;
            return countDocx;
        }

        private void RemplirListeCsvPrésents()
        {
            ListBoxCsvPrésents.Items.Add("1ère période");
            ListBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "1ère période" + "\\", ListBoxCsvPrésents);
            ListBoxCsvPrésents.Items.Add("");
            ListBoxCsvPrésents.Items.Add("2ème période");
            ListBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "2ème période" + "\\", ListBoxCsvPrésents);
            ListBoxCsvPrésents.Items.Add("");
            ListBoxCsvPrésents.Items.Add("3ème période");
            ListBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "3ème période" + "\\", ListBoxCsvPrésents);
            ListBoxCsvPrésents.Items.Add("");
            ListBoxCsvPrésents.Items.Add("Année");
            ListBoxCsvPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "Année" + "\\", ListBoxCsvPrésents);
        }

        private void RemplirListeXlsxPrésents()
        {
            ListBoxXlsxPrésents.Items.Add("1ère période");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "1ère période" + "\\", ListBoxXlsxPrésents);
            ListBoxXlsxPrésents.Items.Add("");
            ListBoxXlsxPrésents.Items.Add("2ème période");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "2ème période" + "\\", ListBoxXlsxPrésents);
            ListBoxXlsxPrésents.Items.Add("");
            ListBoxXlsxPrésents.Items.Add("3ème période");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "3ème période" + "\\", ListBoxXlsxPrésents);
            ListBoxXlsxPrésents.Items.Add("");
            ListBoxXlsxPrésents.Items.Add("Année");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "Année" + "\\", ListBoxXlsxPrésents);
            ListBoxXlsxPrésents.Items.Add("");
            ListBoxXlsxPrésents.Items.Add("DNB");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(LblCheminDossierXlsx.Text, "DNB" + "\\", ListBoxXlsxPrésents);
        }

        private void RemplirListeCsvATraiter()
        {
            Liste_fichiers_présents(LblCheminDossierCsv.Text, "", ListBoxCsvATraiter);
        }

        private void RafraichirListbox()
        {
            EffacerListbox(ListBoxCsvATraiter);
            EffacerListbox(ListBoxCsvPrésents);
            EffacerListbox(ListBoxXlsxPrésents);
            Vérifier_chemins_dossiers();
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            RemplirListeCsvATraiter();
            LblFichiersCsvATraiter.Text = ListBoxCsvATraiter.Items.Count + @" classes à traiter";
            LblFichiersCsvPrésents.Text = CompterFichiersXlsx(ListBoxCsvPrésents) + @" fichiers CSV";
            LblFichiersXlsxPrésents.Text = CompterFichiersXlsx(ListBoxXlsxPrésents) + @" fichiers XLSX et " +
                                           CompterFichiersDnb(ListBoxXlsxPrésents) + @" fichiers DNB";
        }

        private void SelectionClasseTraitementAnnée(object sender, EventArgs e)
        {
            File.WriteAllText("C:\\ELyco\\ELyco_classes_annee.txt", string.Empty);
            foreach (var listBoxItem in ListBoxCsvPrésents.SelectedItems)
                if (listBoxItem.ToString().Contains("competence"))
                    File.AppendAllText("C:\\ELyco\\ELyco_classes_annee.txt",
                        Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
        }

        private void SelectionFichierDnb(object sender, EventArgs e)
        {
            File.WriteAllText("C:\\ELyco\\ELyco_classes_dnb.txt", string.Empty);
            foreach (var listBoxItem in ListBoxXlsxPrésents.SelectedItems)
                if (listBoxItem.ToString().Contains("DNB-"))
                    File.AppendAllText("C:\\ELyco\\ELyco_classes_dnb.txt",
                        Path.GetFileName(listBoxItem.ToString()).Substring(0, 17) + Environment.NewLine);
        }

        private string DétectionPériode()
        {
            string périodeSelect = null;
            foreach (RadioButton période in PanelTrimestre.Controls)
                if (période.Checked)
                    périodeSelect = période.Text;
            return périodeSelect;
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
            var xlApp = new Application();

            Workbook xlWorkBook;

            var sPath = Path.GetTempFileName();
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
            var buffer = new byte[32768];
            while (true)
            {
                var read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        private void GénérerFichiersXlsxDnb()
        {
            var fichiers = Directory.GetFiles(LblCheminDossierXlsx.Text + "Année\\");

            foreach (var file in fichiers)
            {
                var classe = Path.GetFileNameWithoutExtension(file).Substring(17);
                var fichier = Path.GetFileName(file);

                var strPath = LblCheminDossierXlsx.Text + "DNB\\" + "Type_dnb.xlsx";
                if (File.Exists(strPath)) File.Delete(strPath);
                var assembly = Assembly.GetExecutingAssembly();
                //In the next line you should provide      NameSpace.FileName.Extension that you have embedded
                var input = assembly.GetManifestResourceStream("Compétences.Resources.Type_dnb.xlsx");
                var output = File.Open(strPath, FileMode.CreateNew);
                CopieFichierTypeDnb(input, output);
                input?.Dispose();
                output.Dispose();

                var strPath1 = LblCheminDossierXlsx.Text + "DNB\\" + "Type_dnb.docx";
                if (File.Exists(strPath1)) File.Delete(strPath1);
                var assembly1 = Assembly.GetExecutingAssembly();
                var input1 = assembly1.GetManifestResourceStream("Compétences.Resources.Type_dnb.docx");
                var output1 = File.Open(strPath1, FileMode.CreateNew);
                CopieFichierTypeDnb(input1, output1);
                input1?.Dispose();
                output1.Dispose();

                var excelApplication = new Application();

                var srcPath = LblCheminDossierXlsx.Text + "Année\\" + fichier;
                var srcworkBook = excelApplication.Workbooks.Open(srcPath);
                var srcworkSheet = (Worksheet) srcworkBook.Sheets.Item[1];

                var destPath = strPath;
                var destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
                var destworkSheet = (Worksheet) destworkBook.Sheets.Item[1];
                var destworkSheet2 = (Worksheet) destworkBook.Sheets.Item[2];

                var range = srcworkSheet.Range["A2:A50"];
                var cnt = -3;

                foreach (Range element in range.Cells)

                    if (element.Value2 != null)
                        cnt = cnt + 1;

                var from = srcworkSheet.Range["B1:J" + (cnt + 1)];
                var to = destworkSheet.Range["AI1"];

                var from1 = srcworkSheet.Range["A2:A" + (cnt + 1)];
                var to1 = destworkSheet.Range["A2"];

                var from2 = srcworkSheet.Range["A2:A" + (cnt + 1)];
                var to2 = destworkSheet2.Range["A2"];

                from.Copy(to);
                from1.Copy(to1);
                from2.Copy(to2);

                var cells1 = destworkSheet.Range["B2:B" + (cnt + 1)];
                cells1.Value = classe;

                var cells = destworkSheet.Range["A" + (cnt + 2) + ":AG50"];

                var del = cells.EntireRow;

                del.Delete();

                destworkBook.SaveAs(LblCheminDossierXlsx.Text + "DNB\\DNB-" + classe + "_" +
                                    DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                srcworkBook.Close();
                destworkBook.Close();
            }
            RafraichirListbox();
        }

        private void Drag(object sender, DragEventArgs e)
        {
            var fileList = (string[]) e.Data.GetData(DataFormats.FileDrop, false);
            foreach (var file in fileList)
            {
                var filename = Path.GetFullPath(file);
                ListBoxCsvATraiter.Items.Add(filename);
            }

            File.WriteAllText("C:\\ELyco\\ELyco_classes.txt", string.Empty);

            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(),
                        LblCheminDossierCsv.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText("C:\\ELyco\\ELyco_classes.txt",
                    Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
            }

            RafraichirListbox();
        }

        private void Drag_Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
                e.Effect = DragDropEffects.All;
        }

        private void SupprimerSélectionsListbox(object sender, MouseEventArgs e)
        {
            ListBoxCsvATraiter.ClearSelected();
            ListBoxCsvPrésents.ClearSelected();
            ListBoxXlsxPrésents.ClearSelected();
        }
    }
}