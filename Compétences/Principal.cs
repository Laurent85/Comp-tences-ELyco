using Compétences.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
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
        public string CheminElyco = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        public Message MessageTraitement = new Message();

        public Principal()
        {
            InitializeComponent();
        }

        private void OuvertureLogiciel(object sender, EventArgs e)
        {
            Directory.CreateDirectory(CheminElyco + @"\ELyco\Config");
            Directory.CreateDirectory(CheminElyco + @"\ELyco");
            Directory.CreateDirectory(CheminElyco + @"\ELyco\Config");
            Directory.CreateDirectory(CheminElyco + @"\ELyco\Backup");
            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes.txt", string.Empty);
            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_annee.txt", string.Empty);
            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_dnb.txt", string.Empty);
            BtnLancerTraitement.Enabled = false;
            BtnSuppressionFichierCsvATraiter.Enabled = false;
            BtnSuppressionFichierCsv.Enabled = false;
            BtnSuppressionFichierXlsx.Enabled = false;
            BtnGénérerfichiersExcelDnb.Enabled = false;
            BtnGénérerPublipostageDnb.Enabled = false;
            RafraichirListbox();

            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + @"\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(),
                        LblCheminDossierCsv.Text + @"\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText(CheminElyco + @"\ELyco\Config\ELyco_classes.txt",
                    Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
            }

            try
            {
                ComboAnnéeScolaire.Text =
                    File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(2).Take(3).First();
                ComboNiveau6.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(3).Take(4).First();
                ComboNiveau5.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(4).Take(5).First();
                ComboNiveau4.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(5).Take(6).First();
                ComboNiveau3.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(6).Take(7).First();
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void BtnCheminCsv(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var chemin = dlg.SelectedPath + @"\ELyco_CSV\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierCsv.Text = chemin;
                Directory.CreateDirectory(CheminElyco + @"\ELyco\Config");

                if (!File.Exists(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                    using (File.Create(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                    {
                    }
                if (!File.Exists(CheminElyco + @"\ELyco\Config\ELyco_classes.txt"))
                    using (File.Create(CheminElyco + @"\ELyco\Config\ELyco_classes.txt"))
                    {
                    }
                if (!File.Exists(CheminElyco + @"\ELyco\Config\ELyco_classes_annee.txt"))
                    using (File.Create(CheminElyco + @"\ELyco\Config\ELyco_classes_annee.txt"))
                    {
                    }
                if (!File.Exists(CheminElyco + @"\ELyco\Config\ELyco_classes_dnb.txt"))
                    using (File.Create(CheminElyco + @"\ELyco\Config\ELyco_classes_dnb.txt"))
                    {
                    }
                using (var sw = new StreamWriter(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                {
                    sw.WriteLine(LblCheminDossierCsv.Text);
                    sw.WriteLine(dlg.SelectedPath + @"\ELyco_CSV" + "\n");
                }
            }
        }

        private void BtnCheminXlsx(object sender, EventArgs e)
        {
            var dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var cheminXlsx = dlg.SelectedPath + @"\ELyco_Competences\" + ComboAnnéeScolaire.SelectedItem;
                LblCheminDossierXlsx.Text = cheminXlsx;
                Directory.CreateDirectory(CheminElyco + @"\ELyco\Config");

                if (!File.Exists(CheminElyco + @"\ELyco\Config\ELyco_out.txt"))
                    using (File.Create(CheminElyco + @"\ELyco\Config\ELyco_out.txt"))
                    {
                    }
                using (var sw = new StreamWriter(CheminElyco + @"\ELyco\Config\ELyco_out.txt"))
                {
                    sw.WriteLine(LblCheminDossierXlsx.Text);
                    sw.WriteLine(dlg.SelectedPath + @"\ELyco_Competences");
                }
            }
        }

        private void BtnCréerArborescence(object sender, EventArgs e)
        {
            CréationArborescence("6");
            CréationArborescence("5");
            CréationArborescence("4");
            CréationArborescence("3");

            Directory.CreateDirectory(LblCheminDossierXlsx.Text + @"\" + "1ère période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + @"\" + "2ème période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + @"\" + "3ème période");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + @"\" + "Année");
            Directory.CreateDirectory(LblCheminDossierXlsx.Text + @"\" + "DNB");

            ChangerLigneFichierTxt(ComboAnnéeScolaire.SelectedItem + "\n", CheminElyco + @"\ELyco\Config\ELyco_in.txt",
                3);
            ChangerLigneFichierTxt(ComboNiveau6.SelectedItem + "\n", CheminElyco + @"\ELyco\Config\ELyco_in.txt", 4);
            ChangerLigneFichierTxt(ComboNiveau5.SelectedItem + "\n", CheminElyco + @"\ELyco\Config\ELyco_in.txt", 5);
            ChangerLigneFichierTxt(ComboNiveau4.SelectedItem + "\n", CheminElyco + @"\ELyco\Config\ELyco_in.txt", 6);
            ChangerLigneFichierTxt(ComboNiveau3.SelectedItem + "\n", CheminElyco + @"\ELyco\Config\ELyco_in.txt", 7);
        }

        private void BtnTraitementCsv(object sender, EventArgs e)
        {
            VérifierDoublonClasseCsv(DétectionPériode());

            var traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += DébutMacroCompétences;
            traitementMacro.RunWorkerCompleted += FinMacroCompétences;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            if (LblFichiersCsvATraiter.Text == @"0 classes à traiter")

                MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                    ListBoxCsvPrésents.SelectedItems.Count +
                    @" classes à traiter...Veuillez patienter...";
            else
                MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                    LblFichiersCsvATraiter.Text + @"...Veuillez patienter...";

            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = false;
            MessageTraitement.ShowDialog();
        }

        private void DébutMacroCompétences(object sender, DoWorkEventArgs e)
        {
            if (RadioBtnPériode1.Checked)
            {
                ExécuterMacro("Deplacer_P1.Deplacer_P1");
                ExécuterMacro("Compétences_par_lot_P1.Compétences_par_lot_P1");
            }
            if (RadioBtnPériode2.Checked)
            {
                ExécuterMacro("Deplacer_P2.Deplacer_P2");
                ExécuterMacro("Compétences_par_lot_P2.Compétences_par_lot_P2");
            }
            if (RadioBtnPériode3.Checked)
            {
                ExécuterMacro("Deplacer_P3.Deplacer_P3");
                ExécuterMacro("Compétences_par_lot_P3.Compétences_par_lot_P3");
            }
            if (RadioBtnAnnée.Checked)
            {
                ExécuterMacro("Fusionner.Fusionner");
                ExécuterMacro("Compétences_par_lot_Année.Compétences_par_lot_Année");
            }
        }

        private void FinMacroCompétences(object sender, RunWorkerCompletedEventArgs e)
        {
            RafraichirListbox();

            var files = Directory.GetFiles(LblCheminDossierCsv.Text, "*.*");
            foreach (var file in files)
                File.Delete(file);

            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes.txt", string.Empty);

            MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                @"Traitement des fichiers terminé !";
            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void BtnTraitementDnb(object sender, EventArgs e)
        {
            var traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += DébutMacroDnb;
            traitementMacro.RunWorkerCompleted += FinMacroDnb;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                ListBoxXlsxPrésents.SelectedItems.Count +
                @" classes à traiter...Veuillez patienter...";

            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = false;
            MessageTraitement.ShowDialog();
        }

        private void DébutMacroDnb(object sender, DoWorkEventArgs e)
        {
            ExécuterMacro("Publipostage.Publipostage");

            foreach (var fichierSélectionné in ListBoxXlsxPrésents.SelectedItems)
            {
                if (fichierSélectionné.ToString().Contains("xlsx"))
                {
                    Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
                    var strPath = LblCheminDossierXlsx.Text + @"DNB\" + fichierSélectionné;
                    var nomFichier = Path.GetFileNameWithoutExtension(strPath);
                    Document wordDocument = appWord.Documents.Open(LblCheminDossierXlsx.Text + @"DNB\" + nomFichier + @".docx");
                    wordDocument.ExportAsFixedFormat(LblCheminDossierXlsx.Text + @"DNB\" + nomFichier + @".pdf", WdExportFormat.wdExportFormatPDF, false);
                    appWord.Documents.Close();
                    appWord.Quit();
                    GC.Collect();
                }
            }
        }

        private void FinMacroDnb(object sender, RunWorkerCompletedEventArgs e)
        {
            RafraichirListbox();
            MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                @"Traitement des fichiers terminé !";
            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void BtnGénérerfichiersExcelDnb_Click(object sender, EventArgs e)
        {
            var traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += DébutGénérerfichiersExcelDnb;
            traitementMacro.RunWorkerCompleted += FinGénérerfichiersExcelDnb;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;

            MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                ListBoxXlsxPrésents.SelectedItems.Count +
                @" classes à traiter...Veuillez patienter...";

            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = false;
            MessageTraitement.ShowDialog();
        }

        private void DébutGénérerfichiersExcelDnb(object sender, DoWorkEventArgs e)
        {
            GénérerFichiersXlsxDnb();
        }

        private void FinGénérerfichiersExcelDnb(object sender, RunWorkerCompletedEventArgs e)
        {
            EffacerListbox(ListBoxXlsxPrésents);
            VérifierCheminsDossiers();
            RemplirListeXlsxPrésents();
            LblFichiersXlsxPrésents.Text = CompterFichiersPrésents(ListBoxXlsxPrésents) + @" fichiers XLSX et " +
                                           CompterFichiersDnb(ListBoxXlsxPrésents) + @" fichiers DNB";
            MessageTraitement.Controls.Find("LblMessageTraitement", true).First().Text =
                @"Traitement des fichiers terminé !";
            MessageTraitement.Controls.Find("BtnFermerMessageTraitement", true).First().Visible = true;
        }

        private void BtnSuppressionFichierCsvAtraiter(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvATraiter, SearchOption.TopDirectoryOnly);
        }

        private void BtnSuppressionFichierCsv_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierCsv.Text, ListBoxCsvPrésents, SearchOption.AllDirectories);
            LblFichiersCsvPrésents.Text = CompterFichiersPrésents(ListBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersPrésents(ListBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void BtnSuppressionFichierXlsx_Click(object sender, EventArgs e)
        {
            SuppressionFichiersIndividuels(LblCheminDossierXlsx.Text, ListBoxXlsxPrésents, SearchOption.AllDirectories);
            LblFichiersCsvPrésents.Text = CompterFichiersPrésents(ListBoxCsvPrésents) + @" fichiers CSV présents";
            LblFichiersXlsxPrésents.Text = CompterFichiersPrésents(ListBoxXlsxPrésents) + @" fichiers XLSX présents";
        }

        private void BtnSuppressionBases_Click(object sender, EventArgs e)
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
                    Directory.Delete(
                        File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(1).Take(1).First(), true);
                }
                catch
                {
                    // ignored
                }
                try
                {
                    Directory.Delete(
                        File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_out.txt").Skip(1).Take(1).First(), true);
                }
                catch (Exception)
                {
                    // ignored
                }
                try
                {
                    Directory.Delete(CheminElyco + @"\ELyco\Config", true);
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

        private void BtnSauvegarderBases_Click(object sender, EventArgs e)
        {
            var date = DateTime.Now.ToString("dd-MM-yyyy_HH'h'mm");
            var dossierDest = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\ELyco\Backup\";
            var cheminCsv = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(1).Take(1).First();
            var cheminCompétences =
                File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_out.txt").Skip(1).Take(1).First();

            if (!Directory.Exists(dossierDest + date))
            {
                Directory.CreateDirectory(dossierDest + date);
                ZipFile.CreateFromDirectory(CheminElyco + @"\ELyco\Config", dossierDest + date + @"\ELyco.zip");
                ZipFile.CreateFromDirectory(cheminCsv, dossierDest + date + @"\ELyco_CSV.zip");
                ZipFile.CreateFromDirectory(cheminCompétences, dossierDest + date + @"\ELyco_Competences.zip");
            }

            MessageBox.Show(@"Sauvegarde effectuée avec succès vers " + dossierDest + date);
        }

        private void BtnRestaurerBases_Click(object sender, EventArgs e)
        {
            if (File.Exists(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                SuppressionFichiersSauvegarde(File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(1)
                    .Take(1)
                    .First());
            if (File.Exists(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                SuppressionFichiersSauvegarde(File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_out.txt").Skip(1)
                    .Take(1)
                    .First());
            SuppressionFichiersSauvegarde(CheminElyco + @"\ELyco\Config");

            var dlg = new FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.ApplicationData,
                SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\ELyco\Backup\"
            };
            SendKeys.Send("{TAB}{TAB}{RIGHT}");

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var sauvegardeSélectionnée = dlg.SelectedPath;
                ZipFile.ExtractToDirectory(sauvegardeSélectionnée + @"\ELyco.zip", CheminElyco + @"\ELyco\Config");
                ZipFile.ExtractToDirectory(sauvegardeSélectionnée + @"\ELyco_CSV.zip",
                    File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(1).Take(1).First());
                ZipFile.ExtractToDirectory(sauvegardeSélectionnée + @"\ELyco_Competences.zip",
                    File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_out.txt").Skip(1).Take(1).First());
            }
            try
            {
                ComboAnnéeScolaire.Text =
                    File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(2).Take(3).First();
                ComboNiveau6.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(3).Take(4).First();
                ComboNiveau5.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(4).Take(5).First();
                ComboNiveau4.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(5).Take(6).First();
                ComboNiveau3.Text = File.ReadLines(CheminElyco + @"\ELyco\Config\ELyco_in.txt").Skip(6).Take(7).First();
            }
            catch (Exception)
            {
                // ignored
            }
            RafraichirListbox();
            MessageBox.Show(@"Restauration effectuée avec succès depuis " + dlg.SelectedPath);
        }

        private void SuppressionSélectionsListbox(object sender, EventArgs e)
        {
            ListBoxCsvATraiter.ClearSelected();
            ListBoxCsvPrésents.ClearSelected();
            ListBoxXlsxPrésents.ClearSelected();
        }

        private void OuvrirFichierXlsxDocx(object sender, MouseEventArgs e)
        {
            var fichierP1 = LblCheminDossierXlsx.Text + @"1ère période\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierP2 = LblCheminDossierXlsx.Text + @"2ème période\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierP3 = LblCheminDossierXlsx.Text + @"3ème période\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierAnnee = LblCheminDossierXlsx.Text + @"Année\" + ListBoxXlsxPrésents.SelectedItem;
            var fichierDnb = LblCheminDossierXlsx.Text + @"DNB\" + ListBoxXlsxPrésents.SelectedItem;
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

        private void SélectionFichierCsvATraiter(object sender, EventArgs e)
        {
            BtnSuppressionFichierCsvATraiter.Enabled = ListBoxCsvATraiter.SelectedItems.Count != 0;
        }

        private void SélectionFichierCsvPrésent(object sender, EventArgs e)
        {
            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_annee.txt", string.Empty);
            foreach (var listBoxItem in ListBoxCsvPrésents.SelectedItems)
                if (listBoxItem.ToString().Contains("competence"))
                {
                    File.AppendAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_annee.txt",
                        Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
                    BtnSuppressionFichierCsv.Enabled = true;
                }
            SélectionPériode(new object(), new EventArgs());
            if (ListBoxCsvPrésents.SelectedItems.Count != 0 &&
                ListBoxCsvPrésents.SelectedItem.ToString().Contains("competence"))
                BtnSuppressionFichierCsv.Enabled = true;
            else BtnSuppressionFichierCsv.Enabled = false;
        }

        private void SélectionFichierXlsxDocxPrésent(object sender, EventArgs e)
        {
            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_dnb.txt", string.Empty);
            var listeDnbXlsx = new ListBox();
            var listeDocxXlsx = new ListBox();
            var listeAnnéeXlsx = new ListBox();
            var listeSélection = new ListBox();
            foreach (var listBoxItem in ListBoxXlsxPrésents.SelectedItems)
            {
                listeSélection.Items.Add(listBoxItem.ToString());
                if (listBoxItem.ToString().Contains("DNB-") && listBoxItem.ToString().Contains("xlsx"))
                {
                    File.AppendAllText(CheminElyco + @"\ELyco\Config\ELyco_classes_dnb.txt",
                        Path.GetFileName(listBoxItem.ToString()).Substring(0, 17) + Environment.NewLine);
                    listeDnbXlsx.Items.Add(listBoxItem.ToString());
                }
                if (listBoxItem.ToString().Contains("docx") || listBoxItem.ToString().Contains("xlsx") || listBoxItem.ToString().Contains("pdf"))
                    listeDocxXlsx.Items.Add(listBoxItem.ToString());
                if (listBoxItem.ToString().Contains("Annee"))
                    listeAnnéeXlsx.Items.Add(listBoxItem.ToString());
            }

            foreach (var item in listeSélection.Items)
                if (listeDnbXlsx.Items.Contains(item))
                {
                    BtnGénérerPublipostageDnb.Enabled = true;
                }
                else
                {
                    BtnGénérerPublipostageDnb.Enabled = false;
                    break;
                }

            foreach (var item in listeSélection.Items)
                if (listeDocxXlsx.Items.Contains(item))
                {
                    BtnSuppressionFichierXlsx.Enabled = true;
                }
                else
                {
                    BtnSuppressionFichierXlsx.Enabled = false;
                    break;
                }
            foreach (var item in listeSélection.Items)
                if (listeAnnéeXlsx.Items.Contains(item))
                {
                    BtnGénérerfichiersExcelDnb.Enabled = true;
                }
                else
                {
                    BtnGénérerfichiersExcelDnb.Enabled = false;
                    break;
                }

            if (ListBoxXlsxPrésents.SelectedItems.Count == 0)

            {
                BtnGénérerfichiersExcelDnb.Enabled = false;
                BtnGénérerPublipostageDnb.Enabled = false;
                BtnSuppressionFichierXlsx.Enabled = false;
            }
        }

        private void SélectionPériode(object sender, EventArgs e)
        {
            DétectionPériode();
            if (DétectionPériode() != null && DétectionPériode().Contains("période") &&
                ListBoxCsvATraiter.Items.Count != 0 || DétectionPériode() != null &&
                DétectionPériode().Contains("Année") && ListBoxCsvPrésents.SelectedItems.Count != 0 &&
                ListBoxCsvPrésents.SelectedItem.ToString().Contains("competence"))
                BtnLancerTraitement.Enabled = true;
            else BtnLancerTraitement.Enabled = false;
        }

        private void GlisserDéplacerCsvAtraiter(object sender, DragEventArgs e)
        {
            var fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (var file in fileList)
            {
                var filename = Path.GetFullPath(file);
                ListBoxCsvATraiter.Items.Add(filename);
            }

            File.WriteAllText(CheminElyco + @"\ELyco\Config\ELyco_classes.txt", string.Empty);

            foreach (var listBoxItem in ListBoxCsvATraiter.Items)
            {
                if (!File.Exists(LblCheminDossierCsv.Text + @"\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(),
                        LblCheminDossierCsv.Text + @"\" + Path.GetFileName(listBoxItem.ToString()));

                File.AppendAllText(CheminElyco + @"\ELyco\Config\ELyco_classes.txt",
                    Path.GetFileName(listBoxItem.ToString()).Substring(25, 2) + Environment.NewLine);
            }

            RafraichirListbox();
        }

        private void GlisserValiderCsvAtraiter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
                e.Effect = DragDropEffects.All;
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

        private void SuppressionFichiersSauvegarde(string chemin)
        {
            var dossier = new DirectoryInfo(chemin);

            if (Directory.Exists(dossier.ToString()))
            {
                foreach (var file in dossier.GetFiles())
                    file.Delete();
                foreach (var dir in dossier.GetDirectories())
                    dir.Delete(true);
            }
        }

        private void EffacerListbox(ListBox liste)
        {
            for (var i = liste.Items.Count - 1; i >= 0; i--)
                liste.Items.RemoveAt(i);
        }

        private void CréationArborescence(string niveau)
        {
            var classe = 'A';

            var combo = (ComboBox)Controls.Find(string.Format("ComboNiveau" + niveau), false).FirstOrDefault();
            if (combo != null)
                for (var i = 1; i <= int.Parse(combo.Items[combo.SelectedIndex].ToString()); i++)
                {
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + @"\" + "1ère période" + @"\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + @"\" + "2ème période" + @"\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + @"\" + "3ème période" + @"\" + niveau +
                                              classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + @"\" + "Année" + @"\" + niveau + classe);
                    Directory.CreateDirectory(LblCheminDossierCsv.Text + @"\" + niveau + classe);
                    classe++; // c1 is 'B' now
                }
        }

        private void VérifierCheminsDossiers()
        {
            if (File.Exists(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                using (TextReader tr = new StreamReader(CheminElyco + @"\ELyco\Config\ELyco_in.txt"))
                {
                    LblCheminDossierCsv.Text = tr.ReadLine() + @"\";
                }
            if (File.Exists(CheminElyco + @"\ELyco\Config\ELyco_out.txt"))
                using (TextReader tr1 = new StreamReader(CheminElyco + @"\ELyco\Config\ELyco_out.txt"))
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
                    .GetFiles(LblCheminDossierCsv.Text + période + @"\" + classe + @"\", "*.csv")
                    .Select(Path.GetFileName))
                {
                    var classe1 = s.Substring(25, 2);

                    if (classe == classe1)
                    {
                        MessageBox.Show(@"Il existe déjà un fichier pour la classe " + classe1 + @" dans le dossier '" +
                                        période + @"'");
                        templist.Items.Add(t);
                        File.Delete(LblCheminDossierCsv.Text + t);
                    }
                }
            }

            foreach (var v in templist.Items)
                ListBoxCsvATraiter.Items.Remove(v);

            EffacerListbox(ListBoxCsvATraiter);
            VérifierCheminsDossiers();
            RemplirListeCsvATraiter();
            LblFichiersCsvATraiter.Text = ListBoxCsvATraiter.Items.Count + @" classes à traiter";
        }

        private static void ChangerLigneFichierTxt(string newText, string fileName, int lineToEdit)
        {
            var arrLine = File.ReadAllLines(fileName);
            arrLine[lineToEdit - 1] = newText;
            File.WriteAllLines(fileName, arrLine);
        }

        private void ListeFichiersPrésents(string directoryPath, string periode, ListBox liste)

        {
            var directoryInfo = new DirectoryInfo(directoryPath + periode);

            if (directoryInfo.Exists)

            {
                var fileInfo = directoryInfo.GetFiles();

                var subdirectoryInfo = directoryInfo.GetDirectories();

                if (liste != ListBoxCsvATraiter)
                    foreach (var subDirectory in subdirectoryInfo)

                        ListeFichiersPrésents(subDirectory.FullName, "", liste);

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

        private void CopieFichiersTypeDnb(Stream input, Stream output)
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
            var fichiers = Directory.GetFiles(LblCheminDossierXlsx.Text + @"Année\");

            foreach (var file in fichiers)
            {
                var classe = Path.GetFileNameWithoutExtension(file).Substring(17);
                var fichier = Path.GetFileName(file);
                foreach (var fichierSélectionné in ListBoxXlsxPrésents.SelectedItems)
                    if (fichierSélectionné.ToString() == fichier)
                    {
                        var strPath = LblCheminDossierXlsx.Text + @"DNB\" + "Type_dnb.xlsx";
                        if (File.Exists(strPath)) File.Delete(strPath);
                        var assembly = Assembly.GetExecutingAssembly();
                        //In the next line you should provide      NameSpace.FileName.Extension that you have embedded
                        var input = assembly.GetManifestResourceStream("Compétences.Resources.Type_dnb.xlsx");
                        var output = File.Open(strPath, FileMode.CreateNew);
                        CopieFichiersTypeDnb(input, output);
                        input?.Dispose();
                        output.Dispose();

                        var strPath1 = LblCheminDossierXlsx.Text + @"DNB\" + "Type_dnb.docx";
                        if (File.Exists(strPath1)) File.Delete(strPath1);
                        var assembly1 = Assembly.GetExecutingAssembly();
                        var input1 = assembly1.GetManifestResourceStream("Compétences.Resources.Type_dnb.docx");
                        var output1 = File.Open(strPath1, FileMode.CreateNew);
                        CopieFichiersTypeDnb(input1, output1);
                        input1?.Dispose();
                        output1.Dispose();

                        var excelApplication = new Application();

                        var srcPath = LblCheminDossierXlsx.Text + @"Année\" + fichier;
                        var srcworkBook = excelApplication.Workbooks.Open(srcPath);
                        var srcworkSheet = (Worksheet)srcworkBook.Sheets.Item[1];

                        var destPath = strPath;
                        var destworkBook = excelApplication.Workbooks.Open(destPath, 0, false);
                        var destworkSheet = (Worksheet)destworkBook.Sheets.Item[1];
                        var destworkSheet2 = (Worksheet)destworkBook.Sheets.Item[2];

                        var range = srcworkSheet.Range["A2:A50"];
                        var cnt = -3;

                        foreach (Microsoft.Office.Interop.Excel.Range element in range.Cells)

                            if (element.Value2 != null)
                                cnt = cnt + 1;

                        var from = srcworkSheet.Range["B1:J" + (cnt + 1)]; //Copie tableau compétences
                        var to = destworkSheet.Range["AI1"]; //à modifier

                        var from1 = srcworkSheet.Range["A2:A" + (cnt + 1)]; //Copie noms vers récapitilatif
                        var to1 = destworkSheet.Range["A2"];

                        var from2 = srcworkSheet.Range["A2:A" + (cnt + 1)]; //Copie noms vers épreuves écrites
                        var to2 = destworkSheet2.Range["A2"];

                        from.Copy(to);
                        from1.Copy(to1);
                        from2.Copy(to2);

                        var cells1 = destworkSheet.Range["B2:B" + (cnt + 1)]; //Copie classe vers récapitulatif
                        cells1.Value = classe;

                        var cells = destworkSheet.Range[
                            "A" + (cnt + 2) + ":A500"]; //Nettoyage bas tableau récapitulatif

                        var del = cells.EntireRow;

                        del.Delete();

                        destworkBook.SaveAs(LblCheminDossierXlsx.Text + @"DNB\DNB-" + classe + "_" +
                                            DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                        srcworkBook.Close();
                        destworkBook.Close();
                    }
            }
        }

        private int CompterFichiersPrésents(ListBox listbox)
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

        private void RemplirListeCsvATraiter()
        {
            ListeFichiersPrésents(LblCheminDossierCsv.Text, "", ListBoxCsvATraiter);
        }

        private void RemplirListeCsvPrésents()
        {
            foreach (RadioButton période in PanelTrimestre.Controls)
            {
                ListBoxCsvPrésents.Items.Add(période.Text);
                ListBoxCsvPrésents.Items.Add("-----------------------------------");
                ListeFichiersPrésents(LblCheminDossierCsv.Text, période.Text + @"\", ListBoxCsvPrésents);
                ListBoxCsvPrésents.Items.Add("");
            }
        }

        private void RemplirListeXlsxPrésents()
        {
            foreach (RadioButton période in PanelTrimestre.Controls)
            {
                ListBoxXlsxPrésents.Items.Add(période.Text);
                ListBoxXlsxPrésents.Items.Add("-----------------------------------");
                ListeFichiersPrésents(LblCheminDossierXlsx.Text, période.Text + @"\", ListBoxXlsxPrésents);
                ListBoxXlsxPrésents.Items.Add("");
            }
            ListBoxXlsxPrésents.Items.Add("DNB");
            ListBoxXlsxPrésents.Items.Add("-----------------------------------");
            ListeFichiersPrésents(LblCheminDossierXlsx.Text, "DNB" + @"\", ListBoxXlsxPrésents);
        }

        private void RafraichirListbox()
        {
            EffacerListbox(ListBoxCsvATraiter);
            EffacerListbox(ListBoxCsvPrésents);
            EffacerListbox(ListBoxXlsxPrésents);
            VérifierCheminsDossiers();
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            RemplirListeCsvATraiter();
            LblFichiersCsvATraiter.Text = ListBoxCsvATraiter.Items.Count + @" classes à traiter";
            LblFichiersCsvPrésents.Text = CompterFichiersPrésents(ListBoxCsvPrésents) + @" fichiers CSV";
            LblFichiersXlsxPrésents.Text = CompterFichiersPrésents(ListBoxXlsxPrésents) + @" fichiers XLSX et " +
                                           CompterFichiersDnb(ListBoxXlsxPrésents) + @" fichiers DNB";
            SélectionPériode(new object(), new EventArgs());
        }

        private string DétectionPériode()
        {
            string périodeSelect = null;
            foreach (RadioButton période in PanelTrimestre.Controls)
                if (période.Checked)
                    périodeSelect = période.Text;
            return périodeSelect;
        }

        private void SupprimerObjets(object obj)
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

        private void ExécuterMacro(string macro)
        {
            //~~> Define your Excel Objects
            var xlApp = new Application();

            var sPath = Path.GetTempFileName();
            File.WriteAllBytes(sPath, Resources.Compétences);

            //~~> Start Excel and open the workbook.
            var xlWorkBook = xlApp.Workbooks.Open(sPath);

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run(macro);

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();

            //~~> Clean Up
            SupprimerObjets(xlApp);
            SupprimerObjets(xlWorkBook);
        }
    }
}