using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

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
            Vérifier_chemins_dossiers();
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            Lancer_traitement.Enabled = false;
            lbl_fichiers_csv_conservés.Text = Compter_fichiers(Liste_csv_présents) + @" fichiers CSV présents";
            lbl_fichiers_xlsx.Text = Compter_fichiers(Liste_xlsx_présents) + @" fichiers XLSX présents";
            try
            {
                Annee_scolaire.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(2).Take(3).First();
                Niveau_6.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(3).Take(4).First();
                Niveau_5.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(4).Take(5).First();
                Niveau_4.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(5).Take(6).First();
                Niveau_3.Text = File.ReadLines("C:\\ELyco\\ELyco_in.txt").Skip(6).Take(7).First();
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
                string folder = dlg.SelectedPath + "\\ELyco_CSV\\" + Annee_scolaire.SelectedItem;
                Chemin_dossier.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_in.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_in.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco\\ELyco_in.txt"))
                {
                    sw.WriteLine(Chemin_dossier.Text);
                    sw.WriteLine(dlg.SelectedPath + "\\ELyco_CSV" + "\n");
                }
            }
            else
            {
                
            }
        }

        private void Dossier_destination_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string folder = dlg.SelectedPath + "\\ELyco_Competences\\" + Annee_scolaire.SelectedItem;
                Chemin_destination.Text = folder;
                Directory.CreateDirectory("C:\\ELyco");

                if (!File.Exists("C:\\ELyco\\ELyco_out.txt"))
                {
                    using (File.Create("C:\\ELyco\\ELyco_out.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco\\ELyco_out.txt"))
                {
                    sw.WriteLine(Chemin_destination.Text);
                    sw.WriteLine(dlg.SelectedPath + "\\ELyco_Competences");
                }
            }
            else
            {
                
            }
        }

        private void Créer_arborescence_Click(object sender, EventArgs e)
        {
            Création_arborescence("6");
            Création_arborescence("5");
            Création_arborescence("4");
            Création_arborescence("3");

            Directory.CreateDirectory(Chemin_destination.Text + "\\" + "1ère période");
            Directory.CreateDirectory(Chemin_destination.Text + "\\" + "2ème période");
            Directory.CreateDirectory(Chemin_destination.Text + "\\" + "3ème période");
            Directory.CreateDirectory(Chemin_destination.Text + "\\" + "Année");

            Changer_ligne_fichier_txt(Annee_scolaire.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 3);
            Changer_ligne_fichier_txt(Niveau_6.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 4);
            Changer_ligne_fichier_txt(Niveau_5.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 5);
            Changer_ligne_fichier_txt(Niveau_4.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 6);
            Changer_ligne_fichier_txt(Niveau_3.SelectedItem + "\n", "C:\\ELyco\\ELyco_in.txt", 7);
        }

        public void Lancer_traitement_Click(object sender, EventArgs e)
        {
            BackgroundWorker traitementMacro = new BackgroundWorker();
            traitementMacro.DoWork += traitementMacro_Lancement;
            traitementMacro.RunWorkerCompleted += traitementMacro_Fini;
            traitementMacro.RunWorkerAsync();
            traitementMacro.WorkerSupportsCancellation = true;
            Frm2.message = lbl_fichiers_csv_a_traiter.Text + "...Veuillez patienter...";
            Frm2.ShowDialog();
        }

        private void traitementMacro_Lancement(object sender, DoWorkEventArgs e)
        {
            if (bouton_periode1.Checked)
            {
                Executer_macro("Deplacer_P1.Deplacer_P1");
                Executer_macro("Compétences_par_lot_P1.Compétences_par_lot_P1");
            }
            if (bouton_periode2.Checked)
            {
                Executer_macro("Deplacer_P2.Deplacer_P2");
                Executer_macro("Compétences_par_lot_P2.Compétences_par_lot_P2");
            }
            if (bouton_periode3.Checked)
            {
                Executer_macro("Deplacer_P3.Deplacer_P3");
                Executer_macro("Compétences_par_lot_P3.Compétences_par_lot_P3");
            }
            if (bouton_annee.Checked)
            {
                Executer_macro("Fusionner.Fusionner");
                Executer_macro("Compétences_par_lot_Année.Compétences_par_lot_Année");
            }
        }

        private void traitementMacro_Fini(object sender, RunWorkerCompletedEventArgs e)
        {
            EffacerListbox(Liste_CSV_a_traiter);
            EffacerListbox(Liste_csv_présents);
            EffacerListbox(Liste_xlsx_présents);
            Vérifier_chemins_dossiers();
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            lbl_fichiers_csv_a_traiter.Text = "";
            lbl_fichiers_csv_conservés.Text = Compter_fichiers(Liste_csv_présents) + @" fichiers CSV présents";
            lbl_fichiers_xlsx.Text = Compter_fichiers(Liste_xlsx_présents) + @" fichiers XLSX présents";

            string[] files = Directory.GetFiles(Chemin_dossier.Text, "*.*");
            foreach (string file in files)
            {
                File.Delete(file);
            }

            Frm2.Close();
            MessageBox.Show(@"Traitement terminé");
        }

        private void bouton_periode1_CheckedChanged(object sender, EventArgs e)
        {
            Lancer_traitement.Enabled = true;
        }

        private void bouton_periode2_CheckedChanged(object sender, EventArgs e)
        {
            Lancer_traitement.Enabled = true;
        }

        private void bouton_periode3_CheckedChanged(object sender, EventArgs e)
        {
            Lancer_traitement.Enabled = true;
        }

        private void bouton_annee_CheckedChanged(object sender, EventArgs e)
        {
            Lancer_traitement.Enabled = true;
        }

        private void SuppressionFichierCsv_Click(object sender, EventArgs e)
        {
            Suppression_fichiers(Chemin_dossier.Text, Liste_csv_présents);
            lbl_fichiers_csv_conservés.Text = Compter_fichiers(Liste_csv_présents) + @" fichiers CSV présents";
            lbl_fichiers_xlsx.Text = Compter_fichiers(Liste_xlsx_présents) + @" fichiers XLSX présents";
        }

        private void SuppressionFichierXlsx_Click(object sender, EventArgs e)
        {
            Suppression_fichiers(Chemin_destination.Text, Liste_xlsx_présents);
            lbl_fichiers_csv_conservés.Text = Compter_fichiers(Liste_csv_présents) + @" fichiers CSV présents";
            lbl_fichiers_xlsx.Text = Compter_fichiers(Liste_xlsx_présents) + @" fichiers XLSX présents";
        }

        private void Supprimer_tout_Click(object sender, EventArgs e)
        {
            try
            {
                Suppression_fichiers(Chemin_dossier.Text, Liste_csv_présents);
            }
            catch
            {
                // ignored
            }
            try
            {
                Suppression_fichiers(Chemin_destination.Text, Liste_xlsx_présents);
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

            Chemin_dossier.Text = "";
            Chemin_destination.Text = "";
            Annee_scolaire.Text = "";
            Niveau_6.Text = "";
            Niveau_5.Text = "";
            Niveau_4.Text = "";
            Niveau_3.Text = "";

            EffacerListbox(Liste_CSV_a_traiter);
            EffacerListbox(Liste_csv_présents);
            EffacerListbox(Liste_xlsx_présents);
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            Liste_CSV_a_traiter.Refresh();
            Liste_csv_présents.Refresh();
            Liste_xlsx_présents.Refresh();
            lbl_fichiers_csv_a_traiter.Text = "";
        }

        private void Liste_xlsx_présents_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string fichierP1 = Chemin_destination.Text + "1ère période\\" + Liste_xlsx_présents.SelectedItem;
            string fichierP2 = Chemin_destination.Text + "2ème période\\" + Liste_xlsx_présents.SelectedItem;
            string fichierP3 = Chemin_destination.Text + "3ème période\\" + Liste_xlsx_présents.SelectedItem;
            string fichierAnnee = Chemin_destination.Text + "Année\\" + Liste_xlsx_présents.SelectedItem;
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
        }

        private void Création_arborescence(string niveau)
        {
            char classe = 'A';

            ComboBox combo = (ComboBox)Controls.Find(string.Format("Niveau_" + niveau), false).FirstOrDefault();
            if (combo != null)
                for (int i = 1; i <= int.Parse(combo.Items[combo.SelectedIndex].ToString()); i++)
                {
                    Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "1ère période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "2ème période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3ème période" + "\\" + niveau + classe);
                    Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année" + "\\" + niveau + classe);
                    Directory.CreateDirectory(Chemin_dossier.Text + "\\" + niveau + classe);
                    classe++; // c1 is 'B' now
                }
        }

        private void Vérifier_chemins_dossiers()
        {
            if (File.Exists("C:\\ELyco\\ELyco_in.txt"))
            {
                using (TextReader tr = new StreamReader("C:\\ELyco\\ELyco_in.txt"))
                {
                    Chemin_dossier.Text = tr.ReadLine() + @"\";
                }
            }
            if (File.Exists("C:\\ELyco\\ELyco_out.txt"))
            {
                using (TextReader tr1 = new StreamReader("C:\\ELyco\\ELyco_out.txt"))
                {
                    Chemin_destination.Text = tr1.ReadLine() + @"\";
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

                foreach (DirectoryInfo subDirectory in subdirectoryInfo)

                {
                    Liste_fichiers_présents(subDirectory.FullName, "", liste);
                }

                foreach (FileInfo file in fileInfo)

                {
                    if (file.Length > 2000)
                        liste.Items.Add(file.Name);
                    else
                        file.Delete();
                }
            }
        }

        private int Compter_fichiers(ListBox listbox)
        {
            int nombre = 0;
            foreach (var item in listbox.Items)
            {
                if (item.ToString().Contains("competence"))
                {
                    nombre++;
                }
            }
            return nombre;
        }

        private void RemplirListeCsvPrésents()
        {
            Liste_csv_présents.Items.Add("1ère période");
            Liste_csv_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_dossier.Text, "1ère période" + "\\", Liste_csv_présents);
            Liste_csv_présents.Items.Add("");
            Liste_csv_présents.Items.Add("2ème période");
            Liste_csv_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_dossier.Text, "2ème période" + "\\", Liste_csv_présents);
            Liste_csv_présents.Items.Add("");
            Liste_csv_présents.Items.Add("3ème période");
            Liste_csv_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_dossier.Text, "3ème période" + "\\", Liste_csv_présents);
            Liste_csv_présents.Items.Add("");
            Liste_csv_présents.Items.Add("Année");
            Liste_csv_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_dossier.Text, "Année" + "\\", Liste_csv_présents);
        }

        private void RemplirListeXlsxPrésents()
        {
            Liste_xlsx_présents.Items.Add("1ère période");
            Liste_xlsx_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_destination.Text, "1ère période" + "\\", Liste_xlsx_présents);
            Liste_xlsx_présents.Items.Add("");
            Liste_xlsx_présents.Items.Add("2ème période");
            Liste_xlsx_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_destination.Text, "2ème période" + "\\", Liste_xlsx_présents);
            Liste_xlsx_présents.Items.Add("");
            Liste_xlsx_présents.Items.Add("3ème période");
            Liste_xlsx_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_destination.Text, "3ème période" + "\\", Liste_xlsx_présents);
            Liste_xlsx_présents.Items.Add("");
            Liste_xlsx_présents.Items.Add("Année");
            Liste_xlsx_présents.Items.Add("-----------------------------------");
            Liste_fichiers_présents(Chemin_destination.Text, "Année" + "\\", Liste_xlsx_présents);
        }

        private void Suppression_fichiers(string chemin, ListBox liste)
        {
            string[] files = Directory.GetFiles(chemin, "*.*", SearchOption.AllDirectories);

            foreach (string file in files)
            {
                foreach (var selecteditem in liste.SelectedItems)
                {
                    if (file.Contains(selecteditem.ToString()))
                        File.Delete(file);
                }
            }
            var selectedItems = liste.SelectedItems;

            if (liste.SelectedIndex != -1)
            {
                for (int i = selectedItems.Count - 1; i >= 0; i--)
                    liste.Items.Remove(selectedItems[i]);
            }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
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
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;

            string sPath = Path.GetTempFileName();
            File.WriteAllBytes(sPath, Properties.Resources.Compétences);

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

        private void Drag(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string file in fileList)
            {
                string filename = Path.GetFullPath(file);
                Liste_CSV_a_traiter.Items.Add(filename);
            }

            foreach (var listBoxItem in Liste_CSV_a_traiter.Items)
            {
                if (!File.Exists(Chemin_dossier.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(), Chemin_dossier.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));
            }

            lbl_fichiers_csv_a_traiter.Text = Liste_CSV_a_traiter.Items.Count + @" classes à traiter";
        }

        private void Drag_Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
        }
    }
}