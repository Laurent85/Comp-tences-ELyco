using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Compétences
{
    public partial class Principal : Form
    {
        public Principal()
        {
            InitializeComponent();
        }

        private void Principal_Load(object sender, EventArgs e)
        {
            Directory.CreateDirectory("C:\\ELyco");
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
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
            Lancer_traitement.Enabled = false;
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
                // This prevents a crash when you close out of the window with nothing
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
                // This prevents a crash when you close out of the window with nothing
            }
        }

        private void Créer_arborescence_Click(object sender, EventArgs e)
        {
            char c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_6.Items[Niveau_6.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "1ère période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "2ème période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3ème période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "6" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_5.Items[Niveau_5.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "1ère période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "2ème période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3ème période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "5" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_4.Items[Niveau_4.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "1ère période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "2ème période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3ème période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "4" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_3.Items[Niveau_3.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "1ère période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "2ème période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3ème période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "3" + c1);
                c1++; // c1 is 'B' now
            }

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

        private void Lancer_traitement_Click(object sender, EventArgs e)
        {
            if (bouton_periode1.Checked)
            {
                Executer_macro("Deplacer_P1.Deplacer_P1");
                Executer_macro("Compétences_par_lot_P1.Compétences_par_lot_P1");
                MessageBox.Show(@"Traitement terminé");
                EffacerListbox(Liste_CSV);
                EffacerListbox(Liste_csv_présents);
                Liste_CSV.Refresh();
                Liste_csv_présents.Refresh();
            }
            if (bouton_periode2.Checked)
            {
                Executer_macro("Deplacer_P2.Deplacer_P2");
                Executer_macro("Compétences_par_lot_P2.Compétences_par_lot_P2");
                MessageBox.Show(@"Traitement terminé");
                EffacerListbox(Liste_CSV);
                EffacerListbox(Liste_csv_présents);
                Liste_CSV.Refresh();
                Liste_csv_présents.Refresh();
            }
            if (bouton_periode3.Checked)
            {
                Executer_macro("Deplacer_P3.Deplacer_P3");
                Executer_macro("Compétences_par_lot_P3.Compétences_par_lot_P3");
                MessageBox.Show(@"Traitement terminé");
                EffacerListbox(Liste_CSV);
                EffacerListbox(Liste_csv_présents);
                Liste_CSV.Refresh();
                Liste_csv_présents.Refresh();
            }
            if (bouton_annee.Checked)
            {
                Executer_macro("Fusionner.Fusionner");
                Executer_macro("Compétences_par_lot_Année.Compétences_par_lot_Année");
                MessageBox.Show(@"Traitement terminé");
                EffacerListbox(Liste_CSV);
                EffacerListbox(Liste_csv_présents);
                Liste_CSV.Refresh();
                Liste_csv_présents.Refresh();
            }
            RemplirListeCsvPrésents();
            RemplirListeXlsxPrésents();
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
        }

        private void SuppressionFichierXlsx_Click(object sender, EventArgs e)
        {
            Suppression_fichiers(Chemin_destination.Text, Liste_xlsx_présents);
        }

        private void Supprimer_tout_Click(object sender, EventArgs e)
        {
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
                Liste_CSV.Items.Add(filename);
            }

            foreach (var listBoxItem in Liste_CSV.Items)
            {
                if (!File.Exists(Chemin_dossier.Text + "\\" + Path.GetFileName(listBoxItem.ToString())))
                    File.Copy(listBoxItem.ToString(), Chemin_dossier.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));
            }

            lbl_fichiers_csv_a_traiter.Text = Liste_CSV.Items.Count.ToString() + @" classes à traiter";
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