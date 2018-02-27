using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Compétences
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //~~> Define your Excel Objects
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;

            //~~> Start Excel and open the workbook.
            xlWorkBook = xlApp.Workbooks.Open("C:\\Users\\User\\Desktop\\Compétences.xlsm");

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run("Deplacer_P1");

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();

            //~~> Clean Up
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
        }

        //~> Release the objects
        private void releaseObject(object obj)
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
                File.Copy(listBoxItem.ToString(), Chemin_dossier.Text + "\\" + Path.GetFileName(listBoxItem.ToString()));
            }
        }

        private void Drag_Enter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void Dossier_travail_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            // This is what will execute if the user selects a folder and hits OK (File if you change to FileBrowserDialog)
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string folder = dlg.SelectedPath;
                Chemin_dossier.Text = folder;

                if (!File.Exists("C:\\ELyco.txt"))
                {
                    using (File.Create("C:\\ELyco.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco.txt"))
                {
                    sw.Write(Chemin_dossier.Text);
                }
            }
            else
            {
                // This prevents a crash when you close out of the window with nothing
            }
        }

        private void Créer_arborescence_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem);
            Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "1ère période");
            Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "2ème période");
            Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3ème période");
            Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "Année");

            char c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_6.Items[Niveau_6.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "1ère période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "2ème période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3ème période" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "Année" + "\\" + "6" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "6" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_5.Items[Niveau_5.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "1ère période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "2ème période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3ème période" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "Année" + "\\" + "5" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "5" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_4.Items[Niveau_4.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "1ère période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "2ème période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3ème période" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "Année" + "\\" + "4" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "4" + c1);
                c1++; // c1 is 'B' now
            }

            c1 = 'A';

            for (int i = 1; i <= int.Parse(Niveau_3.Items[Niveau_3.SelectedIndex].ToString()); i++)
            {
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "1ère période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "2ème période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3ème période" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "Année" + "\\" + "3" + c1);
                Directory.CreateDirectory(Chemin_dossier.Text + "\\" + "Année scolaire " + Annee_scolaire.SelectedItem + "\\" + "3" + c1);
                c1++; // c1 is 'B' now
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists("C:\\ELyco.txt"))
            {
                using (TextReader tr = new StreamReader("C:\\ELyco.txt"))
                {
                    Chemin_dossier.Text = tr.ReadLine() + @"\";
                }
            }
        }
    }
}