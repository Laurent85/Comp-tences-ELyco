using System;
using System.IO;
using System.Windows.Forms;
using System.Reflection;
using Compétences.Properties;
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
            Executer_macro("Deplacer_P1");
            Executer_macro("Compétences_par_lot_P1");
        }

        //~> Release the objects

            private void Executer_macro(string macro)
        {
            //~~> Define your Excel Objects
            Excel.Application xlApp = new Excel.Application();

            Excel.Workbook xlWorkBook;

            string sPath = System.IO.Path.GetTempFileName();
            System.IO.File.WriteAllBytes(sPath, Properties.Resources.Compétences);

            //~~> Start Excel and open the workbook.
            xlWorkBook = xlApp.Workbooks.Open(sPath);

            //~~> Run the macros by supplying the necessary arguments
            xlApp.Run(macro);

            //~~> Clean-up: Close the workbook
            xlWorkBook.Close(false);

            //~~> Quit the Excel Application
            xlApp.Quit();

            //~~> Clean Up
            releaseObject(xlApp);
            releaseObject(xlWorkBook);
        }
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
                string folder = dlg.SelectedPath + "\\" + Annee_scolaire.SelectedItem;
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
            if (File.Exists("C:\\ELyco1.txt"))
            {
                using (TextReader tr1 = new StreamReader("C:\\ELyco1.txt"))
                {
                    Chemin_destination.Text = tr1.ReadLine() + @"\";
                }
            }
        }

        private void Dossier_destination_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            // This is what will execute if the user selects a folder and hits OK (File if you change to FileBrowserDialog)
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                string folder = dlg.SelectedPath + "\\" + Annee_scolaire.SelectedItem;
                Chemin_destination.Text = folder;

                if (!File.Exists("C:\\ELyco1.txt"))
                {
                    using (File.Create("C:\\ELyco1.txt"))
                    {
                    }
                }
                using (StreamWriter sw = new StreamWriter("C:\\ELyco1.txt"))
                {
                    sw.Write(Chemin_destination.Text);
                }
            }
            else
            {
                // This prevents a crash when you close out of the window with nothing
            }
        }

        private void Créer_arborescence_destination_Click(object sender, EventArgs e)
        {        
            
                Directory.CreateDirectory(Chemin_destination.Text + "\\" + "1ère période");
                Directory.CreateDirectory(Chemin_destination.Text + "\\" + "2ème période");
                Directory.CreateDirectory(Chemin_destination.Text + "\\" + "3ème période");
                Directory.CreateDirectory(Chemin_destination.Text + "\\" + "Année");             
            
        }
    }
}