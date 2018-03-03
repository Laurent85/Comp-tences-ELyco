namespace Compétences
{
    partial class Principal
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Principal));
            this.Lancer_traitement = new System.Windows.Forms.Button();
            this.Liste_CSV = new System.Windows.Forms.ListBox();
            this.Niveau_6 = new System.Windows.Forms.ComboBox();
            this.Niveau_5 = new System.Windows.Forms.ComboBox();
            this.Niveau_4 = new System.Windows.Forms.ComboBox();
            this.Niveau_3 = new System.Windows.Forms.ComboBox();
            this.Annee_scolaire = new System.Windows.Forms.ComboBox();
            this.Dossier_travail = new System.Windows.Forms.Button();
            this.Chemin_dossier = new System.Windows.Forms.Label();
            this.bouton_periode1 = new System.Windows.Forms.RadioButton();
            this.bouton_periode2 = new System.Windows.Forms.RadioButton();
            this.bouton_periode3 = new System.Windows.Forms.RadioButton();
            this.bouton_annee = new System.Windows.Forms.RadioButton();
            this.Créer_arborescence = new System.Windows.Forms.Button();
            this.Dossier_destination = new System.Windows.Forms.Button();
            this.Chemin_destination = new System.Windows.Forms.Label();
            this.Liste_csv_présents = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.Supprimer_tout = new System.Windows.Forms.Button();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.SuppressionFichierCsv = new System.Windows.Forms.Button();
            this.Liste_xlsx_présents = new System.Windows.Forms.ListBox();
            this.SuppressionFichierXlsx = new System.Windows.Forms.Button();
            this.lbl_fichiers_csv_a_traiter = new System.Windows.Forms.Label();
            this.lbl_fichiers_csv_conservés = new System.Windows.Forms.Label();
            this.lbl_fichiers_xlsx = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // Lancer_traitement
            // 
            this.Lancer_traitement.Location = new System.Drawing.Point(443, 553);
            this.Lancer_traitement.Name = "Lancer_traitement";
            this.Lancer_traitement.Size = new System.Drawing.Size(109, 23);
            this.Lancer_traitement.TabIndex = 0;
            this.Lancer_traitement.Text = "Lancer le traitement";
            this.Lancer_traitement.UseVisualStyleBackColor = true;
            this.Lancer_traitement.Click += new System.EventHandler(this.Lancer_traitement_Click);
            // 
            // Liste_CSV
            // 
            this.Liste_CSV.AllowDrop = true;
            this.Liste_CSV.FormattingEnabled = true;
            this.Liste_CSV.Location = new System.Drawing.Point(12, 320);
            this.Liste_CSV.Name = "Liste_CSV";
            this.Liste_CSV.Size = new System.Drawing.Size(540, 212);
            this.Liste_CSV.TabIndex = 1;
            this.Liste_CSV.DragDrop += new System.Windows.Forms.DragEventHandler(this.Drag);
            this.Liste_CSV.DragEnter += new System.Windows.Forms.DragEventHandler(this.Drag_Enter);
            // 
            // Niveau_6
            // 
            this.Niveau_6.FormattingEnabled = true;
            this.Niveau_6.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15"});
            this.Niveau_6.Location = new System.Drawing.Point(363, 93);
            this.Niveau_6.Name = "Niveau_6";
            this.Niveau_6.Size = new System.Drawing.Size(84, 21);
            this.Niveau_6.TabIndex = 2;
            // 
            // Niveau_5
            // 
            this.Niveau_5.FormattingEnabled = true;
            this.Niveau_5.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15"});
            this.Niveau_5.Location = new System.Drawing.Point(495, 93);
            this.Niveau_5.Name = "Niveau_5";
            this.Niveau_5.Size = new System.Drawing.Size(84, 21);
            this.Niveau_5.TabIndex = 3;
            // 
            // Niveau_4
            // 
            this.Niveau_4.FormattingEnabled = true;
            this.Niveau_4.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15"});
            this.Niveau_4.Location = new System.Drawing.Point(629, 93);
            this.Niveau_4.Name = "Niveau_4";
            this.Niveau_4.Size = new System.Drawing.Size(84, 21);
            this.Niveau_4.TabIndex = 4;
            // 
            // Niveau_3
            // 
            this.Niveau_3.FormattingEnabled = true;
            this.Niveau_3.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15"});
            this.Niveau_3.Location = new System.Drawing.Point(748, 93);
            this.Niveau_3.Name = "Niveau_3";
            this.Niveau_3.Size = new System.Drawing.Size(84, 21);
            this.Niveau_3.TabIndex = 5;
            // 
            // Annee_scolaire
            // 
            this.Annee_scolaire.FormattingEnabled = true;
            this.Annee_scolaire.Items.AddRange(new object[] {
            "2017-2018",
            "2018-2019",
            "2019-2020",
            "2020-2021",
            "2021-2022",
            "2022-2023",
            "2023-2024",
            "2024-2025"});
            this.Annee_scolaire.Location = new System.Drawing.Point(155, 93);
            this.Annee_scolaire.Name = "Annee_scolaire";
            this.Annee_scolaire.Size = new System.Drawing.Size(94, 21);
            this.Annee_scolaire.TabIndex = 6;
            // 
            // Dossier_travail
            // 
            this.Dossier_travail.Location = new System.Drawing.Point(155, 139);
            this.Dossier_travail.Name = "Dossier_travail";
            this.Dossier_travail.Size = new System.Drawing.Size(106, 23);
            this.Dossier_travail.TabIndex = 7;
            this.Dossier_travail.Text = "Dossier des csv";
            this.Dossier_travail.UseVisualStyleBackColor = true;
            this.Dossier_travail.Click += new System.EventHandler(this.Dossier_travail_Click);
            // 
            // Chemin_dossier
            // 
            this.Chemin_dossier.AutoSize = true;
            this.Chemin_dossier.Location = new System.Drawing.Point(276, 144);
            this.Chemin_dossier.Name = "Chemin_dossier";
            this.Chemin_dossier.Size = new System.Drawing.Size(0, 13);
            this.Chemin_dossier.TabIndex = 8;
            // 
            // bouton_periode1
            // 
            this.bouton_periode1.AutoSize = true;
            this.bouton_periode1.Location = new System.Drawing.Point(14, 556);
            this.bouton_periode1.Name = "bouton_periode1";
            this.bouton_periode1.Size = new System.Drawing.Size(84, 17);
            this.bouton_periode1.TabIndex = 10;
            this.bouton_periode1.TabStop = true;
            this.bouton_periode1.Text = "1ère période";
            this.bouton_periode1.UseVisualStyleBackColor = true;
            this.bouton_periode1.CheckedChanged += new System.EventHandler(this.bouton_periode1_CheckedChanged);
            // 
            // bouton_periode2
            // 
            this.bouton_periode2.AutoSize = true;
            this.bouton_periode2.Location = new System.Drawing.Point(116, 556);
            this.bouton_periode2.Name = "bouton_periode2";
            this.bouton_periode2.Size = new System.Drawing.Size(89, 17);
            this.bouton_periode2.TabIndex = 11;
            this.bouton_periode2.TabStop = true;
            this.bouton_periode2.Text = "2ème période";
            this.bouton_periode2.UseVisualStyleBackColor = true;
            this.bouton_periode2.CheckedChanged += new System.EventHandler(this.bouton_periode2_CheckedChanged);
            // 
            // bouton_periode3
            // 
            this.bouton_periode3.AutoSize = true;
            this.bouton_periode3.Location = new System.Drawing.Point(223, 556);
            this.bouton_periode3.Name = "bouton_periode3";
            this.bouton_periode3.Size = new System.Drawing.Size(89, 17);
            this.bouton_periode3.TabIndex = 12;
            this.bouton_periode3.TabStop = true;
            this.bouton_periode3.Text = "3ème période";
            this.bouton_periode3.UseVisualStyleBackColor = true;
            this.bouton_periode3.CheckedChanged += new System.EventHandler(this.bouton_periode3_CheckedChanged);
            // 
            // bouton_annee
            // 
            this.bouton_annee.AutoSize = true;
            this.bouton_annee.Location = new System.Drawing.Point(342, 556);
            this.bouton_annee.Name = "bouton_annee";
            this.bouton_annee.Size = new System.Drawing.Size(56, 17);
            this.bouton_annee.TabIndex = 13;
            this.bouton_annee.TabStop = true;
            this.bouton_annee.Text = "Année";
            this.bouton_annee.UseVisualStyleBackColor = true;
            this.bouton_annee.CheckedChanged += new System.EventHandler(this.bouton_annee_CheckedChanged);
            // 
            // Créer_arborescence
            // 
            this.Créer_arborescence.Location = new System.Drawing.Point(155, 197);
            this.Créer_arborescence.Name = "Créer_arborescence";
            this.Créer_arborescence.Size = new System.Drawing.Size(118, 23);
            this.Créer_arborescence.TabIndex = 9;
            this.Créer_arborescence.Text = "Créer l\'arborescence";
            this.Créer_arborescence.UseVisualStyleBackColor = true;
            this.Créer_arborescence.Click += new System.EventHandler(this.Créer_arborescence_Click);
            // 
            // Dossier_destination
            // 
            this.Dossier_destination.Location = new System.Drawing.Point(155, 168);
            this.Dossier_destination.Name = "Dossier_destination";
            this.Dossier_destination.Size = new System.Drawing.Size(106, 23);
            this.Dossier_destination.TabIndex = 14;
            this.Dossier_destination.Text = "Dossier des xlsx";
            this.Dossier_destination.UseVisualStyleBackColor = true;
            this.Dossier_destination.Click += new System.EventHandler(this.Dossier_destination_Click);
            // 
            // Chemin_destination
            // 
            this.Chemin_destination.AutoSize = true;
            this.Chemin_destination.Location = new System.Drawing.Point(273, 173);
            this.Chemin_destination.Name = "Chemin_destination";
            this.Chemin_destination.Size = new System.Drawing.Size(0, 13);
            this.Chemin_destination.TabIndex = 15;
            // 
            // Liste_csv_présents
            // 
            this.Liste_csv_présents.FormattingEnabled = true;
            this.Liste_csv_présents.Location = new System.Drawing.Point(616, 320);
            this.Liste_csv_présents.Name = "Liste_csv_présents";
            this.Liste_csv_présents.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Liste_csv_présents.Size = new System.Drawing.Size(252, 212);
            this.Liste_csv_présents.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Comic Sans MS", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(270, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(498, 35);
            this.label1.TabIndex = 18;
            this.label1.Text = "Traitement des domaines de compétences";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(158, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(77, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Année scolaire";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(360, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Classes de 6ème";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(492, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 13);
            this.label4.TabIndex = 21;
            this.label4.Text = "Classes de 5ème";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(626, 74);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 13);
            this.label5.TabIndex = 22;
            this.label5.Text = "Classes de 4ème";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(745, 74);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(87, 13);
            this.label6.TabIndex = 23;
            this.label6.Text = "Classes de 3ème";
            // 
            // Supprimer_tout
            // 
            this.Supprimer_tout.Location = new System.Drawing.Point(1025, 91);
            this.Supprimer_tout.Name = "Supprimer_tout";
            this.Supprimer_tout.Size = new System.Drawing.Size(112, 23);
            this.Supprimer_tout.TabIndex = 26;
            this.Supprimer_tout.Text = "Supprimer les bases";
            this.Supprimer_tout.UseVisualStyleBackColor = true;
            this.Supprimer_tout.Click += new System.EventHandler(this.Supprimer_tout_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::Compétences.Properties.Resources.St_Jacques;
            this.pictureBox2.Location = new System.Drawing.Point(14, 9);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(94, 55);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 25;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Compétences.Properties.Resources.E_Lyco;
            this.pictureBox1.Location = new System.Drawing.Point(1025, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(151, 56);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 24;
            this.pictureBox1.TabStop = false;
            // 
            // SuppressionFichierCsv
            // 
            this.SuppressionFichierCsv.Location = new System.Drawing.Point(771, 550);
            this.SuppressionFichierCsv.Name = "SuppressionFichierCsv";
            this.SuppressionFichierCsv.Size = new System.Drawing.Size(97, 23);
            this.SuppressionFichierCsv.TabIndex = 27;
            this.SuppressionFichierCsv.Text = "Supprimer fichier";
            this.SuppressionFichierCsv.UseVisualStyleBackColor = true;
            this.SuppressionFichierCsv.Click += new System.EventHandler(this.SuppressionFichierCsv_Click);
            // 
            // Liste_xlsx_présents
            // 
            this.Liste_xlsx_présents.FormattingEnabled = true;
            this.Liste_xlsx_présents.Location = new System.Drawing.Point(924, 320);
            this.Liste_xlsx_présents.Name = "Liste_xlsx_présents";
            this.Liste_xlsx_présents.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.Liste_xlsx_présents.Size = new System.Drawing.Size(252, 212);
            this.Liste_xlsx_présents.TabIndex = 28;
            // 
            // SuppressionFichierXlsx
            // 
            this.SuppressionFichierXlsx.Location = new System.Drawing.Point(1079, 550);
            this.SuppressionFichierXlsx.Name = "SuppressionFichierXlsx";
            this.SuppressionFichierXlsx.Size = new System.Drawing.Size(97, 23);
            this.SuppressionFichierXlsx.TabIndex = 29;
            this.SuppressionFichierXlsx.Text = "Supprimer fichier";
            this.SuppressionFichierXlsx.UseVisualStyleBackColor = true;
            this.SuppressionFichierXlsx.Click += new System.EventHandler(this.SuppressionFichierXlsx_Click);
            // 
            // lbl_fichiers_csv_a_traiter
            // 
            this.lbl_fichiers_csv_a_traiter.AutoSize = true;
            this.lbl_fichiers_csv_a_traiter.Location = new System.Drawing.Point(60, 280);
            this.lbl_fichiers_csv_a_traiter.Name = "lbl_fichiers_csv_a_traiter";
            this.lbl_fichiers_csv_a_traiter.Size = new System.Drawing.Size(0, 13);
            this.lbl_fichiers_csv_a_traiter.TabIndex = 30;
            // 
            // lbl_fichiers_csv_conservés
            // 
            this.lbl_fichiers_csv_conservés.AutoSize = true;
            this.lbl_fichiers_csv_conservés.Location = new System.Drawing.Point(613, 280);
            this.lbl_fichiers_csv_conservés.Name = "lbl_fichiers_csv_conservés";
            this.lbl_fichiers_csv_conservés.Size = new System.Drawing.Size(0, 13);
            this.lbl_fichiers_csv_conservés.TabIndex = 31;
            // 
            // lbl_fichiers_xlsx
            // 
            this.lbl_fichiers_xlsx.AutoSize = true;
            this.lbl_fichiers_xlsx.Location = new System.Drawing.Point(921, 280);
            this.lbl_fichiers_xlsx.Name = "lbl_fichiers_xlsx";
            this.lbl_fichiers_xlsx.Size = new System.Drawing.Size(0, 13);
            this.lbl_fichiers_xlsx.TabIndex = 32;
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.PaleGoldenrod;
            this.ClientSize = new System.Drawing.Size(1188, 594);
            this.Controls.Add(this.lbl_fichiers_xlsx);
            this.Controls.Add(this.lbl_fichiers_csv_conservés);
            this.Controls.Add(this.lbl_fichiers_csv_a_traiter);
            this.Controls.Add(this.SuppressionFichierXlsx);
            this.Controls.Add(this.Liste_xlsx_présents);
            this.Controls.Add(this.SuppressionFichierCsv);
            this.Controls.Add(this.Supprimer_tout);
            this.Controls.Add(this.pictureBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Liste_csv_présents);
            this.Controls.Add(this.Chemin_destination);
            this.Controls.Add(this.Dossier_destination);
            this.Controls.Add(this.bouton_annee);
            this.Controls.Add(this.bouton_periode3);
            this.Controls.Add(this.bouton_periode2);
            this.Controls.Add(this.bouton_periode1);
            this.Controls.Add(this.Créer_arborescence);
            this.Controls.Add(this.Chemin_dossier);
            this.Controls.Add(this.Dossier_travail);
            this.Controls.Add(this.Annee_scolaire);
            this.Controls.Add(this.Niveau_3);
            this.Controls.Add(this.Niveau_4);
            this.Controls.Add(this.Niveau_5);
            this.Controls.Add(this.Niveau_6);
            this.Controls.Add(this.Liste_CSV);
            this.Controls.Add(this.Lancer_traitement);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Principal";
            this.Text = "Conversion des compétences sur E-Lyco";
            this.Load += new System.EventHandler(this.Principal_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Lancer_traitement;
        private System.Windows.Forms.ListBox Liste_CSV;
        private System.Windows.Forms.ComboBox Niveau_6;
        private System.Windows.Forms.ComboBox Niveau_5;
        private System.Windows.Forms.ComboBox Niveau_4;
        private System.Windows.Forms.ComboBox Niveau_3;
        private System.Windows.Forms.ComboBox Annee_scolaire;
        private System.Windows.Forms.Button Dossier_travail;
        private System.Windows.Forms.Label Chemin_dossier;
        private System.Windows.Forms.RadioButton bouton_periode1;
        private System.Windows.Forms.RadioButton bouton_periode2;
        private System.Windows.Forms.RadioButton bouton_periode3;
        private System.Windows.Forms.RadioButton bouton_annee;
        private System.Windows.Forms.Button Créer_arborescence;
        private System.Windows.Forms.Button Dossier_destination;
        private System.Windows.Forms.Label Chemin_destination;
        private System.Windows.Forms.ListBox Liste_csv_présents;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.Button Supprimer_tout;
        private System.Windows.Forms.Button SuppressionFichierCsv;
        private System.Windows.Forms.ListBox Liste_xlsx_présents;
        private System.Windows.Forms.Button SuppressionFichierXlsx;
        private System.Windows.Forms.Label lbl_fichiers_csv_a_traiter;
        private System.Windows.Forms.Label lbl_fichiers_csv_conservés;
        private System.Windows.Forms.Label lbl_fichiers_xlsx;
    }
}

