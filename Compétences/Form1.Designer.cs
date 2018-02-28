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
            this.Créer_arborescence_destination = new System.Windows.Forms.Button();
            this.Liste_csv_présents = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // Lancer_traitement
            // 
            this.Lancer_traitement.Location = new System.Drawing.Point(561, 485);
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
            this.Liste_CSV.Location = new System.Drawing.Point(12, 296);
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
            this.Niveau_6.Location = new System.Drawing.Point(180, 33);
            this.Niveau_6.Name = "Niveau_6";
            this.Niveau_6.Size = new System.Drawing.Size(121, 21);
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
            this.Niveau_5.Location = new System.Drawing.Point(307, 33);
            this.Niveau_5.Name = "Niveau_5";
            this.Niveau_5.Size = new System.Drawing.Size(121, 21);
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
            this.Niveau_4.Location = new System.Drawing.Point(434, 33);
            this.Niveau_4.Name = "Niveau_4";
            this.Niveau_4.Size = new System.Drawing.Size(121, 21);
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
            this.Niveau_3.Location = new System.Drawing.Point(561, 33);
            this.Niveau_3.Name = "Niveau_3";
            this.Niveau_3.Size = new System.Drawing.Size(121, 21);
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
            this.Annee_scolaire.Location = new System.Drawing.Point(35, 33);
            this.Annee_scolaire.Name = "Annee_scolaire";
            this.Annee_scolaire.Size = new System.Drawing.Size(121, 21);
            this.Annee_scolaire.TabIndex = 6;
            // 
            // Dossier_travail
            // 
            this.Dossier_travail.Location = new System.Drawing.Point(35, 69);
            this.Dossier_travail.Name = "Dossier_travail";
            this.Dossier_travail.Size = new System.Drawing.Size(75, 23);
            this.Dossier_travail.TabIndex = 7;
            this.Dossier_travail.Text = "Parcourir...";
            this.Dossier_travail.UseVisualStyleBackColor = true;
            this.Dossier_travail.Click += new System.EventHandler(this.Dossier_travail_Click);
            // 
            // Chemin_dossier
            // 
            this.Chemin_dossier.AutoSize = true;
            this.Chemin_dossier.Location = new System.Drawing.Point(156, 74);
            this.Chemin_dossier.Name = "Chemin_dossier";
            this.Chemin_dossier.Size = new System.Drawing.Size(0, 13);
            this.Chemin_dossier.TabIndex = 8;
            // 
            // bouton_periode1
            // 
            this.bouton_periode1.AutoSize = true;
            this.bouton_periode1.Location = new System.Drawing.Point(130, 260);
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
            this.bouton_periode2.Location = new System.Drawing.Point(232, 260);
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
            this.bouton_periode3.Location = new System.Drawing.Point(339, 260);
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
            this.bouton_annee.Location = new System.Drawing.Point(458, 260);
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
            this.Créer_arborescence.Location = new System.Drawing.Point(35, 98);
            this.Créer_arborescence.Name = "Créer_arborescence";
            this.Créer_arborescence.Size = new System.Drawing.Size(118, 23);
            this.Créer_arborescence.TabIndex = 9;
            this.Créer_arborescence.Text = "Créer l\'arborescence";
            this.Créer_arborescence.UseVisualStyleBackColor = true;
            this.Créer_arborescence.Click += new System.EventHandler(this.Créer_arborescence_Click);
            // 
            // Dossier_destination
            // 
            this.Dossier_destination.Location = new System.Drawing.Point(35, 152);
            this.Dossier_destination.Name = "Dossier_destination";
            this.Dossier_destination.Size = new System.Drawing.Size(75, 23);
            this.Dossier_destination.TabIndex = 14;
            this.Dossier_destination.Text = "Parcourir...";
            this.Dossier_destination.UseVisualStyleBackColor = true;
            this.Dossier_destination.Click += new System.EventHandler(this.Dossier_destination_Click);
            // 
            // Chemin_destination
            // 
            this.Chemin_destination.AutoSize = true;
            this.Chemin_destination.Location = new System.Drawing.Point(153, 157);
            this.Chemin_destination.Name = "Chemin_destination";
            this.Chemin_destination.Size = new System.Drawing.Size(0, 13);
            this.Chemin_destination.TabIndex = 15;
            // 
            // Créer_arborescence_destination
            // 
            this.Créer_arborescence_destination.Location = new System.Drawing.Point(35, 181);
            this.Créer_arborescence_destination.Name = "Créer_arborescence_destination";
            this.Créer_arborescence_destination.Size = new System.Drawing.Size(121, 23);
            this.Créer_arborescence_destination.TabIndex = 16;
            this.Créer_arborescence_destination.Text = "Créer l\'arborescence";
            this.Créer_arborescence_destination.UseVisualStyleBackColor = true;
            this.Créer_arborescence_destination.Click += new System.EventHandler(this.Créer_arborescence_destination_Click);
            // 
            // Liste_csv_présents
            // 
            this.Liste_csv_présents.FormattingEnabled = true;
            this.Liste_csv_présents.Location = new System.Drawing.Point(807, 296);
            this.Liste_csv_présents.Name = "Liste_csv_présents";
            this.Liste_csv_présents.Size = new System.Drawing.Size(270, 212);
            this.Liste_csv_présents.TabIndex = 17;
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1099, 520);
            this.Controls.Add(this.Liste_csv_présents);
            this.Controls.Add(this.Créer_arborescence_destination);
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
            this.Name = "Principal";
            this.Text = "Conversion des compétences sur E-Lyco";
            this.Load += new System.EventHandler(this.Principal_Load);
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
        private System.Windows.Forms.Button Créer_arborescence_destination;
        private System.Windows.Forms.ListBox Liste_csv_présents;
    }
}

