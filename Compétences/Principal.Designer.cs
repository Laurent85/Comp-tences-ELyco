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
            this.BtnLancerTraitement = new System.Windows.Forms.Button();
            this.ListBoxCsvATraiter = new System.Windows.Forms.ListBox();
            this.ComboNiveau6 = new System.Windows.Forms.ComboBox();
            this.ComboNiveau5 = new System.Windows.Forms.ComboBox();
            this.ComboNiveau4 = new System.Windows.Forms.ComboBox();
            this.ComboNiveau3 = new System.Windows.Forms.ComboBox();
            this.ComboAnnéeScolaire = new System.Windows.Forms.ComboBox();
            this.BtnDossierCsv = new System.Windows.Forms.Button();
            this.LblCheminDossierCsv = new System.Windows.Forms.Label();
            this.RadioBtnPériode1 = new System.Windows.Forms.RadioButton();
            this.RadioBtnPériode2 = new System.Windows.Forms.RadioButton();
            this.RadioBtnPériode3 = new System.Windows.Forms.RadioButton();
            this.RadioBtnAnnée = new System.Windows.Forms.RadioButton();
            this.BtnCréationArborescence = new System.Windows.Forms.Button();
            this.BtnDossierXlsx = new System.Windows.Forms.Button();
            this.LblCheminDossierXlsx = new System.Windows.Forms.Label();
            this.ListeBoxCsvPrésents = new System.Windows.Forms.ListBox();
            this.LblTitre = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.BtnSuppressionBases = new System.Windows.Forms.Button();
            this.PictureStJacques = new System.Windows.Forms.PictureBox();
            this.PictureELyco = new System.Windows.Forms.PictureBox();
            this.BtnSuppressionFichierCsv = new System.Windows.Forms.Button();
            this.ListeBoxXlsxPrésents = new System.Windows.Forms.ListBox();
            this.BtnSuppressionFichierXlsx = new System.Windows.Forms.Button();
            this.LblFichiersCsvATraiter = new System.Windows.Forms.Label();
            this.LblFichiersCsvPrésents = new System.Windows.Forms.Label();
            this.LblFichiersXlsxPrésents = new System.Windows.Forms.Label();
            this.BtnGénérerPublipostageDnb = new System.Windows.Forms.Button();
            this.BtnGénérerfichiersExcelDnb = new System.Windows.Forms.Button();
            this.BtnSauvegarderBases = new System.Windows.Forms.Button();
            this.BtnRestaurerBases = new System.Windows.Forms.Button();
            this.BtnSuppressionFichierCsvATraiter = new System.Windows.Forms.Button();
            this.PanelTrimestre = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.PictureStJacques)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureELyco)).BeginInit();
            this.PanelTrimestre.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnLancerTraitement
            // 
            this.BtnLancerTraitement.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BtnLancerTraitement.Location = new System.Drawing.Point(440, 559);
            this.BtnLancerTraitement.Name = "BtnLancerTraitement";
            this.BtnLancerTraitement.Size = new System.Drawing.Size(109, 23);
            this.BtnLancerTraitement.TabIndex = 0;
            this.BtnLancerTraitement.Text = "Lancer le traitement";
            this.BtnLancerTraitement.UseVisualStyleBackColor = true;
            this.BtnLancerTraitement.Click += new System.EventHandler(this.LancerTraitementCsv_Click);
            // 
            // ListBoxCsvATraiter
            // 
            this.ListBoxCsvATraiter.AllowDrop = true;
            this.ListBoxCsvATraiter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.ListBoxCsvATraiter.FormattingEnabled = true;
            this.ListBoxCsvATraiter.Location = new System.Drawing.Point(9, 283);
            this.ListBoxCsvATraiter.Name = "ListBoxCsvATraiter";
            this.ListBoxCsvATraiter.Size = new System.Drawing.Size(540, 212);
            this.ListBoxCsvATraiter.TabIndex = 1;
            this.ListBoxCsvATraiter.DragDrop += new System.Windows.Forms.DragEventHandler(this.Drag);
            this.ListBoxCsvATraiter.DragEnter += new System.Windows.Forms.DragEventHandler(this.Drag_Enter);
            // 
            // ComboNiveau6
            // 
            this.ComboNiveau6.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.ComboNiveau6.FormattingEnabled = true;
            this.ComboNiveau6.Items.AddRange(new object[] {
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
            "12"});
            this.ComboNiveau6.Location = new System.Drawing.Point(343, 93);
            this.ComboNiveau6.Name = "ComboNiveau6";
            this.ComboNiveau6.Size = new System.Drawing.Size(99, 21);
            this.ComboNiveau6.TabIndex = 2;
            // 
            // ComboNiveau5
            // 
            this.ComboNiveau5.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.ComboNiveau5.FormattingEnabled = true;
            this.ComboNiveau5.Items.AddRange(new object[] {
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
            "12"});
            this.ComboNiveau5.Location = new System.Drawing.Point(475, 93);
            this.ComboNiveau5.Name = "ComboNiveau5";
            this.ComboNiveau5.Size = new System.Drawing.Size(99, 21);
            this.ComboNiveau5.TabIndex = 3;
            // 
            // ComboNiveau4
            // 
            this.ComboNiveau4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.ComboNiveau4.FormattingEnabled = true;
            this.ComboNiveau4.Items.AddRange(new object[] {
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
            "12"});
            this.ComboNiveau4.Location = new System.Drawing.Point(609, 93);
            this.ComboNiveau4.Name = "ComboNiveau4";
            this.ComboNiveau4.Size = new System.Drawing.Size(99, 21);
            this.ComboNiveau4.TabIndex = 4;
            // 
            // ComboNiveau3
            // 
            this.ComboNiveau3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.ComboNiveau3.FormattingEnabled = true;
            this.ComboNiveau3.Items.AddRange(new object[] {
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
            "12"});
            this.ComboNiveau3.Location = new System.Drawing.Point(728, 93);
            this.ComboNiveau3.Name = "ComboNiveau3";
            this.ComboNiveau3.Size = new System.Drawing.Size(99, 21);
            this.ComboNiveau3.TabIndex = 5;
            // 
            // ComboAnnéeScolaire
            // 
            this.ComboAnnéeScolaire.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.ComboAnnéeScolaire.FormattingEnabled = true;
            this.ComboAnnéeScolaire.Items.AddRange(new object[] {
            "2017-2018",
            "2018-2019",
            "2019-2020",
            "2020-2021",
            "2021-2022",
            "2022-2023",
            "2023-2024",
            "2024-2025"});
            this.ComboAnnéeScolaire.Location = new System.Drawing.Point(155, 93);
            this.ComboAnnéeScolaire.Name = "ComboAnnéeScolaire";
            this.ComboAnnéeScolaire.Size = new System.Drawing.Size(118, 21);
            this.ComboAnnéeScolaire.TabIndex = 6;
            // 
            // BtnDossierCsv
            // 
            this.BtnDossierCsv.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BtnDossierCsv.Location = new System.Drawing.Point(155, 139);
            this.BtnDossierCsv.Name = "BtnDossierCsv";
            this.BtnDossierCsv.Size = new System.Drawing.Size(118, 23);
            this.BtnDossierCsv.TabIndex = 7;
            this.BtnDossierCsv.Text = "Dossier des csv";
            this.BtnDossierCsv.UseVisualStyleBackColor = true;
            this.BtnDossierCsv.Click += new System.EventHandler(this.Dossier_travail_Click);
            // 
            // LblCheminDossierCsv
            // 
            this.LblCheminDossierCsv.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.LblCheminDossierCsv.AutoSize = true;
            this.LblCheminDossierCsv.ForeColor = System.Drawing.Color.ForestGreen;
            this.LblCheminDossierCsv.Location = new System.Drawing.Point(312, 144);
            this.LblCheminDossierCsv.Name = "LblCheminDossierCsv";
            this.LblCheminDossierCsv.Size = new System.Drawing.Size(0, 13);
            this.LblCheminDossierCsv.TabIndex = 8;
            // 
            // RadioBtnPériode1
            // 
            this.RadioBtnPériode1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioBtnPériode1.AutoSize = true;
            this.RadioBtnPériode1.Location = new System.Drawing.Point(5, 1);
            this.RadioBtnPériode1.Name = "RadioBtnPériode1";
            this.RadioBtnPériode1.Size = new System.Drawing.Size(84, 17);
            this.RadioBtnPériode1.TabIndex = 10;
            this.RadioBtnPériode1.TabStop = true;
            this.RadioBtnPériode1.Text = "1ère période";
            this.RadioBtnPériode1.UseVisualStyleBackColor = true;
            this.RadioBtnPériode1.CheckedChanged += new System.EventHandler(this.bouton_periode1_CheckedChanged);
            // 
            // RadioBtnPériode2
            // 
            this.RadioBtnPériode2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioBtnPériode2.AutoSize = true;
            this.RadioBtnPériode2.Location = new System.Drawing.Point(108, 1);
            this.RadioBtnPériode2.Name = "RadioBtnPériode2";
            this.RadioBtnPériode2.Size = new System.Drawing.Size(89, 17);
            this.RadioBtnPériode2.TabIndex = 11;
            this.RadioBtnPériode2.TabStop = true;
            this.RadioBtnPériode2.Text = "2ème période";
            this.RadioBtnPériode2.UseVisualStyleBackColor = true;
            this.RadioBtnPériode2.CheckedChanged += new System.EventHandler(this.bouton_periode2_CheckedChanged);
            // 
            // RadioBtnPériode3
            // 
            this.RadioBtnPériode3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioBtnPériode3.AutoSize = true;
            this.RadioBtnPériode3.Location = new System.Drawing.Point(214, 1);
            this.RadioBtnPériode3.Name = "RadioBtnPériode3";
            this.RadioBtnPériode3.Size = new System.Drawing.Size(89, 17);
            this.RadioBtnPériode3.TabIndex = 12;
            this.RadioBtnPériode3.TabStop = true;
            this.RadioBtnPériode3.Text = "3ème période";
            this.RadioBtnPériode3.UseVisualStyleBackColor = true;
            this.RadioBtnPériode3.CheckedChanged += new System.EventHandler(this.bouton_periode3_CheckedChanged);
            // 
            // RadioBtnAnnée
            // 
            this.RadioBtnAnnée.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RadioBtnAnnée.AutoSize = true;
            this.RadioBtnAnnée.Location = new System.Drawing.Point(333, 1);
            this.RadioBtnAnnée.Name = "RadioBtnAnnée";
            this.RadioBtnAnnée.Size = new System.Drawing.Size(56, 17);
            this.RadioBtnAnnée.TabIndex = 13;
            this.RadioBtnAnnée.TabStop = true;
            this.RadioBtnAnnée.Text = "Année";
            this.RadioBtnAnnée.UseVisualStyleBackColor = true;
            this.RadioBtnAnnée.CheckedChanged += new System.EventHandler(this.bouton_annee_CheckedChanged);
            // 
            // BtnCréationArborescence
            // 
            this.BtnCréationArborescence.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BtnCréationArborescence.Location = new System.Drawing.Point(155, 197);
            this.BtnCréationArborescence.Name = "BtnCréationArborescence";
            this.BtnCréationArborescence.Size = new System.Drawing.Size(118, 23);
            this.BtnCréationArborescence.TabIndex = 9;
            this.BtnCréationArborescence.Text = "Créer l\'arborescence";
            this.BtnCréationArborescence.UseVisualStyleBackColor = true;
            this.BtnCréationArborescence.Click += new System.EventHandler(this.Créer_arborescence_Click);
            // 
            // BtnDossierXlsx
            // 
            this.BtnDossierXlsx.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.BtnDossierXlsx.Location = new System.Drawing.Point(155, 168);
            this.BtnDossierXlsx.Name = "BtnDossierXlsx";
            this.BtnDossierXlsx.Size = new System.Drawing.Size(118, 23);
            this.BtnDossierXlsx.TabIndex = 14;
            this.BtnDossierXlsx.Text = "Dossier des xlsx";
            this.BtnDossierXlsx.UseVisualStyleBackColor = true;
            this.BtnDossierXlsx.Click += new System.EventHandler(this.Dossier_destination_Click);
            // 
            // LblCheminDossierXlsx
            // 
            this.LblCheminDossierXlsx.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.LblCheminDossierXlsx.AutoSize = true;
            this.LblCheminDossierXlsx.ForeColor = System.Drawing.Color.ForestGreen;
            this.LblCheminDossierXlsx.Location = new System.Drawing.Point(312, 173);
            this.LblCheminDossierXlsx.Name = "LblCheminDossierXlsx";
            this.LblCheminDossierXlsx.Size = new System.Drawing.Size(0, 13);
            this.LblCheminDossierXlsx.TabIndex = 15;
            // 
            // ListeBoxCsvPrésents
            // 
            this.ListeBoxCsvPrésents.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ListeBoxCsvPrésents.FormattingEnabled = true;
            this.ListeBoxCsvPrésents.Location = new System.Drawing.Point(613, 283);
            this.ListeBoxCsvPrésents.Name = "ListeBoxCsvPrésents";
            this.ListeBoxCsvPrésents.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.ListeBoxCsvPrésents.Size = new System.Drawing.Size(252, 212);
            this.ListeBoxCsvPrésents.TabIndex = 17;
            this.ListeBoxCsvPrésents.SelectedIndexChanged += new System.EventHandler(this.SelectionClasseTraitementAnnée);
            // 
            // LblTitre
            // 
            this.LblTitre.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.LblTitre.AutoSize = true;
            this.LblTitre.Font = new System.Drawing.Font("Comic Sans MS", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LblTitre.ForeColor = System.Drawing.Color.Blue;
            this.LblTitre.Location = new System.Drawing.Point(336, 12);
            this.LblTitre.Name = "LblTitre";
            this.LblTitre.Size = new System.Drawing.Size(498, 35);
            this.LblTitre.TabIndex = 18;
            this.LblTitre.Text = "Traitement des domaines de compétences";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label2.Location = new System.Drawing.Point(169, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(91, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Année scolaire";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label3.Location = new System.Drawing.Point(340, 74);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(102, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Classes de 6ème";
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label4.Location = new System.Drawing.Point(472, 74);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 13);
            this.label4.TabIndex = 21;
            this.label4.Text = "Classes de 5ème";
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label5.Location = new System.Drawing.Point(606, 74);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(102, 13);
            this.label5.TabIndex = 22;
            this.label5.Text = "Classes de 4ème";
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.Color.RoyalBlue;
            this.label6.Location = new System.Drawing.Point(725, 74);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(102, 13);
            this.label6.TabIndex = 23;
            this.label6.Text = "Classes de 3ème";
            // 
            // BtnSuppressionBases
            // 
            this.BtnSuppressionBases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSuppressionBases.Location = new System.Drawing.Point(1025, 139);
            this.BtnSuppressionBases.Name = "BtnSuppressionBases";
            this.BtnSuppressionBases.Size = new System.Drawing.Size(151, 23);
            this.BtnSuppressionBases.TabIndex = 26;
            this.BtnSuppressionBases.Text = "Supprimer les bases";
            this.BtnSuppressionBases.UseVisualStyleBackColor = true;
            this.BtnSuppressionBases.Click += new System.EventHandler(this.SuppressionBases_Click);
            // 
            // PictureStJacques
            // 
            this.PictureStJacques.Image = global::Compétences.Properties.Resources.St_Jacques;
            this.PictureStJacques.Location = new System.Drawing.Point(14, 9);
            this.PictureStJacques.Name = "PictureStJacques";
            this.PictureStJacques.Size = new System.Drawing.Size(94, 55);
            this.PictureStJacques.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureStJacques.TabIndex = 25;
            this.PictureStJacques.TabStop = false;
            // 
            // PictureELyco
            // 
            this.PictureELyco.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PictureELyco.Image = global::Compétences.Properties.Resources.E_Lyco;
            this.PictureELyco.Location = new System.Drawing.Point(1025, 12);
            this.PictureELyco.Name = "PictureELyco";
            this.PictureELyco.Size = new System.Drawing.Size(151, 56);
            this.PictureELyco.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.PictureELyco.TabIndex = 24;
            this.PictureELyco.TabStop = false;
            // 
            // BtnSuppressionFichierCsv
            // 
            this.BtnSuppressionFichierCsv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSuppressionFichierCsv.Location = new System.Drawing.Point(768, 516);
            this.BtnSuppressionFichierCsv.Name = "BtnSuppressionFichierCsv";
            this.BtnSuppressionFichierCsv.Size = new System.Drawing.Size(97, 23);
            this.BtnSuppressionFichierCsv.TabIndex = 27;
            this.BtnSuppressionFichierCsv.Text = "Supprimer fichier";
            this.BtnSuppressionFichierCsv.UseVisualStyleBackColor = true;
            this.BtnSuppressionFichierCsv.Click += new System.EventHandler(this.SuppressionFichierCsv_Click);
            // 
            // ListeBoxXlsxPrésents
            // 
            this.ListeBoxXlsxPrésents.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ListeBoxXlsxPrésents.FormattingEnabled = true;
            this.ListeBoxXlsxPrésents.Location = new System.Drawing.Point(921, 283);
            this.ListeBoxXlsxPrésents.Name = "ListeBoxXlsxPrésents";
            this.ListeBoxXlsxPrésents.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.ListeBoxXlsxPrésents.Size = new System.Drawing.Size(255, 212);
            this.ListeBoxXlsxPrésents.TabIndex = 28;
            this.ListeBoxXlsxPrésents.SelectedIndexChanged += new System.EventHandler(this.SelectionFichierDnb);
            this.ListeBoxXlsxPrésents.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.OuvrirFichierXlsx_MouseDoubleClick);
            // 
            // BtnSuppressionFichierXlsx
            // 
            this.BtnSuppressionFichierXlsx.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSuppressionFichierXlsx.Location = new System.Drawing.Point(1079, 516);
            this.BtnSuppressionFichierXlsx.Name = "BtnSuppressionFichierXlsx";
            this.BtnSuppressionFichierXlsx.Size = new System.Drawing.Size(97, 23);
            this.BtnSuppressionFichierXlsx.TabIndex = 29;
            this.BtnSuppressionFichierXlsx.Text = "Supprimer fichier";
            this.BtnSuppressionFichierXlsx.UseVisualStyleBackColor = true;
            this.BtnSuppressionFichierXlsx.Click += new System.EventHandler(this.SuppressionFichierXlsx_Click);
            // 
            // LblFichiersCsvATraiter
            // 
            this.LblFichiersCsvATraiter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.LblFichiersCsvATraiter.AutoSize = true;
            this.LblFichiersCsvATraiter.ForeColor = System.Drawing.Color.Red;
            this.LblFichiersCsvATraiter.Location = new System.Drawing.Point(229, 257);
            this.LblFichiersCsvATraiter.Name = "LblFichiersCsvATraiter";
            this.LblFichiersCsvATraiter.Size = new System.Drawing.Size(0, 13);
            this.LblFichiersCsvATraiter.TabIndex = 30;
            // 
            // LblFichiersCsvPrésents
            // 
            this.LblFichiersCsvPrésents.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LblFichiersCsvPrésents.AutoSize = true;
            this.LblFichiersCsvPrésents.ForeColor = System.Drawing.Color.Red;
            this.LblFichiersCsvPrésents.Location = new System.Drawing.Point(675, 257);
            this.LblFichiersCsvPrésents.Name = "LblFichiersCsvPrésents";
            this.LblFichiersCsvPrésents.Size = new System.Drawing.Size(0, 13);
            this.LblFichiersCsvPrésents.TabIndex = 31;
            // 
            // LblFichiersXlsxPrésents
            // 
            this.LblFichiersXlsxPrésents.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LblFichiersXlsxPrésents.AutoSize = true;
            this.LblFichiersXlsxPrésents.ForeColor = System.Drawing.Color.Red;
            this.LblFichiersXlsxPrésents.Location = new System.Drawing.Point(961, 257);
            this.LblFichiersXlsxPrésents.Name = "LblFichiersXlsxPrésents";
            this.LblFichiersXlsxPrésents.Size = new System.Drawing.Size(0, 13);
            this.LblFichiersXlsxPrésents.TabIndex = 32;
            // 
            // BtnGénérerPublipostageDnb
            // 
            this.BtnGénérerPublipostageDnb.Location = new System.Drawing.Point(921, 559);
            this.BtnGénérerPublipostageDnb.Name = "BtnGénérerPublipostageDnb";
            this.BtnGénérerPublipostageDnb.Size = new System.Drawing.Size(151, 23);
            this.BtnGénérerPublipostageDnb.TabIndex = 35;
            this.BtnGénérerPublipostageDnb.Text = "Générer publipostage DNB";
            this.BtnGénérerPublipostageDnb.UseVisualStyleBackColor = true;
            this.BtnGénérerPublipostageDnb.Click += new System.EventHandler(this.LancerTraitementDnb_Click);
            // 
            // BtnGénérerfichiersExcelDnb
            // 
            this.BtnGénérerfichiersExcelDnb.Location = new System.Drawing.Point(921, 516);
            this.BtnGénérerfichiersExcelDnb.Name = "BtnGénérerfichiersExcelDnb";
            this.BtnGénérerfichiersExcelDnb.Size = new System.Drawing.Size(151, 23);
            this.BtnGénérerfichiersExcelDnb.TabIndex = 36;
            this.BtnGénérerfichiersExcelDnb.Text = "Générer fichiers Excel DNB";
            this.BtnGénérerfichiersExcelDnb.UseVisualStyleBackColor = true;
            this.BtnGénérerfichiersExcelDnb.Click += new System.EventHandler(this.BtnGénérerfichiersExcelDnb_Click);
            // 
            // BtnSauvegarderBases
            // 
            this.BtnSauvegarderBases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSauvegarderBases.Location = new System.Drawing.Point(1025, 168);
            this.BtnSauvegarderBases.Name = "BtnSauvegarderBases";
            this.BtnSauvegarderBases.Size = new System.Drawing.Size(151, 23);
            this.BtnSauvegarderBases.TabIndex = 37;
            this.BtnSauvegarderBases.Text = "Sauvegarder les bases";
            this.BtnSauvegarderBases.UseVisualStyleBackColor = true;
            // 
            // BtnRestaurerBases
            // 
            this.BtnRestaurerBases.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnRestaurerBases.Location = new System.Drawing.Point(1025, 197);
            this.BtnRestaurerBases.Name = "BtnRestaurerBases";
            this.BtnRestaurerBases.Size = new System.Drawing.Size(151, 23);
            this.BtnRestaurerBases.TabIndex = 38;
            this.BtnRestaurerBases.Text = "Restaurer les bases";
            this.BtnRestaurerBases.UseVisualStyleBackColor = true;
            // 
            // BtnSuppressionFichierCsvATraiter
            // 
            this.BtnSuppressionFichierCsvATraiter.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSuppressionFichierCsvATraiter.Location = new System.Drawing.Point(440, 516);
            this.BtnSuppressionFichierCsvATraiter.Name = "BtnSuppressionFichierCsvATraiter";
            this.BtnSuppressionFichierCsvATraiter.Size = new System.Drawing.Size(109, 23);
            this.BtnSuppressionFichierCsvATraiter.TabIndex = 39;
            this.BtnSuppressionFichierCsvATraiter.Text = "Supprimer fichier";
            this.BtnSuppressionFichierCsvATraiter.UseVisualStyleBackColor = true;
            this.BtnSuppressionFichierCsvATraiter.Click += new System.EventHandler(this.BtnSuppressionFichierCsvATraiter_Click);
            // 
            // PanelTrimestre
            // 
            this.PanelTrimestre.Controls.Add(this.RadioBtnPériode1);
            this.PanelTrimestre.Controls.Add(this.RadioBtnPériode2);
            this.PanelTrimestre.Controls.Add(this.RadioBtnPériode3);
            this.PanelTrimestre.Controls.Add(this.RadioBtnAnnée);
            this.PanelTrimestre.Location = new System.Drawing.Point(9, 559);
            this.PanelTrimestre.Name = "PanelTrimestre";
            this.PanelTrimestre.Size = new System.Drawing.Size(417, 23);
            this.PanelTrimestre.TabIndex = 40;
            // 
            // Principal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.PaleGoldenrod;
            this.ClientSize = new System.Drawing.Size(1188, 594);
            this.Controls.Add(this.PanelTrimestre);
            this.Controls.Add(this.BtnSuppressionFichierCsvATraiter);
            this.Controls.Add(this.BtnRestaurerBases);
            this.Controls.Add(this.BtnSauvegarderBases);
            this.Controls.Add(this.BtnGénérerfichiersExcelDnb);
            this.Controls.Add(this.BtnGénérerPublipostageDnb);
            this.Controls.Add(this.LblFichiersXlsxPrésents);
            this.Controls.Add(this.LblFichiersCsvPrésents);
            this.Controls.Add(this.LblFichiersCsvATraiter);
            this.Controls.Add(this.BtnSuppressionFichierXlsx);
            this.Controls.Add(this.ListeBoxXlsxPrésents);
            this.Controls.Add(this.BtnSuppressionFichierCsv);
            this.Controls.Add(this.BtnSuppressionBases);
            this.Controls.Add(this.PictureStJacques);
            this.Controls.Add(this.PictureELyco);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.LblTitre);
            this.Controls.Add(this.ListeBoxCsvPrésents);
            this.Controls.Add(this.LblCheminDossierXlsx);
            this.Controls.Add(this.BtnDossierXlsx);
            this.Controls.Add(this.BtnCréationArborescence);
            this.Controls.Add(this.LblCheminDossierCsv);
            this.Controls.Add(this.BtnDossierCsv);
            this.Controls.Add(this.ComboAnnéeScolaire);
            this.Controls.Add(this.ComboNiveau3);
            this.Controls.Add(this.ComboNiveau4);
            this.Controls.Add(this.ComboNiveau5);
            this.Controls.Add(this.ComboNiveau6);
            this.Controls.Add(this.ListBoxCsvATraiter);
            this.Controls.Add(this.BtnLancerTraitement);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Principal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Conversion des compétences sur E-Lyco";
            this.Load += new System.EventHandler(this.Principal_Load);
            ((System.ComponentModel.ISupportInitialize)(this.PictureStJacques)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.PictureELyco)).EndInit();
            this.PanelTrimestre.ResumeLayout(false);
            this.PanelTrimestre.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button BtnLancerTraitement;
        private System.Windows.Forms.ListBox ListBoxCsvATraiter;
        private System.Windows.Forms.ComboBox ComboNiveau6;
        private System.Windows.Forms.ComboBox ComboNiveau5;
        private System.Windows.Forms.ComboBox ComboNiveau4;
        private System.Windows.Forms.ComboBox ComboNiveau3;
        private System.Windows.Forms.ComboBox ComboAnnéeScolaire;
        private System.Windows.Forms.Button BtnDossierCsv;
        private System.Windows.Forms.Label LblCheminDossierCsv;
        private System.Windows.Forms.RadioButton RadioBtnPériode1;
        private System.Windows.Forms.RadioButton RadioBtnPériode2;
        private System.Windows.Forms.RadioButton RadioBtnPériode3;
        private System.Windows.Forms.RadioButton RadioBtnAnnée;
        private System.Windows.Forms.Button BtnCréationArborescence;
        private System.Windows.Forms.Button BtnDossierXlsx;
        private System.Windows.Forms.Label LblCheminDossierXlsx;
        private System.Windows.Forms.ListBox ListeBoxCsvPrésents;
        private System.Windows.Forms.Label LblTitre;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.PictureBox PictureELyco;
        private System.Windows.Forms.PictureBox PictureStJacques;
        private System.Windows.Forms.Button BtnSuppressionBases;
        private System.Windows.Forms.Button BtnSuppressionFichierCsv;
        private System.Windows.Forms.ListBox ListeBoxXlsxPrésents;
        private System.Windows.Forms.Button BtnSuppressionFichierXlsx;
        private System.Windows.Forms.Label LblFichiersCsvATraiter;
        private System.Windows.Forms.Label LblFichiersCsvPrésents;
        private System.Windows.Forms.Label LblFichiersXlsxPrésents;
        private System.Windows.Forms.Button BtnGénérerPublipostageDnb;
        private System.Windows.Forms.Button BtnGénérerfichiersExcelDnb;
        private System.Windows.Forms.Button BtnSauvegarderBases;
        private System.Windows.Forms.Button BtnRestaurerBases;
        private System.Windows.Forms.Button BtnSuppressionFichierCsvATraiter;
        private System.Windows.Forms.Panel PanelTrimestre;
    }
}

