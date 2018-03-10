namespace Compétences
{
    partial class Message
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.LblMessageTraitement = new System.Windows.Forms.Label();
            this.BtnFermerMessageTraitement = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // LblMessageTraitement
            // 
            this.LblMessageTraitement.AutoSize = true;
            this.LblMessageTraitement.CausesValidation = false;
            this.LblMessageTraitement.Location = new System.Drawing.Point(17, 27);
            this.LblMessageTraitement.Name = "LblMessageTraitement";
            this.LblMessageTraitement.Size = new System.Drawing.Size(255, 13);
            this.LblMessageTraitement.TabIndex = 0;
            this.LblMessageTraitement.Text = "Traitement des fichiers en cours...Veuillez patienter...";
            this.LblMessageTraitement.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // BtnFermerMessageTraitement
            // 
            this.BtnFermerMessageTraitement.ImageAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.BtnFermerMessageTraitement.Location = new System.Drawing.Point(105, 55);
            this.BtnFermerMessageTraitement.Name = "BtnFermerMessageTraitement";
            this.BtnFermerMessageTraitement.Size = new System.Drawing.Size(75, 23);
            this.BtnFermerMessageTraitement.TabIndex = 1;
            this.BtnFermerMessageTraitement.Text = "Fermer";
            this.BtnFermerMessageTraitement.UseVisualStyleBackColor = true;
            this.BtnFermerMessageTraitement.Click += new System.EventHandler(this.BtnFermer_Click);
            // 
            // Message
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 90);
            this.ControlBox = false;
            this.Controls.Add(this.BtnFermerMessageTraitement);
            this.Controls.Add(this.LblMessageTraitement);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Message";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Information";
            this.Load += new System.EventHandler(this.Message_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label LblMessageTraitement;
        private System.Windows.Forms.Button BtnFermerMessageTraitement;
    }
}