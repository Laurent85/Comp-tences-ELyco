using System;
using System.Drawing;
using System.Windows.Forms;

namespace Compétences
{
    public partial class Message : Form
    {
        public Message()

        {
            InitializeComponent();
        }

        private void Message_Load(object sender, EventArgs e)
        {
            LblMessageTraitement.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void BtnFermer_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}