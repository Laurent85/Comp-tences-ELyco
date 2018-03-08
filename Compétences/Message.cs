using System;
using System.Drawing;
using System.Windows.Forms;

namespace Compétences
{
    public partial class Message : Form
    {
        public string message = "";

        public Message()

        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            label1.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void btn_fermer_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}