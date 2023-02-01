using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Subir2
{
    public partial class Form0 : Form
    {
        public Form0()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form0_Load(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Increment(50);
            if(progressBar1.Value == progressBar1.Maximum)
            {
                timer1.Stop();
                this.Hide();
                FormEscoger ventana = new FormEscoger();
                //Main ventana = new Main();
                ventana.ShowDialog();
            }
            
            
        }
    }
}
