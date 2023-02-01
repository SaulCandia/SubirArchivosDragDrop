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
    public partial class FormEscoger : Form
    {
        string seleccion = "";
        public FormEscoger()
        {
            InitializeComponent();
        }

        private void btnSalir_Click_1(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("¿Desea finalizar la aplicación?", "Confirmar acción", MessageBoxButtons.YesNo,
           MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void btnEnter_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("¿Desea finalizar la aplicación?", "Confirmar acción", MessageBoxButtons.YesNo,
           MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void btn_Interfaz_consumos_general_Click(object sender, EventArgs e)
        {
            seleccion = "Consumos (Reporte General)";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_Produccion_Click(object sender, EventArgs e)
        {
            seleccion = "Producción (Reporte Consumos Específicos)";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_PlanmensualxMaquina_Click(object sender, EventArgs e)
        {
            seleccion = "Plan Mensual x Máquina";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_ProductividadPotencial_Click(object sender, EventArgs e)
        {
            seleccion = "Productividad Potencial";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_TurnosProgramados_Click(object sender, EventArgs e)
        {
            seleccion = "Turnos Programados Prácticos";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_TurnosReales_Click(object sender, EventArgs e)
        {
            seleccion = "Turnos Reales";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }

        private void btn_Interfaz_TiemposMuertos_Click(object sender, EventArgs e)
        {
            seleccion = "Tiempos Muertos";
            Main main = new Main();
            main.Seleccion = seleccion;
            main.Show();
            this.Hide();
        }
    }
}
