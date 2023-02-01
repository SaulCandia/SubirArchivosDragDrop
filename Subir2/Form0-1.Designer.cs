
namespace Subir2
{
    partial class FormEscoger
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormEscoger));
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.btnEnter = new System.Windows.Forms.Button();
            this.btnSalir = new System.Windows.Forms.Button();
            this.btn_Interfaz_consumos_general = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btn_Interfaz_ProductividadPotencial = new System.Windows.Forms.Button();
            this.btn_Interfaz_TurnosProgramados = new System.Windows.Forms.Button();
            this.btn_Interfaz_TurnosReales = new System.Windows.Forms.Button();
            this.btn_Interfaz_PlanmensualxMaquina = new System.Windows.Forms.Button();
            this.btn_Interfaz_Produccion = new System.Windows.Forms.Button();
            this.btn_Interfaz_TiemposMuertos = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(69, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(184, 25);
            this.label1.TabIndex = 8;
            this.label1.Text = "Escoja ítems de datos";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.InitialImage = null;
            this.pictureBox1.Location = new System.Drawing.Point(12, 12);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(51, 50);
            this.pictureBox1.TabIndex = 12;
            this.pictureBox1.TabStop = false;
            // 
            // btnEnter
            // 
            this.btnEnter.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEnter.Image = ((System.Drawing.Image)(resources.GetObject("btnEnter.Image")));
            this.btnEnter.Location = new System.Drawing.Point(656, 428);
            this.btnEnter.Name = "btnEnter";
            this.btnEnter.Size = new System.Drawing.Size(72, 32);
            this.btnEnter.TabIndex = 8;
            this.btnEnter.Text = "Salir";
            this.btnEnter.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEnter.UseVisualStyleBackColor = true;
            this.btnEnter.Click += new System.EventHandler(this.btnEnter_Click);
            // 
            // btnSalir
            // 
            this.btnSalir.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
            this.btnSalir.Location = new System.Drawing.Point(695, 12);
            this.btnSalir.Name = "btnSalir";
            this.btnSalir.Size = new System.Drawing.Size(33, 32);
            this.btnSalir.TabIndex = 9;
            this.btnSalir.UseVisualStyleBackColor = true;
            this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click_1);
            // 
            // btn_Interfaz_consumos_general
            // 
            this.btn_Interfaz_consumos_general.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_consumos_general.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_consumos_general.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_consumos_general.Location = new System.Drawing.Point(3, 3);
            this.btn_Interfaz_consumos_general.Name = "btn_Interfaz_consumos_general";
            this.btn_Interfaz_consumos_general.Size = new System.Drawing.Size(234, 98);
            this.btn_Interfaz_consumos_general.TabIndex = 1;
            this.btn_Interfaz_consumos_general.Text = "Consumos \r\n(Reporte General - SIGP)";
            this.btn_Interfaz_consumos_general.UseVisualStyleBackColor = false;
            this.btn_Interfaz_consumos_general.Click += new System.EventHandler(this.btn_Interfaz_consumos_general_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_ProductividadPotencial, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_TurnosProgramados, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_TurnosReales, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_PlanmensualxMaquina, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_Produccion, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_consumos_general, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btn_Interfaz_TiemposMuertos, 1, 2);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 92);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 111F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(716, 330);
            this.tableLayoutPanel1.TabIndex = 14;
            // 
            // btn_Interfaz_ProductividadPotencial
            // 
            this.btn_Interfaz_ProductividadPotencial.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_ProductividadPotencial.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_ProductividadPotencial.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_ProductividadPotencial.Location = new System.Drawing.Point(3, 112);
            this.btn_Interfaz_ProductividadPotencial.Name = "btn_Interfaz_ProductividadPotencial";
            this.btn_Interfaz_ProductividadPotencial.Size = new System.Drawing.Size(234, 98);
            this.btn_Interfaz_ProductividadPotencial.TabIndex = 4;
            this.btn_Interfaz_ProductividadPotencial.Text = "Productividad Potencial";
            this.btn_Interfaz_ProductividadPotencial.UseVisualStyleBackColor = false;
            this.btn_Interfaz_ProductividadPotencial.Click += new System.EventHandler(this.btn_Interfaz_ProductividadPotencial_Click);
            // 
            // btn_Interfaz_TurnosProgramados
            // 
            this.btn_Interfaz_TurnosProgramados.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_TurnosProgramados.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_TurnosProgramados.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_TurnosProgramados.Location = new System.Drawing.Point(243, 112);
            this.btn_Interfaz_TurnosProgramados.Name = "btn_Interfaz_TurnosProgramados";
            this.btn_Interfaz_TurnosProgramados.Size = new System.Drawing.Size(234, 97);
            this.btn_Interfaz_TurnosProgramados.TabIndex = 5;
            this.btn_Interfaz_TurnosProgramados.Text = "Turnos Programados";
            this.btn_Interfaz_TurnosProgramados.UseVisualStyleBackColor = false;
            this.btn_Interfaz_TurnosProgramados.Click += new System.EventHandler(this.btn_Interfaz_TurnosProgramados_Click);
            // 
            // btn_Interfaz_TurnosReales
            // 
            this.btn_Interfaz_TurnosReales.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_TurnosReales.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_TurnosReales.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_TurnosReales.Location = new System.Drawing.Point(483, 112);
            this.btn_Interfaz_TurnosReales.Name = "btn_Interfaz_TurnosReales";
            this.btn_Interfaz_TurnosReales.Size = new System.Drawing.Size(230, 98);
            this.btn_Interfaz_TurnosReales.TabIndex = 6;
            this.btn_Interfaz_TurnosReales.Text = "Turnos Reales";
            this.btn_Interfaz_TurnosReales.UseVisualStyleBackColor = false;
            this.btn_Interfaz_TurnosReales.Click += new System.EventHandler(this.btn_Interfaz_TurnosReales_Click);
            // 
            // btn_Interfaz_PlanmensualxMaquina
            // 
            this.btn_Interfaz_PlanmensualxMaquina.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_PlanmensualxMaquina.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_PlanmensualxMaquina.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_PlanmensualxMaquina.Location = new System.Drawing.Point(483, 3);
            this.btn_Interfaz_PlanmensualxMaquina.Name = "btn_Interfaz_PlanmensualxMaquina";
            this.btn_Interfaz_PlanmensualxMaquina.Size = new System.Drawing.Size(230, 98);
            this.btn_Interfaz_PlanmensualxMaquina.TabIndex = 3;
            this.btn_Interfaz_PlanmensualxMaquina.Text = "Plan Mensual x Máquina";
            this.btn_Interfaz_PlanmensualxMaquina.UseVisualStyleBackColor = false;
            this.btn_Interfaz_PlanmensualxMaquina.Click += new System.EventHandler(this.btn_Interfaz_PlanmensualxMaquina_Click);
            // 
            // btn_Interfaz_Produccion
            // 
            this.btn_Interfaz_Produccion.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_Produccion.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_Produccion.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_Produccion.Location = new System.Drawing.Point(243, 3);
            this.btn_Interfaz_Produccion.Name = "btn_Interfaz_Produccion";
            this.btn_Interfaz_Produccion.Size = new System.Drawing.Size(234, 98);
            this.btn_Interfaz_Produccion.TabIndex = 2;
            this.btn_Interfaz_Produccion.Text = "Producción \r\n(Consumos Específicos)";
            this.btn_Interfaz_Produccion.UseVisualStyleBackColor = false;
            this.btn_Interfaz_Produccion.Click += new System.EventHandler(this.btn_Interfaz_Produccion_Click);
            // 
            // btn_Interfaz_TiemposMuertos
            // 
            this.btn_Interfaz_TiemposMuertos.BackColor = System.Drawing.Color.Green;
            this.btn_Interfaz_TiemposMuertos.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Interfaz_TiemposMuertos.ForeColor = System.Drawing.Color.White;
            this.btn_Interfaz_TiemposMuertos.Location = new System.Drawing.Point(243, 221);
            this.btn_Interfaz_TiemposMuertos.Name = "btn_Interfaz_TiemposMuertos";
            this.btn_Interfaz_TiemposMuertos.Size = new System.Drawing.Size(234, 98);
            this.btn_Interfaz_TiemposMuertos.TabIndex = 7;
            this.btn_Interfaz_TiemposMuertos.Text = "Tiempos Muertos";
            this.btn_Interfaz_TiemposMuertos.UseVisualStyleBackColor = false;
            this.btn_Interfaz_TiemposMuertos.Click += new System.EventHandler(this.btn_Interfaz_TiemposMuertos_Click);
            // 
            // FormEscoger
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(742, 466);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.btnEnter);
            this.Controls.Add(this.btnSalir);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormEscoger";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Escoja su área";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnEnter;
        private System.Windows.Forms.Button btnSalir;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btn_Interfaz_consumos_general;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btn_Interfaz_ProductividadPotencial;
        private System.Windows.Forms.Button btn_Interfaz_TurnosProgramados;
        private System.Windows.Forms.Button btn_Interfaz_TurnosReales;
        private System.Windows.Forms.Button btn_Interfaz_PlanmensualxMaquina;
        private System.Windows.Forms.Button btn_Interfaz_Produccion;
        private System.Windows.Forms.Button btn_Interfaz_TiemposMuertos;
    }
}