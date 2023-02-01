using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Threading;
using System.Configuration;
using Microsoft.Office.Interop.Excel;

namespace Subir2
{
    public partial class Main : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=10.2.0.148\DEVSTAR;Initial Catalog=Cargas;Persist Security Info=True;User ID=sa;Password=L30$2Kv.Tv112.c");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        //ID variable used in Updating and Deleting Record  
        int ID = 0;
        string accionGuardadoPP;
        public string Seleccion { get; set; }
        public Main()
        {
            InitializeComponent();
        }

        //Botón Cargar
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFilename.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFilename.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreview))
                {
                    loadPreview.ShowDialog(this);
                }
            }

        }

        //Sube la data en un solo ciclo
        void UploadData()
        {
            for (int i = 0; i < 1; i++)
            {
                subiendoData();
                Thread.Sleep(2);
            }
        }



        //Vista previa de archivo Excel de SIGP antes de cargarlo
        void ExcelPreview()
        {
            try
            {
                string conexionSIGP = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilename.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionSIGP);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionSIGP);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [PROD$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                grillaExcel.Invoke((MethodInvoker)(() => grillaExcel.DataSource = ds.Tables[0]));
                btnCargar.Invoke((MethodInvoker)(() => btnCargar.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos SIGP." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        //Muestra cuadro de diálogo de carga
        private void btnCargar_Click_1(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de Producción (SIGP) en el servidor." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Carga archivo Excel SIGP
        void subiendoData()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilename.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [PROD$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                grillaExcel.Invoke((MethodInvoker)(() => grillaExcel.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "ProduccionUploadRaw";
                importar.WriteToServer(ds.Tables[0]);

                conexion_destino.Close();

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                txtFilename.Invoke((MethodInvoker)(() => txtFilename.Text = ""));
                btnCargar.Invoke((MethodInvoker)(() => btnCargar.Visible = false));
                int rowcount = grillaExcel.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Carga de pantalla inicial (Main)
        private void Main_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'turnosPracticosDataSet.TurnosPracticosUploadRaw' table. You can move, or remove it, as needed.
            this.turnosPracticosUploadRawTableAdapter.Fill(this.turnosPracticosDataSet.TurnosPracticosUploadRaw);
            // TODO: This line of code loads data into the 'turnosRealesDataSet.TurnosRealesUploadRaw' table. You can move, or remove it, as needed.
            this.turnosRealesUploadRawTableAdapter.Fill(this.turnosRealesDataSet.TurnosRealesUploadRaw);
            // Preparando interfaz estética
            tabPage.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabPage.SizeMode = TabSizeMode.Fixed;

            Size tab_size = tabPage.ItemSize;
            tab_size.Width = 168;
            tab_size.Height += 6;
            tabPage.ItemSize = tab_size;

            lblInterfaz.Text = Seleccion;

            //Carga grillas con actualización en vivo

            switch (Seleccion)
            {
                case "Turnos Programados Prácticos":

                    populateDataGridViewTurnosPracticos();

                    break;

                case "Turnos Reales":

                    populateDataGridViewTurnosReales();
                    break;
            }

            populateDataGridViewProductividadPotencial();

            //Prepara GUI según selección de interfaz/reporte a cargar
            switch (Seleccion)
            {
                case "Tiempos Muertos":
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;

                case "Producción (Reporte Consumos Específicos)":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;

                case "Consumos (Reporte General)":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;

                case "Turnos Programados Prácticos":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;

                case "Turnos Reales":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;

                case "Plan Mensual x Máquina":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabCargaPP);
                    tabPage.TabPages.Remove(tabHistorialPP);
                    break;

                case "Productividad Potencial":
                    tabPage.TabPages.Remove(tabCargarTTMM);
                    tabPage.TabPages.Remove(tabHistorialTTMM);
                    tabPage.TabPages.Remove(tabCargarProduccion);
                    tabPage.TabPages.Remove(tabHistorialProduccion);
                    tabPage.TabPages.Remove(tabCargarConsumos);
                    tabPage.TabPages.Remove(tabHistorialConsumos);
                    tabPage.TabPages.Remove(tabCargarTurnosPracticos);
                    tabPage.TabPages.Remove(tabHistorialTurnosPracticos);
                    tabPage.TabPages.Remove(tabCargarTurnosReales);
                    tabPage.TabPages.Remove(tabHistorialTurnosReales);
                    tabPage.TabPages.Remove(tabPlanMensualxMaquina);
                    tabPage.TabPages.Remove(tabCargaPlanMensualxMaquina);
                    break;
            }
        }

        private const int tab_margin = 3;

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            Brush txt_brush, box_brush;
            Pen box_pen;
            System.Drawing.Rectangle tab_rect = tabPage.GetTabRect(e.Index);

            //Fondo pestaña si está seleccionada
            if (e.State == DrawItemState.Selected)
            {
                e.Graphics.FillRectangle(Brushes.DarkGreen, tab_rect);
                e.DrawFocusRectangle();

                txt_brush = Brushes.White;
                box_brush = Brushes.Silver;
                box_pen = Pens.DarkBlue;
            }
            else
            {
                e.Graphics.FillRectangle(Brushes.LightGray, tab_rect);

                txt_brush = Brushes.Black;
                box_brush = Brushes.LightGray;
                box_pen = Pens.DarkBlue;
            }

            // Márgenes
            RectangleF layout_rect = new RectangleF(
                tab_rect.Left + tab_margin,
                tab_rect.Y + tab_margin,
                tab_rect.Width - 2 * tab_margin,
                tab_rect.Height - 2 * tab_margin);
            using (StringFormat string_format = new StringFormat())
            {

                // Draw the tab's text centered.
                using (System.Drawing.Font big_font = new System.Drawing.Font(this.Font, FontStyle.Regular))
                {
                    string_format.Alignment = StringAlignment.Center;
                    string_format.LineAlignment = StringAlignment.Center;
                    e.Graphics.DrawString(
                        tabPage.TabPages[e.Index].Text,
                        big_font,
                        txt_brush,
                        layout_rect,
                        string_format);
                }
            }
        }

        private void btnBuscarSIGP_Click(object sender, EventArgs e)
        {
            if (dateSIGPHasta.Value < dateSIGPDesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataSIGP))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        //Busca datos en la tabla
        void SearchDataSIGP()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }


                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnBuscarSIGP.Invoke((MethodInvoker)(() => btnBuscarSIGP.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Botón de borrado de datos de tablas
        private void btnDeleteRecordsSIGP_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataSIGP))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado");
                }
            }
        }

        //Elimina datos en la tabla
        void DeleteDataSIGP()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtSIGP = new System.Data.DataTable("SIGP"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_SIGP @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dateSIGPDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dateSIGPHasta.Value);
                            if (chkIncluirAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtSIGP);
                            dtHistorialSIGP.Invoke((MethodInvoker)(() => dtHistorialSIGP.DataSource = dtSIGP));
                            btnDeleteRecordsSIGP.Invoke((MethodInvoker)(() => btnDeleteRecordsSIGP.Enabled = false));
                            lblTotalSIGP.Invoke((MethodInvoker)(() => lblTotalSIGP.Text = "Registros encontrados: 0"));
                            SaveLog("Borrado");
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Buscar Archivo TTMM
        private void btnTTMMbrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFilename.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtArchivoTTMM.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewTTMM))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }
        //Vista previa Excel TTMM
        void ExcelPreviewTTMM()
        {
            try
            {

                string conexionTTMM = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtArchivoTTMM.Text + "';Extended Properties=Excel 12.0;";
                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionTTMM);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();
                adaptador.Fill(ds);
                grillaTTMM.Invoke((MethodInvoker)(() => grillaTTMM.DataSource = ds.Tables[0]));
                btnUploadTtmm.Invoke((MethodInvoker)(() => btnUploadTtmm.Visible = true));
                origen1.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos de Tiempos Muertos." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }



        //Cargar archivo TTMM
        private void btnUploadTtmm_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de Tiempos Muertos en el servidor." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataTTMM))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }
        //Carga datos de Tiempos Muertos a DataBase
        void UploadDataTTMM()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtArchivoTTMM.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                grillaTTMM.Invoke((MethodInvoker)(() => grillaTTMM.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "ProduccionTTMMUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                txtArchivoTTMM.Invoke((MethodInvoker)(() => txtArchivoTTMM.Text = ""));
                btnUploadTtmm.Invoke((MethodInvoker)(() => btnUploadTtmm.Visible = false));
                int rowcount = grillaTTMM.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Acciones botón buscar TTMM
        private void btnBuscarTTMM_Click(object sender, EventArgs e)
        {
            if (dateTTMMHasta.Value < dateTTMMdesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataTTMM))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        void SearchDataTTMM()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtTTMM = new System.Data.DataTable("dtTTMM"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TTMM @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dateTTMMdesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dateTTMMHasta.Value);
                            if (chkIncluirAmbasFechasTTMM.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtTTMM);
                            dtHistorialTTMM.Invoke((MethodInvoker)(() => dtHistorialTTMM.DataSource = dtTTMM));
                            int registrostotales = dtHistorialTTMM.RowCount;
                            if (dtHistorialTTMM.RowCount == 0)
                            {
                                lblRegistrosEncontradosTTMM.Invoke((MethodInvoker)(() => lblRegistrosEncontradosTTMM.Text = "Registros encontrados: 0"));
                                btnDeleteRecordsTTMM.Invoke((MethodInvoker)(() => btnDeleteRecordsTTMM.Enabled = false));
                            }
                            else
                            {
                                btnDeleteRecordsTTMM.Invoke((MethodInvoker)(() => btnDeleteRecordsTTMM.Enabled = true));
                                lblRegistrosEncontradosTTMM.Invoke((MethodInvoker)(() => lblRegistrosEncontradosTTMM.Text = $"Registros encontrados: {dtHistorialTTMM.RowCount - 1}"));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnBuscarTTMM.Invoke((MethodInvoker)(() => btnBuscarTTMM.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteRecordsTTMM_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataTTMM))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado");
                }
            }
        }

        void DeleteDataTTMM()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtTTMM = new System.Data.DataTable("TTMM"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TTMM @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dateTTMMdesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dateTTMMHasta.Value);
                            if (chkIncluirAmbasFechasTTMM.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtTTMM);
                            dtHistorialTTMM.Invoke((MethodInvoker)(() => dtHistorialTTMM.DataSource = dtTTMM));
                            btnDeleteRecordsTTMM.Invoke((MethodInvoker)(() => btnDeleteRecordsTTMM.Enabled = false));
                            lblRegistrosEncontradosTTMM.Invoke((MethodInvoker)(() => lblRegistrosEncontradosTTMM.Text = "Registros encontrados: 0"));
                            
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //Arrastrar y soltar
        private void grillaExcel_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        //Arrastrar y soltar
        private void grillaExcel_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtFilename.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreview))
            {
                loadPreview.ShowDialog(this);
            }
        }
        //Arrastrar y soltar
        private void grillaTTMM_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void grillaTTMM_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivoTTMM = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtArchivoTTMM.Text = rutaArchivoTTMM[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewTTMM))
            {
                loadPreview.ShowDialog(this);
            }
        }

        //Arrastrar y soltar consumos
        private void grillaConsumos_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            tbUbicacionReporteConsumo.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewConsumo))
            {
                loadPreview.ShowDialog(this);
            }
        }

        private void grillaConsumos_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        //Cargar consumos

        private void btnCargarConsumos_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de consumo." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataConsumo))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Sube la data en un solo ciclo
        void UploadDataConsumo()
        {
            for (int i = 0; i < 1; i++)
            {
                subiendoDataConsumo();
                Thread.Sleep(2);
            }
        }

        //Carga archivo Excel Consumo
        void subiendoDataConsumo()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + tbUbicacionReporteConsumo.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                grillaConsumos.Invoke((MethodInvoker)(() => grillaConsumos.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "ConsumoUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();                

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                tbUbicacionReporteConsumo.Invoke((MethodInvoker)(() => tbUbicacionReporteConsumo.Text = ""));
                btnCargar.Invoke((MethodInvoker)(() => btnCargar.Visible = false));
                int rowcount = grillaConsumos.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Vista previa de archivo Excel de consumos
        void ExcelPreviewConsumo()
        {
            try
            {
                string conexionSIGP = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + tbUbicacionReporteConsumo.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionSIGP);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionSIGP);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                grillaConsumos.Invoke((MethodInvoker)(() => grillaConsumos.DataSource = ds.Tables[0]));
                btnCargarConsumos.Invoke((MethodInvoker)(() => btnCargarConsumos.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos de consumos." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnSearchConsumos_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFilename.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                tbUbicacionReporteConsumo.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewConsumo))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }

        private void btnBuscarHistorialConsumos_Click(object sender, EventArgs e)
        {
            if (dateConsumoHasta.Value < dateConsumoDesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataConsumo))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        //Busca datos en la tabla
        void SearchDataConsumo()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("Consumo"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_CONSUMO @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dateConsumoDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dateConsumoHasta.Value);
                            if (chkIncluirAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumo);
                            dtHistorialConsumo.Invoke((MethodInvoker)(() => dtHistorialConsumo.DataSource = dtConsumo));
                            btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                            if (dtHistorialConsumo.RowCount == 0)
                            {
                                lblRegistrosConsumos.Invoke((MethodInvoker)(() => lblRegistrosConsumos.Text = $"Registros encontrados: 0"));
                                btnDeleteRecordsConsumos.Invoke((MethodInvoker)(() => btnDeleteRecordsConsumos.Enabled = false));
                            }
                            else
                            {
                                lblRegistrosConsumos.Invoke((MethodInvoker)(() => lblRegistrosConsumos.Text = $"Registros encontrados: {dtHistorialConsumo.RowCount - 1}"));
                                btnDeleteRecordsConsumos.Invoke((MethodInvoker)(() => btnDeleteRecordsConsumos.Enabled = true));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteRecordsConsumos_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataConsumos))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado");
                }
            }
        }

        void DeleteDataConsumos()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumos = new System.Data.DataTable("Consumos"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_CONSUMO @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dateConsumoDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dateConsumoHasta.Value);
                            if (chkIncluirAmbasFechasTTMM.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumos);
                            dtHistorialConsumo.Invoke((MethodInvoker)(() => dtHistorialConsumo.DataSource = dtConsumos));
                            btnDeleteRecordsConsumos.Invoke((MethodInvoker)(() => btnDeleteRecordsConsumos.Enabled = false));
                            lblRegistrosConsumos.Invoke((MethodInvoker)(() => lblRegistrosConsumos.Text = "Registros encontrados: 0"));
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // TURNOS PRACTICOS


        private void GrillaTurnosPracticos_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtFileTTPP.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewTTPP))
            {
                loadPreview.ShowDialog(this);
            }
        }

        private void GrillaTurnosPracticos_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        void ExcelPreviewTTPP()
        {
            try
            {
                string conexionTTPP = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFileTTPP.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionTTPP);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionTTPP);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                GrillaTurnosPracticos.Invoke((MethodInvoker)(() => GrillaTurnosPracticos.DataSource = ds.Tables[0]));
                btnCargaTTPP.Invoke((MethodInvoker)(() => btnCargaTTPP.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos de consumos." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnOpenTTPPFIle_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFileTTPP.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFileTTPP.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewTTPP))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }

        // Cargar Turnos Practicos
        private void btnCargaTTPP_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de consumo." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataTTPP))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Sube la data en un solo ciclo
        void UploadDataTTPP()
        {
            for (int i = 0; i < 1; i++)
            {
                subiendoDataTTPP();
                Thread.Sleep(2);
            }
        }

        //Carga archivo Excel Turnos Prácticos
        void subiendoDataTTPP()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFileTTPP.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                GrillaTurnosPracticos.Invoke((MethodInvoker)(() => GrillaTurnosPracticos.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "TurnosPracticosUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                btnCargaTTPP.Invoke((MethodInvoker)(() => populateDataGridViewTurnosPracticos()));
                txtFileTTPP.Invoke((MethodInvoker)(() => txtFileTTPP.Text = ""));
                btnCargaTTPP.Invoke((MethodInvoker)(() => btnCargaTTPP.Visible = false));
                int rowcount = GrillaTurnosPracticos.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);


            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Guardar datos actualizados de grilla
        private void btnHistorialTTPPGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtTurno1Practico.Text != "" || txtTurno2Practico.Text != "" || txtTurno3Practico.Text != "")
                {

                    MessageBox.Show("Por favor, seleccione una fila para actualizar");
                }
                else
                {
                    cmd = new SqlCommand("UPDATE TURNOSPRACTICOSUPLOADRAW SET [fecha]=@fecha,[mes ]=@mes,[año]=@anio,[turno1]=@turno1,[turno2]=@turno2,[turno3]=@turno3,[LastUpdate]=@LastUpdate where ID = @ID", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@ID", ID);
                    cmd.Parameters.AddWithValue("@fecha", dtimeFechatxt.Value);
                    cmd.Parameters.AddWithValue("@mes", cbPracticomes.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@anio", txtAnioPractico.Text);
                    cmd.Parameters.AddWithValue("@turno1", txtTurno1Practico.Text);
                    cmd.Parameters.AddWithValue("@turno2", txtTurno2Practico.Text);
                    cmd.Parameters.AddWithValue("@turno3", txtTurno3Practico.Text);
                    cmd.Parameters.AddWithValue("@LastUpdate", dtimeLastUpdatePractico.Value);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                    lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                    populateDataGridViewTurnosPracticos();
                    ClearDataInput();
                    SaveLog("Update");
                    if (panel22.Visible == true)
                    {
                        panel22.Visible = false;
                    }
                    else
                    {
                        panel22.Visible = true;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private void dtHistorialTurnosPracticos_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Delete)
            {
                var resultado = MessageBox.Show("Si elimina este registro, no lo podrá recuperar." + "\n" + "Tendrá que volver a ingresar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (resultado == DialogResult.Yes)
                {
                    try
                    {   //Conecta antes de buscar la data

                        con = new SqlConnection();
                        con.ConnectionString = @"Data Source=10.2.0.148\DEVSTAR;Initial Catalog=Cargas;Persist Security Info=True;User ID=sa;Password=L30$2Kv.Tv112.c";
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }
                        SqlCommand cmd = new SqlCommand("DELETE FROM [TurnosPracticosUploadRaw] WHERE ID = " + dtHistorialTurnosPracticos.SelectedRows[0].Cells[0].Value.ToString() + "", con);
                        cmd.ExecuteNonQuery();
                        dtHistorialTurnosPracticos.Rows.RemoveAt(dtHistorialTurnosPracticos.SelectedRows[0].Index);

                        lblEstatusAcciones.ForeColor = System.Drawing.Color.Blue;
                        lblEstatusAcciones.Text = "Registro eliminado.";

                        {
                            con.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnHistorialTTPPBuscar_Click(object sender, EventArgs e)
        {
            if (dtTTPPHasta.Value < dtTTPPDesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataTurnosPracticos))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        //Busca datos en la tabla Turnos Practicos
        void SearchDataTurnosPracticos()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("Consumo"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TURNOS_PRACTICOS @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtTTPPDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtTTPPHasta.Value);
                            if (chkTTPPAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumo);
                            dtHistorialTurnosPracticos.Invoke((MethodInvoker)(() => dtHistorialTurnosPracticos.DataSource = dtConsumo));
                            btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = true));
                            if (dtHistorialTurnosPracticos.RowCount == 0)
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: 0"));
                                btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = false));
                            }
                            else
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: {dtHistorialTurnosPracticos.RowCount - 1}"));
                                btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = true));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnHistorialTTPPBuscar.Invoke((MethodInvoker)(() => btnHistorialTTPPBuscar.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnHistorialTTPPBorrar_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataTurnosPracticos))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado Masivo");
                }
            }
        }

        void DeleteDataTurnosPracticos()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumos = new System.Data.DataTable("TurnosPracticos"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TURNOS_PRACTICOS @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtTTPPDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtTTPPHasta.Value);
                            if (chkTTPPAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumos);
                            dtHistorialTurnosPracticos.Invoke((MethodInvoker)(() => dtHistorialTurnosPracticos.DataSource = dtConsumos));
                            btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = false));
                            if (dtHistorialTurnosPracticos.RowCount == 0)
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: 0"));
                                btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = false));
                            }
                            else
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: {dtHistorialTurnosPracticos.RowCount - 1}"));
                                btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => btnHistorialTTPPBorrar.Enabled = true));
                            }
                            btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => populateDataGridViewTurnosPracticos()));

                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        //POBLAR GRILLA AL INICIAR PANTALLA Y ACTUALIZAR
        void populateDataGridViewTurnosPracticos()
        {
            try
            {   //Conecta antes de buscar la data

                con.Open();
                System.Data.DataTable dt = new System.Data.DataTable();
                adapt = new SqlDataAdapter("SELECT TOP 1000 [ID] as ID, [fecha] as Fecha, [mes] as Mes, [año] as [Año], [turno1] as [Turno 1], [turno2] as [Turno 2], [turno3] as [Turno 3], [LastUpdate] as [Última actualización] FROM [TurnosPracticosUploadRaw] ORDER BY [Fecha] DESC", con);
                adapt.Fill(dt);
                dtHistorialTurnosPracticos.DataSource = dt;
                con.Close();
            }
            catch (Exception ex)
            {
                btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //TURNOS REALES - ARRASTRAR Y SOLTAR
        private void GrillaTurnosReales_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtTurnosRealesFilePath.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewTTRR))
            {
                loadPreview.ShowDialog(this);
            }
        }

        private void GrillaTurnosReales_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        void ExcelPreviewTTRR()
        {
            try
            {
                string conexionTTRR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtTurnosRealesFilePath.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionTTRR);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionTTRR);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                GrillaTurnosReales.Invoke((MethodInvoker)(() => GrillaTurnosReales.DataSource = ds.Tables[0]));
                btnCargaTTRR.Invoke((MethodInvoker)(() => btnCargaTTRR.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos de consumos." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnOpenFileTTRR_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFileTTPP.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtTurnosRealesFilePath.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewTTRR))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }

        private void btnCargaTTRR_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de consumo." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataTTRR))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Carga datos de Tiempos Muertos a DataBase
        void UploadDataTTRR()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtTurnosRealesFilePath.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                GrillaTurnosReales.Invoke((MethodInvoker)(() => GrillaTurnosReales.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "TurnosRealesUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                txtTurnosRealesFilePath.Invoke((MethodInvoker)(() => txtTurnosRealesFilePath.Text = ""));
                btnCargaTTRR.Invoke((MethodInvoker)(() => btnCargaTTRR.Visible = false));
                int rowcount = GrillaTurnosReales.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnCargaTTRR.Invoke((MethodInvoker)(() => populateDataGridViewTurnosReales()));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //POBLAR GRILLA AL INICIAR PANTALLA Y ACTUALIZAR
        void populateDataGridViewTurnosReales()
        {
            try
            {   //Conecta antes de buscar la data

                try
                {   //Conecta antes de buscar la data

                    con.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapt = new SqlDataAdapter("SELECT TOP 1000 [ID] as ID, [fecha] as Fecha, [mes] as Mes, [año] as [Año], [turno1] as [Turno 1], [turno2] as [Turno 2], [turno3] as [Turno 3], [LastUpdate] as [Última actualización] FROM [TurnosRealesUploadRaw] ORDER BY [Fecha] DESC", con);
                    adapt.Fill(dt);
                    dtHistorialTurnosReales.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    btnHistorialTTRRBuscar.Invoke((MethodInvoker)(() => btnHistorialTTRRBuscar.Enabled = true));
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //POBLAR Y REFRESCAR PRODUCTIVIDAD POTENCIAL
        void populateDataGridViewProductividadPotencial()
        {
            try
            {   //Conecta antes de buscar la data

                try
                {   //Conecta antes de buscar la data

                    con.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapt = new SqlDataAdapter("SELECT [id_maq] as [ID],[maquina] as [Máquina],[m3/hr potencial] [m3/hr Potencial],[LastUpdate] as [última actualización] FROM [dbo].[ProductividadPotencialUploadRaw] order by [LastUpdate] desc", con);
                    adapt.Fill(dt);
                    dtHistoricoPP.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //POBLAR GRILLA AL INICIAR PANTALLA Y ACTUALIZAR
        void populateDataGridViewPlanMensualxMaquina()
        {
            try
            {   //Conecta antes de buscar la data

                try
                {   //Conecta antes de buscar la data

                    con.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapt = new SqlDataAdapter("SELECT TOP 1000 [ID] as ID, [maq] as [Máquina], [plan] as [Plan], [Fecha] as Fecha, [LastUpdate] as [Última actualización] FROM [HorasProgramadasxMaquinaUploadRaw] ORDER BY [Fecha] DESC", con);
                    adapt.Fill(dt);
                    dtGrillaPlanMensualxMaquina.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    btnHistorialTTRRBuscar.Invoke((MethodInvoker)(() => btnHistorialTTRRBuscar.Enabled = true));
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                btnBuscarHistorialConsumos.Invoke((MethodInvoker)(() => btnBuscarHistorialConsumos.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dtHistorialTurnosReales_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                var resultado = MessageBox.Show("Si elimina este registro, no lo podrá recuperar." + "\n" + "Tendrá que volver a ingresar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (resultado == DialogResult.Yes)
                {
                    try
                    {   //Conecta antes de buscar la data

                        con = new SqlConnection();
                        con.ConnectionString = @"Data Source=10.2.0.148\DEVSTAR;Initial Catalog=Cargas;Persist Security Info=True;User ID=sa;Password=L30$2Kv.Tv112.c";
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        SqlCommand sqlCommand = new SqlCommand("DELETE FROM [TurnosRealesUploadRaw] WHERE ID = " + dtHistorialTurnosReales.SelectedRows[0].Cells[0].Value.ToString() + "", con);
                        SqlCommand cmd = sqlCommand;
                        cmd.ExecuteNonQuery();
                        dtHistorialTurnosReales.Rows.RemoveAt(dtHistorialTurnosReales.SelectedRows[0].Index);

                        lblEstatusAcciones.ForeColor = System.Drawing.Color.Blue;
                        lblEstatusAcciones.Text = "Registro eliminado.";

                        {
                            con.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        btnHistorialTTRRBuscar.Invoke((MethodInvoker)(() => btnHistorialTTRRBuscar.Enabled = true));
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnHistorialTTRRBuscar_Click(object sender, EventArgs e)
        {
            if (dtTurnosRealesHasta.Value < dtTurnosRealesDesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataTurnosReales))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        //Busca datos en la tabla Turnos Practicos
        void SearchDataTurnosReales()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("Consumo"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TURNOS_REALES @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtTurnosRealesDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtTurnosRealesHasta.Value);
                            if (chkTTRRAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumo);
                            dtHistorialTurnosReales.Invoke((MethodInvoker)(() => dtHistorialTurnosReales.DataSource = dtConsumo));
                            btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = true));
                            if (dtHistorialTurnosPracticos.RowCount == 0)
                            {
                                lblHistorialTTRREncontrados.Invoke((MethodInvoker)(() => lblHistorialTTRREncontrados.Text = $"Registros encontrados: 0"));
                                btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = false));
                            }
                            else
                            {
                                lblHistorialTTRREncontrados.Invoke((MethodInvoker)(() => lblHistorialTTRREncontrados.Text = $"Registros encontrados: {dtHistorialTurnosReales.RowCount - 1}"));
                                btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = true));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnHistorialTTRRBuscar.Invoke((MethodInvoker)(() => btnHistorialTTRRBuscar.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnHistorialTTRRBorrarMasivo_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataTurnoReal))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado masivo");
                }
            }
        }

        void DeleteDataTurnoReal()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumos = new System.Data.DataTable("TurnosPracticos"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_TURNOS_REALES @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtTurnosRealesDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtTurnosRealesHasta.Value);
                            if (chkTTRRAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumos);
                            dtHistorialTurnosReales.Invoke((MethodInvoker)(() => dtHistorialTurnosReales.DataSource = dtConsumos));
                            btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = false));
                            if (dtHistorialTurnosReales.RowCount == 0)
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: 0"));
                                btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = false));
                            }
                            else
                            {
                                lblRegistrosTTPP.Invoke((MethodInvoker)(() => lblRegistrosTTPP.Text = $"Registros encontrados: {dtHistorialTurnosPracticos.RowCount - 1}"));
                                btnHistorialTTRRBorrarMasivo.Invoke((MethodInvoker)(() => btnHistorialTTRRBorrarMasivo.Enabled = true));
                            }
                            btnHistorialTTPPBorrar.Invoke((MethodInvoker)(() => populateDataGridViewTurnosPracticos()));

                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnRecargar_Click(object sender, EventArgs e)
        {
            populateDataGridViewTurnosPracticos();
            dtHistorialTurnosPracticos.Update();
            dtHistorialTurnosPracticos.Refresh();
        }

        private void btnRecargarTablaTurnosReales_Click(object sender, EventArgs e)
        {
            dtHistorialTurnosReales.Update();
            dtHistorialTurnosReales.Refresh();
            populateDataGridViewTurnosReales();
        }

        private void btnHistorialTTRRGuardar_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                //populateDataGridViewTurnosReales();
                ////Conecta antes de eliminar la data

                //cmdbl = new SqlCommandBuilder(adap);
                //adap.Update(ds, "InformationTiempoReal");

                //if (con.State == ConnectionState.Open)
                //{
                //    con.Close();
                //}
                lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                populateDataGridViewTurnosReales();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private void dtHistorialTurnosPracticos_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (panel22.Visible == true)
            {
                panel22.Visible = false;
            }
            btnGuardarPractico.Visible = false;
            if (btnHistorialTTPPGuardar.Visible == false)
            {
                btnHistorialTTPPGuardar.Visible = true;
            }
            ID = Convert.ToInt32(dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[0].Value.ToString());
            dtimeFechatxt.Value = Convert.ToDateTime(dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[1].Value.ToString());
            cbPracticomes.SelectedIndex = Convert.ToInt32(dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[2].Value.ToString());
            txtAnioPractico.Text = dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[3].Value.ToString(); ;
            txtTurno1Practico.Text = dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[4].Value.ToString(); ;
            txtTurno2Practico.Text = dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[5].Value.ToString(); ;
            txtTurno3Practico.Text = dtHistorialTurnosPracticos.Rows[e.RowIndex].Cells[6].Value.ToString(); ;
            dtimeLastUpdatePractico.Value = DateTime.Today;
        }

        private void btnGuardarPractico_Click(object sender, EventArgs e)
        {
            if (txtTurno1Practico.Text == "" || txtTurno2Practico.Text == "" || txtTurno3Practico.Text == "")
            {
                MessageBox.Show("Debe rellenar los campos en blanco para continuar.", "Campos vacíos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try 
                {
                    cmd = new SqlCommand("insert into TURNOSPRACTICOSUPLOADRAW([fecha],[mes ],[año],[turno1],[turno2],[turno3],[LastUpdate]) values(@fecha, @mes, @anio, @turno1, @turno2, @turno3, @LastUpdate)", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@fecha", dtimeFechatxt.Value);
                    cmd.Parameters.AddWithValue("@mes", cbPracticomes.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@anio", txtAnioPractico.Text);
                    cmd.Parameters.AddWithValue("@turno1", txtTurno1Practico.Text);
                    cmd.Parameters.AddWithValue("@turno2", txtTurno2Practico.Text);
                    cmd.Parameters.AddWithValue("@turno3", txtTurno3Practico.Text);
                    cmd.Parameters.AddWithValue("@LastUpdate", dtimeLastUpdatePractico.Value);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                    lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                    populateDataGridViewTurnosPracticos();
                    ClearDataInput();
                    SaveLog("Insert");

                    if (panel22.Visible == true)
                    {
                        panel22.Visible = false;
                    }
                    else
                    {
                        panel22.Visible = true;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ClearDataInput()
        {
            dtimeFechatxt.Value = DateTime.Today;
            txtAnioPractico.Text = "";
            txtTurno1Practico.Text = "";
            txtTurno2Practico.Text = "";
            txtTurno3Practico.Text = "";
            dtimeLastUpdatePractico.Value = DateTime.Today;

            dtimeFechatxtReal.Value = DateTime.Today;
            txtAnioReal.Text = "";
            txtTurno1Real.Text = "";
            txtTurno2Real.Text = "";
            txtTurno3Real.Text = "";
            dtimeLastUpdateReal.Value = DateTime.Today;
            ID = 0;

            txtMaquinaPP.Text = "";
            txtm3HrPlanPP.Text = "";
        }

        private void btnGuardarReal_Click(object sender, EventArgs e)
        {
            if (txtTurno1Real.Text == "" || txtTurno2Real.Text == "" || txtTurno3Real.Text == "")
            {
                MessageBox.Show("Debe rellenar los campos en blanco para continuar.", "Campos vacíos", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    cmd = new SqlCommand("insert into [TurnosRealesUploadRaw] ([fecha],[mes ],[año],[turno1],[turno2],[turno3],[LastUpdate]) values(@fecha, @mes, @anio, @turno1, @turno2, @turno3, @LastUpdate)", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@fecha", dtimeFechatxtReal.Value);
                    cmd.Parameters.AddWithValue("@mes", cbRealmes.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@anio", txtAnioReal.Text);
                    cmd.Parameters.AddWithValue("@turno1", txtTurno1Real.Text);
                    cmd.Parameters.AddWithValue("@turno2", txtTurno2Real.Text);
                    cmd.Parameters.AddWithValue("@turno3", txtTurno3Real.Text);
                    cmd.Parameters.AddWithValue("@LastUpdate", dtimeLastUpdateReal.Value);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                    lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                    populateDataGridViewTurnosReales();
                    ClearDataInput();
                    SaveLog("Insert");
                    if (panel23.Visible == true)
                    {
                        panel23.Visible = false;
                    }
                    else
                    {
                        panel23.Visible = true;
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } 
            }
        }

        private void dtHistorialTurnosReales_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dtHistorialTurnosReales.Rows[e.RowIndex].Cells[0].Value.ToString());
            dtimeFechatxt.Value = Convert.ToDateTime(dtHistorialTurnosReales.Rows[e.RowIndex].Cells[1].Value.ToString());
            cbRealmes.SelectedIndex = Convert.ToInt32(dtHistorialTurnosReales.Rows[e.RowIndex].Cells[2].Value.ToString());
            txtAnioReal.Text = dtHistorialTurnosReales.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtTurno1Real.Text = dtHistorialTurnosReales.Rows[e.RowIndex].Cells[4].Value.ToString();
            txtTurno2Real.Text = dtHistorialTurnosReales.Rows[e.RowIndex].Cells[5].Value.ToString();
            txtTurno3Real.Text = dtHistorialTurnosReales.Rows[e.RowIndex].Cells[6].Value.ToString();
            dtimeLastUpdateReal.Value = DateTime.Today;
        }

        private void btnActualizarPractico_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtTurno1Real.Text == "" || txtTurno2Real.Text == "" || txtTurno3Real.Text == "")
                {
                    MessageBox.Show("Por favor, seleccione una fila para actualizar");
                }
                else
                {
                    cmd = new SqlCommand("UPDATE TurnosRealesUploadRaw SET [fecha]=@fecha,[mes ]=@mes,[año]=@anio,[turno1]=@turno1,[turno2]=@turno2,[turno3]=@turno3,[LastUpdate]=@LastUpdate where ID = @ID", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@ID", ID);
                    cmd.Parameters.AddWithValue("@fecha", dtimeFechatxtReal.Value);
                    cmd.Parameters.AddWithValue("@mes", cbRealmes.SelectedItem.ToString());
                    cmd.Parameters.AddWithValue("@anio", txtAnioReal.Text);
                    cmd.Parameters.AddWithValue("@turno1", txtTurno1Real.Text);
                    cmd.Parameters.AddWithValue("@turno2", txtTurno2Real.Text);
                    cmd.Parameters.AddWithValue("@turno3", txtTurno3Real.Text);
                    cmd.Parameters.AddWithValue("@LastUpdate", dtimeLastUpdateReal.Value);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                    lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                    populateDataGridViewTurnosReales();
                    ClearDataInput();
                    SaveLog("Update");

                    if (panel23.Visible == true)
                    {
                        panel23.Visible = false;
                    }
                    else
                    {
                        panel23.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            Cursor.Current = Cursors.Default;
        }

        private void dtCargaPP_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtFilelocationPP.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewPP))
            {
                loadPreview.ShowDialog(this);
            }
        }

        private void dtCargaPP_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        void ExcelPreviewPP()
        {
            try
            {
                string conexionSIGP = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilelocationPP.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionSIGP);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionSIGP);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select [maquina], [m3/hr potencial] from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                dtCargaPP.Invoke((MethodInvoker)(() => dtCargaPP.DataSource = ds.Tables[0]));
                btnCargarPP.Invoke((MethodInvoker)(() => btnCargarPP.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos SIGP." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnCargarPP_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de consumo." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataPP))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Carga datos de PP a DataBase
        void UploadDataPP()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilelocationPP.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select '',[maquina], [m3/hr potencial] from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                dtCargaPP.Invoke((MethodInvoker)(() => dtCargaPP.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "ProductividadPotencialUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();



                //CREA LAST UPDATE para Potencial
                try
                {   //Conecta antes de buscar la data

                    con.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapt = new SqlDataAdapter("UPDATE [ProductividadPotencialUploadRaw] SET [LastUpdate] = CURRENT_TIMESTAMP WHERE[LastUpdate] is null", con);
                    adapt.Fill(dt);
                    dtHistoricoPP.DataSource = dt;
                    con.Close();
                }
                catch (Exception ex)
                {
                    btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                txtFilelocationPP.Invoke((MethodInvoker)(() => txtFilelocationPP.Text = ""));
                btnCargarPP.Invoke((MethodInvoker)(() => btnCargarPP.Visible = false));
                int rowcount = dtCargaPP.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnCargarPP.Invoke((MethodInvoker)(() => populateDataGridViewProductividadPotencial()));

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOpenFilePP_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFilelocationPP.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFilelocationPP.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewPP))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }

        private void btnBuscarPP_Click(object sender, EventArgs e)
        {
            try
            {   //Conecta antes de buscar la data

                try
                {   //Conecta antes de buscar la data
                    DateTime UltimaActualizacion = dtLastUpdatePP.Value;
                    con.Open();
                    System.Data.DataTable dt = new System.Data.DataTable();
                    adapt = new SqlDataAdapter("SELECT[id_maq] as [ID Maquina],[maquina] as [Máquina],[m3/hr potencial] [m3/hr Potencial],[LastUpdate] as [última actualización] FROM [dbo].[ProductividadPotencialUploadRaw] order by [LastUpdate] DESC", con);
                    adapt.Fill(dt);
                    dtHistoricoPP.DataSource = dt;
                    btnBorrarPP.Invoke((MethodInvoker)(() => btnBorrarPP.Enabled = true));
                    con.Close();
                }
                catch (Exception ex)
                {
                    btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            catch (Exception ex)
            {
                btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void SearchDataPP()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("Consumo"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_PRODUCTIVIDAD_POTENCIAL @ACCION,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtLastUpdatePP.Value);
                            if (chkIncluirAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumo);
                            dtHistoricoPP.Invoke((MethodInvoker)(() => dtHistoricoPP.DataSource = dtConsumo));
                            btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                            if (dtHistoricoPP.RowCount == 0)
                            {
                                btnBorrarPP.Invoke((MethodInvoker)(() => btnBorrarPP.Enabled = false));
                            }
                            else
                            {
                                btnBorrarPP.Invoke((MethodInvoker)(() => btnBorrarPP.Enabled = true));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnBuscarPP.Invoke((MethodInvoker)(() => btnBuscarPP.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

       //Plan Mensual x Maquina - ARRASTRAR Y SOLTAR
        private void dataGridView2_DragDrop(object sender, DragEventArgs e)
        {
            string[] rutaArchivo = (string[])e.Data.GetData(DataFormats.FileDrop);
            txtFilePlanMensual.Text = rutaArchivo[0];

            using (FormWait loadPreview = new FormWait(ExcelPreviewPlanMensualxMaquina))
            {
                loadPreview.ShowDialog(this);
            }
        }

        private void dtGrillaPlanMensualxMaquina_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Está a punto de cargar datos de consumo." + "\n" + "La operación podría tardar unos segundos." + "\n" + "¿Desea confirmar?", "Confirmar acción", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormUploading wait = new FormUploading(UploadDataPlanMensualxMaquina))
                {
                    wait.ShowDialog(this);
                    SaveLog("Carga");
                }
            }
        }

        //Carga datos de PP a DataBase
        void UploadDataPlanMensualxMaquina()
        {
            try
            {
                string conexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilePlanMensual.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexion);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;

                DataSet ds = new DataSet();
                adaptador.Fill(ds);

                dtGrillaPlanMensualxMaquina.Invoke((MethodInvoker)(() => dtGrillaPlanMensualxMaquina.DataSource = ds.Tables[0]));
                origen.Close();

                SqlConnection conexion_destino = new SqlConnection();
                conexion_destino.ConnectionString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                conexion_destino.Open();

                SqlBulkCopy importar = default(SqlBulkCopy);
                importar = new SqlBulkCopy(conexion_destino);
                importar.DestinationTableName = "HorasProgramadasxMaquinaUploadRaw";
                importar.WriteToServer(ds.Tables[0]);
                conexion_destino.Close();

                using (FormUploading wait = new FormUploading(UploadData))
                {
                    wait.Close();
                }
                txtFilelocationPP.Invoke((MethodInvoker)(() => txtFilelocationPP.Text = ""));
                btnCargarPlanMensualxMaquina.Invoke((MethodInvoker)(() => btnCargarPlanMensualxMaquina.Visible = false));
                int rowcount = dtGrillaPlanMensualxMaquina.RowCount - 1;
                MessageBox.Show("La sábana fue cargada correctamente." + "\n" + "Se cargaron " + rowcount + " registros.", "¡Excelente!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                btnCargarPlanMensualxMaquina.Invoke((MethodInvoker)(() => populateDataGridViewPlanMensualxMaquina()));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al cargar los datos. Verifique el archivo, la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo cargar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        void ExcelPreviewPlanMensualxMaquina()
        {
            try
            {
                string conexionTTRR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + txtFilePlanMensual.Text + "';Extended Properties=Excel 12.0;";
                OleDbConnection origen = default(OleDbConnection);
                origen = new OleDbConnection(conexionTTRR);

                //Carga Excel en la grilla antes de importarlo
                OleDbConnection origen1 = default(OleDbConnection);
                origen1 = new OleDbConnection(conexionTTRR);

                OleDbCommand seleccion = default(OleDbCommand);
                seleccion = new OleDbCommand("select * from [Hoja1$]", origen1);

                OleDbDataAdapter adaptador = new OleDbDataAdapter();
                adaptador.SelectCommand = seleccion;
                DataSet ds = new DataSet();

                adaptador.Fill(ds);
                dtGrillaPlanMensualxMaquina.Invoke((MethodInvoker)(() => dtGrillaPlanMensualxMaquina.DataSource = ds.Tables[0]));
                btnCargarPlanMensualxMaquina.Invoke((MethodInvoker)(() => btnCargarPlanMensualxMaquina.Visible = true));
                origen.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("El formato del archivo o sus columnas no coincide con la sábana de datos de consumos." + "\n\n" + ex.Message, "Error al validar archivo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private void btnOpenFilePlanMensual_Click(object sender, EventArgs e)
        {
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Buscar archivo";
            fdlg.FileName = txtFilePlanMensual.Text;
            fdlg.Filter = "Archivo Excel (*.xlsx)|*.xlsx";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;

            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFilePlanMensual.Text = fdlg.FileName;
                using (FormWait loadPreview = new FormWait(ExcelPreviewPlanMensualxMaquina))
                {
                    loadPreview.ShowDialog(this);
                }
            }
        }

        private void btnBuscarPlanMensualxMaquina_Click(object sender, EventArgs e)
        {

            if (dtPlanMensualxMaquinaHasta.Value < dtPlanMensualxMaquinaDesde.Value)
            {
                MessageBox.Show("La fecha de inicio (DESDE) debe ser más antigua que la fecha de fin (HASTA).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                using (FormWait loading = new FormWait(SearchDataPlanMensualxMaquina))
                {
                    loading.ShowDialog(this);
                }
            }
        }

        //Busca datos en la tabla
        void SearchDataPlanMensualxMaquina()
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("PlanMensualxMaquina"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_PLAN_MENSUAL_POR_MAQUINA @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 1);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtPlanMensualxMaquinaDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtPlanMensualxMaquinaHasta.Value);
                            if (chkIncluirAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Rellena datos a la grilla de Histórico
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumo);
                            dtHistorialPlanMensualxMaquina.Invoke((MethodInvoker)(() => dtHistorialPlanMensualxMaquina.DataSource = dtConsumo));
                            btnBuscarPlanMensualxMaquina.Invoke((MethodInvoker)(() => btnBuscarPlanMensualxMaquina.Enabled = true));
                            if (dtHistorialPlanMensualxMaquina.RowCount == 0)
                            {
                                lblRegistrosPlanMensualxMaquina.Invoke((MethodInvoker)(() => lblRegistrosPlanMensualxMaquina.Text = $"Registros encontrados: 0"));
                                btnDeletePlanMensualxMaquina.Invoke((MethodInvoker)(() => btnDeletePlanMensualxMaquina.Enabled = false));
                            }
                            else
                            {
                                lblRegistrosPlanMensualxMaquina.Invoke((MethodInvoker)(() => lblRegistrosPlanMensualxMaquina.Text = $"Registros encontrados: {dtHistorialPlanMensualxMaquina.RowCount - 1}"));
                                btnDeletePlanMensualxMaquina.Invoke((MethodInvoker)(() => btnDeletePlanMensualxMaquina.Enabled = true));
                            }
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                btnBuscarPlanMensualxMaquina.Invoke((MethodInvoker)(() => btnBuscarPlanMensualxMaquina.Enabled = true));
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void SaveLog(string accion)
        {
            try
            {   //Conecta antes de buscar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumo = new System.Data.DataTable("PlanMensualxMaquina"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("EXEC [INSERT_LOG] @INTERFAZ,@ACCION,@FECHA", con))
                        {
                            // Añade parámetros al StoredProcedure
                            
                            sqlCmd.Parameters.AddWithValue("@INTERFAZ", Seleccion);
                            sqlCmd.Parameters.AddWithValue("@ACCION", accion);
                            sqlCmd.Parameters.AddWithValue("@FECHA", DateTime.Now);
                            sqlCmd.ExecuteNonQuery();
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
              
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnBorrarPP_Click(object sender, EventArgs e)
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtProductividadPotencial = new System.Data.DataTable("PP"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_PRODUCTIVIDAD_POTENCIAL @ACCION,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtLastUpdatePP.Value);
                            sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtProductividadPotencial);
                            dtHistoricoPP.Invoke((MethodInvoker)(() => dtHistoricoPP.DataSource = dtProductividadPotencial));
                            btnBorrarPP.Invoke((MethodInvoker)(() => btnBorrarPP.Enabled = false));
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeletePlanMensualxMaquina_Click(object sender, EventArgs e)
        {
            var resultado = MessageBox.Show("Si elimina estos registros, no los podrá recuperar." + "\n" + "Tendrá que volver a cargar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                using (FormWait loading = new FormWait(DeleteDataPlanMensualxMaquina))
                {
                    loading.ShowDialog(this);
                    SaveLog("Borrado");
                }
            }
        }

        void DeleteDataPlanMensualxMaquina()
        {
            try
            {   //Conecta antes de eliminar la data
                string strConnString = ConfigurationManager.ConnectionStrings["Base"].ConnectionString;
                using (SqlConnection con = new SqlConnection(strConnString))
                {
                    if (con.State == ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    using (System.Data.DataTable dtConsumos = new System.Data.DataTable("Consumos"))
                    {
                        using (SqlCommand sqlCmd = new SqlCommand("BUSCAR_BORRAR_SABANA_PLAN_MENSUAL_POR_MAQUINA @ACCION,@FEC_DESDE,@FEC_HASTA,@CHECK", con))
                        {
                            // Añade parámetros al StoredProcedure
                            sqlCmd.Parameters.AddWithValue("@ACCION", 2);
                            sqlCmd.Parameters.AddWithValue("@FEC_DESDE", dtPlanMensualxMaquinaDesde.Value);
                            sqlCmd.Parameters.AddWithValue("@FEC_HASTA", dtPlanMensualxMaquinaHasta.Value);
                            if (chkPlanMensualxMaquinaAmbasFechas.Checked == true)
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 1);
                            }
                            else
                            {
                                sqlCmd.Parameters.AddWithValue("@CHECK", 0);
                            }
                            // Una vez borrado los datos, refresca la grilla
                            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmd);
                            sqlDataAdapter.Fill(dtConsumos);
                            dtHistorialPlanMensualxMaquina.Invoke((MethodInvoker)(() => dtHistorialPlanMensualxMaquina.DataSource = dtConsumos));
                            btnDeletePlanMensualxMaquina.Invoke((MethodInvoker)(() => btnDeletePlanMensualxMaquina.Enabled = false));
                            lblRegistrosPlanMensualxMaquina.Invoke((MethodInvoker)(() => lblRegistrosPlanMensualxMaquina.Text = "Registros encontrados: 0"));
                        }
                    }
                    if (con.State == ConnectionState.Open)
                    {
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dtHistoricoPP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                var resultado = MessageBox.Show("Si elimina este registro, no lo podrá recuperar." + "\n" + "Tendrá que volver a ingresar los datos en la base." + "\n" + "¿Desea confirmar la acción?", "Confirmar Borrado", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (resultado == DialogResult.Yes)
                {
                    try
                    {   //Conecta antes de buscar la data

                        con = new SqlConnection();
                        con.ConnectionString = @"Data Source=10.2.0.148\DEVSTAR;Initial Catalog=Cargas;Persist Security Info=True;User ID=sa;Password=L30$2Kv.Tv112.c";
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        SqlCommand sqlCommand = new SqlCommand("DELETE FROM [HorasProgramadasxMaquinaUploadRaw] WHERE ID = " + dtHistoricoPP.SelectedRows[0].Cells[0].Value.ToString() + "", con);
                        SqlCommand cmd = sqlCommand;
                        cmd.ExecuteNonQuery();
                        dtHistoricoPP.Rows.RemoveAt(dtHistoricoPP.SelectedRows[0].Index);

                        lblEstatusAcciones.ForeColor = System.Drawing.Color.Blue;
                        lblEstatusAcciones.Text = "Registro eliminado.";

                        {
                            con.Close();
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void dtHistoricoPP_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dtHistoricoPP.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtMaquinaPP.Text = dtHistoricoPP.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtm3HrPlanPP.Text = dtHistoricoPP.Rows[e.RowIndex].Cells[2].Value.ToString();
            dtimeLastUpdatePractico.Value = DateTime.Today;
        }



        private void lblEstatusAcciones_Click_1(object sender, EventArgs e)
        {

        }

        private void cambiarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("¿Desea finalizar?" + "\n\n" + "Cualquier cambio no guardado en la sábana de históricos podría perderse.", "Confirmar acción", MessageBoxButtons.YesNo,
          MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                FormEscoger escoger = new FormEscoger();
                escoger.Show();
                this.Hide();
            }
        }

        private void btnPanelShowTTPPNew_Click(object sender, EventArgs e)
        {
            ID = 0;
            dtimeFechatxt.Value = DateTime.Today;
            cbPracticomes.SelectedIndex = 1;
            txtAnioPractico.Text = "";
            txtTurno1Practico.Text = "";
            txtTurno2Practico.Text = "";
            txtTurno3Practico.Text = "";
            dtimeLastUpdatePractico.Value = DateTime.Today;

            btnHistorialTTPPGuardar.Visible = false;
            btnGuardarPractico.Visible = true;
            if (panel22.Visible == true)
            {
                panel22.Visible = false;
            }
            else
            {
                panel22.Visible = true;
            }
        }

        private void btnPanelShowTTPPEdit_Click(object sender, EventArgs e)
        {

            btnHistorialTTPPGuardar.Visible = true;
            btnGuardarPractico.Visible = false;

            if (panel22.Visible == true)
            {
                panel22.Visible = false;
            }
            else
            {
                panel22.Visible = true;
            }
        }

        private void btnPanelShowTurnoRealNew_Click(object sender, EventArgs e)
        {
            ID = 0;
            dtimeFechatxt.Value = DateTime.Today;
            cbRealmes.SelectedIndex = 1;
            txtAnioReal.Text = "";
            txtTurno1Real.Text = "";
            txtTurno2Real.Text = "";
            txtTurno3Real.Text = "";
            dtimeLastUpdateReal.Value = DateTime.Today;



            btnActualizarPractico.Visible = false;
            btnGuardarReal.Visible = true;
            if (panel23.Visible == true)
            {
                panel23.Visible = false;
            }
            else
            {
                panel23.Visible = true;
            }
        }

        private void btnPanelShowTurnoRealEdit_Click(object sender, EventArgs e)
        {
            btnGuardarReal.Visible = false;
            btnActualizarPractico.Visible = true;
            if (panel23.Visible == true)
            {
                panel23.Visible = false;
            }
            else
            {
                panel23.Visible = true;
            }
        }

        private void btnNuevoPP_Click(object sender, EventArgs e)
        {
            accionGuardadoPP = "Nuevo";
            if (txtMaquinaPP.Visible == false)
            {
                txtIDMaquinaPP.Visible = false;
                txtMaquinaPP.Visible = true;
                txtm3HrPlanPP.Visible = true;
                label23.Visible = false;
                label37.Visible = true;
                label38.Visible = true;
                btnGuardarNewPP.Visible = true;
                txtMaquinaPP.Select();
                txtMaquinaPP.Focus();
                this.btnNuevoPP.Image = Subir2.Properties.Resources.back_24x24_1214369;
                this.btnNuevoPP.Text = "Volver";
                btnModificarPP.Visible = false;
                btnBorrarPP.Visible = false;
                btnBuscarPP.Visible = false;
            }
            else 
            {
                txtIDMaquinaPP.Visible = false;
                txtMaquinaPP.Visible = false;
                txtm3HrPlanPP.Visible = false;
                label23.Visible = false;
                label37.Visible = false;
                label38.Visible = false;
                btnGuardarNewPP.Visible = false;
                btnBuscarPP.Select();
                btnBuscarPP.Focus();
                this.btnNuevoPP.Image = Subir2.Properties.Resources.plus_24x24_1214303;
                this.btnNuevoPP.Text = "Nuevo...";
                btnModificarPP.Visible = true;
                btnBorrarPP.Visible = true;
                btnBuscarPP.Visible = true;
            }           
        }

        private void btnGuardarPP_Click(object sender, EventArgs e)
        {
            accionGuardadoPP = "Editar";
            if (txtMaquinaPP.Text == "" || txtm3HrPlanPP.Text == "" || ID == 0)
            {
                MessageBox.Show("Por favor, seleccione una fila para editar", "Hubo un problema", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                if (txtMaquinaPP.Visible == false)
                {
                    txtIDMaquinaPP.Visible = false;
                    txtMaquinaPP.Visible = true;
                    txtm3HrPlanPP.Visible = true;
                    label23.Visible = false;
                    label37.Visible = true;
                    label38.Visible = true;
                    btnModificarPP.Visible = true;
                    txtMaquinaPP.Select();
                    txtMaquinaPP.Focus();
                    this.btnModificarPP.Image = Subir2.Properties.Resources.back_24x24_1214369;
                    this.btnModificarPP.Text = "Volver";
                    btnNuevoPP.Visible = false;
                    btnBorrarPP.Visible = false;
                    btnGuardarNewPP.Visible = true;
                }
                else
                {
                    txtIDMaquinaPP.Visible = false;
                    txtMaquinaPP.Visible = false;
                    txtm3HrPlanPP.Visible = false;
                    label23.Visible = false;
                    label37.Visible = false;
                    label38.Visible = false;
                    btnNuevoPP.Visible = false;
                    btnBuscarPP.Select();
                    btnBuscarPP.Focus();
                    this.btnModificarPP.Image = Subir2.Properties.Resources.pencil_24x24_1214304;
                    this.btnModificarPP.Text = "Modificar";
                    btnModificarPP.Visible = true;
                    btnNuevoPP.Visible = true;
                    btnBorrarPP.Visible = true;
                    btnGuardarNewPP.Visible = false;
                }
            }
        }

        private void btnGuardarNewPP_Click(object sender, EventArgs e)
        {
            if (accionGuardadoPP == "Nuevo")
            {
                if (txtMaquinaPP.Text == "" || txtm3HrPlanPP.Text == "")
                {
                    MessageBox.Show("Debe rellenar los campos en blanco para continuar.", "Campos vacíos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    try
                    {
                        cmd = new SqlCommand("insert into [ProductividadPotencialUploadRaw]([maquina],[m3/hr potencial],[LastUpdate]) values(@maquina, @m3hr, @LastUpdate)", con);
                        con.Open();
                        cmd.Parameters.AddWithValue("@maquina", txtMaquinaPP.Text);
                        cmd.Parameters.AddWithValue("@m3hr", txtm3HrPlanPP.Text);
                        cmd.Parameters.AddWithValue("@LastUpdate", DateTime.Now);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                        lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                        populateDataGridViewProductividadPotencial();
                        ClearDataInput();
                        SaveLog("Insert");

                        txtIDMaquinaPP.Visible = false;
                        txtMaquinaPP.Visible = false;
                        txtm3HrPlanPP.Visible = false;
                        label23.Visible = false;
                        label37.Visible = false;
                        label38.Visible = false;
                        btnGuardarNewPP.Visible = false;
                        btnBuscarPP.Select();
                        btnBuscarPP.Focus();
                        this.btnNuevoPP.Image = Subir2.Properties.Resources.plus_24x24_1214303;
                        this.btnNuevoPP.Text = "Nuevo...";
                        btnModificarPP.Visible = true;
                        btnBorrarPP.Visible = true;
                        btnBuscarPP.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            if (accionGuardadoPP == "Editar")
            {
                if (txtMaquinaPP.Text == "" || txtm3HrPlanPP.Text == "")
                {
                    MessageBox.Show("Por favor, seleccione una fila para editar", "Hubo un problema", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    try
                    {
                        cmd = new SqlCommand("UPDATE [ProductividadPotencialUploadRaw] SET [maquina] = @maquina,[m3/hr potencial] = @m3hr,[LastUpdate] = @LastUpdate WHERE ID_MAQ = @ID", con);
                        con.Open();
                        cmd.Parameters.AddWithValue("@maquina", txtMaquinaPP.Text);
                        cmd.Parameters.AddWithValue("@m3hr", txtm3HrPlanPP.Text);
                        cmd.Parameters.AddWithValue("@LastUpdate", DateTime.Now);
                        cmd.Parameters.AddWithValue("@ID", ID);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        lblEstatusAcciones.ForeColor = System.Drawing.Color.Green;
                        lblEstatusAcciones.Text = "¡Registro guardado exitosamente!.";
                        populateDataGridViewProductividadPotencial();
                        ClearDataInput();
                        SaveLog("Update");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocurrió un error al guardar los datos. Verifique la conexión a la red y vuelva a intentarlo." + "\n \n" + ex.Message, "No se pudo guardar sábana", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }


            
        }
    }
}
