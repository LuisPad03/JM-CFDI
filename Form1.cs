using System.Threading;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;
using System.Runtime.InteropServices;
using Timer = System.Timers.Timer;
using System.Net;
using System.IO;
using System.Xml;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.Streaming;
using MathNet.Numerics.Distributions;
using NPOI.HPSF;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace JM_CFDI
{
    public partial class Form1 : Form
    {
        //string ConsultaSQLManagement;
        static string carpetaDocs;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel_reporte.Visible = false;
            tamañoForm(670, 160);

            btn_reporte.Enabled = false;
            txt_filtro.Enabled = false;
            lstcbx_descripciones.Enabled = false;
        }

        private void txt_ruta_Click(object sender, EventArgs e)
        {
            if(folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                txt_ruta.Text = folderBrowserDialog1.SelectedPath;
                lbl_ruta.Visible = false;
            }
        }
        private void btn_leerArchivos_Click(object sender, EventArgs e)
        {
            grd_columnas.DataSource = null;
            lstcbx_columnas.DataSource = null;
            cbx_filtro.DataSource = null;
            lstcbx_descripciones.DataSource = null;
            txt_filtro.Text = "";
            txt_filtro.Enabled = false;

            if (txt_ruta.Text != "")
            {
                carpetaDocs = txt_ruta.Text;

                tamañoForm(965, 160);

                string[] sArchivos;
                sArchivos = Directory.GetFiles(carpetaDocs);

                List<string> lstcbxColumnas = new List<string>();
                List<string> cbxFiltro = new List<string>();

                if (obtieneDatos(sArchivos, lstcbxColumnas, cbxFiltro, "columnas"))
                {
                    panel_reporte.Visible = true;
                    tamañoForm(965, 430);

                    List<string> filtroColumnas = lstcbxColumnas.GroupBy(x => x)
                        .Where(g => g.Count() > 1)
                        .Select(x => x.Key)
                        .ToList();
                    lstcbx_columnas.DataSource = filtroColumnas;

                    //List<string> resultOK = cbxFiltro.Distinct().ToList();
                    List<string> filtroFiltro = cbxFiltro.GroupBy(x => x)
                        .Where(g => g.Count() > 1)
                        .Select(x => x.Key)
                        .ToList();
                    cbx_filtro.DataSource = filtroFiltro; ;
                }
                else MessageBox.Show("Los documentos contenidos en esta carpeta no son compatibles. \n\nLos documentos deben de contar con la extencion \".XLSX\", es decir archivos tipo excel.", "Documentos no compatibles", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Ingresa la carpeta que contenga los CFDI, en formato excel.", "Información incompleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void cbx_filtro_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (carpetaDocs != "")
            {
                string[] sArchivos;
                sArchivos = Directory.GetFiles(carpetaDocs);

                List<string> lstFiltro = new List<string>();
                List<string> filtroOK = new List<string>();

                if (obtieneDatos(sArchivos, lstFiltro, filtroOK, "filtro"))
                {
                    filtroOK = lstFiltro.Distinct().ToList();
                    lstcbx_descripciones.DataSource = filtroOK;
                }

                txt_filtro.Enabled = true;
                lstcbx_descripciones.Enabled = true;

                if (lstcbx_descripciones.Items.Count != 0 && lstcbx_columnas.CheckedItems.Count != 0) btn_reporte.Enabled = true;
            }
            else MessageBox.Show("Ingresa la carpeta que contenga los CFDI, en formato excel.", "Información incompleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }
        private bool obtieneDatos(string[] sArchivos, List<string> lst1, List<string> lst2, string tipo)
        {
            bool obtienedatos = false;
            int archs = 0;
            progresBar(archs, sArchivos.Length);
            txt_progreso.Text = "";

            foreach (string rutaarchivo in sArchivos)
            {
                txt_progreso.Text = txt_progreso.Text + "Leyendo archivo: " + Path.GetFileName(rutaarchivo) + "\r\n";
                if (Path.GetExtension(rutaarchivo) == ".xlsx")
                {
                    obtienedatos = true;

                    #region Abre hoja excel a trabajar
                    IWorkbook MiExcel = null;
                    FileStream fs = new FileStream(rutaarchivo, FileMode.Open, FileAccess.Read);

                    if (Path.GetExtension(rutaarchivo) == ".xlsx")
                        MiExcel = new XSSFWorkbook(fs);
                    else
                        MiExcel = new HSSFWorkbook(fs);

                    ISheet hoja = MiExcel.GetSheetAt(1);
                    #endregion

                    if (hoja != null)
                    {
                        if (tipo == "columnas")
                        {
                            txt_progreso.Text = txt_progreso.Text + "--Obteniendo COLUMNAS." + "\r\n";
                            int cantidadcolum = 90;
                            for (int i = 0; i <= cantidadcolum; i++)
                            {
                                IRow fila = hoja.GetRow(0);
                                string celda = fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";
                                if (celda != "")
                                {
                                    lst1.Add(celda);
                                    lst2.Add(celda);
                                }
                            }
                        }
                        if (tipo == "filtro")
                        {
                            string filtro = cbx_filtro.SelectedValue.ToString();
                            int cantidadFilas = hoja.LastRowNum;
                            int colFiltro = 0;
                            bool existefiltro = false;

                            txt_progreso.Text = txt_progreso.Text + "--Obteniendo datos de \"" + filtro.ToUpper() + "\"\r\n";

                            for (int i = 0; i <= 90; i++)
                            {
                                IRow fila = hoja.GetRow(0);
                                string celda = fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";
                                if (celda == filtro)
                                {
                                    colFiltro = i;
                                    existefiltro= true;
                                    break;
                                }
                            }
                            if(existefiltro)
                            {
                                for (int i = 1; i <= cantidadFilas; i++)
                                {
                                    IRow fila = hoja.GetRow(i);
                                    string celda = fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";

                                    if (fila != null && celda != "") lst1.Add(celda);
                                }
                            }
                            else txt_progreso.Text = txt_progreso.Text + "--No se encontraron datos de \"" + filtro.ToUpper() + "\"\r\n";
                        }
                    }
                    MiExcel.Close();
                }
                else txt_progreso.Text = txt_progreso.Text + "\r\n";

                archs++;
                progresBar(archs, sArchivos.Length);
            }

            txt_progreso.Text = txt_progreso.Text.TrimEnd('\n');

            return obtienedatos;
        }
        private void txt_filtro_TextChanged(object sender, EventArgs e)
        {

        }

        private void txt_excel_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txt_excel.Text = openFileDialog1.FileName;
                lbl_excel.Visible = false;
            }
        }
        private void btn_reporte_Click(object sender, EventArgs e)
        {

        }

        private void tamañoForm(int ancho, int alto)
        {
            this.MinimumSize = new System.Drawing.Size(ancho, alto);
            this.MaximumSize = new System.Drawing.Size(ancho, alto);
            this.Size = new System.Drawing.Size(ancho, alto);
        }

        private void lstcbx_columnas_SelectedValueChanged(object sender, EventArgs e)
        { 
            if (lstcbx_columnas.CheckedItems.Count != 0)
            {
                DataTable vistaPrevia = new DataTable();

                for (int x = 0; x < lstcbx_columnas.CheckedItems.Count; x++)
                {
                    vistaPrevia.Columns.Add(lstcbx_columnas.CheckedItems[x].ToString());
                }

                for( int i = 0;i < 7;i++)
                {
                    vistaPrevia.Rows.Add();
                }
                grd_columnas.DataSource = vistaPrevia;

                if (lstcbx_descripciones.Items.Count != 0) btn_reporte.Enabled = true;
            }
            else
            {
                btn_reporte.Enabled = false;
                grd_columnas.DataSource = null;
            }
        }

        private void progresBar(int prog, int total)
        {
            barra_progreso.Value = (prog * 100) / total;
        }
    }
}