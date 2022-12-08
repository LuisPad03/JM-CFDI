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
using NPOI.SS.Formula.Functions;

namespace JM_CFDI
{
    public partial class Form1 : Form
    {
        //string ConsultaSQLManagement;
        static string carpetaDocs;
        private int filasTotales;

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

                    //List<string> filtroColumnas = lstcbxColumnas.GroupBy(x => x)
                    //    .Where(g => g.Count() > 1)
                    //    .Select(x => x.Key)
                    //    .ToList();
                    List<string> filtroColumnas = lstcbxColumnas.Distinct().ToList();
                    lstcbx_columnas.DataSource = filtroColumnas;

                    List<string> filtroFiltro = cbxFiltro.Distinct().ToList();
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
            filasTotales = 0;
            bool obtienedatos = false;
            int archs = 0;
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
                        int cantidadFilas = hoja.LastRowNum;

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
                            //int cantidadFilas = hoja.LastRowNum;
                            int colFiltro = 0;

                            txt_progreso.Text = txt_progreso.Text + "--Obteniendo datos de \"" + filtro.ToUpper() + "\"\r\n";

                            colFiltro = numColumn(hoja, filtro);
                            if (colFiltro != 16385)
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

                        filasTotales = filasTotales + cantidadFilas;
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
            if (lstcbx_descripciones.CheckedItems.Count != 0)
            {
                string palabraFiltro = cbx_filtro.Text;

                List<string> lstDatosFiltro = new List<string>();                
                for (int x = 0; x < lstcbx_descripciones.CheckedItems.Count; x++)
                {
                    lstDatosFiltro.Add(lstcbx_descripciones.CheckedItems[x].ToString());
                }

                List<string> lstColumnasSelec = new List<string>();
                foreach (DataGridViewColumn column in grd_columnas.Columns)
                {
                    lstColumnasSelec.Add(column.Name);
                }

                int columnas = lstColumnasSelec.Count;
                string[,] cfdi = new string[filasTotales, columnas];

                string[] sArchivos;
                sArchivos = Directory.GetFiles(carpetaDocs);

                //int filasFinales = 0;
                var dataSet = new DataSet();
                var dataTable = dataSet.Tables.Add();

                #region Abre hoja excel a trabajar
                string rutaarchivo = sArchivos[0];

                IWorkbook MiExcel = null;
                FileStream fs = new FileStream(rutaarchivo, FileMode.Open, FileAccess.Read);

                if (Path.GetExtension(rutaarchivo) == ".xlsx")
                    MiExcel = new XSSFWorkbook(fs);
                else
                    MiExcel = new HSSFWorkbook(fs);

                ISheet hoja = MiExcel.GetSheetAt(1);
                #endregion
                
                int filas = hoja.LastRowNum;
                List<int> numColumna = new List<int>();
                for (int i = 0; i < lstColumnasSelec.Count; i++)
                {
                    numColumna.Add(numColumn(hoja, lstColumnasSelec[i]));
                }

                #region Primera forma de obtener datos
                //int colFiltro = numColumn(hoja, palabraFiltro);
                //if (colFiltro != 16385) //si existe la columna para filtrar
                //{
                //    for (int i = 0; i < filas; i++) // eje y
                //    {
                //        bool datoAgregado = false;
                //        for (int j = 0; j < columnas; j++) // eje x
                //        {
                //            if (numColumna[j] != 16385) // si existe la columna de donde obtendra datos
                //            {
                //                IRow fila = hoja.GetRow(i);
                //                string celdaFiltro = fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";

                //                for (int k = 0; k < lstDatosFiltro.Count; k++) // recorre lista de datos que se buscan dentro de la columna a filtrar
                //                {
                //                    if (celdaFiltro == lstDatosFiltro[k])
                //                    {
                //                        string celda = fila.GetCell(numColumna[j], MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(numColumna[j], MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";

                //                        cfdi[filasFinales, j] = celda;
                //                        datoAgregado = true;
                //                    }
                //                }
                //            }
                //            //else cfdi[i, j] = "";
                //        }
                //        if (datoAgregado) filasFinales++;
                //    }
                //}
                #endregion

                #region Segunda forma de obtener datos
                for (int i = 0; i < lstColumnasSelec.Count; i++)
                {
                    dataTable.Columns.Add(lstColumnasSelec[i].ToString());
                }

                int colFiltro = numColumn(hoja, palabraFiltro);
                if (colFiltro != 16385)
                {
                    for (var f = 0; f < filas; f++)
                    {
                        var row = dataTable.Rows.Add();
                        for (var c = 0; c < columnas; c++)
                        {
                            if (numColumna[c] != 16385)
                            {
                                IRow fila = hoja.GetRow(f);
                                string celdaFiltro = fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(colFiltro, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";
                                
                                string celda = fila.GetCell(numColumna[c], MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(numColumna[c], MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";

                                if (filtroExiste(lstDatosFiltro, celdaFiltro)) row[c] = celda;
                            }
                        }
                    }
                }
                #endregion

                MiExcel.Close();

                DataTable dtReporte = new DataTable();
                for (int i = 0; i < lstColumnasSelec.Count; i++)
                {
                    dtReporte.Columns.Add(lstColumnasSelec[i].ToString());
                }
                
                for (int fila = 0; fila < dataTable.Rows.Count; fila++)
                {
                    bool borrar = false;
                    for (int columna = 0; columna < lstColumnasSelec.Count; columna++)
                    {
                        string dato = dataTable.Rows[fila][lstColumnasSelec[columna]].ToString();

                        if (dato == "" || dato == null) borrar = true;
                        else borrar = false;
                    }
                    if (!borrar)
                    {
                        dtReporte.ImportRow(dataTable.Rows[fila]);
                    }
                }


                //DataView dtV = dataTable.DefaultView;
                //dtV.Sort = lstColumnasSelec[0].ToString()+ " DESC";
                //dataTable = dtV.ToTable();

                grd_columnas.DataSource = dtReporte;
            }
            else MessageBox.Show("Favor de selcionar por lo menos un elemento de la lista filtro");
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

        private void tamañoForm(int ancho, int alto)
        {
            this.MinimumSize = new System.Drawing.Size(ancho, alto);
            this.MaximumSize = new System.Drawing.Size(ancho, alto);
            this.Size = new System.Drawing.Size(ancho, alto);
        }

        private int numColumn(ISheet hoja, string filtro)
        {
            int numColumna = 16385;
            for (int i = 0; i <= 90; i++)
            {
                IRow fila = hoja.GetRow(0);
                string celda = fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK) != null ? fila.GetCell(i, MissingCellPolicy.RETURN_NULL_AND_BLANK).ToString() : "";
                if (celda == filtro)
                {
                    numColumna = i;
                    break;
                }
            }
            
            return numColumna;
        }

        private bool filtroExiste(List<string> lstDatosFiltro,string celdaFiltro)
        {
            bool filtroExiste = false;
            for (int k = 0; k < lstDatosFiltro.Count; k++)
            {
                if (celdaFiltro == lstDatosFiltro[k]) filtroExiste = true;
            }
            return filtroExiste;

            //var iFila = cfdi.GetLongLength(0);
            //var iCol = cfdi.GetLongLength(1);

            ////Fila
            //for (var f = 0; f < iFila; f++)
            //{
            //    var row = dataTable.Rows.Add();
            //    //Columna
            //    for (var c = 0; c < iCol; c++)
            //    {
            //        //if (f == 1) dataTable.Columns.Add(cfdi[0, c]);
            //        row[c] = "a";
            //    }
            //}
        }
    }
}