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
using System.Collections;

namespace JM_CFDI
{
    public partial class Form1 : Form
    {
        //string ConsultaSQLManagement;
        static string carpetaDocs;
        static string carpetaRep;
        private int filasTotales;
        private bool reporte = true;
        private XSSFWorkbook workbook = new XSSFWorkbook();
        private FileStream file;
        private bool primerHoja = true;
        private List<string> listFiltro = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel_reporte.Visible = false;
            tamañoForm(670, 160);

            txt_ruta.PlaceholderText = "Selecionar Carpeta CFDI's";
            btn_reporte.Enabled = false;
            btn_genera_excel.Enabled = false;
            txt_excel.Size = new Size(453, 23);
            txt_excel.PlaceholderText = "Selecionar Carpeta para guardar Reporte";
            txt_nombre_archivo.Visible = false;
            txt_nombre_archivo.PlaceholderText = "Nombre de la HOJA";

            txt_filtro.Enabled = false;
            cbx_selecciona_todo.Enabled = false;
            lstcbx_descripciones.Enabled = false;

        }

        private void txt_ruta_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) txt_ruta.Text = folderBrowserDialog1.SelectedPath;
        }
        private void btn_leerArchivos_Click(object sender, EventArgs e)
        {
            lstcbx_columnas.DataSource = null;
            cbx_filtro.DataSource = null;
            txt_filtro.Text = "";
            txt_filtro.Enabled = false;
            cbx_selecciona_todo.Checked = false;
            cbx_selecciona_todo.Enabled = false;
            lstcbx_descripciones.DataSource = null;
            grd_columnas.DataSource = null;

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
                    cbx_filtro.DataSource = filtroFiltro;

                }
                else MessageBox.Show("Los documentos contenidos en esta carpeta no son compatibles. \n\nLos documentos deben de contar con la extencion \".XLSX\", es decir archivos tipo excel.", "Documentos no compatibles", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Ingresa la carpeta que contenga los CFDI, en formato excel.", "Información incompleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

                for (int i = 0; i < 9; i++)
                {
                    vistaPrevia.Rows.Add();
                }
                grd_columnas.DataSource = vistaPrevia;

                if (lstcbx_descripciones.Items.Count != 0 && lstcbx_columnas.CheckedItems.Count != 0)
                {
                    btn_reporte.Enabled = true;
                    txt_nombre_archivo.Enabled = true;
                    txt_nombre_archivo.Visible = true;
                    txt_excel.Size = new Size(291, 23);
                }
            }
            else
            {
                btn_reporte.Enabled = false;
                grd_columnas.DataSource = null;
                txt_nombre_archivo.Enabled = false;
            }
        }

        private void cbx_filtro_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (carpetaDocs != "")
            {
                cbx_selecciona_todo.Checked = false;
                cbx_selecciona_todo.Enabled = true;
                lstcbx_descripciones.DataSource = null;
                string[] sArchivos;
                sArchivos = Directory.GetFiles(carpetaDocs);

                List<string> lstFiltro = new List<string>();
                List<string> filtroOK = new List<string>();

                if (obtieneDatos(sArchivos, lstFiltro, filtroOK, "filtro"))
                {
                    filtroOK = lstFiltro.Distinct().ToList();
                    lstcbx_descripciones.DataSource = filtroOK;
                    listFiltro = filtroOK;
                }

                txt_filtro.Enabled = true;
                txt_filtro.Text = "";
                lstcbx_descripciones.Enabled = true;

                if (lstcbx_descripciones.Items.Count != 0 && lstcbx_columnas.CheckedItems.Count != 0)
                {
                    btn_reporte.Enabled = true;
                    txt_nombre_archivo.Enabled = true;
                    txt_nombre_archivo.Visible = true;
                    txt_excel.Size = new Size(291, 23);
                }
            }
            else MessageBox.Show("Ingresa la carpeta que contenga los CFDI, en formato excel.", "Información incompleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        }
        private void cbx_selecciona_todo_CheckedChanged(object sender, EventArgs e)
        {
            bool activo = cbx_selecciona_todo.Checked;

            for (int i = 0; i < lstcbx_descripciones.Items.Count; i++)
            {
                lstcbx_descripciones.SetItemChecked(i, activo);
            }
        }
        private void txt_filtro_TextChanged(object sender, EventArgs e)
       {
            List<string> filtro = new List<string>();
            if (txt_filtro.Text != "")
            {
                List<string> list = listFiltro.Where(x => x.Contains(txt_filtro.Text)).ToList();
                


                lstcbx_descripciones.DataSource = list;
            }
            else lstcbx_descripciones.DataSource = listFiltro;

        }

        private void txt_excel_Click(object sender, EventArgs e)
        {
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    txt_excel.Text = openFileDialog1.FileName;
            //    lbl_excel.Visible = false;
            //}
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK) txt_excel.Text = folderBrowserDialog1.SelectedPath;

        }
        private void btn_reporte_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (sender as System.Windows.Forms.Button);
            string tipobtn = btn.Text;
            string rutaNuevoExcel = txt_excel.Text + "\\CFDI VS ANUAL.xlsx";

            if (tipobtn == "Agrega Hoja al Reporte")
            {
                if (txt_nombre_archivo.Text.Trim() == "" || txt_nombre_archivo.Text.Trim() == null || txt_excel.Text.Trim() == "" || txt_excel.Text.Trim() == null)
                {
                    MessageBox.Show("Favor de agregar ruta para guardar Reporte y escribir el nombre de la hoja");
                }
                else
                {
                    if (lstcbx_descripciones.CheckedItems.Count != 0)
                    {
                        int archs = 0;
                        txt_progreso.Text = "";
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

                        #region Obtiene datos
                        //int filasFinales = 0;
                        var dataSet = new DataSet();
                        var dataTable = dataSet.Tables.Add();
                        for (int i = 0; i < lstColumnasSelec.Count; i++)
                        {
                            dataTable.Columns.Add(lstColumnasSelec[i].ToString());
                        }

                        foreach (string rutaarchivo in sArchivos)
                        {
                            txt_progreso.Text = txt_progreso.Text + "Leyendo archivo: " + Path.GetFileName(rutaarchivo) + "\r\n";

                            #region Abre hoja excel a trabajar
                            //string rutaarchivo = sArchivos[0];

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
                            //for (int i = 0; i < lstColumnasSelec.Count; i++)
                            //{
                            //    dataTable.Columns.Add(lstColumnasSelec[i].ToString());
                            //}                            
                            int colFiltro = numColumn(hoja, palabraFiltro);
                            if (colFiltro != 16385)
                            {
                                txt_progreso.Text = txt_progreso.Text + "--Buscando datos en: " + palabraFiltro.ToUpper() + "\r\n";
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
                            else txt_progreso.Text = txt_progreso.Text + "--No existe: " + palabraFiltro.ToUpper() + "\r\n";
                            #endregion

                            MiExcel.Close();

                            archs++;
                            progresBar(archs, sArchivos.Length);
                        }

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
                        #endregion


                        //string rutaNuevoExcel = txt_excel.Text + "\\" + txt_nombre_archivo.Text + ".xlsx";
                        //GridToExcelByNPOI(dtReporte, rutaNuevoExcel, lstColumnasSelec);

                        txt_progreso.Text = txt_progreso.Text + "---> Generando reporte en: " + rutaNuevoExcel;

                        if (!reporte)
                        {
                            if (txt_nombre_archivo.Text != "")
                            {
                                XSSFWorkbook workbook = new XSSFWorkbook();
                                FileStream file = new FileStream(rutaNuevoExcel, FileMode.OpenOrCreate);
                                reporte = true; 
                                primerHoja = true;
                                GridToExcelByNPOI(dtReporte, file, lstColumnasSelec, workbook, txt_nombre_archivo.Text, false, primerHoja);
                                txt_nombre_archivo.Text = "";
                            }
                            else MessageBox.Show("Favor de cambiar el nombre de la hoja");
                        }
                        else
                        {
                            if (txt_nombre_archivo.Text != "")
                            {
                                file = new FileStream(rutaNuevoExcel, FileMode.OpenOrCreate);
                                GridToExcelByNPOI(dtReporte, file, lstColumnasSelec, workbook, txt_nombre_archivo.Text, false, primerHoja);
                                txt_nombre_archivo.Text = ""; 
                                primerHoja = false;
                            }
                            else MessageBox.Show("Favor de cambiar el nombre de la hoja");
                        }


                        grd_columnas.DataSource = dtReporte;
                        btn_genera_excel.Enabled = true;
                    }
                    else MessageBox.Show("Favor de selcionar por lo menos un elemento de la lista filtro");
                }
            }
            else
            {
                DataTable dtReporte = new DataTable();
                List<string> lstColumnasSelec = new List<string>();
                file = new FileStream(rutaNuevoExcel, FileMode.OpenOrCreate);
                GridToExcelByNPOI(dtReporte, file, lstColumnasSelec, workbook, "", true, primerHoja);
                reporte = false;
                lstcbx_columnas.DataSource = null;
                cbx_filtro.DataSource= null;
                txt_filtro.Text = "";
                lstcbx_descripciones.DataSource = null;
                grd_columnas.DataSource = null;
                txt_progreso.Text = "Reporte Guardado en: \r\n" + rutaNuevoExcel;
            }

        }
        private bool filtroExiste(List<string> lstDatosFiltro, string celdaFiltro)
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
        
        private static void GridToExcelByNPOI(DataTable dt, FileStream file, List<string> lstColumnasSelec, XSSFWorkbook workbook, string nombrehoja, bool finReporte, bool primerHoja)
        {            
            try
            {                
                if (!finReporte)
                {
                    if (primerHoja)
                    {
                        ISheet sheet1 = workbook.CreateSheet("CFDI VS ANUAL");
                        
                        int[] filita = new int[] { 2, 4, 6};
                        string[] valor = new string[] { "Concentardo de gastos declaracion anual vs cfdi", nombrehoja, "ANUAL" };

                        for (int i = 0; i < 3; i++)
                        {
                            IRow headerRow1 = sheet1.CreateRow(filita[i]);

                            if (filita[i] == 6)
                            {
                                ICell cell2 = headerRow1.CreateCell(1);
                                cell2.SetCellValue("CFDI");
                                ICell cell3 = headerRow1.CreateCell(2);
                                cell3.SetCellValue("DIFERENCIA");
                            }

                            ICell cell1 = headerRow1.CreateCell(0);
                            cell1.SetCellValue(valor[i]);
                        }
                    }
                    else
                    {
                        ISheet hoja1 = workbook.GetSheetAt(0);
                        int cantidadFilas = hoja1.LastRowNum;

                        int[] filit = new int[] { cantidadFilas + 3, cantidadFilas + 5 };
                        string[] valo = new string[] { nombrehoja, "ANUAL" };

                        for (int i = 0; i < 2; i++)
                        {
                            IRow headerRow1 = hoja1.CreateRow(filit[i]);

                            if (filit[i] == cantidadFilas + 5)
                            {
                                ICell cell2 = headerRow1.CreateCell(1);
                                cell2.SetCellValue("CFDI");
                                ICell cell3 = headerRow1.CreateCell(2);
                                cell3.SetCellValue("DIFERENCIA");
                            }

                            ICell cell1 = headerRow1.CreateCell(0);
                            cell1.SetCellValue(valo[i]);
                        }
                    }

                    ISheet sheet = workbook.CreateSheet(nombrehoja);
                    ICellStyle HeadercellStyle = workbook.CreateCellStyle();

                    #region HeadercellStyle
                    HeadercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
                    headerfont.Boldweight = (short)FontBoldWeight.Bold;
                    HeadercellStyle.SetFont(headerfont);
                    #endregion

                    //As the column name with [Name,number, value, date]
                    int icolIndex = 0;
                    IRow headerRow = sheet.CreateRow(0);

                    int icolList = 0;
                    foreach (DataColumn item in dt.Columns)
                    {
                        ICell cell = headerRow.CreateCell(icolIndex);
                        cell.SetCellValue(lstColumnasSelec[icolList]);
                        cell.CellStyle = HeadercellStyle;
                        icolIndex++;
                        icolList++;
                    }

                    ICellStyle cellStyle = workbook.CreateCellStyle();

                    #region cellStyle
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                    NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
                    cellfont.Boldweight = (short)FontBoldWeight.Normal;
                    cellStyle.SetFont(cellfont);
                    #endregion

                    int iRowIndex = 1;
                    int iCellIndex = 0;

                    foreach (DataRow Rowitem in dt.Rows)
                    {
                        IRow DataRow = sheet.CreateRow(iRowIndex);
                        foreach (DataColumn Colitem in dt.Columns)
                        {
                            ICell cell = DataRow.CreateCell(iCellIndex);
                            cell.SetCellValue(Rowitem[Colitem].ToString());
                            cell.CellStyle = cellStyle;
                            iCellIndex++;
                        }
                        iCellIndex = 0;
                        iRowIndex++;
                    }

                    for (int i = 0; i < icolIndex; i++)
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    //FileStream file = new FileStream(strExcelFileName, FileMode.OpenOrCreate);
                    workbook.Write(file);
                }
                else
                {
                    file.Flush();
                    file.Close();
                    //limpiar todos los controles
                }
            }
            catch (Exception ex)
            {
            }
            finally { workbook = null; }
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
                
        private void progresBar(int prog, int total)
        {
            barra_progreso.Value = (prog * 100) / total;
        }

        private void tamañoForm(int ancho, int alto)
        {
            this.MinimumSize = new System.Drawing.Size(ancho, alto);
            //this.MaximumSize = new System.Drawing.Size(ancho, alto);
            //this.Size = new System.Drawing.Size(ancho, alto);
        }

        private static void GridToExcelByNPOI(DataTable dt, string strExcelFileName, List<string> lstColumnasSelec)
        {
            if (!File.Exists(strExcelFileName))
            {
                XSSFWorkbook workbook = new XSSFWorkbook();
                try
                {
                    ISheet sheet = workbook.CreateSheet("Sheet1");
                    ICellStyle HeadercellStyle = workbook.CreateCellStyle();

                    #region datatable concepto

                    HeadercellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
                    HeadercellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                    NPOI.SS.UserModel.IFont headerfont = workbook.CreateFont();
                    headerfont.Boldweight = (short)FontBoldWeight.Bold;
                    HeadercellStyle.SetFont(headerfont);


                    //As the column name with [Name,number, value, date]
                    int icolIndex = 0;
                    IRow headerRow = sheet.CreateRow(0);

                    int icolList = 0;
                    foreach (DataColumn item in dt.Columns)
                    {
                        ICell cell = headerRow.CreateCell(icolIndex);
                        cell.SetCellValue(lstColumnasSelec[icolList]);
                        cell.CellStyle = HeadercellStyle;
                        icolIndex++;
                        icolList++;
                    }

                    ICellStyle cellStyle = workbook.CreateCellStyle();

                    //Excel date format in order to avoid being automatically replaced, so the format is set to "@" represents a rate of view as text
                    cellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
                    cellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
                    cellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;

                    NPOI.SS.UserModel.IFont cellfont = workbook.CreateFont();
                    cellfont.Boldweight = (short)FontBoldWeight.Normal;
                    cellStyle.SetFont(cellfont);

                    int iRowIndex = 1;
                    int iCellIndex = 0;

                    foreach (DataRow Rowitem in dt.Rows)
                    {
                        IRow DataRow = sheet.CreateRow(iRowIndex);
                        foreach (DataColumn Colitem in dt.Columns)
                        {
                            ICell cell = DataRow.CreateCell(iCellIndex);
                            cell.SetCellValue(Rowitem[Colitem].ToString());
                            //cell.CellStyle = cellStyle;
                            iCellIndex++;
                        }
                        iCellIndex = 0;
                        iRowIndex++;
                    }

                    for (int i = 0; i < icolIndex; i++)
                    {
                        sheet.AutoSizeColumn(i);
                    }

                    FileStream file = new FileStream(strExcelFileName, FileMode.OpenOrCreate);
                    workbook.Write(file);

                    #endregion

                    file.Flush();
                    file.Close();

                }
                catch (Exception ex)
                {
                }
                finally { workbook = null; }

            }
            else MessageBox.Show("El documento ya existe, favor de cambiar el nombre del archivo");
        }
    }
}