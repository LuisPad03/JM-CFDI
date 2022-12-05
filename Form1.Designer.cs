namespace JM_CFDI
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.txt_excel = new System.Windows.Forms.TextBox();
            this.btn_leerArchivos = new System.Windows.Forms.Button();
            this.txt_ruta = new System.Windows.Forms.TextBox();
            this.lbl_ruta = new System.Windows.Forms.Label();
            this.barra_progreso = new System.Windows.Forms.ProgressBar();
            this.panel_reporte = new System.Windows.Forms.Panel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.grd_columnas = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lstcbx_columnas = new System.Windows.Forms.CheckedListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txt_filtro = new System.Windows.Forms.TextBox();
            this.cbx_filtro = new System.Windows.Forms.ComboBox();
            this.lstcbx_descripciones = new System.Windows.Forms.CheckedListBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.lbl_excel = new System.Windows.Forms.Label();
            this.btn_reporte = new System.Windows.Forms.Button();
            this.panel_lee = new System.Windows.Forms.Panel();
            this.txt_progreso = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel_reporte.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grd_columnas)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.panel_lee.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // txt_excel
            // 
            this.txt_excel.Location = new System.Drawing.Point(34, 59);
            this.txt_excel.Name = "txt_excel";
            this.txt_excel.Size = new System.Drawing.Size(453, 23);
            this.txt_excel.TabIndex = 4;
            this.txt_excel.Click += new System.EventHandler(this.txt_excel_Click);
            // 
            // btn_leerArchivos
            // 
            this.btn_leerArchivos.Location = new System.Drawing.Point(500, 22);
            this.btn_leerArchivos.Name = "btn_leerArchivos";
            this.btn_leerArchivos.Size = new System.Drawing.Size(150, 23);
            this.btn_leerArchivos.TabIndex = 3;
            this.btn_leerArchivos.Text = "Leer Archivos";
            this.btn_leerArchivos.UseVisualStyleBackColor = true;
            this.btn_leerArchivos.Click += new System.EventHandler(this.btn_leerArchivos_Click);
            // 
            // txt_ruta
            // 
            this.txt_ruta.Cursor = System.Windows.Forms.Cursors.Hand;
            this.txt_ruta.Location = new System.Drawing.Point(34, 22);
            this.txt_ruta.Name = "txt_ruta";
            this.txt_ruta.Size = new System.Drawing.Size(453, 23);
            this.txt_ruta.TabIndex = 1;
            this.txt_ruta.Click += new System.EventHandler(this.txt_ruta_Click);
            // 
            // lbl_ruta
            // 
            this.lbl_ruta.AutoSize = true;
            this.lbl_ruta.BackColor = System.Drawing.Color.White;
            this.lbl_ruta.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lbl_ruta.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lbl_ruta.ForeColor = System.Drawing.Color.DimGray;
            this.lbl_ruta.Location = new System.Drawing.Point(50, 25);
            this.lbl_ruta.Name = "lbl_ruta";
            this.lbl_ruta.Size = new System.Drawing.Size(123, 15);
            this.lbl_ruta.TabIndex = 2;
            this.lbl_ruta.Text = "Selecionar Carpeta";
            this.lbl_ruta.Click += new System.EventHandler(this.txt_ruta_Click);
            // 
            // barra_progreso
            // 
            this.barra_progreso.Location = new System.Drawing.Point(8, 97);
            this.barra_progreso.Name = "barra_progreso";
            this.barra_progreso.Size = new System.Drawing.Size(642, 10);
            this.barra_progreso.TabIndex = 7;
            // 
            // panel_reporte
            // 
            this.panel_reporte.Controls.Add(this.groupBox3);
            this.panel_reporte.Controls.Add(this.groupBox2);
            this.panel_reporte.Controls.Add(this.groupBox1);
            this.panel_reporte.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel_reporte.Location = new System.Drawing.Point(0, 126);
            this.panel_reporte.Name = "panel_reporte";
            this.panel_reporte.Size = new System.Drawing.Size(949, 265);
            this.panel_reporte.TabIndex = 2;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.grd_columnas);
            this.groupBox3.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.groupBox3.Location = new System.Drawing.Point(422, 6);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(518, 248);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Vista previa:";
            // 
            // grd_columnas
            // 
            this.grd_columnas.AllowUserToAddRows = false;
            this.grd_columnas.AllowUserToDeleteRows = false;
            this.grd_columnas.AllowUserToOrderColumns = true;
            this.grd_columnas.AllowUserToResizeRows = false;
            this.grd_columnas.BackgroundColor = System.Drawing.SystemColors.Control;
            this.grd_columnas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.grd_columnas.Location = new System.Drawing.Point(6, 22);
            this.grd_columnas.Name = "grd_columnas";
            this.grd_columnas.ReadOnly = true;
            this.grd_columnas.RowHeadersVisible = false;
            this.grd_columnas.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToFirstHeader;
            this.grd_columnas.RowTemplate.Height = 25;
            this.grd_columnas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.grd_columnas.Size = new System.Drawing.Size(506, 217);
            this.grd_columnas.TabIndex = 12;
            this.grd_columnas.TabStop = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lstcbx_columnas);
            this.groupBox2.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.groupBox2.Location = new System.Drawing.Point(6, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 250);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Agregar Columna:";
            // 
            // lstcbx_columnas
            // 
            this.lstcbx_columnas.CheckOnClick = true;
            this.lstcbx_columnas.ForeColor = System.Drawing.Color.Black;
            this.lstcbx_columnas.FormattingEnabled = true;
            this.lstcbx_columnas.HorizontalScrollbar = true;
            this.lstcbx_columnas.Location = new System.Drawing.Point(5, 21);
            this.lstcbx_columnas.Name = "lstcbx_columnas";
            this.lstcbx_columnas.Size = new System.Drawing.Size(189, 220);
            this.lstcbx_columnas.TabIndex = 8;
            this.lstcbx_columnas.SelectedValueChanged += new System.EventHandler(this.lstcbx_columnas_SelectedValueChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txt_filtro);
            this.groupBox1.Controls.Add(this.cbx_filtro);
            this.groupBox1.Controls.Add(this.lstcbx_descripciones);
            this.groupBox1.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.groupBox1.Location = new System.Drawing.Point(215, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 250);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Filtro:";
            // 
            // txt_filtro
            // 
            this.txt_filtro.Location = new System.Drawing.Point(6, 45);
            this.txt_filtro.Name = "txt_filtro";
            this.txt_filtro.Size = new System.Drawing.Size(188, 23);
            this.txt_filtro.TabIndex = 10;
            this.txt_filtro.TextChanged += new System.EventHandler(this.txt_filtro_TextChanged);
            // 
            // cbx_filtro
            // 
            this.cbx_filtro.DropDownHeight = 100;
            this.cbx_filtro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbx_filtro.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.cbx_filtro.FormattingEnabled = true;
            this.cbx_filtro.IntegralHeight = false;
            this.cbx_filtro.ItemHeight = 15;
            this.cbx_filtro.Location = new System.Drawing.Point(6, 16);
            this.cbx_filtro.Name = "cbx_filtro";
            this.cbx_filtro.Size = new System.Drawing.Size(188, 23);
            this.cbx_filtro.TabIndex = 9;
            this.cbx_filtro.SelectionChangeCommitted += new System.EventHandler(this.cbx_filtro_SelectionChangeCommitted);
            // 
            // lstcbx_descripciones
            // 
            this.lstcbx_descripciones.CheckOnClick = true;
            this.lstcbx_descripciones.ForeColor = System.Drawing.Color.Black;
            this.lstcbx_descripciones.FormattingEnabled = true;
            this.lstcbx_descripciones.HorizontalScrollbar = true;
            this.lstcbx_descripciones.Location = new System.Drawing.Point(6, 75);
            this.lstcbx_descripciones.Name = "lstcbx_descripciones";
            this.lstcbx_descripciones.Size = new System.Drawing.Size(188, 166);
            this.lstcbx_descripciones.Sorted = true;
            this.lstcbx_descripciones.TabIndex = 11;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::JM_CFDI.Properties.Resources.excel;
            this.pictureBox2.Location = new System.Drawing.Point(8, 59);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(22, 21);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox2.TabIndex = 6;
            this.pictureBox2.TabStop = false;
            // 
            // lbl_excel
            // 
            this.lbl_excel.AutoSize = true;
            this.lbl_excel.BackColor = System.Drawing.Color.White;
            this.lbl_excel.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lbl_excel.Font = new System.Drawing.Font("Constantia", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.lbl_excel.ForeColor = System.Drawing.Color.DimGray;
            this.lbl_excel.Location = new System.Drawing.Point(50, 62);
            this.lbl_excel.Name = "lbl_excel";
            this.lbl_excel.Size = new System.Drawing.Size(122, 15);
            this.lbl_excel.TabIndex = 5;
            this.lbl_excel.Text = "Selecionar Archivo";
            this.lbl_excel.Click += new System.EventHandler(this.txt_excel_Click);
            // 
            // btn_reporte
            // 
            this.btn_reporte.Location = new System.Drawing.Point(500, 59);
            this.btn_reporte.Name = "btn_reporte";
            this.btn_reporte.Size = new System.Drawing.Size(150, 23);
            this.btn_reporte.TabIndex = 0;
            this.btn_reporte.Text = "Genera Reporte";
            this.btn_reporte.UseVisualStyleBackColor = true;
            this.btn_reporte.Click += new System.EventHandler(this.btn_reporte_Click);
            // 
            // panel_lee
            // 
            this.panel_lee.BackColor = System.Drawing.SystemColors.Control;
            this.panel_lee.Controls.Add(this.txt_progreso);
            this.panel_lee.Controls.Add(this.pictureBox1);
            this.panel_lee.Controls.Add(this.lbl_ruta);
            this.panel_lee.Controls.Add(this.txt_ruta);
            this.panel_lee.Controls.Add(this.pictureBox2);
            this.panel_lee.Controls.Add(this.barra_progreso);
            this.panel_lee.Controls.Add(this.btn_leerArchivos);
            this.panel_lee.Controls.Add(this.lbl_excel);
            this.panel_lee.Controls.Add(this.btn_reporte);
            this.panel_lee.Controls.Add(this.txt_excel);
            this.panel_lee.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_lee.Location = new System.Drawing.Point(0, 0);
            this.panel_lee.Name = "panel_lee";
            this.panel_lee.Size = new System.Drawing.Size(949, 124);
            this.panel_lee.TabIndex = 1;
            // 
            // txt_progreso
            // 
            this.txt_progreso.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txt_progreso.Location = new System.Drawing.Point(660, 10);
            this.txt_progreso.Multiline = true;
            this.txt_progreso.Name = "txt_progreso";
            this.txt_progreso.ReadOnly = true;
            this.txt_progreso.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt_progreso.Size = new System.Drawing.Size(280, 104);
            this.txt_progreso.TabIndex = 6;
            this.txt_progreso.TabStop = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::JM_CFDI.Properties.Resources.carpeta__1_;
            this.pictureBox1.Location = new System.Drawing.Point(8, 22);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(22, 21);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(949, 391);
            this.Controls.Add(this.panel_lee);
            this.Controls.Add(this.panel_reporte);
            this.MaximizeBox = false;
            this.MinimumSize = new System.Drawing.Size(714, 130);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "JM-CFDI";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel_reporte.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.grd_columnas)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.panel_lee.ResumeLayout(false);
            this.panel_lee.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private FolderBrowserDialog folderBrowserDialog1;
        private TextBox txt_excel;
        private Button btn_leerArchivos;
        private TextBox txt_ruta;
        private Label lbl_ruta;
        private ProgressBar barra_progreso;
        private Panel panel_reporte;
        private Panel panel_lee;
        private CheckedListBox lstcbx_descripciones;
        private Button btn_reporte;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private DataGridView grd_columnas;
        private PictureBox pictureBox1;
        private OpenFileDialog openFileDialog1;
        private PictureBox pictureBox2;
        private Label lbl_excel;
        private CheckedListBox lstcbx_columnas;
        private GroupBox groupBox3;
        private ComboBox cbx_filtro;
        private TextBox txt_filtro;
        private TextBox txt_progreso;
    }
}