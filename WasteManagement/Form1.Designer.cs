using System.Windows.Forms;

namespace WasteManagement
{
    partial class Form1
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.Report = new System.Windows.Forms.TabPage();
            this.lbl_ScrapTGT = new System.Windows.Forms.Label();
            this.lblheaderScrapTgt = new System.Windows.Forms.Label();
            this.lbl_PrecentTotalScrap = new System.Windows.Forms.Label();
            this.lblheaderScrapPrecent = new System.Windows.Forms.Label();
            this.lbl_EntryToWarehouse = new System.Windows.Forms.Label();
            this.lblheaderEntryToWarehouse = new System.Windows.Forms.Label();
            this.lbl_totalScrap = new System.Windows.Forms.Label();
            this.lblheaderTotalScrap = new System.Windows.Forms.Label();
            this.lbl_totalCured = new System.Windows.Forms.Label();
            this.lblTotalCured = new System.Windows.Forms.Label();
            this.DgvWasteMonth = new ADGV.AdvancedDataGridView();
            this.WasteUpdate = new System.Windows.Forms.TabPage();
            this.label13 = new System.Windows.Forms.Label();
            this.DataGrid_Desc = new System.Windows.Forms.DataGridView();
            this.btn_SaveData = new System.Windows.Forms.Button();
            this.lbl_totalLeftover = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.txt_AmountSpesific = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.btn_confirm = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_CodeWaste = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.cbo_CodeWaste = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lbl_dateWantedUpdate = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.lbl_AmountWaste = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.lbl_WasteUpdate = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Cbo_MonthReport = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Cbo_YearReport = new System.Windows.Forms.ComboBox();
            this.btn_show = new System.Windows.Forms.Button();
            this.lbl_Report = new System.Windows.Forms.Label();
            this.lbl_MonthReport = new System.Windows.Forms.Label();
            this.MakeExcellRepBTN = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.lbl_LastUpdate = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.Report.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvWasteMonth)).BeginInit();
            this.WasteUpdate.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGrid_Desc)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.Report);
            this.tabControl1.Controls.Add(this.WasteUpdate);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.tabControl1.Location = new System.Drawing.Point(12, 103);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1249, 870);
            this.tabControl1.TabIndex = 0;
            // 
            // Report
            // 
            this.Report.Controls.Add(this.lbl_ScrapTGT);
            this.Report.Controls.Add(this.lblheaderScrapTgt);
            this.Report.Controls.Add(this.lbl_PrecentTotalScrap);
            this.Report.Controls.Add(this.lblheaderScrapPrecent);
            this.Report.Controls.Add(this.lbl_EntryToWarehouse);
            this.Report.Controls.Add(this.lblheaderEntryToWarehouse);
            this.Report.Controls.Add(this.lbl_totalScrap);
            this.Report.Controls.Add(this.lblheaderTotalScrap);
            this.Report.Controls.Add(this.lbl_totalCured);
            this.Report.Controls.Add(this.lblTotalCured);
            this.Report.Controls.Add(this.DgvWasteMonth);
            this.Report.Location = new System.Drawing.Point(4, 29);
            this.Report.Name = "Report";
            this.Report.Padding = new System.Windows.Forms.Padding(3);
            this.Report.Size = new System.Drawing.Size(1241, 837);
            this.Report.TabIndex = 0;
            this.Report.Text = "דו\"ח פסולות";
            this.Report.UseVisualStyleBackColor = true;
            // 
            // lbl_ScrapTGT
            // 
            this.lbl_ScrapTGT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_ScrapTGT.AutoSize = true;
            this.lbl_ScrapTGT.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_ScrapTGT.Location = new System.Drawing.Point(1099, 746);
            this.lbl_ScrapTGT.Name = "lbl_ScrapTGT";
            this.lbl_ScrapTGT.Size = new System.Drawing.Size(136, 32);
            this.lbl_ScrapTGT.TabIndex = 10;
            this.lbl_ScrapTGT.Text = "***********";
            // 
            // lblheaderScrapTgt
            // 
            this.lblheaderScrapTgt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblheaderScrapTgt.AutoSize = true;
            this.lblheaderScrapTgt.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblheaderScrapTgt.Location = new System.Drawing.Point(657, 746);
            this.lblheaderScrapTgt.Name = "lblheaderScrapTgt";
            this.lblheaderScrapTgt.Size = new System.Drawing.Size(232, 32);
            this.lblheaderScrapTgt.TabIndex = 9;
            this.lblheaderScrapTgt.Text = ": Scrap Target %";
            // 
            // lbl_PrecentTotalScrap
            // 
            this.lbl_PrecentTotalScrap.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_PrecentTotalScrap.AutoSize = true;
            this.lbl_PrecentTotalScrap.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_PrecentTotalScrap.Location = new System.Drawing.Point(1099, 801);
            this.lbl_PrecentTotalScrap.Name = "lbl_PrecentTotalScrap";
            this.lbl_PrecentTotalScrap.Size = new System.Drawing.Size(136, 32);
            this.lbl_PrecentTotalScrap.TabIndex = 8;
            this.lbl_PrecentTotalScrap.Text = "***********";
            // 
            // lblheaderScrapPrecent
            // 
            this.lblheaderScrapPrecent.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblheaderScrapPrecent.AutoSize = true;
            this.lblheaderScrapPrecent.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblheaderScrapPrecent.Location = new System.Drawing.Point(657, 792);
            this.lblheaderScrapPrecent.Name = "lblheaderScrapPrecent";
            this.lblheaderScrapPrecent.Size = new System.Drawing.Size(214, 32);
            this.lblheaderScrapPrecent.TabIndex = 7;
            this.lblheaderScrapPrecent.Text = ": Total Scrap %";
            // 
            // lbl_EntryToWarehouse
            // 
            this.lbl_EntryToWarehouse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_EntryToWarehouse.AutoSize = true;
            this.lbl_EntryToWarehouse.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_EntryToWarehouse.Location = new System.Drawing.Point(1099, 697);
            this.lbl_EntryToWarehouse.Name = "lbl_EntryToWarehouse";
            this.lbl_EntryToWarehouse.Size = new System.Drawing.Size(136, 32);
            this.lbl_EntryToWarehouse.TabIndex = 6;
            this.lbl_EntryToWarehouse.Text = "***********";
            // 
            // lblheaderEntryToWarehouse
            // 
            this.lblheaderEntryToWarehouse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblheaderEntryToWarehouse.AutoSize = true;
            this.lblheaderEntryToWarehouse.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblheaderEntryToWarehouse.Location = new System.Drawing.Point(657, 697);
            this.lblheaderEntryToWarehouse.Name = "lblheaderEntryToWarehouse";
            this.lblheaderEntryToWarehouse.Size = new System.Drawing.Size(291, 32);
            this.lblheaderEntryToWarehouse.TabIndex = 5;
            this.lblheaderEntryToWarehouse.Text = ": Entry to Warehouse";
            // 
            // lbl_totalScrap
            // 
            this.lbl_totalScrap.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_totalScrap.AutoSize = true;
            this.lbl_totalScrap.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_totalScrap.Location = new System.Drawing.Point(1099, 647);
            this.lbl_totalScrap.Name = "lbl_totalScrap";
            this.lbl_totalScrap.Size = new System.Drawing.Size(136, 32);
            this.lbl_totalScrap.TabIndex = 4;
            this.lbl_totalScrap.Text = "***********";
            // 
            // lblheaderTotalScrap
            // 
            this.lblheaderTotalScrap.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblheaderTotalScrap.AutoSize = true;
            this.lblheaderTotalScrap.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblheaderTotalScrap.Location = new System.Drawing.Point(657, 647);
            this.lblheaderTotalScrap.Name = "lblheaderTotalScrap";
            this.lblheaderTotalScrap.Size = new System.Drawing.Size(174, 32);
            this.lblheaderTotalScrap.TabIndex = 3;
            this.lblheaderTotalScrap.Text = ":Total Scrap";
            // 
            // lbl_totalCured
            // 
            this.lbl_totalCured.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_totalCured.AutoSize = true;
            this.lbl_totalCured.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_totalCured.Location = new System.Drawing.Point(1099, 598);
            this.lbl_totalCured.Name = "lbl_totalCured";
            this.lbl_totalCured.Size = new System.Drawing.Size(136, 32);
            this.lbl_totalCured.TabIndex = 2;
            this.lbl_totalCured.Text = "***********";
            // 
            // lblTotalCured
            // 
            this.lblTotalCured.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTotalCured.AutoSize = true;
            this.lblTotalCured.Font = new System.Drawing.Font("Arial", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lblTotalCured.Location = new System.Drawing.Point(657, 598);
            this.lblTotalCured.Name = "lblTotalCured";
            this.lblTotalCured.Size = new System.Drawing.Size(219, 32);
            this.lblTotalCured.TabIndex = 1;
            this.lblTotalCured.Text = ":Total Cured-kg";
            // 
            // DgvWasteMonth
            // 
            this.DgvWasteMonth.AllowUserToAddRows = false;
            this.DgvWasteMonth.AllowUserToDeleteRows = false;
            this.DgvWasteMonth.AllowUserToResizeRows = false;
            this.DgvWasteMonth.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DgvWasteMonth.AutoGenerateContextFilters = true;
            this.DgvWasteMonth.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.DgvWasteMonth.BackgroundColor = System.Drawing.Color.LightGray;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DgvWasteMonth.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle7;
            this.DgvWasteMonth.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DgvWasteMonth.Cursor = System.Windows.Forms.Cursors.Hand;
            this.DgvWasteMonth.DateWithTime = false;
            dataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle8.BackColor = System.Drawing.Color.White;
            dataGridViewCellStyle8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            dataGridViewCellStyle8.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.DgvWasteMonth.DefaultCellStyle = dataGridViewCellStyle8;
            this.DgvWasteMonth.Location = new System.Drawing.Point(6, 3);
            this.DgvWasteMonth.Name = "DgvWasteMonth";
            this.DgvWasteMonth.RightToLeft = System.Windows.Forms.RightToLeft.No;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.DgvWasteMonth.RowsDefaultCellStyle = dataGridViewCellStyle9;
            this.DgvWasteMonth.RowTemplate.Height = 40;
            this.DgvWasteMonth.RowTemplate.ReadOnly = true;
            this.DgvWasteMonth.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.DgvWasteMonth.Size = new System.Drawing.Size(1223, 583);
            this.DgvWasteMonth.TabIndex = 0;
            this.DgvWasteMonth.TimeFilter = false;
            this.DgvWasteMonth.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DgvWasteMonth_CellContentDoubleClick);
            this.DgvWasteMonth.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.DgvWasteMonth_CellPainting);
            // 
            // WasteUpdate
            // 
            this.WasteUpdate.BackColor = System.Drawing.Color.White;
            this.WasteUpdate.Controls.Add(this.label13);
            this.WasteUpdate.Controls.Add(this.DataGrid_Desc);
            this.WasteUpdate.Controls.Add(this.btn_SaveData);
            this.WasteUpdate.Controls.Add(this.lbl_totalLeftover);
            this.WasteUpdate.Controls.Add(this.label12);
            this.WasteUpdate.Controls.Add(this.label11);
            this.WasteUpdate.Controls.Add(this.txt_AmountSpesific);
            this.WasteUpdate.Controls.Add(this.label10);
            this.WasteUpdate.Controls.Add(this.btn_confirm);
            this.WasteUpdate.Controls.Add(this.label9);
            this.WasteUpdate.Controls.Add(this.txt_CodeWaste);
            this.WasteUpdate.Controls.Add(this.label8);
            this.WasteUpdate.Controls.Add(this.cbo_CodeWaste);
            this.WasteUpdate.Controls.Add(this.label5);
            this.WasteUpdate.Controls.Add(this.lbl_dateWantedUpdate);
            this.WasteUpdate.Controls.Add(this.label7);
            this.WasteUpdate.Controls.Add(this.lbl_AmountWaste);
            this.WasteUpdate.Controls.Add(this.label6);
            this.WasteUpdate.Controls.Add(this.lbl_WasteUpdate);
            this.WasteUpdate.Controls.Add(this.label4);
            this.WasteUpdate.Location = new System.Drawing.Point(4, 29);
            this.WasteUpdate.Name = "WasteUpdate";
            this.WasteUpdate.Padding = new System.Windows.Forms.Padding(3);
            this.WasteUpdate.Size = new System.Drawing.Size(1241, 837);
            this.WasteUpdate.TabIndex = 1;
            this.WasteUpdate.Text = "עדכון פסולות";
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label13.Location = new System.Drawing.Point(-17, 98);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(6836, 73);
            this.label13.TabIndex = 383;
            this.label13.Text = "_________________________________________________________________________________" +
    "________________________________________________________________________________" +
    "____________________________";
            this.label13.Visible = false;
            // 
            // DataGrid_Desc
            // 
            this.DataGrid_Desc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.DataGrid_Desc.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.DataGrid_Desc.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGrid_Desc.Location = new System.Drawing.Point(446, 467);
            this.DataGrid_Desc.Name = "DataGrid_Desc";
            this.DataGrid_Desc.Size = new System.Drawing.Size(444, 276);
            this.DataGrid_Desc.TabIndex = 382;
            this.DataGrid_Desc.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.Data_Desc_CellClick);
            this.DataGrid_Desc.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.Data_Desc_CellEndEdit);
            this.DataGrid_Desc.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.DataGrid_Desc_UserDeletingRow);
            // 
            // btn_SaveData
            // 
            this.btn_SaveData.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btn_SaveData.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btn_SaveData.Enabled = false;
            this.btn_SaveData.Location = new System.Drawing.Point(520, 791);
            this.btn_SaveData.Name = "btn_SaveData";
            this.btn_SaveData.Size = new System.Drawing.Size(140, 44);
            this.btn_SaveData.TabIndex = 381;
            this.btn_SaveData.Text = "שמור נתונים";
            this.btn_SaveData.UseVisualStyleBackColor = true;
            this.btn_SaveData.Click += new System.EventHandler(this.button1_Click);
            // 
            // lbl_totalLeftover
            // 
            this.lbl_totalLeftover.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_totalLeftover.AutoSize = true;
            this.lbl_totalLeftover.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_totalLeftover.Location = new System.Drawing.Point(607, 759);
            this.lbl_totalLeftover.Name = "lbl_totalLeftover";
            this.lbl_totalLeftover.Size = new System.Drawing.Size(58, 29);
            this.lbl_totalLeftover.TabIndex = 380;
            this.lbl_totalLeftover.Text = "*****";
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label12.Location = new System.Drawing.Point(993, 759);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(146, 29);
            this.label12.TabIndex = 379;
            this.label12.Text = "-כמות נותרת:";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label11.Location = new System.Drawing.Point(896, 467);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(237, 29);
            this.label11.TabIndex = 378;
            this.label11.Text = "-רשימת פירוט פסולות:";
            // 
            // txt_AmountSpesific
            // 
            this.txt_AmountSpesific.AccessibleName = "Code";
            this.txt_AmountSpesific.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txt_AmountSpesific.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.txt_AmountSpesific.Location = new System.Drawing.Point(880, 334);
            this.txt_AmountSpesific.Name = "txt_AmountSpesific";
            this.txt_AmountSpesific.Size = new System.Drawing.Size(152, 29);
            this.txt_AmountSpesific.TabIndex = 377;
            this.txt_AmountSpesific.TextChanged += new System.EventHandler(this.txt_AmountSpesific_TextChanged);
            // 
            // label10
            // 
            this.label10.AccessibleName = "Code";
            this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label10.Location = new System.Drawing.Point(1082, 341);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(42, 19);
            this.label10.TabIndex = 376;
            this.label10.Text = "כמות";
            // 
            // btn_confirm
            // 
            this.btn_confirm.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_confirm.Enabled = false;
            this.btn_confirm.Location = new System.Drawing.Point(338, 381);
            this.btn_confirm.Name = "btn_confirm";
            this.btn_confirm.Size = new System.Drawing.Size(75, 27);
            this.btn_confirm.TabIndex = 374;
            this.btn_confirm.Text = "אשר";
            this.btn_confirm.UseVisualStyleBackColor = true;
            this.btn_confirm.Click += new System.EventHandler(this.btn_confirm_Click);
            // 
            // label9
            // 
            this.label9.AccessibleName = "Code";
            this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label9.Location = new System.Drawing.Point(753, 389);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(86, 19);
            this.label9.TabIndex = 373;
            this.label9.Text = "פירוט תקלה";
            // 
            // txt_CodeWaste
            // 
            this.txt_CodeWaste.AccessibleName = "Code";
            this.txt_CodeWaste.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txt_CodeWaste.Enabled = false;
            this.txt_CodeWaste.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.txt_CodeWaste.Location = new System.Drawing.Point(446, 382);
            this.txt_CodeWaste.Name = "txt_CodeWaste";
            this.txt_CodeWaste.Size = new System.Drawing.Size(301, 29);
            this.txt_CodeWaste.TabIndex = 372;
            // 
            // label8
            // 
            this.label8.AccessibleName = "Code";
            this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label8.Location = new System.Drawing.Point(1047, 389);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(77, 19);
            this.label8.TabIndex = 371;
            this.label8.Text = "קוד פסולת";
            // 
            // cbo_CodeWaste
            // 
            this.cbo_CodeWaste.AccessibleName = "Code";
            this.cbo_CodeWaste.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbo_CodeWaste.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbo_CodeWaste.Font = new System.Drawing.Font("Courier New", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.cbo_CodeWaste.FormattingEnabled = true;
            this.cbo_CodeWaste.Location = new System.Drawing.Point(885, 385);
            this.cbo_CodeWaste.Name = "cbo_CodeWaste";
            this.cbo_CodeWaste.Size = new System.Drawing.Size(147, 30);
            this.cbo_CodeWaste.TabIndex = 370;
            this.cbo_CodeWaste.SelectedIndexChanged += new System.EventHandler(this.cbo_CodeWaste_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label5.Location = new System.Drawing.Point(977, 270);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(162, 29);
            this.label5.TabIndex = 8;
            this.label5.Text = "-פירוט פסולות:";
            // 
            // lbl_dateWantedUpdate
            // 
            this.lbl_dateWantedUpdate.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbl_dateWantedUpdate.AutoSize = true;
            this.lbl_dateWantedUpdate.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_dateWantedUpdate.Location = new System.Drawing.Point(344, 64);
            this.lbl_dateWantedUpdate.Name = "lbl_dateWantedUpdate";
            this.lbl_dateWantedUpdate.Size = new System.Drawing.Size(119, 34);
            this.lbl_dateWantedUpdate.TabIndex = 7;
            this.lbl_dateWantedUpdate.Text = "**/**/****";
            // 
            // label7
            // 
            this.label7.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label7.Location = new System.Drawing.Point(838, 64);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(171, 34);
            this.label7.TabIndex = 6;
            this.label7.Text = "נכון לתאריך :";
            // 
            // lbl_AmountWaste
            // 
            this.lbl_AmountWaste.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbl_AmountWaste.AutoSize = true;
            this.lbl_AmountWaste.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_AmountWaste.Location = new System.Drawing.Point(822, 216);
            this.lbl_AmountWaste.Name = "lbl_AmountWaste";
            this.lbl_AmountWaste.Size = new System.Drawing.Size(58, 29);
            this.lbl_AmountWaste.TabIndex = 3;
            this.lbl_AmountWaste.Text = "*****";
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label6.Location = new System.Drawing.Point(924, 216);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(215, 29);
            this.label6.TabIndex = 2;
            this.label6.Text = "-כמות פסולת כוללת:";
            // 
            // lbl_WasteUpdate
            // 
            this.lbl_WasteUpdate.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbl_WasteUpdate.AutoSize = true;
            this.lbl_WasteUpdate.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_WasteUpdate.Location = new System.Drawing.Point(101, 17);
            this.lbl_WasteUpdate.Name = "lbl_WasteUpdate";
            this.lbl_WasteUpdate.Size = new System.Drawing.Size(356, 34);
            this.lbl_WasteUpdate.TabIndex = 1;
            this.lbl_WasteUpdate.Text = "*******************************";
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label4.Location = new System.Drawing.Point(757, 17);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(252, 34);
            this.label4.TabIndex = 0;
            this.label4.Text = "עדכון פסולות עבור :";
            // 
            // Cbo_MonthReport
            // 
            this.Cbo_MonthReport.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.Cbo_MonthReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cbo_MonthReport.Font = new System.Drawing.Font("Courier New", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.Cbo_MonthReport.FormattingEnabled = true;
            this.Cbo_MonthReport.Location = new System.Drawing.Point(759, 83);
            this.Cbo_MonthReport.Name = "Cbo_MonthReport";
            this.Cbo_MonthReport.Size = new System.Drawing.Size(147, 30);
            this.Cbo_MonthReport.TabIndex = 211;
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Harlow Solid Italic", 15.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(971, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 26);
            this.label1.TabIndex = 214;
            this.label1.Text = ":חודש";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Harlow Solid Italic", 15.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(641, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 26);
            this.label2.TabIndex = 215;
            this.label2.Text = ":שנה";
            // 
            // Cbo_YearReport
            // 
            this.Cbo_YearReport.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.Cbo_YearReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Cbo_YearReport.Font = new System.Drawing.Font("Courier New", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.Cbo_YearReport.FormattingEnabled = true;
            this.Cbo_YearReport.Location = new System.Drawing.Point(488, 83);
            this.Cbo_YearReport.Name = "Cbo_YearReport";
            this.Cbo_YearReport.Size = new System.Drawing.Size(147, 30);
            this.Cbo_YearReport.TabIndex = 216;
            // 
            // btn_show
            // 
            this.btn_show.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.btn_show.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.btn_show.Location = new System.Drawing.Point(410, 86);
            this.btn_show.Name = "btn_show";
            this.btn_show.Size = new System.Drawing.Size(51, 23);
            this.btn_show.TabIndex = 369;
            this.btn_show.Text = "הצג";
            this.btn_show.UseVisualStyleBackColor = true;
            this.btn_show.Click += new System.EventHandler(this.btn_show_Click);
            // 
            // lbl_Report
            // 
            this.lbl_Report.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbl_Report.AutoSize = true;
            this.lbl_Report.Font = new System.Drawing.Font("Harlow Solid Italic", 15.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Report.Location = new System.Drawing.Point(971, 35);
            this.lbl_Report.Name = "lbl_Report";
            this.lbl_Report.Size = new System.Drawing.Size(232, 26);
            this.lbl_Report.TabIndex = 212;
            this.lbl_Report.Text = ":דו\"ח פסולות נכון לחודש";
            this.lbl_Report.Visible = false;
            // 
            // lbl_MonthReport
            // 
            this.lbl_MonthReport.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lbl_MonthReport.AutoSize = true;
            this.lbl_MonthReport.Font = new System.Drawing.Font("Harlow Solid Italic", 15.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_MonthReport.Location = new System.Drawing.Point(767, 35);
            this.lbl_MonthReport.Name = "lbl_MonthReport";
            this.lbl_MonthReport.Size = new System.Drawing.Size(139, 26);
            this.lbl_MonthReport.TabIndex = 213;
            this.lbl_MonthReport.Text = "******* 2018";
            this.lbl_MonthReport.Visible = false;
            // 
            // MakeExcellRepBTN
            // 
            this.MakeExcellRepBTN.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.MakeExcellRepBTN.BackColor = System.Drawing.Color.LightGreen;
            this.MakeExcellRepBTN.BackgroundImage = global::WasteManagement.Properties.Resources.Excel_Icon;
            this.MakeExcellRepBTN.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MakeExcellRepBTN.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.MakeExcellRepBTN.Location = new System.Drawing.Point(333, 79);
            this.MakeExcellRepBTN.Name = "MakeExcellRepBTN";
            this.MakeExcellRepBTN.Size = new System.Drawing.Size(71, 33);
            this.MakeExcellRepBTN.TabIndex = 370;
            this.MakeExcellRepBTN.UseVisualStyleBackColor = false;
            this.MakeExcellRepBTN.Click += new System.EventHandler(this.MakeExcellRepBTN_Click);
            // 
            // label17
            // 
            this.label17.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.label17.AutoSize = true;
            this.label17.BackColor = System.Drawing.Color.Transparent;
            this.label17.Font = new System.Drawing.Font("Tahoma", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.label17.ForeColor = System.Drawing.Color.DarkOliveGreen;
            this.label17.Location = new System.Drawing.Point(513, 3);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(338, 58);
            this.label17.TabIndex = 371;
            this.label17.Text = "ניהול פסולות";
            // 
            // lbl_LastUpdate
            // 
            this.lbl_LastUpdate.AutoSize = true;
            this.lbl_LastUpdate.Font = new System.Drawing.Font("Tahoma", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.lbl_LastUpdate.Location = new System.Drawing.Point(10, 9);
            this.lbl_LastUpdate.Name = "lbl_LastUpdate";
            this.lbl_LastUpdate.Size = new System.Drawing.Size(185, 33);
            this.lbl_LastUpdate.TabIndex = 372;
            this.lbl_LastUpdate.Text = "**********";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LemonChiffon;
            this.ClientSize = new System.Drawing.Size(1273, 974);
            this.Controls.Add(this.lbl_LastUpdate);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.MakeExcellRepBTN);
            this.Controls.Add(this.btn_show);
            this.Controls.Add(this.Cbo_YearReport);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lbl_MonthReport);
            this.Controls.Add(this.lbl_Report);
            this.Controls.Add(this.Cbo_MonthReport);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.Report.ResumeLayout(false);
            this.Report.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DgvWasteMonth)).EndInit();
            this.WasteUpdate.ResumeLayout(false);
            this.WasteUpdate.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGrid_Desc)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage Report;
        private ADGV.AdvancedDataGridView DgvWasteMonth;
        private System.Windows.Forms.TabPage WasteUpdate;
        private System.Windows.Forms.ComboBox Cbo_MonthReport;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox Cbo_YearReport;
        private System.Windows.Forms.Button btn_show;
        private Label lbl_totalCured;
        private Label lblTotalCured;
        private Label lbl_AmountWaste;
        private Label label6;
        private Label lbl_WasteUpdate;
        private Label label4;
        private Label lbl_Report;
        private Label lbl_MonthReport;
        private Label lbl_dateWantedUpdate;
        private Label label7;
        private Label label9;
        private TextBox txt_CodeWaste;
        private Label label8;
        private ComboBox cbo_CodeWaste;
        private Label label5;
        private Button btn_confirm;
        private TextBox txt_AmountSpesific;
        private Label label10;
        private Label label11;
        private Label lbl_totalLeftover;
        private Label label12;
        private Button btn_SaveData;
        private Button MakeExcellRepBTN;
        private DataGridView DataGrid_Desc;
        private Label label13;
        private Label lbl_PrecentTotalScrap;
        private Label lblheaderScrapPrecent;
        private Label lbl_EntryToWarehouse;
        private Label lblheaderEntryToWarehouse;
        private Label lbl_totalScrap;
        private Label lblheaderTotalScrap;
        private Label lbl_ScrapTGT;
        private Label lblheaderScrapTgt;
        private Label label17;
        private Label lbl_LastUpdate;
    }
}

