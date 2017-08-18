namespace WindowsFormsApplication1
{
    partial class frmReport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmReport));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.NPP = new System.Windows.Forms.TabPage();
            this.btnDel_NPP = new System.Windows.Forms.Button();
            this.btnEdit_NPP = new System.Windows.Forms.Button();
            this.dgvNPP = new System.Windows.Forms.DataGridView();
            this.UUT = new System.Windows.Forms.TabPage();
            this.btnDel_UUT = new System.Windows.Forms.Button();
            this.btnEdit_UUT = new System.Windows.Forms.Button();
            this.dgvUUT = new System.Windows.Forms.DataGridView();
            this.SC = new System.Windows.Forms.TabPage();
            this.btnDel_SC = new System.Windows.Forms.Button();
            this.btnEdit_SC = new System.Windows.Forms.Button();
            this.dgvSC = new System.Windows.Forms.DataGridView();
            this.OC = new System.Windows.Forms.TabPage();
            this.btnDel_OC = new System.Windows.Forms.Button();
            this.btnEdit_OC = new System.Windows.Forms.Button();
            this.dgvOC = new System.Windows.Forms.DataGridView();
            this.Terms = new System.Windows.Forms.TabPage();
            this.btnDel_Terms = new System.Windows.Forms.Button();
            this.btnEdit_Terms = new System.Windows.Forms.Button();
            this.dgvTerms = new System.Windows.Forms.DataGridView();
            this.Trans = new System.Windows.Forms.TabPage();
            this.btnDel_Trans = new System.Windows.Forms.Button();
            this.btnEdit_Trans = new System.Windows.Forms.Button();
            this.dgvTrans = new System.Windows.Forms.DataGridView();
            this.cboItemsReport = new System.Windows.Forms.ComboBox();
            this.cboYearPP = new System.Windows.Forms.ComboBox();
            this.cboPP = new System.Windows.Forms.ComboBox();
            this.label34 = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.label35 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.NPP.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvNPP)).BeginInit();
            this.UUT.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUUT)).BeginInit();
            this.SC.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSC)).BeginInit();
            this.OC.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOC)).BeginInit();
            this.Terms.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerms)).BeginInit();
            this.Trans.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTrans)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl1.Controls.Add(this.NPP);
            this.tabControl1.Controls.Add(this.UUT);
            this.tabControl1.Controls.Add(this.SC);
            this.tabControl1.Controls.Add(this.OC);
            this.tabControl1.Controls.Add(this.Terms);
            this.tabControl1.Controls.Add(this.Trans);
            this.tabControl1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl1.Location = new System.Drawing.Point(2, 60);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.Padding = new System.Drawing.Point(6, 6);
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1618, 504);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.Tag = "dgvTerms";
            // 
            // NPP
            // 
            this.NPP.BackColor = System.Drawing.Color.PaleGoldenrod;
            this.NPP.Controls.Add(this.btnDel_NPP);
            this.NPP.Controls.Add(this.btnEdit_NPP);
            this.NPP.Controls.Add(this.dgvNPP);
            this.NPP.Location = new System.Drawing.Point(4, 36);
            this.NPP.Name = "NPP";
            this.NPP.Size = new System.Drawing.Size(1610, 464);
            this.NPP.TabIndex = 2;
            this.NPP.Text = "New Primary Positions";
            // 
            // btnDel_NPP
            // 
            this.btnDel_NPP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_NPP.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_NPP.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_NPP.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_NPP.Image")));
            this.btnDel_NPP.Location = new System.Drawing.Point(116, 416);
            this.btnDel_NPP.Name = "btnDel_NPP";
            this.btnDel_NPP.Size = new System.Drawing.Size(104, 40);
            this.btnDel_NPP.TabIndex = 47;
            this.btnDel_NPP.Tag = "dgvNPP";
            this.btnDel_NPP.Text = "Delete";
            this.btnDel_NPP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_NPP.UseVisualStyleBackColor = false;
            this.btnDel_NPP.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_NPP
            // 
            this.btnEdit_NPP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_NPP.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_NPP.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_NPP.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_NPP.Image")));
            this.btnEdit_NPP.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_NPP.Name = "btnEdit_NPP";
            this.btnEdit_NPP.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_NPP.TabIndex = 46;
            this.btnEdit_NPP.Text = "Edit";
            this.btnEdit_NPP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_NPP.UseVisualStyleBackColor = false;
            this.btnEdit_NPP.Click += new System.EventHandler(this.btnEdit_NPP_Click);
            // 
            // dgvNPP
            // 
            this.dgvNPP.AllowUserToAddRows = false;
            this.dgvNPP.AllowUserToDeleteRows = false;
            this.dgvNPP.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvNPP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvNPP.Location = new System.Drawing.Point(3, 4);
            this.dgvNPP.Name = "dgvNPP";
            this.dgvNPP.ReadOnly = true;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvNPP.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvNPP.Size = new System.Drawing.Size(1604, 404);
            this.dgvNPP.TabIndex = 45;
            this.dgvNPP.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvNPP_CellDoubleClick);
            // 
            // UUT
            // 
            this.UUT.BackColor = System.Drawing.Color.LightPink;
            this.UUT.Controls.Add(this.btnDel_UUT);
            this.UUT.Controls.Add(this.btnEdit_UUT);
            this.UUT.Controls.Add(this.dgvUUT);
            this.UUT.Location = new System.Drawing.Point(4, 36);
            this.UUT.Name = "UUT";
            this.UUT.Padding = new System.Windows.Forms.Padding(3);
            this.UUT.Size = new System.Drawing.Size(1289, 464);
            this.UUT.TabIndex = 0;
            this.UUT.Text = "Unit To Unit Transfer";
            // 
            // btnDel_UUT
            // 
            this.btnDel_UUT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_UUT.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_UUT.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_UUT.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_UUT.Image")));
            this.btnDel_UUT.Location = new System.Drawing.Point(116, 416);
            this.btnDel_UUT.Name = "btnDel_UUT";
            this.btnDel_UUT.Size = new System.Drawing.Size(104, 40);
            this.btnDel_UUT.TabIndex = 49;
            this.btnDel_UUT.Tag = "dgvUUT";
            this.btnDel_UUT.Text = "Delete";
            this.btnDel_UUT.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_UUT.UseVisualStyleBackColor = false;
            this.btnDel_UUT.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_UUT
            // 
            this.btnEdit_UUT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_UUT.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_UUT.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_UUT.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_UUT.Image")));
            this.btnEdit_UUT.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_UUT.Name = "btnEdit_UUT";
            this.btnEdit_UUT.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_UUT.TabIndex = 48;
            this.btnEdit_UUT.Text = "Edit";
            this.btnEdit_UUT.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_UUT.UseVisualStyleBackColor = false;
            this.btnEdit_UUT.Click += new System.EventHandler(this.btnEdit_UUT_Click);
            // 
            // dgvUUT
            // 
            this.dgvUUT.AllowUserToAddRows = false;
            this.dgvUUT.AllowUserToDeleteRows = false;
            this.dgvUUT.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvUUT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUUT.Location = new System.Drawing.Point(3, 3);
            this.dgvUUT.Name = "dgvUUT";
            this.dgvUUT.ReadOnly = true;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvUUT.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvUUT.Size = new System.Drawing.Size(1283, 407);
            this.dgvUUT.TabIndex = 42;
            this.dgvUUT.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvUUT_CellDoubleClick);
            // 
            // SC
            // 
            this.SC.BackColor = System.Drawing.Color.OliveDrab;
            this.SC.Controls.Add(this.btnDel_SC);
            this.SC.Controls.Add(this.btnEdit_SC);
            this.SC.Controls.Add(this.dgvSC);
            this.SC.Location = new System.Drawing.Point(4, 36);
            this.SC.Name = "SC";
            this.SC.Padding = new System.Windows.Forms.Padding(3);
            this.SC.Size = new System.Drawing.Size(1289, 464);
            this.SC.TabIndex = 1;
            this.SC.Text = "Status Change";
            // 
            // btnDel_SC
            // 
            this.btnDel_SC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_SC.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_SC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_SC.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_SC.Image")));
            this.btnDel_SC.Location = new System.Drawing.Point(116, 416);
            this.btnDel_SC.Name = "btnDel_SC";
            this.btnDel_SC.Size = new System.Drawing.Size(104, 40);
            this.btnDel_SC.TabIndex = 51;
            this.btnDel_SC.Tag = "dgvSC";
            this.btnDel_SC.Text = "Delete";
            this.btnDel_SC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_SC.UseVisualStyleBackColor = false;
            this.btnDel_SC.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_SC
            // 
            this.btnEdit_SC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_SC.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_SC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_SC.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_SC.Image")));
            this.btnEdit_SC.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_SC.Name = "btnEdit_SC";
            this.btnEdit_SC.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_SC.TabIndex = 50;
            this.btnEdit_SC.Text = "Edit";
            this.btnEdit_SC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_SC.UseVisualStyleBackColor = false;
            this.btnEdit_SC.Click += new System.EventHandler(this.btnEdit_SC_Click);
            // 
            // dgvSC
            // 
            this.dgvSC.AllowUserToAddRows = false;
            this.dgvSC.AllowUserToDeleteRows = false;
            this.dgvSC.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSC.Location = new System.Drawing.Point(3, 4);
            this.dgvSC.Name = "dgvSC";
            this.dgvSC.ReadOnly = true;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvSC.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgvSC.Size = new System.Drawing.Size(1283, 406);
            this.dgvSC.TabIndex = 48;
            this.dgvSC.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvSC_CellDoubleClick);
            // 
            // OC
            // 
            this.OC.BackColor = System.Drawing.Color.SteelBlue;
            this.OC.Controls.Add(this.btnDel_OC);
            this.OC.Controls.Add(this.btnEdit_OC);
            this.OC.Controls.Add(this.dgvOC);
            this.OC.Location = new System.Drawing.Point(4, 36);
            this.OC.Name = "OC";
            this.OC.Size = new System.Drawing.Size(1289, 464);
            this.OC.TabIndex = 3;
            this.OC.Text = "Occupation Changes";
            // 
            // btnDel_OC
            // 
            this.btnDel_OC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_OC.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_OC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_OC.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_OC.Image")));
            this.btnDel_OC.Location = new System.Drawing.Point(116, 416);
            this.btnDel_OC.Name = "btnDel_OC";
            this.btnDel_OC.Size = new System.Drawing.Size(104, 40);
            this.btnDel_OC.TabIndex = 51;
            this.btnDel_OC.Tag = "dgvOC";
            this.btnDel_OC.Text = "Delete";
            this.btnDel_OC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_OC.UseVisualStyleBackColor = false;
            this.btnDel_OC.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_OC
            // 
            this.btnEdit_OC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_OC.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_OC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_OC.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_OC.Image")));
            this.btnEdit_OC.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_OC.Name = "btnEdit_OC";
            this.btnEdit_OC.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_OC.TabIndex = 50;
            this.btnEdit_OC.Text = "Edit";
            this.btnEdit_OC.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_OC.UseVisualStyleBackColor = false;
            this.btnEdit_OC.Click += new System.EventHandler(this.btnEdit_OC_Click);
            // 
            // dgvOC
            // 
            this.dgvOC.AllowUserToAddRows = false;
            this.dgvOC.AllowUserToDeleteRows = false;
            this.dgvOC.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvOC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvOC.Location = new System.Drawing.Point(3, 4);
            this.dgvOC.Name = "dgvOC";
            this.dgvOC.ReadOnly = true;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvOC.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.dgvOC.Size = new System.Drawing.Size(1283, 406);
            this.dgvOC.TabIndex = 48;
            this.dgvOC.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvOC_CellDoubleClick);
            // 
            // Terms
            // 
            this.Terms.BackColor = System.Drawing.Color.SaddleBrown;
            this.Terms.Controls.Add(this.btnDel_Terms);
            this.Terms.Controls.Add(this.btnEdit_Terms);
            this.Terms.Controls.Add(this.dgvTerms);
            this.Terms.Location = new System.Drawing.Point(4, 36);
            this.Terms.Name = "Terms";
            this.Terms.Size = new System.Drawing.Size(1289, 464);
            this.Terms.TabIndex = 4;
            this.Terms.Text = "Terminations";
            // 
            // btnDel_Terms
            // 
            this.btnDel_Terms.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_Terms.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_Terms.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_Terms.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_Terms.Image")));
            this.btnDel_Terms.Location = new System.Drawing.Point(116, 416);
            this.btnDel_Terms.Name = "btnDel_Terms";
            this.btnDel_Terms.Size = new System.Drawing.Size(104, 40);
            this.btnDel_Terms.TabIndex = 51;
            this.btnDel_Terms.Tag = "dgvTerms";
            this.btnDel_Terms.Text = "Delete";
            this.btnDel_Terms.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_Terms.UseVisualStyleBackColor = false;
            this.btnDel_Terms.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_Terms
            // 
            this.btnEdit_Terms.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_Terms.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_Terms.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_Terms.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_Terms.Image")));
            this.btnEdit_Terms.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_Terms.Name = "btnEdit_Terms";
            this.btnEdit_Terms.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_Terms.TabIndex = 50;
            this.btnEdit_Terms.Text = "Edit";
            this.btnEdit_Terms.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_Terms.UseVisualStyleBackColor = false;
            this.btnEdit_Terms.Click += new System.EventHandler(this.btnEdit_Terms_Click);
            // 
            // dgvTerms
            // 
            this.dgvTerms.AllowUserToAddRows = false;
            this.dgvTerms.AllowUserToDeleteRows = false;
            this.dgvTerms.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTerms.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTerms.Location = new System.Drawing.Point(3, 4);
            this.dgvTerms.Name = "dgvTerms";
            this.dgvTerms.ReadOnly = true;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvTerms.RowsDefaultCellStyle = dataGridViewCellStyle5;
            this.dgvTerms.Size = new System.Drawing.Size(1283, 406);
            this.dgvTerms.TabIndex = 48;
            this.dgvTerms.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTerms_CellDoubleClick);
            // 
            // Trans
            // 
            this.Trans.BackColor = System.Drawing.Color.IndianRed;
            this.Trans.Controls.Add(this.btnDel_Trans);
            this.Trans.Controls.Add(this.btnEdit_Trans);
            this.Trans.Controls.Add(this.dgvTrans);
            this.Trans.Location = new System.Drawing.Point(4, 36);
            this.Trans.Name = "Trans";
            this.Trans.Size = new System.Drawing.Size(1289, 464);
            this.Trans.TabIndex = 5;
            this.Trans.Text = "Transfers";
            // 
            // btnDel_Trans
            // 
            this.btnDel_Trans.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDel_Trans.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnDel_Trans.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDel_Trans.Image = ((System.Drawing.Image)(resources.GetObject("btnDel_Trans.Image")));
            this.btnDel_Trans.Location = new System.Drawing.Point(116, 416);
            this.btnDel_Trans.Name = "btnDel_Trans";
            this.btnDel_Trans.Size = new System.Drawing.Size(104, 40);
            this.btnDel_Trans.TabIndex = 51;
            this.btnDel_Trans.Tag = "dgvTrans";
            this.btnDel_Trans.Text = "Delete";
            this.btnDel_Trans.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnDel_Trans.UseVisualStyleBackColor = false;
            this.btnDel_Trans.Click += new System.EventHandler(this.btnDel_NPP_Click);
            // 
            // btnEdit_Trans
            // 
            this.btnEdit_Trans.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnEdit_Trans.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnEdit_Trans.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEdit_Trans.Image = ((System.Drawing.Image)(resources.GetObject("btnEdit_Trans.Image")));
            this.btnEdit_Trans.Location = new System.Drawing.Point(6, 416);
            this.btnEdit_Trans.Name = "btnEdit_Trans";
            this.btnEdit_Trans.Size = new System.Drawing.Size(104, 40);
            this.btnEdit_Trans.TabIndex = 50;
            this.btnEdit_Trans.Text = "Edit";
            this.btnEdit_Trans.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnEdit_Trans.UseVisualStyleBackColor = false;
            this.btnEdit_Trans.Click += new System.EventHandler(this.btnEdit_Trans_Click);
            // 
            // dgvTrans
            // 
            this.dgvTrans.AllowUserToAddRows = false;
            this.dgvTrans.AllowUserToDeleteRows = false;
            this.dgvTrans.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvTrans.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTrans.Location = new System.Drawing.Point(3, 4);
            this.dgvTrans.Name = "dgvTrans";
            this.dgvTrans.ReadOnly = true;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dgvTrans.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.dgvTrans.Size = new System.Drawing.Size(1283, 406);
            this.dgvTrans.TabIndex = 48;
            this.dgvTrans.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTrans_CellDoubleClick);
            // 
            // cboItemsReport
            // 
            this.cboItemsReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboItemsReport.FormattingEnabled = true;
            this.cboItemsReport.Items.AddRange(new object[] {
            "Items Report A",
            "Items Report B",
            "Items Report C"});
            this.cboItemsReport.Location = new System.Drawing.Point(308, 17);
            this.cboItemsReport.Name = "cboItemsReport";
            this.cboItemsReport.Size = new System.Drawing.Size(121, 21);
            this.cboItemsReport.TabIndex = 63;
            // 
            // cboYearPP
            // 
            this.cboYearPP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboYearPP.FormattingEnabled = true;
            this.cboYearPP.Location = new System.Drawing.Point(181, 17);
            this.cboYearPP.Name = "cboYearPP";
            this.cboYearPP.Size = new System.Drawing.Size(121, 21);
            this.cboYearPP.TabIndex = 61;
            // 
            // cboPP
            // 
            this.cboPP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPP.FormattingEnabled = true;
            this.cboPP.Items.AddRange(new object[] {
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26"});
            this.cboPP.Location = new System.Drawing.Point(81, 17);
            this.cboPP.Name = "cboPP";
            this.cboPP.Size = new System.Drawing.Size(56, 21);
            this.cboPP.TabIndex = 59;
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Location = new System.Drawing.Point(14, 20);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(69, 15);
            this.label34.TabIndex = 60;
            this.label34.Text = "Pay Period:";
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.Location = new System.Drawing.Point(435, 6);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(104, 40);
            this.btnRefresh.TabIndex = 64;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(143, 20);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(35, 15);
            this.label35.TabIndex = 62;
            this.label35.Text = "Year:";
            // 
            // btnClose
            // 
            this.btnClose.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.Location = new System.Drawing.Point(1184, 6);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(104, 40);
            this.btnClose.TabIndex = 65;
            this.btnClose.Text = "Minimize";
            this.btnClose.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnClose.UseVisualStyleBackColor = false;
            this.btnClose.Visible = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // frmReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1621, 564);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.cboItemsReport);
            this.Controls.Add(this.label35);
            this.Controls.Add(this.cboYearPP);
            this.Controls.Add(this.cboPP);
            this.Controls.Add(this.label34);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmReport";
            this.Text = "View Items Report";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmReport_FormClosing);
            this.Load += new System.EventHandler(this.frmReport_Load);
            this.tabControl1.ResumeLayout(false);
            this.NPP.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvNPP)).EndInit();
            this.UUT.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUUT)).EndInit();
            this.SC.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSC)).EndInit();
            this.OC.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvOC)).EndInit();
            this.Terms.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTerms)).EndInit();
            this.Trans.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTrans)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TabPage UUT;
        private System.Windows.Forms.TabPage SC;
        private System.Windows.Forms.DataGridView dgvUUT;
        private System.Windows.Forms.Label label34;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Label label35;
        public System.Windows.Forms.ComboBox cboItemsReport;
        public System.Windows.Forms.ComboBox cboYearPP;
        public System.Windows.Forms.ComboBox cboPP;
        public System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage NPP;
        private System.Windows.Forms.Button btnDel_NPP;
        private System.Windows.Forms.Button btnEdit_NPP;
        private System.Windows.Forms.DataGridView dgvNPP;
        private System.Windows.Forms.DataGridView dgvSC;
        private System.Windows.Forms.TabPage OC;
        private System.Windows.Forms.DataGridView dgvOC;
        private System.Windows.Forms.TabPage Terms;
        private System.Windows.Forms.DataGridView dgvTerms;
        private System.Windows.Forms.TabPage Trans;
        private System.Windows.Forms.DataGridView dgvTrans;
        private System.Windows.Forms.Button btnDel_UUT;
        private System.Windows.Forms.Button btnEdit_UUT;
        private System.Windows.Forms.Button btnDel_SC;
        private System.Windows.Forms.Button btnEdit_SC;
        private System.Windows.Forms.Button btnDel_OC;
        private System.Windows.Forms.Button btnEdit_OC;
        private System.Windows.Forms.Button btnDel_Terms;
        private System.Windows.Forms.Button btnEdit_Terms;
        private System.Windows.Forms.Button btnDel_Trans;
        private System.Windows.Forms.Button btnEdit_Trans;
        private System.Windows.Forms.Button btnClose;
    }
}