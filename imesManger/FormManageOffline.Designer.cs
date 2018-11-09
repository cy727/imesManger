namespace imesManger
{
    partial class FormManageOffline
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormManageOffline));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.printPreviewToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.btnAll = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btnPCF = new System.Windows.Forms.Button();
            this.btnICF = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelC = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.printToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.comboBoxIC = new System.Windows.Forms.ComboBox();
            this.comboBoxPC = new System.Windows.Forms.ComboBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btndisplayfiltersection = new System.Windows.Forms.Button();
            this.btndisplayfilter = new System.Windows.Forms.Button();
            this.checkBoxBefore = new System.Windows.Forms.CheckBox();
            this.dateTimePickerBefore = new System.Windows.Forms.DateTimePicker();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBoxTAC = new System.Windows.Forms.CheckBox();
            this.btnRe = new System.Windows.Forms.Button();
            this.dateTimePickerP = new System.Windows.Forms.DateTimePicker();
            this.dataGridViewP = new System.Windows.Forms.DataGridView();
            this.statusStrip1.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewP)).BeginInit();
            this.SuspendLayout();
            // 
            // printPreviewToolStripButton
            // 
            this.printPreviewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.printPreviewToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("printPreviewToolStripButton.Image")));
            this.printPreviewToolStripButton.ImageTransparentColor = System.Drawing.Color.Black;
            this.printPreviewToolStripButton.Name = "printPreviewToolStripButton";
            this.printPreviewToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.printPreviewToolStripButton.Text = "print previw";
            this.printPreviewToolStripButton.Click += new System.EventHandler(this.printPreviewToolStripButton_Click);
            // 
            // btnAll
            // 
            this.btnAll.Location = new System.Drawing.Point(384, 15);
            this.btnAll.Name = "btnAll";
            this.btnAll.Size = new System.Drawing.Size(55, 45);
            this.btnAll.TabIndex = 10;
            this.btnAll.Text = "Search";
            this.btnAll.UseVisualStyleBackColor = true;
            this.btnAll.Click += new System.EventHandler(this.btnAll_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "Customer Code:";
            // 
            // btnPCF
            // 
            this.btnPCF.Location = new System.Drawing.Point(227, 15);
            this.btnPCF.Name = "btnPCF";
            this.btnPCF.Size = new System.Drawing.Size(151, 21);
            this.btnPCF.TabIndex = 7;
            this.btnPCF.Text = "Product Code filter";
            this.btnPCF.UseVisualStyleBackColor = true;
            this.btnPCF.Click += new System.EventHandler(this.btnPCF_Click);
            // 
            // btnICF
            // 
            this.btnICF.Location = new System.Drawing.Point(227, 41);
            this.btnICF.Name = "btnICF";
            this.btnICF.Size = new System.Drawing.Size(151, 21);
            this.btnICF.TabIndex = 5;
            this.btnICF.Text = "Customer Code filter";
            this.btnICF.UseVisualStyleBackColor = true;
            this.btnICF.Click += new System.EventHandler(this.btnICF_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabelC});
            this.statusStrip1.Location = new System.Drawing.Point(0, 440);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(921, 22);
            this.statusStrip1.TabIndex = 16;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(831, 17);
            this.toolStripStatusLabel1.Spring = true;
            // 
            // toolStripStatusLabelC
            // 
            this.toolStripStatusLabelC.Name = "toolStripStatusLabelC";
            this.toolStripStatusLabelC.Size = new System.Drawing.Size(75, 17);
            this.toolStripStatusLabelC.Text = "单位数量：0";
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.printToolStripButton,
            this.printPreviewToolStripButton});
            this.toolStrip.Location = new System.Drawing.Point(0, 0);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(921, 25);
            this.toolStrip.TabIndex = 15;
            this.toolStrip.Text = "ToolStrip";
            // 
            // printToolStripButton
            // 
            this.printToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.printToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("printToolStripButton.Image")));
            this.printToolStripButton.ImageTransparentColor = System.Drawing.Color.Black;
            this.printToolStripButton.Name = "printToolStripButton";
            this.printToolStripButton.Size = new System.Drawing.Size(23, 22);
            this.printToolStripButton.Text = "print";
            this.printToolStripButton.Click += new System.EventHandler(this.printToolStripButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Product Code:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.comboBoxIC);
            this.groupBox3.Controls.Add(this.comboBoxPC);
            this.groupBox3.Controls.Add(this.btnAll);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.btnPCF);
            this.groupBox3.Controls.Add(this.btnICF);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(7, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(455, 66);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Filter";
            // 
            // comboBoxIC
            // 
            this.comboBoxIC.FormattingEnabled = true;
            this.comboBoxIC.Location = new System.Drawing.Point(100, 41);
            this.comboBoxIC.Name = "comboBoxIC";
            this.comboBoxIC.Size = new System.Drawing.Size(121, 20);
            this.comboBoxIC.TabIndex = 12;
            // 
            // comboBoxPC
            // 
            this.comboBoxPC.FormattingEnabled = true;
            this.comboBoxPC.Location = new System.Drawing.Point(100, 15);
            this.comboBoxPC.Name = "comboBoxPC";
            this.comboBoxPC.Size = new System.Drawing.Size(121, 20);
            this.comboBoxPC.TabIndex = 11;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(0, 25);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.groupBox2);
            this.splitContainer1.Panel1.Controls.Add(this.groupBox1);
            this.splitContainer1.Panel1.Controls.Add(this.groupBox3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewP);
            this.splitContainer1.Size = new System.Drawing.Size(921, 415);
            this.splitContainer1.SplitterDistance = 79;
            this.splitContainer1.TabIndex = 17;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btndisplayfiltersection);
            this.groupBox2.Controls.Add(this.btndisplayfilter);
            this.groupBox2.Controls.Add(this.checkBoxBefore);
            this.groupBox2.Controls.Add(this.dateTimePickerBefore);
            this.groupBox2.Location = new System.Drawing.Point(698, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(217, 66);
            this.groupBox2.TabIndex = 7;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Display Warn Filter";
            // 
            // btndisplayfiltersection
            // 
            this.btndisplayfiltersection.Location = new System.Drawing.Point(112, 43);
            this.btndisplayfiltersection.Name = "btndisplayfiltersection";
            this.btndisplayfiltersection.Size = new System.Drawing.Size(88, 23);
            this.btndisplayfiltersection.TabIndex = 5;
            this.btndisplayfiltersection.Text = "Show Section";
            this.btndisplayfiltersection.UseVisualStyleBackColor = true;
            this.btndisplayfiltersection.Click += new System.EventHandler(this.btndisplayfiltersection_Click);
            // 
            // btndisplayfilter
            // 
            this.btndisplayfilter.Location = new System.Drawing.Point(9, 43);
            this.btndisplayfilter.Name = "btndisplayfilter";
            this.btndisplayfilter.Size = new System.Drawing.Size(88, 23);
            this.btndisplayfilter.TabIndex = 4;
            this.btndisplayfilter.Text = "Show Reserve";
            this.btndisplayfilter.UseVisualStyleBackColor = true;
            this.btndisplayfilter.Click += new System.EventHandler(this.btndisplayfilter_Click);
            // 
            // checkBoxBefore
            // 
            this.checkBoxBefore.AutoSize = true;
            this.checkBoxBefore.Location = new System.Drawing.Point(9, 24);
            this.checkBoxBefore.Name = "checkBoxBefore";
            this.checkBoxBefore.Size = new System.Drawing.Size(60, 16);
            this.checkBoxBefore.TabIndex = 3;
            this.checkBoxBefore.Text = "before";
            this.checkBoxBefore.UseVisualStyleBackColor = true;
            // 
            // dateTimePickerBefore
            // 
            this.dateTimePickerBefore.Location = new System.Drawing.Point(75, 20);
            this.dateTimePickerBefore.Name = "dateTimePickerBefore";
            this.dateTimePickerBefore.Size = new System.Drawing.Size(136, 21);
            this.dateTimePickerBefore.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkBoxTAC);
            this.groupBox1.Controls.Add(this.btnRe);
            this.groupBox1.Controls.Add(this.dateTimePickerP);
            this.groupBox1.Location = new System.Drawing.Point(466, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(226, 66);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Date";
            // 
            // checkBoxTAC
            // 
            this.checkBoxTAC.AutoSize = true;
            this.checkBoxTAC.Checked = true;
            this.checkBoxTAC.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxTAC.Location = new System.Drawing.Point(7, 44);
            this.checkBoxTAC.Name = "checkBoxTAC";
            this.checkBoxTAC.Size = new System.Drawing.Size(210, 16);
            this.checkBoxTAC.TabIndex = 3;
            this.checkBoxTAC.Text = "hidden all recoder which no TAC";
            this.checkBoxTAC.UseVisualStyleBackColor = true;
            // 
            // btnRe
            // 
            this.btnRe.Location = new System.Drawing.Point(148, 18);
            this.btnRe.Name = "btnRe";
            this.btnRe.Size = new System.Drawing.Size(69, 23);
            this.btnRe.TabIndex = 2;
            this.btnRe.Text = "Refresh";
            this.btnRe.UseVisualStyleBackColor = true;
            this.btnRe.Click += new System.EventHandler(this.btnRe_Click);
            // 
            // dateTimePickerP
            // 
            this.dateTimePickerP.Location = new System.Drawing.Point(6, 18);
            this.dateTimePickerP.Name = "dateTimePickerP";
            this.dateTimePickerP.Size = new System.Drawing.Size(136, 21);
            this.dateTimePickerP.TabIndex = 1;
            // 
            // dataGridViewP
            // 
            this.dataGridViewP.AllowUserToAddRows = false;
            this.dataGridViewP.AllowUserToDeleteRows = false;
            this.dataGridViewP.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewP.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridViewP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewP.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewP.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewP.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewP.Name = "dataGridViewP";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewP.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridViewP.RowTemplate.Height = 23;
            this.dataGridViewP.Size = new System.Drawing.Size(921, 332);
            this.dataGridViewP.TabIndex = 6;
            this.dataGridViewP.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewP_CellDoubleClick);
            this.dataGridViewP.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dataGridViewP_CellPainting);
            this.dataGridViewP.CellToolTipTextNeeded += new System.Windows.Forms.DataGridViewCellToolTipTextNeededEventHandler(this.dataGridViewP_CellToolTipTextNeeded);
            this.dataGridViewP.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridViewP_CellValidating);
            this.dataGridViewP.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridViewP_DataError);
            this.dataGridViewP.RowValidating += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridViewP_RowValidating);
            // 
            // FormManageOffline
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(921, 462);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.toolStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormManageOffline";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Forecast Offline";
            this.Load += new System.EventHandler(this.FormManageOffline_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewP)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.ToolStripButton printPreviewToolStripButton;
        private System.Windows.Forms.Button btnAll;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnPCF;
        private System.Windows.Forms.Button btnICF;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelC;
        private System.Windows.Forms.ToolStrip toolStrip;
        public System.Windows.Forms.ToolStripButton printToolStripButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnRe;
        private System.Windows.Forms.DateTimePicker dateTimePickerP;
        private System.Windows.Forms.DataGridView dataGridViewP;
        private System.Windows.Forms.ComboBox comboBoxIC;
        private System.Windows.Forms.ComboBox comboBoxPC;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btndisplayfilter;
        private System.Windows.Forms.CheckBox checkBoxBefore;
        private System.Windows.Forms.DateTimePicker dateTimePickerBefore;
        private System.Windows.Forms.CheckBox checkBoxTAC;
        private System.Windows.Forms.Button btndisplayfiltersection;
    }
}