namespace QcGoldArchive
{
    partial class ArchiveFrm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ArchiveFrm));
            this.csvdata = new System.Windows.Forms.DataGridView();
            this.btnShowAll = new System.Windows.Forms.Button();
            this.btnFilter = new System.Windows.Forms.Button();
            this.combofilter = new System.Windows.Forms.ComboBox();
            this.Archivehead = new System.Windows.Forms.Label();
            this.label122 = new System.Windows.Forms.Label();
            this.txtsearch = new System.Windows.Forms.TextBox();
            this.selectdiffcsv = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.MenuStrip = new System.Windows.Forms.ToolStrip();
            this.generate = new System.Windows.Forms.ToolStripButton();
            this.csvselect = new System.Windows.Forms.ToolStripButton();
            this.Arc_cancel = new System.Windows.Forms.ToolStripButton();
            this.toolTipTest = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.csvdata)).BeginInit();
            this.MenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // csvdata
            // 
            this.csvdata.AllowUserToDeleteRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.csvdata.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.csvdata.ColumnHeadersHeight = 33;
            this.csvdata.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.csvdata.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.csvdata.Location = new System.Drawing.Point(12, 111);
            this.csvdata.Name = "csvdata";
            this.csvdata.ReadOnly = true;
            this.csvdata.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 1.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.csvdata.RowHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.csvdata.RowHeadersWidth = 51;
            this.csvdata.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.csvdata.Size = new System.Drawing.Size(799, 392);
            this.csvdata.TabIndex = 0;
            this.csvdata.VirtualMode = true;
            this.csvdata.MultiSelectChanged += new System.EventHandler(this.BtnFilter_Click);
            this.csvdata.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.csvdata_CellClick);
            this.csvdata.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.csvdata_CellContentClick_1);
            this.csvdata.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.Csvdata_CellValueChanged);
            this.csvdata.CurrentCellDirtyStateChanged += new System.EventHandler(this.csvdata_CurrentCellDirtyStateChanged);
            this.csvdata.RowDirtyStateNeeded += new System.Windows.Forms.QuestionEventHandler(this.csvdata_RowDirtyStateNeeded);
            this.csvdata.RowEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.csvdata_RowEnter);
            this.csvdata.Click += new System.EventHandler(this.csvdata_Click_1);
            this.csvdata.DragDrop += new System.Windows.Forms.DragEventHandler(this.csvdata_DragDrop);
            this.csvdata.DragEnter += new System.Windows.Forms.DragEventHandler(this.csvdata_DragEnter);
            this.csvdata.DragLeave += new System.EventHandler(this.csvdata_DragDrop);
            this.csvdata.DoubleClick += new System.EventHandler(this.csvdata_DoubleClick_1);
            this.csvdata.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Csvdata_KeyDown);
            this.csvdata.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Csvdata_KeyPress);
            this.csvdata.KeyUp += new System.Windows.Forms.KeyEventHandler(this.csvdata_KeyUp);
            this.csvdata.MouseClick += new System.Windows.Forms.MouseEventHandler(this.csvdata_MouseClick);
            this.csvdata.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.csvdata_MouseDoubleClick);
            this.csvdata.MouseUp += new System.Windows.Forms.MouseEventHandler(this.Csvdata_MouseUp);
            // 
            // btnShowAll
            // 
            this.btnShowAll.Location = new System.Drawing.Point(489, 54);
            this.btnShowAll.Name = "btnShowAll";
            this.btnShowAll.Size = new System.Drawing.Size(149, 26);
            this.btnShowAll.TabIndex = 16;
            this.btnShowAll.Text = "Clear Filter";
            this.btnShowAll.UseVisualStyleBackColor = true;
            this.btnShowAll.Click += new System.EventHandler(this.BtnShowAll_Click);
            // 
            // btnFilter
            // 
            this.btnFilter.Location = new System.Drawing.Point(395, 54);
            this.btnFilter.Name = "btnFilter";
            this.btnFilter.Size = new System.Drawing.Size(84, 26);
            this.btnFilter.TabIndex = 15;
            this.btnFilter.Text = "Search";
            this.btnFilter.UseVisualStyleBackColor = true;
            this.btnFilter.Click += new System.EventHandler(this.BtnFilter_Click);
            // 
            // combofilter
            // 
            this.combofilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.combofilter.FormattingEnabled = true;
            this.combofilter.Location = new System.Drawing.Point(100, 58);
            this.combofilter.Name = "combofilter";
            this.combofilter.Size = new System.Drawing.Size(121, 21);
            this.combofilter.TabIndex = 17;
            this.combofilter.SelectedIndexChanged += new System.EventHandler(this.Combofilter_SelectedIndexChanged);
            // 
            // Archivehead
            // 
            this.Archivehead.AutoSize = true;
            this.Archivehead.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Archivehead.Location = new System.Drawing.Point(12, 19);
            this.Archivehead.Name = "Archivehead";
            this.Archivehead.Size = new System.Drawing.Size(174, 20);
            this.Archivehead.TabIndex = 18;
            this.Archivehead.Text = "Archive Test Results";
            // 
            // label122
            // 
            this.label122.AutoSize = true;
            this.label122.Location = new System.Drawing.Point(14, 61);
            this.label122.Name = "label122";
            this.label122.Size = new System.Drawing.Size(58, 13);
            this.label122.TabIndex = 19;
            this.label122.Text = "Search by:";
            // 
            // txtsearch
            // 
            this.txtsearch.Location = new System.Drawing.Point(232, 58);
            this.txtsearch.Name = "txtsearch";
            this.txtsearch.Size = new System.Drawing.Size(151, 20);
            this.txtsearch.TabIndex = 20;
            this.txtsearch.TextChanged += new System.EventHandler(this.Txtsearch_TextChanged_1);
            this.txtsearch.Enter += new System.EventHandler(this.Txtsearch_Enter);
            this.txtsearch.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Txtsearch_KeyDown);
            this.txtsearch.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.Txtsearch_KeyPress);
            this.txtsearch.Leave += new System.EventHandler(this.Txtsearch_Leave);
            // 
            // selectdiffcsv
            // 
            this.selectdiffcsv.Location = new System.Drawing.Point(690, 52);
            this.selectdiffcsv.Name = "selectdiffcsv";
            this.selectdiffcsv.Size = new System.Drawing.Size(121, 26);
            this.selectdiffcsv.TabIndex = 22;
            this.selectdiffcsv.Text = "Select CSV file";
            this.selectdiffcsv.UseVisualStyleBackColor = true;
            this.selectdiffcsv.Visible = false;
            this.selectdiffcsv.Click += new System.EventHandler(this.Selectdiffcsv_Click_1);
            // 
            // MenuStrip
            // 
            this.MenuStrip.AutoSize = false;
            this.MenuStrip.BackColor = System.Drawing.SystemColors.Control;
            this.MenuStrip.Dock = System.Windows.Forms.DockStyle.None;
            this.MenuStrip.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.MenuStrip.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.MenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.generate,
            this.csvselect,
            this.Arc_cancel});
            this.MenuStrip.Location = new System.Drawing.Point(9, 515);
            this.MenuStrip.Name = "MenuStrip";
            this.MenuStrip.Size = new System.Drawing.Size(802, 31);
            this.MenuStrip.TabIndex = 27;
            this.MenuStrip.Text = "Menu";
            // 
            // generate
            // 
            this.generate.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.generate.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.generate.Enabled = false;
            this.generate.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.generate.Margin = new System.Windows.Forms.Padding(230, 1, 0, 2);
            this.generate.Name = "generate";
            this.generate.Padding = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.generate.Size = new System.Drawing.Size(74, 28);
            this.generate.Text = "Average";
            this.generate.ToolTipText = "generate";
            this.generate.Click += new System.EventHandler(this.Average_Click);
            this.generate.MouseHover += new System.EventHandler(this.generate_MouseHover_1);
            this.generate.MouseMove += new System.Windows.Forms.MouseEventHandler(this.generate_MouseMove_1);
            // 
            // csvselect
            // 
            this.csvselect.BackColor = System.Drawing.SystemColors.ButtonShadow;
            this.csvselect.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.csvselect.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.csvselect.Margin = new System.Windows.Forms.Padding(50, 1, 0, 2);
            this.csvselect.Name = "csvselect";
            this.csvselect.Padding = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.csvselect.Size = new System.Drawing.Size(75, 28);
            this.csvselect.Text = "Confirm";
            this.csvselect.Click += new System.EventHandler(this.Confirm_Click);
            // 
            // Arc_cancel
            // 
            this.Arc_cancel.BackColor = System.Drawing.SystemColors.ControlDark;
            this.Arc_cancel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.Arc_cancel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Arc_cancel.Margin = new System.Windows.Forms.Padding(50, 1, 0, 2);
            this.Arc_cancel.Name = "Arc_cancel";
            this.Arc_cancel.Padding = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.Arc_cancel.Size = new System.Drawing.Size(67, 28);
            this.Arc_cancel.Text = "Cancel";
            this.Arc_cancel.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // toolTipTest
            // 
            this.toolTipTest.ToolTipTitle = "This is the Title";
            // 
            // ArchiveFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(830, 555);
            this.Controls.Add(this.MenuStrip);
            this.Controls.Add(this.selectdiffcsv);
            this.Controls.Add(this.txtsearch);
            this.Controls.Add(this.label122);
            this.Controls.Add(this.Archivehead);
            this.Controls.Add(this.combofilter);
            this.Controls.Add(this.btnShowAll);
            this.Controls.Add(this.btnFilter);
            this.Controls.Add(this.csvdata);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ArchiveFrm";
            this.Text = "DesignExample";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ArchiveFrm_FormClosed);
            this.Load += new System.EventHandler(this.DesignExample_Load);
            this.Leave += new System.EventHandler(this.ArchiveFrm_Leave);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ArchiveFrm_MouseMove);
            ((System.ComponentModel.ISupportInitialize)(this.csvdata)).EndInit();
            this.MenuStrip.ResumeLayout(false);
            this.MenuStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView csvdata;
        private System.Windows.Forms.Button btnShowAll;
        private System.Windows.Forms.Button btnFilter;
        private System.Windows.Forms.ComboBox combofilter;
        private System.Windows.Forms.Label Archivehead;
        private System.Windows.Forms.Label label122;
        private System.Windows.Forms.TextBox txtsearch;
        private System.Windows.Forms.Button selectdiffcsv;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolStrip MenuStrip;
        public System.Windows.Forms.ToolStripButton generate;
        private System.Windows.Forms.ToolStripButton csvselect;
        private System.Windows.Forms.ToolTip toolTipTest;
        private System.Windows.Forms.ToolStripButton Arc_cancel;
    }
}