namespace ExcelReconApp;

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
        this.tabControl1 = new System.Windows.Forms.TabControl();
        this.tabPageWorksheet1 = new System.Windows.Forms.TabPage();
        this.dataGridViewWorksheet1 = new System.Windows.Forms.DataGridView();
        this.panel1 = new System.Windows.Forms.Panel();
        this.checkedListBoxValueColumns1 = new System.Windows.Forms.CheckedListBox();
        this.label3 = new System.Windows.Forms.Label();
        this.comboBoxIdColumn1 = new System.Windows.Forms.ComboBox();
        this.label1 = new System.Windows.Forms.Label();
        this.btnLoadWorksheet1 = new System.Windows.Forms.Button();
        this.tabPageWorksheet2 = new System.Windows.Forms.TabPage();
        this.dataGridViewWorksheet2 = new System.Windows.Forms.DataGridView();
        this.panel2 = new System.Windows.Forms.Panel();
        this.checkedListBoxValueColumns2 = new System.Windows.Forms.CheckedListBox();
        this.label4 = new System.Windows.Forms.Label();
        this.comboBoxIdColumn2 = new System.Windows.Forms.ComboBox();
        this.label2 = new System.Windows.Forms.Label();
        this.btnLoadWorksheet2 = new System.Windows.Forms.Button();
        this.tabPageReconResults = new System.Windows.Forms.TabPage();
        this.dataGridViewReconResults = new System.Windows.Forms.DataGridView();
        this.panel3 = new System.Windows.Forms.Panel();
        this.btnRecon = new System.Windows.Forms.Button();
        this.tabControl1.SuspendLayout();
        this.tabPageWorksheet1.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewWorksheet1)).BeginInit();
        this.panel1.SuspendLayout();
        this.tabPageWorksheet2.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewWorksheet2)).BeginInit();
        this.panel2.SuspendLayout();
        this.tabPageReconResults.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewReconResults)).BeginInit();
        this.panel3.SuspendLayout();
        this.SuspendLayout();
        // 
        // tabControl1
        // 
        this.tabControl1.Controls.Add(this.tabPageWorksheet1);
        this.tabControl1.Controls.Add(this.tabPageWorksheet2);
        this.tabControl1.Controls.Add(this.tabPageReconResults);
        this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
        this.tabControl1.Location = new System.Drawing.Point(0, 0);
        this.tabControl1.Name = "tabControl1";
        this.tabControl1.SelectedIndex = 0;
        this.tabControl1.Size = new System.Drawing.Size(1200, 700);
        this.tabControl1.TabIndex = 0;
        // 
        // tabPageWorksheet1
        // 
        this.tabPageWorksheet1.Controls.Add(this.dataGridViewWorksheet1);
        this.tabPageWorksheet1.Controls.Add(this.panel1);
        this.tabPageWorksheet1.Location = new System.Drawing.Point(4, 29);
        this.tabPageWorksheet1.Name = "tabPageWorksheet1";
        this.tabPageWorksheet1.Padding = new System.Windows.Forms.Padding(3);
        this.tabPageWorksheet1.Size = new System.Drawing.Size(1192, 667);
        this.tabPageWorksheet1.TabIndex = 0;
        this.tabPageWorksheet1.Text = "Worksheet 1";
        this.tabPageWorksheet1.UseVisualStyleBackColor = true;
        // 
        // dataGridViewWorksheet1
        // 
        this.dataGridViewWorksheet1.AllowUserToAddRows = false;
        this.dataGridViewWorksheet1.AllowUserToDeleteRows = false;
        this.dataGridViewWorksheet1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridViewWorksheet1.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dataGridViewWorksheet1.Location = new System.Drawing.Point(3, 103);
        this.dataGridViewWorksheet1.Name = "dataGridViewWorksheet1";
        this.dataGridViewWorksheet1.ReadOnly = true;
        this.dataGridViewWorksheet1.RowHeadersWidth = 51;
        this.dataGridViewWorksheet1.Size = new System.Drawing.Size(1186, 561);
        this.dataGridViewWorksheet1.TabIndex = 1;
        // 
        // panel1
        // 
        this.panel1.Controls.Add(this.checkedListBoxValueColumns1);
        this.panel1.Controls.Add(this.label3);
        this.panel1.Controls.Add(this.comboBoxIdColumn1);
        this.panel1.Controls.Add(this.label1);
        this.panel1.Controls.Add(this.btnLoadWorksheet1);
        this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
        this.panel1.Location = new System.Drawing.Point(3, 3);
        this.panel1.Name = "panel1";
        this.panel1.Size = new System.Drawing.Size(1186, 100);
        this.panel1.TabIndex = 0;
        // 
        // checkedListBoxValueColumns1
        // 
        this.checkedListBoxValueColumns1.FormattingEnabled = true;
        this.checkedListBoxValueColumns1.Location = new System.Drawing.Point(450, 35);
        this.checkedListBoxValueColumns1.Name = "checkedListBoxValueColumns1";
        this.checkedListBoxValueColumns1.Size = new System.Drawing.Size(200, 48);
        this.checkedListBoxValueColumns1.TabIndex = 4;
        // 
        // label3
        // 
        this.label3.AutoSize = true;
        this.label3.Location = new System.Drawing.Point(450, 10);
        this.label3.Name = "label3";
        this.label3.Size = new System.Drawing.Size(110, 20);
        this.label3.TabIndex = 3;
        this.label3.Text = "Value Columns:";
        // 
        // comboBoxIdColumn1
        // 
        this.comboBoxIdColumn1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        this.comboBoxIdColumn1.FormattingEnabled = true;
        this.comboBoxIdColumn1.Location = new System.Drawing.Point(220, 35);
        this.comboBoxIdColumn1.Name = "comboBoxIdColumn1";
        this.comboBoxIdColumn1.Size = new System.Drawing.Size(200, 28);
        this.comboBoxIdColumn1.TabIndex = 2;
        // 
        // label1
        // 
        this.label1.AutoSize = true;
        this.label1.Location = new System.Drawing.Point(220, 10);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(85, 20);
        this.label1.TabIndex = 1;
        this.label1.Text = "ID Column:";
        // 
        // btnLoadWorksheet1
        // 
        this.btnLoadWorksheet1.Location = new System.Drawing.Point(10, 10);
        this.btnLoadWorksheet1.Name = "btnLoadWorksheet1";
        this.btnLoadWorksheet1.Size = new System.Drawing.Size(180, 80);
        this.btnLoadWorksheet1.TabIndex = 0;
        this.btnLoadWorksheet1.Text = "Load Worksheet 1";
        this.btnLoadWorksheet1.UseVisualStyleBackColor = true;
        this.btnLoadWorksheet1.Click += new System.EventHandler(this.btnLoadWorksheet1_Click);
        // 
        // tabPageWorksheet2
        // 
        this.tabPageWorksheet2.Controls.Add(this.dataGridViewWorksheet2);
        this.tabPageWorksheet2.Controls.Add(this.panel2);
        this.tabPageWorksheet2.Location = new System.Drawing.Point(4, 29);
        this.tabPageWorksheet2.Name = "tabPageWorksheet2";
        this.tabPageWorksheet2.Padding = new System.Windows.Forms.Padding(3);
        this.tabPageWorksheet2.Size = new System.Drawing.Size(1192, 667);
        this.tabPageWorksheet2.TabIndex = 1;
        this.tabPageWorksheet2.Text = "Worksheet 2";
        this.tabPageWorksheet2.UseVisualStyleBackColor = true;
        // 
        // dataGridViewWorksheet2
        // 
        this.dataGridViewWorksheet2.AllowUserToAddRows = false;
        this.dataGridViewWorksheet2.AllowUserToDeleteRows = false;
        this.dataGridViewWorksheet2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridViewWorksheet2.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dataGridViewWorksheet2.Location = new System.Drawing.Point(3, 103);
        this.dataGridViewWorksheet2.Name = "dataGridViewWorksheet2";
        this.dataGridViewWorksheet2.ReadOnly = true;
        this.dataGridViewWorksheet2.RowHeadersWidth = 51;
        this.dataGridViewWorksheet2.Size = new System.Drawing.Size(1186, 561);
        this.dataGridViewWorksheet2.TabIndex = 1;
        // 
        // panel2
        // 
        this.panel2.Controls.Add(this.checkedListBoxValueColumns2);
        this.panel2.Controls.Add(this.label4);
        this.panel2.Controls.Add(this.comboBoxIdColumn2);
        this.panel2.Controls.Add(this.label2);
        this.panel2.Controls.Add(this.btnLoadWorksheet2);
        this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
        this.panel2.Location = new System.Drawing.Point(3, 3);
        this.panel2.Name = "panel2";
        this.panel2.Size = new System.Drawing.Size(1186, 100);
        this.panel2.TabIndex = 0;
        // 
        // checkedListBoxValueColumns2
        // 
        this.checkedListBoxValueColumns2.FormattingEnabled = true;
        this.checkedListBoxValueColumns2.Location = new System.Drawing.Point(450, 35);
        this.checkedListBoxValueColumns2.Name = "checkedListBoxValueColumns2";
        this.checkedListBoxValueColumns2.Size = new System.Drawing.Size(200, 48);
        this.checkedListBoxValueColumns2.TabIndex = 4;
        // 
        // label4
        // 
        this.label4.AutoSize = true;
        this.label4.Location = new System.Drawing.Point(450, 10);
        this.label4.Name = "label4";
        this.label4.Size = new System.Drawing.Size(110, 20);
        this.label4.TabIndex = 3;
        this.label4.Text = "Value Columns:";
        // 
        // comboBoxIdColumn2
        // 
        this.comboBoxIdColumn2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        this.comboBoxIdColumn2.FormattingEnabled = true;
        this.comboBoxIdColumn2.Location = new System.Drawing.Point(220, 35);
        this.comboBoxIdColumn2.Name = "comboBoxIdColumn2";
        this.comboBoxIdColumn2.Size = new System.Drawing.Size(200, 28);
        this.comboBoxIdColumn2.TabIndex = 2;
        // 
        // label2
        // 
        this.label2.AutoSize = true;
        this.label2.Location = new System.Drawing.Point(220, 10);
        this.label2.Name = "label2";
        this.label2.Size = new System.Drawing.Size(85, 20);
        this.label2.TabIndex = 1;
        this.label2.Text = "ID Column:";
        // 
        // btnLoadWorksheet2
        // 
        this.btnLoadWorksheet2.Location = new System.Drawing.Point(10, 10);
        this.btnLoadWorksheet2.Name = "btnLoadWorksheet2";
        this.btnLoadWorksheet2.Size = new System.Drawing.Size(180, 80);
        this.btnLoadWorksheet2.TabIndex = 0;
        this.btnLoadWorksheet2.Text = "Load Worksheet 2";
        this.btnLoadWorksheet2.UseVisualStyleBackColor = true;
        this.btnLoadWorksheet2.Click += new System.EventHandler(this.btnLoadWorksheet2_Click);
        // 
        // tabPageReconResults
        // 
        this.tabPageReconResults.Controls.Add(this.dataGridViewReconResults);
        this.tabPageReconResults.Controls.Add(this.panel3);
        this.tabPageReconResults.Location = new System.Drawing.Point(4, 29);
        this.tabPageReconResults.Name = "tabPageReconResults";
        this.tabPageReconResults.Size = new System.Drawing.Size(1192, 667);
        this.tabPageReconResults.TabIndex = 2;
        this.tabPageReconResults.Text = "Reconciliation Results";
        this.tabPageReconResults.UseVisualStyleBackColor = true;
        // 
        // dataGridViewReconResults
        // 
        this.dataGridViewReconResults.AllowUserToAddRows = false;
        this.dataGridViewReconResults.AllowUserToDeleteRows = false;
        this.dataGridViewReconResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridViewReconResults.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dataGridViewReconResults.Location = new System.Drawing.Point(0, 60);
        this.dataGridViewReconResults.Name = "dataGridViewReconResults";
        this.dataGridViewReconResults.ReadOnly = true;
        this.dataGridViewReconResults.RowHeadersWidth = 51;
        this.dataGridViewReconResults.Size = new System.Drawing.Size(1192, 607);
        this.dataGridViewReconResults.TabIndex = 1;
        // 
        // panel3
        // 
        this.panel3.Controls.Add(this.btnRecon);
        this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
        this.panel3.Location = new System.Drawing.Point(0, 0);
        this.panel3.Name = "panel3";
        this.panel3.Size = new System.Drawing.Size(1192, 60);
        this.panel3.TabIndex = 0;
        // 
        // btnRecon
        // 
        this.btnRecon.Location = new System.Drawing.Point(10, 10);
        this.btnRecon.Name = "btnRecon";
        this.btnRecon.Size = new System.Drawing.Size(200, 40);
        this.btnRecon.TabIndex = 0;
        this.btnRecon.Text = "Run Reconciliation";
        this.btnRecon.UseVisualStyleBackColor = true;
        this.btnRecon.Click += new System.EventHandler(this.btnRecon_Click);
        // 
        // Form1
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(1200, 700);
        this.Controls.Add(this.tabControl1);
        this.Name = "Form1";
        this.Text = "Excel Reconciliation Tool";
        this.tabControl1.ResumeLayout(false);
        this.tabPageWorksheet1.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewWorksheet1)).EndInit();
        this.panel1.ResumeLayout(false);
        this.panel1.PerformLayout();
        this.tabPageWorksheet2.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewWorksheet2)).EndInit();
        this.panel2.ResumeLayout(false);
        this.panel2.PerformLayout();
        this.tabPageReconResults.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)(this.dataGridViewReconResults)).EndInit();
        this.panel3.ResumeLayout(false);
        this.ResumeLayout(false);
    }

    #endregion

    private System.Windows.Forms.TabControl tabControl1;
    private System.Windows.Forms.TabPage tabPageWorksheet1;
    private System.Windows.Forms.TabPage tabPageWorksheet2;
    private System.Windows.Forms.TabPage tabPageReconResults;
    private System.Windows.Forms.DataGridView dataGridViewWorksheet1;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.Button btnLoadWorksheet1;
    private System.Windows.Forms.DataGridView dataGridViewWorksheet2;
    private System.Windows.Forms.Panel panel2;
    private System.Windows.Forms.Button btnLoadWorksheet2;
    private System.Windows.Forms.DataGridView dataGridViewReconResults;
    private System.Windows.Forms.Panel panel3;
    private System.Windows.Forms.Button btnRecon;
    private System.Windows.Forms.ComboBox comboBoxIdColumn1;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.ComboBox comboBoxIdColumn2;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.CheckedListBox checkedListBoxValueColumns1;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.CheckedListBox checkedListBoxValueColumns2;
    private System.Windows.Forms.Label label4;
}
