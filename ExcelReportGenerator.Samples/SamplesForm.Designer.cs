namespace ExcelReportGenerator.Samples;

partial class SamplesForm
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
        this.lblReport = new System.Windows.Forms.Label();
        this.cmbReports = new System.Windows.Forms.ComboBox();
        this.btnRun = new System.Windows.Forms.Button();
        this.txtOutputFolder = new System.Windows.Forms.TextBox();
        this.lblOutputFolder = new System.Windows.Forms.Label();
        this.progressBar = new System.Windows.Forms.ProgressBar();
        this.SuspendLayout();
        // 
        // lblReport
        // 
        this.lblReport.AutoSize = true;
        this.lblReport.Location = new System.Drawing.Point(14, 10);
        this.lblReport.Name = "lblReport";
        this.lblReport.Size = new System.Drawing.Size(45, 15);
        this.lblReport.TabIndex = 0;
        this.lblReport.Text = "Report:";
        // 
        // cmbReports
        // 
        this.cmbReports.DisplayMember = "Name";
        this.cmbReports.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
        this.cmbReports.FormattingEnabled = true;
        this.cmbReports.Location = new System.Drawing.Point(76, 7);
        this.cmbReports.Name = "cmbReports";
        this.cmbReports.Size = new System.Drawing.Size(359, 23);
        this.cmbReports.TabIndex = 1;
        this.cmbReports.ValueMember = "Name";
        // 
        // btnRun
        // 
        this.btnRun.Location = new System.Drawing.Point(348, 68);
        this.btnRun.Name = "btnRun";
        this.btnRun.Size = new System.Drawing.Size(87, 27);
        this.btnRun.TabIndex = 2;
        this.btnRun.Text = "Run";
        this.btnRun.UseVisualStyleBackColor = true;
        this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
        // 
        // txtOutputFolder
        // 
        this.txtOutputFolder.Location = new System.Drawing.Point(104, 38);
        this.txtOutputFolder.Name = "txtOutputFolder";
        this.txtOutputFolder.Size = new System.Drawing.Size(331, 23);
        this.txtOutputFolder.TabIndex = 3;
        // 
        // lblOutputFolder
        // 
        this.lblOutputFolder.AutoSize = true;
        this.lblOutputFolder.Location = new System.Drawing.Point(14, 43);
        this.lblOutputFolder.Name = "lblOutputFolder";
        this.lblOutputFolder.Size = new System.Drawing.Size(82, 15);
        this.lblOutputFolder.TabIndex = 4;
        this.lblOutputFolder.Text = "Output folder:";
        // 
        // progressBar
        // 
        this.progressBar.Location = new System.Drawing.Point(14, 68);
        this.progressBar.Name = "progressBar";
        this.progressBar.Size = new System.Drawing.Size(327, 27);
        this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
        this.progressBar.TabIndex = 5;
        this.progressBar.Visible = false;
        // 
        // SamplesForm
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(449, 104);
        this.Controls.Add(this.progressBar);
        this.Controls.Add(this.lblOutputFolder);
        this.Controls.Add(this.txtOutputFolder);
        this.Controls.Add(this.btnRun);
        this.Controls.Add(this.cmbReports);
        this.Controls.Add(this.lblReport);
        this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
        this.MaximizeBox = false;
        this.Name = "SamplesForm";
        this.ShowIcon = false;
        this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        this.Text = "Samples";
        this.Load += new System.EventHandler(this.SamplesForm_Load);
        this.ResumeLayout(false);
        this.PerformLayout();
    }

    #endregion

    private System.Windows.Forms.Label lblReport;
    private System.Windows.Forms.ComboBox cmbReports;
    private System.Windows.Forms.Button btnRun;
    private System.Windows.Forms.TextBox txtOutputFolder;
    private System.Windows.Forms.Label lblOutputFolder;
    private System.Windows.Forms.ProgressBar progressBar;
}