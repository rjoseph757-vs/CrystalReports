namespace WindowsFormsAppCrystalReport2021
{
    partial class frmReportViewer
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
            //this.AdbookList1 = new WindowsFormsAppCrystalReport2021.Reports.AdbookList();
            //this.AdbookList2 = new WindowsFormsAppCrystalReport2021.AdbookList();
            this.crystalReportViewer1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            //this.CrystalReport11 = new WindowsFormsAppCrystalReport2021.CrystalReport1();
            this.AdbookList3 = new WindowsFormsAppCrystalReport2021.AdbookList();
            this.SuspendLayout();
            // 
            // crystalReportViewer1
            // 
            this.crystalReportViewer1.ActiveViewIndex = 0;
            this.crystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default;
            this.crystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.crystalReportViewer1.Location = new System.Drawing.Point(0, 0);
            this.crystalReportViewer1.Name = "crystalReportViewer1";
            this.crystalReportViewer1.ReportSource = this.AdbookList3;
            this.crystalReportViewer1.Size = new System.Drawing.Size(1587, 1135);
            this.crystalReportViewer1.TabIndex = 0;
            // 
            // frmReportViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1587, 1135);
            this.Controls.Add(this.crystalReportViewer1);
            this.Name = "frmReportViewer";
            this.Text = "frmReportViewer";
            this.ResumeLayout(false);

        }

        #endregion

        private CrystalDecisions.Windows.Forms.CrystalReportViewer crystalReportViewer1;
        //private AdbookList AdbookList2;
        //private Reports.AdbookList AdbookList1;
        //private CrystalReport1 CrystalReport11;
        private AdbookList AdbookList3;
    }
}