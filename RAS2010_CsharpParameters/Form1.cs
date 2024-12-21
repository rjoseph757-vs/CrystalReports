using System.Reflection;
using System.Runtime.InteropServices;
using System;
using System.IO;
using System.Xml;
using System.Drawing;
using System.Drawing.Printing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Threading;
using System.Diagnostics;
using System.Globalization;
using ADODB;   
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.ReportAppServer;
using CrystalDecisions.ReportAppServer.ClientDoc;
using CrystalDecisions.ReportAppServer.Controllers;
using CrystalDecisions.ReportAppServer.ReportDefModel;
using CrystalDecisions.ReportAppServer.CommonControls;
using CrystalDecisions.ReportAppServer.CommLayer;
using CrystalDecisions.ReportAppServer.CommonObjectModel;
using CrystalDecisions.ReportAppServer.ObjectFactory;
using CrystalDecisions.ReportAppServer.Prompting;
using System.Data.OleDb;
using CrystalDecisions.ReportAppServer.DataSetConversion;
using CrystalDecisions.ReportAppServer.DataDefModel;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Windows.Forms;
using CrystalDecisions.ReportAppServer.XmlSerialize;
using System.Timers;

namespace Unmanaged_RAS10_CSharp_Printers
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class frmMain : System.Windows.Forms.Form
	{
        CrystalDecisions.CrystalReports.Engine.ReportDocument rpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc;

        CrystalDecisions.CrystalReports.Engine.ReportDocument rpt1 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
        CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc1;

        CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument reportClientDocument;

        CrystalDecisions.ReportAppServer.ReportDefModel.EromPageSettings eromPageSettings;
        CrystalDecisions.ReportAppServer.ReportDefModel.EromPrinterSettings eromPrinterSettings;

        #region CR Labels
        // this line is used with crystalras.exe running
        //ReportClientDocument rptClientDoc;
        private System.Windows.Forms.Button btnOpenReport;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Button btnSaveRptAs;
        private Button ViewReport;
        private Button SetData;
        private Button SetParam;
        private Button Export;
        private Button SetLogonInfo;
        private Button ReplaceConnection;
        private CrystalReportViewer crystalReportViewer1;
        private Button btnCloserpt;
        private System.Windows.Forms.Label ReportVersion;
        private TextBox btnReportVersion;
        private System.Windows.Forms.Label RecordSelectionForm;
        private TextBox btnRecordSelectionForm;
        private TextBox btnReportName;
        private System.Windows.Forms.Label ReportName;
        private System.Windows.Forms.Label rtnSQLStatement;
        private System.Windows.Forms.Label rtnReportObjects;
        internal ComboBox ReportObjectComboBox1;
        private Button SetPCDatabase;
        private System.Windows.Forms.Label label8;
        private TextBox btnCount;
        private System.Windows.Forms.Label label11;
        private TextBox btnReportKind;
        private TextBox btnRecordCount;
        private System.Windows.Forms.Label label12;
        private CheckBox chkUseRAS;
        private System.Windows.Forms.Label label15;
        private TextBox btnDBDriver;
        private RichTextBox btnSQLStatement;
        private TextBox btrDataFile;
        private TextBox btrSearchPath;
        private Button btnRasOpen;
        private Label btrRuntimeVersion;
        private TextBox txtRuntimeVersion;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
        private CheckBox btrVerifyDatabase;
        #endregion CR Labels

        string CRVer;
        private TextBox btrFileLocation;
        private GroupBox groupBox1;
        private TextBox btrPassword;
        private TextBox lblRPTRev;
        private Label label18;
        private Label label19;
        private Label label20;
        private Label label21;
        private Label label22;
        private Label label;
        private ComboBox cbLastSaveHistory;
        private Button btrRefreshViewer;
        private TextBox btnErrorCount;
        private Label label23;
        private CheckBox chkTrusted;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel tsslThread;
        private ToolStripStatusLabel tsslHandles;
        private ToolStripStatusLabel tsslPrivateBytes;
        private Button SetToXML;
        private CheckBox btnHasSavedData;
        private RichTextBox btnReportObjects;
        private Label label31;
        private ComboBox LstInterpolationMode;
        private Button CommandTable;
        float isMetric = 0;

        public void createTimer()
        {
            System.Windows.Forms.Timer timerKeepTrack = new System.Windows.Forms.Timer();
            timerKeepTrack.Interval = 1000;
        }

		public frmMain()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

            //PerformanceCounter pcPrivateBytes = new PerformanceCounter();
            //PerformanceCounter pcHandles = new PerformanceCounter();
            //PerformanceCounter pcThreads = new PerformanceCounter();

            //pcThreads.CategoryName = "Threads";
            //pcThreads.InstanceLifetime = PerformanceCounterInstanceLifetime.Process;
            //pcThreads.MachineName = Name;
            //pcThreads.InstanceName = Application.ProductName;
            //pcHandles.CategoryName = "Handles";
            //pcHandles.InstanceLifetime = PerformanceCounterInstanceLifetime.Process;
            //pcHandles.MachineName = Name;
            //pcHandles.InstanceName = Application.ProductName;
            //pcPrivateBytes.CategoryName = "pcPrivateBytesName";
            //pcPrivateBytes.InstanceLifetime = PerformanceCounterInstanceLifetime.Process;
            //pcPrivateBytes.MachineName = Name;
            //pcPrivateBytes.InstanceName = Application.ProductName;


            //pcThreads.CounterName = "Thread Count";
            //pcHandles.CounterName = "Handle Count";
            //pcPrivateBytes.CounterName = "Private Bytes";

            //tsslThread.Text = string.Format("Threads: {0}", pcThreads.NextValue().ToString());
            //tsslHandles.Text = string.Format("Handles: {0}", pcHandles.NextValue().ToString());

            //// Divide the return of Private Bytes to something we will recognize.
            //tsslPrivateBytes.Text = string.Format("Private Bytes: {0}", (pcPrivateBytes.NextValue() / 1000).ToString());

            LstInterpolationMode.Enabled = true;
            Array CRinterpolationMode = Enum.GetValues(typeof(System.Drawing.Drawing2D.InterpolationMode));
            foreach (object obj in CRinterpolationMode)
            {
                //CRInterpolMode.GetTypeCode(CRinterpolationMode);
                LstInterpolationMode.Items.Add(obj);
            }
            LstInterpolationMode.SelectedItem = System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor;

            if (System.Globalization.RegionInfo.CurrentRegion.IsMetric)
                isMetric = 567;
            else
                isMetric = 1440;

            foreach (Assembly MyVerison in AppDomain.CurrentDomain.GetAssemblies())
            {
                if (MyVerison.FullName.Substring(0, 38) == "CrystalDecisions.CrystalReports.Engine")
                {
                    //File:             C:\Windows\assembly\GAC_MSIL\CrystalDecisions.CrystalReports.Engine\13.0.2000.0__692fbea5521e1304\CrystalDecisions.CrystalReports.Engine.dll
                    //InternalName:     Crystal Reports
                    //OriginalFilename: 
                    //FileVersion:      13.0.9.1312
                    //FileDescription:  Crystal Reports
                    //Product:          SBOP Crystal Reports
                    //ProductVersion:   13.0.9.1312
                    //Debug:            False
                    //Patched:          False
                    //PreRelease:       False
                    //PrivateBuild:     False
                    //SpecialBuild:     False
                    //Language:         English (United States)

                    System.Diagnostics.FileVersionInfo fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(MyVerison.Location);
                    txtRuntimeVersion.Text += fileVersionInfo.FileVersion.ToString();
                    // check if CrsytalDecisions.Enterprise dll's can be loaded ( Anything but Cortez - managed reporting )
                    if (fileVersionInfo.FileVersion.Substring(0, 2) == "13")
                    {
                        btnRasOpen.Enabled = false;
                    }
                    CRVer = fileVersionInfo.FileVersion.Substring(0, 2);
                    return;
                }
            }

            //VB code
            //Imports CrystalDecisions.CrystalReports.Engine
            //Imports CrystalDecisions.Shared
            //Imports System.Reflection
            //Imports System.Runtime.InteropServices

            //Public Class Form1

            //    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load

            //        For Each MyVerison As Assembly In AppDomain.CurrentDomain.GetAssemblies()
            //            If MyVerison.FullName.Substring(0, 38) = "CrystalDecisions.CrystalReports.Engine" Then
            //                'File:             C:\Windows\assembly\GAC_MSIL\CrystalDecisions.CrystalReports.Engine\13.0.2000.0__692fbea5521e1304\CrystalDecisions.CrystalReports.Engine.dll
            //                'InternalName:     Crystal Reports
            //                'OriginalFilename: 
            //                'FileVersion:      13.0.9.1312
            //                'FileDescription:  Crystal Reports
            //                'Product:          SBOP Crystal Reports
            //                'ProductVersion:   13.0.9.1312
            //                'Debug:            False
            //                'Patched:          False
            //                'PreRelease:       False
            //                'PrivateBuild:     False
            //                'SpecialBuild:     False
            //                'Language:         English (United States)

            //                Dim fileVersionInfo As System.Diagnostics.FileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(MyVerison.Location)
            //                MessageBox.Show(fileVersionInfo.FileVersion.ToString())

            //                Return
            //            End If
            //        Next


            //        Dim Report As New CrystalDecisions.CrystalReports.Engine.ReportDocument
            //        Report.Load("C:\reports\formulas.rpt")
            //        CrystalReportViewer1.ReportSource = Report
            //    End Sub

            //End Class 


			// 
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>

		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnOpenReport = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnSaveRptAs = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.ViewReport = new System.Windows.Forms.Button();
            this.SetData = new System.Windows.Forms.Button();
            this.SetParam = new System.Windows.Forms.Button();
            this.Export = new System.Windows.Forms.Button();
            this.SetLogonInfo = new System.Windows.Forms.Button();
            this.ReplaceConnection = new System.Windows.Forms.Button();
            this.crystalReportViewer1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.btnCloserpt = new System.Windows.Forms.Button();
            this.ReportVersion = new System.Windows.Forms.Label();
            this.btnReportVersion = new System.Windows.Forms.TextBox();
            this.RecordSelectionForm = new System.Windows.Forms.Label();
            this.btnRecordSelectionForm = new System.Windows.Forms.TextBox();
            this.btnReportName = new System.Windows.Forms.TextBox();
            this.ReportName = new System.Windows.Forms.Label();
            this.rtnSQLStatement = new System.Windows.Forms.Label();
            this.rtnReportObjects = new System.Windows.Forms.Label();
            this.ReportObjectComboBox1 = new System.Windows.Forms.ComboBox();
            this.SetPCDatabase = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.btnCount = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.btnReportKind = new System.Windows.Forms.TextBox();
            this.btnRecordCount = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.chkUseRAS = new System.Windows.Forms.CheckBox();
            this.label15 = new System.Windows.Forms.Label();
            this.btnDBDriver = new System.Windows.Forms.TextBox();
            this.btnSQLStatement = new System.Windows.Forms.RichTextBox();
            this.btrDataFile = new System.Windows.Forms.TextBox();
            this.btrSearchPath = new System.Windows.Forms.TextBox();
            this.btnRasOpen = new System.Windows.Forms.Button();
            this.btrRuntimeVersion = new System.Windows.Forms.Label();
            this.txtRuntimeVersion = new System.Windows.Forms.TextBox();
            this.btrVerifyDatabase = new System.Windows.Forms.CheckBox();
            this.btrFileLocation = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.SetToXML = new System.Windows.Forms.Button();
            this.btrPassword = new System.Windows.Forms.TextBox();
            this.lblRPTRev = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label = new System.Windows.Forms.Label();
            this.cbLastSaveHistory = new System.Windows.Forms.ComboBox();
            this.btrRefreshViewer = new System.Windows.Forms.Button();
            this.btnErrorCount = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.chkTrusted = new System.Windows.Forms.CheckBox();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.tsslThread = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsslHandles = new System.Windows.Forms.ToolStripStatusLabel();
            this.tsslPrivateBytes = new System.Windows.Forms.ToolStripStatusLabel();
            this.btnHasSavedData = new System.Windows.Forms.CheckBox();
            this.btnReportObjects = new System.Windows.Forms.RichTextBox();
            this.label31 = new System.Windows.Forms.Label();
            this.LstInterpolationMode = new System.Windows.Forms.ComboBox();
            this.CommandTable = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpenReport
            // 
            this.btnOpenReport.Location = new System.Drawing.Point(8, 7);
            this.btnOpenReport.Name = "btnOpenReport";
            this.btnOpenReport.Size = new System.Drawing.Size(79, 24);
            this.btnOpenReport.TabIndex = 11;
            this.btnOpenReport.Text = "Open rpt...";
            this.btnOpenReport.Click += new System.EventHandler(this.btnOpenReport_Click);
            // 
            // btnSaveRptAs
            // 
            this.btnSaveRptAs.Enabled = false;
            this.btnSaveRptAs.Location = new System.Drawing.Point(332, 7);
            this.btnSaveRptAs.Name = "btnSaveRptAs";
            this.btnSaveRptAs.Size = new System.Drawing.Size(79, 24);
            this.btnSaveRptAs.TabIndex = 12;
            this.btnSaveRptAs.Text = "save as...";
            this.btnSaveRptAs.Click += new System.EventHandler(this.btnSaveReportAs_Click);
            // 
            // ViewReport
            // 
            this.ViewReport.Enabled = false;
            this.ViewReport.Location = new System.Drawing.Point(170, 7);
            this.ViewReport.Name = "ViewReport";
            this.ViewReport.Size = new System.Drawing.Size(79, 24);
            this.ViewReport.TabIndex = 15;
            this.ViewReport.Text = "View Report";
            this.ViewReport.Click += new System.EventHandler(this.ViewReport_Click);
            // 
            // SetData
            // 
            this.SetData.Location = new System.Drawing.Point(887, 36);
            this.SetData.Name = "SetData";
            this.SetData.Size = new System.Drawing.Size(75, 23);
            this.SetData.TabIndex = 17;
            this.SetData.Text = "Set Data";
            this.SetData.Click += new System.EventHandler(this.SetData_Click);
            // 
            // SetParam
            // 
            this.SetParam.Location = new System.Drawing.Point(887, 64);
            this.SetParam.Name = "SetParam";
            this.SetParam.Size = new System.Drawing.Size(75, 23);
            this.SetParam.TabIndex = 18;
            this.SetParam.Text = "Set Param";
            this.SetParam.UseVisualStyleBackColor = true;
            this.SetParam.Click += new System.EventHandler(this.SetParam_Click);
            // 
            // Export
            // 
            this.Export.Location = new System.Drawing.Point(887, 93);
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(75, 23);
            this.Export.TabIndex = 19;
            this.Export.Text = "Export";
            this.Export.UseVisualStyleBackColor = true;
            this.Export.Click += new System.EventHandler(this.Export_Click);
            // 
            // SetLogonInfo
            // 
            this.SetLogonInfo.Location = new System.Drawing.Point(887, 8);
            this.SetLogonInfo.Name = "SetLogonInfo";
            this.SetLogonInfo.Size = new System.Drawing.Size(75, 23);
            this.SetLogonInfo.TabIndex = 21;
            this.SetLogonInfo.Text = "Set Logon";
            this.SetLogonInfo.UseVisualStyleBackColor = true;
            this.SetLogonInfo.Click += new System.EventHandler(this.SetLogonInfo_Click);
            // 
            // ReplaceConnection
            // 
            this.ReplaceConnection.Location = new System.Drawing.Point(1084, 8);
            this.ReplaceConnection.Name = "ReplaceConnection";
            this.ReplaceConnection.Size = new System.Drawing.Size(98, 23);
            this.ReplaceConnection.TabIndex = 33;
            this.ReplaceConnection.Text = "Replace Conn";
            this.ReplaceConnection.UseVisualStyleBackColor = true;
            this.ReplaceConnection.Click += new System.EventHandler(this.ReplaceConnection_Click);
            // 
            // crystalReportViewer1
            // 
            this.crystalReportViewer1.ActiveViewIndex = -1;
            this.crystalReportViewer1.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            this.crystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default;
            this.crystalReportViewer1.Location = new System.Drawing.Point(8, 150);
            this.crystalReportViewer1.Name = "crystalReportViewer1";
            this.crystalReportViewer1.Size = new System.Drawing.Size(828, 717);
            this.crystalReportViewer1.TabIndex = 11;
            this.crystalReportViewer1.ViewZoom += new CrystalDecisions.Windows.Forms.ZoomEventHandler(this.crystalReportViewer1_ViewZoom);
            this.crystalReportViewer1.Navigate += new CrystalDecisions.Windows.Forms.NavigateEventHandler(this.crystalReportViewer1_Navigate);
            this.crystalReportViewer1.ClickPage += new CrystalDecisions.Windows.Forms.PageMouseEventHandler(this.crystalReportViewer1_ClickPage_1);
            this.crystalReportViewer1.DoubleClickPage += new CrystalDecisions.Windows.Forms.PageMouseEventHandler(this.crystalReportViewer1_DoubleClickPage);
            this.crystalReportViewer1.Drill += new CrystalDecisions.Windows.Forms.DrillEventHandler(this.crystalReportViewer1_Drill_1);
            this.crystalReportViewer1.DrillDownSubreport += new CrystalDecisions.Windows.Forms.DrillSubreportEventHandler(this.crystalReportViewer1_DrillDownSubreport);
            this.crystalReportViewer1.Error += new CrystalDecisions.Windows.Forms.ExceptionEventHandler(this.crystalReportViewer1_Error);
            this.crystalReportViewer1.EnabledChanged += new System.EventHandler(this.crystalReportViewer1_EnabledChanged);
            this.crystalReportViewer1.RegionChanged += new System.EventHandler(this.crystalReportViewer1_RegionChanged);
            this.crystalReportViewer1.VisibleChanged += new System.EventHandler(this.crystalReportViewer1_VisibleChanged);
            this.crystalReportViewer1.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.crystalReportViewer1_HelpRequested);
            this.crystalReportViewer1.Paint += new System.Windows.Forms.PaintEventHandler(this.crystalReportViewer1_Paint);
            this.crystalReportViewer1.MouseClick += new System.Windows.Forms.MouseEventHandler(this.crystalReportViewer1_MouseClick_1);
            this.crystalReportViewer1.Validating += new System.ComponentModel.CancelEventHandler(this.crystalReportViewer1_Validating);
            this.crystalReportViewer1.Validated += new System.EventHandler(this.crystalReportViewer1_Validated);
            this.crystalReportViewer1.ImeModeChanged += new System.EventHandler(this.crystalReportViewer1_ImeModeChanged);
            // 
            // btnCloserpt
            // 
            this.btnCloserpt.Enabled = false;
            this.btnCloserpt.Location = new System.Drawing.Point(89, 7);
            this.btnCloserpt.Name = "btnCloserpt";
            this.btnCloserpt.Size = new System.Drawing.Size(79, 24);
            this.btnCloserpt.TabIndex = 38;
            this.btnCloserpt.Text = "Close rpt...";
            this.btnCloserpt.UseVisualStyleBackColor = true;
            this.btnCloserpt.Click += new System.EventHandler(this.Closerpt_Click);
            // 
            // ReportVersion
            // 
            this.ReportVersion.AutoSize = true;
            this.ReportVersion.Location = new System.Drawing.Point(13, 59);
            this.ReportVersion.Name = "ReportVersion";
            this.ReportVersion.Size = new System.Drawing.Size(83, 13);
            this.ReportVersion.TabIndex = 40;
            this.ReportVersion.Text = "Report Version: ";
            // 
            // btnReportVersion
            // 
            this.btnReportVersion.Location = new System.Drawing.Point(97, 56);
            this.btnReportVersion.Name = "btnReportVersion";
            this.btnReportVersion.Size = new System.Drawing.Size(41, 20);
            this.btnReportVersion.TabIndex = 41;
            // 
            // RecordSelectionForm
            // 
            this.RecordSelectionForm.AutoSize = true;
            this.RecordSelectionForm.Location = new System.Drawing.Point(853, 200);
            this.RecordSelectionForm.Name = "RecordSelectionForm";
            this.RecordSelectionForm.Size = new System.Drawing.Size(129, 13);
            this.RecordSelectionForm.TabIndex = 42;
            this.RecordSelectionForm.Text = "Record Selection Formula";
            // 
            // btnRecordSelectionForm
            // 
            this.btnRecordSelectionForm.Location = new System.Drawing.Point(853, 216);
            this.btnRecordSelectionForm.Multiline = true;
            this.btnRecordSelectionForm.Name = "btnRecordSelectionForm";
            this.btnRecordSelectionForm.ReadOnly = true;
            this.btnRecordSelectionForm.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.btnRecordSelectionForm.Size = new System.Drawing.Size(332, 274);
            this.btnRecordSelectionForm.TabIndex = 43;
            this.btnRecordSelectionForm.MouseClick += new System.Windows.Forms.MouseEventHandler(this.crystalReportViewer1_MouseClick);
            this.btnRecordSelectionForm.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.crystalReportViewer1_DoubleClickPage);
            // 
            // btnReportName
            // 
            this.btnReportName.Location = new System.Drawing.Point(84, 34);
            this.btnReportName.Name = "btnReportName";
            this.btnReportName.Size = new System.Drawing.Size(417, 20);
            this.btnReportName.TabIndex = 44;
            // 
            // ReportName
            // 
            this.ReportName.AutoSize = true;
            this.ReportName.Location = new System.Drawing.Point(10, 38);
            this.ReportName.Name = "ReportName";
            this.ReportName.Size = new System.Drawing.Size(73, 13);
            this.ReportName.TabIndex = 45;
            this.ReportName.Text = "Report Name:";
            // 
            // rtnSQLStatement
            // 
            this.rtnSQLStatement.AutoSize = true;
            this.rtnSQLStatement.Location = new System.Drawing.Point(853, 501);
            this.rtnSQLStatement.Name = "rtnSQLStatement";
            this.rtnSQLStatement.Size = new System.Drawing.Size(82, 13);
            this.rtnSQLStatement.TabIndex = 47;
            this.rtnSQLStatement.Text = "SQL Statement:";
            // 
            // rtnReportObjects
            // 
            this.rtnReportObjects.AutoSize = true;
            this.rtnReportObjects.Location = new System.Drawing.Point(1188, 10);
            this.rtnReportObjects.Name = "rtnReportObjects";
            this.rtnReportObjects.Size = new System.Drawing.Size(78, 13);
            this.rtnReportObjects.TabIndex = 51;
            this.rtnReportObjects.Text = "Report Objects";
            // 
            // ReportObjectComboBox1
            // 
            this.ReportObjectComboBox1.AllowDrop = true;
            this.ReportObjectComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ReportObjectComboBox1.Items.AddRange(new object[] {
            "",
            "Formula Fields",
            "Groups",
            "Parameter Fields",
            "Fields used in the report",
            "Sorts",
            "Summary Fields",
            "Running Totals",
            "Main Report Data Sources",
            "Main Report Data Links",
            "SubReports",
            "SubReport Data Sources",
            "SubReport Links",
            "Special Fields",
            "Section Print Orientation",
            "Hyperlinks",
            "Summary Info",
            "Table Joins",
            "Alerts",
            "Object Name -> Field Name",
            "LOV - List of Values",
            "Graphic Location Formula",
            "Fonts used in the report",
            "Charts",
            "Condition Fields/Formula",
            "Section Condition Formula",
            "OLE Objects"});
            this.ReportObjectComboBox1.Location = new System.Drawing.Point(1191, 27);
            this.ReportObjectComboBox1.Name = "ReportObjectComboBox1";
            this.ReportObjectComboBox1.Size = new System.Drawing.Size(189, 21);
            this.ReportObjectComboBox1.TabIndex = 52;
            this.ReportObjectComboBox1.SelectedIndexChanged += new System.EventHandler(this.ReportObjectComboBox1_SelectedIndexChanged);
            // 
            // SetPCDatabase
            // 
            this.SetPCDatabase.Location = new System.Drawing.Point(9, 37);
            this.SetPCDatabase.Name = "SetPCDatabase";
            this.SetPCDatabase.Size = new System.Drawing.Size(72, 23);
            this.SetPCDatabase.TabIndex = 54;
            this.SetPCDatabase.Text = "PC Data";
            this.SetPCDatabase.UseVisualStyleBackColor = true;
            this.SetPCDatabase.Click += new System.EventHandler(this.SetPCDatabase_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(1386, 28);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 16);
            this.label8.TabIndex = 60;
            this.label8.Text = "Count:";
            // 
            // btnCount
            // 
            this.btnCount.Location = new System.Drawing.Point(1437, 28);
            this.btnCount.Name = "btnCount";
            this.btnCount.Size = new System.Drawing.Size(44, 20);
            this.btnCount.TabIndex = 61;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(247, 59);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(66, 13);
            this.label11.TabIndex = 62;
            this.label11.Text = "Report Kind:";
            // 
            // btnReportKind
            // 
            this.btnReportKind.Location = new System.Drawing.Point(316, 56);
            this.btnReportKind.Name = "btnReportKind";
            this.btnReportKind.Size = new System.Drawing.Size(195, 20);
            this.btnReportKind.TabIndex = 63;
            // 
            // btnRecordCount
            // 
            this.btnRecordCount.Location = new System.Drawing.Point(1035, 842);
            this.btnRecordCount.Name = "btnRecordCount";
            this.btnRecordCount.Size = new System.Drawing.Size(109, 20);
            this.btnRecordCount.TabIndex = 64;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(941, 845);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(88, 13);
            this.label12.TabIndex = 65;
            this.label12.Text = "Data Row count:";
            // 
            // chkUseRAS
            // 
            this.chkUseRAS.AutoSize = true;
            this.chkUseRAS.Location = new System.Drawing.Point(10, 65);
            this.chkUseRAS.Name = "chkUseRAS";
            this.chkUseRAS.Size = new System.Drawing.Size(70, 17);
            this.chkUseRAS.TabIndex = 69;
            this.chkUseRAS.Text = "Use RAS";
            this.chkUseRAS.UseVisualStyleBackColor = true;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(232, 88);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(56, 13);
            this.label15.TabIndex = 74;
            this.label15.Text = "DB Driver:";
            // 
            // btnDBDriver
            // 
            this.btnDBDriver.Location = new System.Drawing.Point(289, 85);
            this.btnDBDriver.Name = "btnDBDriver";
            this.btnDBDriver.Size = new System.Drawing.Size(222, 20);
            this.btnDBDriver.TabIndex = 75;
            // 
            // btnSQLStatement
            // 
            this.btnSQLStatement.Location = new System.Drawing.Point(856, 517);
            this.btnSQLStatement.Name = "btnSQLStatement";
            this.btnSQLStatement.Size = new System.Drawing.Size(329, 319);
            this.btnSQLStatement.TabIndex = 76;
            this.btnSQLStatement.Text = "Must be logged on before SQL can be retrieved";
            // 
            // btrDataFile
            // 
            this.btrDataFile.Location = new System.Drawing.Point(598, 9);
            this.btrDataFile.Name = "btrDataFile";
            this.btrDataFile.Size = new System.Drawing.Size(238, 20);
            this.btrDataFile.TabIndex = 77;
            this.btrDataFile.Text = "10.161.27.110";
            // 
            // btrSearchPath
            // 
            this.btrSearchPath.Location = new System.Drawing.Point(598, 31);
            this.btrSearchPath.Name = "btrSearchPath";
            this.btrSearchPath.Size = new System.Drawing.Size(238, 20);
            this.btrSearchPath.TabIndex = 78;
            this.btrSearchPath.Text = "Will be populated in code";
            // 
            // btnRasOpen
            // 
            this.btnRasOpen.Location = new System.Drawing.Point(11, 83);
            this.btnRasOpen.Name = "btnRasOpen";
            this.btnRasOpen.Size = new System.Drawing.Size(75, 23);
            this.btnRasOpen.TabIndex = 84;
            this.btnRasOpen.Text = "RAS Open";
            this.btnRasOpen.UseVisualStyleBackColor = true;
            this.btnRasOpen.Click += new System.EventHandler(this.btnRasOpen_Click);
            // 
            // btrRuntimeVersion
            // 
            this.btrRuntimeVersion.AutoSize = true;
            this.btrRuntimeVersion.Location = new System.Drawing.Point(92, 88);
            this.btrRuntimeVersion.Name = "btrRuntimeVersion";
            this.btrRuntimeVersion.Size = new System.Drawing.Size(46, 13);
            this.btrRuntimeVersion.TabIndex = 85;
            this.btrRuntimeVersion.Text = "Runtime";
            // 
            // txtRuntimeVersion
            // 
            this.txtRuntimeVersion.Location = new System.Drawing.Point(144, 85);
            this.txtRuntimeVersion.Name = "txtRuntimeVersion";
            this.txtRuntimeVersion.Size = new System.Drawing.Size(84, 20);
            this.txtRuntimeVersion.TabIndex = 86;
            // 
            // btrVerifyDatabase
            // 
            this.btrVerifyDatabase.AutoSize = true;
            this.btrVerifyDatabase.Location = new System.Drawing.Point(733, 82);
            this.btrVerifyDatabase.Name = "btrVerifyDatabase";
            this.btrVerifyDatabase.Size = new System.Drawing.Size(101, 17);
            this.btrVerifyDatabase.TabIndex = 81;
            this.btrVerifyDatabase.Text = "Verify Database";
            this.btrVerifyDatabase.UseVisualStyleBackColor = true;
            // 
            // btrFileLocation
            // 
            this.btrFileLocation.Location = new System.Drawing.Point(598, 58);
            this.btrFileLocation.Name = "btrFileLocation";
            this.btrFileLocation.Size = new System.Drawing.Size(129, 20);
            this.btrFileLocation.TabIndex = 79;
            this.btrFileLocation.Text = "sa";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkUseRAS);
            this.groupBox1.Controls.Add(this.SetPCDatabase);
            this.groupBox1.Controls.Add(this.SetToXML);
            this.groupBox1.Location = new System.Drawing.Point(984, 7);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(94, 90);
            this.groupBox1.TabIndex = 90;
            this.groupBox1.TabStop = false;
            // 
            // SetToXML
            // 
            this.SetToXML.Location = new System.Drawing.Point(7, 10);
            this.SetToXML.Name = "SetToXML";
            this.SetToXML.Size = new System.Drawing.Size(75, 23);
            this.SetToXML.TabIndex = 30;
            this.SetToXML.Text = "Set to XML";
            this.SetToXML.UseVisualStyleBackColor = true;
            this.SetToXML.Click += new System.EventHandler(this.SetToXML_Click);
            // 
            // btrPassword
            // 
            this.btrPassword.Location = new System.Drawing.Point(598, 81);
            this.btrPassword.Name = "btrPassword";
            this.btrPassword.Size = new System.Drawing.Size(129, 20);
            this.btrPassword.TabIndex = 80;
            this.btrPassword.Text = "1Oem2000";
            // 
            // lblRPTRev
            // 
            this.lblRPTRev.Location = new System.Drawing.Point(194, 56);
            this.lblRPTRev.Name = "lblRPTRev";
            this.lblRPTRev.Size = new System.Drawing.Size(52, 20);
            this.lblRPTRev.TabIndex = 92;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(144, 59);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(51, 13);
            this.label18.TabIndex = 93;
            this.label18.Text = "Revision:";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(528, 13);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(69, 13);
            this.label19.TabIndex = 94;
            this.label19.Text = "DSN/Server:";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(541, 35);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(56, 13);
            this.label20.TabIndex = 95;
            this.label20.Text = "Database:";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(565, 62);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(32, 13);
            this.label21.TabIndex = 96;
            this.label21.Text = "User:";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(541, 85);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(56, 13);
            this.label22.TabIndex = 97;
            this.label22.Text = "Password:";
            // 
            // label
            // 
            this.label.AutoSize = true;
            this.label.Location = new System.Drawing.Point(8, 114);
            this.label.Name = "label";
            this.label.Size = new System.Drawing.Size(96, 13);
            this.label.TabIndex = 98;
            this.label.Text = "CRD Save History:";
            // 
            // cbLastSaveHistory
            // 
            this.cbLastSaveHistory.FormattingEnabled = true;
            this.cbLastSaveHistory.Location = new System.Drawing.Point(109, 111);
            this.cbLastSaveHistory.MaxDropDownItems = 20;
            this.cbLastSaveHistory.Name = "cbLastSaveHistory";
            this.cbLastSaveHistory.Size = new System.Drawing.Size(392, 21);
            this.cbLastSaveHistory.TabIndex = 99;
            // 
            // btrRefreshViewer
            // 
            this.btrRefreshViewer.Location = new System.Drawing.Point(251, 7);
            this.btrRefreshViewer.Name = "btrRefreshViewer";
            this.btrRefreshViewer.Size = new System.Drawing.Size(79, 24);
            this.btrRefreshViewer.TabIndex = 101;
            this.btrRefreshViewer.Text = "Refresh";
            this.btrRefreshViewer.UseVisualStyleBackColor = true;
            this.btrRefreshViewer.Click += new System.EventHandler(this.btrRefreshViewer_Click);
            // 
            // btnErrorCount
            // 
            this.btnErrorCount.Location = new System.Drawing.Point(1494, 28);
            this.btnErrorCount.Name = "btnErrorCount";
            this.btnErrorCount.Size = new System.Drawing.Size(44, 20);
            this.btnErrorCount.TabIndex = 102;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(1494, 11);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(37, 13);
            this.label23.TabIndex = 103;
            this.label23.Text = "Errors:";
            // 
            // chkTrusted
            // 
            this.chkTrusted.AutoSize = true;
            this.chkTrusted.Location = new System.Drawing.Point(733, 61);
            this.chkTrusted.Name = "chkTrusted";
            this.chkTrusted.Size = new System.Drawing.Size(62, 17);
            this.chkTrusted.TabIndex = 104;
            this.chkTrusted.Text = "Trusted";
            this.chkTrusted.UseVisualStyleBackColor = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsslThread,
            this.tsslHandles,
            this.tsslPrivateBytes});
            this.statusStrip1.Location = new System.Drawing.Point(0, 891);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1570, 22);
            this.statusStrip1.TabIndex = 110;
            this.statusStrip1.Text = "statusStrip1";
            this.statusStrip1.Click += new System.EventHandler(this.timeCounter);
            // 
            // tsslThread
            // 
            this.tsslThread.Name = "tsslThread";
            this.tsslThread.Size = new System.Drawing.Size(56, 17);
            this.tsslThread.Text = "Handles: ";
            // 
            // tsslHandles
            // 
            this.tsslHandles.Name = "tsslHandles";
            this.tsslHandles.Size = new System.Drawing.Size(55, 17);
            this.tsslHandles.Text = "Threads: ";
            // 
            // tsslPrivateBytes
            // 
            this.tsslPrivateBytes.Name = "tsslPrivateBytes";
            this.tsslPrivateBytes.Size = new System.Drawing.Size(80, 17);
            this.tsslPrivateBytes.Text = "Private Bytes: ";
            // 
            // btnHasSavedData
            // 
            this.btnHasSavedData.AutoSize = true;
            this.btnHasSavedData.Location = new System.Drawing.Point(1054, 103);
            this.btnHasSavedData.Name = "btnHasSavedData";
            this.btnHasSavedData.Size = new System.Drawing.Size(105, 17);
            this.btnHasSavedData.TabIndex = 112;
            this.btnHasSavedData.Text = "Has Saved Data";
            this.btnHasSavedData.UseVisualStyleBackColor = true;
            // 
            // btnReportObjects
            // 
            this.btnReportObjects.Location = new System.Drawing.Point(1191, 62);
            this.btnReportObjects.Name = "btnReportObjects";
            this.btnReportObjects.Size = new System.Drawing.Size(367, 805);
            this.btnReportObjects.TabIndex = 111;
            this.btnReportObjects.Text = "";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(853, 150);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(137, 13);
            this.label31.TabIndex = 123;
            this.label31.Text = "GDIPlus InterpolationMode:";
            // 
            // LstInterpolationMode
            // 
            this.LstInterpolationMode.FormattingEnabled = true;
            this.LstInterpolationMode.Location = new System.Drawing.Point(853, 166);
            this.LstInterpolationMode.Name = "LstInterpolationMode";
            this.LstInterpolationMode.Size = new System.Drawing.Size(329, 21);
            this.LstInterpolationMode.TabIndex = 124;
            // 
            // CommandTable
            // 
            this.CommandTable.Location = new System.Drawing.Point(1085, 35);
            this.CommandTable.Name = "CommandTable";
            this.CommandTable.Size = new System.Drawing.Size(97, 23);
            this.CommandTable.TabIndex = 125;
            this.CommandTable.Text = "Set Command";
            this.CommandTable.UseVisualStyleBackColor = true;
            this.CommandTable.Click += new System.EventHandler(this.CommandTable_Click_1);
            // 
            // frmMain
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(1570, 913);
            this.Controls.Add(this.CommandTable);
            this.Controls.Add(this.LstInterpolationMode);
            this.Controls.Add(this.label31);
            this.Controls.Add(this.btnHasSavedData);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.chkTrusted);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.btnErrorCount);
            this.Controls.Add(this.btrRefreshViewer);
            this.Controls.Add(this.cbLastSaveHistory);
            this.Controls.Add(this.label);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.lblRPTRev);
            this.Controls.Add(this.btrPassword);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btrFileLocation);
            this.Controls.Add(this.btrVerifyDatabase);
            this.Controls.Add(this.txtRuntimeVersion);
            this.Controls.Add(this.btrRuntimeVersion);
            this.Controls.Add(this.btnRasOpen);
            this.Controls.Add(this.btrSearchPath);
            this.Controls.Add(this.btrDataFile);
            this.Controls.Add(this.btnSQLStatement);
            this.Controls.Add(this.btnDBDriver);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.btnRecordCount);
            this.Controls.Add(this.btnReportKind);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.btnCount);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.ReportObjectComboBox1);
            this.Controls.Add(this.rtnReportObjects);
            this.Controls.Add(this.btnReportObjects);
            this.Controls.Add(this.rtnSQLStatement);
            this.Controls.Add(this.ReportName);
            this.Controls.Add(this.btnReportName);
            this.Controls.Add(this.btnRecordSelectionForm);
            this.Controls.Add(this.RecordSelectionForm);
            this.Controls.Add(this.btnReportVersion);
            this.Controls.Add(this.ReportVersion);
            this.Controls.Add(this.btnCloserpt);
            this.Controls.Add(this.crystalReportViewer1);
            this.Controls.Add(this.ReplaceConnection);
            this.Controls.Add(this.SetLogonInfo);
            this.Controls.Add(this.Export);
            this.Controls.Add(this.SetParam);
            this.Controls.Add(this.SetData);
            this.Controls.Add(this.ViewReport);
            this.Controls.Add(this.btnSaveRptAs);
            this.Controls.Add(this.btnOpenReport);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Unmanaged RAS 2010 Parameters";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
        
        static void Main() 
		{
			Application.Run(new frmMain());
		}

		private void frmMain_Load(object sender, System.EventArgs e)
		{

		}

        bool IsRpt = true;
        bool IsLoggedOn; // = false;
        bool IsCRSE = false;
        bool IsCMD = false;
        bool IsBEX = false;

		private void grpBoxCurrent_Enter(object sender, System.EventArgs e)
		{
		
		}

        private void btnRasOpen_Click(object sender, EventArgs e)
        {
            // NOTE: Cortez does not support unmanaged reporting so have to look into using version 14 Enterprise dll's to connect to CRSE 2011.
            CrystalDecisions.CrystalReports.Engine.ReportDocument rpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            rptClientDoc = new CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument(); // ReportClientDocumentClass();
            DateTime dtStart;
            TimeSpan difference;

            // this line should return the number of concurent reports being executed
            int i = CrystalDecisions.CrystalReports.Engine.ReportDocument.GetConcurrentUsage();

            // do not modify this works
            // this line uses the local PC as the RAS server/services running
            //rptClientDoc.ReportAppServer = "VMDWCR2k8RAS"; //System.Environment.MachineName;
            // rassdk means load from local HD
            // ras requires access to the Ras Server folder
            String RPTPath = btnReportName.Text; // "rassdk://c:\\RASReports\\formulas.rpt"; c:\\RASReports\\formulas.rpt
            //Set the ReportAppServer and open the document
            rptClientDoc.ReportAppServer = "10.160.0.00"; 

            object RPTObject = (object)RPTPath;
            try
            {
                dtStart = DateTime.Now;

                rptClientDoc.Open(ref RPTObject, 0);
                IsRpt = false;
                difference = DateTime.Now.Subtract(dtStart);
                btnReportObjects.Text += "Report Document Load: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + "\r\n";
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
                return;
            }

            // do not modify this works
            btnOpenReport.Enabled = false;
            btnSaveRptAs.Enabled = true;
            ViewReport.Enabled = true;
            btnCloserpt.Enabled = true;
            ViewReport.Enabled = true;

            // determine the report locale:
            CrystalDecisions.ReportAppServer.DataDefModel.CeLocale preferredViewingLocaleID;
            preferredViewingLocaleID = rptClientDoc.LocaleID;
            CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag PromptProperties = new PropertyBag();

            int dbConCount = rptClientDoc.DatabaseController.GetConnectionInfos(PromptProperties).Count;
            String DBDriver = "";

            for (int x = 0; x < dbConCount; x++)
            {
                try
                {
                    DBDriver = rptClientDoc.DatabaseController.GetConnectionInfos(PromptProperties)[x].Attributes.get_StringValue("Database DLL").ToString();
                    btnDBDriver.Text += DBDriver + " :";
                }
                catch (Exception ex)
                {
                    btnDBDriver.Text = "ERROR: " + ex.Message;
                    return;
                }
            }

            if (dbConCount == 0)
                btnDBDriver.Text = "NO Datasource in report";


            //// SP 4
            //System.Diagnostics.FileVersionInfo fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(MyVerison.Location);
            //txtRuntimeVersion.Text += fileVersionInfo.FileVersion.ToString();
            if (CRVer == "12")
            {
                btnReportKind.Text = ".ReportKind Not supported in CR 2008";
            }

            // Record selection formula with comments included can only be retrieve via RAS
            CrystalDecisions.ReportAppServer.DataDefModel.ISCRFilter myRecordSelectionWithComments; // = new CrystalDecisions.ReportAppServer.DataDefModel.;
            myRecordSelectionWithComments = rptClientDoc.DataDefController.DataDefinition.RecordFilter;
            if (myRecordSelectionWithComments.FreeEditingText != null)
            {
                myRecordSelectionWithComments.FreeEditingText = rptClientDoc.DataDefController.RecordFilterController.GetFormulaText();
                btnRecordSelectionForm.Text = myRecordSelectionWithComments.FreeEditingText.ToString();
            }
            else
                btnRecordSelectionForm.Text = "No Record Selection formula";

            btnReportVersion.Text = rptClientDoc.MajorVersion.ToString() + "." + rptClientDoc.MinorVersion.ToString();
        }

		private void btnOpenReport_Click(object sender, System.EventArgs e) 
		{

            //bool msOutlook = false;
            //foreach (Process otlk in Process.GetProcesses())
            //{
            //    if (otlk.ProcessName.Contains("OUTLOOK"))
            //    {
            //        msOutlook = true;
            //    }
            //}
            //if (msOutlook.Equals(true))
            //{
            //    btnSQLStatement.AppendText("\nOutlook is running");
            //}
            //else
            //{
            //    //MessageBox.Show("Ms Outlook needs to be running to send email", "Mail error", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
            //    DialogResult dr;
            //    dr = MessageBox.Show("Ms Outlook needs to be running to Export to MAPI e-mail, \nclick Cancel or No to not start Outlook", "Mail error", MessageBoxButtons.YesNoCancel);
            //    if (dr == DialogResult.Yes)
            //        System.Diagnostics.Process.Start("OUTLOOK.EXE");
            //}

            //CrystalDecisions.CrystalReports.Engine.ReportDocument rpt = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            rptClientDoc = new CrystalDecisions.ReportAppServer.ClientDoc.ReportClientDocument(); // ReportClientDocumentClass();
            DateTime dtStart;
            TimeSpan difference;

            openFileDialog.Filter = "Crystal Reports (*.rpt)|*.rpt|Crystal Reports Secure (*.rptr)|*.rptr";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;

			if (openFileDialog.ShowDialog() == DialogResult.OK) 
			{
				btnOpenReport.Enabled = false;
				btnSaveRptAs.Enabled = false;
                btnCloserpt.Enabled = false;
				object rptName = openFileDialog.FileName;

                dtStart = DateTime.Now;

                try
                {
                    rpt.Load(rptName.ToString(), OpenReportMethod.OpenReportByTempCopy);
                    difference = DateTime.Now.Subtract(dtStart);
                    btnReportObjects.Text += "Report Document Load: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + "\r\n";
 
                    cbLastSaveHistory.Text = "";
                    try
                    {
                        for (int x = 0; x < rpt.HistoryInfos.Count; x++)
                        {
                            cbLastSaveHistory.Items.Add(rpt.HistoryInfos[x].BuildVersion.ToString() + ": Date: " + rpt.HistoryInfos[x].SavedDate.ToString());
                        }

                        cbLastSaveHistory.SelectedIndex = 0; // rpt.HistoryInfos.Count - 1;
                        //SP 13
                        //•	RAS .NET SDK 
                        //ReportClientDocument.HistoryInfos[i].SavedDate
                        //ReportClientDocument.HistoryInfos[i].BuildVersion

                        //•	CR .NET SDK 
                        //ReportDocument.HistoryInfos[i].SavedDate
                        //ReportDocument.HistoryInfos[i].BuildVersion
                    }
                    catch (Exception ex)
                    {
                        cbLastSaveHistory.Text = "This report has no Save history";
                    }

                    btnOpenReport.Enabled = false;
                    btnSaveRptAs.Enabled = true;
                    ViewReport.Enabled = true;
                    btnCloserpt.Enabled = true;
                    ViewReport.Enabled = true;
                    btnReportName.Text = rptName.ToString(); 
                    btnReportKind.Text = rpt.ReportDefinition.ReportKind.ToString();
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToString() == "Object reference not set to an instance of an object.")
                        MessageBox.Show("ERROR: Object reference not set to an instance of an object.");
                    else
                        if (ex.Message.ToString() == "External component has thrown an exception.")
                        {
                            MessageBox.Show("ERROR: External component has thrown an exception.");
                        }
                        else
                        {
                            {
                                if (ex.InnerException.Message != null)
                                {
                                    MessageBox.Show("ERROR: " + ex.Message + " ;" + ex.InnerException.Message);

                                    // to handle your own error message search or help files
                                    //string myURL = @"http://search.sap.com/ui/scn#query=" + ex.InnerException.Message + "&startindex=1&filter=scm_a_site%28scm_v_Site11%29&filter=scm_a_modDate%28*%29&timeScope=all";
                                    //string fixedString = myURL.Replace(" ", "%20");

                                    ////System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                                    //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);

                                    ////string myURL = @"C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports 2011\Help\en\crw.chm";
                                    ////System.Diagnostics.Process.Start(myURL);
                                    ////cool but of no use to release mode

                                }
                                else
                                {
                                    MessageBox.Show("ERROR: " + ex.Message);
                                }
                            }
                        }

                    btnOpenReport.Enabled = true;
                    ViewReport.Enabled = false;
                    btnSaveRptAs.Enabled = false;
                    ViewReport.Enabled = false; 
                    btnCloserpt.Enabled = false;
                    btnReportKind.Text = "";
                    return;
                }
                // then clone the report
                rpt1 = (CrystalDecisions.CrystalReports.Engine.ReportDocument)rpt.Clone();
                // this assigns the report to RAS for modification
                rptClientDoc = rpt.ReportClientDocument;
                rptClientDoc1 = rpt.ReportClientDocument;

                // determine the report locale:
                CrystalDecisions.ReportAppServer.DataDefModel.CeLocale preferredViewingLocaleID;
                preferredViewingLocaleID = rptClientDoc.LocaleID;

                // check if the report is based on a Command and if so then display the SQL. This causes a huge delay opening report
                btnReportObjects.AppendText("");
                dtStart = DateTime.Now;
                int dbConCount = rptClientDoc.DatabaseController.GetConnectionInfos().Count;
                difference = DateTime.Now.Subtract(dtStart);
                btnReportObjects.Text += "GetConnectionInfos().Count took: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + "\r\n";

                String DBDriver = "";
                for (int x = 0; x < dbConCount; x++)
                {
                    try
                    {
                        DBDriver = rptClientDoc.DatabaseController.GetConnectionInfos()[x].Attributes.get_StringValue("Database DLL").ToString();
                        btnDBDriver.Text += DBDriver + " :";
                        if (((dynamic)rptClientDoc.Database.Tables[0].Name) == "Command")
                        {
                            CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = rpt.ReportClientDocument.DatabaseController;
                            ISCRTable oldTable = (ISCRTable)databaseController.Database.Tables[0];

                            btnSQLStatement.Text = "Report is using Command Object: \n" + ((dynamic)oldTable).CommandText.ToString();
                            btnSQLStatement.Text += "\n";

                            IsLoggedOn = false;
                            IsCMD = true;
                        }
                        if (DBDriver.ToString() == "crdb_bwmdx.dll")
                            IsBEX = true;

                    }
                    catch (Exception ex)
                    {
                        //btnDBDriver.Text = "ERROR: " + ex.Message;
                        btnDBDriver.Text += "Main Report has no Data Driver";
                    }
                }

                if (dbConCount == 0)
                    btnDBDriver.Text = "NO Datasource in report";

                // this line should return the number of concurrent reports being executed
                int i = CrystalDecisions.CrystalReports.Engine.ReportDocument.GetConcurrentUsage();
                btnSQLStatement.Text += "\nGetConcurrentUsage: " + i + "\n";

                // SP 4
                btnReportKind.Text = "Eng: " + rpt.ReportDefinition.ReportKind.ToString() + " : RAS - " + rptClientDoc.ReportDefController.ReportDefinition.ReportKind.ToString();

                // SP 2 addition
                //crystalReportViewer1.SetFocusOn(UIComponent.Page);

                //CrystalDecisions.Shared.PrintLayoutSettings.PrintScaling.Scale = PrintLayoutSettings.PrintScaling.DoNotScale;

                //rpt.FormatEngine.GetLastPageNumber(new ReportPageRequestContext());
                //rpt.FormatSection += new FormatSectionEventHandler(report_FormatSection);
                //EventEnabledArgs enableArgs = new EventEnabledArgs();
                //enableArgs.FormattingEnabled = true;
                //rpt.EnableEvent(enableArgs);

                //CrystalDecisions.CrystalReports.Engine.FormatSectionEventArgs myTrigger = new FormatSectionEventArgs();


                // this one uses the RAS Service
                //rptClientDoc.Open(ref rptName, (int) CdReportClientDocumentOpenOptionsEnum.cdOpenReportRetrieveMinimumReportDocument);

                //MessageBox.Show("Report version: " + rptClientDoc.MajorVersion.ToString() + "." + rptClientDoc.MinorVersion.ToString(), "RAS" ,MessageBoxButtons.OK,MessageBoxIcon.Information );

                //MessageBox.Show("Report opened.","RAS",MessageBoxButtons.OK,MessageBoxIcon.Information );

                //btnSQLString.text = "";

                // Record selection formula with comments included can only be retrieve via RAS
                CrystalDecisions.ReportAppServer.DataDefModel.ISCRFilter myRecordSelectionWithComments; // = new CrystalDecisions.ReportAppServer.DataDefModel.;
                myRecordSelectionWithComments = rptClientDoc.DataDefController.DataDefinition.RecordFilter;
                if (myRecordSelectionWithComments.FreeEditingText != null)
                {
                    btnRecordSelectionForm.Text = "\nWith Comments:\n" + myRecordSelectionWithComments.FreeEditingText.ToString();
                    btnRecordSelectionForm.AppendText("\n\n");
                    btnRecordSelectionForm.AppendText("\nWithout Comments:\n" + rpt.RecordSelectionFormula.ToString());
                    btnRecordSelectionForm.AppendText("\n");
                }
                else
                    btnRecordSelectionForm.Text = "No Record Selection formula";

                //rpt.RecordSelectionFormula = @"{Orders.Customer ID} IN [2]";
                //btnRecordSelectionForm.Text = rpt.RecordSelectionFormula.ToString();

                btnReportVersion.Text = rptClientDoc.MajorVersion.ToString() + "." + rptClientDoc.MinorVersion.ToString();
                // SP 13
                lblRPTRev.Text = rpt.SummaryInfo.RevisionNumber.ToString();

                getReportOptionsOnOpen(rptClientDoc);

                // get the DB name from the report
                CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                CrystalDecisions.CrystalReports.Engine.Tables crTables;
                crDatabase = rpt.Database;
                crTables = crDatabase.Tables;
                int dbx = 0;

                foreach (CrystalDecisions.CrystalReports.Engine.Table crTable in crTables)
                {
                    if (crDatabase.Tables.Count != 0)
                    {
                        CrystalDecisions.Shared.NameValuePair2 nvp2 = (NameValuePair2)rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.Attributes.Collection[1];
                        btnSQLStatement.Text += "\nRPT Data Source Info: \n" + " Server Name: " + rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.ServerName.ToString() + "\n Database Name: " + nvp2.Value.ToString() + "\n Table Name: " + rpt.Database.Tables[dbx].Name.ToString();
                        if (rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.UserID != null)
                            btnSQLStatement.Text += "\n User ID: " + rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.UserID.ToString() + "\r\n";
                        btrSearchPath.Text = rpt.Database.Tables[dbx].Name.ToString();
                        dbx++;
                    }
                    else
                        btnSQLStatement.Text += "No Data source or not found\r\n";
                }

                //get the subreport connection infos
                //btnReportObjects.Text = "";
                string SecName = "";

                CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;

                //set the crSections object to the current report's sections
                CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;
                int flcnt = 0;
                bool SecureDB;

                //loop through all the sections to find all the report objects
                foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                {
                    crReportObjects = crSection.ReportObjects;
                    //loop through all the report objects to find all the subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                    {
                        if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                        {
                            try
                            {
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                dbx = 0;

                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                                SubreportClientDocument subRCD = rptClientDoc.SubreportController.GetSubreport(crSubreportObject.SubreportName);
                                string mysubname = crSubreportObject.SubreportName.ToString();

                                try
                                {
                                    //open the subreport object
                                    //subRCD = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                    CrystalDecisions.Shared.ConnectionInfo crSubConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();
                                    btnSQLStatement.Text += "\n\nSubReport Table count: " + subRCD.DatabaseController.Database.Tables.Count.ToString();

                                    // get the DB names from the subreport
                                    //crDatabase = subRCD.DatabaseController.Database;
                                    //crTables = crDatabase.Tables;

                                    if (subRCD.DatabaseController.Database.Tables.Count != 0)
                                    {
                                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table crTable in subRCD.DatabaseController.Database.Tables)
                                        {
                                            // using engine if subreprot is not using a Command
                                            if (subRCD.DatabaseController.Database.Tables.Count != 0 && (((dynamic)crTable.Name) != "Command"))
                                            {
                                                CrystalDecisions.Shared.NameValuePair2 nvp2Sub = (NameValuePair2)crSubreportDocument.Database.Tables[dbx].LogOnInfo.ConnectionInfo.Attributes.Collection[1];
                                                btnSQLStatement.Text += "\nSubreport Data Source Info: \n" + " Server Name: " + crSubreportDocument.Database.Tables[dbx].LogOnInfo.ConnectionInfo.ServerName.ToString() + "\n Database Name: " + nvp2Sub.Value.ToString() + "\n Table Name: " + crSubreportDocument.Database.Tables[dbx].Name.ToString();

                                                try
                                                {
                                                    if (rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.UserID != null)
                                                        btnSQLStatement.Text += "\n User ID: " + rpt.Database.Tables[dbx].LogOnInfo.ConnectionInfo.UserID.ToString() + "\r\n";
                                                    dbx++;
                                                }
                                                catch (Exception ex)
                                                {
                                                    btnReportObjects.AppendText("\n User ID Does not exist for Subreport Table\n");
                                                }                                            //return;
                                            }
                                            else
                                                if (((dynamic)crTable.Name) != "Command")
                                                btnSQLStatement.Text += "No Data source or not found\r\n";

                                            try
                                            {
                                                // Subreport is using a Command so use RAS to get the SQL
                                                btnDBDriver.Text += DBDriver + " :";
                                                if (((dynamic)crTable.Name) == "Command")
                                                {

                                                    CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = subRCD.DatabaseController;
                                                    CommandTable SuboldTable = (CommandTable)databaseController.Database.Tables[0];

                                                    btnSQLStatement.Text += "\n\n";
                                                    btnSQLStatement.Text += "SubReport is using Command Object: \n" + ((dynamic)SuboldTable).CommandText.ToString();
                                                    btnSQLStatement.Text += "\n";

                                                    IsLoggedOn = false;
                                                    IsCMD = true;
                                                }
                                                if (DBDriver.ToString() == "crdb_bwmdx.dll")
                                                    IsBEX = true;

                                            }
                                            catch (Exception ex)
                                            {
                                                //btnDBDriver.Text += "ERROR: " + ex.Message;
                                                btnDBDriver.Text += "Main Report has no Data Driver";
                                            }
                                        }
                                    }
                                    else
                                    {
                                        btnReportObjects.Text += "\nSubreport: " + subRCD.Name.ToString() + ": Has no Data Source\n";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    btnReportObjects.Text += "SubreportName: " + subRCD.Name + "\n";
                                    btnReportObjects.Update();
                                    btnReportObjects.AppendText("Error in " + SecName + " : " + ex.Message + "\n\n");
                                }
                                break;

                            }
                            catch (Exception ex)
                            {
                                btnReportObjects.AppendText("Error in " + SecName + " : " + ex.Message + "\n");
                            }
                        }
                    }
                }

                if (IsLoggedOn)
                {
                    dtStart = DateTime.Now;
                    GroupPath gp = new GroupPath();
                    string tmp = String.Empty;
                    try
                    {
                        rptClientDoc.RowsetController.GetSQLStatement(gp, out tmp);
                        btnSQLStatement.Text = tmp;
                    }
                    catch (Exception ex)
                    {
                        btnSQLStatement.Text = "ERROR: " + ex.Message;
                        return;
                    }
                    difference = DateTime.Now.Subtract(dtStart);
                    btnReportObjects.Text += "\nGet SQL Statement: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + "\r\n";
                }
            }
		}

        private void getReportOptionsOnOpen(CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc)
        {
            CrystalDecisions.ReportAppServer.ReportDefModel.ReportOptions iro = rptClientDoc.ReportOptions;
            CrystalDecisions.ReportAppServer.ReportDefModel.ReportOptions myRPTOpts = iro.Clone();

            //CrystalDecisions.ReportAppServer.Controllers.RowsetController crRowsetController; // = new RowsetController();
            //CrystalDecisions.ReportAppServer.Controllers.RowsetCursor crRowsetCursor;

            //crRowsetController = rptClientDoc.RowsetController;
            //crRowsetCursor = crRowsetController.CreateCursor(null);
            //btnRecordCount.Text = crRowsetCursor.RecordCount.ToString();

            if (rpt.HasSavedData == true)
            {
                try
                {
                    bool doesIt = rpt.HasRecords;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(rpt.HasRecords.ToString());
                }

                btnHasSavedData.Checked = true;
                RowsetController boRowsetController;
                RowsetCursor boRowsetCursor;

                GroupPath gp = new GroupPath();
                gp.FromString("*"); // this changes the results but it is not the complete count and it does not get the subreport SQL
                int rows;
                boRowsetController = rptClientDoc.RowsetController;
                try
                {
                    if (!IsLoggedOn)
                    {
                        boRowsetCursor = boRowsetController.CreateCursor(gp, new RowsetMetaData(), 0);

                        (boRowsetCursor).RowsetController.RowsetBatchSize = 1000;
                        (boRowsetCursor).Rowset.BatchSize = 1000;
                        (boRowsetCursor).MoveLast();
                        rows = boRowsetCursor.RecordCount;
                        //txtRecordCount.Text = rows.ToString();
                    }
                    else
                    {
                        btnReportObjects.Text += "\nSee SQL Statement box for SQL info for main and/or subreport";
                    }
                }
                catch (Exception ex)
                {
                    if (ex.Message.ToString() == "Failed to get data source.")
                    {
                        btnReportObjects.Text += "\nMust be logged onto Database before SQL for subreport can be retrieved or...\nReports SQL Query may have been manually modified, this is no longer supported - please check the report with CR Designer\n";
                    }
                    else
                    {
                        if (ex.Message == "\rNo error.")
                            btnReportObjects.Text += ex.Message + "\n";
                        else
                        {
                            if (ex.Message == "The Report Application Server failed")
                            {
                                btnReportObjects.Text += "\nThe Report Application Server failed";
                            }
                            else
                            {
                                if (ex.Message.Substring(0, 56) == "Invalid group number.\rError in creating new data source.")
                                {
                                    btnReportObjects.Text += "\nMust be logged onto Database before SQL for subreport can be retrieved or...\nReports SQL Query may have been manually modified, this is no longer supported - please check the report with CR Designer\n";
                                }
                                else
                                    btnReportObjects.Text += "Report has Dynamic Cascading Parameter - SDK cannot get SQL for these\n"; // +ex.Message + "\n";
                            }
                        }
                    }

                    gp.FromString("");
                    try
                    {
                        boRowsetCursor = boRowsetController.CreateCursor(gp, new RowsetMetaData(), 0);

                        (boRowsetCursor).RowsetController.RowsetBatchSize = 1000;
                        (boRowsetCursor).Rowset.BatchSize = 1000;
                        (boRowsetCursor).MoveLast();
                        rows = boRowsetCursor.RecordCount;
                        //txtRecordCount.Text = rows.ToString();
                    }
                    catch (Exception exx)
                    {
                        {
                            btnReportObjects.Text += exx.Message + "\n";
                        } // main report does not have a name so ignore the exception
                    }

                }
            }
            else
                btnHasSavedData.Checked = false;

            // used to se the viewer to the RAS Report object
            IsRpt = false;
        }

        private void timeCounter(object sender, EventArgs e)
        {
            DateTime dtStart;
            dtStart = DateTime.Now;
            btnSQLStatement.Text += "\nClicked View Report End time: \n" + dtStart.ToString();

            string[] dateValues = { "30/12/2011", "12/30/2011", "30/12/11", "12/30/11" };
            string pattern = "MM/dd/yyyy";
            DateTime parsedDate;

            foreach (var dateValue in dateValues)
            {
                if (DateTime.TryParseExact(dateValue, pattern, null, DateTimeStyles.None, out parsedDate))
                    MessageBox.Show("Converted: " + dateValue + " to: " + parsedDate);
                else
                    MessageBox.Show("Unable to convert: " + dateValue + " to a date and time.");
            }
            // The example displays the following output: 
            //    Unable to convert '30-12-2011' to a date and time. 
            //    Unable to convert '12-30-2011' to a date and time. 
            //    Unable to convert '30-12-11' to a date and time. 
            //    Converted '12-30-11' to 12/30/2011.
        }

        private void Closerpt_Click(object sender, EventArgs e)
        {
            btnOpenReport.Enabled = true;
            btnSaveRptAs.Enabled = false;
            btnCloserpt.Enabled = false;
            ViewReport.Enabled = false;
            btnRecordSelectionForm.Text = "";
            btnReportVersion.Text = "";
            btnReportName.Text = "";
            btnSQLStatement.Text = "Log on must be set first before SQL can be retrieved";
            btnReportObjects.Text = "";
            ReportObjectComboBox1.ResetText();
            btnReportKind.Text = "";
            btnRecordCount.Text = "";
            btnDBDriver.Text = "";
            btnDBDriver.Text = "";
            crystalReportViewer1.SetProductLocale(1033);
            lblRPTRev.Text = "";
            cbLastSaveHistory.Items.Clear();
            cbLastSaveHistory.Text = "";

            btrVerifyDatabase.Checked = false;
            IsLoggedOn = false;
            bool IsCMD = false;
            bool IsBEX = false;

            chkUseRAS.CheckState = CheckState.Unchecked;

            rpt.Close();

            crystalReportViewer1.ReportSource = null;
            if (!crystalReportViewer1.Disposing)
                btnSQLStatement.Text += "\nViewer is disposing";
            else
                btnSQLStatement.Text += "\nViewer is disposed";
            crystalReportViewer1.Refresh();

            GC.Collect();
            IsRpt = true;

        }

		private void btnSaveReportAs_Click(object sender, System.EventArgs e)
		{
            saveFileDialog.Filter = "Crystal Reports (*.rpt)|*.rpt";
            if (DialogResult.OK == saveFileDialog.ShowDialog())
            {

                object saveFolder = System.IO.Path.GetDirectoryName(saveFileDialog.FileName);
                string saveFileName = System.IO.Path.GetFileName(saveFileDialog.FileName);

                if (!IsRpt)
                {
                    try
                    {
                    rptClientDoc.SaveAs(saveFileName, ref saveFolder,
                        (int)CdReportClientDocumentSaveAsOptionsEnum.cdReportClientDocumentSaveAsOverwriteExisting);
                    }
                    catch (Exception ex)
                    {
                        btnSQLStatement.Text = "ERROR: " + ex.Message;
                        return;
                    }
                }
                else
                {
                    try
                    {
                        rpt.SaveAs(saveFolder + "\\" + saveFileName, true);
                    }
                    catch (Exception ex)
                    {
                        btnSQLStatement.Text = "ERROR: " + ex.Message;
                    }
                }
            }
		}

        private void ViewReport_Click(object sender, EventArgs e)
        {
            //// from Soda to handle Viewer OLE Object refresh issue 175343
            //crystalReportViewer1.CachedPageNumberPerDoc = 1;

            DateTime dtStart;
            dtStart = DateTime.Now;
            btnSQLStatement.Text += "\nClicked View Report Start time: \n" + dtStart.ToString();

            //// SP 17 - SAP NOTE 2321691 
            //System.Drawing.Drawing2D.InterpolationMode CRInterpolMode = new System.Drawing.Drawing2D.InterpolationMode();
            //CRInterpolMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            crystalReportViewer1.InterpolationMode = (System.Drawing.Drawing2D.InterpolationMode.NearestNeighbor);

            //crystalReportViewer1.ReuseParameterValuesOnRefresh = false;
            //crystalReportViewer1.ViewTimeSelectionFormula = "hello";

            // next line causes a delay due to the engine having to go to the last page.
            try
            {
                if (rpt.Rows.Count > 0)
                {
                    btnRecordCount.Text = rpt.Rows.Count.ToString();

                    try
                    {
                        btnSQLStatement.Text += rpt.FormatEngine.GetLastPageNumber(new ReportPageRequestContext()).ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("no pages");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("no data Mr. Dumas, must log on");
            }

            try
            {
                if (!IsRpt)
                {
                    crystalReportViewer1.ReportSource = rptClientDoc.ReportSource;
                }
                else
                {
                    crystalReportViewer1.ReportSource = rpt;
                }
            }
            catch (Exception ex)
            {
                //crystalReportViewer1.ReportSource = rpt;
                btnSQLStatement.Text = "ERROR: " + ex.Message;
            }

            // if no groups in report hide the group tree in viewer
            if (rptClientDoc.DataDefController.DataDefinition.Groups.Count == 0)
            {
                crystalReportViewer1.SetFocusOn(UIComponent.Page);
                crystalReportViewer1.ToolPanelView = ToolPanelViewType.None;
            }
            else
            {
                /// this sets the focus on the group set to none and then sets it to the group again to get rid of the dotted line around the main viewer area.
                crystalReportViewer1.ToolPanelView = ToolPanelViewType.None;
                crystalReportViewer1.SetFocusOn(UIComponent.GroupTree);
                crystalReportViewer1.ToolPanelView = ToolPanelViewType.GroupTree;
                crystalReportViewer1.SetFocusOn(UIComponent.GroupTree);
            }

            // next line causes a delay due to the engine having to go to the last page.
            //if (rpt.Rows.Count > 0)
            //{
            //    rpt.FormatEngine.GetLastPageNumber(new ReportPageRequestContext());
            //}

            // set up the format export types:
            int myFOpts = (int)(
                CrystalDecisions.Shared.ViewerExportFormats.RptFormat |
                CrystalDecisions.Shared.ViewerExportFormats.PdfFormat |
                CrystalDecisions.Shared.ViewerExportFormats.RptrFormat |
                CrystalDecisions.Shared.ViewerExportFormats.XLSXFormat |
                CrystalDecisions.Shared.ViewerExportFormats.CsvFormat |
                CrystalDecisions.Shared.ViewerExportFormats.EditableRtfFormat |
                CrystalDecisions.Shared.ViewerExportFormats.ExcelRecordFormat |
                CrystalDecisions.Shared.ViewerExportFormats.RtfFormat |
                CrystalDecisions.Shared.ViewerExportFormats.WordFormat |
                CrystalDecisions.Shared.ViewerExportFormats.XmlFormat |
                CrystalDecisions.Shared.ViewerExportFormats.ExcelFormat |
                CrystalDecisions.Shared.ViewerExportFormats.ExcelRecordFormat);
            //CrystalDecisions.Shared.ViewerExportFormats.NoFormat); // no exports allowed
            //int myFOpts = (int)(CrystalDecisions.Shared.ViewerExportFormats.AllFormats);

            crystalReportViewer1.AllowedExportFormats = myFOpts;
            crystalReportViewer1.EnableToolTips = true;
            //crystalReportViewer1.ReuseParameterValuesOnRefresh = true;
            crystalReportViewer1.ShowParameterPanelButton = true;

            CrystalDecisions.Windows.Forms.ParameterPromptControl ParamClick = new ParameterPromptControl();
            //ParamClick.Capture.ToString();

            GroupPath gp = new GroupPath();
            string tmp = String.Empty;
            if (IsLoggedOn)
            {
                try
                {
                    rptClientDoc.RowsetController.GetSQLStatement(gp, out tmp);
                    btnSQLStatement.Text += "\n" + tmp + "\n";
                }
                catch (Exception ex)
                {
                    btnSQLStatement.Text += "\nERROR: " + ex.Message;
                }
            }
        }

        private void SetData_Click(object sender, EventArgs e)
        {
            string connString = "Provider=SQLOLEDB;Data Source=VMMSSQL2K8;Database=xtreme;User ID=sb;Password=1Oem2000";
            string sqlString = "SELECT 'Orders'.'Customer ID', 'Orders'.'Order Date' FROM 'xtreme'.'dbo'.'Orders' 'Orders'";
            string sqlString2 = "Select * From \"Orders Detail\"";

            OleDbConnection oleConn = new OleDbConnection(connString);
            OleDbDataAdapter oleAdapter = new OleDbDataAdapter(sqlString, oleConn);
            OleDbDataAdapter oleAdapter2 = new OleDbDataAdapter(sqlString2, oleConn);
            DataTable dt1 = new DataTable("Query");
            DataTable dt2 = new DataTable("\"Orders Detail\"");

            oleAdapter.Fill(dt1);
            oleAdapter2.Fill(dt2);

            System.Data.DataSet ds = new System.Data.DataSet();
            ds.Tables.Add(dt1);
            ds.Tables.Add(dt2);
            
            ds.WriteXml(@"D:\Atest\list_of_workflows.xml");
            ds.WriteXmlSchema(@"D:\Atest\list_of_workflows.xsd");

            ds.ReadXml(@"D:\Atest\Dev Element\data_ReplicateCRIssue.xml", XmlReadMode.ReadSchema);

            ISCRDataSet DS1 = (ISCRDataSet) CrystalDecisions.ReportAppServer.DataSetConversion.DataSetConverter.Convert(ds);
            
            // uses this for OLE DB DS record set
            rptClientDoc.DatabaseController.SetDataSource(DS1, "FunctionSales", "FunctionTotalPrice");

            MessageBox.Show("Data Source Set", "RAS", MessageBoxButtons.OK, MessageBoxIcon.Information);
            IsRpt = false;
        }

        private void SetParam_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition pField in (ParameterFieldDefinitions)rpt.DataDefinition.ParameterFields)
                {
                    if (pField.ParameterType == ParameterType.StoreProcedureParameter)
                    {
                        rpt.SetParameterValue("servProvCode", "1");
                        rpt.SetParameterValue("servProvCode2", "1");
                        //rpt.SetParameterValue("@StartDate", "12/23/2017 00:00:00");
                        //rpt.SetParameterValue("@EndDate", "1/23/2018 00:00:00");
                        //rpt.SetParameterValue("@ResultsBy", "Queue");
                        //rpt.SetParameterValue("@Detail", "All");
                        //rpt.SetParameterValue("@QueueGroups", "All");
                    }
                    if (pField.ParameterType == ParameterType.ReportParameter)
                    {
                        if (pField.Name == "StartDate")
                            rpt.SetParameterValue("StartDate", "12/08/2017 00:00:00");
                        if (pField.Name == "EndDate")
                            rpt.SetParameterValue("EndDate", "12/09/2017 00:00:00");
                        if (pField.Name == "ResultsBy")
                            rpt.SetParameterValue("ResultsBy", "Queue");
                        if (pField.Name == "Detail")
                            rpt.SetParameterValue("Detail", "All");
                        if (pField.Name == "QueueGroups")
                            rpt.SetParameterValue("QueueGroups", "All");
                        if (pField.Name != "StartDate" && pField.Name != "EndDate" && pField.Name != "ResultsBy" && pField.Name != "Detail" && pField.Name != "QueueGroups")
                        {
                            if (pField.ValueType.ToString() == "StringField")
                                rpt.SetParameterValue(pField.Name, "");
                        }

                        //if (pField.ParameterFieldName == "My Parameter") // && pField.ParameterFieldName != "Pm-".Substring(0, 3))
                        //    SetCrystalParam(rpt, "@My Parameter", "1005");
                        ////if (pField.ParameterFieldName == "Pm-Orders.Order ID")
                        //    SetCrystalParam(rpt, "@Pm-Orders.Order ID", "1005");
                        //if (pField.ParameterFieldName == "Don")
                        //    SetCrystalParam(rpt, "@Don", "2/23/2000 00:00:00");
                        //if (pField.ParameterFieldName == "My Parameter")
                        //    SetCrystalParam(rpt, "@My Parameter", "1002");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            btnSQLStatement.Text += "\nParameters Set";
        }

        private void ReportDocumentSetParameters(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {
            try
            {
                SetCrystalParam(rpt, "@percentage", "100");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //CrystalReportViewer viewer = new CrystalReportViewer();
            //viewer.ReportSource = rpt;
            //viewer.RefreshReport();
            //return rpt.ParameterFields&#91;0&#93;.HasCurrentValue;

            //SetCrystalParam(rpt, "@FROMSEQ", "0000");
            //SetCrystalParam(rpt, "@TOSEQ", "9999");
            //SetCrystalParam(rpt, "@REPRINT", "1");
            //SetCrystalParam(rpt, "@MULTICUR", "Y");
            //SetCrystalParam(rpt, "@CMPNAME", "Sample Company Litmited");
            //SetCrystalParam(rpt, "@FUNCDEC", "2");
            //SetCrystalParam(rpt, "@INCTAX", "1");


            ////SetCrystalParamsToNull(rpt);
            //SetCrystalParam(rpt, "@CMPNAME", "Universal Corporation");
            //SetCrystalParam(rpt, "@LEVEL1NAME", "Contract");
            //SetCrystalParam(rpt, "@FROMCONT", "1");
            //SetCrystalParam(rpt, "@TOCONT", "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ");
            //SetCrystalParam(rpt, "@FROMUFMTCONT", "1");
            //SetCrystalParam(rpt, "@TOUFMTCONT", "ZZZZZZZZZZZZZZZZ");
            //SetCrystalParam(rpt, "@LEVEL2NAME", "Project");
            //SetCrystalParam(rpt, "@FROMPROJ", "1");
            //SetCrystalParam(rpt, "@TOPROJ", "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ");
            //SetCrystalParam(rpt, "@FROMCUSTOMER", "1");
            //SetCrystalParam(rpt, "@TOCUSTOMER", "ZZZZZZZZZZZZ");
            //SetCrystalParam(rpt, "@FROMCONTSTART", "19000101");
            //SetCrystalParam(rpt, "@TOCONTSTART", "99991231");
            //SetCrystalParam(rpt, "@FROMPROJSTART", "19000101");
            //SetCrystalParam(rpt, "@TOPROJSTART", "99991231");
            //SetCrystalParam(rpt, "@CONTSTATUS", "0");
            //SetCrystalParam(rpt, "@PROJSTATUS", "0");
            //SetCrystalParam(rpt, "@SWINCLLABOR", "1");
            //SetCrystalParam(rpt, "@SWINCLOVERHEAD", "1");
            //SetCrystalParam(rpt, "@SWCUSTCUR", "1");
            //SetCrystalParam(rpt, "@SWMULTI", "1");
            //SetCrystalParam(rpt, "@LEVEL4NAME", "Contracts");
            //SetCrystalParam(rpt, "@HCURN", "CAD");
            //SetCrystalParam(rpt, "@HCURNDEC", "2");
            //SetCrystalParam(rpt, "@USEPERCENT", "1");
            //SetCrystalParam(rpt, "@LEVEL3NAME", "Category");
            ////rpt.SetParameterValue("LEVEL3NAME", "Category");
            ////rpt.SetParameterValue("USEPERCENT", "1", "category");

            foreach (CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition pField in (ParameterFieldDefinitions)rpt.DataDefinition.ParameterFields)
            {
                switch (pField.Name)
                {
                    case "@percentage":
                        {
                            SetCrystalParam(rpt, "@percentage", "100");
                        }
                        break;

                    //case "@st":
                    //    {
                    //        SetCrystalParam(rpt, "@st", "184");
                    //    }
                    //    break;
                    //case "@sg":
                    //    {
                    //        SetCrystalParam(rpt, "@sg", "1,1");
                    //    }
                    //    break;
                    //case "@si":
                    //    {
                    //        SetCrystalParam(rpt, "@si", null);
                    //    }
                    //    break;
                    //case "@sd":
                    //    {
                    //        SetCrystalParam(rpt, "@sd", "4/1/1825 00:00:00");
                    //    }
                    //    break;
                    //case "@ed":
                    //    {
                    //        SetCrystalParam(rpt, "@ed", "4/1/2012 00:00:00");
                    //    }
                    //    break;
                    //case "@Range":
                    //    {
                    //        SetCrystalParam(rpt, "@Range", "M");
                    //    }
                    //    break;
                    //case "@uid":
                    //    {
                    //        SetCrystalParam(rpt, "@uid", "3");
                    //    }
                    //    break;
                    //case "@rid":
                    //    {
                    //        SetCrystalParam(rpt, "@rid", "6");
                    //    }
                    //    break;
                    //case "test":
                    //    {
                    //        SetCrystalParam(rpt, "test", "6");
                    //    }
                    //    break;
                }
            }
        }

        private void SetCrystalParam(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt, string parameterName, string listOfValues)
        {
            int ParCount = rpt.DataDefinition.ParameterFields.Count;

            CrystalDecisions.ReportAppServer.Controllers.ParameterFieldController myNewPAram = new ParameterFieldController();
            rptClientDoc = rpt.ReportClientDocument;

            rpt.ParameterFields.Add(parameterName.ToString());
                
            foreach (CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition crParamField in rpt.DataDefinition.ParameterFields)
            {
                //if (crParamField.Name.ToLower() == "@" + parameterName.ToLower())
                if (crParamField.ReportName == "") // && crParamField.ReportName != "Pm-".Substring(0,3))
                {
                    try
                    {
                        CrystalDecisions.Shared.ParameterValues myparameterValues = new CrystalDecisions.Shared.ParameterValues();
                        CrystalDecisions.Shared.ParameterDiscreteValue crDiscreteValue = new CrystalDecisions.Shared.ParameterDiscreteValue();

                        crDiscreteValue.Value = listOfValues; // "1";
                        myparameterValues.IndexOf(crParamField.Name);
                        myparameterValues.Add(crDiscreteValue);
                        crParamField.ApplyCurrentValues(myparameterValues);
                        //return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        //return;
                    }
                }
                    //else
                    //{
                    //try
                    //{
                    //    CrystalDecisions.Shared.ParameterValues myparameterValues = new CrystalDecisions.Shared.ParameterValues();
                    //    CrystalDecisions.Shared.ParameterDiscreteValue crDiscreteValue = new CrystalDecisions.Shared.ParameterDiscreteValue();

                    //    crDiscreteValue.Value = listOfValues; // "0";
                    //    myparameterValues.IndexOf(crParamField.Name);
                    //    myparameterValues.Add(crDiscreteValue);
                    //    crParamField.ApplyCurrentValues(myparameterValues);
                    //    //return;
                    //}
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show(ex.Message);
                    //    //return;
                    //}
                //}
            }
        }

        private void Export_Click(object sender, EventArgs e)
        {
            CrystalDecisions.Shared.PdfRtfWordFormatOptions pdfOpts = CrystalDecisions.Shared.ExportOptions.CreatePdfRtfWordFormatOptions();
            CrystalDecisions.Shared.ExcelDataOnlyFormatOptions excelOptsDataOnly = CrystalDecisions.Shared.ExportOptions.CreateDataOnlyExcelFormatOptions();
            CrystalDecisions.Shared.ExcelFormatOptions excelOpts = CrystalDecisions.Shared.ExportOptions.CreateExcelFormatOptions();
            CrystalDecisions.Shared.MicrosoftMailDestinationOptions mailOpts = CrystalDecisions.Shared.ExportOptions.CreateMicrosoftMailDestinationOptions();
            CrystalDecisions.Shared.DiskFileDestinationOptions diskOpts = CrystalDecisions.Shared.ExportOptions.CreateDiskFileDestinationOptions();
            CrystalDecisions.Shared.ExportOptions exportOpts = new CrystalDecisions.Shared.ExportOptions();
            CrystalDecisions.Shared.PdfFormatOptions PDFExpOpts = new CrystalDecisions.Shared.PdfFormatOptions();
            CrystalDecisions.Shared.HTMLFormatOptions HTMLExpOpts = CrystalDecisions.Shared.ExportOptions.CreateHTMLFormatOptions(); 
            CrystalDecisions.Shared.ExchangeFolderDestinationOptions ExchExpOpts = new ExchangeFolderDestinationOptions();
            CrystalDecisions.Shared.TextFormatOptions TxtExpOpts = new TextFormatOptions();
            CrystalDecisions.Shared.CharacterSeparatedValuesFormatOptions csvExpOpts = new CrystalDecisions.Shared.CharacterSeparatedValuesFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.HTMLExportFormatOptions RasHTMLExpOpts = new HTMLExportFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.TextExportFormatOptions txtFmtOpts = new TextExportFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions RASexportOpts = new CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.CharacterSeparatedValuesExportFormatOptions RASCSVExptOpts = new CharacterSeparatedValuesExportFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.PDFExportFormatOptions RasPDFExpOpts = new PDFExportFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.DataOnlyExcelExportFormatOptions RasExcelOpts = new DataOnlyExcelExportFormatOptions();
            CrystalDecisions.ReportAppServer.ReportDefModel.ExcelExportFormatOptions myExcel = new ExcelExportFormatOptions();

            CrystalDecisions.ReportAppServer.ReportDefModel.CrReportExportFormatEnum RASXLXSExportOpts = CrystalDecisions.ReportAppServer.ReportDefModel.CrReportExportFormatEnum.crReportExportFormatXLSX;

            rptClientDoc = rpt.ReportClientDocument;

            CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions ExPOEx = new CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions();

            //csvExpOpts.GroupSectionsOption = CsvExportSectionsOption.DoNotExport;
            //RASCSVExptOpts.GroupSectionsOption = CrCSVExportSectionsOptionEnum.crCSVExportSectionsOptionDoNotExport;

            //// This works do not modify
            string MyRptName = rpt.FileName.ToString();
            int PathIndex = rpt.FileName.ToString().LastIndexOf(@"\");
            int lastIndex = MyRptName.Substring(MyRptName.LastIndexOf(@"\") + 1).Length;
            MyRptName = MyRptName.Substring(MyRptName.LastIndexOf(@"\") + 1, (rpt.FileName.Length - 3) - (MyRptName.LastIndexOf(@"\") + 1)) + "pdf";

            diskOpts.DiskFileName = @rpt.FileName.Substring(9, PathIndex - 8) + MyRptName;
            exportOpts.ExportFormatType = ExportFormatType.PortableDocFormat;
            exportOpts.ExportFormatOptions = new PdfFormatOptions();
            exportOpts.ExportDestinationOptions = diskOpts;
            exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;

            rpt.Export(exportOpts);
            MessageBox.Show("Export to PDF Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //// This works do not modify

            //// This works do not modify
            //// This requires you to set the report export option in the report designer first

            //diskOpts.DiskFileName = @"D:\Atest\Without.csv";

            //csvExpOpts = new CharacterSeparatedValuesFormatOptions();

            //csvExpOpts.ExportMode = CsvExportMode.Standard;
            //csvExpOpts.GroupSectionsOption = CsvExportSectionsOption.ExportIsolated;
            //csvExpOpts.ReportSectionsOption = CsvExportSectionsOption.DoNotExport;
            //csvExpOpts.SeparatorText = ",";
            //csvExpOpts.Delimiter = "|";

            //exportOpts.DestinationOptions = diskOpts;
            //exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
            //exportOpts.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
            //exportOpts.ExportFormatOptions = csvExpOpts;

            //try
            //{
            //    rpt.Export(exportOpts);
            //}
            //catch (Exception ex)
            //{
            //    btnSQLStatement.Text = ex.Message;
            //    return;
            //}
            //// This works do not modify

            //// This works do not modify
            //RASexportOpts.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatPDF;
            //RASexportOpts.FormatOptions = RASCSVExptOpts;

            //rptClientDoc.PrintOutputController.ExportEx(RASexportOpts).Save(@"D:\Atest\CortezSPs1.pdf", true);
            //MessageBox.Show("RAS Export to PDF Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //// This works do not modify


            // This works do not modify
            //rpt.Load(@"C:\CortezReports\CortezSPs1.rpt");
            //rptClientDoc = rpt.ReportClientDocument;

            //RASCSVExptOpts.ExportMode = CrCSVExportModeEnum.crCSVExportModeStandard;
            ////RASCSVExptOpts.GroupSectionsOption = CrCSVExportSectionsOptionEnum.crCSVExportSectionsOptionExportIsolated;
            ////RASCSVExptOpts.ReportSectionsOption = CrCSVExportSectionsOptionEnum.crCSVExportSectionsOptionDoNotExport;
            //RASCSVExptOpts.Separator = ",";
            //RASCSVExptOpts.Delimiter = "\"";

            //RASexportOpts.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatCharacterSeparatedValues;
            //RASexportOpts.FormatOptions = RASCSVExptOpts;

            //rptClientDoc.PrintOutputController.ExportEx(RASexportOpts).Save(@"D:\Atest\471324\CortezSPs1.csv", true);
            //MessageBox.Show("Export to CSV Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // This works do not modify

            //// This works do not alter
            //string MyRptName = rpt.FileName.ToString();
            //MyRptName = "D:\\Atest\\" + MyRptName.Substring(MyRptName.LastIndexOf(@"\") + 1, (rpt.FileName.Length - 3) - (MyRptName.LastIndexOf(@"\") + 1)) + "xls";
            //try
            //{
            //    if (File.Exists(MyRptName))
            //    {
            //        File.Delete(MyRptName);
            //    }

            //    CrystalDecisions.ReportAppServer.ReportDefModel.ExcelExportFormatOptions RasXLSExpOpts = new ExcelExportFormatOptions();
            //    //RasXLSExpOpts = rptClientDoc.get_SavedExportOptions(CrReportExportFormatEnum.crReportExportFormatMSExcel);
            //    // textBox1 = "Excel - BaseAreaGroupNumber:       " + RasXLSExpOpts.BaseAreaGroupNumber.ToString() + "\n";
            //    //textBox1 += "Excel - BaseAreaType:              " + RasXLSExpOpts.BaseAreaType.ToString() + "\n";
            //    //textBox1 += "Excel - FormulaExportPageAreaType: " + RasXLSExpOpts.ExportPageAreaPairType.ToString() + "\n";
            //    //textBox1 += "Excel - ExportPageBreaks:          " + RasXLSExpOpts.ExportPageBreaks.ToString() + "\n";
            //    //textBox1 += "Excel - ConstantColWidth:          " + RasXLSExpOpts.ConstantColWidth.ToString() + "\n";
            //    //textBox1 += "Excel - ConvertDatesToStrings:     " + RasXLSExpOpts.ConvertDatesToStrings.ToString() + "\n";
            //    //textBox1 += "Excel - StartPageNumber:           " + RasXLSExpOpts.StartPageNumber.ToString() + "\n";
            //    //textBox1 += "Excel - EndPageNumber:             " + RasXLSExpOpts.EndPageNumber.ToString() + "\n";
            //    //textBox1 += "Excel - ExportPageBreaks:          " + RasXLSExpOpts.ExportPageBreaks.ToString() + "\n";
            //    //textBox1 += "Excel - MRelativeObjectPosition:   " + RasXLSExpOpts.MaintainRelativeObjectPosition.ToString() + "\n";
            //    //textBox1 += "Excel - ShowGridlines:             " + RasXLSExpOpts.ShowGridlines.ToString() + "\n";
            //    //textBox1 += "Excel - UseConstantColWidth:       " + RasXLSExpOpts.UseConstantColWidth.ToString() + "\n";

            //    // Set them now:
            //    RasXLSExpOpts.BaseAreaType = CrAreaSectionKindEnum.crAreaSectionKindPageHeader;
            //    RasXLSExpOpts.UseConstantColWidth = false;
            //    RasXLSExpOpts.ShowGridlines = false;
            //    //RasXLSExpOpts.StartPageNumber = 3;
            //    //RasXLSExpOpts.EndPageNumber = 10;

            //    // Save the udpated info
            //    //rptClientDoc.set_SavedExportOptions(CrReportExportFormatEnum.crReportExportFormatMSExcel, RasXLSExpOpts);

            //    CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions exportOpts1 = new CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions();
            //    exportOpts1.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatMSExcel;
            //    exportOpts1.FormatOptions = RasXLSExpOpts;

            //    // And Export
            //    rptClientDoc.PrintOutputController.ExportEx(exportOpts1).Save(MyRptName, true);
            //    MessageBox.Show("Export to Excel Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //// This works do not alter

            ////DO not CHANGE This Works
            //// this gets the report name and sets the export name to be the same less the extension
            //string MyRptName = rpt.FileName.ToString();
            //MyRptName = "Daily Sales Report";

            //HTMLExportFormatOptions htmlOpts = new HTMLExportFormatOptions();
            //HTMLExpOpts.HTMLBaseFolderName = @"c:\test\Daily Sales Report\SDK\";
            //HTMLExpOpts.HTMLFileName = MyRptName; // @"c:\test\" + MyRptName;
            //HTMLExpOpts.HTMLHasPageNavigator = false;
            //HTMLExpOpts.HTMLEnableSeparatedPages = false;
            //HTMLExpOpts.HTMLHasPageSection = true;
            //// New to SP 3 - Engine only, no RAS equivolent
            //// To export the same way as CR Designer does
            //exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
            //exportOpts.ExportFormatType = ExportFormatType.HTML40;
            //exportOpts.ExportFormatOptions = HTMLExpOpts;

            //rpt.Export(exportOpts);

            //string sa = HTMLExpOpts.HTMLFileName;

            //try
            //{ // this is cool, don't delete it. appends files doens't work for html but will for others
            //    string[] tmpfiles = Directory.GetFiles(sa, "*.html");
            //    using (StreamWriter writer = File.CreateText(HTMLExpOpts.HTMLFileName + @"\Master.html"))
            //    {
            //        foreach (string tempFile in tmpfiles)
            //        {
            //            using (StreamReader reader = File.OpenText(tempFile))
            //            {
            //                writer.Write(reader.ReadToEnd());
            //            }
            //        }
            //    }
            //}
            //catch (Exception ex)
            //{
            //    btnSQLStatement.Text = "ERROR: " + ex.Message;
            //    return;
            //}             

            // RAS doesn't work
            //RasHTMLExpOpts.BaseFile = @"c:\test\" + MyRptName;
            //RasHTMLExpOpts.PageNavigator = false;
            //RasHTMLExpOpts.SeparatePages = false;

            //RASexportOpts.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatMHTML;
            //RASexportOpts.FormatOptions = RasHTMLExpOpts;

            //rptClientDoc.PrintOutputController.ExportEx(RASexportOpts);

            //MessageBox.Show("Export to HTML Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ////DO not CHANGE This Works
                
            ////DO not CHANGE This Works
            //CrystalDecisions.CrystalReports.Engine.ReportDocument rpt1 = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            //ISCDReportClientDocument rcd;

            ////rpt1.ReportClientDocument.LocaleID = CrystalDecisions.ReportAppServer.DataDefModel.CeLocale.ceLocaleGerman;
            //rpt1.Load(@"D:\Atest\Bhushan\World Sales Report.rpt");

            //rcd = rpt1.ReportClientDocument;
            //HTMLExportFormatOptions htmlOpts = new HTMLExportFormatOptions();

            //htmlOpts.BaseFile = @"D:\Atest\html\World Sales Report.html";
            //htmlOpts.PageNavigator = false;
            //htmlOpts.SeparatePages = false;
            //CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions exportOpts1 = new CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions();
            //exportOpts1.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatMHTML; // crReportExportFormatHTML;
            //exportOpts1.FormatOptions = htmlOpts;

            //rcd.PrintOutputController.ExportEx(exportOpts1).Save(@"c:\test\html6\World Sales Report.html", true);
            //MessageBox.Show("Export to HTML Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //rpt1.Close();
            //rpt1.Dispose();
            ////DO not CHANGE This Works

            // Export to RAS HTML 4.0 NOT SUPPORTED 
            ////HTMLExportFormatOptions htmlOpts = new HTMLExportFormatOptions();
            //HTMLExpOpts.HTMLBaseFolderName = @"c:\test\" + MyRptName.ToString();
            //HTMLExpOpts.HTMLFileName = "\\" + MyRptName + ".html";
            //HTMLExpOpts.HTMLHasPageNavigator = false;
            //HTMLExpOpts.HTMLEnableSeparatedPages = false;
            //HTMLExpOpts.HTMLHasPageSection = true;
            ////New to SP 3 - Engine only, no RAS equivolent
            //HTMLExpOpts.HTMLEnableSeparatedPages = true;

            //exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
            //exportOpts.ExportFormatType = ExportFormatType.HTML32;
            //exportOpts.ExportFormatOptions = HTMLExpOpts;
            //rpt.Export(exportOpts);
            //MessageBox.Show("Export to HTML Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            ////// this works DO NOT EDIT
            //ExchExpOpts.FolderPath = "YourName@YourCompany.com#Inbox"; // Should be an exchange folder...
            //ExchExpOpts.Profile = "Default Outlook Profile"; //_exchangeProfile; // "MS Exchange Settings"; // "Outlook";
            //ExchExpOpts.Password = "password";

            //exportOpts.ExportFormatOptions = PDFExpOpts;
            //exportOpts.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.ExchangeFolder;
            //exportOpts.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat;
            //exportOpts.ExportDestinationOptions = ExchExpOpts;
            //rpt.Export(exportOpts);
            //MessageBox.Show("Mail to Exchange folder sent Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ////// this works DO NOT EDIT

            ////// this works DO NOT EDIT
            //exportOpts.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.MicrosoftMail;
            //exportOpts.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat;
            ////exportOpts.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.HTML32;

            //exportOpts.ExportFormatOptions = pdfOpts;
            ////pdfOpts.FirstPageNumber = 3;
            ////pdfOpts.LastPageNumber = 3;
            ////pdfOpts.UsePageRange = true;

            //// this does not work, name of the file is the Title saved in the RPT
            //rpt.SummaryInfo.ReportTitle = "Dons Report";

            //mailOpts.MailToList = "Don Williams";
            //mailOpts.MailSubject = "Attached is a PDF file - .net Export test ";
            //mailOpts.MailSubject = "This is the Subject";
            ////MailOpts.MailCCList = "John Doe";
            //mailOpts.UserName = "YourName@YourCompany.com";
            //mailOpts.Password = "Password";
            //exportOpts.ExportDestinationOptions = mailOpts;
            //rpt.Export(exportOpts);
            //MessageBox.Show("Mail sent Completed", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ////// this works DO NOT EDIT

        }

        String subreportLinkValue = "1003"; // MUST CHANGE THIS

        private void SetLogonInfo_Click(object sender, EventArgs e)
        {
            CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
            CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
            CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
            CrystalDecisions.CrystalReports.Engine.Database crDatabase;
            CrystalDecisions.CrystalReports.Engine.Tables crTables;
            TableLogOnInfo crTableLogOnInfo;

            CrystalDecisions.Shared.ConnectionInfo crConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();

            //btrDataFile.Text = rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.ServerName.ToString();
            //btrSearchPath.Text = rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.DatabaseName.ToString();
            //btrFileLocation.Text = rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.UserID.ToString();
            //btrPassword.Text = rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.Password.ToString();

            rpt.DataSourceConnections.ToString();

            //set up the database and tables objects to refer to the current report
            crDatabase = rpt.Database;
            crTables = crDatabase.Tables;

            crConnectioninfo.ServerName = btrDataFile.Text.ToString();
            crConnectioninfo.UserID = btrFileLocation.Text.ToString(); // "sa"; // "sa"; // Hana i819003

            int tableIndex = 0;
            bool mainSecureDB;

            if (chkTrusted.CheckState != CheckState.Checked)
            {
                crConnectioninfo.Password = btrPassword.Text.ToString(); //  "1Oem2000"; // Hana Password4
                crConnectioninfo.DatabaseName = btrSearchPath.Text.ToString();
            }
            else
            {
                crConnectioninfo.IntegratedSecurity = true;
                mainSecureDB = true;
                //mainSecureDB = rpt.Database.Tables[tableIndex].LogOnInfo.ConnectionInfo.IntegratedSecurity;
                //btrFileLocation.Text = "<UseTrustedNTAuthentication>";
                //btrPassword.Text = "<UseTrustedNTAuthentication>";
            }

            CrystalDecisions.Shared.DataSourceConnections DSCon;
            //rpt.DataSourceConnections.Count.ToString();

            ////Get the Database Tables Collection for your report
            //CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
            //boTables = rptClientDoc.DatabaseController.Database.Tables;

            //CrystalDecisions.ReportAppServer.DataDefModel.Procedure boTable = new CrystalDecisions.ReportAppServer.DataDefModel.Procedure();
            ////For each table in the report:
            //// - Set the Table Name properties.
            //// - Set the table location in the report to use the new modified table
            //boTable.Name = "cspRpt_Astellas_Query_Rate_Per_Month_Wk;1";
            //boTable.QualifiedName = "astellastest.dbo.cspRpt_Astellas_Query_Rate_Per_Month_Wk;1";
            //boTable.Alias = "cspRpt_Astellas_Query_Rate_Per_Month_Wk;1";

            //rpt.ReportClientDocument.DatabaseController.SetTableLocation(boTables[0], boTable);

            // Don't delete this line - used for testing report with Parms.
            //ReportDocumentSetParameters(rpt);
            bool ConWorks = false;

            CrystalDecisions.Shared.NameValuePair2 nvp2 = new NameValuePair2();
            if (crDatabase.Tables.Count != 0)
                nvp2 = (NameValuePair2)rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.Attributes.Collection[1];

            //NameValuePair2 test = rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.LogonProperties.LookupNameValuePair("UseDSNProperties");
            ////MessageBox.Show(test.Name.ToString());
            //test.Value = true;

            //((CrystalDecisions.Shared.NameValuePair2)((new System.Collections.ArrayList.ArrayListDebugView(((CrystalDecisions.Shared.DbConnectionAttributes)(((CrystalDecisions.Shared.NameValuePair2)
            //    ((new System.Collections.ArrayList.ArrayListDebugView(rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.Attributes.Collection)).Items[3])).Value)).Collection)).Items[2])).Value = true;

            //CrystalDecisions.Shared.DbConnectionAttributes nvp2DSN = 
            //rpt.Database.Tables[0].LogOnInfo.ConnectionInfo.Attributes.Collection)).Items[3])).Value)).Collection)).Items[2]) = true;

            //btnSQLStatement.Text = "Data Sources are the not the same: Database: " + nvp2.Value.ToString();

            //SetCrystalParam(rpt, "servProvCode", "1");
            //SetCrystalParam(rpt, "servProvCode2", "1");

            //loop through all the tables and pass in the connection info
            foreach (CrystalDecisions.CrystalReports.Engine.Table crTable in crTables)
            {
                mainSecureDB = rpt.Database.Tables[tableIndex].LogOnInfo.ConnectionInfo.IntegratedSecurity;
                string mainTableName = crTable.Name.ToString();
                tableIndex++;

                //btrDataFile - Server
                //btrSearchPath - Database
                //btrFileLocation - User Name 
                //btrPassword - password

                //pass the necessary parameters to the connectionInfo object
                crConnectioninfo.ServerName = btrDataFile.Text.ToString(); 
                if (!mainSecureDB)
                {
                    crConnectioninfo.ServerName = btrDataFile.Text.ToString();
                    crConnectioninfo.UserID = btrFileLocation.Text.ToString();
                    crConnectioninfo.Password = btrPassword.Text.ToString(); 
                    crConnectioninfo.DatabaseName = btrSearchPath.Text.ToString(); 
                }
                else
                {
                    crConnectioninfo.IntegratedSecurity = true;
                }

                crTable.LogOnInfo.ConnectionInfo.DatabaseName = crConnectioninfo.DatabaseName;
                crTable.LogOnInfo.ConnectionInfo.UserID = btrFileLocation.Text.ToString(); ;
                crTable.LogOnInfo.ConnectionInfo.Password = btrPassword.Text.ToString(); ;
                crTableLogOnInfo = crTable.LogOnInfo;
                crTableLogOnInfo.ConnectionInfo = crConnectioninfo;

                // this is where you set the main report parameter values
                // if the subreport is not linked by the record selection formula it returns the first value, or last value depending on the subreport link
                // so in this case where the subreport is linked via the record selection formula the value needs to be passed to the subreport Link.
                //subreportLinkValue = "EnterMainReportParametrValue";
                //SetCrystalParam(rpt, "@file_id", subreportLinkValue.ToString());

                try
                {
                    crTable.ApplyLogOnInfo(crTableLogOnInfo);
                }
                catch (Exception ex)
                {
                    btnSQLStatement.Text += "ApplyLogOnInfo failed: " + ex.ToString();
                    btnSQLStatement.AppendText("\n");
                }

                //// Postgres
                //try
                //{
                //    crTable.Location = btrSearchPath.Text.ToString() + ".public." + crTable.Location; //"orders"; // crTable.Location;
                //}
                //catch (Exception ex)
                //{
                //    if (ex.Message == "Fetching SQL statements is not supported for this report.")
                //    {
                //        CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = rpt.ReportClientDocument.DatabaseController;
                //        ISCRTable oldTable = (ISCRTable)databaseController.Database.Tables[0];

                //        if (nvp2.Name.ToString() == crConnectioninfo.DatabaseName.ToString())
                //        {
                //            if (((dynamic)oldTable).CommandText.ToString() != null)
                //            {
                //                btnSQLStatement.Text = "Main Report Command: " + ((dynamic)oldTable).CommandText.ToString();
                //                btnSQLStatement.Text += "\n";
                //            }
                //        }
                //        else
                //            btnSQLStatement.Text += "\nWrong log on Server Name: Database: " + nvp2.Value.ToString() + " does not exist\n";
                //    }
                //    else
                //    {
                //        btnSQLStatement.Text = "ERROR: " + ex.Message;
                //    }
                //}

                // SQL Server
                try
                {
                    crTable.Location = btrSearchPath.Text.ToString() + ".dbo." + crTable.Location; //"orders"; // crTable.Location;
                }
                catch (Exception ex)
                {
                    if (ex.Message == "Fetching SQL statements is not supported for this report.")
                    {
                        CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = rpt.ReportClientDocument.DatabaseController;
                        ISCRTable oldTable = (ISCRTable)databaseController.Database.Tables[0];

                        if (nvp2.Name.ToString() == crConnectioninfo.DatabaseName.ToString())
                        {
                            if (((dynamic)oldTable).CommandText.ToString() != null)
                            {
                                btnSQLStatement.Text = "Main Report Command: " + ((dynamic)oldTable).CommandText.ToString();
                                btnSQLStatement.Text += "\n";
                            }
                        }
                        else
                            btnSQLStatement.Text += "\nWrong log on Server Name: Database: " + nvp2.Value.ToString() + " does not exist\n";
                    }
                    else
                    {
                        btnSQLStatement.Text = "ERROR: " + ex.Message;
                    }
                }

                try
                {
                    if (crTable.TestConnectivity())
                    {
                        ConWorks = true;
                        btnSQLStatement.Text += "Test Connectivity worked\n";
                    }
                    else
                        ConWorks = false;
                }
                catch (Exception ex)
                {
                    btnSQLStatement.Text += "Test Connectivity failed: " + ex.ToString();
                }


                //rpt.VerifyDatabase();

                //ISCRTable CurTable = rptClientDoc.DatabaseController.Database.Tables.FindTableByAlias("Customer1");
                //if (CurTable.Alias == "Customer1")
                //{
                //    ISCRTable NewTable = CurTable.Clone();
                //    NewTable.Name = "Customer";
                //    NewTable.QualifiedName = "xtreme.dbo.Customer";
                //    NewTable.Alias = "CustomerYoshi";
                //    rptClientDoc.DatabaseController.SetTableLocation(CurTable, NewTable);
                //    IsRpt = false;
                //}


                //crTable.Location = "Addresses_cws";
                //crTable.Name = "Address"; // read only 

                // SQL Server
                //crTable.Location = crDatabase + ".dbo." + "Addresses_cws"; // crTable.Location;
                //crTable.Location = "exec cspRpt_Astellas_Query_Rate_Per_Month_Wk NULL, NULL, NULL, NULL, NULL, NULL, NULL";

                //foreach (CrystalDecisions.CrystalReports.Engine.Table tbl in rpt.Database.Tables)
                //{
                //    tbl.Location = "Addresses_cws";
                //}

            }

            if (ConWorks || crDatabase.Tables.Count == 0)
            {
                // test if no main report DB then go to subreport logon
                if (crDatabase.Tables.Count != 0)
                {
                    //gp.FromString("2-0");
                    GroupPath gp = new GroupPath();
                    string tmp = String.Empty;
                    try
                    {
                        rptClientDoc.RowsetController.GetSQLStatement(gp, out tmp);
                        btnSQLStatement.Text = tmp;
                        btnSQLStatement.AppendText("\n");
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message == "Fetching SQL statements is not supported for this report.")
                        {
                            CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = rpt.ReportClientDocument.DatabaseController;
                            ISCRTable oldTable = (ISCRTable)databaseController.Database.Tables[0];

                            if (nvp2.Name.ToString() == crConnectioninfo.DatabaseName.ToString())
                            {
                                if (((dynamic)oldTable).CommandText.ToString() != null)
                                {
                                    btnSQLStatement.Text += "Main Report Command: " + ((dynamic)oldTable).CommandText.ToString();
                                    btnSQLStatement.Text += "\n";
                                }
                            }
                            else
                                btnSQLStatement.Text += "\nWrong log on Server Name: Database: " + nvp2.Value.ToString() + " does not exist\n";
                        }
                        else
                        {
                            if (ex.Message == "Missing parameter values.")
                            {
                                // set parameter values
                                ReportDocumentSetParameters(rpt);
                            }
                            else
                                btnSQLStatement.Text = "ERROR: " + ex.Message;
                            //return;
                        }
                    }
                }

                //set the crSections object to the current report's sections
                CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;
                int flcnt = 0;
                bool SecureDB;

                //loop through all the sections to find all the subreport objects
                foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                {
                    crReportObjects = crSection.ReportObjects;
                    //loop through all the report objects to find all the subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                    {
                        if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                        {
                            try
                            {
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();

                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                string mysubname = crSubreportObject.SubreportName.ToString();
                                //mysubname = " ";
                                //open the subreport object
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                CrystalDecisions.CrystalReports.Engine.Database crSubDatabase;
                                CrystalDecisions.CrystalReports.Engine.Tables crSubTables;
                                TableLogOnInfo crTableSubLogOnInfo;

                                CrystalDecisions.Shared.ConnectionInfo crSubConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();

                                //set the database and tables objects to work with the subreport
                                crSubDatabase = crSubreportDocument.Database;
                                //crSubTables = crSubTables.Tables;
                                tableIndex = 0;

                                // get the DB name from the report
                                CrystalDecisions.Shared.NameValuePair2 nvp2Sub = (NameValuePair2)crSubreportDocument.Database.Tables[tableIndex].LogOnInfo.ConnectionInfo.Attributes.Collection[1];

                                // Log on and Get subreport SQl 
                                SubreportController subreportController = rptClientDoc.SubreportController;
                                SubreportClientDocument subreportClientDocument = subreportController.GetSubreport(crSubreportObject.SubreportName);

                                // this gets the subreport linked fields so if more than one you will need to handle the values individually
                                // since the main report parameter and record selection formula are only linked in the record selection formula you need to get the actual field the subreport is linked on to be able to get the SQL statement

                                try
                                {
                                    SubreportLinks SubLinks = rptClientDoc.SubreportController.GetSubreportLinks(crSubreportDocument.Name.ToString());
                                    for (int I = 0; I < SubLinks.Count; I++)
                                    {
                                        SubreportLink subLink = SubLinks[I];
                                        string trimCurlies = subLink.LinkedParameterName.ToString();
                                        char[] charsToTrim = { '{', '}', '?' }; // remove the {? ... } from the field name
                                        trimCurlies = trimCurlies.Trim(charsToTrim);
                                        subreportClientDocument.DataDefController.ParameterFieldController.SetCurrentValue(crSubreportObject.SubreportName.ToString(), trimCurlies.ToString(), subreportLinkValue);  
                                    }

                                    //SetSubCrystalParam(subreportClientDocument, "Pm-Orders.Order ID", "1002");

                                    subreportClientDocument.DatabaseController.LogonEx(crConnectioninfo.ServerName, crConnectioninfo.DatabaseName, crConnectioninfo.UserID, crConnectioninfo.Password);

                                    GroupPath gp1 = new GroupPath();
                                    gp1.FromString("");
                                    string sql = String.Empty;
                                    subreportClientDocument.RowsetController.GetSQLStatement(gp1, out sql);

                                    btnSQLStatement.Text += "\nSubreport: " + crSubreportObject.SubreportName.ToString() + "\n" + sql;
                                    btnSQLStatement.AppendText("\n");
                                    //break;
                                }
                                catch (Exception ex)
                                {
                                    {
                                        if (nvp2Sub.Name.ToString() == crConnectioninfo.DatabaseName.ToString())
                                        {
                                            if (ex.Message == "Fetching SQL statements is not supported for this report.")
                                            {
                                                CrystalDecisions.ReportAppServer.Controllers.DatabaseController databaseController = rpt.ReportClientDocument.SubreportController.GetSubreport(crSubreportObject.SubreportName.ToString()).DatabaseController;
                                                ISCRTable oldTable = (ISCRTable)databaseController.Database.Tables[0];
                                                btnSQLStatement.Text += "\nSubreport: " + crSubreportObject.SubreportName.ToString() + " - Command:\n" + ((dynamic)oldTable).CommandText.ToString();
                                                btnSQLStatement.AppendText("\n");
                                            }
                                            else
                                            {
                                                btnSQLStatement.Text += "Logon Subreport: " + crSubreportObject.SubreportName.ToString() + " - ERROR: " + ex.Message;
                                                btnSQLStatement.AppendText("\n");
                                            }
                                        }
                                        else
                                            btnSQLStatement.Text += "\nWrong Subreport log on Server Name or Database: " + nvp2Sub.Value.ToString() + " does not exist or \ngetting SQL from subreport not supported\n";
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                {
                                    if (ex.Message == "Fetching SQL statements is not supported for this report.")
                                    {
                                        btnSQLStatement.Text += "ERROR: " + ex.Message;
                                        btnSQLStatement.AppendText("\n");
                                    }
                                    else
                                    {
                                        //btnSQLStatement.Text += "ERROR: " + ex.Message;
                                        //btnSQLStatement.AppendText("\n");
                                    }
                                }
                            }
                        }
                    }
                }
                if (btrVerifyDatabase.Checked)
                    rpt.VerifyDatabase();
            }
            else
            {
                btnSQLStatement.Text += "No data source or failed to log on";
            }

            IsLoggedOn = true;

            // update the SQL info
            getReportOptionsOnOpen(rptClientDoc);
        }

        private void SubReportDocumentSetParameters(SubreportClientDocument subreportClinetDocument)
        {
            SetSubCrystalParam(subreportClinetDocument, "@My Parameter", "1");
            //SetSubCrystalParam(subreportClinetDocument, "@LEVEL1NAME", "Contract");
        }

        private void SetSubCrystalParam(SubreportClientDocument subreportClinetDocument, string parameterName, string listOfValues)
        {
            int ParCount = rpt.DataDefinition.ParameterFields.Count;

            foreach (CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition crParamField in rpt.DataDefinition.ParameterFields)
            {

                if ("@" + crParamField.Name.ToLower() == parameterName.ToLower())
                {
                    try
                    {
                        CrystalDecisions.Shared.ParameterValues myparameterValues = new CrystalDecisions.Shared.ParameterValues();
                        CrystalDecisions.Shared.ParameterDiscreteValue crDiscreteValue = new CrystalDecisions.Shared.ParameterDiscreteValue();

                        crDiscreteValue.Value = listOfValues;
                        myparameterValues.IndexOf(crParamField.Name);
                        myparameterValues.Add(crDiscreteValue);
                        crParamField.ApplyCurrentValues(myparameterValues);
                        return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }

                // this should catch if it's a CR Linked Parameter
                if (crParamField.Name.ToLower() == parameterName.ToLower())
                {
                    try
                    {
                        CrystalDecisions.Shared.ParameterValues myparameterValues = new CrystalDecisions.Shared.ParameterValues();
                        CrystalDecisions.Shared.ParameterDiscreteValue crDiscreteValue = new CrystalDecisions.Shared.ParameterDiscreteValue();

                        CrystalDecisions.Shared.ParameterField crParamValue = new CrystalDecisions.Shared.ParameterField();

                        crDiscreteValue.Value = listOfValues;
                        myparameterValues.IndexOf(crParamField.Name);
                        myparameterValues.Add(crDiscreteValue);
                        crParamField.ApplyCurrentValues(myparameterValues);
                        return;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        return;
                    }
                }
            }
        }

        private void TextObjects_Click(object sender, EventArgs e)
        {
            string textBox1 = String.Empty;
            textBox1 = rptClientDoc.DisplayName.ToString();
            btnReportObjects.Text = String.Empty;

            // List all fields used in the document
            foreach (CrystalDecisions.ReportAppServer.DataDefModel.Field resultField in rptClientDoc.DataDefController.DataDefinition.ResultFields)
            {
                textBox1 = resultField.FormulaForm.ToString();
                btnReportObjects.AppendText("\n");
                btnReportObjects.Text += textBox1;
                //MessageBox.Show(resultField.FormulaForm, "text objects", MessageBoxButtons.OK, MessageBoxIcon.Information);

            } 
        }

        private void ReportObjectComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string textBox1 = String.Empty;
            string textBox2 = String.Empty;
            string MyObjectType = ReportObjectComboBox1.SelectedItem.ToString();
            btnCount.Text = "";
            btnErrorCount.Text = "";
            CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj;

            switch (MyObjectType)
            {
                case "Formula Fields":
                    #region FormulaFields
                    btnReportObjects.Text = "";
                    int flcnt = 0;
                    int Errorcnt = 0;

                    #region do not delete
                    ////do not delete - this works ... sort of
                    ////rptClientDoc.DataDefController.FormulaFieldController.Remove(resultField);
                    //CrystalDecisions.ReportAppServer.ObjectFactory.ObjectFactory objFactory = new CrystalDecisions.ReportAppServer.ObjectFactory.ObjectFactory();
                    //CrystalDecisions.ReportAppServer.DataDefModel.FormulaField Formula = (CrystalDecisions.ReportAppServer.DataDefModel.FormulaField)objFactory.CreateObject("CrystalReports.FormulaField");
                    //Formula.Type = CrystalDecisions.ReportAppServer.DataDefModel.CrFieldValueTypeEnum.crFieldValueTypeStringField;
                    //Formula.Syntax = CrystalDecisions.ReportAppServer.DataDefModel.CrFormulaSyntaxEnum.crFormulaSyntaxCrystal;
                    //Formula.Text = @"whilereadingrecords; ""A"""; // "hello Ludek = 1"; // resultField.Text; // "n=3"; {Customer.Customer Credit ID} + 1 @"whilereadingrecords; ""A"""
                    //Formula.Name = "TestDon"; // "TestDon"; //  resultField.Name; //"testformula";
                    //String FormulaMessage = rptClientDoc.DataDefController.FormulaFieldController.Check(Formula);
                    //if (FormulaMessage == null)
                    //    rptClientDoc.DataDefController.FormulaFieldController.Add(Formula);
                    //else
                    //    btnReportObjects.Text += "There are errors in the formula: " + FormulaMessage.ToString() + "\n";
                    ////do not delete - this works
                    #endregion do not delete

                    btnCount.Text = rptClientDoc.DataDefController.DataDefinition.FormulaFields.Count.ToString();

                    foreach (FormulaField resultField in rptClientDoc.DataDefController.DataDefinition.FormulaFields)
                    {
                        String FormulaMessage; 
                        textBox1 = resultField.LongName.ToString();
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" : Used: " + resultField.UseCount.ToString() + " times'End' ");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                        try
                        {
                            FormulaMessage = "\n" + rptClientDoc.DataDefController.FormulaFieldController.Check(resultField);
                        }
                        catch (Exception ex)
                        {
                            btnReportObjects.Text += "\n" + ex.Message + "\n";
                            FormulaMessage = "";

                            //MessageBox.Show("ERROR: " + ex.Message); // + " ;" + ex.InnerException.Message);

                            ////cool but of no use to release mode
                            //string myURL = @"http://search.sap.com/ui/scn#query=" + ex.Message + "&startindex=1&filter=scm_a_site%28scm_v_Site11%29&filter=scm_a_modDate%28*%29&timeScope=all";
                            //string fixedString = myURL.Replace(" ", "%20");

                            ////System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                            //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);

                            ////string myURL = @"C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports 2011\Help\en\crw.chm";
                            ////System.Diagnostics.Process.Start(myURL);
                            ////cool but of no use to release mode

                        }

                        if (FormulaMessage != null)
                        {
                            btnReportObjects.Text += FormulaMessage.ToString() + "\n";
                            if (FormulaMessage != "\n") // || FormulaMessage.Substring(0, 14) == "Error in formula") //\nError in formula  ~: \n'NumberVar DoBVar := IIF((100 * MONTH(currentdate) + DAY(currentdate)) < (100 * MONTH({Command.BirthDate}) + DAY({Command.BirthDate})), 1, 0);\r'\nThis field name is not known.\nDetails: errorKind"
                            {
                                FormulaMessage = @"This field name is not known";
                                btnReportObjects.AppendText("\n");
                                ++Errorcnt;
                                btnErrorCount.Text = Errorcnt.ToString();

                                //string myURL = @"http://search.sap.com/ui/scn#query=" + FormulaMessage + "&startindex=1&filter=scm_a_site%28scm_v_Site11%29&filter=scm_a_modDate%28*%29&timeScope=all";
                                //string fixedString = myURL.Replace(" ", "%20");

                                ////System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);

                            }
                        }
                    }

                    // now look in the subreports also
                    CrystalDecisions.ReportAppServer.ReportDefModel.ReportObjects rptObjs1;
                    rptObjs1 = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects();
                    int SubObjects1 = rpt.Subreports.Count;
                    btnReportObjects.Text += "Engine Subreport count: " + SubObjects1.ToString() + "\n\n";

                    try
                    {
                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj1 in rptObjs1)
                        {
                            switch (rptObj1.Kind)
                            {
                                // look for sub report object and display info
                                case CrReportObjectKindEnum.crReportObjectKindSubreport:
                                    SubreportClientDocument mySub;
                                    CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject subObj1;
                                    subObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject)rptObj1;
                                    mySub = (SubreportClientDocument)rpt.ReportClientDocument.SubreportController.GetSubreport(subObj1.SubreportName);

                                    textBox1 = "Subreport Name: " + subObj1.SubreportName + "\n";

                                    foreach (FormulaField resultField in mySub.DataDefController.DataDefinition.FormulaFields)
                                    {
                                        String FormulaMessage;
                                        textBox1 += "Formula: " + resultField.LongName.ToString();
                                        btnReportObjects.Text += textBox1;
                                        btnReportObjects.AppendText(" : Used: " + resultField.UseCount.ToString() + " times'End'\n");
                                        ++flcnt;
                                        btnCount.Text = flcnt.ToString();
                                        try
                                        {
                                            FormulaMessage = "\n" + rptClientDoc.DataDefController.FormulaFieldController.Check(resultField);
                                        }
                                        catch (Exception ex)
                                        {
                                            btnReportObjects.Text += "\n" + ex.Message + "\n";
                                            FormulaMessage = "";
                                            btnCount.Text = flcnt.ToString();
                                            break;
                                        }
                                        textBox1 = "";

                                        if (FormulaMessage != null)
                                        {
                                            btnReportObjects.Text += FormulaMessage.ToString() + "\n";
                                            if (FormulaMessage != "\n") 
                                            {
                                                FormulaMessage = @"This field name is not known";
                                                btnReportObjects.AppendText("\n");

                                                btnReportObjects.Text += "Formula Text: " + resultField.Text + "\n";
                                                btnReportObjects.AppendText("\n");

                                                ++Errorcnt;
                                                btnErrorCount.Text = Errorcnt.ToString();

                                                //string myURL = @"http://search.sap.com/ui/scn#query=" + FormulaMessage + "&startindex=1&filter=scm_a_site%28scm_v_Site11%29&filter=scm_a_modDate%28*%29&timeScope=all";
                                                //string fixedString = myURL.Replace(" ", "%20");

                                                ////System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                                                //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);

                                            }
                                        }
                                        btnReportObjects.AppendText("End of Subreport Formula: " + rptObj1.Name + "\n\n");
                                    }
                                        break;
                            }
                        }
                    }


                    catch (Exception ex)
                    {
                        btnReportObjects.AppendText("Error in: " + ex.Message + "\n");
                        ++Errorcnt;
                        btnErrorCount.Text = Errorcnt.ToString();
                    }



                    #endregion FormulaFields
                    break;
                case "Fields used in the report":
                    #region Fields Used
                    btnReportObjects.Text = "";
                    flcnt = 0;

                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table tbl in rptClientDoc.DataDefController.Database.Tables)
                    {
                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.DBField dbfld in tbl.DataFields)
                        {
                            if (dbfld.UseCount > 0)
                            {
                                textBox1 = dbfld.FormulaForm.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" : Used: " + dbfld.UseCount.ToString() + " times'End' \n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                            }
                        }
                    }
                    #endregion Fields Used
                    break;
                case "Object Name -> Field Name":
                    #region OBJtoFName
                    btnReportObjects.Text = "";
                    textBox1 = "";
                    flcnt = 0;

                    CrystalDecisions.ReportAppServer.ReportDefModel.ReportObjects rptObjs;
                    //rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetReportObjectsByKind(CrReportObjectKindEnum.crReportObjectKindField);
                    rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects();

                    foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj1 in rptObjs)
                    {
                        switch (rptObj1.Kind)
                        {
                            case CrReportObjectKindEnum.crReportObjectKindField:
                                CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject fldObj1;
                                fldObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)rptObj1;
                                
                                textBox1 = fldObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += fldObj1.DataSourceName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();

                                //// This works do not change
                                //if (fldObj1.Name == "LINETOTALTAXAMOUNT1")
                                //{
                                //    CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject OldfieldObject = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)rptObj1;
                                //    CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject NewfieldObject = new CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject();

                                //    //OldfieldObject.CopyTo(NewfieldObject, true);

                                //    NewfieldObject = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)OldfieldObject.Clone(true);

                                //    CrystalDecisions.ReportAppServer.ReportDefModel.NumericFieldFormat numericFieldFormat = NewfieldObject.FieldFormat.NumericFormat;
                                //    //CrystalDecisions.ReportAppServer.ReportDefModel.CrCurrencySymbolTypeEnum = CrCurrencySymbolTypeEnum.crCurrencySymbolTypeFixedSymbol;

                                //    numericFieldFormat.NDecimalPlaces = 8;
                                //    numericFieldFormat.EnableUseLeadZero = false;
                                //    //numericFieldFormat.ThousandSymbol = "&";
                                //    //numericFieldFormat.ThousandsSeparator = true;
                                //    numericFieldFormat.NegativeFormat = CrystalDecisions.ReportAppServer.ReportDefModel.CrNegativeTypeEnum.crNegativeTypeLeadingMinus;
                                //    numericFieldFormat.RoundingFormat = CrystalDecisions.ReportAppServer.ReportDefModel.CrRoundingTypeEnum.crRoundingTypeRoundToTenBillionth;

                                //    rptClientDoc.ReportDefController.ReportObjectController.Modify(OldfieldObject, NewfieldObject);
                                //    // Another bug in the controller need to call this 2 times to take the Rounding property - same as ADAPT01727457
                                //    rptClientDoc.ReportDefController.ReportObjectController.Modify(OldfieldObject, NewfieldObject);

                                //    IsRpt = false;
                                //}
                                //// This works do not change

                                break;
                            case CrReportObjectKindEnum.crReportObjectKindText:
                                CrystalDecisions.ReportAppServer.ReportDefModel.TextObject txtObj;
                                txtObj = (CrystalDecisions.ReportAppServer.ReportDefModel.TextObject)rptObj1;
                                textBox1 = txtObj.Name.ToString();
                                textBox1 += " -> " + txtObj.Text.ToString();
                                textBox1 += "\nEnableSuppressEmbedBlankLines: " + ((dynamic) rptObj1).TextObjectFormat.EnableSuppressEmbedBlankLines + "\n";
                                textBox1 += "SectionCode: " + ((dynamic) rptObj1).SectionCode;
                                btnReportObjects.Text += textBox1 + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();

                                //// This works do not change
                                //if (txtObj.Name == "Text1")
                                //{
                                //    CrystalDecisions.ReportAppServer.ReportDefModel.TextObject OldTextObject = (CrystalDecisions.ReportAppServer.ReportDefModel.TextObject)rptObj1;
                                //    CrystalDecisions.ReportAppServer.ReportDefModel.TextObject NewTextObject = new CrystalDecisions.ReportAppServer.ReportDefModel.TextObject();

                                //    //OldfieldObject.CopyTo(NewfieldObject, true);

                                //    NewTextObject = (CrystalDecisions.ReportAppServer.ReportDefModel.TextObject)OldTextObject.Clone(true);
                                //    NewTextObject.Border.BackgroundColor = 255;

                                //    rptClientDoc.ReportDefController.ReportObjectController.Modify(OldTextObject, NewTextObject);
                                //    //rptClientDoc.ReportDefController.ReportObjectController.Modify(OldTextObject, NewTextObject);

                                //    IsRpt = false;
                                //}
                                //// This works do not change

                                break;

                            case CrReportObjectKindEnum.crReportObjectKindBox:
                                CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject boxObj1;
                                CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject boxObj2 = new CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject();
                                boxObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject)rptObj1;
                                boxObj2 = boxObj1;

                                //// modify the box size
                                //boxObj.Bottom = 360;
                                //boxObj1.Height = 2385;
                                //boxObj1.Format.EnableCanGrow = true;
                                //boxObj1.EndSectionName = "ReportFooterSection1";
                                //boxObj1.EnableExtendToBottomOfSection = true;
                                //rptClientDoc.ReportDefController.ReportObjectController.Modify(boxObj, boxObj1);

                                textBox1 = boxObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += boxObj1.Name.ToString();
                                textBox1 += " : X - ";
                                textBox1 += (boxObj1.Left / 1440.00).ToString();
                                textBox1 += " : Y - ";
                                textBox1 += ((boxObj1.Right - boxObj1.Left) / 1440.00).ToString();
                                textBox1 += " : Width - ";
                                textBox1 += (boxObj1.Width / 1440.00).ToString();
                                textBox1 += " : Height - ";
                                textBox1 += (boxObj1.Height / 1440.00).ToString();
                                btnReportObjects.Text += textBox1 + "\n";
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                IsRpt = false;

                                break;

                            case CrReportObjectKindEnum.crReportObjectKindPicture:
                                textBox1 = rptObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += rptObj1.Name.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "Height: " + rptObj1.Height.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Width: " + rptObj1.Width.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Top: " + rptObj1.Top.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Left: " + rptObj1.Left.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindChart:
                                textBox1 = rptObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += rptObj1.Name.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "Height: " + rptObj1.Height.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Width: " + rptObj1.Width.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Top: " + rptObj1.Top.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Left: " + rptObj1.Left.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindSubreport:
                                CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject subObj1;
                                subObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject)rptObj1;

                                textBox1 = subObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += subObj1.SubreportName.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindOlapGrid:
                                CrystalDecisions.ReportAppServer.ReportDefModel.OlapGridObject OLAPObj1;
                                OLAPObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.OlapGridObject)rptObj1;

                                textBox1 = OLAPObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += OLAPObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindMap:
                                CrystalDecisions.ReportAppServer.ReportDefModel.MapObject MapObj1;
                                MapObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.MapObject)rptObj1;

                                textBox1 = MapObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += MapObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindLine:
                                CrystalDecisions.ReportAppServer.ReportDefModel.LineObject LineObj1;
                                LineObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.LineObject)rptObj1;

                                textBox1 = LineObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += LineObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindFlash:
                                CrystalDecisions.ReportAppServer.ReportDefModel.FlashObject FlashObj1;
                                FlashObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.FlashObject)rptObj1;

                                textBox1 = FlashObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += FlashObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindFieldHeading:
                                CrystalDecisions.ReportAppServer.ReportDefModel.FieldHeadingObject FieldHeadingObj1;
                                FieldHeadingObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldHeadingObject)rptObj1;

                                textBox1 = FieldHeadingObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += FieldHeadingObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindCrosstab:
                                CrystalDecisions.ReportAppServer.ReportDefModel.CrossTabObject CrossObj1;
                                CrossObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.CrossTabObject)rptObj1;

                                textBox1 = CrossObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += CrossObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            case CrReportObjectKindEnum.crReportObjectKindBlobField:
                                CrystalDecisions.ReportAppServer.ReportDefModel.BlobFieldObject BlobObj1;
                                BlobObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.BlobFieldObject)rptObj1;

                                textBox1 = BlobObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += BlobObj1.Name.ToString() + " - ";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                                btnReportObjects.AppendText("ToolTip: " + rptObj1.Format.ToolTipText.ToString() + "\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;

                            //case CrReportObjectKindEnum.crReportObjectKindInvalid:
                            //    CrystalDecisions.ReportAppServer.ReportDefModel. InvalidObj1;
                            //    BlobObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.BlobFieldObject)rptObj1;

                            //    textBox1 = BlobObj1.Name.ToString();
                            //    textBox1 += " -> ";
                            //    textBox1 += BlobObj1.Name.ToString() + " - ";
                            //    btnReportObjects.Text += textBox1;
                            //    btnReportObjects.Text += "Section Name: " + rptObj1.SectionCode.ToString() + " " + rptObj1.SectionName.ToString() + "\n";
                            //    btnReportObjects.AppendText("\n");
                            //    ++flcnt;
                            //    btnCount.Text = flcnt.ToString();
                            //    break;
                            //    //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindInvalid;
                        }
                    }
                    #endregion OBJtoFName
                    break;
                case "Groups":
                    #region Groups
                    btnReportObjects.Text = "";
                    flcnt = 0;

                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.Group resultField in rptClientDoc.DataDefController.DataDefinition.Groups)
                    {
                        textBox1 = resultField.ConditionField.FormulaForm.ToString();
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                     
                        ISCRAreas GrAreas = rptClientDoc.ReportDefController.ReportDefinition.Areas;
                        foreach (ISCRArea area in GrAreas)
                        {
                            //int xxxx = 0;
                            if (area.Kind == CrAreaSectionKindEnum.crAreaSectionKindGroupHeader)
                            {
                                int xxx = GrAreas.Count;
                                foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSectName in area.Sections)
                                {
                                    for (int xx = 0; xx < xxx; xx++)
                                    {
                                        if (((dynamic)(GrAreas[xx].Name == (dynamic)area.Name)))
                                        {
                                            btnReportObjects.AppendText("\nRAS:\nBackground Color: " + crSectName.Format.BackgroundColor.ToString());
                                            btnReportObjects.AppendText("\nSuppress If Blank: " + crSectName.Format.EnableSuppressIfBlank.ToString());
                                            btnReportObjects.AppendText("\nSuppress: " + crSectName.Format.EnableSuppress.ToString());
                                            btnReportObjects.AppendText("\nKeep Together: " + crSectName.Format.EnableKeepTogether.ToString());
                                            btnReportObjects.AppendText("\nNew Page Before: " + crSectName.Format.EnableNewPageBefore.ToString());
                                            btnReportObjects.AppendText("\nNew Page After: " + crSectName.Format.EnableNewPageAfter.ToString());
                                            btnReportObjects.AppendText("\nPrint At Bottom Of Page: " + crSectName.Format.EnablePrintAtBottomOfPage.ToString());
                                            btnReportObjects.AppendText("\nReset Page Number After: " + crSectName.Format.EnableResetPageNumberAfter.ToString());
                                            btnReportObjects.AppendText("\nReset Underlay Section: " + crSectName.Format.EnableUnderlaySection.ToString());
                                            btnReportObjects.AppendText("\n\n");
                                            textBox1 = "";
                                            //return;
                                        }
                                    }

                                    // add summary sort

                                    //DataDefController dataDefController = rptClientDoc.DataDefController;
                                    //CrystalDecisions.ReportAppServer.DataDefModel.Group group = dataDefController.DataDefinition.Groups[0];

                                    //CrystalDecisions.ReportAppServer.DataDefModel.Fields resultFields = dataDefController.DataDefinition.ResultFields;
                                    //CrystalDecisions.ReportAppServer.DataDefModel.Field fieldToSummarize = (CrystalDecisions.ReportAppServer.DataDefModel.Field)resultFields.FindField("{Employee_Options.EMPLOYEE Full Name}", CrystalDecisions.ReportAppServer.DataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameFormula, CrystalDecisions.ReportAppServer.DataDefModel.CeLocale.ceLocaleEnglishUS);

                                    //SummaryFieldController summaryFieldController = dataDefController.SummaryFieldController;

                                    //if (summaryFieldController.CanSummarizeOn(fieldToSummarize))
                                    //{
                                    //    SummaryField summaryField = new SummaryFieldClass();
                                    //    summaryField.Group = group;
                                    //    summaryField.SummarizedField = fieldToSummarize;
                                    //    summaryField.Operation = CrSummaryOperationEnum.crSummaryOperationDistinctCount;
                                    //    summaryFieldController.Add(-1, summaryField);
                                    //}

                                    ////Above code will automatically add the summary field object to group footer section. 
                                    ////Without a summaryField in GroupHeader/GroupFooter section, in CRW designer, “Group Sort Export” item will be gray out in menu.
                                    ////If you want to change group sort to Top N, please just add following code:


                                    //CrystalDecisions.ReportAppServer.DataDefModel.ISCRField NewSummaryField = dataDefController.DataDefinition.SummaryFields[0];

                                    //SortController sortController = dataDefController.SortController;

                                    //TopNSort topNSort = new TopNSortClass();
                                    //topNSort.SortField = NewSummaryField;
                                    //topNSort.Direction = CrSortDirectionEnum.crSortDirectionNoSortOrder;

                                    //topNSort.DiscardOthers = false;
                                    //topNSort.NIndividualGroups = 5;
                                    //topNSort.NotInTopBottomName = "Other";

                                    //sortController.Add(-1, topNSort);


                                    // add a summary based on this field Summaryfields has no controller so only a get to add to Group change after adding to field summary.
                                    //DistinctCount (Employee_Options.EMPLOYEE Full Name, Employee_Options.EMPLOYEE Hire Date, "Monthly")

                                    //m_boFieldIndex = rptClientDoc.DataDefController.DataDefinition.SummaryFields.Find("Sum ({frhsrg.pr_bedr}, {frhsrg.artcode})",
                                    //    //m_boFieldIndex = m_boReportClientDocument.DataDefController.DataDefinition.SummaryFields.Find("Sum ({@F_Quantity}, {frhsrg.artcode})",
                                    //    CrystalDecisions.ReportAppServer.DataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameFormula,
                                    //                      CrystalDecisions.ReportAppServer.DataDefModel.CeLocale.ceLocaleUserDefault);


                                    //m_boField = (Field)rptClientDoc.DataDefController.DataDefinition.SummaryFields[m_boFieldIndex];

                                    //rptClientDoc.DataDefController.DataDefinition.SummaryFields[0].FormulaForm.ToString() = @"DistinctCount (Employee_Options.EMPLOYEE Full Name, Employee_Options.EMPLOYEE Hire Date, 'Monthly')";

                                    #region GroupNote
                                    //CrystalDecisions.CrystalReports.Engine.Areas areas = rpt.ReportDefinition.Areas;
                                    //foreach (CrystalDecisions.CrystalReports.Engine.Area ENarea in areas)
                                    //{
                                    //    if (ENarea.Kind == AreaSectionKind.GroupHeader)
                                    //    {
                                    //        CrystalDecisions.CrystalReports.Engine.GroupAreaFormat groupAreaFormat;
                                    //        groupAreaFormat = (CrystalDecisions.CrystalReports.Engine.GroupAreaFormat)ENarea.AreaFormat;
                                    //        textBox1 = "ENGINE:\nEnableRepeatGroupHeader: " + groupAreaFormat.EnableRepeatGroupHeader.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableNewPageAfter: " + groupAreaFormat.EnableNewPageAfter.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableHideForDrillDown: " + groupAreaFormat.EnableHideForDrillDown.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableKeepGroupTogether: " + groupAreaFormat.EnableKeepGroupTogether.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableKeepTogether: " + groupAreaFormat.EnableKeepTogether.ToString(); // always returns false
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableNewPageBefore: " + groupAreaFormat.EnableNewPageBefore.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnablePrintAtBottomOfPage: " + groupAreaFormat.EnablePrintAtBottomOfPage.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableRepeatGroupHeader: " + groupAreaFormat.EnableRepeatGroupHeader.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableResetPageNumberAfter: " + groupAreaFormat.EnableResetPageNumberAfter.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "EnableSuppress: " + groupAreaFormat.EnableSuppress.ToString(); // always returns false
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n\n");

                                    //        CrystalDecisions.ReportAppServer.ReportDefModel.Area oldGHArea = (CrystalDecisions.ReportAppServer.ReportDefModel.Area)rptClientDoc.ReportDefController.ReportDefinition.get_GroupHeaderArea(0);
                                    //        CrystalDecisions.ReportAppServer.ReportDefModel.ISCRGroupAreaFormat grpAreaFormat = (CrystalDecisions.ReportAppServer.ReportDefModel.ISCRGroupAreaFormat)oldGHArea.Format;
                                    //        textBox1 = "RAS:\nKeepGroupTogether: " + grpAreaFormat.EnableKeepGroupTogether.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "VisibleGroupNumberPerPage: " + grpAreaFormat.VisibleGroupNumberPerPage.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n");
                                    //        textBox1 = "RepeatGroupHeader: " + grpAreaFormat.EnableRepeatGroupHeader.ToString();
                                    //        btnReportObjects.Text += textBox1;
                                    //        btnReportObjects.AppendText("\n 'End' \n");
                                    //        textBox1 = "";
                                    //    }
                                    //}

                                    //try
                                    //{
                                    //    //CrystalDecisions.ReportAppServer.ReportDefModel.Area oldGHArea = (CrystalDecisions.ReportAppServer.ReportDefModel.Area)rptClientDoc.ReportDefController.ReportDefinition.get_GroupHeaderArea(0);
                                    //    //CrystalDecisions.ReportAppServer.ReportDefModel.Area newGHArea = new CrystalDecisions.ReportAppServer.ReportDefModel.Area();
                                    //    //CrystalDecisions.ReportAppServer.ReportDefModel.GroupAreaFormat group = (CrystalDecisions.ReportAppServer.ReportDefModel.GroupAreaFormat)oldGHArea.Format.Clone(true);
                                    //    CrystalDecisions.ReportAppServer.ReportDefModel.ISCRGroupAreaFormat grpAreaFormat = (CrystalDecisions.ReportAppServer.ReportDefModel.ISCRGroupAreaFormat)area.Format;
                                    //    //grpAreaFormat.EnableNewPageAfter = true;

                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text != null)
                                    //        textBox1 += "\nGroup BackgroundColor Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeBackgroundColor].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text != null)
                                    //        textBox1 += "\nGroup ClampPageFooter Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableClampPageFooter].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text != null)
                                    //        textBox1 += "\nGroup HideForDrillDown Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableHideForDrillDown].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text != null)
                                    //        textBox1 += "\nGroup KeepTogether Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableKeepTogether].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text != null)
                                    //        textBox1 += "\nGroup NewPageAfter Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text != null)
                                    //        textBox1 += "\nGroup NewPageBefore Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageBefore].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text != null)
                                    //        textBox1 += "\nGroup PrintAtBottomOfPage Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnablePrintAtBottomOfPage].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text != null)
                                    //        textBox1 += "\nGroup ResetPageNumberAfter Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableResetPageNumberAfter].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text != null)
                                    //        textBox1 += "\nGroup Suppress Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text != null)
                                    //        textBox1 += "\nGroup SuppressIfBlank Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppressIfBlank].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text != null)
                                    //        textBox1 += "\nGroup UnderlaySection Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableUnderlaySection].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text != null)
                                    //        textBox1 += "\nGroup GroupNumberPerPage Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text.ToString()) + "\n";
                                    //    if (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text != null)
                                    //        textBox1 += "\nGroup RecordNumberPerPage Conditional formula: \n" + (grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeRecordNumberPerPage].Text.ToString()) + "\n";
                                    //    btnReportObjects.Text += textBox1;
                                    //    btnReportObjects.AppendText(" 'End' \n\n");

                                    //    //// this sets the new page after formula
                                    //    //grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Text = @"if {Orders.Customer ID} < 2 then false"; //.ToString();
                                    //    //grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableNewPageAfter].Syntax = CrFormulaSyntaxEnum.crFormulaSyntaxCrystal;
                                    //    //rptClientDoc.ReportDefController.ReportAreaController.SetProperty(oldGHArea, CrReportAreaPropertyEnum.crReportAreaPropertyFormat, grpAreaFormat);

                                    //    //grpAreaFormat.EnableKeepGroupTogether = true;
                                    //    //grpAreaFormat.VisibleGroupNumberPerPage = 1;
                                    //    //grpAreaFormat.EnableRepeatGroupHeader = true;

                                    //    //grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Text = @"9-8";
                                    //    //grpAreaFormat.ConditionFormulas[CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeGroupNumberPerPage].Syntax = CrFormulaSyntaxEnum.crFormulaSyntaxCrystal;
                                    //    //rptClientDoc.ReportDefController.ReportAreaController.SetProperty(oldGHArea, CrReportAreaPropertyEnum.crReportAreaPropertyFormat, grpAreaFormat);

                                    //}
                                    //catch (Exception ex)
                                    //{
                                    //    btnReportObjects.Text += "\n" + ex.Message + "\n";
                                    //}
                                    #endregion GroupNote
                                }
                            }

                        //// Do NOT delete - how to change date filtering for groups based on a date field
                        //CrystalDecisions.ReportAppServer.DataDefModel.Group oldGroup;
                        //CrystalDecisions.ReportAppServer.DataDefModel.Group newgroup;
                        //CrystalDecisions.ReportAppServer.DataDefModel.DateGroupOptions dgOpts = new CrystalDecisions.ReportAppServer.DataDefModel.DateGroupOptions();

                        //if (resultField.ConditionField.FormulaForm == "{Orders.Order Date}")
                        //{
                        //    dgOpts.DateCondition = CrystalDecisions.ReportAppServer.DataDefModel.CrDateConditionEnum.crDateConditionMonthly;

                        //    oldGroup = rptClientDoc.DataDefController.DataDefinition.Groups[flcnt-1];
                        //    newgroup = oldGroup.Clone(true);

                        //    newgroup.Options = dgOpts;
                        //    rptClientDoc.DataDefController.GroupController.Modify(oldGroup, newgroup); 

                        //    IsRpt = false;
                        }
                    }
                    #endregion Groups
                    break;
                case "Parameter Fields":
                    getParameterFields(rpt);
                    break;
                case "Sorts":
                    #region Sorts
                        btnReportObjects.Text = "";
                        flcnt = 0;
                        int GRsortIndex = 1;
                        int FLDsortIndex = 0;

                        //IsRpt = false;
                        //SortDirection sd = rpt.DataDefinition.SortFields[1].SortDirection;
                        //rpt.DataDefinition.SortFields[1].SortDirection = SortDirection.NoSortOrder;

                        //SortController sortController = rptClientDoc.DataDefController.SortController;
                        ////sortController.Remove(GRsortIndex); // Sort of works for Groups - set the sort type to original order
                        //try
                        //{
                        //    sortController.ModifySortDirection(GRsortIndex, CrSortDirectionEnum.crSortDirectionNoSortOrder);
                        //}
                        //catch
                        //{
                        //}

                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.Sort RASSortField in rptClientDoc.DataDefController.DataDefinition.Sorts)
                        {
                            foreach (SortField crSortField in rpt.DataDefinition.SortFields)
                            {
                                if (RASSortField.SortField.FormulaForm == crSortField.Field.FormulaName)
                                {
                                    if (GRsortIndex >= 0)
                                    {
                                        if (crSortField.SortType.ToString() == CrystalDecisions.Shared.SortFieldType.GroupSortField.ToString())
                                        {
                                            textBox1 = "Group Sort:\n ";
                                            IsRpt = false;
                                            GRsortIndex++;
                                        }
                                        else
                                        {
                                            textBox1 = "Record Sort:\n ";
                                            IsRpt = false;
                                            FLDsortIndex++;
                                        }

                                        textBox1 += crSortField.Field.Name.ToString();
                                        textBox1 += "  ";
                                    }
                                }
                            }

                            textBox1 += RASSortField.SortField.FormulaForm.ToString() + "\n";
                            textBox1 += "  ";
                            textBox1 += RASSortField.Direction.ToString() + "\n";
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText("\n");
                        }
                    #endregion Sorts
                    break;
                case "Summary Fields":
                    #region SummaryFields
                    btnReportObjects.Text = "";
                    flcnt = 0;

                    // add summary sort

                    DataDefController dataDefController = rptClientDoc.DataDefController;
                    ////CrystalDecisions.ReportAppServer.DataDefModel.Group group = dataDefController.DataDefinition.Groups[0];

                    //CrystalDecisions.ReportAppServer.DataDefModel.Fields resultFields = dataDefController.DataDefinition.ResultFields;
                    //CrystalDecisions.ReportAppServer.DataDefModel.Field fieldToSummarize = (CrystalDecisions.ReportAppServer.DataDefModel.Field)resultFields.FindField("{Employee_Options.EMPLOYEE Full Name}", CrystalDecisions.ReportAppServer.DataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameFormula, CrystalDecisions.ReportAppServer.DataDefModel.CeLocale.ceLocaleEnglishUS);

                    //SummaryFieldController summaryFieldController = dataDefController.SummaryFieldController;

                    //if (summaryFieldController.CanSummarizeOn(fieldToSummarize))
                    //{
                    //    SummaryField summaryField = new SummaryFieldClass();
                    //    summaryField.Group = new CrystalDecisions.ReportAppServer.DataDefModel.Group(); //group;
                    //    summaryField.SummarizedField = fieldToSummarize;
                    //    summaryField.Operation = CrSummaryOperationEnum.crSummaryOperationDistinctCount;
                    //    summaryFieldController.Add(-1, summaryField);
                    //}

                    ////Above code will automatically add the summary field object to group footer section. 
                    ////Without a summaryField in GroupHeader/GroupFooter section, in CRW designer, “Group Sort Expert” item will be gray out in menu.
                    ////If you want to change group sort to Top N, please just add following code:

                    //// this works but group must exist
                    //CrystalDecisions.ReportAppServer.DataDefModel.ISCRField NewSummaryField = dataDefController.DataDefinition.SummaryFields[0];
                    //SortController sortController = dataDefController.SortController;
                    //TopNSort topNSort = new TopNSortClass();
                    //topNSort.SortField = NewSummaryField;
                    //topNSort.Direction = CrSortDirectionEnum.crSortDirectionDescendingOrder;

                    //topNSort.DiscardOthers = false;
                    //topNSort.NIndividualGroups = 5;
                    //topNSort.NotInTopBottomName = "Other";

                    //sortController.Add(-1, topNSort);
                    //// this works

                    foreach (SummaryField resultField in rptClientDoc.DataDefController.DataDefinition.SummaryFields)
                    {
                        textBox1 = resultField.LongName.ToString();
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                    }
                    #endregion SummaryFields
                    break;
                case "Running Totals":
                    #region Running Totals
                    btnReportObjects.Text = "";
                    flcnt = 0;

                    // using RAS to get the collection
                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.RunningTotalField resultField in rptClientDoc.DataDefController.DataDefinition.RunningTotalFields)
                    // using Engine to get collection
                    //foreach (CrystalDecisions.CrystalReports.Engine.RunningTotalFieldDefinition resultField in rpt.DataDefinition.RunningTotalFields)
                    {
                        textBox1 = resultField.Name.ToString() + " - " + resultField.SummarizedField.Name.ToString(); // rpt needs .FormulaName - RAS LongName
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" : Used: " + resultField.UseCount.ToString() + " 'End' \n");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                    }
                    #endregion Running Totals
                    break;
                case "SubReports":
                    #region Subreports
                    btnReportObjects.Text = "";
                    flcnt = 0;
                    Errorcnt = 0;

                    btnReportObjects.Text = "";
                    //int flcnt = 0;
                    string SecName = "";
                    string SecName1 = "";

                    CrystalDecisions.CrystalReports.Engine.ReportObjects rptObjsRPT;
                    rptObjsRPT = rpt.ReportDefinition.ReportObjects;

                    rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects();
                    int SubObjects = rpt.Subreports.Count;
                    btnReportObjects.Text += "Engine Sub count: " + SubObjects.ToString() + "\n\n";

                    if (SubObjects == 0)
                        return;

                    int crtoChr = 97;

                    try
                    {
                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj1 in rptObjs)
                        {
                            switch (rptObj1.Kind)
                            {
                                // look for sub report object and display info
                                case CrReportObjectKindEnum.crReportObjectKindSubreport:
                                    CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject subObj1;
                                    try
                                    {
                                        subObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject)rptObj1;
                                        ++flcnt;

                                        SubreportController subreportController = rptClientDoc.SubreportController;
                                        SubreportClientDocument subreportClientDocument = subreportController.GetSubreport(subObj1.SubreportName);

                                        foreach (CrystalDecisions.ReportAppServer.DataDefModel.ParameterField pField in subreportController.GetSubreport(subreportClientDocument.Name).DataDefController.DataDefinition.ParameterFields)
                                        {
                                            Console.WriteLine(pField.UseCount.ToString());
                                            foreach (ParameterFieldValue pValue in pField.CurrentValues)
                                            {
                                                Console.WriteLine(pValue.ToString());
                                            }
                                        }

                                        int SecNameL = SecName.Length;
                                        textBox1 = "Imported: " + subreportClientDocument.SubreportLocation.ToString() + "\n";
                                        int c = (rptObj1.SectionCode / 6000) % 1000;
                                        int d = (rptObj1.SectionCode % 1000); // zero based Section
                                        int dd = (d % 1000) % 50;
                                        int cc = d / 50; // section letter character
                                        int l = d - ((((d) / 26) % 26) * 26);

                                        int sectionCodeArea1 = (rptObj1.SectionCode / 1000) % 1000; // Area
                                        int sectionCodeGroup = (((rptObj1.SectionCode) / 50) % 120); // group sectio
                                        int sectionCodeSection = (rptObj1.SectionCode % 1000); // zero based Section
                                        //int sectionCodeGroup = (((rptObj1.SectionCode) / 25) % 40); // group section
                                        int sectionCodeGroupNo = (((rptObj1.SectionCode) % 25)); // group area

                                        string sectionCodeArea = "";

                                        switch (c)
                                        {
                                            case 2:
                                                sectionCodeArea = "Page Header";
                                                break;
                                            case 1:
                                                sectionCodeArea = "Report Header";
                                                break;
                                            case 7:
                                                sectionCodeArea = "Page Footer";
                                                break;
                                            case 8:
                                                sectionCodeArea = "Report Footer";
                                                break;
                                            case 3:
                                                sectionCodeArea = "Group Header";
                                                break;
                                            case 5:
                                                sectionCodeArea = "Group Footer";
                                                break;
                                            case 4:
                                                sectionCodeArea = "Detail";
                                                break;
                                            default:
                                                break;
                                        }

                                        int RawSectID = 1;
                                        {
                                            if (c == 3 || c == 5) // group header/footer
                                            {
                                                d = (d / 50) + 97;
                                                textBox1 += SecName.ToString() + " " + sectionCodeArea + " " + (char)(dd + 49) + (char)(97 + cc) + "\n SubObjt: " + rptObj1.Name.ToString() + " -> Name: " + subObj1.SubreportName.ToString() + "\n";
                                                btnReportObjects.Text += SecName1 + "Don #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr) + ": " + "\n";

                                            }
                                            else
                                            {
                                                textBox1 += SecName.ToString() + " " + sectionCodeArea + " ";

                                                {
                                                    if (RawSectID <= 26 && d <= 26) // a to z
                                                    {
                                                        btnReportObjects.Text += textBox1 + " " + (char)(97 + dd) + "\n SubObjt: " + rptObj1.Name.ToString() + " -> Name: " + subObj1.SubreportName.ToString() + "\n";
                                                        btnReportObjects.Text += SecName1 + "Don #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr) + ": " + "\n";
                                                    }
                                                    else
                                                    { // aa and up
                                                        //btnReportObjects.AppendText(SecName);
                                                        char b = (char)(((((d) / 26) % 26) - 1) + (crtoChr));
                                                        btnReportObjects.Text += textBox1 + " " + b + (char)(97 + l) + "\n SubObjt: " + rptObj1.Name.ToString() + " -> Name: " + subObj1.SubreportName.ToString() + "\n";
                                                        btnReportObjects.Text += textBox1 + "Don #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr) + "\n";

                                                        //+ " : " + RawSectID + " - ";
                                                    }
                                                }

                                            }

                                        }

                                        btnReportObjects.Text += textBox1;
                                        btnReportObjects.Text += " Section Name: " + rptObj1.SectionCode.ToString() + "  " + rptObj1.SectionName.ToString() + "\n";
                                        btnReportObjects.AppendText("ToolTip: " + flcnt.ToString() + rptObj1.Format.ToolTipText.ToString() + "\n");
                                        textBox1 = "";

                                        btnCount.Text = flcnt.ToString(); // SubObjects

                                        //// get the SQL THIS REALLY SLOWS THE REPORT PROCESSING DOWN IF NOT CONNECTED FIRST
                                        //GroupPath gp = new GroupPath();
                                        //string tmp = String.Empty;
                                        //btnReportObjects.Text += "Subreport SQL: ";
                                        //try
                                        //{
                                        //    subreportClientDocument.RowsetController.GetSQLStatement(gp, out tmp);
                                        //    btnReportObjects.Text += tmp;
                                        //}
                                        //catch (Exception ex)
                                        //{
                                        //    btnReportObjects.Text += "ERROR: " + ex.Message;
                                        //    //return;
                                        //}

                                        // seems like subreports with saved data cannot be had...
                                        //if (rpt.HasSavedData == true)
                                        //{
                                        //    btnHasSavedData.Checked = true;

                                        //    // get the record count
                                        //    RowsetController boRowsetController;
                                        //    RowsetCursor boRowsetCursor;
                                        //    GroupPath gp1 = new GroupPath();
                                        //    int rows;
                                        //    //CrystalDecisions.ReportAppServer.CrystalReportDataRowView rowsetMetaData =  new CrystalReportDataRowView();
                                        //    RowsetMetaData rowsetMetaData = new RowsetMetaData();
                                        //    boRowsetController = subreportClientDocument.RowsetController;
                                        //    boRowsetCursor = boRowsetController.CreateCursor(gp1, rowsetMetaData, 0);
                                        //    (boRowsetCursor).RowsetController.RowsetBatchSize = 1000;
                                        //    (boRowsetCursor).Rowset.BatchSize = 1000;
                                        //    (boRowsetCursor).MoveLast();
                                        //    rows = boRowsetCursor.RecordCount;
                                        //    btnReportObjects.Text += "\nSubreport Record Count: " + rows.ToString();
                                        //    btnReportObjects.AppendText("\n");
                                        //}
                                        //else
                                        //    btnHasSavedData.Checked = false;

                                        btnReportObjects.AppendText("----------------------------------------------------------------------------------------------------------\n\n");
                                    }
                                    catch (Exception ex)
                                    {
                                        btnReportObjects.AppendText("Error in " + SecName + " : " + ex.Message + "\nMust log onto to get Subreport Row count even with saved data");
                                        ++Errorcnt;
                                        btnErrorCount.Text = Errorcnt.ToString();
                                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
                                    }

                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        btnReportObjects.AppendText("Error in " + SecName + " : " + ex.Message + "\nMust log onto to get Subreport Row count even with saved data");
                        ++Errorcnt;
                        btnErrorCount.Text = Errorcnt.ToString();
                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
                    }

                    #endregion Subreports
                    break;
                case "Main Report Data Sources":
                    #region MainDataSources
                    btnReportObjects.Text = "";
                    flcnt = 0;
                    foreach (String resultField in rptClientDoc.DatabaseController.GetServerNames())
                    {
                        textBox1 = "Server: " + resultField.ToString();
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\nDatabase: ");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();

                        CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                        CrystalDecisions.CrystalReports.Engine.Tables crTables;

                        CrystalDecisions.Shared.ConnectionInfo crConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();

                        crConnectioninfo.DatabaseName.ToString();

                        //set up the database and tables objects to refer to the current report
                        crDatabase = rpt.Database;
                        crTables = crDatabase.Tables;

                        try
                        {
                            foreach (CrystalDecisions.CrystalReports.Engine.Table crTable in crTables)
                            {
                                textBox1 = crTable.Name.ToString();
                                btnReportObjects.Text += textBox1;
                                // checks for table alias
                                if (crTable.Location != crTable.Name)
                                {
                                    btnReportObjects.AppendText("(" + crTable.Location.ToString() + ")");
                                }
                                btnReportObjects.AppendText("'End' \n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                            return;
                        }

                    }
                    #endregion MainDataSources
                    break;
                case "Main Report Data Links":
                    #region MainDataLinks
                    btnReportObjects.Text = "";
                    flcnt = 0;
                    foreach (String resultField in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        //CrystalDecisions.ReportAppServer.Controllers.SubreportClientDocument mysub;
                        //mysub.EnableOnDemand = true;
                        textBox1 = resultField.ToString();
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("Need to add code to get main report data source info 'End' \n");
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                    }
                    #endregion MainDataLinks
                    break;
                case "SubReport Data Sources":
                    #region SubDataSources
                    btnReportObjects.Text = "";
                    flcnt = 0;
                    Errorcnt = 0;

                    CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                    CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                    CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;

                    //set the crSections object to the current report's sections
                    CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;

                    //loop through all the sections to find all the report objects
                    foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                    {
                        crReportObjects = crSection.ReportObjects;
                        //loop through all the report objects to find all the subreports
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                        {
                            if (crReportObject.Kind == ReportObjectKind.SubreportObject)
                            {
                                CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                                CrystalDecisions.CrystalReports.Engine.Tables crTables;

                                CrystalDecisions.Shared.ConnectionInfo crConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();
                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;
                                string mysubname = crSubreportObject.SubreportName.ToString();

                                crSubreportObject.ToString();
                                //open the subreport object
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                //// This does not work - trying to export subreport
                                CrystalDecisions.CrystalReports.Engine.ReportDocument myNewSubObj = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
                                CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc2;

                                //// This works do not modify
                                //myNewSubObj = crSubreportDocument;
                                //rptClientDoc2 = myNewSubObj.ReportClientDocument;

                                //CrystalDecisions.Shared.DiskFileDestinationOptions diskOpts = CrystalDecisions.Shared.ExportOptions.CreateDiskFileDestinationOptions();
                                //CrystalDecisions.Shared.ExportOptions exportOpts = new CrystalDecisions.Shared.ExportOptions();

                                //CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions RASexportOpts = new CrystalDecisions.ReportAppServer.ReportDefModel.ExportOptions();
                                //CrystalDecisions.ReportAppServer.ReportDefModel.RPTExportFormatOptions RasPDFExpOpts = new RPTExportFormatOptions();
                                ////exportOpts.ExportFormatOptions = new RPTExportFormatOptions();

                                //// this gets the report name and sets the export name to be the same less the extension
                                //string MyRptName = myNewSubObj.Name.ToString();

                                //diskOpts.DiskFileName = @"C:\Reports\" + MyRptName;

                                //exportOpts.DestinationOptions = diskOpts;
                                //exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
                                //exportOpts.ExportFormatType = ExportFormatType.CrystalReport;

                                //RASexportOpts.ExportFormatType = CrReportExportFormatEnum.crReportExportFormatCrystalReports;
                                ////myNewSubObj.ExportToDisk(ExportFormatType.CrystalReport, @"C:\Reports\" + MyRptName);

                                //rptClientDoc2.PrintOutputController.ExportEx(RASexportOpts).Save(@"C:\Reports\" + MyRptName, true);

                                //try
                                //{
                                //    myNewSubObj.Export(exportOpts);
                                //}
                                //catch (Exception ex)
                                //{
                                //    btnSQLStatement.Text = ex.Message;
                                //    return;
                                //}
                                //// This does not work

                                try
                                {
                                    foreach (IConnectionInfo subinfo in crSubreportDocument.DataSourceConnections)
                                    {
                                        textBox1 = "Subreport Name: " + mysubname.ToString();
                                        btnReportObjects.Text += textBox1;
                                        btnReportObjects.AppendText("\nServer: ");
                                        btnReportObjects.AppendText(subinfo.ServerName.ToString());
                                        btnReportObjects.AppendText("\nDatabase: ");
                                        ++flcnt;
                                        btnCount.Text = flcnt.ToString();

                                        //set the database and tables objects to work with the subreport
                                        crDatabase = crSubreportDocument.Database;
                                        crTables = crDatabase.Tables;

                                        try
                                        {
                                            foreach (CrystalDecisions.CrystalReports.Engine.Table crTable in crTables)
                                            {
                                                textBox1 = crTable.Name.ToString();
                                                btnReportObjects.Text += textBox1;
                                                // checks for table alias
                                                if (crTable.Location != crTable.Name)
                                                {
                                                    btnReportObjects.AppendText("(" + crTable.Location.ToString() + ")");
                                                }
                                                btnReportObjects.AppendText("'End' \n");
                                                ++flcnt;
                                                btnCount.Text = flcnt.ToString();
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.Message);
                                            return;
                                        }
                                        btnReportObjects.AppendText("\n");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    btnReportObjects.AppendText("Error: " + ex.Message + "\n");
                                    ++Errorcnt;
                                    btnErrorCount.Text = Errorcnt.ToString();
                                }
                            }
                        }

                        //++flcnt;
                        btnCount.Text = flcnt.ToString();
                    }
                    #endregion SubDataSources
                    break;
                case "SubReport Links":
                    #region SubLinks
                    btnReportObjects.Text = "";
                    btnCount.Text = "";
                    foreach (String resultField in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        textBox1 = resultField.ToString();
                        btnReportObjects.Text += "Subreport Name: " + textBox1;
                        btnReportObjects.AppendText(":\n");

                        SubreportLinks SubLinks = rptClientDoc.SubreportController.GetSubreportLinks(resultField.ToString());
                        for (int I = 0; I < SubLinks.Count; I++)
                        {
                            SubreportLink subLink = SubLinks[I];
                            textBox1 = subLink.LinkedParameterName.ToString();
                            btnReportObjects.Text += "PM-Name: " + textBox1;
                            btnReportObjects.AppendText("\n");
                            textBox1 = subLink.MainReportFieldName.ToString();
                            btnReportObjects.Text += "Main Field: " + textBox1;
                            btnReportObjects.AppendText("\n");
                            textBox1 = subLink.SubreportFieldName.ToString();
                            btnReportObjects.Text += " Sub Field: " + textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                        }
                        btnReportObjects.AppendText("\n");
                        btnCount.Text = SubLinks.Count.ToString();
                    }
                    #endregion SubLinks
                    break;
                case "Special Fields":
                    btnReportObjects.Text = "";
                    getSpecialFields(rpt);
                    break;
                case "Hyperlinks":
                    btnReportObjects.Text = "";
                    getHyperlinks(rpt);
                    break;
                case "Section Print Orientation":
                    btnReportObjects.Text = "";
                    getSectionPrintOrientation(rpt);
                    break;
                case "Summary Info":
                    #region SummaryInfo
                    btnReportObjects.Text = "";
                    try
                    {
                        if (rpt.SummaryInfo.ReportAuthor != null)
                        {
                            textBox1 = "Author: " + rpt.SummaryInfo.ReportAuthor.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                            btnReportObjects.AppendText("Length: " + rpt.SummaryInfo.ReportAuthor.Length.ToString());
                            btnReportObjects.AppendText(" 'End' \n\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1 = "Author: ERROR: " + ex.Message;
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n"); 
                    }
                    try
                    {
                        if (rpt.SummaryInfo.ReportComments != null)
                        {
                            textBox1 = "Comments: " + rpt.SummaryInfo.ReportComments.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                            btnReportObjects.AppendText("Length: " + rpt.SummaryInfo.ReportComments.Length.ToString());
                            btnReportObjects.AppendText(" 'End' \n\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1 = "Comments: ERROR: " + ex.Message;
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                    }
                    try
                    {
                        //rpt.SummaryInfo.ReportTitle = "Vinits Report";
                        if (rpt.SummaryInfo.ReportTitle != null)
                        {
                            textBox1 = "Title: " + rpt.SummaryInfo.ReportTitle.ToString();
                            rpt.SummaryInfo.ReportTitle = "Overview Report" + "\r\n" + "Hello";
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                            btnReportObjects.AppendText("Length: " + rpt.SummaryInfo.ReportTitle.Length.ToString());
                            btnReportObjects.AppendText(" 'End' \n\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1 = "Title: ERROR: " + ex.Message;
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                    }

                    try // enhancement in SP 13
                    {
                        if (rpt.SummaryInfo.LastSavedBy != null)
                        {
                            textBox1 += "Last Saved By: " + rpt.SummaryInfo.LastSavedBy.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText("\n\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1 += "Last Saved By: ERROR: " + ex.Message;
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                    }

                    try // enhancement in SP 13
                    {
                        if (rpt.SummaryInfo.RevisionNumber != null)
                        {
                            textBox1 += "Revision Number: " + rpt.SummaryInfo.RevisionNumber.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText("\n\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        textBox1 = "Revision Number: ERROR: " + ex.Message;
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                    }

                    #endregion SummaryInfo
                    break;
                case "Alerts":
                    #region Alerts
                    btnReportObjects.Text = "";
                    Alerts alerts;
                    btnCount.Text = "";
                    alerts = rptClientDoc.SearchController.GetTriggeredAlerts();
                    if (alerts.Count > 0)
                    {
                        foreach (Alert alert in alerts)
                        {
                            textBox1 = "Alert Name:     " + alert.Name + "\n";
                            textBox1 += "Alert Message:  " + alert.Message.ToString() + "\n";
                            textBox1 += "Alert Formula:  " + alert.ConditionFilter.FreeEditingText.ToString() + "\n";
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                        }
                        btnCount.Text = alerts.Count.ToString();
                    }
                    #endregion Alerts
                    break;
                case "LOV - List of Values":
                    btnReportObjects.Text = "";
                    CrystalDecisions.ReportAppServer.Prompting.LOVField myLOV = new CrystalDecisions.ReportAppServer.Prompting.LOVField();
                    btnCount.Text = "";
                    myLOV.Name = rptClientDoc.DataDefController.ResultFieldController.ToString();
                    MessageBox.Show("Use Parameter menu option");
                    break;
                case "Graphic Location Formula":
                    #region GRaphicsLocal
                    btnReportObjects.Text = "";

                    CrystalDecisions.ReportAppServer.ReportDefModel.ConditionFormula myGraphicLocationFormula;
                    rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetReportObjectsByKind(CrReportObjectKindEnum.crReportObjectKindPicture);

                    // get the count of picture objects
                    btnCount.Text = rptObjs.Count.ToString();
                    foreach (CrystalDecisions.ReportAppServer.ReportDefModel.PictureObject MyrptObj in rptObjs)
                    {
                        if (((dynamic)MyrptObj).GraphicLocationFormula.Text != null)
                        {
                            textBox1 = "Picture Name: " + MyrptObj.Name.ToString() + "\n";
                            textBox1 += "Graphic Location Formula:     " + MyrptObj.GraphicLocationFormula.Text.ToString() + "\n";
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                        }
                        else
                        {
                            textBox1 = "Picture Name: " + MyrptObj.Name.ToString() + "\n";
                            textBox1 += "Graphic Location Formula: None\n";
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText(" 'End' \n");
                        }
                    }
                    #endregion GRaphicsLocal
                    break;
                case "Table Joins":
                    #region TableJoins
                    btnReportObjects.Text = "";
//                    string crJoinOperator = "";

                    //CrystalDecisions.ReportAppServer.DataDefModel.TableJoin RASTableJoins = new TableJoins();

                    //RASTableJoins = rptClientDoc.DataDefController.Database.TableLinks[0].JoinType;

                    //foreach (CrystalDecisions.ReportAppServer.DataDefModel.TableJoin RASTableJoins in rptClientDoc.DataDefController.Database.TableJoins)
                    //{
                    //    string crJoinOperator = "";

                    //}

                    //int x = 0;

                    string crJoinOperator = "";

                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.TableLink rasTableLink in rptClientDoc.DataDefController.Database.TableLinks)
                    {
                        //get the link properties
                        btnCount.Text = "";
                        int y = rptClientDoc.DataDefController.Database.TableLinks.Count;
                        btnCount.Text = y.ToString();
                        string crJoinType = "";

                        //for (int x = 0; x <= y; x++)
                        //{
                        //    try
                        //    {
                        //        if (rptClientDoc.DataDefController.Database.TableJoins[0].JoinOperator == CrTableJoinOperatorEnum.crTableJoinOperatorInnerJoin)
                        //        {
                        //            crJoinOperator = "InnerJoin";
                        //        }
                        //    }
                        //    catch
                        //    {
                        //        btnReportObjects.AppendText("Invalid Index 'End' \n");
                        //        //x++;
                        //    }

                        //}

                        //try
                        //{
                        //    if (rptClientDoc.DataDefController.Database.TableJoins[x].JoinOperator == CrTableJoinOperatorEnum.crTableJoinOperatorInnerJoin)
                        //    {
                        //        crJoinOperator = "InnerJoin ";
                        //    }
                        //}


                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeAdvance)
                        {
                            crJoinType = "-> Advanced ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeEqualJoin)
                        {
                            crJoinType = "-> = ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeGreaterOrEqualJoin)
                        {
                            crJoinType = "-> >= ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeGreaterThanJoin)
                        {
                            crJoinType = "-> > ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeLeftOuterJoin)
                        {
                            crJoinType = "-> LOJ ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeLessOrEqualJoin)
                        {
                            crJoinType = "-> <= ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeLessThanJoin)
                        {
                            crJoinType = "-> < ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeNotEqualJoin)
                        {
                            crJoinType = "-> != ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeOuterJoin)
                        {
                            crJoinType = "-> OJ ->";
                        }
                        if (rasTableLink.JoinType == CrTableJoinTypeEnum.crTableJoinTypeRightOuterJoin)
                        {
                            crJoinType = "-> ROJ ->";
                        }

                        textBox1 = "Only gets Link type:" + rasTableLink.SourceTableAlias.ToString() + "." + rasTableLink.SourceFieldNames[0].ToString() +
                            crJoinOperator + "." + crJoinType + rasTableLink.TargetTableAlias.ToString() + "." + rasTableLink.TargetFieldNames[0].ToString() + "\n";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText(" 'End' \n");
                    }
                    #endregion TableJoins
                    break;
                case "Charts":
                    #region Charts
                    btnReportObjects.Text = "";
                    flcnt = 0;
                    rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects();

                    foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj1 in rptObjs)
                    {
                        switch (rptObj1.Kind)
                        {
                            case CrReportObjectKindEnum.crReportObjectKindChart:
                                textBox1 = rptObj1.Name.ToString();
                                textBox1 += " -> ";
                                textBox1 += rptObj1.Name.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "Height: " + rptObj1.Height.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Width: " + rptObj1.Width.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Top: " + rptObj1.Top.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "Left: " + rptObj1.Left.ToString();
                                btnReportObjects.AppendText(" Twips\n");
                                btnReportObjects.Text += "BottomLineStyle: " + rptObj1.Border.BottomLineStyle.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "BackgroundColor: " + rptObj1.Border.BackgroundColor.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "BorderColor: " + rptObj1.Border.BorderColor.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "BottomLineStyle: " + rptObj1.Border.BottomLineStyle.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "EnableTightHorizontal: " + rptObj1.Border.EnableTightHorizontal.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "HasDropShadow: " + rptObj1.Border.HasDropShadow.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "LeftLineStyle: " + rptObj1.Border.LeftLineStyle.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "RightLineStyle: " + rptObj1.Border.RightLineStyle.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "TopLineStyle: " + rptObj1.Border.TopLineStyle.ToString();
                                btnReportObjects.AppendText("\n");
                                btnReportObjects.Text += "Section Name: " + rptObj1.SectionName.ToString();
                                btnReportObjects.AppendText("\n\n");
                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;
                        }
                    }
                    #endregion Charts
                    break;
                case "Condition Fields/Formula":
                    #region Conditional
                    {
                        btnReportObjects.Text = "";
                        flcnt = 0;

                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject reportObject in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
                        {
                            var reportSection = rptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(reportObject.SectionName);
                            for (int index = 0; index < reportSection.Format.ConditionFormulas.Count; ++index)
                            {
                                var formula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index];
                                var NewFormula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index];

                                try
                                {
                                    if (formula.Text != null)
                                    {
                                        textBox1 = "Condition Formula: " + formula.Text.ToString() + "\n";
                                        textBox1 += "Section " + reportSection.Name + "\n";
                                        btnReportObjects.Text += textBox1;
                                        btnReportObjects.AppendText(" 'End' \n");
                                        textBox1 = "";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    btnReportObjects.Text += "\nException: " + ex.ToString();
                                }
                            }
                        }
                    }
                    #endregion Conditional
                    break;
                case "Section Condition Formula":
                    #region Section Conditional
                    {
                        btnReportObjects.Text = "";
                        flcnt = 0;

                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Area CrArea in rptClientDoc.ReportDefinition.Areas)
                        {
                            try
                            {
                                int RawSectID = 1;
                                if (CrArea.Sections.Count > 1)
                                {
                                    //foreach (CrystalDecisions.CrystalReports.Engine.Section crSect in CrArea.Sections) // "DetailSection1"
                                    foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSect in CrArea.Sections)
                                    {
                                        //int sectionCodeArea = (crSect.SectionCode / 6000) % 6000; // Area
                                        //int sectionCodeSection = (crSect.SectionCode % 6000); // zero based Section
                                        //int sectionCodeGroup = (((crSect.SectionCode) / 50) % 120); // group section
                                        //int sectionCodeGroupNo = (((crSect.SectionCode) % 50)); // group area
                                        int sectionCodeArea = (crSect.SectionCode / 1000) % 1000; // Area
                                        int sectionCodeSection = (crSect.SectionCode % 1000); // zero based Section
                                        int sectionCodeGroup = (((crSect.SectionCode) / 25) % 40); // group section
                                        int sectionCodeGroupNo = (((crSect.SectionCode) % 25)); // group area
                                        bool isGRoup = false;

                                        SecName1 = CrArea.Kind.ToString();
                                        SecName1 = SecName1.Substring(17, (SecName1.Length - 17));

                                        crtoChr = 97;
                                        char b, c;

                                        if (sectionCodeGroup <= 26 && (SecName1 == "GroupFooter" || SecName1 == "GroupHeader"))
                                        {
                                            isGRoup = true;
                                        }
                                        else
                                        {  // aa and up
                                            if (RawSectID <= 26 && !isGRoup) // a to z
                                            {
                                                btnReportObjects.AppendText(SecName1);
                                                c = (char)(RawSectID + crtoChr - 1);
                                                btnReportObjects.Text += textBox1 + "  " + c + " : ";
                                                textBox1 = "";
                                            }
                                            else
                                            {
                                                btnReportObjects.AppendText(SecName1);
                                                if ((char)((RawSectID) % 26) == 0) // test if next "b" and set b back 1 and last letter to z
                                                {
                                                    b = (char)(((((RawSectID) / 26) % 26) - 2) + (crtoChr));
                                                    c = (char)(122);
                                                }
                                                else
                                                {   // this determines if the name is a or aa.
                                                    b = (char)((((RawSectID) / 26) % 26) + crtoChr - 1);
                                                    c = (char)((RawSectID % 26) + crtoChr - 1);
                                                }

                                                btnReportObjects.Text += textBox1 + " " + b + c + " : " + RawSectID + " - ";
                                            }
                                        }

                                        if (isGRoup)
                                        {
                                            if (sectionCodeGroup == 0) // && (SecName1 == "GroupFooter" || SecName1 == "GroupHeader")) // only one group area now sections below
                                            {
                                                btnReportObjects.AppendText(SecName1 + " #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr) + ": "); // + crSect.Format.PageOrientation.ToString() + "\n");
                                                foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject reportObject in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
                                                {
                                                    var reportSection = rptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(reportObject.SectionName);
                                                    if (reportSection.Format.ConditionFormulas.Count != 0)
                                                    {
                                                        for (int index = 0; index < reportSection.Format.ConditionFormulas.Count; ++index)
                                                        {
                                                            //if (reportObject.SectionCode == reportSection.SectionCode)
                                                            if (crSect.Name == reportSection.Name)
                                                            {
                                                                if (reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index].Text != null)
                                                                {
                                                                    var formula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index].Text.ToString();
                                                                    try
                                                                    {
                                                                        if (formula != null)
                                                                        {
                                                                            //btnReportObjects.AppendText(SecName1 + " : ");
                                                                            btnReportObjects.Text += "\nCondition Formula: \n";
                                                                            btnReportObjects.AppendText(formula.ToString() + "\n");
                                                                            textBox1 = "";
                                                                        }
                                                                    }
                                                                    catch (Exception ex)
                                                                    {
                                                                        btnReportObjects.Text += "\nException: " + ex.ToString();
                                                                    }
                                                                    isGRoup = false;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                btnReportObjects.AppendText("None\n");
                                                textBox1 = "";
                                                isGRoup = false;
                                            }
                                            else
                                            {
                                                btnReportObjects.AppendText(SecName1 + " #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + (crtoChr - 1))); // + ": " + crSect.Format.PageOrientation.ToString() + "\n");
                                                foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject reportObject in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
                                                {
                                                    var reportSection = rptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(reportObject.SectionName);
                                                    if (reportSection.Format.ConditionFormulas.Count != 0)
                                                    {
                                                        for (int index = 0; index < reportSection.Format.ConditionFormulas.Count; ++index)
                                                            if (crSect.Name == reportSection.Name)
                                                            {
                                                                var formula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index].Text.ToString();
                                                                try
                                                                {
                                                                    if (formula != null)
                                                                    {
                                                                        btnReportObjects.AppendText(SecName1 + " : ");
                                                                        btnReportObjects.AppendText("\nCondition Formula: \n" + formula.ToString() + "\n");
                                                                        textBox1 = "";
                                                                    }
                                                                }
                                                                catch (Exception ex)
                                                                {
                                                                    btnReportObjects.Text += "\nException: " + ex.ToString();
                                                                }
                                                            }
                                                            else
                                                            {
                                                                btnReportObjects.AppendText(SecName1 + " : ");
                                                                btnReportObjects.AppendText("None\n");
                                                                textBox1 = "";

                                                                ++flcnt;
                                                                btnCount.Text = flcnt.ToString();
                                                            }
                                                        btnReportObjects.AppendText(SecName1 + " : ");
                                                        btnReportObjects.AppendText("None\n");
                                                        textBox1 = "";
                                                    }
                                                    else
                                                    {   // end of if none found
                                                        btnReportObjects.AppendText(SecName1 + " : ");
                                                        btnReportObjects.AppendText("None\n");
                                                        textBox1 = "";
                                                    }

                                                    ++flcnt;
                                                    btnCount.Text = flcnt.ToString();
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            textBox1 += crSect.Format.PageOrientation.ToString() + "\n";
                                            isGRoup = false;
                                        }
                                        btnReportObjects.Text += textBox1;
                                        textBox1 = "";
                                        btnCount.Text = flcnt.ToString();
                                        ++flcnt;
                                        RawSectID++;
                                    }
                                }
                                else
                                {
                                    SecName = CrArea.Kind.ToString();

                                    if (SecName == "crAreaSectionKindGroupHeader" || SecName == "crAreaSectionKindGroupFooter")
                                    {
                                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSectName in CrArea.Sections)
                                        {
                                            //int sectionCodeArea = (crSectName.SectionCode / 6000) % 6000; // Area
                                            //int sectionCodeSection = (crSectName.SectionCode % 6000); // zero based Section
                                            //int sectionCodeGroup = (((crSectName.SectionCode) / 50) % 120); // group section
                                            //int sectionCodeGroupNo = (((crSectName.SectionCode) % 50)); // group area
                                            int sectionCodeArea = (crSectName.SectionCode / 1000) % 1000; // Area
                                            int sectionCodeSection = (crSectName.SectionCode % 1000); // zero based Section
                                            int sectionCodeGroup = (((crSectName.SectionCode) / 25) % 40); // group section
                                            int sectionCodeGroupNo = (((crSectName.SectionCode) % 40)); // group area
                                            crtoChr = 97;

                                            if (sectionCodeGroup <= 26)
                                            {
                                                if (SecName == "crAreaSectionKindGroupFooter")
                                                    btnReportObjects.AppendText("GroupFooter #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr));
                                                else
                                                {
                                                    btnReportObjects.AppendText("GroupHeader #" + (sectionCodeGroupNo + 1) + (char)(sectionCodeGroup + crtoChr));
                                                }
                                            }
                                            else
                                            { // Group higher than -aa
                                                char b, c;
                                                if ((char)((sectionCodeSection) % 26) == 0) // test if next "b" and set b back 1 and last letter to z
                                                {
                                                    b = (char)(((sectionCodeSection / 26) % 26) + (crtoChr));
                                                    c = (char)(122);

                                                    btnReportObjects.AppendText(c + "\n");
                                                }
                                                else
                                                {   // this determines if the name is a or aa.
                                                    b = (char)((((sectionCodeSection) / 26) % 26) + (crtoChr));
                                                    c = (char)((sectionCodeSection % 26) + crtoChr - 1);

                                                    btnReportObjects.AppendText(" " + b + c + "\n");
                                                }
                                            }

                                            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject reportObject in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
                                            {
                                                var reportSection = rptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(reportObject.SectionName);
                                                for (int index = 0; index < reportSection.Format.ConditionFormulas.Count; ++index)
                                                    if (crSectName.Name == reportSection.Name)
                                                    {
                                                        var formula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index].Text.ToString();
                                                        try
                                                        {
                                                            if (formula != null)
                                                            {
                                                                //btnReportObjects.AppendText(SecName1 + " : ");
                                                                btnReportObjects.AppendText("\nCondition Formula: \n" + formula.ToString() + "\n");
                                                                textBox1 = "";
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            btnReportObjects.Text += "\nException: " + ex.ToString();
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //btnReportObjects.AppendText(SecName1 + " : ");
                                                        btnReportObjects.AppendText("None\n");
                                                        textBox1 = "";

                                                        ++flcnt;
                                                        btnCount.Text = flcnt.ToString();
                                                    }
                                                break;
                                            }
                                            //if (CrArea.Name.Substring(0, 11) == crSectName.Name.Substring(0, 11))
                                            {
                                                btnReportObjects.AppendText(" - ");
                                                btnReportObjects.Text += textBox1;
                                                btnReportObjects.AppendText("None\n");
                                                textBox1 = "";

                                                ++flcnt;
                                                btnCount.Text = flcnt.ToString();
                                                break;
                                            }
                                        }
                                    }

                                    foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSectName in CrArea.Sections)
                                    {
                                        SecName1 = CrArea.Kind.ToString();
                                        SecName1 = SecName1.Substring(17, (SecName1.Length - 17));
                                        if (SecName1 == "GroupFooter" || SecName1 == "GroupHeader")
                                        {
                                            //btnReportObjects.AppendText(SecName1 + " #" + (((crSectName.SectionCode) % 40) + 1) + " : ");
                                        }
                                        else
                                        {
                                            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject reportObject in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
                                            {
                                                var reportSection = rptClientDoc.ReportDefController.ReportDefinition.FindSectionByName(reportObject.SectionName);
                                                if (reportSection.Format.ConditionFormulas.Count != 0)
                                                {
                                                    for (int index = 0; index < reportSection.Format.ConditionFormulas.Count; ++index)
                                                        if (crSectName.Name == reportSection.Name)
                                                        {
                                                            var formula = reportSection.Format.ConditionFormulas[(CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum)index].Text.ToString();
                                                            try
                                                            {
                                                                if (formula != null)
                                                                {
                                                                    btnReportObjects.AppendText(SecName1 + " : ");
                                                                    btnReportObjects.AppendText("\nCondition Formula: \n" + formula.ToString() + "\n");
                                                                    textBox1 = "";
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                btnReportObjects.Text += "\nException: " + ex.ToString();
                                                            }
                                                        }
                                                        else
                                                        {
                                                            btnReportObjects.AppendText(SecName1 + " : ");
                                                            btnReportObjects.AppendText("None\n");
                                                            textBox1 = "";

                                                            ++flcnt;
                                                            btnCount.Text = flcnt.ToString();
                                                        }
                                                }
                                                else
                                                {   // end of if none found
                                                    btnReportObjects.AppendText(SecName1 + " : ");
                                                    btnReportObjects.AppendText("None\n");
                                                    textBox1 = "";
                                                }

                                                ++flcnt;
                                                btnCount.Text = flcnt.ToString();
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            catch
                            {
                                btnReportObjects.AppendText("unhandled exception being thrown. 'End' \n");
                            }
                        }
                    }

                    #endregion Section Conditional
                    break;
                case "Fonts used in the report":
                    #region Fonts
                    {
                        btnReportObjects.Text = "";
                        flcnt = 0;

                        //CrystalDecisions.ReportAppServer.ReportDefModel.ReportObjects rptObjs;
                        //rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetReportObjectsByKind(CrReportObjectKindEnum.crReportObjectKindField);
                        rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects();


                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj1 in rptObjs)
                        {
                            switch (rptObj1.Kind)
                            {
                                case CrReportObjectKindEnum.crReportObjectKindField:
                                    CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject fldObj1;
                                    fldObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)rptObj1;

                                    textBox1 = fldObj1.Name.ToString();
                                    textBox1 += " -> ";

                                    rptObj = (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject)rptClientDoc.ReportDefController.ReportDefinition.FindObjectByName(fldObj1.Name.ToString());


                                    CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject objField = rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects()[rptObj.Name] as CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject;
                                        //objField.FontColor.Font.Name = "Tahoma";

                                    textBox1 += objField.FontColor.Font.Name.ToString();
                                    btnReportObjects.Text += textBox1;
                                    btnReportObjects.AppendText("\n");
                                    ++flcnt;
                                    btnCount.Text = flcnt.ToString();

                                
                                // this always returns an empty string. will be tracked
                                    CrystalDecisions.ReportAppServer.ReportDefModel.ConditionFormula displayStringCondition;
                                    displayStringCondition = rptObj.Format.ConditionFormulas[CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeDisplayString];

                                    break;
                                case CrReportObjectKindEnum.crReportObjectKindText:
                                    CrystalDecisions.ReportAppServer.ReportDefModel.TextObject txtObj;
                                    txtObj = (CrystalDecisions.ReportAppServer.ReportDefModel.TextObject)rptObj1;

                                    textBox1 = txtObj.Name.ToString();
                                    textBox1 += " -> ";
                                    textBox1 += txtObj.Text.ToString();
                                    textBox1 += "\nFont: " + txtObj.FontColor.Font.Name.ToString();
                                    //if txtObj.FontColor.Font.
                                    btnReportObjects.Text += textBox1;
                                    btnReportObjects.AppendText("\n");
                                    ++flcnt;
                                    btnCount.Text = flcnt.ToString();

                                    break;
                                case CrReportObjectKindEnum.crReportObjectKindBox:
                                    CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject boxObj1;
                                    CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject boxObj2 = new CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject();
                                    boxObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.BoxObject)rptObj1;
                                    boxObj2 = boxObj1;

                                    //// modify the box size
                                    //boxObj.Bottom = 360;
                                    //boxObj1.Height = 2385;
                                    //boxObj1.Format.EnableCanGrow = true;
                                    //boxObj1.EndSectionName = "ReportFooterSection1";
                                    //boxObj1.EnableExtendToBottomOfSection = true;
                                    //rptClientDoc.ReportDefController.ReportObjectController.Modify(boxObj, boxObj1);

                                    textBox1 = boxObj1.Name.ToString();
                                    textBox1 += " -> ";
                                    textBox1 += boxObj1.Name.ToString();
                                    textBox1 += " : X - ";
                                    textBox1 += (boxObj1.Left / 1440.00).ToString();
                                    textBox1 += " : Y - ";
                                    textBox1 += ((boxObj1.Right - boxObj1.Left) / 1440.00).ToString();
                                    textBox1 += " : Width - ";
                                    textBox1 += (boxObj1.Width / 1440.00).ToString();
                                    textBox1 += " : Height - ";
                                    textBox1 += (boxObj1.Height / 1440.00).ToString();
                                    btnReportObjects.Text += textBox1;

                                    btnReportObjects.AppendText("\n");
                                    ++flcnt;
                                    btnCount.Text = flcnt.ToString();
                                    IsRpt = false;

                                    break;
                                case CrReportObjectKindEnum.crReportObjectKindPicture:

                                    CrystalDecisions.ReportAppServer.ReportDefModel.PictureObject MyGraphicLocationFormula; // = new PictureObject();

                                    //txtObj = MyGraphicLocationFormula.GraphicLocationFormula.Text.ToString();

                                    //txtObj = (CrystalDecisions.ReportAppServer.ReportDefModel.TextObject)rptObj1;

                                    //textBox1 = txtObj.Name.ToString();
                                    //textBox1 += " -> ";
                                    //textBox1 += txtObj.Text.ToString();
                                    //btnReportObjects.Text += textBox1;
                                    //btnReportObjects.AppendText("\n");
                                    //++flcnt;
                                    //btnCount.Text = flcnt.ToString();
                                    break;

                                  //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindSubreport;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindPicture;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindOlapGrid;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindMap;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindLine;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindInvalid;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindFlash;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindFieldHeading;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindField;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindCrosstab;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindChart;
                                //CrystalDecisions.ReportAppServer.ReportDefModel.CrReportObjectKindEnum.crReportObjectKindBlobField;

                            }
                        }
                    }
                    #endregion Fonts
                    break;
                default:
                    btnReportObjects.Text = "";
                    //btnReportObjects.Text = "None Used"; 
                    break;
            }
        }

        int iCnt = -1;
        bool YorN;
        bool IsParamRange;

        private void getParameterFields(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {
            string textBox1 = "";
            string textBox2 = "";
            string MyObjectType = ReportObjectComboBox1.SelectedItem.ToString();
            btnReportObjects.Text = "";

            iCnt = -1;

            // NOTE: WHEN GETTING THE PARAMETER LISTS IT DOES NOT MATTER THE TYPE OF PARAMETER. ONLY WHEN SETTING THE VALUES DOES IT MATTER. 
            // SEE SetParam_Click, ReportDocumentSetParameters AND SetCrystalParam FOR MORE DETAILS ON HOW TO

            if (rptClientDoc.DataDefController.DataDefinition.ParameterFields.Count > 0) //there are parameters
            {
                btnCount.Text = rptClientDoc.DataDefController.DataDefinition.ParameterFields.Count.ToString();
                foreach (CrystalDecisions.ReportAppServer.DataDefModel.ParameterField paramfield in rptClientDoc.DataDefController.DataDefinition.ParameterFields)
                {
                    iCnt++;
                    // this line gets the parameter name by index
                    //rptClientDoc.DataDefController.DataDefinition.ParameterFields[1].LongName.ToString();
                    switch (paramfield.ValueRangeKind)
                    {
                        case CrParameterValueRangeKindEnum.crParameterValueRangeKindDiscrete:
                            {
                                try
                                {
                                    if (paramfield.ReportName.ToString() != null)
                                        btnReportObjects.Text += "Subreport Name: " + paramfield.ReportName.ToString() + "\n";
                                }
                                catch (Exception ex)
                                {
                                    { } // main report does not have a name so ignore the exception
                                }
                                btnReportObjects.Text += "Discrete Param Name: \"";
                                btnReportObjects.AppendText(paramfield.Name.ToString() + "\"");
                                textBox1 = "";
                                getDiscreteValues(paramfield);
                                break;
                            }
                        case CrParameterValueRangeKindEnum.crParameterValueRangeKindDiscreteAndRange:
                            {
                                try
                                {
                                    if (paramfield.ReportName.ToString() != null)
                                        btnReportObjects.Text += "Subreport Name: " + paramfield.ReportName.ToString() + "\n";
                                }
                                catch (Exception ex)
                                {
                                    { } // main report does not have a name so ignore the exception
                                }
                                //getRangeAndDiscreteValues(paramfield);
                                btnReportObjects.Text += "Discrete and Range Param Name: ";
                                btnReportObjects.AppendText(paramfield.Name.ToString() + "\"");
                                textBox1 = "";
                                getDiscreteValues(paramfield);
                                IsParamRange = true;
                                break;
                            }
                        case CrParameterValueRangeKindEnum.crParameterValueRangeKindRange:
                            {
                                try
                                {
                                    if (paramfield.ReportName.ToString() != null)
                                        btnReportObjects.Text += "Subreport Name: " + paramfield.ReportName.ToString() + "\n";
                                }
                                catch (Exception ex)
                                {
                                    { } // main report does not have a name so ignore the exception
                                }

                                btnReportObjects.Text += "Range Parameter Name: ";
                                btnReportObjects.AppendText(paramfield.Name.ToString() + "\"");
                                textBox1 = "";
                                //getRangeValues(paramfield);
                                getDiscreteValues(paramfield);
                                IsParamRange = true;
                                break;
                            }
                    }
                }
            }

            //rpt.SetParameterValue("pcname", Convert.ToDateTime(parmPCname));
            //rpt.SetParameterValue("pcname", Convert.ToDecimal(parmPCname));
            //rpt.SetParameterValue("pcname", Convert.ToString(parmPCname));

            #region Commented
            //int x = 0;
            //int z = 0;
            //foreach (CrystalDecisions.ReportAppServer.DataDefModel.ParameterField paramFields in rptClientDoc.DataDefController.DataDefinition.ParameterFields)
            //{
            //    textBox1 = paramFields.LongName.ToString();
            //    btnReportObjects.Text += textBox1;
            //    btnReportObjects.AppendText(" = ");

            //    CrystalDecisions.ReportAppServer.DataDefModel.Fields flds;
            //    flds = rptClientDoc.DataDefController.ParameterFieldController.GetPromptParameterFields(null);

            //    CrystalDecisions.ReportAppServer.DataDefModel.ParameterField pfld1 = (CrystalDecisions.ReportAppServer.DataDefModel.ParameterField)flds[z];
            //    int currValCount = pfld1.InitialValues.Count;
            //    int defValCount = pfld1.DefaultValues.Count;

            //    try
            //    {
            //        CrystalDecisions.ReportAppServer.DataDefModel.Value val = (CrystalDecisions.ReportAppServer.DataDefModel.Value)pfld1.InitialValues[z];
            //        btnReportObjects.AppendText(pfld1.InitialValues.ToString());
            //        btnReportObjects.AppendText(" is Initial Value \n");
            //    }
            //    catch
            //    {
            //        btnReportObjects.AppendText("No Initial Value \n");
            //    }

            //    if (paramFields.ValueRangeKind == CrParameterValueRangeKindEnum.crParameterValueRangeKindDiscrete)
            //    {
            //        try
            //        {
            //            z = 0;
            //            //if (currValCount > 0)  //this parameter has default values
            //            {
            //                foreach (ParameterFieldDiscreteValue discreteDefaultVal in pfld1.InitialValues)
            //                {
            //                    btnReportObjects.AppendText(discreteDefaultVal.Value.ToString());
            //                    btnReportObjects.AppendText(": is Default Value \n");
            //                    pfld1.InitialValues.RemoveAll();
            //                    pfld1.InitialValues.Add("Donw");
            //                    rpt.VerifyDatabase();
            //                    z++;
            //                }
            //            }
            //            //else
            //            //{
            //            //    //CrystalDecisions.ReportAppServer.DataDefModel.ParameterFieldDiscreteValue paramField1;
            //            //    //paramField1 = (CrystalDecisions.ReportAppServer.DataDefModel.ParameterFieldDiscreteValue)paramFields.DefaultValues[0];
            //            //    btnReportObjects.AppendText(pfld1.InitialValues.ToString());
            //            //    btnReportObjects.AppendText(": is Default Value \n");
            //            //    z++;
            //            //}
            //        }
            //        catch
            //        {
            //            btnReportObjects.AppendText("No Default Values \n");
            //        }
            //    }
            //    else
            //    {
            //        int y = 0;
            //        CrystalDecisions.ReportAppServer.DataDefModel.ParameterField pfld = (CrystalDecisions.ReportAppServer.DataDefModel.ParameterField)rptClientDoc.DataDefinition.ParameterFields[x];

            //        int pCount = pfld.DefaultValues.Count;
            //        CrystalDecisions.ReportAppServer.DataDefModel.ISCRValues currentvalues;
            //        currentvalues = pfld.CurrentValues;
            //        flds = rptClientDoc.DataDefController.ParameterFieldController.GetPromptParameterFields(null);

            //        btnReportObjects.AppendText("Range Values:\n ");

            //        foreach (CrystalDecisions.ReportAppServer.DataDefModel.ISCRValue RangeVal in pfld.DefaultValues)
            //        {
            //            CrystalDecisions.Shared.ParameterDiscreteValue oValue = (CrystalDecisions.Shared.ParameterDiscreteValue)rpt.DataDefinition.ParameterFields[x].DefaultValues[y];
            //            btnReportObjects.AppendText(oValue.Value.ToString());
            //            // Get the original default values
            //            ParameterFieldDefinition boParameterFieldDefinition;
            //            ParameterValues boParameterDefaultValues;
            //            //ParameterDiscreteValue boParameterDiscreteValue;
            //            bool myDefault;
            //            myDefault = rpt.DataDefinition.ParameterFields[x].HasCurrentValue;
            //            boParameterFieldDefinition = rpt.DataDefinition.ParameterFields[0];
            //            boParameterDefaultValues = rpt.DataDefinition.ParameterFields[0].CurrentValues;
            //            //CrystalDecisions.ReportAppServer.DataDefModel.ISCRParameterFieldRangeValue.EndValue

            //            btnReportObjects.AppendText("\n ");
            //            y++;
            //        }
            //    }

            //    btnReportObjects.AppendText(" 'End' \n");
            //    x++;
            //}
            #endregion Commented
        }

        public Boolean isParameterDynamic(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt, int iCnt, bool YorN)
        {
            if (rpt.DataDefinition.ParameterFields[iCnt].Attributes != null && rpt.DataDefinition.ParameterFields[iCnt].Attributes.ContainsKey("IsDCP"))
            {
                Hashtable objAttributes = rpt.DataDefinition.ParameterFields[iCnt].Attributes;
                YorN = (Boolean)objAttributes["IsDCP"];
                return YorN;
            }
            return YorN;
        }

        private void getDiscreteValues(CrystalDecisions.ReportAppServer.DataDefModel.ISCRParameterField paramfield)
        {
            string textBox1 = "";

            int currValCount = paramfield.InitialValues.Count; // this is the old Parameter value count. Seen in CRD as "Default"
            int defValCount = paramfield.DefaultValues.Count;
            string LOVFieldName = "";
            bool currValCountValue = false;

            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeQueryParameter")
                btnReportObjects.AppendText(" (Query Parameter)");

            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeConnectionParameter")
                btnReportObjects.AppendText(" (Connection Parameter)");

            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeMetaDataParameter")
                btnReportObjects.AppendText(" (MetaData Parameter)");

            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeReportParameter")
                btnReportObjects.AppendText(" (Crystal");

            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeStoredProcedureParameter")
            {
                if (IsCMD)
                    btnReportObjects.AppendText(" (Command)");
                else
                    if (IsBEX == false)
                        btnReportObjects.AppendText(" (Stored Procedure)");
                    else
                        btnReportObjects.AppendText(" (BEX Query)");
            }

            try
            { // provided by DEV to tell if the param is dynamic or not
                if (rpt.DataDefinition.ParameterFields[iCnt].Attributes != null)
                {
                    Hashtable objAttributes = rpt.DataDefinition.ParameterFields[iCnt].Attributes;
                    if ((Boolean)objAttributes["IsDCP"] != false)
                    {
                        try
                        {
                            LOVFieldName = objAttributes["FieldID"].ToString();
                        }
                        catch (Exception ex)
                        {
                            //btnReportObjects.AppendText("\nParameter not based on any field");
                        }

                        YorN = (Boolean)objAttributes["IsDCP"];
                        if ((dynamic)paramfield.Usage == 25 || (dynamic)paramfield.Usage == 57)
                            btnReportObjects.AppendText(" Dynamic Cascading Parameter)");
                        else
                            btnReportObjects.AppendText(" Dynamic Parameter)");

                    }
                    else
                        if (defValCount > 0)
                            btnReportObjects.AppendText(" Static LOV Parameter)");
                        else
                            btnReportObjects.AppendText(" Static Parameter)");
                }
            }
            catch (Exception ex)
            {
                btnReportObjects.AppendText(" Error in Parameter)");
            }

            if (paramfield.Type != CrystalDecisions.ReportAppServer.DataDefModel.CrFieldValueTypeEnum.crFieldValueTypeBooleanField)
            {
                if (currValCount == 0)
                    btnReportObjects.AppendText("\n Allow Custom Values: True\n"); // Always true if no current values specified
                else
                    btnReportObjects.AppendText(" Allow Multiple Values: " + paramfield.AllowCustomCurrentValues.ToString() + "\n");
                btnReportObjects.AppendText(" Allow Multiple Values: " + paramfield.AllowMultiValue.ToString() + "\n");
                btnReportObjects.AppendText(" Allow Discrete Values: " + paramfield.AllowCustomCurrentValues.ToString() + "\n");
                btnReportObjects.AppendText(" Allow Null Values: " + paramfield.AllowNullValue.ToString());
            }

            if (currValCount > 0)
            {
                foreach (ParameterFieldDiscreteValue CurrentInitialVal in paramfield.InitialValues)
                {
                    // Always only going to be 1
                    // Initial Value is the legacy Default Value, lower part of the dialog box.
                    textBox1 = "\n Current Initial Value (legacy): " + CurrentInitialVal.Value.ToString();
                    btnReportObjects.Text += textBox1;
                    btnReportObjects.AppendText("");
                    textBox1 = "";
                    currValCountValue = true;
                    if (defValCount == 0)
                    {
                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
                        return;
                    }
                }
            }

            if (defValCount > -1)  //this parameter has default values
            #region DefCount1
            {
                int xxx = 1;

                foreach (ParameterFieldDiscreteValue CurrentInitialVal in paramfield.DefaultValues)
                {
                    if (defValCount > 0)
                    {
                        foreach (ParameterFieldDiscreteValue discreteDefaultVal in paramfield.DefaultValues)
                        {
                            // I think I have to be logged on first
                            if (xxx == 1 && paramfield.BrowseField != null)
                            {
                                textBox1 += "\nBased on Field: {" + paramfield.BrowseField.LongName.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("}");
                            }
                            else
                            {
                                if (xxx == 1)
                                    btnReportObjects.AppendText("\nBased on Field: Parameter not based on any field");
                            }

                            textBox1 = "\n Default Value #" + xxx + ": \"" + discreteDefaultVal.Value.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText("\"");
                            textBox1 = "";

                            if (discreteDefaultVal.Description != null)
                            {
                                textBox1 += " - Description: \"" + discreteDefaultVal.Description.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("");
                                textBox1 = "";
                            }
                            else
                            {
                                textBox1 += " - Description: (blank)";
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("");
                                textBox1 = "";
                            }
                            ++xxx;
                            // do this because CR will create a list for each value as an Initial value so you get count * values
                            if (xxx >= defValCount + 1)
                            {
                                btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
                                return;
                            }
                        }
                    }
                }

                // this gets the field name the Dynamic param is based on
                if (paramfield.Attributes.PropertyIDs[0].ToString() != null)
                {
                    // provided by DEV to tell if the param is dynamic or not
                    try
                    {
                        if (rpt.DataDefinition.ParameterFields[iCnt].Attributes != null)
                        {
                            Hashtable objAttributes = rpt.DataDefinition.ParameterFields[iCnt].Attributes;
                            if ((Boolean)objAttributes["IsDCP"] != false)
                            {
                                try
                                {
                                    LOVFieldName = objAttributes["FieldID"].ToString();
                                }
                                catch (Exception ex)
                                {
                                    //btnReportObjects.AppendText("\nParameter not based on any field");
                                }
                                YorN = (Boolean)objAttributes["IsDCP"];
                            }
                            else
                                LOVFieldName = "Parameter not based on any field";
                        }
                    }
                    catch (Exception ex)
                    {
                        btnReportObjects.AppendText(" Error in Parameter)");
                    }

                    //paramfield.Attributes.PropertyIDs[0].ToString();
                    textBox1 += "\nBased on Field: " + LOVFieldName.ToString();
                    btnReportObjects.Text += textBox1;

                    if (!YorN && rpt.HasSavedData != true)
                    {
                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n");
                        textBox1 = "";
                    }
                }
                else
                {
                    textBox1 += " - Description: (blank)";
                    btnReportObjects.Text += textBox1;
                    btnReportObjects.AppendText("");
                    textBox1 = "";
                }

                // if SP parameter requires being logged on first. To be updated, rpt file saves the values from the last time previewed, if saved of course.
                if (YorN || rpt.HasSavedData == true)
                {   // if report based on BV then count = 0
                    if (defValCount != 0)
                    {
                        foreach (ParameterFieldDiscreteValue CurrentInitialVal in paramfield.CurrentValues)
                        {
                            textBox1 = "\n Saved Data Parameter Value #" + xxx + ": \"" + CurrentInitialVal.Value.ToString();
                            btnReportObjects.Text += textBox1;
                            btnReportObjects.AppendText("\"");
                            textBox1 = "";

                            if (CurrentInitialVal.Description != null)
                            {
                                textBox1 += " - Description: \"" + CurrentInitialVal.Description.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("");
                                textBox1 = "";
                            }
                            ++xxx;
                        }
                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n");
                        textBox1 = "";
                    }
                    else
                    {
                        textBox1 = "\n No Default or Current Value saved in the Report";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\"");
                        textBox1 = "";
                        btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n");
                        textBox1 = "";
                    }
                }
                //btnReportObjects.AppendText("\n");
                ++xxx;
                YorN = false;
                // do this because CR will create a list for each value as an Initial value so you get count * values
                if (xxx >= defValCount + 1)
                {
                    btnReportObjects.AppendText("\n");
                    return;
                }
            #endregion DefCount1
            }

            /*            else
                        {
                            // this else never jumped into
                            //isParameterDynamic(rpt, iCnt, YorN);
                            btnReportObjects.AppendText(paramfield.Name.ToString());

                            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeStoredProcedureParameter")
                                btnReportObjects.AppendText(": Store Procedure Parameter \nBased on field: " + paramfield.FormulaForm + "----------------------------------------------------------------------------------------------------------\n\n");

                            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeReportParameter")
                            {
                                // provided by DEV to tell if the param is dynamic or not
                                if (rpt.DataDefinition.ParameterFields[iCnt].Attributes != null && rpt.DataDefinition.ParameterFields[iCnt].Attributes.ContainsKey("IsDCP"))
                                {
                                    Hashtable objAttributes = rpt.DataDefinition.ParameterFields[iCnt].Attributes;
                                    LOVFieldName = objAttributes["FieldID"].ToString();
                                    YorN = (Boolean)objAttributes["IsDCP"];
                                }
                                if (YorN)
                                {
                                    if ((dynamic)paramfield.Usage == 25)
                                        btnReportObjects.AppendText(": Dynamic Cascading Parameter ( LOV )\nBased on field: " + paramfield.BrowseField.ToString() + "----------------------------------------------------------------------------------------------------------\n\n"); //.Remove(0, 6) + "\n"); // paramfield.FormulaForm + "\n");
                                    else
                                        btnReportObjects.AppendText(": Dynamic Parameter \nBased on field: " + LOVFieldName + "----------------------------------------------------------------------------------------------------------\n\n"); //.Remove(0, 6) + "\n"); // paramfield.FormulaForm + "\n");
                                    //btnReportObjects.AppendText(": No Initial Value \n");
                                }
                                else
                                {
                                    btnReportObjects.AppendText(": Crystal Parameter \nBased on field: " + paramfield.FormulaForm + "----------------------------------------------------------------------------------------------------------\n\n");
                                    if ((dynamic)(paramfield.Values != null))
                                        btnReportObjects.AppendText("Current Value: " + ((dynamic)((paramfield.Values)).ToString()) + "----------------------------------------------------------------------------------------------------------\n\n"); // paramfield.InitialValues.Value.ToString());
                                }
                            }

                            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeQueryParameter")
                                btnReportObjects.AppendText(": Query Parameter \nBased on field: " + paramfield.FormulaForm + "----------------------------------------------------------------------------------------------------------\n\n");

                            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeConnectionParameter")
                                btnReportObjects.AppendText(": Connection Parameter \nBased on field: " + paramfield.FormulaForm + "----------------------------------------------------------------------------------------------------------\n\n");

                            if (paramfield.ParameterType.ToString() == "crParameterFieldTypeMetaDataParameter")
                                btnReportObjects.AppendText(": MetaData Parameter \nBased on field: " + paramfield.FormulaForm + "----------------------------------------------------------------------------------------------------------\n\n");
                        }
             * */
            btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
        }

        private void getSpecialFields(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {
            string textBox1 = String.Empty;
            string textBox2 = String.Empty;
            string MyObjectType = ReportObjectComboBox1.SelectedItem.ToString();
            btnReportObjects.Text = "";
            int flcnt = 0;

            CrystalDecisions.ReportAppServer.ReportDefModel.ReportObjects rptObjs;

            rptObjs = rptClientDoc.ReportDefController.ReportObjectController.GetReportObjectsByKind(CrReportObjectKindEnum.crReportObjectKindField);
            //t = rptClientDoc.DataDefController.DataDefinition.FormulaFields.FindIndexOf(rptObjs);
            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj in rptObjs)
            {
                CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject fldObj;
                fldObj = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)rptObj;
                string MySpecialType = fldObj.DataSourceName.ToString();

                switch (MySpecialType)
                {
                    case "ContentLocale":
                        textBox1 = "Content Locale";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "CurrentUserID":
                        textBox1 = "Current CE User ID";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "CurrentUserName":
                        textBox1 = "Current CE User Name";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "CurrentTimeZone":
                        textBox1 = "Current CE User Time Zone";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "DataDate":
                        textBox1 = "Data Date";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "DataTime":
                        textBox1 = "Data Time";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "DataTimeZone":
                        textBox1 = "Data Time Zone";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "FileAuthor":
                        textBox1 = "File Author";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "FileCreationDate":
                        textBox1 = "File Creation Date";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "Filename":
                        textBox1 = "File Path and Name";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "GroupNumber":
                        textBox1 = "Group Number";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "GroupSelection":
                        textBox1 = "Group Selection";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "HPageNumber":
                        textBox1 = "Horizontal Page Number";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "ModificationDate":
                        textBox1 = "Modification Date";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "ModificationTime":
                        textBox1 = "Modification Time";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "PageNofM":
                        textBox1 = "Page N of M";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "PageNumber":
                        textBox1 = "Page Number";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "PrintDate":
                        textBox1 = "Print Date";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "PrintTime":
                        textBox1 = "Print Time";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "PrintTimeZone":
                        textBox1 = "Print Time Zone";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "RecordNumber":
                        textBox1 = "Record Number";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "RecordSelection":
                        textBox1 = "Record Selection Formula";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "ReportComments":
                        textBox1 = "Report Comments";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "ReportTitle":
                        textBox1 = "Report Title";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "SelectionLocale":
                        textBox1 = "Selection Locale";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;
                    case "TotalPageCount":
                        textBox1 = "Total Page Count";
                        btnReportObjects.Text += textBox1;
                        btnReportObjects.AppendText("\n");
                        break;

                    default:

                        break;
                        ++flcnt;
                        btnCount.Text = flcnt.ToString();
                }
            }
        }

        private void getSectionPrintOrientation(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {
            string textBox1 = String.Empty;
            string textBox2 = String.Empty;
            string MyObjectType = ReportObjectComboBox1.SelectedItem.ToString();
            btnReportObjects.Text = "";
            int flcnt = 0;
            string SecName = "";
            string SecName1 = "";
            int sectionCodeArea = 0; // Area
            int sectionCodeSection = 0; // zero based Section
            int sectionCodeGroup = 0; // group section
            int sectionCodeGroupNo = 0; // group area
            int x = 0;

            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Area CrArea in rptClientDoc.ReportDefinition.Areas)
            {
                try
                {
                    int RawSectID = 1;
                    if (CrArea.Sections.Count > 1)
                    {
                        //foreach (CrystalDecisions.CrystalReports.Engine.Section crSect in CrArea.Sections) // "DetailSection1"
                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSect in CrArea.Sections)
                        {
                            //int sectionCodeArea = (crSect.SectionCode / 6000) % 6000; // Area
                            //int sectionCodeSection = (crSect.SectionCode % 6000); // zero based Section
                            //int sectionCodeGroup = (((crSect.SectionCode) / 50) % 120); // group section
                            //int sectionCodeGroupNo = (((crSect.SectionCode) % 50)); // group area
                            sectionCodeArea = (crSect.SectionCode / 1000) % 1000; // Area
                            sectionCodeSection = (crSect.SectionCode % 1000); // zero based Section
                            sectionCodeGroup = (((crSect.SectionCode) / 25) % 40); // group section
                            sectionCodeGroupNo = (((crSect.SectionCode) % 25)); // group area
                            bool isGRoup = false;

                            SecName1 = CrArea.Kind.ToString();
                            SecName1 = SecName1.Substring(17, (SecName1.Length - 17));

                            //// this deletes ReportFooters
                            //if (CrArea.Kind.ToString() == "crAreaSectionKindReportFooter" && sectionCodeSection == 1)
                            //{
                            //    rptClientDoc.ReportDefController.ReportSectionController.Remove(crSect);
                            //}

                            int crtoChr = 97;
                            char b, c;

                            if (SecName1 == "GroupFooter" || SecName1 == "GroupHeader")
                            {
                                isGRoup = true;
                            }

                            if (RawSectID <= 26 && !isGRoup) // a to z
                            {
                                //btnReportObjects.Text += "\n" + ++x + "\n"; 
                                btnReportObjects.AppendText(SecName1);
                                c = (char)(RawSectID + crtoChr - 1);
                                // this is group A, B, C, etc.
                                if (!isGRoup)
                                    btnReportObjects.AppendText(" " + c + " : " + RawSectID + " - ");
                            }
                            else
                            {
                                btnReportObjects.AppendText(SecName1);
                                if ((char)((RawSectID) % 26) == 0) // test if next "b" and set b back 1 and last letter to z
                                {
                                    b = (char)(((((RawSectID) / 26) % 26) - 2) + (crtoChr));
                                    c = (char)(122);
                                    if (b == 96)
                                        btnReportObjects.AppendText(" #" + (sectionCodeGroupNo + 1) + c + " : " + RawSectID + " - ");
                                    else
                                        btnReportObjects.AppendText(" #" + (sectionCodeGroupNo + 1) + b + c + " : " + RawSectID + " - ");
                                }
                                else
                                {   // this determines if the name is a or aa.
                                    b = (char)((((RawSectID) / 26) % 26) + crtoChr);
                                    c = (char)((RawSectID % 26) + crtoChr - 1);
                                    if (b == 97)
                                        btnReportObjects.AppendText(" #" + (sectionCodeGroupNo + 1) + c + " : " + RawSectID + " - ");
                                    else
                                    {
                                        b = (char)((((RawSectID) / 26) % 26) + crtoChr - 1);
                                        if (SecName1 == "GroupFooter" || SecName1 == "GroupHeader")
                                            //if (SecName1 != "Detail")
                                            btnReportObjects.AppendText(" #" + (sectionCodeGroupNo + 1) + b + c + " : " + RawSectID + " - ");
                                        else
                                            btnReportObjects.AppendText(" " + b + (char)c + " : " + RawSectID + " - ");
                                    }
                                }
                                // group aa etc.   need to fix this, shows 2b and should be 2a

                            }

                            textBox1 += crSect.Format.PageOrientation.ToString() + "\n";
                            isGRoup = false;
                            btnReportObjects.Text += textBox1;
                            textBox1 = "";
                            btnCount.Text = flcnt.ToString();
                            ++flcnt;
                            RawSectID++;

                            //// want to keep this, it gets subreport objects by section
                            ////int objCnt = crSect.ReportObjects.Count;
                            //foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject subobj in crSect.ReportObjects)
                            //{
                            //    switch (subobj.ClassName)
                            //    {
                            //        // look for sub report object and display info
                            //        case "CrystalReports.SubreportObject": // CrReportObjectKindEnum.crReportObjectKindSubreport:
                            //            CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject subObj1;
                            //            try
                            //            {
                            //                subObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject)subobj;
                            //                btnReportObjects.Text += ": SubName: ";
                            //                btnReportObjects.AppendText(subObj1.SubreportName.ToString());
                            //                btnReportObjects.Text += " : ";
                            //                ++flcnt;

                            //                SubreportController subreportController = rptClientDoc.SubreportController;
                            //                SubreportClientDocument subreportClientDocument = subreportController.GetSubreport(subObj1.SubreportName);
                            //            }
                            //            catch (Exception ex)
                            //            {
                            //                btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
                            //            }
                            //            break;
                            //    }
                            //}
                        }
                    }
                    else
                    {
                        SecName = CrArea.Kind.ToString();

                        if (SecName == "crAreaSectionKindGroupHeader" || SecName == "crAreaSectionKindGroupFooter")
                        {
                            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSectName in CrArea.Sections)
                            {
                                //int sectionCodeArea = (crSectName.SectionCode / 6000) % 6000; // Area
                                //int sectionCodeSection = (crSectName.SectionCode % 6000); // zero based Section
                                //int sectionCodeGroup = (((crSectName.SectionCode) / 50) % 120); // group section
                                //int sectionCodeGroupNo = (((crSectName.SectionCode) % 50)); // group area
                                sectionCodeArea = (crSectName.SectionCode / 1000) % 1000; // Area
                                sectionCodeSection = (crSectName.SectionCode % 1000); // zero based Section
                                sectionCodeGroup = (((crSectName.SectionCode) / 25) % 40); // group section
                                sectionCodeGroupNo = (((crSectName.SectionCode) % 40)); // group area

                                int crtoChr = 97;

                                if (sectionCodeGroup <= 26)
                                {
                                    if (SecName == "crAreaSectionKindGroupFooter")
                                        btnReportObjects.AppendText("GroupFooter #" + (sectionCodeGroupNo + 1)); // + (char)(sectionCodeGroup + crtoChr));
                                    else
                                    {
                                        btnReportObjects.AppendText("GroupHeader #" + (sectionCodeGroupNo + 1)); // + (char)(sectionCodeGroup + crtoChr));
                                    }
                                }
                                else
                                { // Group higher than -aa
                                    char b, c;
                                    if ((char)((sectionCodeSection) % 26) == 0) // test if next "b" and set b back 1 and last letter to z
                                    {
                                        b = (char)(((sectionCodeSection / 26) % 26) + (crtoChr));
                                        c = (char)(122);

                                        btnReportObjects.AppendText(" " + c + "\n");
                                    }
                                    else
                                    {   // this determines if the name is a or aa.
                                        b = (char)((((sectionCodeSection) / 26) % 26) + (crtoChr));
                                        c = (char)((sectionCodeSection % 26) + crtoChr - 1);

                                        btnReportObjects.AppendText(" " + b + c + "\n");
                                    }
                                }

                                btnReportObjects.AppendText(" - ");
                                textBox1 += crSectName.Format.PageOrientation.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("\n");
                                textBox1 = "";

                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                                break;
                            }
                        }

                        foreach (CrystalDecisions.ReportAppServer.ReportDefModel.Section crSectName in CrArea.Sections)
                        {
                            SecName1 = CrArea.Kind.ToString();
                            SecName1 = SecName1.Substring(17, (SecName1.Length - 17));
                            if (SecName1 == "GroupFooter" || SecName1 == "GroupHeader")
                            {
                                //btnReportObjects.AppendText(SecName1 + " #" + (((crSectName.SectionCode) % 40) + 1) + " : ");
                            }
                            else
                            {   // single section print
                                crSectName.Format.ConditionFormulas.ToString();
                                btnReportObjects.AppendText(SecName1 + " : ");
                                textBox1 += crSectName.Format.PageOrientation.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText("\n");
                                textBox1 = "";

                                ++flcnt;
                                btnCount.Text = flcnt.ToString();
                            }
                            break;
                        }
                    }
                }
                catch
                {
                    btnReportObjects.AppendText("More than one Section in Area. 'End' \n");
                }
            }
            // want to keep this, it gets subreport objects by section
            //int objCnt = crSect.ReportObjects.Count;
            //foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject subobj in crSect.ReportObjects)
            //{
            //    switch (subobj.ClassName)
            //    {
            //        // look for sub report object and display info
            //        case "CrystalReports.SubreportObject": // CrReportObjectKindEnum.crReportObjectKindSubreport:
            //            CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject subObj1;
            //            try
            //            {
            //                subObj1 = (CrystalDecisions.ReportAppServer.ReportDefModel.SubreportObject)subobj;
            //                btnReportObjects.Text += ": SubName: ";
            //                btnReportObjects.AppendText(subObj1.SubreportName.ToString());
            //                btnReportObjects.Text += " : ";
            //                ++flcnt;

            //                SubreportController subreportController = rptClientDoc.SubreportController;
            //                SubreportClientDocument subreportClientDocument = subreportController.GetSubreport(subObj1.SubreportName);
            //            }
            //            catch (Exception ex)
            //            {
            //                btnReportObjects.AppendText("\n----------------------------------------------------------------------------------------------------------\n\n");
            //            }
            //            break;
            //    }
            //}
        }
        private void getHyperlinks(CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {
            string textBox1 = String.Empty;
            string MyObjectType = ReportObjectComboBox1.SelectedItem.ToString();

            btnReportObjects.Text = "";
            foreach (CrystalDecisions.ReportAppServer.ReportDefModel.ReportObject rptObj in rptClientDoc.ReportDefController.ReportObjectController.GetAllReportObjects())
            {
                CrystalDecisions.ReportAppServer.ReportDefModel.ObjectFormat objfmt;
                objfmt = rptObj.Format;
                switch (objfmt.HyperlinkType)
                {
                    case CrHyperlinkTypeEnum.crHyperlinkTypeCrystalReport:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeDrilldown:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": Report Part Drill Down - DHTML only".ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeEmail:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeEmailFieldValue:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                // MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeHtml:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeReportObject:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": Another Report Object - DHTML only".ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeUndefined:
                        {
                            //try
                            //{
                            //    textBox1 = rptObj.Name.ToString();
                            //    textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                            //    btnReportObjects.Text += textBox1;
                            //    btnReportObjects.AppendText(" 'End' \n");
                            //}
                            //catch (Exception ex)
                            //{
                            //    //MessageBox.Show("ERROR: " + ex.Message);
                            //}
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeWebsite:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                    case CrHyperlinkTypeEnum.crHyperlinkTypeWebsiteFieldValue:
                        {
                            try
                            {
                                textBox1 = rptObj.Name.ToString();
                                textBox1 += ": " + rptObj.Format.HyperlinkText.ToString();
                                btnReportObjects.Text += textBox1;
                                btnReportObjects.AppendText(" 'End' \n");
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("ERROR: " + ex.Message);
                            }
                            break;
                        }
                }
            }
        
        }

        static string aa = "";

        public static CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo CreateConnectionInfo()
        {
            PropertyBag pBag = new PropertyBag();
            PropertyBag crConnProperties = new PropertyBag();
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo crConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();

            // Set the userID and password values.
            crConnInfo.UserName = "sa";
            crConnInfo.Password = "PW";
            crConnInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;

            pBag.set_StringValue("Database DLL", "crdb_odbc.dll");
            pBag.set_StringValue("QE_DatabaseName", "xtreme");
            pBag.set_StringValue("QE_DatabaseType", "ODBC (RDO)");
            pBag.set_StringValue("QE_ServerDescription", "DSN Name");

            crConnProperties = new PropertyBag();

            crConnProperties.set_StringValue("Trusted Connection", "false");
            crConnProperties.set_StringValue("Server", "DSN Name");

            pBag.Add("QE_LogonProperties", crConnProperties);

            crConnInfo.Attributes = pBag;

            return crConnInfo;
        }

        private void GetFormula_Click(object sender, EventArgs e)
        {
            string myFF = "True and ";

            myFF += rpt.RecordSelectionFormula;

            rpt.RecordSelectionFormula = myFF.ToString();

            MessageBox.Show("my formula: " + myFF);

            //CrystalDecisions.ReportAppServer.DataDefModel.FormulaField ffc = new FormulaField(); //   CrystalDecisions.ReportAppServer.Controllers.FormulaFieldController();
            //CrystalDecisions.ReportAppServer.DataDefModel.FormulaField ffNew = new CrystalDecisions.ReportAppServer.DataDefModel.FormulaField();
            //CrystalDecisions.ReportAppServer.DataDefModel.FormulaField ffOld = new CrystalDecisions.ReportAppServer.DataDefModel.FormulaField();

            //try
            //{
            //    //Get the old field, this is the field defined in the report.
            //    ffOld = rptClientDoc.DataDefController.DataDefinition.FormulaFields.FindField(myFF, CrystalDecisions.ReportAppServer.DataDefModel.CrFieldDisplayNameTypeEnum.crFieldDisplayNameName);
            //    //'For Each ffOld In arpt.DataDefController.DataDefinition.FormulaFields
            //    //'    If ffOld.Name = asParameterName Then
            //    //'        Exit For
            //    //'    End If
            //    //'Next
            //    //If Not ffOld Is Nothing Then
            //    //Clone a new field and update values
            //    ffNew = ffOld.Clone();
            //    ffNew.Text = asParameterValue;
            //    ffc = rptClientDoc.DataDefController.FormulaFieldController;
            //    //Replace old with new
            //    ffc.Modify(ffOld, ffNew);

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("ERROR: " + ex.Message);
            //}            

        }

        private void SetSchema_Click(object sender, EventArgs e)
        {

            //rptClientDoc.DatabaseController.SetTableLocationByServerDatabaseName("Orders", "dwcb12003", "xtreme", "sb", "1Oem2000");
            rptClientDoc.DatabaseController.LogonEx("ServerNameOrIP", "xtreme", "sa", "PW");

            //Create the logon propertybag for the connection we wish to use
            CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag logonDetails = new CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag();
            logonDetails.Add("Auto Translate", -1);
            logonDetails.Add("Connect Timeout", 15);
            logonDetails.Add("Data Source", "ServerNameOrIP");
            logonDetails.Add("General Timeout", 0);
            logonDetails.Add("Initial Catalog", "Orders");
            logonDetails.Add("Integrated Security", "True");
            logonDetails.Add("Locale Identifier", 1033);
            logonDetails.Add("OLE DB Services", -5);
            logonDetails.Add("Provider", "SQLOLEDB");
            logonDetails.Add("Use Encryption for Data", 0);
            logonDetails.Add("Owner", "dbo"); // schema

            //Create the QE (query engine) propertybag with the provider details and logon property bag.
            CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag QE_Details = new CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag();
            QE_Details.Add("Database DLL", "crdb_ado.dll");
            QE_Details.Add("QE_DatabaseName", "Orders");
            QE_Details.Add("QE_DatabaseType", "OLE DB (ADO)");
            QE_Details.Add("QE_LogonProperties", logonDetails);
            QE_Details.Add("QE_ServerDescription", "ServerNameOrIP");
            QE_Details.Add("QE_SQLDB", "True");
            QE_Details.Add("SSO Enabled", "False");
            QE_Details.Add("Owner", "dbo");

            //CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rcd = reportDocument.ReportClientDocument;
            CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rcd = rptClientDoc;

            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;

            oldConnInfos = rcd.DatabaseController.GetConnectionInfos(null);
            for (int I = 0; I < oldConnInfos.Count; I++)
            {
                oldConnInfo = oldConnInfos[I];
                newConnInfo.Attributes = QE_Details;
                newConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                rcd.DatabaseController.ReplaceConnection(oldConnInfo, newConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
            }

        }

        private void DataSet1_Click(object sender, EventArgs e)
        {
            rptClientDoc = rpt.ReportClientDocument;

            string connString = "Provider=SQLNCLI10;Server=ServerNameOrIP;Database=xtreme;User ID=sa;Password=PW";

            System.Data.DataSet thisDataSet = new System.Data.DataSet();

            string sqlString = @"SELECT top 10*  FROM  ""xtreme"".""dbo"".""Financials"" ""Financials""";

            System.Data.OleDb.OleDbConnection oleConn = new System.Data.OleDb.OleDbConnection(connString);
            System.Data.OleDb.OleDbCommand cmd = oleConn.CreateCommand();
            cmd.CommandText = sqlString;

            System.Data.DataSet ds = new System.Data.DataSet();

            OleDbDataAdapter oleAdapter = new OleDbDataAdapter(sqlString, oleConn);
            //OleDbDataAdapter oleAdapter2 = new OleDbDataAdapter(sqlString2, oleConn);
            DataTable dt1 = new DataTable("Financials");
            //DataTable dt2 = new DataTable("Orders Detail");

            oleAdapter.Fill(dt1);
            //oleAdapter2.Fill(dt2);

            ds.Tables.Add(dt1);
            //ds.Tables.Add(dt2);
            ds.WriteXml("c:\\reports\\sc2.xml", XmlWriteMode.WriteSchema);

            // as long as the field names match exactly Cr has no problems setting report to a DS.
            try
            {
                rpt.SetDataSource(ds.Tables[0]);
                rpt.SetDataSource(ds);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: Schema Mismatch. Error reported by CR: " + ex.Message);
            }

            //Now check for subreport and set to same DS
            foreach (CrystalDecisions.CrystalReports.Engine.Section section in rpt.ReportDefinition.Sections)
            {
                foreach (CrystalDecisions.CrystalReports.Engine.ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        CrystalDecisions.CrystalReports.Engine.SubreportObject subReport = (CrystalDecisions.CrystalReports.Engine.SubreportObject)reportObject;

                        CrystalDecisions.CrystalReports.Engine.ReportDocument subDocument = subReport.OpenSubreport(subReport.SubreportName);

                        subDocument.SetDataSource(ds);
                    }
                }
            }
            MessageBox.Show("Data Source Set", "RAS", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ReplaceConnection_Click(object sender, EventArgs e)
        {
            DateTime dtStart;
            TimeSpan difference;
            DateTime TSTotal;
            TSTotal = DateTime.Now;

            CrystalDecisions.Shared.ConnectionInfo crConnectioninfo = new CrystalDecisions.Shared.ConnectionInfo();

            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;

            //Create a new Database Table to replace the reports current table.
            CrystalDecisions.ReportAppServer.DataDefModel.Table boTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
            CrystalDecisions.ReportAppServer.DataDefModel.Table subboTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();

            //boMainPropertyBag: These hold the attributes of the tables ConnectionInfo object
            PropertyBag boMainPropertyBag = new PropertyBag();
            //boInnerPropertyBag: These hold the attributes for the QE_LogonProperties
            //In the main property bag (boMainPropertyBag)
            PropertyBag boInnerPropertyBag = new PropertyBag();

            //Create a new ConnectionInfo object
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
            CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag PromptProperties = new PropertyBag();
            int dbConCount = rptClientDoc.DatabaseController.GetConnectionInfos(PromptProperties).Count;
            string DBDriver = rptClientDoc.DatabaseController.GetConnectionInfos(PromptProperties)[0].Attributes.get_StringValue("Database DLL").ToString(); // = "crdb_xxx.dll"; // 

            if (DBDriver.ToString() == "crdb_odbc.dll")
            {
                //Set the attributes for the boInnerPropertyBag
                boInnerPropertyBag.Add("Database", "xtreme");
                boInnerPropertyBag.Add("DSN", "10.161.121.28 NT auth");
                boInnerPropertyBag.Add("Trusted_Connection", "1");
                boInnerPropertyBag.Add("UseDSNProperties", "False");

                //Set the attributes for the boMainPropertyBag
                boMainPropertyBag.Add("Database DLL", "crdb_odbc.dll");
                boMainPropertyBag.Add("QE_DatabaseName", "xtreme");
                boMainPropertyBag.Add("QE_DatabaseType", "ODBC (RDO)");
                //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                boMainPropertyBag.Add("QE_ServerDescription", "10.161.121.28 NT auth");
                boMainPropertyBag.Add("QE_SQLDB", "True");
                boMainPropertyBag.Add("SSO Enabled", "False");

                //Create a new ConnectionInfo object
                //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                //Pass the database properties to a connection info object
                boConnectionInfo.Attributes = boMainPropertyBag;
                //Set the connection kind
                boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                //**EDIT** Set the User Name and Password if required.
                //boConnectionInfo.UserName = "";
                //boConnectionInfo.Password = "";
                //Pass the connection information to the table
                boTable.ConnectionInfo = boConnectionInfo;
            }

            if (DBDriver.ToString() == "crdb_oracle.dll")
            {
                //Set the attributes for the boInnerPropertyBag
                //boInnerPropertyBag.Add("Data Source", btrDataFile.Text);
                boInnerPropertyBag.Add("PreQEServerName", btrDataFile.Text);
                boInnerPropertyBag.Add("PreQEServerType", "Oracle Server");

                //Set the attributes for the boMainPropertyBag
                boMainPropertyBag.Add("Database DLL", "crdb_oracle.dll");
                boMainPropertyBag.Add("QE_DatabaseName", "");
                boMainPropertyBag.Add("QE_DatabaseType", "Oracle Server");
                //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                //boMainPropertyBag.Add("QE_ServerDescription", "fun11rac");
                boMainPropertyBag.Add("QE_SQLDB", "True");
                boMainPropertyBag.Add("SSO Enabled", "False");

                //Create a new ConnectionInfo object
                //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                //Pass the database properties to a connection info object
                boConnectionInfo.Attributes = boMainPropertyBag;
                //Set the connection kind
                boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                //**EDIT** Set the User Name and Password if required.
                boConnectionInfo.UserName = btrFileLocation.Text;
                boConnectionInfo.Password = btrPassword.Text;
                //Pass the connection information to the table
                boTable.ConnectionInfo = boConnectionInfo;

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                //For each table in the report:
                // - Set the Table Name properties.
                // - Set the table location in the report to use the new modified table

                for (int S = 0; S < boTables.Count; S++)
                {
                    boTable.Name = (dynamic)boTables[S].Name;
                    boTable.QualifiedName = (dynamic)boTables[S].QualifiedName;
                    boTable.Alias = (dynamic)boTables[S].Alias;

                    try
                    {
                        dtStart = DateTime.Now;
                        rptClientDoc.DatabaseController.SetTableLocation(boTables[0], boTable);
                        difference = DateTime.Now.Subtract(dtStart);
                        btnReportObjects.Text += boTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message); // + " ;" + ex.InnerException.Message);
                        return;
                    }
                }

                //Verify the database after adding substituting the new table.
                //To ensure that the table updates properly when adding Command tables or Stored Procedures.
                rptClientDoc.VerifyDatabase();
                IsRpt = false;
            }

            if (DBDriver.ToString() == "crdb_adoplus.dll")
            {
                //Set the attributes for the boInnerPropertyBag
                boInnerPropertyBag.Add("File Path ", @"D:\Atest\321145\PretEmpr.XML");
                boInnerPropertyBag.Add("Internal Connection ID", "{2c6ae642-8c24-4fe9-a4ad-f95f78f142b8}");

                //Set the attributes for the boMainPropertyBag
                boMainPropertyBag.Add("Database DLL", "crdb_adoplus.dll");
                boMainPropertyBag.Add("QE_DatabaseName", "");
                boMainPropertyBag.Add("QE_DatabaseType", "ADO.NET (XML)");
                //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                boMainPropertyBag.Add("QE_ServerDescription", "NewDataSet");
                boMainPropertyBag.Add("QE_SQLDB", "False");
                boMainPropertyBag.Add("SSO Enabled", "False");

                //Pass the database properties to a connection info object
                boConnectionInfo.Attributes = boMainPropertyBag;
                //Set the connection kind
                boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                //**EDIT** Set the User Name and Password if required.
                //boConnectionInfo.UserName = "UserName";
                //boConnectionInfo.Password = "Password";
                //Pass the connection information to the table
                boTable.ConnectionInfo = boConnectionInfo;

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                //For each table in the report:
                // - Set the Table Name properties.
                // - Set the table location in the report to use the new modified table
                boTable.Name = "Table";
                boTable.QualifiedName = "Table";
                boTable.Alias = "PretEmpr";

                rptClientDoc.DatabaseController.SetTableLocation(boTables[0], boTable);

                //Verify the database after adding substituting the new table.
                //To ensure that the table updates properly when adding Command tables or Stored Procedures.
                rptClientDoc.VerifyDatabase();
                try
                {
                    rptClientDoc.DatabaseController.SetTableLocation(boTables[0], boTable);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR: " + ex.Message);
                    return;
                }
            }

            if (DBDriver.ToString() == "crdb_p2bxbse.dll")
            {
                //Set the attributes for the boInnerPropertyBag
                boInnerPropertyBag.Add("Database Name", btrDataFile.Text);
                boInnerPropertyBag.Add("Database Type", "dBASE 5.0");

                //Set the attributes for the boMainPropertyBag
                boMainPropertyBag.Add("Database DLL", "crdb_dao.dll");
                boMainPropertyBag.Add("QE_DatabaseName", btrDataFile.Text);
                boMainPropertyBag.Add("QE_DatabaseType", "Access/Excel (DAO)");
                //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                boMainPropertyBag.Add("QE_ServerDescription", btrDataFile.Text);
                boMainPropertyBag.Add("QE_SQLDB", "True");
                boMainPropertyBag.Add("SSO Enabled", "False");

                //Pass the database properties to a connection info object
                boConnectionInfo.Attributes = boMainPropertyBag;
                //Set the connection kind
                boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                //**EDIT** Set the User Name and Password if required.
                //boConnectionInfo.UserName = "UserName";
                //boConnectionInfo.Password = "Password";
                //Pass the connection information to the table
                boTable.ConnectionInfo = boConnectionInfo;

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                //For each table in the report:
                // - Set the Table Name properties.
                // - Set the table location in the report to use the new modified table
                boTable.Name = btrSearchPath.Text;
                boTable.QualifiedName = btrSearchPath.Text;
                boTable.Alias = btrSearchPath.Text;

                rptClientDoc.DatabaseController.SetTableLocation(boTables[0], boTable);

                //Verify the database after adding substituting the new table.
                //To ensure that the table updates properly when adding Command tables or Stored Procedures.
                rptClientDoc.VerifyDatabase();
                IsRpt = false;
            }

            if (DBDriver.ToString() == "crdb_ado.dll")
            {
                //Set the attributes for the boInnerPropertyBag
                boInnerPropertyBag.Add("Data Source", btrDataFile.Text);
                boInnerPropertyBag.Add("Locale Identifier", "1033");
                boInnerPropertyBag.Add("OLE DB Services", "-5");
                boInnerPropertyBag.Add("Provider", "OraOLEDB.Oracle");
                boInnerPropertyBag.Add("Use DSN Default Properties", "False");

                //Set the attributes for the boMainPropertyBag
                boMainPropertyBag.Add("Database DLL", "crdb_ado.dll");
                boMainPropertyBag.Add("QE_DatabaseName", "");
                boMainPropertyBag.Add("QE_DatabaseType", "OLE DB (ADO)");
                //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                //boMainPropertyBag.Add("QE_ServerDescription", "fun11rac");
                boMainPropertyBag.Add("QE_SQLDB", "True");
                boMainPropertyBag.Add("SSO Enabled", "False");

                //Create a new ConnectionInfo object
                //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                //Pass the database properties to a connection info object
                boConnectionInfo.Attributes = boMainPropertyBag;
                //Set the connection kind
                boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                //**EDIT** Set the User Name and Password if required.
                boConnectionInfo.UserName = btrFileLocation.Text;
                boConnectionInfo.Password = btrPassword.Text;
                //Pass the connection information to the table
                boTable.ConnectionInfo = boConnectionInfo;

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                //For each table in the report:
                // - Set the Table Name properties.
                // - Set the table location in the report to use the new modified table

                for (int S = 0; S < boTables.Count; S++)
                {
                    boTable.Name = (dynamic)boTables[S].Name;
                    boTable.QualifiedName = (dynamic)boTables[S].QualifiedName;
                    boTable.Alias = (dynamic)boTables[S].Alias;

                    try
                    {
                        dtStart = DateTime.Now;
                        rptClientDoc.DatabaseController.SetTableLocation(boTables[0], boTable);
                        difference = DateTime.Now.Subtract(dtStart);
                        btnReportObjects.Text += boTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message); // + " ;" + ex.InnerException.Message);
                        return;
                    }
                }

                //Verify the database after adding substituting the new table.
                //To ensure that the table updates properly when adding Command tables or Stored Procedures.
                rptClientDoc.VerifyDatabase();
                IsRpt = false;
            }

            // Now look for each subreport and do the same
            #region Subreport
            foreach (string crSubreportDocument1 in rptClientDoc.SubreportController.GetSubreportNames())
            {
                SubreportClientDocument SubRCD = rptClientDoc.SubreportController.GetSubreport(crSubreportDocument1);
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newSubConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables subboTables;
                subboTables = SubRCD.DatabaseController.Database.Tables;

                //**EDIT** Set the User Name and Password if required.
                //boConnectionInfo.UserName = "Admin";
                //boConnectionInfo.Password = "holly";
                //Pass the connection information to the table
                subboTable.ConnectionInfo = boConnectionInfo;

                btnReportObjects.Text += "\nSubreport: " + crSubreportDocument1.ToString() + "\n";

                for (int S = 0; S < subboTables.Count; S++)
                {
                    subboTable.Name = (dynamic)subboTables[S].Name;
                    subboTable.QualifiedName = (dynamic)subboTables[S].QualifiedName;
                    subboTable.Alias = (dynamic)subboTables[S].Alias;

                    try
                    {
                        dtStart = DateTime.Now;
                        SubRCD.DatabaseController.SetTableLocation(subboTables[S], subboTable);
                        difference = DateTime.Now.Subtract(dtStart);
                        btnReportObjects.Text += subboTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message); // + " ;" + ex.InnerException.Message);
                        return;
                    }


                    // Subreport links
                    foreach (String resultField in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        SubreportLinks OldSubLinks = rptClientDoc.SubreportController.GetSubreportLinks(resultField.ToString());
                        SubreportLinks NewSubLinks = OldSubLinks.Clone(true);

                        for (int I = 0; I < OldSubLinks.Count; I++)
                        {
                            SubreportLink subLink = NewSubLinks[I];
                        }
                    }
                }
            }
            #endregion Subreport

            GroupPath gp = new GroupPath();
            string tmp = String.Empty;
            try
            {
                rptClientDoc.RowsetController.GetSQLStatement(gp, out tmp);
                btnSQLStatement.Text = tmp;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex.Message);
                return;
            }

            IsRpt = false;

            //CrystalDecisions.CrystalReports.Engine.Subreports mySubs = new Subreports();

            //mySubs. Subreports
        }

        private void SetToXML_Click(object sender, EventArgs e)
        {
            DateTime dtStart;
            TimeSpan difference;
            DateTime TSTotal;
            TSTotal = DateTime.Now;

            // Access
            rptClientDoc = rpt.ReportClientDocument;

            //Create a new Database Table to replace the reports current table.
            CrystalDecisions.ReportAppServer.DataDefModel.Table boTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
            CrystalDecisions.ReportAppServer.DataDefModel.Table subboTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();

            //Get the Database Tables Collection for your report
            CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
            boTables = rptClientDoc.DatabaseController.Database.Tables;

            // Get the old connection info
            oldConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
            boTable.ConnectionInfo = boConnectionInfo;

            oldConnInfo = oldConnInfos[0];

            // RAS XML - need to use ReplaceConnection
            # region XML
            #region RASXMLReplaceconnection;
            if (chkUseRAS.CheckState == CheckState.Checked)
            {
                rptClientDoc = rpt.ReportClientDocument;

                //Get the Database Tables Collection for your report
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                // Get the old connection info
                oldConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                boTable.ConnectionInfo = boConnectionInfo;

                oldConnInfo = oldConnInfos[0];

                //if ((dynamic)oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_adoplus.dll")
                {
                    PropertyBag logonDetails = new PropertyBag();
                    PropertyBag QeDetails = new PropertyBag();

                    //boMainPropertyBag: These hold the attributes of the tables ConnectionInfo object
                    PropertyBag boMainPropertyBag = new PropertyBag();
                    //boInnerPropertyBag: These hold the attributes for the QE_LogonProperties
                    //In the main property bag (boMainPropertyBag)
                    PropertyBag boInnerPropertyBag = new PropertyBag();

                    //Set the attributes for the boInnerPropertyBag
                    boInnerPropertyBag.Add("File Path ", btrDataFile.Text + btrSearchPath.Text);
                    boInnerPropertyBag.Add("Server Path", btrSearchPath.Text);
                    boInnerPropertyBag.Add("Database Name", btrSearchPath.Text); // btrDataFile.Text
                    boInnerPropertyBag.Add("Database Type", "ADO.NET (XML)");

                    //Set the attributes for the boMainPropertyBag
                    boMainPropertyBag.Add("Database DLL", "crdb_adoplus.dll");
                    boMainPropertyBag.Add("QE_DatabaseName", btrDataFile.Text);
                    boMainPropertyBag.Add("QE_DatabaseType", "ADO.NET (XML)"); 
                    //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                    boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                    boMainPropertyBag.Add("QE_ServerDescription", btrSearchPath.Text);
                    boMainPropertyBag.Add("QE_SQLDB", "false");
                    boMainPropertyBag.Add("SSO Enabled", "False");

                    // ReplaceConnection method This works also
                    for (int I = 0; I < oldConnInfos.Count; I++)
                    {
                        // CrystalDecisions.ReportAppServer.DataDefModel.Table tbl;
                        for (int S = 0; S < boTables.Count; S++)
                        {

                            oldConnInfo = oldConnInfos[I];
                            newConnInfo.Attributes = boMainPropertyBag;
                            newConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            try
                            {
                                rptClientDoc.DatabaseController.ReplaceConnection(oldConnInfo, newConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message); //  + " ;" + ex.InnerException.Message.ToString());
                            }
                        }
                    }
                    // ReplaceConnection method This works also

                    // check for subreports
                    //loop through all the report objects to find all the subreports
                    foreach (string crSubreportDocument1 in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        SubreportClientDocument SubRCD = rptClientDoc.SubreportController.GetSubreport(crSubreportDocument1);
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newSubConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldSubConnInfo;
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldSubConnInfos;

                        PropertyBag SublogonDetails = new PropertyBag();
                        PropertyBag SubQeDetails = new PropertyBag();

                        oldSubConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                        btnReportObjects.Text += "\nSubreport: " + crSubreportDocument1.ToString() + "\n";

                        for (int I = 0; I < oldSubConnInfos.Count; I++)
                        {
                            oldSubConnInfo = oldSubConnInfos[I];
                            newSubConnInfo.Attributes = boMainPropertyBag;
                            newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            // this works also
                            for (int S = 0; S < boTables.Count; S++)
                            {
                                oldSubConnInfo = oldSubConnInfos[I];
                                newSubConnInfo.Attributes = boMainPropertyBag;
                                newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                                try
                                {
                                    rptClientDoc.DatabaseController.ReplaceConnection(oldSubConnInfo, newSubConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("ERROR: " + ex.Message);
                                }
                            }
                        }
                        IsRpt = false;
                    }
                }
            }

            #endregion RASXMLReplaceconnection;

            // now set the PC data to the xml file
            if (chkUseRAS.CheckState != CheckState.Checked)
            {
                if (oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_adoplus.dll")
                {
                    // Engine
                    CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                    CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                    CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
                    CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                    CrystalDecisions.CrystalReports.Engine.Tables crTables;

                    CrystalDecisions.Shared.TableLogOnInfo tLogonInfo;

                    btnSQLStatement.Text = "";

                    try
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in rpt.Database.Tables)
                        {
                            tLogonInfo = rptTable.LogOnInfo;
                            tLogonInfo.ConnectionInfo.ServerName = btrDataFile.Text; 
                            tLogonInfo.ConnectionInfo.DatabaseName = btrSearchPath.Text; 
                            tLogonInfo.ConnectionInfo.UserID = "";
                            tLogonInfo.ConnectionInfo.Password = "";
                            tLogonInfo.TableName = rptTable.Name;

                            dtStart = DateTime.Now;

                            try
                            {
                                rptTable.ApplyLogOnInfo(tLogonInfo);
                                rptTable.Location = btrDataFile.Text + "\\" + btrSearchPath.Text;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message);
                                //return;
                            }

                            difference = DateTime.Now.Subtract(dtStart);

                            btnSQLStatement.Text += /*rptTable.Name.ToString() +*/ " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";

                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                        //return;
                    }

                    #region XML Subreport
                    // check for subreports
                    //set the crSections object to the current report's sections
                    CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;

                    //loop through all the sections to find all the report objects
                    foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                    {
                        crReportObjects = crSection.ReportObjects;
                        //loop through all the report objects to find all the subreports
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                        {
                            if (crReportObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;

                                //open the subreport object
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                CrystalDecisions.CrystalReports.Engine.Database crSubDatabase;
                                CrystalDecisions.CrystalReports.Engine.Tables crSubTables;

                                //set the database and tables objects to work with the subreport
                                crSubDatabase = crSubreportDocument.Database;
                                crSubTables = crSubDatabase.Tables;

                                //loop through all the tables in the subreport and 
                                //set up the connection info and apply it to the tables
                                try
                                {
                                    //foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in crTables)
                                    foreach (CrystalDecisions.CrystalReports.Engine.Table subrptTable in crSubreportDocument.Database.Tables)
                                    {
                                        tLogonInfo = subrptTable.LogOnInfo;
                                        //tLogonInfo.ConnectionInfo.DatabaseName = newDataFile;
                                        //tLogonInfo.ConnectionInfo.ServerName = newDataFile;
                                        tLogonInfo.TableName = subrptTable.Name;
                                        tLogonInfo.ConnectionInfo.UserID = "";
                                        tLogonInfo.ConnectionInfo.Password = "";

                                        dtStart = DateTime.Now;

                                        try
                                        {
                                            subrptTable.ApplyLogOnInfo(tLogonInfo);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("ERROR: " + ex.Message);
                                            //return;
                                        }

                                        difference = DateTime.Now.Subtract(dtStart);
                                        btnSQLStatement.Text += "Subreport Table: " + subrptTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("SubReport ERROR: " + ex.Message);
                                    //return;
                                }
                            }
                        }
                    }
                    #endregion XML Subreport
                }
            }

            # endregion XML 

            if (btrVerifyDatabase.Checked)
            {
                if (chkUseRAS.Checked)
                    try
                    {
                        rptClientDoc.VerifyDatabase();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Verify Database ERROR: " + ex.Message);
                        //return;
                    }
                else
                    try
                    {
                        rpt.VerifyDatabase();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Verify Database ERROR: " + ex.Message);
                        //return;
                    }
            }

            difference = DateTime.Now.Subtract(TSTotal);
            btnSQLStatement.Text += "\nTotal time: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";

            MessageBox.Show("Data Source Set", "RAS", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SetPCDatabase_Click(object sender, EventArgs e)
        {
            DateTime dtStart;
            TimeSpan difference;
            DateTime TSTotal;
            TSTotal = DateTime.Now;

            string newDataFile = btrDataFile.Text;
            string newSearchPath = btrSearchPath.Text;

            // RAS Btrieve - Access need to use ReplaceConnection
            #region RASReplaceconnection;
            if (chkUseRAS.CheckState == CheckState.Checked)
            {
                rptClientDoc = rpt.ReportClientDocument;

                //Create a new Database Table to replace the reports current table.
                CrystalDecisions.ReportAppServer.DataDefModel.Table boTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                CrystalDecisions.ReportAppServer.DataDefModel.Table subboTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                // Get the old connection info
                oldConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                boTable.ConnectionInfo = boConnectionInfo;

                oldConnInfo = oldConnInfos[0];

                // Access
                #region DAO RAS Access
                if ((dynamic)oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_dao.dll")
                {
                    rptClientDoc.DatabaseController.LogonEx(@"D:\Atest" + @"\xtreme.mdb", "", "", "");

                    CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag logonProperties =
                        new CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag();
                    logonProperties.Add("Database Name", @"D:\Atest" + @"\xtreme.mdb");
                    logonProperties.Add("Database Type", "Access");

                    CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag connectionAttributes =
                        new CrystalDecisions.ReportAppServer.DataDefModel.PropertyBag();
                    connectionAttributes.Add("Database DLL", "crdb_dao.dll");
                    connectionAttributes.Add("QE_DatabaseName", @"D:\Atest" + @"\xtreme.mdb");
                    connectionAttributes.Add("QE_DatabaseType", "Access/Excel (DAO)");
                    connectionAttributes.Add("QE_LogonProperties", logonProperties);
                    connectionAttributes.Add("QE_ServerDescription", @"D:\Atest" + @"\xtreme.mdb");
                    connectionAttributes.Add("QE_SQLDB", true);
                    connectionAttributes.Add("SSO Enabled", false);

                    //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                    //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
                    //CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;

                    oldConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                    for (int I = 0; I < oldConnInfos.Count; I++)
                    {
                        oldConnInfo = oldConnInfos[I];
                        newConnInfo.Attributes = connectionAttributes;
                        newConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                        rptClientDoc.DatabaseController.ReplaceConnection(oldConnInfo, newConnInfo, null, crDBOptionUseDefault:
                            CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                    }
                    
/*                    PropertyBag logonDetails = new PropertyBag();
                    PropertyBag QeDetails = new PropertyBag();

                    //boMainPropertyBag: These hold the attributes of the tables ConnectionInfo object
                    PropertyBag boMainPropertyBag = new PropertyBag();
                    //boInnerPropertyBag: These hold the attributes for the QE_LogonProperties
                    //In the main property bag (boMainPropertyBag)
                    PropertyBag boInnerPropertyBag = new PropertyBag();

                    //Set the attributes for the boInnerPropertyBag
                    boInnerPropertyBag.Add("Database Name", newDataFile);
                    boInnerPropertyBag.Add("Database Type", "Access/Excel (DAO)");
                    boInnerPropertyBag.Add("Session UserID", btrFileLocation.Text);
                    boInnerPropertyBag.Add("Session Password", btrPassword.Text);

                    //Set the attributes for the boMainPropertyBag
                    boMainPropertyBag.Add("Database DLL", "crdb_dao.dll");
                    boMainPropertyBag.Add("QE_DatabaseName", newDataFile);
                    boMainPropertyBag.Add("QE_DatabaseType", "Access/Excel (DAO)"); // C:\xtreme.mdb
                    //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                    boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                    boMainPropertyBag.Add("QE_ServerDescription", newDataFile);
                    boMainPropertyBag.Add("QE_SQLDB", "false");
                    boMainPropertyBag.Add("SSO Enabled", "False");

                    // ReplaceConnection method This works also
                    for (int I = 0; I < oldConnInfos.Count; I++)
                    {
                        // CrystalDecisions.ReportAppServer.DataDefModel.Table tbl;
                        for (int S = 0; S < boTables.Count; S++)
                        {

                            oldConnInfo = oldConnInfos[I];
                            newConnInfo.Attributes = boMainPropertyBag;
                            newConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            try
                            {
                                //rptClientDoc.DatabaseController.SetConnectionInfos(oldConnInfo, newConnInfo);
                                rptClientDoc.DatabaseController.ReplaceConnection(oldConnInfos[I], newConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                                //rptClientDoc.DatabaseController.SetTableLocationEx(oldConnInfo, newConnInfo);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message); //  + " ;" + ex.InnerException.Message.ToString());
                                //return;
                            }
                        }
                    }
                    // ReplaceConnection method This works also

                    // check for subreports
                    //loop through all the report objects to find all the subreports
                    foreach (string crSubreportDocument1 in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        SubreportClientDocument SubRCD = rptClientDoc.SubreportController.GetSubreport(crSubreportDocument1);
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newSubConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldSubConnInfo;
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldSubConnInfos;

                        PropertyBag SublogonDetails = new PropertyBag();
                        PropertyBag SubQeDetails = new PropertyBag();

                        oldSubConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                        btnReportObjects.Text += "\nSubreport: " + crSubreportDocument1.ToString() + "\n";

                        for (int I = 0; I < oldSubConnInfos.Count; I++)
                        {
                            oldSubConnInfo = oldSubConnInfos[I];
                            newSubConnInfo.Attributes = boMainPropertyBag;
                            newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            // this works also
                            for (int S = 0; S < boTables.Count; S++)
                            {
                                oldSubConnInfo = oldSubConnInfos[I];
                                newSubConnInfo.Attributes = boMainPropertyBag;
                                newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                                try
                                {
                                    rptClientDoc.DatabaseController.ReplaceConnection(oldSubConnInfos[I], newSubConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                                    //rptClientDoc.DatabaseController.SetTableLocationEx(oldSubConnInfo, newSubConnInfo);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("ERROR: " + ex.Message);
                                    //return;
                                }
                            }
                        }
                        //rptClientDoc.VerifyDatabase();
                    }
 * */
                }
                #endregion DAO RAS Access
                // Access

                // Btrieve
                #region Btrieve RAS Access
                if ((dynamic)oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_p2bbtrv.dll")
                {
                    PropertyBag logonDetails = new PropertyBag();
                    PropertyBag QeDetails = new PropertyBag();

                    //boMainPropertyBag: These hold the attributes of the tables ConnectionInfo object
                    PropertyBag boMainPropertyBag = new PropertyBag();
                    //boInnerPropertyBag: These hold the attributes for the QE_LogonProperties
                    //In the main property bag (boMainPropertyBag)
                    PropertyBag boInnerPropertyBag = new PropertyBag();

                    //Set the attributes for the boInnerPropertyBag
                    boInnerPropertyBag.Add("Data File", newDataFile);
                    boInnerPropertyBag.Add("Data File Search Path", newSearchPath);
                    //boInnerPropertyBag.Add("Table Name", 
                    //Set the attributes for the boMainPropertyBag
                    boMainPropertyBag.Add("Database DLL", "crdb_p2bbtrv.dll");
                    boMainPropertyBag.Add("QE_DatabaseName", newSearchPath);
                    boMainPropertyBag.Add("QE_DatabaseType", "Btrieve");
                    //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
                    boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
                    boMainPropertyBag.Add("QE_ServerDescription", newDataFile);
                    boMainPropertyBag.Add("QE_SQLDB", "False");
                    boMainPropertyBag.Add("SSO Enabled", "False");

                    #region ReplaceConnection method This works also
                    //for (int I = 0; I < oldConnInfos.Count; I++)
                    //{
                    //    // CrystalDecisions.ReportAppServer.DataDefModel.Table tbl;
                    //    for (int S = 0; S < boTables.Count; S++)
                    //    {

                    //        oldConnInfo = oldConnInfos[I];
                    //        newConnInfo.Attributes = boMainPropertyBag;
                    //        newConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                    //        try
                    //        {
                    //            //rptClientDoc.DatabaseController.SetConnectionInfos(oldConnInfo, newConnInfo);
                    //            rptClientDoc.DatabaseController.ReplaceConnection(oldConnInfo, newConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                    //            //rptClientDoc.DatabaseController.SetTableLocationEx(oldConnInfo, newConnInfo);
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            MessageBox.Show("ERROR: " + ex.Message); //  + " ;" + ex.InnerException.Message.ToString());
                    //            //return;
                    //        }
                    //    }
                    //}
                    #endregion ReplaceConnection method This works also

                    // Set Location
                    foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table oldTable in rptClientDoc.DatabaseController.Database.Tables)
                    {
                        for (int I = 0; I < oldConnInfos.Count; I++)
                        {
                            //newDataFile = btrDataFile.ToString(); //@"D:\Atest\131755\Data\FILE.DDF";
                            //newSearchPath = btrSearchPath.ToString(); // @"D:\Atest\131755\Temp1\";

                            //newDataFile = @"D:\Atest\FILE.DDF";
                            //newSearchPath = @"D:\Atest\Temp1\";

                            logonDetails["Data File"] = newDataFile;
                            logonDetails["Data File Search Path"] = newSearchPath;

                            QeDetails.Add("Database DLL", "crdb_p2bbtrv.dll");
                            QeDetails.Add("QE_DatabaseType", "Btrieve");
                            QeDetails.Add("QE_LogonProperties", logonDetails);
                            QeDetails.Add("QE_LogonProperties", boInnerPropertyBag);
                            QeDetails.Add("QE_DatabaseName", newSearchPath);
                            QeDetails.Add("QE_ServerDescription", newDataFile);
                            QeDetails.Add("QE_SQLDB", "False");
                            QeDetails.Add("SSO Enabled", "False");

                            btrFileLocation.Text += oldTable.Alias.ToString() + ": ";

                            if (oldTable.Name == "TRIPS")
                                QeDetails.Add("File Name", "legrpt.tmp");
                            if (oldTable.Name == "LEG")
                                QeDetails.Add("File Name", "legpxrpt.tmp");

                            newConnInfo.Attributes = QeDetails;
                            newConnInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                            CrystalDecisions.ReportAppServer.DataDefModel.Table newTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                            newTable.ConnectionInfo = newConnInfo;

                            // this sets the table name
                            newTable.Name = oldTable.Name;
                            
                            dtStart = DateTime.Now;
                            rptClientDoc.DatabaseController.SetTableLocation(oldTable, newTable);
                            difference = DateTime.Now.Subtract(dtStart);
                            btnReportObjects.Text += oldTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                        }
                    }

                    # region Subreport
                    // check for subreports
                    //loop through all the report objects to find all the subreports
                    foreach (string crSubreportDocument1 in rptClientDoc.SubreportController.GetSubreportNames())
                    {
                        SubreportClientDocument SubRCD = rptClientDoc.SubreportController.GetSubreport(crSubreportDocument1);
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newSubConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldSubConnInfo;
                        CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldSubConnInfos;

                        PropertyBag SublogonDetails = new PropertyBag();
                        PropertyBag SubQeDetails = new PropertyBag();

                        oldSubConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                        btnReportObjects.Text += "\nSubreport: " + crSubreportDocument1.ToString() + "\n";

                        for (int I = 0; I < oldSubConnInfos.Count; I++)
                        {
                            oldSubConnInfo = oldSubConnInfos[I];
                            newSubConnInfo.Attributes = boMainPropertyBag;
                            newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            // this works also
                            //for (int S = 0; S < boTables.Count; S++)
                            //{
                            //    oldSubConnInfo = oldSubConnInfos[I];
                            //    newSubConnInfo.Attributes = boMainPropertyBag;
                            //    newSubConnInfo.Kind = CrystalDecisions.ReportAppServer.DataDefModel.CrConnectionInfoKindEnum.crConnectionInfoKindDBFile;

                            //    try
                            //    {
                            //        rptClientDoc.DatabaseController.ReplaceConnection(oldSubConnInfo, newSubConnInfo, null, CrystalDecisions.ReportAppServer.DataDefModel.CrDBOptionsEnum.crDBOptionDoNotVerifyDB);
                            //        //rptClientDoc.DatabaseController.SetTableLocationEx(oldSubConnInfo, newSubConnInfo);
                            //    }
                            //    catch (Exception ex)
                            //    {
                            //        MessageBox.Show("ERROR: " + ex.Message);
                            //        //return;
                            //    }
                            //}
                            //this works also

                            // Set Location 
                            foreach (CrystalDecisions.ReportAppServer.DataDefModel.Table oldSubTable in SubRCD.DatabaseController.Database.Tables)
                            {
                                for (I = 0; I < oldConnInfos.Count; I++)
                                {
                                    SublogonDetails["Data File"] = newDataFile;
                                    SublogonDetails["Data File Search Path"] = newSearchPath;

                                    SubQeDetails.Add("Database DLL", "crdb_p2bbtrv.dll");
                                    SubQeDetails.Add("QE_DatabaseType", "Btrieve");
                                    SubQeDetails.Add("QE_LogonProperties", logonDetails);
                                    SubQeDetails.Add("QE_LogonProperties", boInnerPropertyBag);
                                    SubQeDetails.Add("QE_DatabaseName", newSearchPath);
                                    SubQeDetails.Add("QE_ServerDescription", newDataFile);
                                    SubQeDetails.Add("QE_SQLDB", "False");
                                    SubQeDetails.Add("SSO Enabled", "False");

                                    if (oldSubTable.Name == "CREW")
                                        SubQeDetails.Add("File Name", "legcwrpt.tmp");
                                    if (oldSubTable.Name == "TRIP")
                                        SubQeDetails.Add("File Name", "legrpt.tmp");

                                    newSubConnInfo.Attributes = SubQeDetails;
                                    newSubConnInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
                                    newSubConnInfo.UserName = "";
                                    newSubConnInfo.Password = "";
                                    CrystalDecisions.ReportAppServer.DataDefModel.Table newSubTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                                    newSubTable.ConnectionInfo = newConnInfo;
                                    newSubTable.Name = oldSubTable.Name;

                                    dtStart = DateTime.Now;
                                    SubRCD.DatabaseController.SetTableLocation(oldSubTable, newSubTable);
                                    difference = DateTime.Now.Subtract(dtStart);
                                    btnReportObjects.Text += newSubTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";

                                }
                            }
                        }
                        # endregion Subreport
                    }
                    #endregion Btrieve RAS Access
                // Btrieve
                }
                IsRpt = false;
            }
            #endregion RASReplaceConnection;
            else
            #region Engine
            {
                // Access
                rptClientDoc = rpt.ReportClientDocument;

                //Create a new Database Table to replace the reports current table.
                CrystalDecisions.ReportAppServer.DataDefModel.Table boTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                CrystalDecisions.ReportAppServer.DataDefModel.Table subboTable = new CrystalDecisions.ReportAppServer.DataDefModel.Table();
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo newConnInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo oldConnInfo;
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfos oldConnInfos;
                CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();

                //Get the Database Tables Collection for your report
                CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
                boTables = rptClientDoc.DatabaseController.Database.Tables;

                // Get the old connection info
                oldConnInfos = rptClientDoc.DatabaseController.GetConnectionInfos(null);
                boTable.ConnectionInfo = boConnectionInfo;

                oldConnInfo = oldConnInfos[0];

                # region DAO Access
                if (oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_dao.dll")
                {
                    // Engine
                    CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                    CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                    CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
                    CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                    CrystalDecisions.CrystalReports.Engine.Tables crTables;

                    CrystalDecisions.Shared.TableLogOnInfo tLogonInfo;

                    btnSQLStatement.Text = "";

                    try
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in rpt.Database.Tables)
                        {
                            tLogonInfo = rptTable.LogOnInfo;
                            tLogonInfo.ConnectionInfo.ServerName = newDataFile; // @"D:\Atest\dsTimesheet.xml";
                            tLogonInfo.ConnectionInfo.DatabaseName = btrSearchPath.Text; // D:\Atest\extreme.mdb
                            tLogonInfo.ConnectionInfo.UserID = btrFileLocation.Text;
                            tLogonInfo.ConnectionInfo.Password = btrPassword.Text;
                            tLogonInfo.TableName = rptTable.Name;

                            dtStart = DateTime.Now;

                            try
                            {
                                rptTable.ApplyLogOnInfo(tLogonInfo);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message);
                                //return;
                            }

                            difference = DateTime.Now.Subtract(dtStart);

                            //rptTable.Location = rptTable.Name;
                            btnSQLStatement.Text += /*rptTable.Name.ToString() +*/ " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                        //return;
                    }

                    // check for subreports
                    //set the crSections object to the current report's sections
                    CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;

                    //loop through all the sections to find all the report objects
                    foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                    {
                        crReportObjects = crSection.ReportObjects;
                        //loop through all the report objects to find all the subreports
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                        {
                            if (crReportObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;

                                //open the subreport object
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                CrystalDecisions.CrystalReports.Engine.Database crSubDatabase;
                                CrystalDecisions.CrystalReports.Engine.Tables crSubTables;

                                //set the database and tables objects to work with the subreport
                                crSubDatabase = crSubreportDocument.Database;
                                crSubTables = crSubDatabase.Tables;

                                //loop through all the tables in the subreport and 
                                //set up the connection info and apply it to the tables
                                try
                                {
                                    //foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in crTables)
                                    foreach (CrystalDecisions.CrystalReports.Engine.Table subrptTable in crSubreportDocument.Database.Tables)
                                    {
                                        tLogonInfo = subrptTable.LogOnInfo;
                                        //tLogonInfo.ConnectionInfo.DatabaseName = newDataFile;
                                        //tLogonInfo.ConnectionInfo.ServerName = newDataFile;
                                        tLogonInfo.TableName = subrptTable.Name;
                                        //tLogonInfo.ConnectionInfo.UserID = "";
                                        //tLogonInfo.ConnectionInfo.Password = "";

                                        tLogonInfo.ConnectionInfo.ServerName = newDataFile; // @"D:\Atest\dsTimesheet.xml";
                                        tLogonInfo.ConnectionInfo.DatabaseName = btrSearchPath.Text; // D:\Atest\extreme.mdb
                                        tLogonInfo.ConnectionInfo.UserID = btrFileLocation.Text;
                                        tLogonInfo.ConnectionInfo.Password = btrPassword.Text;

                                        dtStart = DateTime.Now;

                                        try
                                        {
                                            subrptTable.ApplyLogOnInfo(tLogonInfo);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("ERROR: " + ex.Message);
                                            //return;
                                        }

                                        difference = DateTime.Now.Subtract(dtStart);
                                        btnSQLStatement.Text += "Subreport Table: " + subrptTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("SubReport ERROR: " + ex.Message);
                                    //return;
                                }
                            }
                        }
                    }
                }

                # endregion DAO Access
                    // Access
                else
                {                    
                    // Btrieve
                    #region Btrieve

                    // Engine
                    if (oldConnInfo.Attributes["Database DLL"].ToString() == "crdb_p2bbtrv.dll")
                    {
                                            // Engine
                    CrystalDecisions.CrystalReports.Engine.ReportObjects crReportObjects;
                    CrystalDecisions.CrystalReports.Engine.SubreportObject crSubreportObject;
                    CrystalDecisions.CrystalReports.Engine.ReportDocument crSubreportDocument;
                    CrystalDecisions.CrystalReports.Engine.Database crDatabase;
                    CrystalDecisions.CrystalReports.Engine.Tables crTables;

                    CrystalDecisions.Shared.TableLogOnInfo tLogonInfo;

                    btnSQLStatement.Text = "";

                    try
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in rpt.Database.Tables)
                        {
                            tLogonInfo = rptTable.LogOnInfo;
                            tLogonInfo.ConnectionInfo.DatabaseName = newSearchPath;
                            tLogonInfo.ConnectionInfo.ServerName = newDataFile;
                            tLogonInfo.ConnectionInfo.UserID = "";
                            tLogonInfo.ConnectionInfo.Password = "";
                            tLogonInfo.TableName = rptTable.Name;

                            dtStart = DateTime.Now;

                            try
                            {
                                rptTable.ApplyLogOnInfo(tLogonInfo);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message);
                                //return;
                            }

                            difference = DateTime.Now.Subtract(dtStart);

                            //rptTable.Location = rptTable.Name;
                            btnSQLStatement.Text += /*rptTable.Name.ToString() +*/ " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";

                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                        //return;
                    }

                    // check for subreports
                    //set the crSections object to the current report's sections
                    CrystalDecisions.CrystalReports.Engine.Sections crSections = rpt.ReportDefinition.Sections;

                    crSections = rpt.ReportDefinition.Sections;

                    //loop through all the sections to find all the report objects
                    foreach (CrystalDecisions.CrystalReports.Engine.Section crSection in crSections)
                    {
                        crReportObjects = crSection.ReportObjects;
                        //loop through all the report objects to find all the subreports
                        foreach (CrystalDecisions.CrystalReports.Engine.ReportObject crReportObject in crReportObjects)
                        {
                            if (crReportObject.Kind == CrystalDecisions.Shared.ReportObjectKind.SubreportObject)
                            {
                                //you will need to typecast the reportobject to a subreport 
                                //object once you find it
                                crSubreportObject = (CrystalDecisions.CrystalReports.Engine.SubreportObject)crReportObject;

                                //open the subreport object
                                crSubreportDocument = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);

                                CrystalDecisions.CrystalReports.Engine.Database crSubDatabase;
                                CrystalDecisions.CrystalReports.Engine.Tables crSubTables;

                                //set the database and tables objects to work with the subreport
                                crSubDatabase = crSubreportDocument.Database;
                                crSubTables = crSubDatabase.Tables;

                                //loop through all the tables in the subreport and 
                                //set up the connection info and apply it to the tables
                                try
                                {
                                    //foreach (CrystalDecisions.CrystalReports.Engine.Table rptTable in crTables)
                                    foreach (CrystalDecisions.CrystalReports.Engine.Table subrptTable in crSubreportDocument.Database.Tables)
                                    {
                                        tLogonInfo = subrptTable.LogOnInfo;
                                        tLogonInfo.ConnectionInfo.DatabaseName = newDataFile;
                                        tLogonInfo.ConnectionInfo.ServerName = newDataFile;
                                        tLogonInfo.TableName = subrptTable.Name;

                                        dtStart = DateTime.Now;

                                        try
                                        {
                                            subrptTable.ApplyLogOnInfo(tLogonInfo);
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("ERROR: " + ex.Message);
                                            //return;
                                        }

                                        difference = DateTime.Now.Subtract(dtStart);
                                        btnSQLStatement.Text += "Subreport Table: " + subrptTable.Name.ToString() + " Set in " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("SubReport ERROR: " + ex.Message);
                                    //return;
                                }
                            }
                        }
                    }
                }
                #endregion Btrieve
                // btrieve
            }
            #endregion Engine
            }
            if (btrVerifyDatabase.Checked)
            {
                if (chkUseRAS.Checked)
                    try
                    {
                        rptClientDoc.VerifyDatabase();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                    }
                else
                {
                    try
                    {
                        rpt.VerifyDatabase();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                    }
                }
            }

            difference = DateTime.Now.Subtract(TSTotal);
            btnSQLStatement.Text += "\nTotal time: " + difference.Minutes.ToString() + ":" + difference.Seconds.ToString() + ":" + difference.Milliseconds.ToString() + "\n";
        }

        #region All Viewer Events

        private void crystalReportViewer1_Navigate(object source, CrystalDecisions.Windows.Forms.NavigateEventArgs e) 
        {
            if (!e.Handled)
            {
                crystalReportViewer1.ShowProgressAnimation(true);
                //MessageBox.Show(e.CurrentPageNumber.ToString());
            }
        }

        private void crystalReportViewer1_DoubleClickPage(object sender, PageMouseEventArgs e)
        {
            //// do not edit this works
            //CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc;
            //rptClientDoc = rpt.ReportClientDocument;
            //crystalReportViewer1.SetFocusOn(UIComponent.Page);

            //if (e.ObjectInfos != null)
            //{
            //    btnReportObjects.Text = "";
            //    btnReportObjects.Text += "Number of items in selected Row: " + e.ObjectInfos.Length.ToString() + "\n\n";
            //    foreach (ObjectInfo oi in e.ObjectInfos)
            //    {
            //        //btnReportObjects.Text += ">>>OBJECT : " + t + " <<<" + "\n";
            //        btnReportObjects.Text += "Fld Text: " + oi.Text.ToString() + "\n";
            //        btnReportObjects.Text += "DataContext: " + oi.DataContext.ToString() + "\n";
            //        btnReportObjects.Text += "GroupNamePath: " + oi.GroupNamePath.ToString() + "\n";
            //        btnReportObjects.Text += "Hyperlink: " + oi.Hyperlink.ToString() + "\n";
            //        try
            //        {
            //            foreach (FormulaField resultField in rptClientDoc.DataDefController.DataDefinition.FormulaFields)
            //            {
            //                if ((resultField.UseCount > 0) && (resultField.Text == oi.Text))
            //                {
            //                    CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject objField;
            //                    objField = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)rptClientDoc.ReportDefController.ReportDefinition.FindObjectByName(resultField.Text.ToString());
            //                    richTextBox1.Text += "Formula field object: " + objField.Name.ToString() + "\n"; // oi.DataContext.ToString() ;
            //                }
            //            }
            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show("ERROR: " + ex.Message);
            //            return;
            //        }
            //        btnReportObjects.Text += "ObjectType: " + oi.ObjectType.ToString() + "\n";
            //        btnReportObjects.Text += "Tooltip: " + oi.ToolTip.ToString() + "\n";
            //        btnReportObjects.Text += "========================\n";
            //    }
            //    btnReportObjects.AppendText(" 'End' \n");
            //}
            ////do not edit - this works

            // Gary Chang sent me this not sure how it's diff than above but it's better
            CrystalDecisions.ReportAppServer.ClientDoc.ISCDReportClientDocument rptClientDoc;
            rptClientDoc = rpt.ReportClientDocument;

            rptClientDoc.ReportOptions.ConvertNullFieldToDefault = true;
            rptClientDoc.ReportOptions.ConvertOtherNullsToDefault = true;

            btnReportObjects.Text = "";
            int t = 0;
            if (e.ObjectInfos != null)
            {
                foreach (ObjectInfo oi in e.ObjectInfos)
                {
                    try
                    {
                        if (oi.ObjectType.ToString() != "GroupChart")
                        {
                            if (oi.ObjectType.ToString() != "OleObject")
                            {
                                if (oi.ObjectType.ToString() != "CrossTab")
                                {
                                    // This is the field you double clicked on, in the viewer
                                    if (oi.Name == e.ObjectInfo.Name)
                                        btnReportObjects.Text += @">>> SELECTED OBJECT - you clicked on<<<\n";

                                    btnReportObjects.Text += "DataContext: " + oi.DataContext.ToString() + "\n";
                                    btnReportObjects.Text += "GroupNamePath: " + oi.GroupNamePath.ToString() + "\n";
                                    btnReportObjects.Text += "Hyperlink: " + oi.Hyperlink.ToString() + "\n";

                                    string textBox1 = String.Empty;
                                    foreach (FormulaField resultField in rptClientDoc.DataDefController.DataDefinition.FormulaFields)
                                    {
                                        CrystalDecisions.ReportAppServer.ReportDefModel.ReportObjects FormulaFldobjs;
                                        FormulaFldobjs = rptClientDoc.ReportDefController.ReportObjectController.GetReportObjectsByKind(CrReportObjectKindEnum.crReportObjectKindField);
                                        t = rptClientDoc.DataDefController.DataDefinition.FormulaFields.FindIndexOf(resultField);
                                        CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject objField;
                                        objField = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)FormulaFldobjs[t];
                                        CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject objField2 = (CrystalDecisions.ReportAppServer.ReportDefModel.FieldObject)objField.Clone(true);

                                        if (oi.Name.ToString() == objField2.Name.ToString())
                                        {
                                            btnReportObjects.Text += "Name: " + oi.Name.ToString() + " -> " + objField2.DataSourceName.ToString() + "\n";
                                        }

                                        t++;
                                    }

                                    //btnReportObjects.Text += "Name: " + oi.Name.ToString() + "\n";
                                    btnReportObjects.Text += "Value: " + oi.Text.ToString() + "\n";
                                    btnReportObjects.Text += "ObjectType: " + oi.ObjectType.ToString() + "\n";
                                    btnReportObjects.Text += "Tooltip: " + oi.ToolTip.ToString() + "\n";
                                    btnReportObjects.Text += "========================\n";
                                    //btnReportObjects.Text += richTextBox1;
                                    btnReportObjects.AppendText("\n\n");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("ERROR: " + ex.Message);
                    }
                }
            }
        }

        private void crystalReportViewer1_MouseClick(object sender, MouseEventArgs e)
        {
            int i;
            i = 0;
            //crystalReportViewer1.Zoom(1);
        }

        private void crystalReportViewer1_ClickPage_1(object sender, PageMouseEventArgs e)
        {
            //crystalReportViewer1.ExportReport();
            //crystalReportViewer1.Zoom(1);
        }

        private void crystalReportViewer1_Drill_1(object source, DrillEventArgs e)
        {
            //crystalReportViewer1.Zoom(1);
            //crystalReportViewer1.SetFocusOn(UIComponent.Page);
            string tmp = string.Empty;
            tmp = e.NewGroupPath;

            CrystalDecisions.ReportAppServer.DataDefModel.ISCRField iscrField;

            if (rptClientDoc.DataDefController.DataDefinition.FormulaFields.Count > 0)
            {
                iscrField = (CrystalDecisions.ReportAppServer.DataDefModel.ISCRField)rptClientDoc.DataDefController.DataDefinition.FormulaFields[0];
                if (rptClientDoc.RowsetController.CanBrowseField(iscrField))
                {
                    rptClientDoc.RowsetController.BrowseFieldValues(iscrField, 10);
                }
                else
                    btnRecordSelectionForm.Text = "Field is wrong type to browse";
            }

        }

        private void crystalReportViewer1_MouseClick_1(object sender, MouseEventArgs e)
        {
            //crystalReportViewer1.Zoom(2);
        }

        private void crystalReportViewer1_Error(object source, ExceptionEventArgs e)
        {
            string eError = "";
            int StringLen = 0;
            eError = e.Exception.Message.ToString();
            StringLen = eError.Length;
            if (StringLen >= 61)
            {
                if (eError.Substring(0, 61) == @"The remaining text does not appear to be part of the formula.")
                {
                    MessageBox.Show("Possible missing UFL or actual error in formula");
                    eError = eError.Substring(0, StringLen - 2);
                    ////cool but of no use to release mode
                    string myURL = @"http://search.sap.com/ui/scn#query=" + eError + "&startindex=1&filter=scm_a_resourceType(scm_v_resType252)";
                    string fixedString = myURL.Replace(" ", "%20");

                    //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);
                    e.Handled = true;
                }
                else
                {
                    MessageBox.Show("Trigger Event - Error in Viewer: " + e.ToString());
                }
            }
            else
            {
                if (eError.ToString() == "The system cannot find the path specified.\r")
                {
                    MessageBox.Show(eError);
                    eError = eError.Substring(0, StringLen - 2);
                    ////cool but of no use to release mode
                    string myURL = @"http://search.sap.com/ui/scn#query=" + eError + "&startindex=1&filter=scm_a_resourceType(scm_v_resType252)";
                    string fixedString = myURL.Replace(" ", "%20");

                    //System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Internet Explorer\iexplore.exe", myURL);
                    System.Diagnostics.Process.Start(@"C:\Program Files (x86)\Mozilla Firefox\firefox.exe", fixedString);

                    ////string myURL = @"C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports 2011\Help\en\crw.chm";
                    ////System.Diagnostics.Process.Start(myURL);
                    ////cool but of no use to release mode
                }
                else
                    MessageBox.Show("Break on this line to get the actaul error being trapped");

                e.Handled = true;
            }
        }

        private void Form2_Click(object sender, EventArgs e, CrystalDecisions.CrystalReports.Engine.ReportDocument rpt)
        {

        }

        private void crystalReportViewer1_Validated(object sender, EventArgs e)
        {
            //MessageBox.Show("Hello");
        }

        private void crystalReportViewer1_Validating(object sender, CancelEventArgs e)
        {
            //MessageBox.Show("Hello 1");
        }

        private void crystalReportViewer1_ImeModeChanged(object sender, EventArgs e)
        {
            MessageBox.Show("Hello 2");
        }

        private void crystalReportViewer1_VisibleChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Hello");
        }

        private void crystalReportViewer1_RegionChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Hello");
        }

        private void crystalReportViewer1_EnabledChanged(object sender, EventArgs e)
        {
            MessageBox.Show("Hello");
        }

        int myZoom;

        private void crystalReportViewer1_ViewZoom(object source, ZoomEventArgs e)
        {
            myZoom = e.NewZoomFactor;
            System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog("event.log");
            btnSQLStatement.Text +=
              "\nZoom event:" + "\n" +
              "CurrentZoomFactor: " + e.CurrentZoomFactor + "\n" +
              "NewZoomFactor: " + e.NewZoomFactor;
            btnSQLStatement.AppendText("\n");
            e.Handled = false;
        }

        private void crystalReportViewer1_DrillDownSubreport(object source, DrillSubreportEventArgs e)
        {
            CrystalDecisions.Windows.Forms.ZoomEventArgs myZoomFactor = new ZoomEventArgs();
            //crystalReportViewer1.ExportReport();
            //crystalReportViewer1.Zoom(myZoom);
        }

        private void crystalReportViewer1_HelpRequested(object sender, HelpEventArgs hlpevent)
        {
            Help.ShowHelp(this, @"C:\Program Files (x86)\SAP BusinessObjects\Crystal Reports 2011\Help\en\Crw.chm");
        }

        private void crystalReportViewer1_DoubleClickPage(object sender, MouseEventArgs e)
        {

        }
        #endregion All Viewer Events

        public static DBField DataColumnToField(string ColumnName)
        {
            DBField class2 = new DBField();
            int rnBytesInField = 0;
            class2.Name = ColumnName;
            class2.Type = CrystalDecisions.ReportAppServer.DataDefModel.CrFieldValueTypeEnum.crFieldValueTypeStringField;
            class2.Length = 1024;
            //class2.Description = column.Caption;
            //class2.HeadingText = column.Expression;
            return class2;
        }

        private void btrRefreshViewer_Click(object sender, EventArgs e)
        {
            crystalReportViewer1.RefreshReport();
            //crystalReportViewer1.ReportRefresh;
        }

        private void crystalReportViewer1_Paint(object sender, PaintEventArgs e)
        {
            //crystalReportViewer1.Zoom(300);
        }

        private void CommandTable_Click_1(object sender, EventArgs e)
        {
            String newTableName = "Command";
            CrystalDecisions.Shared.ConnectionInfo connectionInfo = new CrystalDecisions.Shared.ConnectionInfo();
            DbConnectionAttributes dbconn = new DbConnectionAttributes();
            NameValuePairs2 propertyBag = new NameValuePairs2();

            // OLE DB Connection
            //propertyBag.Set("QE_DatabaseDLL", "crdb_ado.dll");
            //propertyBag.Set("QE_Servertype", "OLE DB (ADO)");
            //propertyBag.Set("QE_ConnectionString", "Provider=SQLNCLI11;User ID=;Initial Catalog=xtreme;Data Source=10.161.27.110;");
            ////propertyBag.Set("QE_ConnectionString", "Provider=SQLNCLI11.1;Persist Security Info=False;User ID=sa;Initial Catalog=;Data Source=10.161.27.110;Application Name=Ducky Connection;Initial File Name=;Server SPN=;");
            //propertyBag.Set("QE_Servername", "10.161.27.110");

            ////ODBC connection 
            //propertyBag.Set("QE_DatabaseDLL", "crdb_odbc.dll");
            //propertyBag.Set("QE_Servertype", "ODBC (RDO)");
            //propertyBag.Set("QE_ConnectionString", "Provider=ODBCTest;User ID=;Initial Catalog=xtreme;Data Source=10.161.27.110;");
            ////propertyBag.Set("QE_ConnectionString", "Provider=SQLNCLI11.1;Persist Security Info=False;User ID=sa;Initial Catalog=;Data Source=10.161.27.110;Application Name=Ducky Connection;Initial File Name=;Server SPN=;");
            //propertyBag.Set("QE_Servername", "ODBCTest");

            //dbconn.Collection = propertyBag;
            //connectionInfo.Attributes = dbconn;
            //connectionInfo.UserID = "sa";
            //connectionInfo.Password = "1Oem2000";

            //SetCrystalParam(rpt, "DonTest", "1007");

            //// OLE DB Server Name or IP Connection
            ////connectionInfo.ServerName = "10.161.27.110";
            ////ODBC DSN Name
            //connectionInfo.ServerName = "ODBCTest";

            //connectionInfo.DatabaseName = "xtreme";

            //String sqlQueryString = "SELECT \"Orders_Detail\".\"Order ID\", \"Orders_Detail\".\"Product ID\", \"Orders_Detail\".\"Unit Price\", \"Orders_Detail\".\"Quantity\" FROM \"xtreme\".\"dbo\".\"Orders Detail\" \"Orders_Detail\" WHERE \"Orders_Detail\".\"Order ID\"<={?DonTest}";

            //try
            //{
            //    rpt.SetSQLCommandTable(connectionInfo, newTableName, sqlQueryString.ToString());
            //    btnSQLStatement.Text = "Command Table SQL updated: \n" + sqlQueryString;
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("ERROR: " + ex.Message);
            //    btnSQLStatement.Text = "Command Table SQL update failed";
            //}

            //Create a new Command Table to replace the reports current table.
            CrystalDecisions.ReportAppServer.DataDefModel.CommandTable boTable = new CrystalDecisions.ReportAppServer.DataDefModel.CommandTable();
            //Get the Database Tables Collection for your report
            CrystalDecisions.ReportAppServer.DataDefModel.Tables boTables;
            boTables = rptClientDoc.DatabaseController.Database.Tables;

            // something missing from the collection property bag from gen code app, cloning gets what ever was missing...
            CommandTable TableOld = (CommandTable)boTables[0];
            boTable = (CommandTable)TableOld.Clone(true);

            //boMainPropertyBag: These hold the attributes of the tables ConnectionInfo object
            PropertyBag boMainPropertyBag = new PropertyBag();
            //boInnerPropertyBag: These hold the attributes for the QE_LogonProperties
            //In the main property bag (boMainPropertyBag)
            PropertyBag boInnerPropertyBag = new PropertyBag();

            //Set the attributes for the boInnerPropertyBag
            boInnerPropertyBag.Add("Database", "xtreme");
            boInnerPropertyBag.Add("DSN", "DonSQL");
            boInnerPropertyBag.Add("UseDSNProperties", "False");

            //Set the attributes for the boMainPropertyBag
            boMainPropertyBag.Add("Database DLL", "crdb_odbc.dll");
            boMainPropertyBag.Add("QE_DatabaseName", "xtreme");
            boMainPropertyBag.Add("QE_DatabaseType", "ODBC (RDO)");
            //Add the QE_LogonProperties we set in the boInnerPropertyBag Object
            boMainPropertyBag.Add("QE_LogonProperties", boInnerPropertyBag);
            boMainPropertyBag.Add("QE_ServerDescription", "DonSQL");
            boMainPropertyBag.Add("QE_SQLDB", "True");
            boMainPropertyBag.Add("SSO Enabled", "False");

            //Create a new ConnectionInfo object
            CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo boConnectionInfo = new CrystalDecisions.ReportAppServer.DataDefModel.ConnectionInfo();
            //Pass the database properties to a connection info object
            boConnectionInfo.Attributes = boMainPropertyBag;
            //Set the connection kind
            boConnectionInfo.Kind = CrConnectionInfoKindEnum.crConnectionInfoKindCRQE;
            //**EDIT** Set the User Name and Password if required.
            boConnectionInfo.UserName = "sa";
            boConnectionInfo.Password = "1Oem2000";
            //Pass the connection information to the table
            boTable.ConnectionInfo = boConnectionInfo;

            boTable.Name = "DonSQL";
            boTable.QualifiedName = "Command";
            boTable.Alias = "Command";
            boTable.CommandText = "Select * from xtreme.dbo.credit where " + (char)34 + "Customer Credit ID" + (char)34 + " = {?test}";

            IsRpt = false;

            rptClientDoc.DatabaseController.SetTableLocation(TableOld, boTable);
            // this works do not alter
            // MUST verify the DB so it updates all of the various property bags.
            rptClientDoc.VerifyDatabase();
        }

    }
}