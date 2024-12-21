'******************************************************************************************
'
'File Name:     SimplePreviewReport.sln
'Created:       May 9, 2003
'Author ID:     ALD
'Purpose:       This sample .NET Windows application demonstrates five different methods of 
'               loading a report in the Windows Forms viewer. The types of reports set as 
'               the reportsource for the viewer are:
'
'               - By Report Name (from string path)
'               - By Report Object (from string path)
'               - By Untyped Report Component (from string path)
'               - By Report Object (added to Project)
'               - By Strongly-Typed Report Component
'             
'
'******************************************************************************************


Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Report load method explanations"
    Dim text1 As String = "This method of loading a report simply specifies a string path to the physical location of the report and " _
    & "sets it as the reportsource for the viewer." & vbCrLf & vbCrLf & "The code used for this is:" & vbCrLf & vbCrLf _
    & "CrystalReportViewer1.ReportSource = " & Chr(34) & "C:\fullPathToReport\World Sales Report.rpt" & Chr(34) & vbCrLf

    Dim text2 As String = "This method of loading a report creates a new instance of the ReportDocument object from the " _
    & "CrystalDecisions.CrystalReports.Engine namespace and loads a report into the object by specifying the physical path to the report." & vbCrLf _
    & "The code used to bind the report to the viewer is as follows:" & vbCrLf & vbCrLf _
    & "Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()" & vbCrLf _
    & "crReportDocument.Load(" & Chr(34) & "C:\fullPathToReport\World Sales Report.rpt" & Chr(34) & ")" & vbCrLf _
    & "CrystalReportViewer1.ReportSource = crReportDocument" & vbCrLf

    Dim text3 As String = "This method of loading a report uses the ReportDocument component from the Components tab of the ToolBox.  " _
    & "When this component is added to the form, a dialogue box appears and prompts the Developer as to which type of Component to add " _
    & "to the form.  The options selected were 'Untyped ReportDocument' and a component called ReportDocument1 was added to the form.  " _
    & "This component acts as a proxy to the ReportDocument class." & vbCrLf _
    & "The code used to bind the report to the viewer is as follows:" & vbCrLf & vbCrLf _
    & "Dim ReportDocument1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()" & vbCrLf _
    & "ReportDocument1.Load(" & Chr(34) & "C:\fullPathToReport\World Sales Report.rpt" & Chr(34) & ")" & vbCrLf _
    & "CrystalReportViewer1.ReportSource = ReportDocument1" & vbCrLf


    Dim text4 As String = "This scenario uses a report object that has been added to the current Project.  When a report is added to a project, " _
    & "a class file is generated for the report making the report a strongly-typed object, in this case 'World Sales Report.rpt'" & vbCrLf _
    & "The code used to bind the report to the viewer is as follows:" & vbCrLf & vbCrLf _
    & "CrystalReportViewer1.ReportSource = New World_Sales_Report()" & vbCrLf

    Dim text5 As String = "This method of loading a report uses the ReportDocument component from the Components tab of the ToolBox.  " _
    & "When this component is added to the form, a dialogue box appears and prompts the Developer as to which type of Component to add " _
    & "to the form.  The 'Name' option selected was the strongly-typed report that was added to the project (simplepreviewreport.World_Sales_Report).  The option to generate a Cached " _
    & "Strongly-Typed Report was deselected and a report component called 'world_Sales_Report1' was created and added to the form." & vbCrLf _
    & "The code used to bind the report to the viewer is as follows:" & vbCrLf & vbCrLf _
    & "CrystalReportViewer1.ReportSource = world_Sales_Report1" & vbCrLf

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        CrystalReportViewer1.Visible = False
        TextBox1.Visible = False

        With ComboBox1.Items
            .Add("")
            .Add("By Report Name (from string path)")
            .Add("By Report Object (from string path)")
            .Add("By Untyped Report Component (from string path)")
            .Add("By Report Object (added to Project)")
            .Add("By Strongly-Typed Report Component")
        End With

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ReportDocument1 As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Friend WithEvents world_Sales_Report1 As SimplePreviewReport.World_Sales_Report
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.ReportDocument1 = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Me.world_Sales_Report1 = New SimplePreviewReport.World_Sales_Report()
        Me.SuspendLayout()
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(0, 173)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(616, 400)
        Me.CrystalReportViewer1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(232, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select a method for viewing the report"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Location = New System.Drawing.Point(8, 24)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(344, 24)
        Me.ComboBox1.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(360, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 40)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Load Report"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(8, 56)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(520, 96)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = ""
        '
        'ReportDocument1
        '
        Me.ReportDocument1.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation
        Me.ReportDocument1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        Me.ReportDocument1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Upper
        Me.ReportDocument1.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default
        '
        'world_Sales_Report1
        '
        Me.world_Sales_Report1.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation
        Me.world_Sales_Report1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        Me.world_Sales_Report1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Upper
        Me.world_Sales_Report1.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(616, 573)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.TextBox1, Me.Button1, Me.ComboBox1, Me.Label1, Me.CrystalReportViewer1})
        Me.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Sub LoadReportBySelectedMethod()
        CrystalReportViewer1.Visible = True
        TextBox1.Visible = True
        Select Case ComboBox1.Text 'this contains the value of the selected load method.
            '--------------------------------------------------------------------
        Case ("By Report Name (from string path)")
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text1

                'Bind the report to the viewer
                CrystalReportViewer1.ReportSource = "C:\Crystal\crnet\vbnet_win_simplepreviewreport\World Sales Report.rpt"


                '--------------------------------------------------------------------
            Case "By Report Object (from string path)"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text2

                'create a new instance of the reportdocument object and load the report
                Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
                crReportDocument.Load("C:\Crystal\crnet\vbnet_win_simplepreviewreport\World Sales Report.rpt")

                'Bind the report to the viewer
                CrystalReportViewer1.ReportSource = crReportDocument


                '--------------------------------------------------------------------
            Case "By Untyped Report Component (from string path)"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text3
                'Here 


                Dim reportdocument1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
                'load the report into the report document component and set the component as the reportsource for the viewer
                reportdocument1.Load("C:\Crystal\crnet\vbnet_win_simplepreviewreport\World Sales Report.rpt")
                CrystalReportViewer1.ReportSource = reportdocument1


                '--------------------------------------------------------------------
            Case "By Report Object (added to Project)"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text4

                'Bind a new instance of the strongly-typed report object to the viewer
                CrystalReportViewer1.ReportSource = New World_Sales_Report()


                '--------------------------------------------------------------------
            Case "By Strongly-Typed Report Component"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text5

                'Bind the strongly typed report component to the viewer
                CrystalReportViewer1.ReportSource = world_Sales_Report1

        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadReportBySelectedMethod()
    End Sub
End Class
