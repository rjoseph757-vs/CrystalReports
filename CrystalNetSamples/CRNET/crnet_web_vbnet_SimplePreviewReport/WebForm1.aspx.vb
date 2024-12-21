'********************************************
'
'File Name:     SimplePreviewReport.sln
'Created:       Jan 15, 2002
'Author ID:     DAK
'Purpose:       This sample .NET Web application demonstrates six different methods of 
'               loading a report in the Web Forms viewer. The types of reports set as 
'               the reportsource for the viewer are:
'
'               - By Report Name (from string path)
'               - By Report Object (from string path)
'               - By Untyped Report Component (from string path)
'               - By Report Object (added to Project)
'               - By Strongly-Typed Report Component
'               - By Cached Report Component 
'
'********************************************


Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents CrystalReportViewer1 As CrystalDecisions.Web.CrystalReportViewer
    Protected WithEvents DropDownList1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents reportDocument1 As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Protected WithEvents world_Sales_Report1 As crnet_web_vbnet_SimplePreviewReport.World_Sales_Report
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents cachedWorld_Sales_Report2 As crnet_web_vbnet_SimplePreviewReport.CachedWorld_Sales_Report
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
    & "Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()" & vbCrLf _
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


    Dim text6 As String = "This method of loading a report uses the ReportDocument component from the Components tab of the ToolBox.  " _
    & "When this component is added to the form, a dialogue box appears and prompts the Developer as to which type of Component to add " _
    & "to the form.  The 'Name' option selected was the strongly-typed report that was added to the project(simplepreviewreport.World_Sales_Report)." _
    & "The option to generate a Cached Strongly-Typed Report was selected and a report component called 'CachedWorld_Sales_Report2' was created and added to the form." & vbCrLf _
    & "The code used to bind the report to the viewer is as follows:" & vbCrLf & vbCrLf _
    & "CrystalReportViewer1.ReportSource = cachedWorld_Sales_Report2"

#End Region


#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.reportDocument1 = New CrystalDecisions.CrystalReports.Engine.ReportDocument()
        Me.world_Sales_Report1 = New crnet_web_vbnet_SimplePreviewReport.World_Sales_Report()
        Me.cachedWorld_Sales_Report2 = New crnet_web_vbnet_SimplePreviewReport.CachedWorld_Sales_Report()
        '
        'reportDocument1
        '
        Me.reportDocument1.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation
        Me.reportDocument1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        Me.reportDocument1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Upper
        Me.reportDocument1.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default
        '
        'world_Sales_Report1
        '
        Me.world_Sales_Report1.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.DefaultPaperOrientation
        Me.world_Sales_Report1.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize
        Me.world_Sales_Report1.PrintOptions.PaperSource = CrystalDecisions.Shared.PaperSource.Upper
        Me.world_Sales_Report1.PrintOptions.PrinterDuplex = CrystalDecisions.Shared.PrinterDuplex.Default
        '
        'cachedWorld_Sales_Report2
        '
        Me.cachedWorld_Sales_Report2.CacheTimeOut = System.TimeSpan.Parse("02:10:00")
        Me.cachedWorld_Sales_Report2.IsCacheable = True
        Me.cachedWorld_Sales_Report2.ShareDBLogonInfo = False

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()


        CrystalReportViewer1.Visible = False
        TextBox1.Visible = False

        With DropDownList1.Items
            .Add("")
            .Add("By Report Name (from string path)")
            .Add("By Report Object (from string path)")
            .Add("By Untyped Report Component (from string path)")
            .Add("By Report Object (added to Project)")
            .Add("By Strongly-Typed Report Component")
            .Add("By Cached Report Component")
        End With


    End Sub

#End Region
    Sub LoadReportBySelectedMethod()
        CrystalReportViewer1.Visible = True
        TextBox1.Visible = True
        Select Case DropDownList1.SelectedItem.Text 'this contains the value of the selected export format.
            '--------------------------------------------------------------------
        Case ("By Report Name (from string path)")
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text1

                'Bind the report to the viewer
                CrystalReportViewer1.ReportSource = "C:\Crystal\CRNET\crnet_web_vbnet_SimplePreviewReport\World Sales Report.rpt"


                '--------------------------------------------------------------------
            Case "By Report Object (from string path)"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text2

                'create a new instance of the reportdocument object and load the report
                Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
                crReportDocument.Load("C:\Crystal\CRNET\crnet_web_vbnet_SimplePreviewReport\World Sales Report.rpt")

                'Bind the report to the viewer
                CrystalReportViewer1.ReportSource = crReportDocument


                '--------------------------------------------------------------------
            Case "By Untyped Report Component (from string path)"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text3

                'load the report into the report document component and set the component as the reportsource for the viewer
                reportDocument1.Load("C:\Crystal\CRNET\crnet_web_vbnet_SimplePreviewReport\World Sales Report.rpt")
                CrystalReportViewer1.ReportSource = reportDocument1


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


                '--------------------------------------------------------------------
            Case "By Cached Report Component"
                ' Display text explaining this method of binding a report to the viewer
                TextBox1.Text = text6

                'Bind the strongly typed cached report component to the viewer
                CrystalReportViewer1.ReportSource = cachedWorld_Sales_Report2

        End Select
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoadReportBySelectedMethod()
    End Sub

    Private Sub DropDownList1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged

    End Sub
End Class
