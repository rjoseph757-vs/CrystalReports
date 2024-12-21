'********************************************
'
'File Name:     Crnet_win_vbnet_Export.sln
'Created:       May 1, 2003
'Author ID:     ALD
'Purpose:       This sample .net windows application demonstrates how to export your report 
'               to the following formats:
'
'               - Rich Text Format (RTF)
'               - Microsoft Word Format (DOC)
'               - Portable Document Format (PDF)
'               - Microsoft Excel (XLS)
'               - Crystal Report (RPT)
'               - HTML 3.2 (HTML)
'               - HTML 4.0 (DHTML)  
'
'********************************************


Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO


Public Class Form1
    Inherits System.Windows.Forms.Form

    ' CR Variables
    Dim crReportDocument As ReportDocument
    Dim crExportOptions As ExportOptions
    Dim crDiskFileDestinationOptions As DiskFileDestinationOptions

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        crReportDocument = New Chart()

        CrystalReportViewer1.ReportSource = crReportDocument

        ' *******************************************
        ' Initialize Combobox1 for Format types
        ' *******************************************
        With ComboBox1.Items
            .Add("Rich Text (RTF)")
            .Add("Portable Document (PDF)")
            .Add("MS Word (DOC)")
            .Add("MS Excel (XLS)")
            .Add("Crystal Report (RPT)")
            .Add("HTML 3.2 (HTML)")
            .Add("HTML 4.0 (HTML)")
        End With

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
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents CrystalReportViewer1 As CrystalDecisions.Windows.Forms.CrystalReportViewer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.CrystalReportViewer1 = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.SuspendLayout()
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(8, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(224, 21)
        Me.ComboBox1.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(264, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 24)
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Export Report"
        '
        'CrystalReportViewer1
        '
        Me.CrystalReportViewer1.ActiveViewIndex = -1
        Me.CrystalReportViewer1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.CrystalReportViewer1.Location = New System.Drawing.Point(16, 56)
        Me.CrystalReportViewer1.Name = "CrystalReportViewer1"
        Me.CrystalReportViewer1.ReportSource = Nothing
        Me.CrystalReportViewer1.Size = New System.Drawing.Size(480, 296)
        Me.CrystalReportViewer1.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(512, 365)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.CrystalReportViewer1, Me.Button1, Me.ComboBox1})
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ExportReport()
    End Sub

    Sub ExportReport()
        ' This subroutine uses a case statement to determine the selected export format from the Combobox1
        ' menu and then sets the appropriate export options for the selected export format.  The report is 
        ' exported to a subdirectory called "Exported".

        ' ********************************
        'Check to see if the application directory has a subdirectory called "Exported".
        'If not, create the directory since exported files will be placed here.
        'This uses the Directory class of the System.IO namespace.
        Dim ExportPath As String
        ExportPath = Application.StartupPath + "\Exported\"
        If Directory.Exists(ExportPath) = False Then
            Directory.CreateDirectory(ExportPath)
        End If
        ' ********************************


        ' First we must create a new instance of the diskfiledestinationoptions class and
        ' set variable called crExportOptions to the exportoptions class of the reportdocument.
        crDiskFileDestinationOptions = New DiskFileDestinationOptions()
        crExportOptions = crReportDocument.ExportOptions


        'Find the export type specified in the Combobox1 and export the report. The possible export format
        'types are Rich Text(RTF), Portable Document (PDF), MS Word (DOC), MS Excel (XLS), Crystal Report (RPT),
        'HTML 3.2 (HTML) and HTML 4.0 (HTML)
        '
        'Though not used in this sample application, there are options that can be specified for various format types.
        'When exporting to Rich Text, Word, or PDF, you can use the PdfRtfWordFormatOptions class to specify the
        'first page, last page or page range to be exported.
        'When exporting to Excel, you can use the ExcelFormatOptions class to specify export properties such as
        'the column width etc.

        Select Case ComboBox1.Text 'this contains the value of the selected export format.

            Case "Rich Text (RTF)"
                '--------------------------------------------------------------------
                'Export to RTF. 

                'append a filename to the export path and set this file as the filename property for
                'the DestinationOptions class
                crDiskFileDestinationOptions.DiskFileName = ExportPath + "RichTextFormat.rtf"

                'set the required report ExportOptions properties
                With crReportDocument.ExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.RichText
                    .DestinationOptions = crDiskFileDestinationOptions
                End With
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "Portable Document (PDF)"
                'Export to PDF


                'append a filename to the export path and set this file as the filename property for
                'the DestinationOptions class
                crDiskFileDestinationOptions.DiskFileName = ExportPath + "PortableDoc.pdf"

                'set the required report ExportOptions properties
                With crExportOptions
                    .DestinationOptions = crDiskFileDestinationOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.PortableDocFormat
                End With
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "MS Word (DOC)"
                'Export to Word


                'append a filename to the export path and set this file as the filename property for
                'the DestinationOptions class
                crDiskFileDestinationOptions.DiskFileName = ExportPath + "Word.doc"

                'set the required report ExportOptions properties
                With crReportDocument.ExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.WordForWindows
                    .DestinationOptions = crDiskFileDestinationOptions
                End With
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "MS Excel (XLS)"
                'Export to Excel

                'append a filename to the export path and set this file as the filename property for
                'the DestinationOptions class
                crDiskFileDestinationOptions.DiskFileName = ExportPath + "Excel.xls"

                'set the required report ExportOptions properties
                With crReportDocument.ExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.Excel
                    .DestinationOptions = crDiskFileDestinationOptions
                End With
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "Crystal Report (RPT)"
                'Export to Crystal reports:

                'append a filename to the export path and set this file as the filename property for
                'the DestinationOptions class
                crDiskFileDestinationOptions.DiskFileName = ExportPath + "Report.rpt"

                'set the required report ExportOptions properties
                With crReportDocument.ExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.CrystalReport
                    .DestinationOptions = crDiskFileDestinationOptions
                End With
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "HTML 3.2 (HTML)"
                'Export to HTML32:

                Dim HTML32Formatopts As New HTMLFormatOptions()

                With crExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.HTML32
                End With

                With HTML32Formatopts
                    .HTMLBaseFolderName = ExportPath + "Html32Folder" 'Foldername to place HTML files
                    .HTMLFileName = "HTML32.html"
                    .HTMLEnableSeparatedPages = False
                    .HTMLHasPageNavigator = False
                End With

                crExportOptions.FormatOptions = HTML32Formatopts
                '--------------------------------------------------------------------


                '--------------------------------------------------------------------
            Case "HTML 4.0 (HTML)"
                'Export to Html 4.0:

                Dim HTML40Formatopts As New HTMLFormatOptions()

                With crExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile
                    .ExportFormatType = ExportFormatType.HTML40
                End With

                With HTML40Formatopts
                    .HTMLBaseFolderName = ExportPath + "Html40Folder" ' Foldername to place HTML files
                    .HTMLFileName = "HTML40.html"
                    .HTMLEnableSeparatedPages = True
                    .HTMLHasPageNavigator = True
                    .FirstPageNumber = 1
                    .LastPageNumber = 3
                End With

                crExportOptions.FormatOptions = HTML40Formatopts


        End Select 'export format

        'Once the export options have been set for the report, the report can be exported. The Export command
        'does not take any arguments
        Try
            ' Export the report
            crReportDocument.Export()
            MsgBox("Report exported successfully.")
        Catch err As Exception
            MsgBox(err.Message.ToString)
        End Try

    End Sub
End Class
