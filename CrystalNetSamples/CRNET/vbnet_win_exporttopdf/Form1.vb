'********************************************
'Created:       January 17, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               export a report to a Portable Document Format (PDF) file.
'********************************************

''Both of these CR assemblies are required to be able to
''access, load, and export a report.
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form1
    Inherits System.Windows.Forms.Form

    ''CR Variables
    Dim crReportDocument As ReportDocument
    Dim crExportOptions As ExportOptions
    Dim crDiskFileDestinationOptions As DiskFileDestinationOptions

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(64, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(160, 120)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Export report to Portable Document Format (PDF)"
        '
        'Form1
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(10, 25)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1})
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ''The path/location where the exported file will be saved
        Dim exportFilePath As String = "c:\exported.pdf"

        ''Create an instance of the untyped report object
        crReportDocument = New ReportDocument()
        ''Load the report from disk
        crReportDocument.Load("C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Reports\Feature Examples\Chart.rpt")

        ''Set the options for saving the exported file to disk
        crDiskFileDestinationOptions = New DiskFileDestinationOptions()
        crDiskFileDestinationOptions.DiskFileName = exportFilePath

        ''Set the exporting information
        crExportOptions = crReportDocument.ExportOptions
        With crExportOptions
            .DestinationOptions = crDiskFileDestinationOptions
            .ExportDestinationType = ExportDestinationType.DiskFile
            .ExportFormatType = ExportFormatType.PortableDocFormat
        End With

        ''Export the report
        crReportDocument.Export()

        ''Display a message letting the user know the export is complete
        MessageBox.Show("Report exported to: '" & exportFilePath & "'")

    End Sub
End Class
