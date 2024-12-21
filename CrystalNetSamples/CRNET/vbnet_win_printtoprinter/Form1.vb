'********************************************
'Created:       January 17, 2002
'Author ID:     JAS
'Purpose:       This Visual Basic .NET sample Windows application demonstrates how to '               print a report directly to a printer at runtime using the engine object model.
'********************************************


''This CR assembly is required to be able to
''access, load, and print a report.
Imports CrystalDecisions.CrystalReports.Engine

Public Class Form1
    Inherits System.Windows.Forms.Form

    ''CR Variables
    Dim crReportDocument As ReportDocument

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
    Friend WithEvents Button1 As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(72, 64)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(144, 128)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Print Report"
        '
        'Form1
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1})
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print to Printer"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ''Create an instance of a ReportDocument object.  
        ''Load a sample report file.
        crReportDocument = New ReportDocument()
        crReportDocument.Load("C:\Program Files\Microsoft Visual Studio .NET\Crystal Reports\Samples\Reports\General Business\Income Statement.rpt")

        ''Use error handling in case an error occurs
        Try
            ''Set the printer name to print the report to.  By default the sample
            ''report does not have a defult printer specified.  This will tell the
            ''engine to use the specified printer to print the report.  Print out 
            ''a test page (from Printer properties) to get the correct value.
            crReportDocument.PrintOptions.PrinterName = "\\vanprt04\P6-4050TN-TS1"

            ''Start the printing process.  Provide details of the print job
            ''using the arguments.
            crReportDocument.PrintToPrinter(1, True, 1, 1)

            ''Let the user know that the print job is completed.
            MessageBox.Show("Report finished printing!")

        Catch err As Exception
            MessageBox.Show(err.ToString())
        End Try

    End Sub
End Class
