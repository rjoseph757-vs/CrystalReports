Public Class frmRecSelForm
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents lblRecSelForm As System.Windows.Forms.Label
    Friend WithEvents txtRecSelForm As System.Windows.Forms.TextBox
    Friend WithEvents btnRecSelForm As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblRecSelForm = New System.Windows.Forms.Label()
        Me.txtRecSelForm = New System.Windows.Forms.TextBox()
        Me.btnRecSelForm = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblRecSelForm
        '
        Me.lblRecSelForm.Location = New System.Drawing.Point(16, 8)
        Me.lblRecSelForm.Name = "lblRecSelForm"
        Me.lblRecSelForm.Size = New System.Drawing.Size(168, 16)
        Me.lblRecSelForm.TabIndex = 0
        Me.lblRecSelForm.Text = "Enter Record Selection Formula:"
        '
        'txtRecSelForm
        '
        Me.txtRecSelForm.Location = New System.Drawing.Point(8, 40)
        Me.txtRecSelForm.Multiline = True
        Me.txtRecSelForm.Name = "txtRecSelForm"
        Me.txtRecSelForm.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtRecSelForm.Size = New System.Drawing.Size(272, 72)
        Me.txtRecSelForm.TabIndex = 1
        Me.txtRecSelForm.Text = "{table.field} = ""value"""
        '
        'btnRecSelForm
        '
        Me.btnRecSelForm.Location = New System.Drawing.Point(8, 120)
        Me.btnRecSelForm.Name = "btnRecSelForm"
        Me.btnRecSelForm.Size = New System.Drawing.Size(272, 23)
        Me.btnRecSelForm.TabIndex = 2
        Me.btnRecSelForm.Text = "Set Selection Formula"
        '
        'frmRecSelForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(296, 149)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnRecSelForm, Me.txtRecSelForm, Me.lblRecSelForm})
        Me.Name = "frmRecSelForm"
        Me.Text = "frmRecSelForm"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmRecSelForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If strRecSelForm <> "" Then
            txtRecSelForm.Text = strRecSelForm
        End If
        txtRecSelForm.Focus()
        'txtRecSelForm.Text
    End Sub

    Private Sub btnRecSelForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRecSelForm.Click
        strRecSelForm = txtRecSelForm.Text
        Me.Close()
    End Sub
End Class
