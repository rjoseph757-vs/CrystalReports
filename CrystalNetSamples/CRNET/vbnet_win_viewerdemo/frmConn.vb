Public Class frmConn
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
    Friend WithEvents txtServer As System.Windows.Forms.TextBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents txtUserID As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents btnSetConn As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtServer = New System.Windows.Forms.TextBox()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.txtUserID = New System.Windows.Forms.TextBox()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btnSetConn = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtServer
        '
        Me.txtServer.Location = New System.Drawing.Point(112, 8)
        Me.txtServer.Name = "txtServer"
        Me.txtServer.Size = New System.Drawing.Size(136, 20)
        Me.txtServer.TabIndex = 0
        Me.txtServer.Text = ""
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(112, 32)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(136, 20)
        Me.txtDatabase.TabIndex = 1
        Me.txtDatabase.Text = ""
        '
        'txtUserID
        '
        Me.txtUserID.Location = New System.Drawing.Point(112, 56)
        Me.txtUserID.Name = "txtUserID"
        Me.txtUserID.Size = New System.Drawing.Size(136, 20)
        Me.txtUserID.TabIndex = 2
        Me.txtUserID.Text = ""
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(112, 80)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.Size = New System.Drawing.Size(136, 20)
        Me.txtPassword.TabIndex = 3
        Me.txtPassword.Text = ""
        '
        'btnSetConn
        '
        Me.btnSetConn.Location = New System.Drawing.Point(8, 112)
        Me.btnSetConn.Name = "btnSetConn"
        Me.btnSetConn.Size = New System.Drawing.Size(240, 23)
        Me.btnSetConn.TabIndex = 4
        Me.btnSetConn.Text = "Set Connection Information"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Server Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Database:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "User ID:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Password:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmConn
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(264, 141)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.btnSetConn, Me.txtPassword, Me.txtUserID, Me.txtDatabase, Me.txtServer})
        Me.Name = "frmConn"
        Me.Text = "frmConn"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnSetConn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetConn.Click
        strServer = txtServer.Text
        strDatabase = txtDatabase.Text
        strUserID = txtUserID.Text
        strPassword = txtPassword.Text
        Me.Close()
    End Sub

    Private Sub frmConn_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtServer.Text = strServer
        txtDatabase.Text = strDatabase
        txtUserID.Text = strUserID
        txtPassword.Focus()
    End Sub
End Class
