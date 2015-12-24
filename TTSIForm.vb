Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text
Imports TTSI.MEDIA

Public Class TTSIForm
    Inherits System.Windows.Forms.Form
    Private m_iSelected As Integer

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
    Friend WithEvents btnFirst As System.Windows.Forms.Button
    Friend WithEvents btnSecond As System.Windows.Forms.Button
    Friend WithEvents btnThird As System.Windows.Forms.Button
    Friend WithEvents gbFooter As System.Windows.Forms.GroupBox
    Friend WithEvents gbButtons As System.Windows.Forms.GroupBox
    Friend WithEvents gbHeader As System.Windows.Forms.GroupBox
    Friend WithEvents bgBody As System.Windows.Forms.GroupBox
    Friend WithEvents lbHeader As System.Windows.Forms.Label
    Friend WithEvents lbFooter As System.Windows.Forms.Label
    Friend WithEvents lbBody As System.Windows.Forms.Label
    Friend WithEvents pbPicture As Infragistics.Win.UltraWinEditors.UltraPictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.gbFooter = New System.Windows.Forms.GroupBox
        Me.lbFooter = New System.Windows.Forms.Label
        Me.btnThird = New System.Windows.Forms.Button
        Me.btnSecond = New System.Windows.Forms.Button
        Me.btnFirst = New System.Windows.Forms.Button
        Me.gbButtons = New System.Windows.Forms.GroupBox
        Me.gbHeader = New System.Windows.Forms.GroupBox
        Me.lbHeader = New System.Windows.Forms.Label
        Me.bgBody = New System.Windows.Forms.GroupBox
        Me.lbBody = New System.Windows.Forms.Label
        Me.pbPicture = New Infragistics.Win.UltraWinEditors.UltraPictureBox
        Me.gbFooter.SuspendLayout()
        Me.gbButtons.SuspendLayout()
        Me.gbHeader.SuspendLayout()
        Me.bgBody.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbFooter
        '
        Me.gbFooter.Controls.Add(Me.lbFooter)
        Me.gbFooter.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.gbFooter.Location = New System.Drawing.Point(0, 282)
        Me.gbFooter.Name = "gbFooter"
        Me.gbFooter.Size = New System.Drawing.Size(559, 30)
        Me.gbFooter.TabIndex = 3
        Me.gbFooter.TabStop = False
        '
        'lbFooter
        '
        Me.lbFooter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbFooter.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lbFooter.Location = New System.Drawing.Point(3, 16)
        Me.lbFooter.Name = "lbFooter"
        Me.lbFooter.Size = New System.Drawing.Size(553, 11)
        Me.lbFooter.TabIndex = 0
        '
        'btnThird
        '
        Me.btnThird.Location = New System.Drawing.Point(218, 20)
        Me.btnThird.Name = "btnThird"
        Me.btnThird.Size = New System.Drawing.Size(100, 20)
        Me.btnThird.TabIndex = 4
        Me.btnThird.Text = "Button3"
        '
        'btnSecond
        '
        Me.btnSecond.Location = New System.Drawing.Point(112, 20)
        Me.btnSecond.Name = "btnSecond"
        Me.btnSecond.Size = New System.Drawing.Size(100, 20)
        Me.btnSecond.TabIndex = 3
        Me.btnSecond.Text = "Button2"
        '
        'btnFirst
        '
        Me.btnFirst.Location = New System.Drawing.Point(7, 20)
        Me.btnFirst.Name = "btnFirst"
        Me.btnFirst.Size = New System.Drawing.Size(100, 20)
        Me.btnFirst.TabIndex = 2
        Me.btnFirst.Text = "Button1"
        '
        'gbButtons
        '
        Me.gbButtons.Controls.Add(Me.btnFirst)
        Me.gbButtons.Controls.Add(Me.btnThird)
        Me.gbButtons.Controls.Add(Me.btnSecond)
        Me.gbButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.gbButtons.Location = New System.Drawing.Point(0, 312)
        Me.gbButtons.Name = "gbButtons"
        Me.gbButtons.Size = New System.Drawing.Size(559, 52)
        Me.gbButtons.TabIndex = 2
        Me.gbButtons.TabStop = False
        '
        'gbHeader
        '
        Me.gbHeader.Controls.Add(Me.lbHeader)
        Me.gbHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.gbHeader.Location = New System.Drawing.Point(0, 0)
        Me.gbHeader.Name = "gbHeader"
        Me.gbHeader.Size = New System.Drawing.Size(559, 73)
        Me.gbHeader.TabIndex = 4
        Me.gbHeader.TabStop = False
        '
        'lbHeader
        '
        Me.lbHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbHeader.Location = New System.Drawing.Point(3, 16)
        Me.lbHeader.Name = "lbHeader"
        Me.lbHeader.Size = New System.Drawing.Size(553, 54)
        Me.lbHeader.TabIndex = 0
        '
        'bgBody
        '
        Me.bgBody.Controls.Add(Me.lbBody)
        Me.bgBody.Controls.Add(Me.pbPicture)
        Me.bgBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.bgBody.Location = New System.Drawing.Point(0, 73)
        Me.bgBody.Name = "bgBody"
        Me.bgBody.Size = New System.Drawing.Size(559, 209)
        Me.bgBody.TabIndex = 5
        Me.bgBody.TabStop = False
        '
        'lbBody
        '
        Me.lbBody.BackColor = System.Drawing.SystemColors.Control
        Me.lbBody.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lbBody.Location = New System.Drawing.Point(170, 16)
        Me.lbBody.Name = "lbBody"
        Me.lbBody.Size = New System.Drawing.Size(386, 190)
        Me.lbBody.TabIndex = 5
        '
        'pbPicture
        '
        Me.pbPicture.BorderShadowColor = System.Drawing.Color.Empty
        Me.pbPicture.Dock = System.Windows.Forms.DockStyle.Left
        Me.pbPicture.Location = New System.Drawing.Point(3, 16)
        Me.pbPicture.Name = "pbPicture"
        Me.pbPicture.Size = New System.Drawing.Size(167, 190)
        Me.pbPicture.TabIndex = 4
        '
        'TTSIForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(559, 364)
        Me.Controls.Add(Me.bgBody)
        Me.Controls.Add(Me.gbHeader)
        Me.Controls.Add(Me.gbFooter)
        Me.Controls.Add(Me.gbButtons)
        Me.Name = "TTSIForm"
        Me.Text = "TTSI Form"
        Me.gbFooter.ResumeLayout(False)
        Me.gbButtons.ResumeLayout(False)
        Me.gbHeader.ResumeLayout(False)
        Me.bgBody.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub UseTheMouse()
        Sound.PlayWaveFile("Use-the-Mouse.wav")
    End Sub

    Public ReadOnly Property ButtonSelected() As Integer
        Get
            Return m_iSelected
        End Get
    End Property

    Private Sub btnFirst_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnFirst.MouseDown
        m_iSelected = 1
        Me.Close()
    End Sub

    Private Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        UseTheMouse()
    End Sub

    Private Sub btnSecond_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnSecond.MouseDown
        m_iSelected = 2
        Me.Close()
    End Sub

    Private Sub btnSecond_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSecond.Click
        UseTheMouse()
    End Sub

    Private Sub btnThird_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnThird.MouseDown
        m_iSelected = 3
        Me.Close()
    End Sub

    Private Sub btnThird_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnThird.Click
        UseTheMouse()
    End Sub

End Class
