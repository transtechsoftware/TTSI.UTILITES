Option Explicit On 

Imports System.Drawing

Public Class ButtonProperties

    Public Visible As Boolean
    Public Text As String
    Public IsDefault As Boolean

    Public Sub New()
        Visible = True
        Text = "Button"
        IsDefault = False
    End Sub

End Class

Public Class FormProperties

    Public Height As Integer
    Public Width As Integer
    Public Name As String
    Public X As Integer
    Public Y As Integer

    Public Sub New()
        Height = 480
        Width = 640
        Name = "IMPORTANT MESSAGE"
        X = 25
        Y = 25
    End Sub

End Class

Public Class TextProperties

    Public Text As String
    Public Color As System.Drawing.Color
    Public TextAlign As System.Drawing.ContentAlignment

    Private m_oFont As System.Drawing.Font

    Public Property Font() As System.Drawing.Font
        Get
            Return m_oFont
        End Get
        Set(ByVal Value As System.Drawing.Font)
            'If Not m_oFont Is Nothing Then
            '    m_oFont = Nothing
            'End If
            m_oFont = Value
        End Set
    End Property

    Public Sub New()

        Text = "Uninitialized"
        Color = Color.Black
        TextAlign = ContentAlignment.MiddleCenter
        m_oFont = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

    End Sub

End Class

Public Class ImageProperties

    Public Enum DefaultImages
        Asterisk = 1
        Erroneous = 2
        Exclamation = 3
        Hand = 4
        Information = 5
        None = 6
        Question = 7
        Halt = 8
        Warning = 9
    End Enum

    Private m_oPictureToDisplay As Image
    Private m_bSetToNone As Boolean

    Public Property Picture() As Image
        Get
            Return m_oPictureToDisplay
        End Get
        Set(ByVal Value As Image)
            m_oPictureToDisplay = Value
        End Set
    End Property

    Private m_sAsteriskImagePath As String
    Private m_sErrorImagePath As String
    Private m_sExclamationImagePath As String
    Private m_sHandImagePath As String
    Private m_sInformationImagePath As String
    Private m_sNoneImagePath As String
    Private m_sQuestionImagePath As String
    Private m_sStopImagePath As String
    Private m_sWarningImagePath As String

    'By Wrapping the Image Paths in Properties, we can perform file validation
    Public Property AsteriskImagePath() As String
        Get
            Return m_sAsteriskImagePath
        End Get
        Set(ByVal Value As String)
            m_sAsteriskImagePath = Value
        End Set
    End Property

    Public Property ErrorImagePath() As String
        Get
            Return m_sErrorImagePath
        End Get
        Set(ByVal Value As String)
            m_sErrorImagePath = Value
        End Set
    End Property

    Public Property ExclamationImagePath() As String
        Get
            Return m_sExclamationImagePath
        End Get
        Set(ByVal Value As String)
            m_sExclamationImagePath = Value
        End Set
    End Property

    Public Property HandImagePath() As String
        Get
            Return m_sHandImagePath
        End Get
        Set(ByVal Value As String)
            m_sHandImagePath = Value
        End Set
    End Property

    Public Property InformationImagePath() As String
        Get
            Return m_sInformationImagePath
        End Get
        Set(ByVal Value As String)
            m_sInformationImagePath = Value
        End Set
    End Property

    Public Property NoneImagePath() As String
        Get
            Return m_sNoneImagePath
        End Get
        Set(ByVal Value As String)
            m_sNoneImagePath = Value
        End Set
    End Property

    Public Property QuestionImagePath() As String
        Get
            Return m_sQuestionImagePath
        End Get
        Set(ByVal Value As String)
            m_sQuestionImagePath = Value
        End Set
    End Property

    Public Property StopImagePath() As String
        Get
            Return m_sStopImagePath
        End Get
        Set(ByVal Value As String)
            m_sStopImagePath = Value
        End Set
    End Property

    Public Property WarningImagePath() As String
        Get
            Return m_sWarningImagePath
        End Get
        Set(ByVal Value As String)
            m_sWarningImagePath = Value
        End Set
    End Property

    Public ReadOnly Property SetToNone() As Boolean
        Get
            Return m_bSetToNone
        End Get
    End Property

    'This is the preferred method of assigning a default picture to the Image control
    Public Function PictureToDisplay(ByVal p_eDefaultImage As ImageProperties.DefaultImages) As Boolean

        Try

            m_bSetToNone = False

            Select Case p_eDefaultImage
                Case DefaultImages.Asterisk
                    Picture = Image.FromFile(AsteriskImagePath)
                Case DefaultImages.Erroneous
                    Picture = Image.FromFile(ErrorImagePath)
                Case DefaultImages.Exclamation
                    Picture = Image.FromFile(ExclamationImagePath)
                Case DefaultImages.Hand
                    Picture = Image.FromFile(HandImagePath)
                Case DefaultImages.Information
                    Picture = Image.FromFile(InformationImagePath)
                Case DefaultImages.Question
                    Picture = Image.FromFile(QuestionImagePath)
                Case DefaultImages.Halt
                    Picture = Image.FromFile(StopImagePath)
                Case DefaultImages.Warning
                    Picture = Image.FromFile(WarningImagePath)
                Case Else
                    Picture = Image.FromFile(NoneImagePath)
                    m_bSetToNone = True
            End Select

        Catch ex As Exception

            Picture = Nothing
            m_bSetToNone = True

        End Try

        Return Not m_bSetToNone

    End Function

    Public Function PictureToDisplay(ByVal p_sImagePath As String) As Boolean

        Try
            Picture = Image.FromFile(p_sImagePath)
            m_bSetToNone = False
        Catch ex As Exception
            Picture = Nothing
            m_bSetToNone = True
        End Try

        Return Not m_bSetToNone

    End Function

    Public Function PictureToDisplay(ByVal p_oImage As Image) As Boolean

        Try
            Picture = p_oImage
            m_bSetToNone = False
        Catch ex As Exception

            Picture = Nothing
            m_bSetToNone = True

        End Try

        Return Not m_bSetToNone

    End Function

    Public Sub New()

        m_sAsteriskImagePath = "C:\Program Files\Common Files\TransTechSoftware\AsteriskMark.png"
        m_sErrorImagePath = "C:\Program Files\Common Files\TransTechSoftware\ErrorMark.png"
        m_sExclamationImagePath = "C:\Program Files\Common Files\TransTechSoftware\ExclamationMark.png"
        m_sHandImagePath = "C:\Program Files\Common Files\TransTechSoftware\HandMark.png"
        m_sInformationImagePath = "C:\Program Files\Common Files\TransTechSoftware\InformationMark.png"
        m_sNoneImagePath = "C:\Program Files\Common Files\TransTechSoftware\NoneMark.png"
        m_sQuestionImagePath = "C:\Program Files\Common Files\TransTechSoftware\QuestionMark.png"
        m_sStopImagePath = "C:\Program Files\Common Files\TransTechSoftware\StopMark.png"
        m_sWarningImagePath = "C:\Program Files\Common Files\TransTechSoftware\WarningMark.png"

        'm_oPictureToDisplay.FromFile(NoneImagePath)
        m_oPictureToDisplay = Image.FromFile(NoneImagePath)

    End Sub

End Class

Public Class TTSIFormProperties

    Protected m_oFormProperties As FormProperties
    Protected m_oFirstButtonProperties As ButtonProperties
    Protected m_oSecondButtonProperties As ButtonProperties
    Protected m_oThirdButtonProperties As ButtonProperties
    Protected m_oBodyTextProperties As TextProperties
    Protected m_oHeaderTextProperties As TextProperties
    Protected m_oFooterTextProperties As TextProperties
    Protected m_oPicture As ImageProperties

    Public Property Form() As FormProperties
        Get
            Return m_oFormProperties
        End Get
        Set(ByVal Value As FormProperties)
            m_oFormProperties = Value
        End Set
    End Property

    Public Property Button1() As ButtonProperties
        Get
            Return m_oFirstButtonProperties
        End Get
        Set(ByVal Value As ButtonProperties)
            m_oFirstButtonProperties = Value
        End Set
    End Property

    Public Property Button2() As ButtonProperties
        Get
            Return m_oSecondButtonProperties
        End Get
        Set(ByVal Value As ButtonProperties)
            m_oSecondButtonProperties = Value
        End Set
    End Property

    Public Property Button3() As ButtonProperties
        Get
            Return m_oThirdButtonProperties
        End Get
        Set(ByVal Value As ButtonProperties)
            m_oThirdButtonProperties = Value
        End Set
    End Property

    Public Property Header() As TextProperties
        Get
            Return m_oHeaderTextProperties
        End Get
        Set(ByVal Value As TextProperties)
            m_oHeaderTextProperties = Value
        End Set
    End Property

    Public Property Footer() As TextProperties
        Get
            Return m_oFooterTextProperties
        End Get
        Set(ByVal Value As TextProperties)
            m_oFooterTextProperties = Value
        End Set
    End Property

    Public Property Body() As TextProperties
        Get
            Return m_oBodyTextProperties
        End Get
        Set(ByVal Value As TextProperties)
            m_oBodyTextProperties = Value
        End Set
    End Property

    Public Property Image() As ImageProperties
        Get
            Return m_oPicture
        End Get
        Set(ByVal Value As ImageProperties)
            m_oPicture = Value
        End Set
    End Property

    Public Sub New()

        m_oFormProperties = New FormProperties
        m_oFirstButtonProperties = New ButtonProperties
        m_oSecondButtonProperties = New ButtonProperties
        m_oThirdButtonProperties = New ButtonProperties
        m_oBodyTextProperties = New TextProperties
        m_oHeaderTextProperties = New TextProperties
        m_oFooterTextProperties = New TextProperties
        m_oPicture = New ImageProperties

    End Sub

End Class

Public Class CustomMessageBoxBase

    Private m_oForm As TTSIForm
    Private m_oFormProperties As TTSIFormProperties

    Public Property Properties() As TTSIFormProperties
        Get
            Return m_oFormProperties
        End Get
        Set(ByVal Value As TTSIFormProperties)
            m_oFormProperties = Value
        End Set
    End Property

    ''The following properties are unnecessary since consumers can access Properties directy, but are being temporarily kept for legacy reasons.
    'Public Property Height() As Integer
    '    Get
    '        Return Properties.Form.Height
    '        ''Return m_oFormProperties.Height
    '        ''m_iHeight
    '    End Get
    '    Set(ByVal Value As Integer)
    '        Properties.Form.Height = Value
    '        ''m_oFormProperties.Height = Value
    '        ''m_iHeight = Value
    '    End Set
    'End Property

    'Public Property Width() As Integer
    '    Get
    '        Return Properties.Form.Width
    '        ''Return m_iWidth
    '    End Get
    '    Set(ByVal Value As Integer)
    '        Properties.Form.Width = Value
    '        ''m_iWidth = Value
    '    End Set
    'End Property

    'Public Property BodyColor() As System.Drawing.Color
    '    Get
    '        'Return m_sdcBodyColor
    '        Return Properties.Body.Color
    '    End Get
    '    Set(ByVal Value As System.Drawing.Color)
    '        'm_sdcBodyColor = Value
    '        Properties.Body.Color = Value
    '    End Set
    'End Property

    'Public Property BodyFont() As System.Drawing.Font
    '    Get
    '        'Return m_sdfBodyFont
    '        Return Properties.Body.Font
    '    End Get
    '    Set(ByVal Value As System.Drawing.Font)
    '        'm_sdfBodyFont = Value
    '        Properties.Body.Font = Value
    '    End Set
    'End Property

    'Public Property Picture() As Image
    '    Get
    '        'Return m_bPicture
    '        Return Properties.Image.Picture
    '    End Get
    '    Set(ByVal Value As Image)
    '        'm_bPicture = Value
    '        Properties.Image.Picture = Value
    '    End Set
    'End Property

    'Public Property HeaderText() As String
    '    Get
    '        'Return m_sHeaderText
    '        Return Properties.Header.Text
    '    End Get
    '    Set(ByVal Value As String)
    '        'm_sHeaderText = Value
    '        Properties.Header.Text = Value
    '    End Set
    'End Property

    'Public Property HeaderColor() As System.Drawing.Color
    '    Get
    '        'Return m_sdcHeaderColor
    '        Return Properties.Header.Color
    '    End Get
    '    Set(ByVal Value As System.Drawing.Color)
    '        'm_sdcHeaderColor = Value
    '        Properties.Header.Color = Value
    '    End Set
    'End Property

    'Public Property HeaderFont() As System.Drawing.Font
    '    Get
    '        'Return m_sdfHeaderFont
    '        Return Properties.Header.Font
    '    End Get
    '    Set(ByVal Value As System.Drawing.Font)
    '        'm_sdfHeaderFont = Value
    '        Properties.Header.Font = Value
    '    End Set
    'End Property

    'Public Property FooterText() As String
    '    Get
    '        'Return m_sFooterText
    '        Return Properties.Footer.Text
    '    End Get
    '    Set(ByVal Value As String)
    '        'm_sFooterText = Value
    '        Properties.Footer.Text = Value
    '    End Set
    'End Property

    'Public Property FooterColor() As System.Drawing.Color
    '    Get
    '        'Return m_sdcFooterColor
    '        Return Properties.Footer.Color
    '    End Get
    '    Set(ByVal Value As System.Drawing.Color)
    '        'm_sdcFooterColor = Value
    '        Properties.Footer.Color = Value
    '    End Set
    'End Property

    'Public Property FooterFont() As System.Drawing.Font
    '    Get
    '        'Return m_sdfFooterFont
    '        Return Properties.Footer.Font
    '    End Get
    '    Set(ByVal Value As System.Drawing.Font)
    '        'm_sdfFooterFont = Value
    '        Properties.Footer.Font = Value
    '    End Set
    'End Property

    Public Sub New()

        m_oFormProperties = New TTSIFormProperties

        ''m_iHeight = 480
        ''m_iWidth = 640

        ''m_sdcBodyColor = Color.Black
        ''m_sdfBodyFont = New Font("Microsoft Sans Serif", 8, FontStyle.Regular)

        'm_sHeaderText = String.Empty
        ''m_sdcHeaderColor = Color.Black
        ''m_sdfHeaderFont = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        Properties.Header.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Bold)

        ''m_sFooterText = "You must use the mouse to click on the appropriate button."
        ''m_sdcFooterColor = Color.Black
        ''m_sdfFooterFont = New Font("Microsoft Sans Serif", 10, FontStyle.Regular)
        Properties.Footer.Text = "You must use the mouse to click on the appropriate button."
        Properties.Footer.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Regular)

    End Sub

    Protected Overridable Sub PrepareForm(ByRef p_oForm As TTSIForm)
        ' Set Height and Width of Dialog

        With Properties

            'Set FORM properties
            p_oForm.Height = .Form.Height
            p_oForm.Width = .Form.Width
            p_oForm.Text = .Form.Name
            p_oForm.Location = New Point(.Form.X, .Form.Y)

            'Set HEADER properties
            p_oForm.lbHeader.Text = .Header.Text
            p_oForm.lbHeader.ForeColor = .Header.Color
            p_oForm.lbHeader.Font = .Header.Font
            p_oForm.lbHeader.TextAlign = .Header.TextAlign

            'Set BODY properties
            p_oForm.lbBody.Text = .Body.Text
            p_oForm.lbBody.ForeColor = .Body.Color
            p_oForm.lbBody.Font = .Body.Font
            p_oForm.lbBody.TextAlign = .Body.TextAlign

            'Set FOOTER properties
            p_oForm.lbFooter.Text = .Footer.Text
            p_oForm.lbFooter.ForeColor = .Footer.Color
            p_oForm.lbFooter.Font = .Footer.Font
            p_oForm.lbFooter.TextAlign = .Footer.TextAlign

            'Set BUTTON1 properties
            p_oForm.btnFirst.Visible = .Button1.Visible
            p_oForm.btnFirst.Text = .Button1.Text

            'Set BUTTON2 properties
            p_oForm.btnSecond.Visible = .Button2.Visible
            p_oForm.btnSecond.Text = .Button2.Text

            'Set BUTTON3 properties
            p_oForm.btnThird.Visible = .Button3.Visible
            p_oForm.btnThird.Text = .Button3.Text

            'Decide which button is selected by Default
            If .Button1.IsDefault = True Then
                p_oForm.btnFirst.Select()
            ElseIf .Button2.IsDefault = True Then
                p_oForm.btnSecond.Select()
            ElseIf .Button3.IsDefault = True Then
                p_oForm.btnThird.Select()
            Else
                p_oForm.btnFirst.Select()
            End If

            'Set IMAGE properties
            If Not .Image.Picture Is Nothing Then
                p_oForm.pbPicture.Image = .Image.Picture
                If (.Image.Picture.Width > p_oForm.pbPicture.Size.Width) Or (.Image.Picture.Height > p_oForm.pbPicture.Size.Height) Then
                    p_oForm.pbPicture.ScaleImage = Infragistics.Win.ScaleImage.Always
                End If
            End If

            'If (Not .Image.Picture Is Nothing) And (.Image.SetToNone = True) Then
            '    p_oForm.pbPicture.Image = .Image.Picture
            '    If (.Image.Picture.Size.Width > p_oForm.pbPicture.Size.Width) Or (.Image.Picture.Size.Height > p_oForm.pbPicture.Size.Height) Then
            '        p_oForm.pbPicture.ScaleImage = Infragistics.Win.ScaleImage.Always
            '    End If
            'Else
            '    If (p_oForm.pbPicture.Image.Size.Width > p_oForm.pbPicture.Size.Width) Or (p_oForm.pbPicture.Image.Size.Height > p_oForm.pbPicture.Size.Height) Then
            '        p_oForm.pbPicture.ScaleImage = Infragistics.Win.ScaleImage.Always
            '    End If
            'End If

        End With

    End Sub

    Public Overridable Sub Show()

        Dim oForm As New TTSIForm

        PrepareForm(oForm)

        oForm.Show()

    End Sub

    Public Overridable Function ShowDialog() As Windows.Forms.DialogResult

        Dim oForm As New TTSIForm

        PrepareForm(oForm)

        Return oForm.ShowDialog()

    End Function

    Public Overridable Function ShowDialog(ByVal p_oOwner As System.Windows.Forms.IWin32Window) As Windows.Forms.DialogResult

        Dim oForm As New TTSIForm

        PrepareForm(oForm)

        Return oForm.ShowDialog(p_oOwner)

    End Function

End Class
