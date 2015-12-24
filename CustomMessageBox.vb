Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
'Imports System.Drawing.Imaging
Imports System.Text

Public Class CustomMessageBox
    Inherits CustomMessageBoxBase

    Public Overloads Function Show(ByVal p_sBodyText As String, Optional ByVal p_sHeaderText As String = "MESSAGE", Optional ByVal p_oButtons As System.Windows.Forms.MessageBoxButtons = Windows.Forms.MessageBoxButtons.OK, Optional ByVal p_oIcon As System.Windows.Forms.MessageBoxIcon = Windows.Forms.MessageBoxIcon.None, Optional ByVal p_oDefaultButton As System.Windows.Forms.MessageBoxDefaultButton = Windows.Forms.MessageBoxDefaultButton.Button1, Optional ByVal p_sButton1Text As String = Nothing, Optional ByVal p_sButton2Text As String = Nothing, Optional ByVal p_sButton3Text As String = Nothing) As System.Windows.Forms.DialogResult

        Dim drReturn As System.Windows.Forms.DialogResult

        With Properties

            .Header.Text = p_sHeaderText
            .Body.Text = p_sBodyText

            ' Decide which Buttons to display
            Select Case p_oButtons
                Case Windows.Forms.MessageBoxButtons.AbortRetryIgnore
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oThirdButtonProperties.Visible = True
                    .Button3.Visible = True
                    'm_oFirstButtonProperties.Text = "Abort"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "Abort"
                    Else
                        .Button1.Text = p_sButton1Text
                    End If
                    'm_oSecondButtonProperties.Text = "Retry"
                    If p_sButton2Text Is Nothing Then
                        .Button2.Text = "Retry"
                    Else
                        .Button2.Text = p_sButton2Text
                    End If
                    'm_oThirdButtonProperties.Text = "Ignore"
                    If p_sButton3Text Is Nothing Then
                        .Button3.Text = "Ignore"
                    Else
                        .Button3.Text = p_sButton3Text
                    End If
                Case Windows.Forms.MessageBoxButtons.OK
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = False
                    .Button2.Visible = False
                    'm_oThirdButtonProperties.Visible = False
                    .Button3.Visible = False
                    'm_oFirstButtonProperties.Text = "OK"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "OK"
                    Else
                        .Button1.Text = p_sButton2Text
                    End If
                Case Windows.Forms.MessageBoxButtons.OKCancel
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = True
                    .Button2.Visible = True
                    'm_oThirdButtonProperties.Visible = False
                    .Button3.Visible = False
                    'm_oFirstButtonProperties.Text = "OK"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "OK"
                    Else
                        .Button1.Text = p_sButton1Text
                    End If
                    'm_oSecondButtonProperties.Text = "Cancel"
                    If p_sButton2Text Is Nothing Then
                        .Button2.Text = "Cancel"
                    Else
                        .Button2.Text = p_sButton2Text
                    End If
                Case Windows.Forms.MessageBoxButtons.RetryCancel
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = True
                    .Button2.Visible = True
                    'm_oThirdButtonProperties.Visible = False
                    .Button3.Visible = False
                    'm_oFirstButtonProperties.Text = "Retry"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "Retry"
                    Else
                        .Button1.Text = p_sButton1Text
                    End If
                    'm_oSecondButtonProperties.Text = "Cancel"
                    If p_sButton2Text Is Nothing Then
                        .Button2.Text = "Cancel"
                    Else
                        .Button2.Text = p_sButton2Text
                    End If
                Case Windows.Forms.MessageBoxButtons.YesNo
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = True
                    .Button2.Visible = True
                    'm_oThirdButtonProperties.Visible = False
                    .Button3.Visible = False
                    'm_oFirstButtonProperties.Text = "Yes"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "Yes"
                    Else
                        .Button1.Text = p_sButton1Text
                    End If
                    'm_oSecondButtonProperties.Text = "No"
                    If p_sButton2Text Is Nothing Then
                        .Button2.Text = "No"
                    Else
                        .Button2.Text = p_sButton2Text
                    End If
                Case Windows.Forms.MessageBoxButtons.YesNoCancel
                    'm_oFirstButtonProperties.Visible = True
                    .Button1.Visible = True
                    'm_oSecondButtonProperties.Visible = True
                    .Button2.Visible = True
                    'm_oThirdButtonProperties.Visible = True
                    .Button3.Visible = True
                    'm_oFirstButtonProperties.Text = "Yes"
                    If p_sButton1Text Is Nothing Then
                        .Button1.Text = "Yes"
                    Else
                        .Button1.Text = p_sButton1Text
                    End If
                    'm_oSecondButtonProperties.Text = "No"
                    If p_sButton2Text Is Nothing Then
                        .Button2.Text = "No"
                    Else
                        .Button2.Text = p_sButton2Text
                    End If
                    'm_oThirdButtonProperties.Text = "Cancel"
                    If p_sButton3Text Is Nothing Then
                        .Button3.Text = "Cancel"
                    Else
                        .Button3.Text = p_sButton3Text
                    End If
            End Select

            'Decide which button to set as the Default Button
            Select Case p_oDefaultButton
                Case Windows.Forms.MessageBoxDefaultButton.Button1
                    'm_oFirstButtonProperties.IsDefault = True
                    .Button1.IsDefault = True
                    'm_oSecondButtonProperties.IsDefault = False
                    .Button2.IsDefault = False
                    'm_oThirdButtonProperties.IsDefault = False
                    .Button3.IsDefault = False
                Case Windows.Forms.MessageBoxDefaultButton.Button2
                    'm_oFirstButtonProperties.IsDefault = False
                    .Button1.IsDefault = False
                    'm_oSecondButtonProperties.IsDefault = True
                    .Button2.IsDefault = True
                    'm_oThirdButtonProperties.IsDefault = False
                    .Button3.IsDefault = False
                Case Windows.Forms.MessageBoxDefaultButton.Button3
                    'm_oFirstButtonProperties.IsDefault = False
                    .Button1.IsDefault = False
                    'm_oSecondButtonProperties.IsDefault = False
                    .Button2.IsDefault = False
                    'm_oThirdButtonProperties.IsDefault = True
                    .Button3.IsDefault = True
            End Select

            'Decide which Image is going to be displayed.  This will also set the default header text.
            Select Case p_oIcon
                Case Windows.Forms.MessageBoxIcon.Asterisk
                    'x.pbPicture.Image = Image.FromFile(m_sAsteriskImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Asterisk)
                    'm_sHeaderText = "Read the Prompt."
                    .Form.Name = "Read the Prompt"
                Case Windows.Forms.MessageBoxIcon.Error
                    'x.pbPicture.Image = Image.FromFile(m_sErrorImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Erroneous)
                    'm_sHeaderText = "Error:  Read the Prompt."
                    .Form.Name = "ERROR:  Read the Prompt."
                Case Windows.Forms.MessageBoxIcon.Exclamation
                    'x.pbPicture.Image = Image.FromFile(m_sExclamationImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Exclamation)
                    'm_sHeaderText = "Warning:  Read the Prompt."
                    .Form.Name = "WARNING:  Read the Prompt."
                Case Windows.Forms.MessageBoxIcon.Hand
                    'x.pbPicture.Image = Image.FromFile(m_sHandImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Hand)
                    'm_sHeaderText = "Stop:  Read the Prompt."
                    .Form.Name = "STOP: Read the Prompt."
                Case Windows.Forms.MessageBoxIcon.Information
                    'x.pbPicture.Image = Image.FromFile(m_sInformationImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Information)
                    'm_sHeaderText = "Follow the Instructions."
                    .Form.Name = "Follow the Instructions."
                Case Windows.Forms.MessageBoxIcon.None
                    'x.pbPicture.Image = Image.FromFile(m_sNoneImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.None)
                    'm_sHeaderText = String.Empty
                    .Form.Name = String.Empty
                Case Windows.Forms.MessageBoxIcon.Question
                    'x.pbPicture.Image = Image.FromFile(m_sQuestionImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Question)
                    'm_sHeaderText = "Answer the Question"
                    .Form.Name = "Answer the Question."
                Case Windows.Forms.MessageBoxIcon.Stop
                    'x.pbPicture.Image = Image.FromFile(m_sStopImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Halt)
                    'm_sHeaderText = "Stop:  Read the Prompt"
                    .Form.Name = "STOP:  Read the Prompt."
                Case Windows.Forms.MessageBoxIcon.Warning
                    'x.pbPicture.Image = Image.FromFile(m_sWarningImagePath)
                    .Image.PictureToDisplay(ImageProperties.DefaultImages.Warning)
                    'm_sHeaderText = "Warning:  Read the Prompt"
                    .Form.Name = "WARNING:  Read the Prompt."
            End Select

            ''Adjust image properties:  This was Karina's Test, but she is not sure why she did it this way.  I think most of this will be 
            ''irrelevant now that I've changed the way the Image is assigned a picture.  The problem we are trying to avoid is an image control with
            ''no picture assigned.  Things can get a little confusing between a picture file and image control.
            'If (Not Picture Is Nothing) And (Icon = Windows.Forms.MessageBoxIcon.None) Then
            '    x.pbPicture.Image = Picture
            '    If (Picture.Size.Width > x.pbPicture.Size.Width) Or (Picture.Size.Height > x.pbPicture.Size.Height) Then
            '        x.pbPicture.ScaleImage = Infragistics.Win.ScaleImage.Always
            '    End If
            'Else
            '    If (x.pbPicture.Image.Size.Width > x.pbPicture.Size.Width) Or (x.pbPicture.Image.Size.Height > x.pbPicture.Size.Height) Then
            '        x.pbPicture.ScaleImage = Infragistics.Win.ScaleImage.Always
            '    End If
            'End If

            ''Set Properties for Body Text
            'x.lbBody.Text = BodyText
            ''x.lbBody.ForeColor = m_sdcBodyColor
            'x.lbBody.ForeColor = .Body.Color
            ''x.lbBody.Font = m_sdfBodyFont
            'x.lbBody.Font = .Body.Font
            ''x.lbBody.TextAlign = ContentAlignment.MiddleLeft
            'x.lbBody.TextAlign = .Body.TextAlign = ContentAlignment.MiddleLeft

            ''Set Properties for Header Text
            ''x.lbHeader.Text = m_sHeaderText
            'x.lbHeader.Text = .Header.Text
            ''x.lbHeader.ForeColor = m_sdcHeaderColor
            'x.lbHeader.ForeColor = .Header.Color
            ''x.lbHeader.Font = m_sdfHeaderFont
            'x.lbHeader.Font = .Header.Font

            ''Set Properties for Footer Text
            ''x.lbFooter.Text = m_sFooterText
            'x.lbFooter.Text = .Footer.Text
            ''x.lbFooter.ForeColor = m_sdcFooterColor
            'x.lbFooter.ForeColor = .Footer.Color
            ''x.lbFooter.Font = m_sdfFooterFont
            'x.lbFooter.Font = .Footer.Font
            ''Display the dialog in Modal mode
        End With

        Dim x As New TTSIForm
        PrepareForm(x)
        x.ShowDialog()

        'Return the appropriate DialogResult
        Select Case x.ButtonSelected()
            Case 1
                Select Case p_oButtons
                    Case Windows.Forms.MessageBoxButtons.AbortRetryIgnore
                        drReturn = (Windows.Forms.DialogResult.Abort)
                    Case Windows.Forms.MessageBoxButtons.OK
                        drReturn = Windows.Forms.DialogResult.OK
                    Case Windows.Forms.MessageBoxButtons.OKCancel
                        drReturn = Windows.Forms.DialogResult.OK
                    Case Windows.Forms.MessageBoxButtons.RetryCancel
                        drReturn = Windows.Forms.DialogResult.Retry
                    Case Windows.Forms.MessageBoxButtons.YesNo
                        drReturn = Windows.Forms.DialogResult.Yes
                    Case Windows.Forms.MessageBoxButtons.YesNoCancel
                        drReturn = Windows.Forms.DialogResult.Yes
                End Select
            Case 2
                Select Case p_oButtons
                    Case Windows.Forms.MessageBoxButtons.AbortRetryIgnore
                        drReturn = Windows.Forms.DialogResult.Retry
                    Case Windows.Forms.MessageBoxButtons.OKCancel
                        drReturn = Windows.Forms.DialogResult.Cancel
                    Case Windows.Forms.MessageBoxButtons.RetryCancel
                        drReturn = Windows.Forms.DialogResult.Cancel
                    Case Windows.Forms.MessageBoxButtons.YesNo
                        drReturn = Windows.Forms.DialogResult.No
                    Case Windows.Forms.MessageBoxButtons.YesNoCancel
                        drReturn = Windows.Forms.DialogResult.No
                End Select
            Case 3
                Select Case p_oButtons
                    Case Windows.Forms.MessageBoxButtons.AbortRetryIgnore
                        drReturn = Windows.Forms.DialogResult.Ignore
                    Case Windows.Forms.MessageBoxButtons.YesNoCancel
                        drReturn = Windows.Forms.DialogResult.Cancel
                End Select
        End Select

        Return drReturn

    End Function

    Public Function ShowPrompt() As Integer

        Dim oForm As New TTSIForm

        PrepareForm(oForm)

        oForm.ShowDialog()

        Return oForm.ButtonSelected()

    End Function

    Public Overridable Function ShowPrompt(ByVal p_oOwner As System.Windows.Forms.IWin32Window) As Integer

        Dim oForm As New TTSIForm

        PrepareForm(oForm)

        oForm.ShowDialog(p_oOwner)

        Return oForm.ButtonSelected()

    End Function

    Public Sub New()

        MyBase.New()

    End Sub

End Class
