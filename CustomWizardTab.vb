Public Class CustomWizardTab
    Inherits CustomMessageBox

    Private m_drAnswer As Windows.Forms.DialogResult

    Public ReadOnly Property Answer() As Windows.Forms.DialogResult
        Get
            Return m_drAnswer
        End Get
    End Property

    Public Sub New()
        MyBase.New()
    End Sub

    Public Overloads Function ShowDialog() As Windows.Forms.DialogResult

        m_drAnswer = MyBase.ShowDialog()

        Return m_drAnswer

    End Function

End Class
