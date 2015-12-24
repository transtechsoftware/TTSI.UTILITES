Imports System
Imports System.Collections

'A CustomWizard is a collection of CustomWizardTabs.  Makes sense right?

Public Class CustomWizard
    Implements IEnumerable

    Private m_oElements As ArrayList

    Public Sub New()
        m_oElements = New ArrayList
    End Sub

    Public Sub AddNew(ByVal p_oTab As CustomWizardTab)
        Dim oTab As New CustomWizardTab
        oTab.Properties.Header.Text = p_oTab.Properties.Header.Text
        oTab.Properties.Body.Text = p_oTab.Properties.Body.Text
        oTab.Properties.Image.PictureToDisplay(p_oTab.Properties.Image.Picture)
        m_oElements.Add(oTab)
    End Sub

    Public Sub Add(ByVal p_oTab As CustomWizardTab)
        m_oElements.Add(p_oTab)
    End Sub

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return New ScanRecordEnumerator(Me)
    End Function

    Private Class ScanRecordEnumerator
        Implements IEnumerator

        Private m_iPos As Integer = -1
        Private m_oCustomWizard As CustomWizard

        Public Sub New(ByVal p_objCustomWizard As CustomWizard)
            Me.m_oCustomWizard = p_objCustomWizard
        End Sub

        Public ReadOnly Property Current() As Object Implements System.Collections.IEnumerator.Current
            Get
                Return m_oCustomWizard.m_oElements.Item(m_iPos)
            End Get
        End Property

        Public Function MoveNext() As Boolean Implements System.Collections.IEnumerator.MoveNext
            If (m_iPos < m_oCustomWizard.m_oElements.Count - 1) Then
                m_iPos += 1
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub Reset() Implements System.Collections.IEnumerator.Reset
            m_iPos = -1
        End Sub
    End Class

End Class

