Imports System
Imports System.IO

Public Module ConversionUtilities

    Public Function ByteArrayToDecimal(ByVal p_bArray As Byte()) As Decimal

        Dim sValue As String = Nothing
        Dim fValue As Decimal

        For i As Integer = 0 To p_bArray.Length - 1
            sValue += Chr(p_bArray(i))
        Next i

        Try
            fValue = CDec(sValue)
            Return fValue
        Catch ex As Exception
            Dim strException As String
            strException = "Error in ConversionUtilites.ByteArrayToDecimal: " & ex.Message
            'MessageBox.Show(ex.Message, "Error in ConversionUtilities.ByteArrayToDecimal", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Throw New System.Exception(strException)
        End Try

    End Function


End Module
