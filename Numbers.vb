Public Module Numbers

    Public Function DecimalPortion(ByVal p_fNum As Decimal) As Decimal

        Dim lPos As Long

        lPos = InStr(1, CStr(p_fNum), ".")

        If lPos > 0 Then
            DecimalPortion = Val(Mid(p_fNum, lPos))
        Else
            DecimalPortion = 0
        End If

    End Function

    Public Function IntegerPortion(ByVal p_fNum As Decimal) As Integer

        IntegerPortion = p_fNum - DecimalPortion(p_fNum)

    End Function

End Module
