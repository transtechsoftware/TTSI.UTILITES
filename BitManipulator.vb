Public Class BitManipulator

    Private m_iMyByte As Int16

    Public Sub New(ByVal p_iByte As Int16)
        m_iMyByte = p_iByte
    End Sub

    Public ReadOnly Property ByteValue() As Short
        Get
            Return m_iMyByte
        End Get
    End Property

    'The ClearBit Sub clears the 1 based, nth bit (p_MyBit) of an integer (p_MyByte)
    Public Sub ClearBit(ByVal p_MyBit As Int16)
        Dim BitMask As Int16
        'Create a bitmask with the 2 to the nth power bit set
        BitMask = 2 ^ (p_MyBit - 1)
        'Clear the nth Bit
        m_iMyByte = m_iMyByte And Not BitMask
    End Sub

    'The ExamineBit function wil return True or False depending on the value of the 1 based, nth bit (p_MyBit) of an integer (p_MyByte)
    Function ExamineBit(ByVal p_MyBit As Int16) As Boolean
        Dim BitMask As Int16
        BitMask = 2 ^ (p_MyBit - 1)
        ExamineBit = ((m_iMyByte And BitMask) > 0)
    End Function

    'The SetBit Sub will set the 1 based, nth bit (p_MyBit) of an integer (p_MyByte)
    Sub SetBit(ByVal p_MyBit As Int16)
        Dim BitMask As Int16
        BitMask = 2 ^ (p_MyBit - 1)
        m_iMyByte = m_iMyByte Or BitMask
    End Sub

    'The ToggleBit Sub will change the state of the 1 based, nth bit (p_MyBit) of an integer (p_MyByte)
    Sub ToggleBit(ByVal p_MyBit As Int16)
        Dim BitMask As Int16
        BitMask = 2 ^ (p_MyBit - 1)
        m_iMyByte = m_iMyByte Xor BitMask
    End Sub


End Class
