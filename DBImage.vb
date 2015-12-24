Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Text


Public Enum DBCommandType
    DoInsert = 0
    DoParamInsert = 1
    DoSelect = 2
    DoUpdate = 3
    DoParamUpdate = 4
    DoDelete = 5
End Enum


Public Class DBImage

    Private m_iImageHeight As Integer
    Private m_iImageWidth As Integer
    Private m_iImageLength As Integer
    Private m_sImageType As String
    Private m_sKeyColumnName As String
    Private m_iKeyColumn As Integer
    Private m_sHair As String
    Private m_sEyes As String
    Private m_sEmployeeHeight As String
    Private m_sEmployeeWeight As String

    Private m_sImageFilePath As String 'File Path Of The Location Machine

    Private m_sConnection As String
    Private m_sInsert As String
    Private m_sParamInsert As String
    Private m_sSelect As String
    Private m_sDelete As String
    Private m_sUpdate As String
    Private m_sParamUpdate As String
    Private m_sFileName As String

    Private m_sTableName As String

    Sub New(ByVal p_eDBCommandType As DBCommandType, ByVal p_sImageFilePath As String, ByVal p_sSqlParameterizedInsert As String, ByVal p_sConnection As String)
        m_sImageFilePath = p_sImageFilePath
        m_sConnection = p_sConnection
        m_sParamInsert = p_sSqlParameterizedInsert
        m_iImageHeight = 0
        m_iImageWidth = 0
        m_iImageLength = 0
        m_sImageType = 0
        m_sHair = "Unknown"
        m_sEyes = "Unknown"
        m_sEmployeeHeight = "Unknown"
        m_sEmployeeWeight = ""
    End Sub

    Public Property KeyColumnName() As String
        Get
            Return m_sKeyColumnName
        End Get
        Set(ByVal Value As String)
            m_sKeyColumnName = Value
        End Set
    End Property

    Public Property KeyColumn() As Integer
        Get
            Return m_iKeyColumn
        End Get
        Set(ByVal Value As Integer)
            m_iKeyColumn = Value
        End Set
    End Property


    Public Property TableName() As String
        Get
            Return m_sTableName
        End Get
        Set(ByVal Value As String)
            m_sTableName = Value
        End Set
    End Property

    Public Property Connection() As String
        Get
            Return m_sConnection
        End Get
        Set(ByVal Value As String)
            m_sConnection = Value
        End Set
    End Property

    Public Property InsertString() As String
        Get
            Return m_sInsert
        End Get
        Set(ByVal Value As String)
            m_sInsert = Value
        End Set
    End Property

    Public Property ParamInsertString() As String
        Get
            Return m_sParamInsert
        End Get
        Set(ByVal Value As String)
            m_sParamInsert = Value
        End Set
    End Property

    Public Property SelectString() As String
        Get
            Return m_sSelect
        End Get
        Set(ByVal Value As String)
            m_sSelect = Value
        End Set
    End Property

    Public Property DeleteString() As String
        Get
            Return m_sDelete
        End Get
        Set(ByVal Value As String)
            m_sDelete = Value
        End Set
    End Property

    Public Property UpdateString() As String
        Get
            Return m_sUpdate
        End Get
        Set(ByVal Value As String)
            m_sUpdate = Value
        End Set
    End Property

    Public Property ParamUpdateString() As String
        Get
            Return m_sParamUpdate
        End Get
        Set(ByVal Value As String)
            m_sParamUpdate = Value
        End Set
    End Property

    Public Property FileName() As String
        Get
            Return m_sFileName
        End Get
        Set(ByVal Value As String)
            m_sFileName = Value
        End Set
    End Property

    Public Function Save() As Boolean
        m_iImageHeight = Nothing
        m_iImageWidth = Nothing
        m_iImageLength = Nothing
        m_sImageType = Nothing
        m_sHair = "Unknown"
        m_sEyes = "Unknown"
        m_sEmployeeHeight = "Unknown"
        m_sEmployeeWeight = ""

        Try
            Dim sFilePath As String
            Dim iImageFilePathLength As Integer = m_sImageFilePath.ToString().Length - (m_sImageFilePath.ToString().Length - m_sImageFilePath.ToString().LastIndexOf("."))
            sFilePath = m_sImageFilePath.ToString().Substring(iImageFilePathLength)
            If sFilePath <> ".tif" And sFilePath <> ".jpg" And sFilePath <> ".jpeg" And sFilePath <> ".jif" And sFilePath <> ".jfif" And sFilePath <> ".png" And sFilePath <> ".gif" And sFilePath <> ".tiff" Then
                Return False
            End If


            Dim fs As FileStream = New FileStream(m_sImageFilePath.ToString(), FileMode.Open, FileAccess.Read)

            'Read File Into Byte Array
            Dim img As Byte() = New Byte(fs.Length) {}
            fs.Read(img, 0, fs.Length)
            fs.Close()

            'Obtain Information About the File
            Dim oImage As Image = Image.FromFile(m_sImageFilePath.ToString())
            m_iImageHeight = oImage.Height
            m_iImageWidth = oImage.Width
            m_iImageLength = oImage.PropertyItems.Length
            m_sImageType = Path.GetExtension(m_sImageFilePath)
            oImage = Nothing

            'Image Data Validation
            If m_iImageHeight <> 90 Or m_iImageWidth <> 90 Or m_iImageLength = 0 Or (m_sImageType <> ".tif" And m_sImageType <> ".jpg" And m_sImageType <> ".jpeg" And m_sImageType <> ".jif" And m_sImageType <> ".jfif" And m_sImageType <> ".png" And m_sImageType <> ".gif" And m_sImageType <> ".tiff") Then
                Return False
            Else
                'Insert into database
                Dim sb As New StringBuilder
                sb.Append("INSERT INTO ")
                sb.Append(m_sTableName.ToString())
                sb.Append(" (")
                sb.Append(m_sKeyColumnName.ToString())
                sb.Append(",Photo, Height, Width, Length, Type, Hair, Eyes, EmployeeHeight, EmployeeWeight) VALUES (")
                sb.Append(m_iKeyColumn.ToString())
                sb.Append(",@image")
                sb.Append(",")
                sb.Append(m_iImageHeight.ToString())
                sb.Append(",")
                sb.Append(m_iImageWidth.ToString())
                sb.Append(",")
                sb.Append(m_iImageLength.ToString())
                sb.Append(",'")
                sb.Append(m_sImageType.ToString())
                sb.Append("', '")
                sb.Append("Unknown")
                sb.Append("', '")
                sb.Append("Unknown")
                sb.Append("', '")
                sb.Append("Unknown")
                sb.Append("', '")
                sb.Append("")
                sb.Append("')")

                Dim sInsert As String = sb.ToString()

                Dim oConn As SqlConnection = New SqlConnection(m_sConnection)
                Dim oCmd As SqlCommand = New SqlCommand(sInsert, oConn)

                'Image content
                Dim oPic As SqlParameter = New SqlParameter("@image", SqlDbType.Image)
                oPic.Value = img
                oCmd.Parameters.Add(oPic)

                'Execute Insert
                oConn.Open()
                oCmd.ExecuteNonQuery()
                oConn.Close()

                Return True
            End If
        Catch ex As Exception
            Dim s As String = ex.Message
        End Try
    End Function

    Public Function Update() As Boolean
        m_iImageHeight = Nothing
        m_iImageWidth = Nothing
        m_iImageLength = Nothing
        m_sImageType = Nothing
        Try
            Dim sFilePath As String
            Dim iImageFilePathLength As Integer = m_sImageFilePath.ToString().Length - (m_sImageFilePath.ToString().Length - m_sImageFilePath.ToString().LastIndexOf("."))
            sFilePath = m_sImageFilePath.ToString().Substring(iImageFilePathLength)
            If sFilePath <> ".tif" And sFilePath <> ".jpg" And sFilePath <> ".jpeg" And sFilePath <> ".jif" And sFilePath <> ".jfif" And sFilePath <> ".png" And sFilePath <> ".gif" And sFilePath <> ".tiff" Then
                Return False
            End If

            Dim fs As FileStream = New FileStream(m_sImageFilePath.ToString(), FileMode.Open, FileAccess.Read)

            'Read File Into Byte Array
            Dim img As Byte() = New Byte(fs.Length) {}
            fs.Read(img, 0, fs.Length)
            fs.Close()

            'Obtain Information About the File
            Dim oImage As Image = Image.FromFile(m_sImageFilePath.ToString())
            m_iImageHeight = oImage.Height
            m_iImageWidth = oImage.Width
            m_iImageLength = oImage.PropertyItems.Length
            m_sImageType = Path.GetExtension(m_sImageFilePath)
            oImage = Nothing

            'Image Data Validation
            If m_iImageHeight <> 90 Or m_iImageWidth <> 90 Or m_iImageLength <= 0 Or (m_sImageType <> ".tif" And m_sImageType <> ".jpg" And m_sImageType <> ".jpeg" And m_sImageType <> ".jif" And m_sImageType <> ".jfif" And m_sImageType <> ".png" And m_sImageType <> ".gif" And m_sImageType <> ".tiff") Then
                Return False
            Else

                'Update the database
                Dim sb As New StringBuilder
                sb.Append("UPDATE ")
                sb.Append(m_sTableName.ToString())
                sb.Append(" SET ")
                sb.Append(" Photo = ")
                sb.Append("@image")
                sb.Append(",")
                sb.Append(" Height = ")
                sb.Append(m_iImageHeight.ToString())
                sb.Append(",")
                sb.Append(" Width = ")
                sb.Append(m_iImageWidth.ToString())
                sb.Append(",")
                sb.Append(" Length = ")
                sb.Append(m_iImageLength.ToString())
                sb.Append(",")
                sb.Append(" Type = '")
                sb.Append(m_sImageType.ToString())
                sb.Append("' Where  ")
                sb.Append(m_sKeyColumnName.ToString())
                sb.Append(" = ")
                sb.Append(m_iKeyColumn.ToString())

                Dim sInsert As String = sb.ToString()

                Dim oConn As SqlConnection = New SqlConnection(m_sConnection)
                Dim oCmd As SqlCommand = New SqlCommand(sInsert, oConn)

                'Image content
                Dim oPic As SqlParameter = New SqlParameter("@image", SqlDbType.Image)
                oPic.Value = img
                oCmd.Parameters.Add(oPic)

                'Execute Insert
                oConn.Open()
                oCmd.ExecuteNonQuery()
                oConn.Close()

                Return True
            End If
        Catch ex As Exception
            Dim s As String = ex.Message
        End Try
    End Function

    Public Function SaveWithParameter() As Boolean
        Try
            Dim sFilePath As String
            Dim iImageFilePathLength As Integer = m_sImageFilePath.ToString().Length - (m_sImageFilePath.ToString().Length - m_sImageFilePath.ToString().LastIndexOf("."))
            sFilePath = m_sImageFilePath.ToString().Substring(iImageFilePathLength)
            If sFilePath <> ".tif" And sFilePath <> ".jpg" And sFilePath <> ".jpeg" And sFilePath <> ".jif" And sFilePath <> ".jfif" And sFilePath <> ".png" And sFilePath <> ".gif" And sFilePath <> ".tiff" Then
                Return False
            End If


            Dim fs As FileStream = New FileStream(m_sImageFilePath.ToString(), FileMode.Open, FileAccess.Read)

            'Read File Into Byte Array
            Dim img As Byte() = New Byte(fs.Length) {}
            fs.Read(img, 0, fs.Length)
            fs.Close()

            'Obtain Information About the File
            Dim oImage As Image = Image.FromFile(m_sImageFilePath.ToString())
            m_iImageHeight = oImage.Height
            m_iImageWidth = oImage.Width
            m_iImageLength = oImage.PropertyItems.Length
            m_sImageType = Path.GetExtension(m_sImageFilePath)
            oImage = Nothing

            'Insert into database
            Dim oConn As SqlConnection = New SqlConnection(m_sConnection)
            Dim oCmd As SqlCommand = New SqlCommand(m_sParamInsert, oConn)

            'Image content
            Dim oPic As SqlParameter = New SqlParameter("@image", SqlDbType.Image)
            oPic.Value = img
            oCmd.Parameters.Add(oPic)

            'Execute Insert
            oConn.Open()
            oCmd.ExecuteNonQuery()
            oConn.Close()
        Catch ex As Exception
            Dim s As String = ex.Message
        End Try
    End Function
    'Summary: Displaying the image from database on the form
    'Details: Loading image from the database into file and then from the file to the form
    Public Sub Load(ByVal p_sSelect As String)
        Dim sExceptionMessage As String
        Dim DataAdapter As New SqlDataAdapter
        Dim localConn As New SqlConnection(m_sConnection)
        Dim dtView As New DataView

        DataAdapter.SelectCommand = New SqlCommand

        With DataAdapter.SelectCommand
            .Connection = localConn
            .CommandTimeout = 120
            .CommandText = p_sSelect
            .CommandType = CommandType.Text
        End With
        Try
            Dim dsData As DataSet = New DataSet
            localConn.Open()

            DataAdapter.Fill(dsData)
            dtView.Table = dsData.Tables(0)

            If dtView.Table.Rows.Count > 0 Then
                Dim sType As String = dsData.Tables(0).Rows(0).Item("Type")
                Dim sFile As Byte() = dsData.Tables(0).Rows(0).Item("Photo")
                Dim sFileName As String = dsData.Tables(0).Rows(0).Item("EmployeeID")
                m_iImageLength = sFile.Length


                m_sFileName = sFileName & sType
                Dim oFileStream As FileStream = New FileStream(m_sFileName, FileMode.Create, FileAccess.Write)

                If (m_iImageLength > 0) Then
                    oFileStream.Write(sFile, 0, m_iImageLength)
                    oFileStream.Close()
                End If
            Else
                m_sFileName = ""
            End If
        Catch ex As Exception
            sExceptionMessage = ex.Message()
        End Try
    End Sub
End Class
