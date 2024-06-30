Imports System.IO
Imports System.Data.OleDb
Imports System.Security.AccessControl
Imports System.Configuration
Imports System.Data

Partial Class SCM_UploadAvData
    Inherits System.Web.UI.Page
    Private _Description As String
    Private _AVNumber As String
    Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Private ReadOnly Property ProductVersionID() As String
        Get
            Return Request.QueryString("PVID")
        End Get
    End Property

    Private ReadOnly Property BrandID() As String
        Get
            Return Request.QueryString("BID")
        End Get
    End Property

    Private ReadOnly Property UserName() As String
        Get
            Return Request.QueryString("UserName")
        End Get
    End Property

    Public Property sAVNumber() As String
        Get
            Return _AVNumber
        End Get
        Set(ByVal value As String)
            _AVNumber = value
        End Set
    End Property


    Public Property sDescription() As String
        Get
            Return _Description
        End Get
        Set(ByVal value As String)
            _Description = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "no-cache"
    End Sub

#Region " Upload Button Clicked "
    Private Sub Button_Clicked(ByVal sender As Object, ByVal e As EventArgs) Handles submitButton.Click
        If Not Me.IsValid Then
            bodyPre.InnerText = "Page is not valid!"
            Exit Sub
        End If
        bodyPre.InnerText = ""

        If FileUpload1.HasFile Then
            Dim saNow As String
            saNow = Path.GetExtension(FileUpload1.PostedFile.FileName)
            Dim tempPath As String = Path.GetTempPath
            Dim tempFileName As String = String.Format("{0}{1}", Path.GetTempFileName, Path.GetExtension(FileUpload1.PostedFile.FileName)) 'String.Format("{0}{1}{2}{3}{4}{5}{6}", saNow)
            Dim tempFileSpec As String = Path.Combine(tempPath, tempFileName)
            Dim uploadedFileName As String = Path.GetFileName(FileUpload1.PostedFile.FileName)

            bodyPre.InnerText += "File Name: " & uploadedFileName & vbLf
            bodyPre.InnerText += " - Size: " & FileUpload1.PostedFile.ContentLength & vbLf
            'bodyPre.InnerText += " - Content type: " & FileUpload1.PostedFile.ContentType & vbLf

            FileUpload1.PostedFile.SaveAs(tempFileSpec)

            Dim connectionString As String = String.Empty
            Select Case Path.GetExtension(uploadedFileName).ToLower()
                Case ".xls"
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;MaxScanRows=0""", tempFileSpec)
                Case ".xlsx"
                    connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1;MaxScanRows=0""", tempFileSpec)
                Case Else
                    bodyPre.InnerText += vbCrLf & String.Format("Invalid File Extension ({0})", Path.GetExtension(uploadedFileName))
            End Select

            If connectionString <> String.Empty Then
                Dim cn As OleDbConnection = New OleDbConnection(connectionString)
                Dim cmd As OleDbCommand = New OleDbCommand()
                Dim da As OleDbDataAdapter = New OleDbDataAdapter()
                Dim ds As DataSet = New DataSet()
                Dim dt As New DataTable
                Dim xlDt As New DataTable
                Dim xlTab As String = String.Empty
                Dim sqlQuery As String = String.Empty

                cn.Open()
                dt = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "Table"})
                cn.Close()

                bodyPre.InnerText += vbCrLf


                For Each row As DataRow In dt.Rows

                    'bodyPre.InnerText += vbCrLf

                    xlTab = row("TABLE_NAME").ToString()
                    sqlQuery = String.Format("SELECT * FROM [{0}]", xlTab)
                    cmd.Connection = cn
                    cmd.CommandText = sqlQuery
                    cmd.CommandType = CommandType.Text
                    da.SelectCommand = cmd
                    ds = New DataSet()
                    da.Fill(ds)
                    xlDt = ds.Tables(0)

                    saNow = Regex.Replace(xlTab, "[\W]", "")

                    bodyPre.InnerText += String.Format("Processing Sheet: {0}", xlTab)
                    bodyPre.InnerText += vbCrLf


                    Dim sColumnNames As String = String.Empty
                    For i As Integer = 0 To xlDt.Columns.Count - 1
                        sColumnNames += xlDt.Columns(i).ColumnName & "|"
                        ''Display All Columns
                        'bodyPre.InnerText += xlDt.Columns(i).ColumnName & "|"
                    Next

                    Dim bMaterialNumberExists As Boolean = False
                    Dim bDescription As Boolean = False
                    Dim saColumnNames As String() = Split(sColumnNames, "|")
                    Dim k As Integer = 0
                    For k = 0 To saColumnNames.Length - 1
                        Select Case saColumnNames(k)
                            Case "MaterialNumber"
                                bMaterialNumberExists = True
                            Case "Description"
                                bDescription = True
                        End Select
                    Next

                    If xlDt.Columns(0).ColumnName.ToString() <> (String.Empty) And bMaterialNumberExists = True And bDescription = True Then
                        bodyPre.InnerText += " - Columns: MaterialNumber | Description" & vbCrLf
                        For Each xlRow As DataRow In xlDt.Rows
                            Dim j As Integer = 0
                            For j = 0 To saColumnNames.Length - 1
                                Select Case saColumnNames(j)
                                    Case "MaterialNumber"
                                        sAVNumber = xlRow("MaterialNumber").ToString().Trim
                                    Case "Description"
                                        sDescription = xlRow("Description").ToString().Trim
                                End Select
                            Next
                            Dim i As Integer = -100

                            i = hpqData.UpdateAvDetailViaUpload(sAVNumber, sDescription, ProductVersionID, BrandID, UserName)

                            Select Case i
                                'Case 1
                                '    bodyPre.InnerText += " - Values:  " & sAVNumber & ", " & sDescription & " (EXISTS)" & vbLf
                                Case 2
                                    bodyPre.InnerText += " - Values:  " & sAVNumber & ", " & sDescription & " (NOT FOUND ON SCM)" & vbLf
                                Case 3
                                    bodyPre.InnerText += " - Values:  " & sAVNumber & ", " & sDescription & " (UPDATED)" & vbLf
                                Case Else
                                    bodyPre.InnerText += " - Values:  " & sAVNumber & ", " & sDescription & " (FAILED)" & vbLf
                            End Select
                            'If saColumnNames.Contains("MaterialNumber") Then sAVNumber = xlRow("MaterialNumber").ToString()
                            'If saColumnNames.Contains("Description") Then sDescription = sAVNumber = xlRow("MaterialNumber").ToString()

                        Next
                    End If
                    Exit For
                Next
                bodyPre.InnerText += vbLf & vbLf & "Please Verify The Results And Close The Pop Up..."
                If File.Exists(tempFileSpec) Then
                    File.Delete(tempFileSpec)
                End If
            End If
        End If
    End Sub
#End Region

End Class
