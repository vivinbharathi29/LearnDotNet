Imports System.IO

Partial Class DownloadZip
    Inherits System.Web.UI.Page
    Private _fileSpec As String
    Public ReadOnly Property FilePath() As String
        Get
            Return Request("file")
        End Get
    End Property

    Public ReadOnly Property DeliverablePath() As String
        Get
            Return Request("DeliverablePath")
        End Get
    End Property

    Public ReadOnly Property Path2Location() As String
        Get
            Return Request("Path2Location")
        End Get
    End Property

    Public ReadOnly Property Path3Location() As String
        Get
            Return Request("Path3Location")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim fileSpec As String = Request.Form("hfZipPath")
        If fileSpec = String.Empty Then
            fileSpec = Path3Location
        End If
        If fileSpec = String.Empty Then
            fileSpec = DeliverablePath
        End If


        Dim iStream As System.IO.Stream
        Dim bufferSize As Integer = 1024 * 1024 * 3
        Dim buffer(bufferSize) As Byte
        Dim length As Integer
        Dim dataToRead As Long
        Dim filename As String = Path.GetFileName(fileSpec)
        Try
            iStream = New FileStream(fileSpec, FileMode.Open, FileAccess.Read, FileShare.Read)
            dataToRead = iStream.Length

            Response.ContentType = "application/zip"
            Response.AddHeader("Content-Length", dataToRead.ToString)
            Response.AddHeader("Content-Disposition", String.Format("attachment; filename={0}", Path.GetFileName(fileSpec)))

            While dataToRead > 0
                If Response.IsClientConnected Then
                    length = iStream.Read(buffer, 0, bufferSize)
                    Dim result As IAsyncResult = Response.OutputStream.BeginWrite(buffer, 0, length, Nothing, Nothing)
                    result.AsyncWaitHandle.WaitOne()
                    Response.OutputStream.EndWrite(result)
                    Response.Flush()

                    ReDim buffer(bufferSize)
                    dataToRead = dataToRead - length
                Else
                    dataToRead = -1
                End If
            End While
        Catch ex As Exception
            Response.Write("Error : " & ex.Message)
        Finally
            If Not IsNothing(iStream) Then
                iStream.Close()
            End If
            Response.Close()
        End Try

    End Sub
End Class
