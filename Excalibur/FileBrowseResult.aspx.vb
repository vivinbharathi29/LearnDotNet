Imports System.IO
Imports System.Net.Mail
Imports HPQ.Excalibur

Partial Class FileBrowseResult
    Inherits System.Web.UI.Page

#Region " Page Properties "
    Private _currentDomain As String
    Public Property CurrentDomain() As String
        Get
            If _currentDomain = String.Empty Then
                _currentDomain = ViewState.Item("currentDomain").ToString()
            End If
            Return _currentDomain
        End Get
        Set(ByVal value As String)
            _currentDomain = value
            ViewState.Item("currentDomain") = _currentDomain
        End Set
    End Property

    Private _currentUser As String
    Public Property CurrentUser() As String
        Get
            If _currentUser = String.Empty Then
                _currentUser = ViewState.Item("currentUser").ToString()
            End If
            Return _currentUser
        End Get
        Set(ByVal value As String)
            _currentUser = value
            ViewState.Item("currentUser") = _currentUser
        End Set
    End Property

    Public ReadOnly Property DeliverablePath() As String
        Get
            Return Request("DeliverablePath").ToLower()
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

    Public ReadOnly Property TDCImagePath() As String
        Get
            Return Request("TDCImagePath")
        End Get
    End Property

    Public ReadOnly Property DeliverableName() As String
        Get
            Return Request("DeliverableName")
        End Get
    End Property

    Public ReadOnly Property DeliverableVersionId() As String
        Get
            Return Request("DeliverableID")
        End Get
    End Property
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            If Not Page.IsPostBack Then
                Dim secObj As HPQ.Excalibur.Security = New Security(Session("LoggedInUser"))
                CurrentUser = secObj.CurrentUser
                CurrentDomain = secObj.CurrentUserDomain


                If Right(CurrentUser.ToLower, 7) <> "@hp.com" And CurrentDomain.ToLower <> "americas" And CurrentDomain.ToLower <> "emea" And CurrentDomain.ToLower <> "asiapacific" Then
                    LoadFileWindow.Visible = False

                    Dim filePath As String = Path3Location.Trim()
                    If filePath = String.Empty And
                        (DeliverablePath.EndsWith(".zip") Or
                         DeliverablePath.EndsWith(".exe") Or
                         DeliverablePath.EndsWith(".html") Or
                         DeliverablePath.EndsWith(".txt") Or
                         DeliverablePath.EndsWith(".cva") Or
                         DeliverablePath.EndsWith(".tar") Or
                         DeliverablePath.EndsWith(".tgz") Or
                         DeliverablePath.EndsWith(".iso")) Then
                        filePath = DeliverablePath
                    End If

                    If filePath = String.Empty Then
                        filePath = String.Format("\\houhpqprtrel05\Zips\{0}\{0}.zip", DeliverableVersionId)
                    End If

                    If filePath <> String.Empty And File.Exists(filePath) Then

                        Dim myFileInfo As FileInfo = New FileInfo(filePath)
                        Dim fileSize As Long = myFileInfo.Length
                        Dim fileSizeString As String

                        If fileSize > 1024 * 1024 * 1024 Then
                            fileSizeString = Format(myFileInfo.Length / (1024 * 1024 * 1024), "#,### GB")
                        ElseIf fileSize > 1024 * 1024 Then
                            fileSizeString = Format(myFileInfo.Length / (1024 * 1024), "#,### MB")
                        ElseIf fileSize > 1024 Then
                            fileSizeString = Format(myFileInfo.Length / 1024, "#,### KB")
                        Else
                            fileSizeString = Format(myFileInfo.Length, "#,### Bytes")
                        End If

                        pnlFileInfo.Visible = True
                        lblFileNameText.Text = myFileInfo.Name
                        lblFileSizeText.Text = fileSizeString
                        hfZipPath.Value = filePath
                    Else
                        pnlFileMissing.Visible = True
                        If filePath = String.Empty Then
                            lblFileError.Text = String.Format("{0} is not currently available for download.  The Release Lab has been notified of your request.", DeliverableName)
                        Else
                            lblFileError.Text = String.Format("Excalibur can not retrieve {0}.  The Release Lab has been notified.", filePath)

                        End If

                        Dim msg As MailMessage = New MailMessage()
                        msg.From = New MailAddress(secObj.CurrentUserEmail)
                        msg.To.Add(New MailAddress("tammy.schapiro@hp.com"))
                        msg.To.Add(New MailAddress("jones.ramsey@hp.com"))
                        msg.Bcc.Add(New MailAddress("kenneth.berntsen@hp.com"))
                        msg.Subject = "Excalibur File Download Request"
                        msg.IsBodyHtml = False
                        msg.ReplyTo = New MailAddress(secObj.CurrentUserEmail)

                        Dim msgBody As StringBuilder = New StringBuilder()
                        msgBody.Append(secObj.CurrentUserFullName)
                        msgBody.Append(" has requested the following deliverable for download.")
                        msgBody.Append(vbCrLf)
                        msgBody.Append(vbCrLf)
                        msgBody.Append("Deliverable Name: " + DeliverableName)
                        msgBody.Append(vbCrLf)
                        msgBody.Append(vbCrLf)
                        msgBody.Append("Deliverable Version ID: " + DeliverableVersionId)
                        msgBody.Append(vbCrLf)
                        msgBody.Append(vbCrLf)

                        If Path3Location <> String.Empty Then
                            msgBody.Append("Zip File Path: " + filePath)
                            msgBody.Append(vbCrLf)
                            msgBody.Append(vbCrLf)
                        End If

                        If DeliverablePath <> String.Empty Then
                            msgBody.Append("Server Path: " + DeliverablePath)
                        Else
                            msgBody.Append("Server Path: " + TDCImagePath)
                        End If
                        msgBody.Append(vbCrLf)

                        msg.Body = msgBody.ToString

                        Dim smtp As SmtpClient = New SmtpClient()

                        Try
                            smtp.Send(msg)
                        Catch ex As SmtpException
                            Throw ex
                        Catch ex As Exception
                            Throw ex
                        End Try

                        msg = Nothing
                        msgBody = Nothing
                        smtp = Nothing

                    End If

                    Exit Sub
                End If

                If DeliverablePath <> String.Empty Then
                    hfPath.Value = DeliverablePath
                ElseIf TDCImagePath <> String.Empty Then
                    hfPath.Value = TDCImagePath
                End If

                If CurrentDomain.ToLower <> "asiapacific" And CurrentDomain.ToLower <> Server.MachineName.ToLower Then

                    If DeliverablePath <> String.Empty Then
                        LoadFileWindow.Attributes.Add("src", String.Format("file://{0}", DeliverablePath))
                    ElseIf TDCImagePath <> String.Empty Then
                        LoadFileWindow.Attributes.Add("src", String.Format("file://{0}", TDCImagePath))
                    End If
                Else
                    If TDCImagePath <> String.Empty Then
                        LoadFileWindow.Attributes.Add("src", String.Format("file://{0}", TDCImagePath))
                    ElseIf DeliverablePath <> String.Empty Then
                        LoadFileWindow.Attributes.Add("src", String.Format("file://{0}", DeliverablePath))
                    End If
                End If

            End If

        Catch ex As Exception
            Response.Write(ex)
        End Try

    End Sub


    Protected Sub lbDownload_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbDownload.Click
        Dim filePath As String = hfZipPath.Value

        If filePath = String.Empty Then
            filePath = String.Format("\\houhpqprtrel05\Zips\{0}\{0}.zip", DeliverableVersionId)
        End If

        Server.Transfer("DownLoadZip.aspx", True)
    End Sub
End Class
