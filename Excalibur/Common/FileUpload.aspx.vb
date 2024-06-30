Imports System.IO
Imports Brettle.Web.NeatUpload

Partial Class Common_FileUpload
    Inherits System.Web.UI.Page

#Region " Properties "
    Private ReadOnly Property DestinationFolder() As String
        Get
            Dim QueryFolder As String = Request.QueryString("Folder")
            If QueryFolder = String.Empty Then
                Return Path.GetRandomFileName
            Else
                Return QueryFolder
            End If
        End Get
    End Property

    Private ReadOnly Property FileTypeFilter() As String
        Get
            Dim strFileTypeFilter As String = Request.QueryString("Filter")
            If strFileTypeFilter = String.Empty Then
                Return String.Empty
            Else
                Return strFileTypeFilter
            End If
        End Get
    End Property

    Private ReadOnly Property FilterExcludeList() As String
        Get
            Dim strFilterExcludeList As String = Request.QueryString("FilterExclude").ToString()
            If strFilterExcludeList = String.Empty Then
                Return String.Empty
            Else
                Return strFilterExcludeList
            End If
        End Get
    End Property

    Private ReadOnly Property TitleText() As String
        Get
            Dim strTitle As String = Request.QueryString("Title")
            If strTitle = String.Empty Then
                Return "Pulsar File Upload"
            Else
                Return strTitle
            End If
        End Get
    End Property

    Private ReadOnly Property KeepLocal() As Boolean
        Get
            Dim strKeepLocal As String = Request.QueryString("KeepLocal")
            Dim bKeepLocal As Boolean = False
            Boolean.TryParse(strKeepLocal, bKeepLocal)
            Return bKeepLocal
        End Get
    End Property
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetExpires(DateTime.Now())
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        hidAppName.Value = Request.QueryString("AppName")
        hidControlId.Value = Request.QueryString("ControlId")
        If Not Page.IsPostBack Then
            submitButton.OnClientClick = "this.value='Uploading...';setTimeout('document.getElementById(\'submitButton\').disabled=true',100);"
            lblTitle.Text = TitleText
            InputFile1.StorageConfig("tempDirectory") = Path.Combine(Request.PhysicalApplicationPath, "temp")
            errorText.Visible = False
        Else
            submitButton.Enabled = False
            submitButton.Text = "Upload Complete"
        End If
        If FileTypeFilter.ToLower = "zip" Then
            RegularExpressionValidator1.ValidationExpression = ".*((.exe)|(.zip)|(.tar)|(.tgz))$"
            RegularExpressionValidator1.ErrorMessage = "Only Zip, Tar, Tgz & Softpaq files are allowed."
        Else
            RegularExpressionValidator1.ValidationExpression = "^(?!.*\.(exe|bat|msi|js|cmd)$).*"
            RegularExpressionValidator1.ErrorMessage = "Executable Files Are Not Allowed."
        End If
    End Sub

    Protected Sub submitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles submitButton.Click
        Dim fileStorePath As String = ConfigurationManager.AppSettings("FileStorePath")
        Dim tempFileFolder As String = Path.Combine(Request.PhysicalApplicationPath, "temp")
        Dim tempFileLocation As String = Path.Combine(tempFileFolder, Path.GetRandomFileName)

        errorText.Visible = False

        If InputFile1.HasFile And Page.IsValid Then

            Try
                If KeepLocal Then
                    tempFileLocation = Path.Combine(tempFileFolder, InputFile1.FileName)
                Else
                    tempFileLocation = tempFileLocation & Path.GetExtension(InputFile1.FileName)
                End If
                InputFile1.MoveTo(tempFileLocation, MoveToOptions.Overwrite)
            Catch
            End Try

            If File.Exists(tempFileLocation) Then

                Dim destPath As String = Path.Combine(fileStorePath, DestinationFolder)
                If Not Directory.Exists(destPath) Then
                    Directory.CreateDirectory(destPath)
                End If

                Dim fileLocation As String = Path.Combine(destPath, InputFile1.FileName)

                File.Copy(tempFileLocation, fileLocation, True)


                If KeepLocal Then
                    fileLocation = fileLocation & "|" & tempFileLocation
                Else
                    File.Delete(tempFileLocation)
                End If

                hidReturnValue.Value = fileLocation
            Else
                errorText.Visible = True
                lblErrUploading.Text = "An error occured while uploading your file.  Please try again."

            End If


        End If
    End Sub

End Class
