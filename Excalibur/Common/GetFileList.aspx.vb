Imports System.IO

Partial Class Common_GetFileList
    Inherits System.Web.UI.Page

    <System.Web.Services.WebMethod()> _
    Public Function FileList(ByVal Path As String) As String
        Dim sFiles As String
        sFiles = ""

        If Directory.Exists(Path) Then
            For Each sFile As String In Directory.GetFiles(Path)
                If sFiles <> "" Then
                    sFiles = sFiles & ","
                End If

                sFiles = sFiles & sFile

            Next
        End If

        Return sFiles
    End Function


End Class
