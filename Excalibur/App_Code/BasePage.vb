Imports Microsoft.VisualBasic
Imports System.Text
Imports System.IO

Public Class BasePage
    Inherits System.Web.UI.Page

    Protected Overloads Overrides Sub Render(ByVal writer As HtmlTextWriter)
        Try
            Dim renderedOutput As New StringBuilder()
            Dim strWriter As New StringWriter(renderedOutput)
            Dim tWriter As New HtmlTextWriter(strWriter)
            MyBase.Render(tWriter)

            'this string is to be searched for src="/" mce_src="/" and replace it with correct src="./" mce_src="./". 

            Dim s As String = renderedOutput.ToString()
            s = Regex.Replace(s, "(?<=<img[^>]*)(src=\""/)", String.Format("src=""{0}/", Session("ApplicationRoot")), RegexOptions.IgnoreCase)
            s = Regex.Replace(s, "(?<=<script[^>]*)(src=\""/)", String.Format("src=""{0}/", Session("ApplicationRoot")), RegexOptions.IgnoreCase)
            s = Regex.Replace(s, "(?<=<link[^>]*)(href=\""/)", String.Format("href=""{0}/", Session("ApplicationRoot")), RegexOptions.IgnoreCase)
            s = Regex.Replace(s, "(?<=<a[^>]*)(href=\""/)", String.Format("href=""{0}/", Session("ApplicationRoot")), RegexOptions.IgnoreCase)
            's = s.Replace("/NeatUpload/", "../NeatUpload/")

            writer.Write(s)
        Catch e As Exception
        End Try
    End Sub

End Class
