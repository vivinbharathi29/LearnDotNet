Partial Class SCM_LocalizeAVsSeriesNum
    Inherits System.Web.UI.Page

    Public ReadOnly Property SeriesNumbers() As String
        Get
            Return Request("strSeriesSummary")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property UserName() As String
        Get
            Return Request("UserName")
        End Get
    End Property

    Public ReadOnly Property PulsarplusDivId() As String
        Get
            Return Request("pulsarplusDivId")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            lblAvType.ForeColor = Drawing.Color.Black
            Dim strSeriesNumber() As String = SeriesNumbers.Trim.Split(",")
            Dim i As Integer = 0
            ddlSeriesNumbers.Items.Add("")
            For i = 0 To strSeriesNumber.Length - 1
                ddlSeriesNumbers.Items.Add(strSeriesNumber(i))
            Next
            If Request("KMAT") = "" Then
                trKMAT.Visible = True
                btnSubmit.Enabled = False
            Else
                trKMAT.Visible = False
                btnCancel.Enabled = True
            End If
        Catch ex As Exception
            Response.Write("Error")
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Select Case ddlAvType.SelectedItem.Value
                Case 0
                    lblAvType.ForeColor = Drawing.Color.Red
                Case 1, 2, 3, 4 '1=Images, 2=HWKits, 3=Keyboards, 4=OS Restore Media
                    Dim strPath As String = Me.ResolveUrl("~/SCM/LocalizeAVs.asp")
                    strPath = """" & strPath & "?Mode=add&PVID=" & PVID & "&strSeriesSummary=" & ddlSeriesNumbers.SelectedItem.Text & "&BID=" & BID & "&AvType=" & ddlAvType.SelectedItem.Value & "&UserName=" & UserName & """, """", ""dialogWidth:875px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No"""
                    If PulsarplusDivId <> "" Then
                        strPath += "&pulsarplusDivId=" & PulsarplusDivId
                    End If
                    thisBody.Attributes.Add("onload", "parent.window.parent.OpenLocalizeAVs(" & strPath & ");")
            End Select
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub
End Class
