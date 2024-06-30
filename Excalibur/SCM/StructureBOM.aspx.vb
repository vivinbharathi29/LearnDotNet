Partial Class SCM_StructureBOM
    Inherits System.Web.UI.Page
    Public ReadOnly Property CurrentUser() As String
        Get
            Return Request("User")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property KMAT() As String
        Get
            Return Request("KMAT")
        End Get
    End Property

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim URL As String = Nothing
            Dim applicationRoot As String = Session("ApplicationRoot")
            btnCancel.Enabled = False
            btnSubmit.Enabled = False
            If ddlMain.SelectedItem.Value = 0 Then
                ddlMain.BackColor = Drawing.Color.MistyRose
                Exit Sub
            ElseIf Request.Form("ddlSub") = 1 And ddlMain.SelectedItem.Value = 1 Then 'Doc Kit SA - Image AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 2 And ddlMain.SelectedItem.Value = 1 Then 'Image SA - Image AV
                URL = "/iPulsar/ExcelExport/ImageBOM.aspx?BID=" & BID & "&UserName=" & CurrentUser & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 3 And ddlMain.SelectedItem.Value = 1 Then 'COA SA - Image AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 4 And ddlMain.SelectedItem.Value = 2 Then 'AC Adapter SA - HWKit AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 5 And ddlMain.SelectedItem.Value = 2 Then 'Power Cord SA - HWKit AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 5 And ddlMain.SelectedItem.Value = 3 Then 'Power Cord SA - Keyboard AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&PVID=" & PVID & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            ElseIf Request.Form("ddlSub") = 6 And ddlMain.SelectedItem.Value = 3 Then 'Keyboard SA - Keyboard AV
                URL = "/iPulsar/ExcelExport/StructureToBOM.aspx?User=" & CurrentUser & "&PVID=" & PVID & "&BID=" & BID & "&KMAT=" & KMAT & _
                "&Child=" & Request.Form("ddlSub") & "&Parent=" & ddlMain.SelectedItem.Value & "&ItemNumber=" & Request.Form("txtItemNumber")
            End If

            Response.Redirect(URL)
        Catch ex As Exception
            Response.Write(ex.Message.ToString)
        End Try
    End Sub

End Class
