
Partial Class maint_Default
    Inherits System.Web.UI.Page

    Protected Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles frmMain.Load
        txtRollback.text = Server.UrlEncode(Request.QueryString("ID"))
    End Sub

    Protected Sub cmdRollback_Click(sender As Object, e As System.EventArgs) Handles cmdRollback.Click

    End Sub
End Class
