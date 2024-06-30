Partial Class InitialOfferingFilter
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim dtCategories As Data.DataTable = dw.SelectInitialOfferingCategories()

                dtCategories.Rows.Add("", 0)
                dtCategories.DefaultView.Sort = String.Format("Name", "{0}")
                dtCategories = dtCategories.DefaultView.ToTable

                ddlCategory.DataTextField = "Name"
                ddlCategory.DataValueField = "ID"

                ddlCategory.DataSource = dtCategories

                ddlCategory.DataBind()
            End If
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If ddlBusUnit.SelectedItem.Value = 0 Then
                lblBusUnit.ForeColor = Drawing.Color.Red
                Exit Sub
            ElseIf ddlCategory.SelectedItem.Value = 0 Then
                lblCategory.ForeColor = Drawing.Color.Red
                Exit Sub
            End If

            Dim URL As String = Nothing
            Dim applicationRoot As String = Session("ApplicationRoot")

            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", ddlBusUnit.SelectedItem.Value & "," & ddlCategory.SelectedItem.Value))
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub
End Class
