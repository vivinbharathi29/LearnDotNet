Imports System.Data
Imports HPQ.Excalibur

Partial Class Image_ImageDriveDefinitionMain
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            BindGvDrives()
        End If
    End Sub
    Sub BindGvDrives()
        Dim dtDrives As DataTable = Images.usp_ListImageDriveDefinitions(True)
        gvDrives.DataSource = dtDrives.DefaultView
        gvDrives.DataBind()
    End Sub

    Protected Sub gvDrives_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs) Handles gvDrives.RowCancelingEdit
        gvDrives.EditIndex = -1
        gvDrives.ShowFooter = True
        BindGvDrives()
    End Sub
    Protected Sub gvDrives_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gvDrives.RowCommand
        If e.CommandName = "Insert" Then
            Dim tbDivCd As TextBox = gvDrives.FooterRow.FindControl("tbDivCd")
            Dim tbSiteCd As TextBox = gvDrives.FooterRow.FindControl("tbSiteCd")
            Dim tbDriveName As TextBox = gvDrives.FooterRow.FindControl("tbDriveName")
            Dim tbPartNo As TextBox = gvDrives.FooterRow.FindControl("tbPartNo")
            Dim tbPartNoRev As TextBox = gvDrives.FooterRow.FindControl("tbPartNoRev")
            Dim cbIsAssembly As CheckBox = gvDrives.FooterRow.FindControl("cbIsAssembly")
            'Dim cbActive As CheckBox = gvDrives.FooterRow.FindControl("cbActive")

            'Response.Write(String.Format("You Inserted {0}:{1}:{2}", tbDriveName.Text, tbPartNo.Text, tbPartNoRev.Text))

            Dim returnValue As Integer = Images.usp_ImageDriveDefinitionInsert(tbDivCd.Text, tbSiteCd.Text, tbDriveName.Text, tbPartNo.Text, tbPartNoRev.Text, cbIsAssembly.Checked.ToString(), True.ToString(), Session("LoggedInUser"), DateTime.Now())

            BindGvDrives()
        End If

    End Sub

    Protected Sub gvDrives_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles gvDrives.RowEditing
        gvDrives.EditIndex = e.NewEditIndex
        gvDrives.ShowFooter = False
        BindGvDrives()
    End Sub

    Protected Sub gvDrives_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs) Handles gvDrives.RowUpdating

        Dim hfId As HiddenField = gvDrives.Rows(gvDrives.EditIndex).FindControl("hfId")
        Dim tbDivCd As TextBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("tbDivCd")
        Dim tbSiteCd As TextBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("tbSiteCd")
        Dim tbDriveName As TextBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("tbDriveName")
        Dim tbPartNo As TextBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("tbPartNo")
        Dim tbPartNoRev As TextBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("tbPartNoRev")
        Dim cbIsAssembly As CheckBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("cbIsAssembly")
        Dim cbActive As CheckBox = gvDrives.Rows(gvDrives.EditIndex).FindControl("cbActive")

        'Response.Write(String.Format("You Updated {0}:{1}:{2}", tbDriveName.Text, tbPartNo.Text, tbPartNoRev.Text))

        Dim returnValue As Integer = Images.usp_ImageDriveDefinitionUpdate(hfId.Value, tbDivCd.Text, tbSiteCd.Text, tbDriveName.Text, tbPartNo.Text, tbPartNoRev.Text, cbIsAssembly.Checked.ToString(), cbActive.Checked.ToString, Session("LoggedInUser"), DateTime.Now())

        gvDrives.EditIndex = -1
        gvDrives.ShowFooter = True
        BindGvDrives()

    End Sub
End Class
