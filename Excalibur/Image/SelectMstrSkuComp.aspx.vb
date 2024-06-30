Imports System.Data
Imports HPQ.Excalibur

Partial Class Image_SelectMstrSkuComp
    Inherits System.Web.UI.Page

    Private ReadOnly Property ImageId() As String
        Get
            Return Request.QueryString("ID")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Dim dtDrives As DataTable = Images.usp_ListImageDriveDefinitions(False)

            ddlMstrSkuComp.AppendDataBoundItems = True
            ddlMstrSkuComp.DataTextField = "DriveName"
            ddlMstrSkuComp.DataValueField = "ID"
            ddlMstrSkuComp.DataSource = dtDrives
            ddlMstrSkuComp.DataBind()

            Dim ImageDriveDefinitionId As String = String.Empty
            If Not String.IsNullOrEmpty(ImageId) Then
                Dim dtImage As DataTable = Images.spGetImageProperties(ImageId)
                ' Response.Write(dtImage.Rows.Count)
                If dtImage.Rows.Count > 0 Then
                    ImageDriveDefinitionId = dtImage.Rows(0)("ImageDriveDefinitionId").ToString()
                End If
            End If

            If Not String.IsNullOrEmpty(ImageDriveDefinitionId) Then
                ddlMstrSkuComp.SelectedValue = ImageDriveDefinitionId
            End If

        End If
    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.body.Attributes.Add("onload", "window.close();")
    End Sub

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        '
        ' Update Image & Set ImageDriveDefinitionId
        '
        Images.usp_UpdateImagesImageDriveDefinitionId(ImageId, ddlMstrSkuComp.SelectedValue)
        Me.body.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", ddlMstrSkuComp.SelectedItem.Text))
    End Sub
End Class
