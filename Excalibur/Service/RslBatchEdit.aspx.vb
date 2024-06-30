Imports System.Data

Partial Class Service_RslBatchEdit
    Inherits System.Web.UI.Page


    Private productVersionIdValue As String
    Public ReadOnly Property ProductVersionId() As String
        Get
            If productVersionIdValue = String.Empty Then
                productVersionIdValue = Request.QueryString("PVID").ToString()
            End If
            Return productVersionIdValue
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
            Dim dt As DataTable = dw.usp_SelectServiceSpareKitsForProduct(ProductVersionId)
            Repeater1.DataSource = dt.DefaultView
            Repeater1.DataBind()
        End If
    End Sub
End Class
