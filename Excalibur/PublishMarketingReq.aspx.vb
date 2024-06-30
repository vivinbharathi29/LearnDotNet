Imports System.Data

Partial Class PublishMarketingReq
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                Dim dt As DataTable = dw.SelectInitialOfferingMarketingReq(PVID)
                If dt.Rows.Count = 0 Then
                    lblHeader.Text = "There Are No Initial Offering Deliverables Linked To This Product"
                Else
                    lblHeader.Text = "Please click 'OK' to Confirm - " & dt.Rows(1).Item("DOTSName")
                    gvMarketingReq.DataSource = dt
                    gvMarketingReq.DataBind()
                End If
            End If
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            For Each gvRow As GridViewRow In gvMarketingReq.Rows
                Dim lblDRID As System.Web.UI.WebControls.Label = gvRow.FindControl("lblDRID")
                Dim lblProductReqID As System.Web.UI.WebControls.Label = gvRow.FindControl("lblProductReqID")
                Dim sAlert As String = lblDRID.Text & "," & lblProductReqID.Text & "," & PVID
                Response.Write("<script language='javascript'>{ alert('" & sAlert & "'); }</script>")
                'dw.AddDelRoot2ProductReq(lblProductReqID.Text, PVID, lblDRID.Text)
            Next
            Response.Write("<script language='javascript'> { window.close();}</script>")
        Catch ex As Exception
            'Response.Write(ex.ToString)
            lblHeader.Text = ex.ToString
            gvMarketingReq.Visible = False
        End Try
    End Sub

End Class
