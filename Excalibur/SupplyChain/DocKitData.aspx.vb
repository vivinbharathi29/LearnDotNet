Imports System.Data
Partial Class SupSCM_DocKitData
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    'Dim de As HPQ.Excalibur.EmailMessage = New HPQ.Excalibur.EmailMessage

    Public ReadOnly Property KMAT() As String
        Get
            Return Request("KMAT")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                Dim dt As New DataTable
                dt = dw.SelectDocKitsByKMAT(KMAT)

                If dt.Rows.Count = 0 Then
                    lblHeader.Text = "There Is No Doc Kit Data Available."
                Else
                    gvDocKitData.DataSource = dt
                    gvDocKitData.DataBind()
                End If

            End If
        Catch ex As Exception
            Response.Write(ex.InnerException)
        End Try
    End Sub

    Protected Sub ClosePopup(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Dim pulsarplusDivId As String = Request.QueryString("pulsarplusDivId")
            If Not String.IsNullOrEmpty(pulsarplusDivId) Then
                Response.Write("<script language='javascript'>parent.window.parent.closeExternalPopup();</script>")
            Else
                Response.Write("<script language='javascript'>{ if (parent.window.parent.document.getElementById('modal_dialog')) { parent.window.parent.modalDialog.cancel();} else {window.parent.close();} }</script>")
            End If
        Catch ex As Exception
            Response.Write(ex.InnerException)
        End Try
    End Sub


End Class
