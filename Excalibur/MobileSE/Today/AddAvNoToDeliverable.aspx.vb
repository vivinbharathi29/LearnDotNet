
Partial Class AddAvNoToDeliverable
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property AvCreateID() As String
        Get
            Return Request("ID")
        End Get
    End Property

    Public ReadOnly Property CurrentUserId() As String
        Get
            Return Request("CurrentUserId")
        End Get
    End Property

    Public ReadOnly Property DeliverableRootId() As String
        Get
            Return Request("DeliverableRootId")
        End Get
    End Property

    Public ReadOnly Property DeliverableName() As String
        Get
            Return Request("DeliverableName")
        End Get
    End Property

    Public ReadOnly Property ProductBrandId() As String
        Get
            Return Request("ProductBrandId")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                lblDelName.Text = DeliverableName
                cbUpdateDesc.Checked = False
            End If
        Catch ex As Exception
            Response.Write("Error")
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If txtAV.Text <> "" Then
                If cbUpdateDesc.Checked Then
                    dw.UpdateAvDetailDeliverableRootID(AvCreateID, DeliverableRootId, ProductBrandId, CurrentUserId, txtAV.Text, 1)
                Else
                    dw.UpdateAvDetailDeliverableRootID(AvCreateID, DeliverableRootId, ProductBrandId, CurrentUserId, txtAV.Text, 0)
                End If
                'Response.Write("<script language='javascript'> if (IsFromPulsarPlus()) {ClosePulsarPlusPopup();} else { window.close();}</script>")
                Dim scriptKey As String = "UniqueKeyForThisScript"
                Dim javaScript As String = "<script language='javascript'> if (IsFromPulsarPlus()) {window.parent.parent.parent.popupCallBack(1); ClosePulsarPlusPopup();} else { window.close();}</script>"
                ClientScript.RegisterStartupScript(Me.GetType(), scriptKey, javaScript)
            Else
                lblAV.ForeColor = Drawing.Color.Red
            End If
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub
End Class
