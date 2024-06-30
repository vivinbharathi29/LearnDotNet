
Partial Class InitialOfferingLTF
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property ActionItemID() As String
        Get
            Return Request("ID")
        End Get
    End Property

    Public ReadOnly Property DeliverableName() As String
        Get
            Return Request("Name")
        End Get
    End Property

    Public ReadOnly Property LTFAVNo() As String
        Get
            Return Request("LTFAvNo")
        End Get
    End Property

    Public ReadOnly Property LTFSANo() As String
        Get
            Return Request("LTFSANo")
        End Get
    End Property

    Public ReadOnly Property CurrentUserID() As String
        Get
            Return Request("CurrentUserID")
        End Get
    End Property

    Public ReadOnly Property DeliverableRootID() As String
        Get
            Return Request("DRID")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                lblDelName.Text = DeliverableName
                txtLTFAV.Text = LTFAVNo
                txtLTFSA.Text = LTFSANo
            End If
        Catch ex As Exception
            Response.Write("Error")
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            dw.UpdateInitialOfferingLTFAVSAs(DeliverableRootID, txtLTFAV.Text.Trim, txtLTFSA.Text.Trim, ActionItemID, CurrentUserID)
            Response.Write("<script language='javascript'> { window.close();}</script>")
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub
End Class
