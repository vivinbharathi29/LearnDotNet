Imports System.Data

Partial Class MobileSE_Today_AddAvNoToFeature
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()



    Public ReadOnly Property AvCreateID() As String
        Get
            Return Request("AvCreateID")
        End Get
    End Property

    Public ReadOnly Property CurrentUserId() As String
        Get
            Return Request("CurrentUserId")
        End Get
    End Property

    Public ReadOnly Property FeatureName() As String
        Get
            Return Request("FeatureName")
        End Get
    End Property

    Public ReadOnly Property FeatureId() As String
        Get
            Return Request("FeatureId")
        End Get
    End Property

    Public ReadOnly Property ProductBrandId() As String
        Get
            Return Request("ProductBrandId")
        End Get
    End Property

    Public ReadOnly Property SCMCategoryId() As String
        Get
            Return Request("SCMCategoryId")
        End Get
    End Property

    Public ReadOnly Property ProductVersionID() As String
        Get
            Return Request("ProductVersionID")
        End Get
    End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            btnSubmit.Enabled = True
            cboAVDescription.Visible = True
            lblAV.Visible = True
            If Not Me.Page.IsPostBack Then
                lblDelName.Text = FeatureName
                GetAVNoDescription()
            End If
        Catch ex As Exception
            Response.Write("Error")
        End Try
    End Sub


    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Try
            If cboAVDescription.SelectedValue <> "" Then

                dw.UpdateAvDetailFeatureID(AvCreateID, FeatureId, CurrentUserId, cboAVDescription.SelectedValue.Trim)

                'Response.Write("<script language='javascript'> { window.close();}</script>")
                Response.Write("<script language='javascript'> { parent.window.parent.CloseExistingAVNoToFeature('1','" & AvCreateID & "'); }</script>")
            Else
                lblAV.ForeColor = Drawing.Color.Red
            End If
            
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Private Sub GetAVNoDescription()
        Try

            Dim dtData As New DataTable

            lblNoData.Visible = False
            dtData = dw.SelectAVNoDescriptions(ProductVersionID, ProductBrandId, SCMCategoryId)
            If dtData.Rows.Count > 0 Then
                cboAVDescription.DataSource = dtData
                cboAVDescription.DataTextField = "AvNoGPGDescription"
                cboAVDescription.DataValueField = "AvNo"
                cboAVDescription.DataBind()
            Else
                lblNoData.Visible = True
                lblNoData.Text = "The selected Feature's Product, Brand and SCM Category (if applicable) does not have any existing AV No.'s available."
                cboAVDescription.Visible = False
                lblAV.Visible = False
                btnSubmit.Enabled = False
            End If


        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

End Class
