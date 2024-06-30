Imports System.Data
Partial Class CommodityGuidanceCategoryFilter
    Inherits System.Web.UI.Page
    Dim CategoryIDs As String = ""
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim dtCommercial As Data.DataTable = dw.SelectAvFeatureCategoriesFilter(1)
                Dim dtConsumer As Data.DataTable = dw.SelectAvFeatureCategoriesFilter(2)

                lbCommercial.DataSource = dtCommercial
                lbConsumer.DataSource = dtConsumer
                lbCommercial.DataBind()
                lbConsumer.DataBind()

                lbConsumer.Visible = False
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.ToString
        End Try
    End Sub

    Protected Sub rblBusiness_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblBusiness.SelectedIndexChanged
        Try
            If rblBusiness.SelectedItem.Value = 1 Then
                lbCommercial.Visible = True
                lbConsumer.Visible = False
            ElseIf rblBusiness.SelectedItem.Value = 2 Then
                lbCommercial.Visible = False
                lbConsumer.Visible = True
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.ToString
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If rblBusiness.SelectedItem.Value = 1 Then
                CategoryIDs = ""
                For Each item As ListItem In lbCommercial.Items
                    If item.Selected Then
                        If CategoryIDs = "" Then
                            CategoryIDs = item.Value
                        Else
                            CategoryIDs = CategoryIDs & "," & item.Value
                        End If
                    End If
                Next
            Else
                CategoryIDs = ""
                For Each item As ListItem In lbConsumer.Items
                    If item.Selected Then
                        If CategoryIDs = "" Then
                            CategoryIDs = item.Value
                        Else
                            CategoryIDs = CategoryIDs & "," & item.Value
                        End If
                    End If
                Next
            End If

            If CategoryIDs <> "" Then
                Dim URL As String = Nothing
                Dim applicationRoot As String = Session("ApplicationRoot")
                Me.Body1.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", CategoryIDs))
            Else
                lblHeader1.ForeColor = Drawing.Color.Red
                lblHeader1.Text = "Please Select At Least One Category"
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.ToString
        End Try
    End Sub
End Class