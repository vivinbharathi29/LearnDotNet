Imports System.Data
Partial Class SCM_FilterByCategory
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property Categories() As String
        Get
            Return Request("Categories")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim dtCategories As DataTable
                dtCategories = dw.SelectAvFeatureCategoriesFilter(BusinessID)
                lbCategories.DataSource = dtCategories
                lbCategories.DataBind()

                If Categories <> "" Then
                    'btnDeselect.Visible = True
                    Dim sCategories() As String = Categories.Split(",")
                    Dim item As ListItem
                    For Each item In lbCategories.Items
                        Dim i As Integer = 0
                        For i = 0 To sCategories.Length - 1
                            If sCategories(i) = item.Value Then
                                item.Selected = True
                            End If
                        Next
                    Next
                End If
            End If
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim sCategories As String = ""
            Dim item As ListItem
            Dim n As Integer = 0
            For Each item In lbCategories.Items
                If item.Selected Then
                    If sCategories = "" Then
                        sCategories = item.Value
                    Else
                        sCategories = sCategories & "," & item.Value
                    End If
                End If
            Next
            ProcessFilter(sCategories)
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Private Function ProcessFilter(ByVal sCategories As String) As Boolean
        Try
            Dim URL As String = Nothing
            'Dim applicationRoot As String = Session("ApplicationRoot")
            'If sCategories = "" Then
            'Response.Write("<script language='javascript'> {window.close();}</script>")
            'Else
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sCategories))
            'End If
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub btnDeselect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeselect.Click
        Dim item As ListItem
        For Each item In lbCategories.Items
            If item.Selected Then
                item.Selected = False
            End If
        Next
        'btnDeselect.Visible = False
    End Sub

    'Protected Sub lbCategories_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbCategories.SelectedIndexChanged
    '    Dim item As ListItem
    '    For Each item In lbCategories.Items
    '        If item.Selected Then
    '            btnDeselect.Visible = True
    '            Exit Sub
    '        End If
    '    Next
    '    btnDeselect.Visible = False
    'End Sub
End Class
