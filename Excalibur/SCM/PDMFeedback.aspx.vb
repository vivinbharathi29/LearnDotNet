Imports System.Data
Partial Class PDMFeedback
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property AvId() As String
        Get
            Return Request("AvId")
        End Get
    End Property

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public Shared Property dtActionItems() As Data.DataTable
        Get
            Return (GetSessionStateValue("dtActionItems"))
        End Get
        Set(ByVal value As Data.DataTable)
            AddSessionStateValue("dtActionItems", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Dim AvNo As String = ""
            Dim Feedback As String = ""
            Dim dt As New DataTable
            Dim row As DataRow
            dt = dw.SelectAvActionItems(AvId)
            dtActionItems = dt
            For Each row In dt.Rows
                AvNo = row("AvNo")
                'If Not row("Feedback") Is DBNull.Value Then
                '    Feedback = row("PDMFeedback")
                '    Feedback = Feedback.Replace("]", " " & Environment.NewLine)
                '    Feedback = Feedback.Replace("[", "")
                '    row("AvNo") = Feedback
                'End If
                Exit For
            Next
            lblHeader.Text = "AV Action Item(s)"
            lblAvNoText.Text = AvNo
            gvAvActionItems.DataSource = dt
            gvAvActionItems.DataBind()
        Catch ex As Exception
            lblHeader.Text = ex.ToString
        End Try
    End Sub

    Protected Sub gvAvActionItems_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvAvActionItems.RowDataBound
        Try
            Dim i As Integer = 0
            Dim row As GridViewRow
            For Each row In gvAvActionItems.Rows
                Dim lblNotActionable As System.Web.UI.WebControls.Label = row.FindControl("lblNotActionable")
                If lblNotActionable.Text = "True" Then
                    row.BackColor = Drawing.Color.MistyRose
                End If
            Next
        Catch ex As Exception
            lblHeader.Text = ex.ToString
        End Try
    End Sub

    Protected Sub ClosePopup(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Response.Write("<script language='javascript'> { window.close();}</script>")
        Catch ex As Exception
            lblHeader.Text = ex.ToString
        End Try
    End Sub

    Protected Sub gvAvActionItems_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvAvActionItems.Sorting
        GridViewSortExpression = e.SortExpression
        gvAvActionItems.DataSource = SortDataTable(dtActionItems, False)
        gvAvActionItems.DataBind()
    End Sub

    Private Property GridViewSortDirection() As String
        Get
            Return IIf(ViewState("SortDirection") = Nothing, "DESC", ViewState("SortDirection"))
        End Get
        Set(ByVal value As String)
            ViewState("SortDirection") = value
        End Set
    End Property

    Private Property GridViewSortExpression() As String
        Get
            Return IIf(ViewState("SortExpression") = Nothing, String.Empty, ViewState("SortExpression"))
        End Get
        Set(ByVal value As String)
            ViewState("SortExpression") = value
        End Set
    End Property

    Private Function GetSortDirection() As String
        Select Case GridViewSortDirection
            Case "ASC"
                GridViewSortDirection = "DESC"
            Case "DESC"
                GridViewSortDirection = "ASC"
        End Select
        Return GridViewSortDirection
    End Function

    Protected Function SortDataTable(ByVal dataTable As Data.DataTable, ByVal isPageIndexChanging As Boolean) As Data.DataView
        If Not dataTable Is Nothing Then
            Dim dataView As New Data.DataView(dataTable)
            If GridViewSortExpression <> String.Empty Then
                If isPageIndexChanging Then
                    dataView.Sort = String.Format("{0} {1}", GridViewSortExpression, GridViewSortDirection)
                Else
                    dataView.Sort = String.Format("{0} {1}", GridViewSortExpression, GetSortDirection())
                End If
            End If
            Return dataView
        Else
            Return New Data.DataView
        End If
    End Function

End Class
