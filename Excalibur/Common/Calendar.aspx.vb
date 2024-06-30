Imports System.Configuration
Imports System.IO

Partial Class Common_Calendar
    Inherits System.Web.UI.Page

    Property StartDate() As String
        Set(ByVal value As String)
            ViewState("StartDate") = value
        End Set
        Get
            Return CStr(ViewState("StartDate"))
        End Get
    End Property

    Property EndDate() As String
        Set(ByVal value As String)
            ViewState("EndDate") = value
        End Set
        Get
            Return CStr(ViewState("EndDate"))
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                StartDate = Request.QueryString("StartDate")
                EndDate = Request.QueryString("EndDate")
                ctl.Value = Request.QueryString("ctl")
            End If
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Protected Sub Calendar1_DayRender(sender As Object, e As DayRenderEventArgs)
        Dim dEndDate As Date
        Dim dStartDate As Date
        Dim hl As New HyperLink()
        If Date.TryParse(EndDate, dEndDate) = True Then
            If e.Day.Date <= dEndDate Then
                hl.Text = CType(e.Cell.Controls(0), LiteralControl).Text
                hl.NavigateUrl = "javascript:SetDate('" & e.Day.Date.ToShortDateString() & "');"
                e.Cell.Controls.Clear()
                e.Cell.Controls.Add(hl)
            End If
        ElseIf Date.TryParse(StartDate, dStartDate) = True Then
            If e.Day.Date >= dStartDate Then
                hl.Text = CType(e.Cell.Controls(0), LiteralControl).Text
                hl.NavigateUrl = "javascript:SetDate('" & e.Day.Date.ToShortDateString() & "');"
                e.Cell.Controls.Clear()
                e.Cell.Controls.Add(hl)
            End If
        Else
            hl.Text = CType(e.Cell.Controls(0), LiteralControl).Text
            hl.NavigateUrl = "javascript:SetDate('" & e.Day.Date.ToShortDateString() & "');"
            e.Cell.Controls.Clear()
            e.Cell.Controls.Add(hl)
        End If
    End Sub

End Class
