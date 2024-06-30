Imports System.Data

Partial Class Query_ScheduleChanges
    Inherits System.Web.UI.Page

    Private Property ProductVersions() As String
        Get
            Return ViewState("ProductVersions")
        End Get
        Set(ByVal value As String)
            ViewState("ProductVersions") = value
        End Set
    End Property

    Private Property Partners() As String
        Get
            Return ViewState("Partners")
        End Get
        Set(ByVal value As String)
            ViewState("Partners") = value
        End Set
    End Property

    Private Property Programs() As String
        Get
            Return ViewState("Programs")
        End Get
        Set(ByVal value As String)
            ViewState("Programs") = value
        End Set
    End Property

    Private Property DevCenters() As String
        Get
            Return ViewState("DevCenters")
        End Get
        Set(ByVal value As String)
            ViewState("DevCenters") = value
        End Set
    End Property

    Private Property Status() As String
        Get
            Return ViewState("Status")
        End Get
        Set(ByVal value As String)
            ViewState("Status") = value
        End Set
    End Property

    Private Property Milestones() As String
        Get
            Return ViewState("Milestones")
        End Get
        Set(ByVal value As String)
            ViewState("Milestones") = value
        End Set
    End Property

    Private Property ShowSystemTeam() As Boolean
        Set(ByVal value As Boolean)
            ViewState("ShowSystemTeam") = value
        End Set
        Get
            Return ViewState("ShowSystemTeam")
        End Get
    End Property

    Private _lastProductName As String = String.Empty
    Private _lastScheduleName As String = String.Empty
    Private _lastMilestoneName As String = String.Empty

    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
        Response.CacheControl = "No-cache"

        If Not Page.IsPostBack Then
            If Not PreviousPage Is Nothing Then
                Dim sbProductVersions As StringBuilder = New StringBuilder()
                Dim sbPartners As StringBuilder = New StringBuilder()
                Dim sbPrograms As StringBuilder = New StringBuilder()
                Dim sbDevCenters As StringBuilder = New StringBuilder()
                Dim sbStatus As StringBuilder = New StringBuilder()
                Dim sbMilestones As StringBuilder = New StringBuilder()

                Dim lbProducts As ListBox = PreviousPage.FindControl("lbProducts")
                Dim lbProductGroups As ListBox = PreviousPage.FindControl("lbProductGroups")
                Dim rptMilestones As Repeater = PreviousPage.FindControl("rptMilestones")
                Dim cbShowSystemTeam As CheckBox = PreviousPage.FindControl("cbIncludeSystemTeam")
                Dim ddlReportFormat As DropDownList = PreviousPage.FindControl("ddlReportFormat")

                Select Case ddlReportFormat.SelectedValue
                    Case 1
                        Response.ContentType = "application/vnd.ms-excel"
                    Case 2
                        Response.ContentType = "application/msword"
                End Select

                ShowSystemTeam = cbShowSystemTeam.Checked

                For Each oItem As RepeaterItem In rptMilestones.Items
                    Dim cbEnabled As CheckBox = oItem.FindControl("cbEnabled")
                    Dim milestoneID As String = cbEnabled.Attributes("MilestoneID")

                    If cbEnabled.Checked Then sbMilestones.Append(milestoneID & ",")
                Next

                For Each item As ListItem In lbProducts.Items
                    If item.Selected And item.Value <> String.Empty Then
                        sbProductVersions.Append(item.Value & ",")
                    End If
                Next

                For Each item As ListItem In lbProductGroups.Items
                    If item.Selected And item.Value <> String.Empty Then
                        Select Case item.Value.Substring(0, 1)
                            Case 1
                                sbPartners.Append(item.Value.Substring(2) & ",")
                            Case 2
                                sbPrograms.Append(item.Value.Substring(2) & ",")
                            Case 3
                                sbDevCenters.Append(item.Value.Substring(2) & ",")
                            Case 4
                                sbStatus.Append(item.Value.Substring(2) & ",")
                        End Select
                    End If
                Next

                If sbProductVersions.Length > 0 Then sbProductVersions.Remove(sbProductVersions.Length - 1, 1)
                If sbPartners.Length > 0 Then sbPartners.Remove(sbPartners.Length - 1, 1)
                If sbPrograms.Length > 0 Then sbPrograms.Remove(sbPrograms.Length - 1, 1)
                If sbDevCenters.Length > 0 Then sbDevCenters.Remove(sbDevCenters.Length - 1, 1)
                If sbStatus.Length > 0 Then sbStatus.Remove(sbStatus.Length - 1, 1)
                If sbMilestones.Length > 0 Then sbMilestones.Remove(sbMilestones.Length - 1, 1)

                ProductVersions = sbProductVersions.ToString()
                Partners = sbPartners.ToString()
                Programs = sbPrograms.ToString()
                DevCenters = sbDevCenters.ToString()
                Status = sbStatus.ToString()
                Milestones = sbMilestones.ToString()

                Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
                Dim dt As DataTable = dw.SelectScheduleHistoryData(ProductVersions, Partners, DevCenters, Programs, Status, Milestones)

                rptrHistoryDetails.DataSource = dt
                rptrHistoryDetails.DataBind()
                lblLastRunDate.Text = Date.Now.ToLongDateString()
            Else
                Response.Write("<h1>You must enter this page through the Schedule Advanced Search & Report screen.</h1>")
                Response.End()
            End If
        End If
    End Sub

    Protected Sub rptrHistoryDetails_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptrHistoryDetails.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim dRow As DataRowView = e.Item.DataItem
            Dim tr As HtmlTableRow = New HtmlTableRow()
            Dim cell As HtmlTableCell = New HtmlTableCell()
            If dRow("DotsName").ToString.Trim() <> _lastProductName Then
                _lastProductName = dRow("DotsName").ToString.Trim()
                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.ColSpan = 11
                cell.InnerHtml = "&nbsp;"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.ColSpan = 11
                cell.InnerHtml = String.Format("<span style=""font-face:verdana;font-size:medium;font-weight:bold;"">{0}</span>", dRow("DotsName").ToString.Trim())
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)
            End If
            If dRow("schedule_name").ToString.Trim() <> _lastScheduleName Then
                _lastScheduleName = dRow("schedule_name").ToString.Trim()
                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.ColSpan = 11
                cell.InnerHtml = "&nbsp;"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader")
                cell.ColSpan = 11
                cell.InnerHtml = String.Format("<span style=""font-face:verdana;font-size:small;font-weight:bold;"">Schedule:&nbsp;{0}</span>", dRow("schedule_name").ToString.Trim())
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)
            End If
            If dRow("item_description").ToString.Trim() <> _lastMilestoneName Then
                _lastMilestoneName = dRow("item_description").ToString.Trim()
                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.ColSpan = 11
                cell.InnerHtml = "&nbsp;"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.ColSpan = 11
                cell.InnerHtml = String.Format("<span style=""font-face:verdana;font-size:x-small;font-weight:bold;"">{0}</span>", dRow("item_description").ToString.Trim())
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow()
                cell = New HtmlTableCell
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 4
                cell.InnerHtml = "Projected"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 4
                cell.InnerHtml = "Actual"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.RowSpan = 3
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Change Notes"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.RowSpan = 3
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Changed By"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.RowSpan = 3
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Change Date"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow
                cell = New HtmlTableCell
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 2
                cell.InnerHtml = "Start"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 2
                cell.InnerHtml = "End"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 2
                cell.InnerHtml = "Start"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell
                cell.Attributes.Add("class", "TableHeader1")
                cell.ColSpan = 2
                cell.InnerHtml = "End"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)

                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Old"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "New"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Old"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "New"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Old"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "New"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "Old"
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableHeader1")
                cell.InnerHtml = "New"
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)
            End If
            If Not (dRow("old_projected_start_dt").ToString = String.Empty And _
                dRow("new_projected_start_dt").ToString = String.Empty And _
                dRow("old_projected_end_dt").ToString = String.Empty And _
                dRow("new_projected_end_dt").ToString = String.Empty And _
                dRow("old_actual_start_dt").ToString = String.Empty And _
                dRow("new_actual_start_dt").ToString = String.Empty And _
                dRow("old_actual_end_dt").ToString = String.Empty And _
                dRow("new_actual_end_dt").ToString = String.Empty) Then
                tr = New HtmlTableRow()
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("old_projected_start_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("new_projected_start_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("old_projected_end_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("new_projected_end_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("old_actual_start_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("new_actual_start_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("old_actual_end_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = String.Format("{0:d}", dRow("new_actual_end_dt"))
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellLeft")
                cell.InnerHtml = dRow("notes").ToString.Trim()
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellLeft")
                cell.InnerHtml = dRow("name").ToString.Trim()
                tr.Cells.Add(cell)
                cell = New HtmlTableCell()
                cell.Attributes.Add("class", "TableCellCentered")
                cell.InnerHtml = dRow("last_upd_date").ToString()
                tr.Cells.Add(cell)
                e.Item.Controls.Add(tr)
            End If
        End If
    End Sub
End Class
