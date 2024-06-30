Imports System.Data

Partial Class Query_ScheduleSummary
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


    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
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

                gvScheduleSummary.DataSource = GetScheduleSummaryData()
                gvScheduleSummary.DataBind()
                lblLastRunDate.Text = Date.Now.ToLongDateString()
            Else
                Response.Write("<h1>You must enter this page through the Schedule Advanced Search & Report screen.</h1>")
                Response.End()
            End If
        End If
    End Sub

#Region " GetScheduleSummary Data "
    Private Function GetScheduleSummaryData() As DataTable
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtData As DataTable = dw.SelectScheduleSummaryData(ProductVersions, Partners, DevCenters, Programs, Status, Milestones)
        Dim dtSummaryData As DataTable = New DataTable()
        Dim dtSystemTeam As DataTable
        Dim sSystemManager As String, sConfigManager As String, sSepm As String
        Dim sSupplyChain As String, sPlatformDev As String, sService As String
        Dim sMarketing As String, sCommodities As String, sFinance As String

        dtSummaryData.Columns.Add("ID")
        dtSummaryData.Columns.Add("Schedule")
        dtSummaryData.Columns.Add("Series")
        dtSummaryData.Columns.Add("ODM")

        For Each row As DataRow In dtData.Rows
            If row("milestone_yn").ToString.ToLower() = "y" Then
                If Not dtSummaryData.Columns.Contains(row("item_description")) Then
                    dtSummaryData.Columns.Add(row("item_description").ToString(), Type.GetType("System.DateTime"))
                End If
            Else
                If Not dtSummaryData.Columns.Contains(row("item_description") & " start") Then
                    dtSummaryData.Columns.Add(row("item_description").ToString() & " start", Type.GetType("System.DateTime"))
                    dtSummaryData.Columns.Add(row("item_description").ToString() & " end", Type.GetType("System.DateTime"))
                End If
            End If
        Next

        If ShowSystemTeam Then
            dtSummaryData.Columns.Add("System Manager")
            dtSummaryData.Columns.Add("Configuration Manager")
            dtSummaryData.Columns.Add("SE PM")
            dtSummaryData.Columns.Add("Supply Chain")
            dtSummaryData.Columns.Add("Platform Development")
            dtSummaryData.Columns.Add("Service")
            dtSummaryData.Columns.Add("Marketing")
            dtSummaryData.Columns.Add("Commodity PM")
            dtSummaryData.Columns.Add("Finance")
        End If

        Dim newRow As DataRow
        Dim lastProductID As Integer = 0
        Dim lastScheduleID As Integer = 0
        Dim myDate As DateTime

        For Each row As DataRow In dtData.Rows
            If lastScheduleID <> row("schedule_id") Then
                If lastScheduleID <> 0 Then dtSummaryData.Rows.Add(newRow)
                lastProductID = row("ProductVersionID")
                lastScheduleID = row("schedule_id")
                newRow = dtSummaryData.NewRow()
                newRow("ID") = row("ProductVersionID")
                newRow("Schedule") = row("DotsName") & " - " & row("Schedule_Name")
                newRow("Series") = row("SeriesList")
                newRow("ODM") = row("PartnerName")

                If ShowSystemTeam Then
                    dtSystemTeam = dw.ListSystemTeam(lastProductID)
                    sSystemManager = String.Empty
                    sConfigManager = String.Empty
                    sMarketing = String.Empty
                    sSepm = String.Empty
                    sSupplyChain = String.Empty
                    sPlatformDev = String.Empty
                    sService = String.Empty
                    sCommodities = String.Empty
                    sFinance = String.Empty

                    For Each teamMember As DataRow In dtSystemTeam.Rows
                        Select Case teamMember("role")
                            Case "System Manager"
                                sSystemManager = teamMember("name") & "|" & teamMember("email")
                            Case "Configuration Manager"
                                sConfigManager = teamMember("name") & "|" & teamMember("email")
                            Case "Commercial Marketing"
                                sMarketing = teamMember("Name") & "|" & teamMember("email")
                            Case "Consumer Marketing"
                                sMarketing = teamMember("Name") & "|" & teamMember("email")
                            Case "SE PM"
                                sSepm = teamMember("Name") & "|" & teamMember("email")
                            Case "Supply Chain"
                                sSupplyChain = teamMember("Name") & "|" & teamMember("email")
                            Case "Platform Development"
                                sPlatformDev = teamMember("Name") & "|" & teamMember("email")
                            Case "Service"
                                sService = teamMember("Name") & "|" & teamMember("email")
                            Case "Commodity PM"
                                sCommodities = teamMember("Name") & "|" & teamMember("email")
                            Case "Finance"
                                sFinance = teamMember("Name") & "|" & teamMember("email")
                        End Select

                    Next
                End If
            End If 'End New Schedule

            If row("milestone_yn").ToString.ToLower = "y" Then
                If IsDBNull(row("actual_start_dt")) Then
                    If Date.TryParse(row("projected_start_dt").ToString(), myDate) Then
                        newRow(row("item_description")) = myDate.ToShortDateString()
                    End If
                Else
                    If Date.TryParse(row("actual_start_dt").ToString(), myDate) Then
                        newRow(row("item_description")) = myDate.ToShortDateString() & " 12:00 PM"
                    End If
                End If
            Else
                If IsDBNull(row("actual_start_dt")) Then
                    If Date.TryParse(row("projected_start_dt").ToString(), myDate) Then
                        newRow(row("item_description") & " start") = myDate.ToShortDateString()
                    End If
                Else
                    If Date.TryParse(row("actual_start_dt").ToString(), myDate) Then
                        newRow(row("item_description") & " start") = myDate.ToShortDateString() & " 12:00 PM"
                    End If
                End If
                If IsDBNull(row("actual_end_dt")) Then
                    If Date.TryParse(row("projected_end_dt").ToString(), myDate) Then
                        newRow(row("item_description") & " end") = myDate.ToShortDateString()
                    End If
                Else
                    If Date.TryParse(row("actual_end_dt").ToString(), myDate) Then
                        newRow(row("item_description") & " end") = myDate.ToShortDateString() & " 12:00 PM"
                    End If
                End If
            End If

            If ShowSystemTeam Then
                newRow("System Manager") = sSystemManager
                newRow("Configuration Manager") = sConfigManager
                newRow("Marketing") = sMarketing
                newRow("SE PM") = sSepm
                newRow("Supply Chain") = sSupplyChain
                newRow("Platform Development") = sPlatformDev
                newRow("Service") = sService
                newRow("Commodity PM") = sCommodities
                newRow("Finance") = sFinance
            End If
        Next

        If Not newRow Is Nothing Then
            dtSummaryData.Rows.Add(newRow)
        End If


        Return dtSummaryData

    End Function
#End Region

#Region " gvScheduleSummary Event Handlers "
    Protected Sub gvScheduleSummary_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvScheduleSummary.RowDataBound
        If e.Row.RowType = DataControlRowType.Header Then
            For Each cell As TableCell In e.Row.Cells
                cell.Style.Add("white-space", "wrap")
            Next
        End If
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim cellDate As Date
            Dim saTeamMember As String()
            If e.Row.Cells(0).Text.Length > 0 Then
                e.Row.Cells(0).Text = String.Format("<a href=""/pmview.asp?ID={0}&List=Schedule"" target=""_blank"">{0}</a>", e.Row.Cells(0).Text)
            End If
            e.Row.Cells(0).Style.Add("white-space", "nowrap")
            e.Row.Cells(1).Style.Add("white-space", "nowrap")
            e.Row.Cells(2).Style.Add("white-space", "nowrap")
            e.Row.Cells(3).Style.Add("white-space", "nowrap")
            Dim cellCount As Integer = e.Row.Cells.Count - 1
            If ShowSystemTeam Then
                cellCount = cellCount - 10
            End If
            If ShowSystemTeam Then
                For i As Integer = e.Row.Cells.Count - 9 To e.Row.Cells.Count - 1
                    e.Row.Cells(i).Style.Add("white-space", "nowrap")
                    If e.Row.Cells(i).Text.Length > 0 And e.Row.Cells(i).Text.Contains("|") Then
                        saTeamMember = e.Row.Cells(i).Text.Split("|")
                        e.Row.Cells(i).Text = String.Format("<a href=""mailto:{0}"">{1}</a>", saTeamMember(1), saTeamMember(0))
                    End If
                Next
            End If
            For i As Integer = 4 To cellCount
                If e.Row.Cells(i).Text.Length > 0 Then
                    If Date.TryParse(e.Row.Cells(i).Text, cellDate) Then
                        e.Row.Cells(i).Style.Add("text-align", "center")
                        e.Row.Cells(i).Text = cellDate.ToShortDateString()
                        If cellDate.TimeOfDay() = New TimeSpan(12, 0, 0) Then
                            e.Row.Cells(i).Style.Add("text-align", "center")
                            e.Row.Cells(i).BackColor = Drawing.Color.PaleGreen
                        ElseIf DateDiff(DateInterval.Day, cellDate, Date.Now) > 0 Then
                            e.Row.Cells(i).ForeColor = Drawing.Color.Red
                            e.Row.Cells(i).Font.Bold = True
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Protected Sub gvScheduleSummary_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvScheduleSummary.Sorting
        Dim dt As DataTable = GetScheduleSummaryData()

        dt.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
        gvScheduleSummary.DataSource = dt
        gvScheduleSummary.DataBind()

    End Sub

    Private Function GetSortDirection(ByVal column As String) As String

        ' By default, set the sort direction to ascending.
        Dim sortDirection = "ASC"

        ' Retrieve the last column that was sorted.
        Dim sortExpression = TryCast(ViewState("SortExpression"), String)

        If sortExpression IsNot Nothing Then
            ' Check if the same column is being sorted.
            ' Otherwise, the default value can be returned.
            If sortExpression = column Then
                Dim lastDirection = TryCast(ViewState("SortDirection"), String)
                If lastDirection IsNot Nothing _
                  AndAlso lastDirection = "ASC" Then

                    sortDirection = "DESC"

                End If
            End If
        End If

        ' Save new values in ViewState.
        ViewState("SortDirection") = sortDirection
        ViewState("SortExpression") = column

        Return sortDirection

    End Function

#End Region
End Class
