Imports System.Data

Partial Class Query_schedule
    Inherits BasePage

    '
    ' Report Profile Type 5
    '
    Private Const REPORT_PROFILE_TYPE_ID As Integer = 5
    Private Const REPORT_TRIGGER_TYPE_ID As Integer = 1
    Private _employeeID As Integer = 0
    Private ReadOnly Property EmployeeID() As Integer
        Get
            If _employeeID = 0 Then
                _employeeID = HPQ.Excalibur.Employee.GetUserID(Session("LoggedInUser"))
            End If
            Return _employeeID
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Response.CacheControl = "No-cache"

        If Not Page.IsPostBack Then
            FillReportProfiles()
            FillProducts()
            FillProductGroups()
            FillMilestoneList()
            lbAddProfile.Attributes.Add("onclick", "getProfileName();")
            lbRenameProfile.Attributes.Add("onclick", "getProfileName();")
            btnSummaryReport.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
            btnSummaryReport.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
            btnHistoryReport.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
            btnHistoryReport.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
            Dim currentUserName As String = Session("LoggedInUser").ToLower()
            btnReset.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
            btnReset.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
        End If
    End Sub

#Region " Fill Milestone List "
    Private Sub FillMilestoneList()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtMilestones As DataTable = dw.ListScheduleMilestones(hidProfileId.Value.Trim())

        rptMilestones.DataSource = dtMilestones
        rptMilestones.DataBind()
    End Sub

    Private Sub FillMilestoneList(ByVal selectedProductVerIds As String)
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtMilestones As DataTable = dw.ListScheduleMilestones(hidProfileId.Value.Trim(), selectedProductVerIds)

        rptMilestones.DataSource = dtMilestones
        rptMilestones.DataBind()
    End Sub

    Protected Sub rptMilestones_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptMilestones.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then

            Dim row As DataRowView = e.Item.DataItem
            Dim cbEnabled As CheckBox = e.Item.FindControl("cbEnabled")
            cbEnabled.Attributes.Add("TriggerID", row("ReportTriggerID").ToString())
            cbEnabled.Attributes.Add("MilestoneID", row("schedule_definition_data_id").ToString())
            cbEnabled.Checked = (row("ReportTriggerID") > 0)
            'cbEnabled.Attributes.Add("onclick", "javascript:alert(this.parentNode.id);")

            Dim cbSendEmail As CheckBox = e.Item.FindControl("cbSendEmail")
            cbSendEmail.Checked = row("SendEmail")

            Dim cbCreateAction As CheckBox = e.Item.FindControl("cbCreateAction")
            cbCreateAction.Checked = row("CreateActionItem")

            Dim tbDaysDiff As TextBox = e.Item.FindControl("tbDaysDiff")
            tbDaysDiff.Text = row("DaysDiff").ToString()
            tbDaysDiff.Attributes.Add("onkeydown", "if (event.keyCode==13){return false};")

            Dim tbNoteToSelf As TextBox = e.Item.FindControl("tbNoteToSelf")
            tbNoteToSelf.Text = row("NoteToSelf").ToString()
            tbNoteToSelf.Attributes.Add("onblur", "this.rows=2;")
            tbNoteToSelf.Attributes.Add("onfocus", "this.rows=8;")

            Dim tblRow As HtmlTableRow = e.Item.FindControl("tblRow")
            If CBool(row("active_yn_default")) Then
                tblRow.Attributes.Remove("class")
                tblRow.Attributes.Add("class", "td-DarkSeaGreen1")
            End If
        End If
    End Sub

    Protected Sub btnRefreshMilestones_Click(sender As Object, e As EventArgs) Handles btnRefreshMilestones.Click
        Dim selProdVIds As String
        selProdVIds = String.Empty
        Dim selProdGroupIds As String
        selProdGroupIds = String.Empty

        For Each selItem As ListItem In lbProductGroups.Items
            If selItem.Selected And selItem.Value.StartsWith("2:") Then
                selProdGroupIds = addToCommaSeparatedStrCollection(selProdGroupIds, selItem.Value.Replace("2:", ""))
            End If
        Next
        If Len(selProdGroupIds) > 0 Then
            selProdVIds = getProdVIDByGroupIds(selProdGroupIds)
        End If

        Dim counter As Integer
        counter = 0
        For Each selItem As ListItem In lbProducts.Items
            If selItem.Selected Then
                selProdVIds = addToCommaSeparatedStrCollection(selProdVIds, selItem.Value)
                counter = counter + 1
            End If
        Next

        If Len(selProdVIds) > 0 Then
            FillMilestoneList(selProdVIds)
        End If

    End Sub

    Private Function getProdVIDByGroupIds(commaSeparatedGIds As String) As String
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim productVersionIds As String = dw.GetProductVersionIdsByGroupIds(commaSeparatedGIds)
        Return productVersionIds
    End Function

    Private Function addToCommaSeparatedStrCollection(csStrCollection As String, newStrElement As String) As String
        If Len(csStrCollection) <= 0 Then
            Return newStrElement
        Else
            If InStr("," & csStrCollection & ",", "," & newStrElement & ",") < 1 Then
                Return csStrCollection & "," & newStrElement
            Else
                Return csStrCollection
            End If
        End If
    End Function

#End Region

#Region " Fill Products "
    Private Sub FillProducts()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtProducts As DataTable = dw.ListProducts("0", "5")

        Dim dv As DataView = dtProducts.DefaultView
        dv.Sort = "ProductVersionName asc"

        lbProducts.DataSource = dv
        lbProducts.DataTextField = "ProductVersionName"
        lbProducts.DataValueField = "ProductVersionID"
        lbProducts.DataBind()
    End Sub
#End Region

#Region " Fill Product Groups "
    Private Sub FillProductGroups()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtPartners As DataTable = dw.ListPartners(0)
        Dim dtPrograms As DataTable = dw.ListPrograms("")
        Dim dtDevCenters As DataTable = dw.ListDevCenters()
        Dim dtProductStatuses As DataTable = dw.ListProductStatuses()

        Dim dtProductGroups As DataTable = New DataTable()
        dtProductGroups.Columns.Add(New DataColumn("value"))
        dtProductGroups.Columns.Add(New DataColumn("text"))

        Dim newRow As DataRow = dtProductGroups.NewRow()
        newRow("value") = ""
        newRow("Text") = "------------ODM-------------"
        dtProductGroups.Rows.Add(newRow)
        For Each row As DataRow In dtPartners.Rows()
            newRow = dtProductGroups.NewRow()
            newRow("value") = String.Format("1:{0}", row("ID"))
            newRow("text") = row("Name").ToString().Trim()
            dtProductGroups.Rows.Add(newRow)
        Next

        newRow = dtProductGroups.NewRow()
        newRow("value") = ""
        newRow("Text") = "-----------Cycle------------"
        dtProductGroups.Rows.Add(newRow)
        For Each row As DataRow In dtPrograms.Rows()
            newRow = dtProductGroups.NewRow()
            newRow("value") = String.Format("2:{0}", row("ID"))
            newRow("text") = row("FullName").ToString().Trim()
            'If row("BusinessID") = 2 Then
            '    newRow("text") = String.Format("CNB {0}", row("Name").ToString().Trim())
            'Else
            '    newRow("text") = String.Format("BNB {0}", row("Name").ToString().Trim())
            'End If
            dtProductGroups.Rows.Add(newRow)
        Next

        newRow = dtProductGroups.NewRow()
        newRow("value") = ""
        newRow("Text") = "-------Dev. Center---------"
        dtProductGroups.Rows.Add(newRow)
        For Each row As DataRow In dtDevCenters.Rows()
            newRow = dtProductGroups.NewRow()
            newRow("value") = String.Format("3:{0}", row("ID"))
            newRow("text") = row("Name").ToString().Trim()
            dtProductGroups.Rows.Add(newRow)
        Next

        newRow = dtProductGroups.NewRow()
        newRow("value") = ""
        newRow("Text") = "-----Product Phase-----"
        dtProductGroups.Rows.Add(newRow)
        For Each row As DataRow In dtProductStatuses.Rows()
            newRow = dtProductGroups.NewRow()
            newRow("value") = String.Format("4:{0}", row("ID"))
            newRow("text") = row("Name").ToString().Trim()
            dtProductGroups.Rows.Add(newRow)
        Next

        lbProductGroups.DataSource = dtProductGroups.DefaultView()
        lbProductGroups.DataTextField = "text"
        lbProductGroups.DataValueField = "value"
        lbProductGroups.DataBind()

    End Sub
#End Region

#Region " Fill Report Profiles "
    Private Sub FillReportProfiles()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtReportProfiles As DataTable = dw.ListReportProfiles(EmployeeID, REPORT_PROFILE_TYPE_ID)
        Dim dtReportProfilesShared As DataTable = dw.ListReportProfilesShared(EmployeeID, REPORT_PROFILE_TYPE_ID)
        Dim dtReportProfilesGroupShared As DataTable = dw.ListReportProfilesGroupShared(EmployeeID, REPORT_PROFILE_TYPE_ID)

        Dim dtReportProfileList As DataTable = New DataTable()
        dtReportProfileList.Columns.Add(New DataColumn("value"))
        dtReportProfileList.Columns.Add(New DataColumn("text"))

        Dim newRow As DataRow = dtReportProfileList.NewRow()
        newRow("value") = ""
        newRow("text") = "Use Options Selected Below"
        dtReportProfileList.Rows.Add(newRow)

        If dtReportProfiles.Rows.Count > 0 Then
            newRow = dtReportProfileList.NewRow()
            newRow("value") = ""
            newRow("text") = "----------------------------------------"
            dtReportProfileList.Rows.Add(newRow)

            For Each row As DataRow In dtReportProfiles.Rows()
                newRow = dtReportProfileList.NewRow()
                newRow("value") = row("ID")
                newRow("text") = row("ProfileName")
                dtReportProfileList.Rows.Add(newRow)
            Next
        End If

        'newRow = dtReportProfileList.NewRow()
        'newRow("value") = "S:1"
        'newRow("text") = "My Test Profile"
        'dtReportProfileList.Rows.Add(newRow)

        If dtReportProfilesShared.Rows.Count > 0 Then
            newRow = dtReportProfileList.NewRow()
            newRow("value") = ""
            newRow("text") = "----------- Shared Profiles -----------"
            dtReportProfileList.Rows.Add(newRow)

            For Each row As DataRow In dtReportProfilesShared.Rows()
                newRow = dtReportProfileList.NewRow()
                newRow("value") = String.Format("S:{0}", row("ID"))
                newRow("text") = row("ProfileName")
                dtReportProfileList.Rows.Add(newRow)
            Next
        End If

        If dtReportProfilesGroupShared.Rows.Count > 0 Then
            newRow = dtReportProfileList.NewRow()
            newRow("value") = ""
            newRow("text") = "----------- Group Shared Profiles -----------"
            dtReportProfileList.Rows.Add(newRow)

            For Each row As DataRow In dtReportProfilesGroupShared.Rows()
                newRow = dtReportProfileList.NewRow()
                newRow("value") = String.Format("G:{0}", row("ID"))
                newRow("text") = row("ProfileName")
                dtReportProfileList.Rows.Add(newRow)
            Next
        End If

        ddlReportProfiles.DataSource = dtReportProfileList
        ddlReportProfiles.DataValueField = "value"
        ddlReportProfiles.DataTextField = "text"
        ddlReportProfiles.DataBind()
    End Sub
#End Region

#Region " Report Profiles Index Changed "
    Protected Sub ddlReportProfiles_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlReportProfiles.SelectedIndexChanged
        lbUpdateProfile.Visible = False
        lbDeleteProfile.Visible = False
        lbRenameProfile.Visible = False
        lbRemoveProfile.Visible = False
        lbShareProfile.Visible = False
        lblProfileOwnerHdr.Visible = False
        lblProfileOwnerName.Visible = False


        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

        If ddlReportProfiles.SelectedValue <> String.Empty Then
            Dim dt As DataTable
            Dim dr As DataRow
            If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
                dt = dw.GetReportProfileShared(ddlReportProfiles.SelectedValue.Substring(2), EmployeeID)
                dr = dt.Rows(0)

                If dr("CanEdit") Then
                    lbUpdateProfile.Visible = True
                    lbRenameProfile.Visible = True
                End If
                If dr("CanDelete") Then
                    lbDeleteProfile.Visible = True
                End If
                lbRemoveProfile.Visible = True

                lblProfileOwnerHdr.Visible = True
                lblProfileOwnerName.Visible = True
                lblProfileOwnerName.Text = dr("PrimaryOwner")

            ElseIf ddlReportProfiles.SelectedValue.Substring(0, 1) = "G" Then
                dt = dw.GetReportProfileGroup(ddlReportProfiles.SelectedValue.Substring(2))
                dr = dt.Rows(0)

                If dr("CanEdit") Then
                    lbUpdateProfile.Visible = True
                    lbRenameProfile.Visible = True
                End If
                If dr("CanDelete") Then
                    lbDeleteProfile.Visible = True
                End If
                'lbRemoveProfile.Visible = True

                lblProfileOwnerHdr.Visible = True
                lblProfileOwnerName.Visible = True
                lblProfileOwnerName.Text = dr("PrimaryOwner")

                dt = dw.GetReportProfile(ddlReportProfiles.SelectedValue.Substring(2))
                dr = dt.Rows(0)
            Else
                dt = dw.GetReportProfile(ddlReportProfiles.SelectedValue.ToString())
                dr = dt.Rows(0)

                lbUpdateProfile.Visible = True
                lbDeleteProfile.Visible = True
                lbRenameProfile.Visible = True
                lbShareProfile.Visible = True
            End If

            hidProducts.Value = dr("Value15").ToString()
            hidGroups.Value = dr("Value45").ToString()
            hidProfileName.Value = dr("ProfileName").ToString()
            hidProfileId.Value = dr("ID").ToString()
            cbIncludeSystemTeam.Checked = dr("Value17")
            LoadProductSavedValues()
            LoadProductGroupsSavedValues()
            FillMilestoneList()
        End If
    End Sub
#End Region

#Region " Load Product Saved Values "
    Private Sub LoadProductSavedValues()
        lbProducts.SelectedIndex = -1
        If hidProducts.Value.Trim <> String.Empty Then
            Dim saSelectedItems As String() = hidProducts.Value.Split("|")

            For Each value As String In saSelectedItems
                lbProducts.Items.FindByValue(value).Selected = True
            Next
        End If
    End Sub
#End Region

#Region " Load Product Group Saved Values "
    Private Sub LoadProductGroupsSavedValues()
        lbProductGroups.SelectedIndex = -1
        If hidGroups.Value.Trim <> String.Empty Then
            Dim saSelectedItems As String() = hidGroups.Value.Split("|")

            For Each value As String In saSelectedItems
                lbProductGroups.Items.FindByValue(value).Selected = True
            Next
        End If
    End Sub
#End Region

#Region " btnReset Events "
    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Server.Transfer("schedule.aspx")
    End Sub
#End Region

#Region " Profile Link Button Click Event Handlers "
    Protected Sub lbAddProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbAddProfile.Click
        GetSelectedItems()
        CreateProfile()
        SaveReportTriggers()
    End Sub

    Protected Sub lbUpdateProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbUpdateProfile.Click
        GetSelectedItems()
        UpdateProfile()
        SaveReportTriggers()
        FillMilestoneList()
    End Sub

    Protected Sub lbDeleteProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbDeleteProfile.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim profileID As String

        If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
            profileID = ddlReportProfiles.SelectedValue.Substring(2)
        Else
            profileID = ddlReportProfiles.SelectedValue.ToString()
            dw.DeleteReportProfile(profileID)
        End If

        dw.DeleteReportTriggerByProfileID(profileID)

        'FillReportProfiles()
        Server.Transfer("schedule.aspx")
    End Sub

    Protected Sub lbRemoveProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbRemoveProfile.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim profileID As String

        If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
            profileID = ddlReportProfiles.SelectedValue.Substring(2)
            dw.RemoveReportProfile(profileID, EmployeeID)
        End If

        'FillReportProfiles()
        Server.Transfer("schedule.aspx")

    End Sub

    Protected Sub lbRenameProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbRenameProfile.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim profileSelectedValue As String = ddlReportProfiles.SelectedValue
        Dim profileID As String

        If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
            profileID = ddlReportProfiles.SelectedValue.Substring(2)
        Else
            profileID = ddlReportProfiles.SelectedValue.ToString()
        End If

        dw.RenameProfile(profileID, hidProfileName.Value)

        FillReportProfiles()

        ddlReportProfiles.SelectedValue = profileSelectedValue

    End Sub

#End Region

#Region " Profile Management Support "
    Private Sub CreateProfile()
        If hidProfileName.Value.Trim() <> String.Empty Then
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim newID As Long = dw.AddReportProfile(hidProfileName.Value, REPORT_PROFILE_TYPE_ID.ToString(), EmployeeID, hidProducts.Value, hidGroups.Value, cbIncludeSystemTeam.Checked.ToString())
            FillReportProfiles()
            ddlReportProfiles.SelectedValue = newID
            hidProfileId.Value = newID.ToString()

            lbUpdateProfile.Visible = True
            lbDeleteProfile.Visible = True
            lbRenameProfile.Visible = True
            lbShareProfile.Visible = True

        End If
    End Sub

    Private Sub UpdateProfile()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim profileID As String

        If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
            profileID = ddlReportProfiles.SelectedValue.Substring(2)
        Else
            profileID = ddlReportProfiles.SelectedValue.ToString()
        End If

        dw.UpdateProfile(profileID, hidProducts.Value, hidGroups.Value, cbIncludeSystemTeam.Checked.ToString())

    End Sub

    Private Sub GetSelectedItems()
        Dim sbSelectedPrograms As Text.StringBuilder = New Text.StringBuilder
        For Each item As ListItem In lbProducts.Items
            If item.Selected Then
                sbSelectedPrograms.Append(String.Format("{0}|", item.Value))
            End If
        Next
        If sbSelectedPrograms.Length > 0 Then sbSelectedPrograms.Remove(sbSelectedPrograms.Length - 1, 1)
        hidProducts.Value = sbSelectedPrograms.ToString()

        Dim sbSelectedProductGroups As Text.StringBuilder = New Text.StringBuilder
        For Each item As ListItem In lbProductGroups.Items
            If item.Selected And item.Value <> String.Empty Then
                sbSelectedProductGroups.Append(String.Format("{0}|", item.Value))
            End If
        Next
        If sbSelectedProductGroups.Length > 0 Then sbSelectedProductGroups.Remove(sbSelectedProductGroups.Length - 1, 1)
        hidGroups.Value = sbSelectedProductGroups.ToString()
    End Sub

    Private Sub SaveReportTriggers()
        For Each oItem As RepeaterItem In rptMilestones.Items
            Dim cbEnabled As CheckBox = oItem.FindControl("cbEnabled")
            Dim cbSendEmail As CheckBox = oItem.FindControl("cbSendEmail")
            Dim cbCreateAction As CheckBox = oItem.FindControl("cbCreateAction")
            Dim tbDaysDiff As TextBox = oItem.FindControl("tbDaysDiff")
            Dim tbNoteToSelf As TextBox = oItem.FindControl("tbNoteToSelf")
            Dim triggerID As Long = Long.Parse(cbEnabled.Attributes("TriggerID").ToString())
            Dim milestoneID As String = cbEnabled.Attributes("MilestoneID")
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim sendEmail As String = IIf(cbSendEmail.Checked, "1", "0")
            Dim createAction As String = IIf(cbCreateAction.Checked, "1", "0")
            Dim CreateEmailJob As Boolean = False
            Dim curUserName As String = HPQ.Excalibur.Employee.GetUserName(Session("LoggedInUser"))

            'Response.Write(String.Format("{0}:{1}:{2}<br />", triggerID, cbSendEmail.Checked, cbCreateAction.Checked))

            If triggerID > 0 And (Not cbEnabled.Checked) Then
                dw.DeleteReportTrigger(triggerID.ToString())
            ElseIf cbEnabled.Checked Then
                If sendEmail Then
                    CreateEmailJob = True
                End If
                triggerID = dw.UpdateReportTrigger(REPORT_TRIGGER_TYPE_ID, milestoneID, hidProfileId.Value, tbDaysDiff.Text, sendEmail, createAction, tbNoteToSelf.Text, curUserName)
                dw.DeleteReportTriggerProductVersionByTriggerID(triggerID)
                SaveProductRelations(triggerID)
                SaveGroupRelations(triggerID)
            End If

            If CreateEmailJob Then
                'dw.InsertJobDefinition4ScheduleEmailAlert(HPQ.Excalibur.Employee.GetUserID(Session("LoggedInUser")))
            End If
        Next
    End Sub

    Private Sub SaveProductRelations(ByVal TriggerID As String)
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

        For Each item As ListItem In lbProducts.Items
            If item.Selected Then
                dw.InsertReportTriggerProductVersion(TriggerID, item.Value, String.Empty, String.Empty, String.Empty, String.Empty)
            End If
        Next
    End Sub

    Private Sub SaveGroupRelations(ByVal TriggerID As String)
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        For Each item As ListItem In lbProductGroups.Items
            If item.Selected And item.Value <> String.Empty Then
                Select Case item.Value.Substring(0, 1)
                    Case 1
                        dw.InsertReportTriggerProductVersion(TriggerID, String.Empty, item.Value.Substring(2), String.Empty, String.Empty, String.Empty)
                        'dw.InsertReportTriggerProductVersionByGroup(TriggerID, item.Value.Substring(2), String.Empty, String.Empty)
                    Case 2
                        dw.InsertReportTriggerProductVersion(TriggerID, String.Empty, String.Empty, item.Value.Substring(2), String.Empty, String.Empty)
                        'dw.InsertReportTriggerProductVersionByProgram(TriggerID, item.Value.Substring(2))
                    Case 3
                        dw.InsertReportTriggerProductVersion(TriggerID, String.Empty, String.Empty, String.Empty, item.Value.Substring(2), String.Empty)
                        'dw.InsertReportTriggerProductVersionByGroup(TriggerID, String.Empty, item.Value.Substring(2), String.Empty)
                    Case 4
                        dw.InsertReportTriggerProductVersion(TriggerID, String.Empty, String.Empty, String.Empty, String.Empty, item.Value.Substring(2))
                        'dw.InsertReportTriggerProductVersionByGroup(TriggerID, String.Empty, String.Empty, item.Value.Substring(2))
                End Select
            End If
        Next

    End Sub
#End Region

End Class
