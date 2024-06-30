Imports System.Data
Imports System.Data.SqlClient

Partial Class Service_SpareKitMapMain3
    Inherits System.Web.UI.Page

#Region " Page Properties "
    Public ReadOnly Property ProductBrandId() As String
        Get
            Return Request.QueryString("PBID")
        End Get
    End Property

    Public ReadOnly Property ProductVersionId() As String
        Get
            Return Request.QueryString("PVID")
        End Get
    End Property

    Public ReadOnly Property SpareKitId() As String
        Get
            Return Request.QueryString("SKID")
        End Get
    End Property

    Dim _categoryId As Integer = 0
    Public Property CategoryID() As Integer
        Set(ByVal value As Integer)
            _categoryId = value
        End Set
        Get
            Return _categoryId
        End Get
    End Property

    Private blnOSSPUser As Boolean = False

    Public Property OSSPUser() As Boolean
        Get
            Return blnOSSPUser
        End Get
        Set(ByVal value As Boolean)
            blnOSSPUser = value
        End Set
    End Property


    Dim _editMode As Boolean = False
    Private Property EditMode() As Boolean
        Get

            If Not (ViewState("EditMode") = Nothing) Then
                _editMode = ViewState("EditMode")
            End If
            Return _editMode
        End Get
        Set(ByVal value As Boolean)
            _editMode = value
            ViewState("EditMode") = _editMode
        End Set
    End Property

    Dim _SpareKitMapId As String = String.Empty
    Public Property SpareKitMapId() As String
        Get
            If _SpareKitMapId = String.Empty Then _SpareKitMapId = ViewState("SpareKitMapId")
            If _SpareKitMapId = String.Empty Then _SpareKitMapId = Request.QueryString("MapId")
            Return _SpareKitMapId
        End Get
        Set(ByVal value As String)
            _SpareKitMapId = value
            ViewState("SpareKitMapId") = _SpareKitMapId
        End Set
    End Property

    Dim _selections As String
    Dim blnSelectAll As Boolean

    Private Property Selections() As String
        Get
            Return _selections
        End Get
        Set(ByVal value As String)
            _selections = value
        End Set
    End Property

#End Region

#Region " Page Events "
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoServerCaching()


        If Not Page.IsPostBack Then
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
            If SpareKitMapId = 0 Then
                'Get New SpareKitMapId
                SpareKitMapId = dw.GetNewSpareKitMapId(ProductBrandId, SpareKitId)
            Else
                EditMode = True
            End If

            Dim dt As DataTable = dw.SelectServiceSpareKitAvMap(ProductBrandId, SpareKitMapId)
            ViewState("mapTable") = dt
            ViewState("orgMapTable") = dt
            mapGrid.DataSource = dt
            mapGrid.DataBind()

        End If

        If body.Attributes.Item("onload") <> String.Empty Then
            body.Attributes.Item("onload") = String.Empty
        End If

        Me.OSSPUser = IsOSSPUser()

        If (Me.OSSPUser) Then
            Me.btnSave.Visible = False
            Me.btnSave.Enabled = False
            Me.btnCancel.Text = "Close"
        End If

        If pnlFind.Visible Then
            Me.Selections = getSelections()

            Try
                blnSelectAll = CType(CType(ViewState("SelectAll"), String), Boolean)
            Catch ex As Exception
                blnSelectAll = False
            End Try

        End If

    End Sub
#End Region

#Region " Support Functions "

    Private Function IsOSSPUser() As Boolean
        Dim blnResult As Boolean = False
        Dim strCurrentUser As String = LCase(Trim(Session("LoggedInUser")))
        Dim strCurrentDomain As String = ""
        Dim intCurrentPartnerTypeID As Integer
        Dim objRow As DataRow

        If InStr(strCurrentUser, "\") > 0 Then
            strCurrentDomain = Left(strCurrentUser, InStr(strCurrentUser, "\") - 1)
            strCurrentUser = Mid(strCurrentUser, InStr(strCurrentUser, "\") + 1)
        End If

        Dim dw As HPQ.Data.DataWrapper = New HPQ.Data.DataWrapper()
        Dim cmd As SqlCommand

        cmd = dw.CreateCommand("usp_GetUserType", CommandType.StoredProcedure)

        dw.CreateParameter(cmd, "@UserName", SqlDbType.VarChar, strCurrentUser, 30)
        dw.CreateParameter(cmd, "@Domain", SqlDbType.VarChar, strCurrentDomain, 30)

        Dim dt As DataTable = dw.ExecuteCommandTable(cmd)

        Try
            If (Not dt Is Nothing) Then
                objRow = dt.Rows(0)

                If (IsNumeric(objRow("PartnerTypeID"))) Then
                    intCurrentPartnerTypeID = CInt(objRow("PartnerTypeID"))

                    If (intCurrentPartnerTypeID = 2) Then
                        blnResult = True
                    End If
                End If
            End If

        Catch ex As Exception
        Finally
            If (Not dt Is Nothing) Then
                dt.Dispose()
                dt = Nothing
            End If
            cmd = Nothing
            dw = Nothing
        End Try


        IsOSSPUser = blnResult
    End Function

    Private Sub setDirtyFlag()

        Dim objDTOrg As DataTable = ViewState("orgMapTable")
        Dim objDTCurr As DataTable = ViewState("mapTable")
        Dim intOrgRowCount As Integer = objDTOrg.Rows.Count
        Dim intCurrRowCount As Integer = objDTCurr.Select("", "", DataViewRowState.CurrentRows).Length
        Dim intNumMatches As Integer = 0

        If intOrgRowCount > 0 And intCurrRowCount > 0 Then
            For Each objRow As DataRow In objDTCurr.Rows

                If objRow.RowState <> DataRowState.Deleted Then

                    If (objDTOrg.Select("AvNo='" & objRow("AvNo") & "'").Length > 0) Then
                        intNumMatches = intNumMatches + 1
                    End If

                End If

            Next

            If (intNumMatches = intCurrRowCount) And (intCurrRowCount = intOrgRowCount) Then
                Me.hidDirtyFlag.Value = "false"
            Else
                Me.hidDirtyFlag.Value = "true"
            End If

        ElseIf intOrgRowCount <> intCurrRowCount Then
            Me.hidDirtyFlag.Value = "true"
        Else
            Me.hidDirtyFlag.Value = "false"
        End If
    End Sub
#End Region

#Region " Grid Events "
    Protected Sub mapGrid_EditCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles mapGrid.EditCommand
        mapGrid.ShowFooter = False
        mapGrid.EditItemIndex = e.Item.ItemIndex
        mapGrid.DataSource = ViewState("mapTable")
        mapGrid.DataBind()
    End Sub

    Protected Sub mapGrid_UpdateCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles mapGrid.UpdateCommand
        Dim hidAvNo As HiddenField = e.Item.FindControl("hidAvNo")
        Dim txtAvNo As TextBox = e.Item.FindControl("txtAvNo")
        Dim ddlAvCategory As DropDownList = e.Item.FindControl("ddlAvCategory")

        Dim dt As DataTable = ViewState("mapTable")

        For i As Integer = dt.Rows.Count - 1 To 0 Step -1
            If dt.Rows(i).RowState <> DataRowState.Deleted Then
                If dt.Rows(i)("AvNo").ToString().Trim.ToUpper = hidAvNo.Value.Trim.ToUpper Then
                    dt.Rows(i)("AvCategoryId") = ddlAvCategory.SelectedValue
                    dt.Rows(i)("AvCategoryName") = ddlAvCategory.SelectedItem.Text
                    dt.Rows(i)("AvNo") = txtAvNo.Text.Trim.ToUpper
                End If
            End If
        Next

        mapGrid.EditItemIndex = -1
        mapGrid.ShowFooter = True
        ViewState("mapTable") = dt
        mapGrid.DataSource = dt
        mapGrid.DataBind()

        setDirtyFlag()

    End Sub

    Protected Sub mapGrid_CancelCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles mapGrid.CancelCommand
        mapGrid.ShowFooter = True
        mapGrid.EditItemIndex = -1
        mapGrid.DataSource = ViewState("mapTable")
        mapGrid.DataBind()
    End Sub

    Protected Sub mapGrid_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles mapGrid.ItemCommand
        If e.CommandName = "Insert" Then
            Dim add_ddlAvCategory As DropDownList
            Dim add_avNo As TextBox
            Dim avCategoryId As Integer
            Dim avCategoryName As String
            Dim avNumbers As String()

            add_ddlAvCategory = e.Item.FindControl("add_ddlAvCategory")
            add_avNo = e.Item.FindControl("add_avNo")
            avCategoryId = add_ddlAvCategory.SelectedValue
            avCategoryName = add_ddlAvCategory.SelectedItem.Text
            avNumbers = add_avNo.Text.Split(",")

            Dim dt As DataTable = ViewState("mapTable")
            For Each avNumber As String In avNumbers
                Dim row As DataRow = dt.NewRow()
                row("AvCategoryId") = avCategoryId
                row("AvCategoryName") = avCategoryName
                row("AvNo") = avNumber.Trim().ToUpper()
                dt.Rows.Add(row)
            Next

            ViewState("mapTable") = dt
            mapGrid.DataSource = dt
            mapGrid.DataBind()

            setDirtyFlag()
        End If

        If e.CommandName = "Delete" Then
            Dim hidAvNo As HiddenField
            hidAvNo = e.Item.FindControl("hidAvNo")
            Dim avNo As String = hidAvNo.Value
            Dim dt As DataTable = ViewState("mapTable")

            For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                If dt.Rows(i).RowState <> DataRowState.Deleted Then
                    If dt.Rows(i)("AvNo").ToString().Trim.ToUpper = avNo.Trim.ToUpper Then
                        dt.Rows(i).Delete()
                    End If
                End If
            Next

            ViewState("mapTable") = dt
            mapGrid.DataSource = dt
            mapGrid.DataBind()

            setDirtyFlag()

        End If

        If e.CommandName = "Find" Then
            'Dim script As StringBuilder = New StringBuilder()
            Dim add_ddlAvCategory As DropDownList = e.Item.FindControl("add_ddlAvCategory")

            ''Response.Write(e.Item.FindControl("add_AvNo").UniqueID.ToString())

            'script.Append(String.Format("var txtAvNo = document.getElementById('{0}');", e.Item.FindControl("add_AvNo").ClientID.ToString()))
            'script.Append("var retValue;")
            'script.Append(String.Format("retValue = window.parent.showModalDialog('RSLAVMapping.asp?CategoryId={0}&ProductBrandID={1}', '', 'dialogWidth:675px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No');", add_ddlAvCategory.SelectedValue.ToString(), ProductBrandId))
            'script.Append("if (retValue != undefined){")
            'script.Append("txtAvNo.value = retValue;}")

            'Me.body.Attributes.Add("onload", script.ToString())
            CategoryID = Convert.ToInt32(add_ddlAvCategory.SelectedValue)
            LoadFindPanel()
            Me.pnlMain.Visible = False
            Me.pnlFind.Visible = True

        End If
    End Sub

    Protected Sub mapGrid_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles mapGrid.ItemDataBound
        If e.Item.ItemType = ListItemType.EditItem Then
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
            Dim dtAvCategory As DataTable = dw.SelectAvFeatureCategoriesForService(ProductBrandId)

            Dim dvAvCategory As DataView = dtAvCategory.DefaultView
            'dvAvCategory.RowFilter = "ISNULL(ParentCategoryId, 0) = 0"

            Dim ddlAvCategory As DropDownList = e.Item.FindControl("ddlAvCategory")
            ddlAvCategory.DataValueField = "AvFeatureCategoryID"
            ddlAvCategory.DataTextField = "AvFeatureCategory"
            ddlAvCategory.DataSource = dvAvCategory
            ddlAvCategory.DataBind()

            Dim currentCategory As Integer = Convert.ToInt32(DataBinder.Eval(e.Item.DataItem, "AvCategoryId"))
            Dim liCategory As ListItem = ddlAvCategory.Items.FindByValue(currentCategory.ToString())
            If liCategory IsNot Nothing Then liCategory.Selected = True
        End If

        If e.Item.ItemType = ListItemType.Footer Then
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
            Dim dtAvCategory As DataTable = dw.SelectAvFeatureCategoriesForService(ProductBrandId)

            Dim dvAvCategory As DataView = dtAvCategory.DefaultView
            'dvAvCategory.RowFilter = "ISNULL(ParentCategoryId, 0) = 0"

            Dim add_ddlAvCategory As DropDownList = e.Item.FindControl("add_ddlAvCategory")
            add_ddlAvCategory.DataValueField = "AvFeatureCategoryID"
            add_ddlAvCategory.DataTextField = "AvFeatureCategory"
            add_ddlAvCategory.DataSource = dvAvCategory
            add_ddlAvCategory.DataBind()
        End If
    End Sub
#End Region

#Region " Button Events "
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Me.body.Attributes.Add("onload", "confirmSave(this);")
        cValidator.IsValid = True
        cValidator.ErrorMessage = String.Empty

        Dim secObj As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(Session("LoggedInUser"))
        If Not (secObj.IsSysAdmin Or
                secObj.UserInRole(HPQ.Excalibur.Security.ProgramRoles.GPLM) Or
                secObj.UserInRole(HPQ.Excalibur.Security.ProgramRoles.ServiceBomAnalyst)) Then

            'Dim cValidator As CustomValidator = New CustomValidator
            cValidator.IsValid = False
            cValidator.ErrorMessage = "You must be either a GPLM or Service Bom Analyst to save these changes.<br />Please click Cancel."
            cValidator.CssClass = "ErrorText"
            'divErrors.Controls.Add(cValidator)
            Exit Sub

        End If

        Dim add_AvNo As TextBox = New TextBox

        Dim dgFooter As DataGridItem =
            mapGrid.Controls(0).Controls(mapGrid.Controls(0).Controls.Count - 1)

        add_AvNo = dgFooter.FindControl("add_AvNo")

        If add_AvNo.Text <> String.Empty Then
            'Dim cValidator As CustomValidator = New CustomValidator
            cValidator.IsValid = False
            cValidator.ErrorMessage = "You must click the Add link before attempting to save your changes."
            cValidator.CssClass = "ErrorText"
            'divErrors.Controls.Add(cValidator)
            Exit Sub
        End If

        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable = ViewState("mapTable")
        Dim dv As DataView = dt.DefaultView
        dv.Sort = "AvCategoryId"


        'Count the number of distinct categories
        Dim prevCategoryId As Integer = 0
        Dim categoryCount As Integer = 0
        Dim rowCount As Integer
        For Each row As DataRowView In dv
            If prevCategoryId <> Convert.ToInt32(row("AvCategoryId")) Then
                prevCategoryId = Convert.ToInt32(row("AvCategoryId"))
                categoryCount += 1
            End If
            rowCount += 1

        Next

        'If we are in edit mode and have a category count not equal to the row count stop.
        If EditMode And categoryCount <> rowCount Then
            'Dim cValidator As CustomValidator = New CustomValidator
            cValidator.IsValid = False
            cValidator.ErrorMessage = "Each AV must be from a different category in edit mode."
            cValidator.CssClass = "ErrorText"
            'divErrors.Controls.Add(cValidator)
            Exit Sub
        End If

        'If we have more than one distinct category and more than one row per category stop.
        If categoryCount > 3 And categoryCount <> rowCount Then
            'Dim cValidator As CustomValidator = New CustomValidator
            cValidator.IsValid = False
            cValidator.ErrorMessage = "You must use unique categories for each row or the same category for every row."
            cValidator.CssClass = "ErrorText"
            'divErrors.Controls.Add(cValidator)
            Exit Sub
        End If

        Dim prevSpareKitMapId As Integer = 0

        If categoryCount = 2 And rowCount <> categoryCount Then

            Dim Category1 As Integer = 0
            Dim Category2 As Integer = 0

            For Each row As DataRowView In dv
                'Response.Write(Convert.ToInt32(row("AvCategoryId")) & "<br>")
                If Category1 = 0 Then
                    Category1 = Convert.ToInt32(row("AvCategoryId"))
                ElseIf Category1 <> Convert.ToInt32(row("AvCategoryId")) And Category2 = 0 Then
                    Category2 = Convert.ToInt32(row("AvCategoryId"))
                End If
                'Response.Write(Category1 & ":" & Category2 & "<br>")
            Next

            dv.RowFilter = String.Format("(AvCategoryId = {0})", Category1)
            Dim dt1 As DataTable = dv.ToTable()
            dv.RowFilter = String.Format("(AvCategoryId = {0})", Category2)
            Dim dt2 As DataTable = dv.ToTable

            For Each row1 As DataRow In dt1.Rows

                'Response.Write(row1(1) & " " & row1(2) & "<br>")
                For Each row2 As DataRow In dt2.Rows
                    'Response.Write(row2(1) & " " & row2(2) & "<br>")
                    If prevSpareKitMapId = SpareKitMapId Then
                        SpareKitMapId = dw.GetNewSpareKitMapId(ProductBrandId, SpareKitId)
                    End If
                    'Response.Write(SpareKitMapId & " " & row1(1) & " " & row1(2) & "<br>")
                    'Response.Write(SpareKitMapId & " " & row2(1) & " " & row2(2) & "<br>")
                    dw.InsertServiceSpareKitAvMap(SpareKitMapId, row1(0), row1(2))
                    dw.InsertServiceSpareKitAvMap(SpareKitMapId, row2(0), row2(2))
                    prevSpareKitMapId = SpareKitMapId

                Next
            Next
        ElseIf categoryCount = 3 And rowCount <> categoryCount Then

            Dim Category1 As Integer = 0
            Dim Category2 As Integer = 0
            Dim Category3 As Integer = 0

            For Each row As DataRowView In dv
                'Response.Write(Convert.ToInt32(row("AvCategoryId")) & "<br>")
                If Category1 = 0 Then
                    Category1 = Convert.ToInt32(row("AvCategoryId"))
                ElseIf Category1 <> Convert.ToInt32(row("AvCategoryId")) And Category2 = 0 Then
                    Category2 = Convert.ToInt32(row("AvCategoryId"))
                ElseIf Category1 <> Convert.ToInt32(row("AvCategoryId")) And Category2 <> Convert.ToInt32(row("AvCategoryId")) And Category3 = 0 Then
                    Category3 = Convert.ToInt32(row("AvCategoryId"))
                End If
                'Response.Write(Category1 & ":" & Category2 & ":" & Category3 & "<br>")
            Next

            dv.RowFilter = String.Format("(AvCategoryId = {0})", Category1)
            Dim dt1 As DataTable = dv.ToTable()
            dv.RowFilter = String.Format("(AvCategoryId = {0})", Category2)
            Dim dt2 As DataTable = dv.ToTable
            dv.RowFilter = String.Format("(AvCategoryId = {0})", Category3)
            Dim dt3 As DataTable = dv.ToTable

            For Each row1 As DataRow In dt1.Rows
                'Response.Write(row1(1) & " " & row1(2) & "<br>")
                For Each row2 As DataRow In dt2.Rows
                    'Response.Write(row2(1) & " " & row2(2) & "<br>")
                    For Each row3 As DataRow In dt3.Rows
                        'Response.Write(row3(1) & " " & row3(2) & "<br>")
                        If prevSpareKitMapId = SpareKitMapId Then
                            SpareKitMapId = dw.GetNewSpareKitMapId(ProductBrandId, SpareKitId)
                        End If
                        'Response.Write(SpareKitMapId & " " & row1(1) & " " & row1(2) & "<br>")
                        'Response.Write(SpareKitMapId & " " & row2(1) & " " & row2(2) & "<br>")
                        'Response.Write(SpareKitMapId & " " & row3(1) & " " & row3(2) & "<br>")
                        dw.InsertServiceSpareKitAvMap(SpareKitMapId, row1(0), row1(2))
                        dw.InsertServiceSpareKitAvMap(SpareKitMapId, row2(0), row2(2))
                        dw.InsertServiceSpareKitAvMap(SpareKitMapId, row3(0), row3(2))
                        prevSpareKitMapId = SpareKitMapId

                    Next
                Next
            Next

        Else
            '   checkDuplicates()
            '
            ' Look for rows that are deleted and added in the same category
            '
            Dim deletedRowCategoryId As Integer = 0
            Dim deleteAndAddInSameTransaction As Boolean
            For Each row As DataRow In dt.Rows
                If row.RowState = DataRowState.Deleted Then
                    deletedRowCategoryId = Convert.ToInt32(row("AvCategoryId", DataRowVersion.Original))
                End If
            Next
            If deletedRowCategoryId <> 0 Then
                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Added Then
                        deleteAndAddInSameTransaction = (Convert.ToInt32(row("AvCategoryId")) = deletedRowCategoryId)
                    End If
                Next
            End If
            If deleteAndAddInSameTransaction Then
                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then
                        Dim AvNo As String = row(2, DataRowVersion.Original).ToString()
                        dw.DeleteServiceSpareKitAvMap(SpareKitMapId, AvNo)
                    ElseIf row.RowState <> DataRowState.Deleted Then
                        dw.DeleteServiceSpareKitAvMap(SpareKitMapId, row(2))
                    End If
                Next
                SpareKitMapId = dw.GetNewSpareKitMapId(ProductBrandId, SpareKitId)
                For Each row As DataRow In dt.Rows
                    If row.RowState <> DataRowState.Deleted Then
                        dw.InsertServiceSpareKitAvMap(SpareKitMapId, row(0), row(2))
                    End If
                Next
            Else
                For Each row As DataRow In dt.Rows
                    If row.RowState = DataRowState.Deleted Then
                        Dim AvNo As String = row(2, DataRowVersion.Original).ToString()
                        dw.DeleteServiceSpareKitAvMap(SpareKitMapId, AvNo)
                    ElseIf row.RowState = DataRowState.Added Then
                        If prevSpareKitMapId = SpareKitMapId And categoryCount = 1 And rowCount > 1 Then
                            SpareKitMapId = dw.GetNewSpareKitMapId(ProductBrandId, SpareKitId)
                        End If
                        'If checkDuplicates(SpareKitMapId, row(0), row(2), ProductBrandId) = True Then
                        '    Exit Sub
                        'End If
                        dw.InsertServiceSpareKitAvMap(SpareKitMapId, row(0), row(2))
                        prevSpareKitMapId = SpareKitMapId

                    End If
                Next
            End If
        End If

        'Me.body.Attributes.Add("onload", "bPostBack=true;window.returnValue='refresh';window.close();")
        System.Web.UI.ScriptManager.RegisterClientScriptBlock(Me.btnSave, Me.GetType(), "Close_Window", "window.returnValue='refresh';window.close();", True)

    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
    End Sub

    Private Function checkDuplicates(ByVal SpareKitMapId As String, ByVal catid As String, ByVal avno As String, ByVal pbId As String) As Boolean
        'Dim sql As String
        'Dim retVal As Boolean

        'sql = " select * from servicesparekitmapav tb1 inner join servicesparekitmap tb2 on tb1.ServiceSpareKitMapId = tb2.id where tb1.avno = '" & avno & "'" & _
        '        " and tb1.AvCategoryId = '" & catid & "'" & _
        '        " and tb2.ProductBrandId = '" & pbId & "'" & _
        '        " and tb1.status = 'A' "

        'Dim objDT As DataTable = Nothing
        'Dim objDW As HPQ.Data.DataWrapper = Nothing
        'Dim objComm As Data.SqlClient.SqlCommand
        'objDW = New HPQ.Data.DataWrapper()
        'objComm = objDW.CreateCommand(sql, Data.CommandType.Text)
        'objDT = objDW.ExecuteCommandTable(objComm)

        'If objDT.Rows.Count > 0 Then
        '    cValidator.IsValid = False
        '    cValidator.ErrorMessage = "AV number " & avno & " as already mapped and cannot be added"
        '    cValidator.CssClass = "ErrorText"
        '    retVal = True
        'Else
        '    retVal = False
        'End If
        ' System.Web.UI.ScriptManager.RegisterClientScriptBlock(Me.btnSave, Me.GetType(), "Close_Window", "alert('" & sql & "');", True)
        Return False
    End Function
#End Region

#Region " Find Dialog "
    Private Sub LoadFindPanel()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable = dw.SelectAvByBrandCategory(ProductBrandId, CategoryID)

        ViewState("CategoryTable") = dt
        gvAVNumbers.DataSource = dt
        gvAVNumbers.DataBind()

        ViewState("SortDirection") = "ASC"
        ViewState("SortExpression") = "AvNo"

        ViewState("SelectAll") = "FALSE"

        ' Store any selected items (checkboxes)

    End Sub

    Protected Sub cbxAll_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
        '''''''UNCOMMENT LINES BELOW WHEN AJAX UPDATE PANEL IS ENABLED''''''''
        Dim cbxAll As CheckBox = sender
        Dim row As GridViewRow
        'Dim strAVs As String = ""

        ViewState("SelectAll") = cbxAll.Checked.ToString.ToUpper

        For Each row In gvAVNumbers.Rows
            Dim cbxAVNumber As CheckBox = CType(row.FindControl("cbxAVNumber"), CheckBox)
            'Dim lblAvNo As Label = CType(row.FindControl("lblAvNo"), Label)
            If cbxAll.Checked Then
                cbxAVNumber.Checked = True
                'If strAVs = "" Then
                'strAVs = lblAvNo.text
                'Else
                'strAVs = strAVs & "," & lblAvNo.text
                'End If
            Else
                cbxAVNumber.Checked = False
            End If
        Next
        'If strAVs = "" Then
        '    lblAVs.text = "NONE"
        'Else
        '    lblAVs.text = strAVs
        'End If
    End Sub

    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data

        Dim dt As New DataTable
        Dim txtAVFilter As TextBox = CType(Me.form1.FindControl("txtAVFilter"), TextBox)

        dt = dw.SelectAvByBrandCategory(416, 1)
        ViewState("CategoryTable") = dt
        gvAVNumbers.DataSource = dt
        gvAVNumbers.DataBind()
        txtAVFilter.Text = ""
    End Sub

    'Protected Sub cbxAVNumber_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
    '''''''UNCOMMENT LINES BELOW WHEN AJAX UPDATE PANEL IS ENABLED''''''''
    '    Dim strAVs As String = ""
    '    Dim row As GridViewRow
    '    For Each row In gvAVNumbers.Rows
    '        Dim cbxAVNumber As CheckBox = CType(row.FindControl("cbxAVNumber"), CheckBox)
    '        If cbxAVNumber.Checked Then
    '            Dim lblAvNo As Label = CType(row.FindControl("lblAvNo"), Label)
    '            If strAVs = "" Then
    '                strAVs = lblAvNo.text
    '            Else
    '                strAVs = strAVs & "," & lblAvNo.text
    '            End If
    '        End If
    '    Next
    '    If strAVs = "" Then
    '        lblAVs.text = "NONE"
    '    Else
    '        lblAVs.text = strAVs
    '    End If
    'End Sub

    Protected Sub btnAVFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAVFilter.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
        Dim dt As New DataTable
        Dim txtAVFilter As TextBox = CType(Me.form1.FindControl("txtAVFilter"), TextBox)

        If txtAVFilter.Text.Trim = "" Then
            dt = dw.SelectAvByBrandCategory(416, 1)
            ViewState("CategoryTable") = dt
            gvAVNumbers.DataSource = dt
            gvAVNumbers.DataBind()
        Else
            dt = ViewState("CategoryTable")
            dt.DefaultView.RowFilter = "(AvNo LIKE '%" & txtAVFilter.Text & "%')"
            If dt.DefaultView.ToTable.Rows.Count > 0 Then
                ViewState("CategoryTable") = dt.DefaultView.ToTable
                gvAVNumbers.DataSource = ViewState("CategoryTable")
                gvAVNumbers.DataBind()
            End If
        End If

        ' Iterate through the filtered grid rows and set the check boxes
        updateSelections()

    End Sub

    Protected Sub btnFindSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindSubmit.Click
        Dim row As GridViewRow
        Dim strAVs As String = ""
        For Each row In gvAVNumbers.Rows
            Dim cbxAVNumber As CheckBox = CType(row.FindControl("cbxAVNumber"), CheckBox)
            Dim lblAvNo As Label = CType(row.FindControl("lblAvNo"), Label)
            If cbxAVNumber.Checked Then
                If strAVs = "" Then
                    strAVs = lblAvNo.Text
                Else
                    strAVs += "," & lblAvNo.Text
                End If
            End If
        Next



        Dim footerIndex As Integer = mapGrid.Controls(0).Controls.Count - 1
        Dim txtAvNo As TextBox = mapGrid.Controls(0).Controls(footerIndex).FindControl("add_AvNo")

        txtAvNo.Text = strAVs

        pnlFind.Visible = False
        pnlMain.Visible = True
    End Sub

    Protected Sub btnFindCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFindCancel.Click
        pnlFind.Visible = False
        pnlMain.Visible = True
    End Sub

    Private Function GetSortDirection(ByVal column As String) As String

        ' By default, set the sort direction to ascending.
        Dim sortDirection As String = "ASC"

        ' Retrieve the last column that was sorted.
        Dim sortExpression As String = TryCast(ViewState("SortExpression"), String)

        If sortExpression IsNot Nothing Then
            ' Check if the same column is being sorted.
            ' Otherwise, the default value can be returned.
            If sortExpression = column Then

                Dim lastDirection As String = TryCast(ViewState("SortDirection"), String)

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


    Function getSelections() As String
        Dim oRow As GridViewRow
        Dim cbxAVNumber As CheckBox
        Dim lblAvNo As Label
        Dim strAVs As String = ""

        For Each oRow In gvAVNumbers.Rows
            cbxAVNumber = CType(oRow.FindControl("cbxAVNumber"), CheckBox)
            lblAvNo = CType(oRow.FindControl("lblAvNo"), Label)

            If cbxAVNumber.Checked Then
                If strAVs = "" Then
                    strAVs = lblAvNo.Text.ToUpper

                Else
                    strAVs = strAVs & "," & lblAvNo.Text.ToUpper

                End If
            End If
        Next

        Return strAVs

    End Function


    Function FlagSelections(ByVal strSelections As String) As Boolean
        Dim blnResult As Boolean = False
        Dim intCount As Integer = 0
        Dim intTotalRows As Integer = gvAVNumbers.Rows.Count
        Dim oRow As GridViewRow
        Dim cbxAVNumber As CheckBox
        Dim lblAvNo As Label

        For Each oRow In gvAVNumbers.Rows
            cbxAVNumber = CType(oRow.FindControl("cbxAVNumber"), CheckBox)
            lblAvNo = CType(oRow.FindControl("lblAvNo"), Label)

            If strSelections.ToUpper.IndexOf("," & lblAvNo.Text.ToUpper & ",") > -1 Then
                cbxAVNumber.Checked = True
                intCount = intCount + 1
            End If
        Next

        If (intCount = intTotalRows) Then blnResult = True

        Return blnResult

    End Function

    Protected Sub gvAVNumbers_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs)
        Dim dt As New DataTable
        dt = ViewState("CategoryTable")
        dt.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)

        ViewState("CategoryTable") = dt.DefaultView.ToTable
        gvAVNumbers.DataSource = ViewState("CategoryTable")
        gvAVNumbers.DataBind()

        ' Iterate through the sorted grid rows and set the check boxes
        updateSelections()

    End Sub

    Sub updateSelections()

        If Me.Selections.Trim.Length > 0 Then
            blnSelectAll = FlagSelections("," & Me.Selections & ",")
        Else
            blnSelectAll = False
        End If

        ViewState("SelectAll") = blnSelectAll.ToString.ToUpper

        Try
            Dim cbxAVAll As CheckBox = CType(gvAVNumbers.HeaderRow.FindControl("cbxAll"), CheckBox)
            cbxAVAll.Checked = blnSelectAll
        Catch ex As Exception

        End Try

    End Sub
#End Region
End Class
