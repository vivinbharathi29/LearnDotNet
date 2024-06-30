Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml


Partial Class Service_SKUBOMAS
    Inherits System.Web.UI.Page

    Private strParameters As String = ""
    Protected Property Parameters As String
        Get
            Return strParameters
        End Get
        Set(ByVal value As String)
            strParameters = value
        End Set
    End Property

    Private intUserId As Integer = GetUserID()

    Protected Property UserId As Integer
        Get
            Return intUserId
        End Get
        Set(ByVal value As Integer)
            intUserId = value
        End Set
    End Property

    Private intReturnCode As Integer
    Protected Property ReturnCode As Integer
        Get
            Return intReturnCode
        End Get
        Set(ByVal value As Integer)
            intReturnCode = value
        End Set
    End Property

    Private strReturnDesc As String
    Protected Property ReturnDesc As String
        Get
            Return strReturnDesc
        End Get
        Set(ByVal value As String)
            strReturnDesc = value
        End Set
    End Property

    Private strResults As String = ""
    Protected Property Results As String
        Get
            Return strResults
        End Get
        Set(ByVal value As String)
            strResults = value
        End Set
    End Property

    Private strCurrProfileName As String = ""
    Protected Property CurrProfileName As String
        Get
            Return strCurrProfileName
        End Get
        Set(ByVal value As String)
            strCurrProfileName = value
        End Set
    End Property

    Private intRowsPerPage As Integer = 20
    Protected Property RowsPerPage As Integer
        Get
            Return intRowsPerPage
        End Get
        Set(ByVal value As Integer)
            intRowsPerPage = value
        End Set
    End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim objSec As HPQ.Excalibur.Security = Nothing

        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoServerCaching()

        Dim intUID As Integer = 0

        Try

            Me.profileFlag.Value = "false"

            Me.UserId = GetUserID()

            If (Me.UserId <> 0) Then

                If Not Me.Page.IsPostBack Then ' Initial load

                    ViewState("UserID") = Me.UserId

                    ' Populate Selected columns from the default column list
                    Me.selectedCols.Value = PopulateDefaultSelectedColumns(Me.selectedColumns, Nothing) ' Me.orderByCols)
                    Me.orderCols.Value = Me.selectedColumns.Items(0).Value.ToString() & "-0"
                    Me.profileName.Value = Me.CurrProfileName
                    PopulateProfileList(False, Nothing)

                    Me.Parameters = ""

                    SetProductFilter("0")

                End If

            Else
                DisplayMessage("<b>User credentials not verified for this action.</b>", "lblStatus", Drawing.Color.Red)
            End If

        Catch ex As Exception
            DisplayMessage("<b>" & ex.Message & "</b>", "lblStatus", Drawing.Color.Red)
        Finally
            objSec = Nothing
        End Try

    End Sub


    Private Sub PopulateProfileList(ByVal blnSuppressSuccessMsg As Boolean, ByVal strSelectedItemText As String)

        ' Retrieve the User Profiles
        Dim objDT As DataTable = Nothing
        ' Dim objDW As HPQ.Data.DataWrapper = Nothing
        'Dim objComm As Data.SqlClient.SqlCommand
        Dim objRow As DataRow

        Try

            ' Validate user requesting this action
            If ((ViewState("UserID") = GetUserID())) Then

                objDT = GetAllProfiles(Me.UserId, Me.chkBxIncludeRemSProfs.Checked)

                ddlRptProfile.Items.Clear()

                ddlRptProfile.Items.Add(New ListItem("-- Select Profile --", "0"))

                If (Not objDT Is Nothing) Then

                    For Each objRow In objDT.Rows
                        ddlRptProfile.Items.Add(New ListItem(objRow("ProfileName"), objRow("ID")))
                    Next

                    If Not blnSuppressSuccessMsg Then
                        DisplayMessage("<b>Successfully retrieved Profile(s).</b>", "lblProfileStatus", Drawing.Color.Black)
                    End If

                    If (Not strSelectedItemText Is Nothing) Then
                        Me.ddlRptProfile.Items.FindByText(strSelectedItemText).Selected = True
                    Else
                        Me.ddlRptProfile.SelectedIndex = 0
                    End If

                End If

            Else
                DisplayMessage("<b>User credentials not verified for this action.</b>", "lblProfileStatus", Drawing.Color.Red)
            End If

        Catch ex As Exception
            DisplayMessage("<b>PopulateProfileList</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDT = Nothing
            'objComm = Nothing
            'objDW = Nothing
        End Try

    End Sub

    Private Function GetRemovedSharedProfilesList(ByVal strUID As String) As String
        Dim strList As String = ""
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand


        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("SELECT PROFILE_ID FROM dbo.BTOSSREMOVEDPROFILES WITH(NOLOCK) WHERE USER_ID=" & strUID, Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)

        If (objDT.Rows.Count > 0) Then

            For Each objRow As DataRow In objDT.Rows
                If (strList.Length = 0) Then
                    strList = objRow(0).ToString
                Else
                    strList += "," & objRow(0).ToString
                End If
            Next
        End If

        objDT = Nothing
        objComm = Nothing
        objDW = Nothing


        GetRemovedSharedProfilesList = strList

    End Function

    Private Function IsRemovedSharedProfile(ByVal strRemovedSharedProfileList As String, ByVal intProfileID As Integer) As Boolean
        Dim blnResult As Boolean = False
        Dim strTmp As String = "," & strRemovedSharedProfileList & ","

        If (strTmp.IndexOf("," & intProfileID.ToString & ",") <> -1) Then
            blnResult = True
        End If

        IsRemovedSharedProfile = blnResult

    End Function

    Private Function GetAllProfiles(ByVal intUserID As Integer, ByVal blnIncludeRemovedSharedProfiles As Boolean) As DataTable
        ' ADD LOGIC TO ACCOMMODATE DIFFERENT PROFILES (i.e. - Group Shared/Shared/Etc) USE EXISTING SOURCE AS REFERENCED IN WJs TASK DESC
        '       Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        '       Dim dtReportProfiles As DataTable = dw.ListReportPYPE_ID) --- 8
        '       Dim dtReportProfilesShared As DataTable = dw.ListReportProfilesShared(EmployeeID, REPORT_PROFILE_TYPE_ID) --- 8
        '       Dim dtReportProfilesGroupShared As DataTable = dw.ListReportProfilesGroupShared(EmployeeID, REPORT_PROFILE_TYPE_ID) --- 8rofiles(EmployeeID, REPORT_PROFILE_T
        Dim objDW As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim objRPDT As DataTable = objDW.ListReportProfiles(intUserID, 8)
        Dim objSRPDT As DataTable = objDW.ListReportProfilesShared(intUserID, 8)
        Dim objGSRPDT As DataTable = objDW.ListReportProfilesGroupShared(intUserID, 8)
        Dim objALLDT As New DataTable

        Dim objRow As DataRow
        Dim strRow As String
        Dim aryCols As Object()
        Dim strRemovedSharedProfileList As String = GetRemovedSharedProfilesList(intUserID.ToString)


        Try ' MAY NEED TO CHECK FOR DUPLICATE NAMES/IDs---> PERMISSIONS MAY VARY (i.e. DELETE/EDIT PERMS) -- CONSIDER S AS PREFIX FOR ALL SHARED PROFS, OR USE NEGATIVE NUMBER IDENTIFIER
            ' S and G Prefixes ARE NOT REALLY NECESSARY, CONSIDER REWORKING (PERHAPS, NEGATIVE [-] SIGN)
            If (strRemovedSharedProfileList.Length > 0) Then
                Me.chkBxIncludeRemSProfs.Visible = True
                Me.chkBxIncludeRemSProfs.Enabled = True
            Else
                Me.chkBxIncludeRemSProfs.Visible = False
                Me.chkBxIncludeRemSProfs.Enabled = False
            End If

            If (blnIncludeRemovedSharedProfiles) Then
                strRemovedSharedProfileList = ""
            End If


            objALLDT.Columns.Add("ID")
            objALLDT.Columns.Add("ProfileName")

            For Each objRow In objRPDT.Rows
                If (objALLDT.Select("ID='" & objRow.Item("ID").ToString() & "'").Length = 0) Then
                    strRow = objRow.Item("ID").ToString() & "|" & objRow.Item("ProfileName")
                    aryCols = strRow.Split("|")
                    objALLDT.Rows.Add(aryCols)
                End If
            Next

            For Each objRow In objSRPDT.Rows
                If (objALLDT.Select("ID='" & objRow.Item("ID").ToString() & "'").Length = 0) And (objALLDT.Select("ID='S" & objRow.Item("ID").ToString() & "'").Length = 0) And (Not IsRemovedSharedProfile(strRemovedSharedProfileList, objRow.Item("ID"))) Then
                    strRow = "S" & objRow.Item("ID").ToString() & "|" & objRow.Item("ProfileName") & " (Shared)"
                    aryCols = strRow.Split("|")
                    objALLDT.Rows.Add(aryCols)
                End If
            Next

            For Each objRow In objGSRPDT.Rows
                If (objALLDT.Select("ID='" & objRow.Item("ID").ToString() & "'").Length = 0) And (objALLDT.Select("ID='G" & objRow.Item("ID").ToString() & "'").Length = 0) And (objALLDT.Select("ID='S" & objRow.Item("ID").ToString() & "'").Length = 0) And (Not IsRemovedSharedProfile(strRemovedSharedProfileList, objRow.Item("ID"))) Then
                    strRow = "G" & objRow.Item("ID").ToString() & "|" & objRow.Item("ProfileName") & " (Group Shared)"
                    aryCols = strRow.Split("|")
                    objALLDT.Rows.Add(aryCols)
                End If
            Next


            objALLDT.DefaultView.Sort = "ProfileName ASC"
            objALLDT = objALLDT.DefaultView.ToTable

        Catch ex As Exception

            DisplayMessage("<b>GetAllProfiles</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDW = Nothing
            objRPDT = Nothing
            objSRPDT = Nothing
            objGSRPDT = Nothing
        End Try

        GetAllProfiles = objALLDT

    End Function


    Private Function GetUserID() As Integer
        Dim intUserID As Integer = 0
        Dim objSec As HPQ.Excalibur.Security
        Dim strLOGON_USER As String = ""
        Dim strHTTPCTXT_USER As String = ""

        Try
            strLOGON_USER = Session("LoggedInUser")
            strHTTPCTXT_USER = Session("LoggedInUser")

            If (strHTTPCTXT_USER.Trim.Length > 0) Then
                objSec = New HPQ.Excalibur.Security(strHTTPCTXT_USER.Trim)
                intUserID = objSec.CurrentUserID
            ElseIf (strLOGON_USER.Trim.Length > 0) Then
                objSec = New HPQ.Excalibur.Security(strLOGON_USER.Trim)
                intUserID = objSec.CurrentUserID
            Else
                Me.ReturnCode = -1
                Me.ReturnDesc = "<b>ERROR DETERMINING USER VERIFICATION PARAMETERS</b>"
                intUserID = 0
            End If

        Catch ex As Exception
            Me.ReturnCode = -1
            Me.ReturnDesc = "<b>USER VERIFICATION ERROR</b> - " & ex.Message
            intUserID = 0
        Finally
            objSec = Nothing
        End Try

        GetUserID = intUserID

    End Function


    Private Function PopulateDefaultSelectedColumns(ByRef objList As ListBox, ByRef objDropList As DropDownList) As String
        Dim strDefSelCols As String = ""
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = New HPQ.Data.DataWrapper()
        Dim objComm As Data.SqlClient.SqlCommand

        objComm = objDW.CreateCommand("SELECT ColumnID, ColumnDesc FROM BTOSSASColumns WITH(NOLOCK) WHERE Active=1 AND [Default]=1 ORDER BY OrderIndex ASC", Data.CommandType.Text)

        objDT = objDW.ExecuteCommandTable(objComm)
        Dim objRow As DataRow

        For Each objRow In objDT.Rows
            objList.Items.Add(New ListItem(objRow(1).ToString(), objRow(0).ToString()))

            If strDefSelCols.Length = 0 Then
                strDefSelCols = objRow(0).ToString
            Else
                strDefSelCols += "," + objRow(0).ToString()
            End If
        Next

        objComm = Nothing
        objDW = Nothing

        PopulateDefaultSelectedColumns = strDefSelCols

    End Function


    Private Sub DisplayMessage(ByVal strMsg As String, ByVal strControlName As String, ByVal objColor As System.Drawing.Color)
        Dim objStatusLabel As Label = DirectCast(Me.FindControl(strControlName), Label)

        Try
            objStatusLabel.ForeColor = objColor
            objStatusLabel.Text = strMsg
        Catch ex As Exception
            ' Do Nothing
        End Try

    End Sub


    Function GetColumnName(ByVal intColumnID As Integer) As String
        Dim strName As String = ""
        Dim objListItem As ListItem

        For Each objListItem In Me.allColumns.Items
            If objListItem.Value = intColumnID.ToString() Then
                strName = objListItem.Text
                Exit For
            End If

        Next

        GetColumnName = strName

    End Function


    Function GetColumnID(ByVal strColumnName As String) As Integer
        Dim intID As Integer = 0
        Dim objListItem As ListItem

        For Each objListItem In Me.allColumns.Items
            If objListItem.Text.ToUpper = strColumnName.ToUpper Then
                If Integer.TryParse(objListItem.Value, intID) Then
                    Exit For
                End If
            End If
        Next

        GetColumnID = intID

    End Function


    Private Function GetProfileParameters(ByVal strProfileID As String) As String
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim intProfileUserID As Integer = 0
        Dim strParameters As String = ""

        Try
            objDW = New HPQ.Data.DataWrapper()
            objComm = objDW.CreateCommand("spGetReportProfile", Data.CommandType.StoredProcedure)
            objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
            objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, Me.UserID.ToString, 8)

            objDT = objDW.ExecuteCommandTable(objComm)

            If (objDT.Rows.Count > 0) Then
                Dim objRow As DataRow = objDT.Rows(0)
                strParameters = objRow("SelectedFilters")

                intProfileUserID = objRow("EmployeeID")

                If (intProfileUserID <> Me.UserId) Then
                    Me.ReturnCode = -1
                    Me.ReturnDesc = "Invalid User Profile specified."
                    strParameters = ""
                End If

            End If

        Catch ex As Exception
            DisplayMessage("<b>GetProfileParameters</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDT = Nothing
            objComm = Nothing
            objDW = Nothing
        End Try

        GetProfileParameters = strParameters

    End Function

    Private Function getOrderByColumnData(ByVal strData As String, ByVal intMode As Integer) As String


        Dim strRetValue As String = ""
        Dim aryData As String() = strData.Split(",")
        Dim arySubData As String()
        Dim intIdx As Integer

        For intIdx = 0 To UBound(aryData)
            arySubData = aryData(intIdx).Split("-")

            If ((Not arySubData(intMode) Is Nothing) And (arySubData(intMode) <> vbNull)) Then
                If (strRetValue.Length = 0) Then
                    strRetValue = arySubData(intMode)
                Else
                    strRetValue = strRetValue & "," & arySubData(intMode)
                End If
            End If
        Next

        getOrderByColumnData = strRetValue

    End Function

    Private Function GetParameters() As String
        Dim strValues As String = ""
        ' NEED TO ADD LOGIC TO REPLACE PIPE DELIMITER ON OPEN INPUTS
        Dim strOrderData As String = Me.orderCols.Value

        ' @ColumnIDs
        strValues = Me.selectedCols.Value

        ' @ColumnOrderByIDs
        strValues += "|" & strOrderData

        ' @ColumnOrderAscDesc
        strValues += "|"

        '@ServiceGeoNA
        strValues += "|" & Me.chkbxGeoNA.Checked.ToString

        '@ServiceGeoLA
        strValues += "|" & Me.chkbxGeoLA.Checked.ToString

        '@ServiceGeoAPJ
        strValues += "|" & Me.chkbxGeoAPJ.Checked.ToString

        '@ServiceGeoEMEA
        strValues += "|" & Me.chkbxGeoEMEA.Checked.ToString

        '@SKGeoNA
        strValues += "|" & Me.chkbxSKGeoNA.Checked.ToString

        '@SKGeoLA
        strValues += "|" & Me.chkbxSKGeoLA.Checked.ToString

        '@SKGeoAPJ
        strValues += "|" & Me.chkbxSKGeoAPJ.Checked.ToString

        '@SKGeoEMEA
        strValues += "|" & Me.chkbxSKGeoEMEA.Checked.ToString

        '@ProductBrandIDs
        strValues += "|" & GetListSelections(Me.lstbxBrands)

        '@ServiceCategoryIDs
        strValues += "|" & GetListSelections(Me.lstbxCats)

        '@OSSPIDs
        strValues += "|" & GetListSelections(Me.lstbxOSSPs)

        '@KMATs
        strValues += "|" & Me.txtbxKMATS.Text.Replace("|", "")

        '@SKUs
        strValues += "|" & Me.txtbxSKUS.Text.Replace("|", "")

        '@AVs
        strValues += "|" & Me.txtbxAVS.Text.Replace("|", "")

        '@RequireAllAVs
        strValues += "|" & Me.chkbxAllAVs.Checked.ToString

        '@SKs
        strValues += "|" & Me.txtbxSKs.Text.Replace("|", "")

        '@SFPNS
        strValues += "|" & Me.txtbxSFPNS.Text.Replace("|", "")

        '@LastAction
        Dim strLastAction As String = ""
        If (Me.chkbxAdded.Checked) Then strLastAction = "I"
        If (Me.chkbxUpdated.Checked) Then
            If Len(strLastAction) = 0 Then
                strLastAction = "U"
            Else
                strLastAction += ",U"
            End If
        End If

        strValues += "|" & strLastAction

        '@ActionDateFrom
        strValues += "|" & Me.txtbxFromDate.Text.Replace("|", "")

        '@ActionDateTo
        strValues += "|" & Me.txtbxToDate.Text.Replace("|", "")

        'Rows/Page
        strValues += "|" & Me.ddlRowsPerPage.SelectedItem.Value

        '@SKUAVs
        strValues += "|" & Me.txtbxSKUAVS.Text.Replace("|", "")

        '@RequireAllSKUAVs
        strValues += "|" & Me.chkbxAllSKUAVs.Checked.ToString

        'ReportType
        strValues += "|" & Me.ddlReportType.SelectedItem.Value

        '@ActionDateColumnID
        strValues += "|" & Me.ddlActionDateColumn.SelectedItem.Value

        '@SAs
        strValues += "|" & Me.txtbxSAs.Text.Replace("|", "")

        '@COMPs
        strValues += "|" & Me.txtbxComps.Text.Replace("|", "")

        '@PROD_DIV
        strValues += "|" & Me.rdoProdDiv.SelectedValue.ToString()

        '@SPBLogFilter
        strValues += "|" & Me.chkbxSPBLogFilter.Checked.ToString()

        '@SKUAVLogFilter
        strValues += "|" & Me.chkbxSKUAVLogFilter.Checked.ToString()


        GetParameters = strValues

    End Function

    Sub SetListSelections(ByVal objList As ListBox, ByVal strSelections As String)
        Dim arySelections() As String
        Dim i As Integer
        Dim iLBIdx As Integer
        Dim iUBIdx As Integer

        Try

            arySelections = strSelections.Split(",")
            objList.ClearSelection()

            If (strSelections.Length > 0) Then
                If (arySelections.Length > 0) Then
                    iLBIdx = arySelections.GetLowerBound(0)
                    iUBIdx = arySelections.GetUpperBound(0)
                    For i = iLBIdx To iUBIdx
                        objList.Items.FindByValue(arySelections(i)).Selected = True
                    Next
                End If
            End If
            'DisplayMessage("<b>SetListSelections</b> - " & strSelections, "lblProfileStatus", Drawing.Color.Red)
        Catch ex As Exception
            DisplayMessage("<b>SetListSelections</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        End Try

    End Sub

    Sub SetDDListSelection(ByVal objList As DropDownList, ByVal strSelection As String)
        Dim objListItem As ListItem = Nothing

        Try
            objListItem = objList.Items.FindByValue(strSelection)

            If (Not objListItem Is Nothing) Then
                objList.ClearSelection()
                objListItem.Selected = True
            End If
        Catch ex As Exception
            DisplayMessage("<b>SetDDListSelection</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        End Try

    End Sub

    Private Sub PopulateResultParameters(ByVal strColumnIDs As String)
        Dim aryColumnIDs() As String
        Dim objListItem As ListItem = Nothing
        Dim i As Integer

        Try
            aryColumnIDs = strColumnIDs.Split(",")
            Me.selectedColumns.Items.Clear()

            For i = 0 To aryColumnIDs.Length - 1
                Me.selectedColumns.Items.Add(New ListItem(GetColumnName(Integer.Parse(aryColumnIDs(i))), aryColumnIDs(i)))
            Next


        Catch ex As Exception
            DisplayMessage("<b>" & ex.Message & "</b>", "lblProfileStatus", Drawing.Color.Red)
        End Try
    End Sub

    Private Sub ApplyProfile(ByVal strParameters As String)
        Dim aryParameters() As String = strParameters.Split("|")
        Dim blnChecked As Boolean = False

        ' ADD ERROR TRAPPING FOR CORRUPT ARRAY
        Try

            ' @ColumnIDs
            Me.selectedCols.Value = aryParameters(0)
            'PopulateResultParameters(aryParameters(0))

            ' @ColumnOrderByIDs
            Me.orderCols.Value = aryParameters(1)

            ' @ServiceGeoNA
            If (Boolean.TryParse(aryParameters(3), blnChecked)) Then
                Me.chkbxGeoNA.Checked = blnChecked
            End If

            ' @ServiceGeoLA
            If (Boolean.TryParse(aryParameters(4), blnChecked)) Then
                Me.chkbxGeoLA.Checked = blnChecked
            End If

            ' @ServiceGeoAPJ
            If (Boolean.TryParse(aryParameters(5), blnChecked)) Then
                Me.chkbxGeoAPJ.Checked = blnChecked
            End If

            ' @ServiceGeoEMEA
            If (Boolean.TryParse(aryParameters(6), blnChecked)) Then
                Me.chkbxGeoEMEA.Checked = blnChecked
            End If

            ' @SKGeoNA
            If (Boolean.TryParse(aryParameters(7), blnChecked)) Then
                Me.chkbxSKGeoNA.Checked = blnChecked
            End If

            ' @SKGeoLA
            If (Boolean.TryParse(aryParameters(8), blnChecked)) Then
                Me.chkbxSKGeoLA.Checked = blnChecked
            End If

            ' @SKGeoAPJ
            If (Boolean.TryParse(aryParameters(9), blnChecked)) Then
                Me.chkbxSKGeoAPJ.Checked = blnChecked
            End If

            ' @SKGeoEMEA
            Boolean.TryParse(aryParameters(10), Me.chkbxSKGeoEMEA.Checked)
            If (Boolean.TryParse(aryParameters(10), blnChecked)) Then
                Me.chkbxSKGeoEMEA.Checked = blnChecked
            End If

            ' @ProductBrandIDs (MOVED TO BOTTOM TO ACCOMMODATE Filter By Division-->DevCenter)
            'SetListSelections(Me.lstbxBrands, aryParameters(11))

            ' @ServiceCategoryIDs
            SetListSelections(Me.lstbxCats, aryParameters(12))

            ' @OSSPIDs
            SetListSelections(Me.lstbxOSSPs, aryParameters(13))

            ' @KMATs
            If (aryParameters(14).Length = 0) Then
                aryParameters(14) = ""
            End If
            Me.txtbxKMATS.Text = aryParameters(14)

            ' @SKUs
            Me.txtbxSKUS.Text = aryParameters(15)

            ' @AVs
            Me.txtbxAVS.Text = aryParameters(16)

            ' @RequireAllAVs
            If (Boolean.TryParse(aryParameters(17), blnChecked)) Then
                Me.chkbxAllAVs.Checked = blnChecked
            End If

            ' @SKs
            Me.txtbxSKs.Text = aryParameters(18)

            ' @SFPNS
            Me.txtbxSFPNS.Text = aryParameters(19)

            ' @LastAction
            If (aryParameters(20).IndexOf("I") = -1) Then
                Me.chkbxAdded.Checked = False
            Else
                Me.chkbxAdded.Checked = True
            End If

            If (aryParameters(20).IndexOf("U") = -1) Then
                Me.chkbxUpdated.Checked = False
            Else
                Me.chkbxUpdated.Checked = True
            End If

            ' @ActionDateFrom
            Me.txtbxFromDate.Text = aryParameters(21)

            ' @ActionDateTo
            Me.txtbxToDate.Text = aryParameters(22)

            ' Rows Per Page
            If (Not IsNumeric(aryParameters(23))) Then
                aryParameters(23) = "20"
            End If

            Me.RowsPerPage = Integer.Parse(aryParameters(23))
            SetDDListSelection(Me.ddlRowsPerPage, aryParameters(23))

            ' @SKUAVs
            Me.txtbxSKUAVS.Text = aryParameters(24)

            ' @RequireAllSKUAVs
            If (Boolean.TryParse(aryParameters(25), blnChecked)) Then
                Me.chkbxAllSKUAVs.Checked = blnChecked
            End If

            ' Report Type
            SetDDListSelection(Me.ddlReportType, aryParameters(26))

            ' @ActionDateColumnID
            If (aryParameters(27).Length = 0) Then
                aryParameters(27) = "0"
            End If

            SetDDListSelection(Me.ddlActionDateColumn, aryParameters(27))

            ' @SAs
            Me.txtbxSAs.Text = aryParameters(28)

            ' @COMPs
            Me.txtbxComps.Text = aryParameters(29)

            ' @PROD_DIV
            rdoProdDiv.SelectedIndex = 0
            If (UBound(aryParameters)) > 29 Then
                rdoProdDiv.SelectedIndex = Integer.Parse(aryParameters(30))
            End If

            SetProductFilter(rdoProdDiv.SelectedIndex.ToString())

            ' @ProductBrandIDs (MOVED TO HERE TO ACCOMMODATE Filter By DevCenter)
            SetListSelections(Me.lstbxBrands, aryParameters(11))

            ' @SPBLogFilter - Apply Filter to SPB Change Log 
            If (UBound(aryParameters)) > 30 Then
                If (Boolean.TryParse(aryParameters(31), blnChecked)) Then
                    Me.chkbxSPBLogFilter.Checked = blnChecked
                Else
                    Me.chkbxSPBLogFilter.Checked = False
                End If
            Else
                Me.chkbxSPBLogFilter.Checked = False
            End If

            ' @SKUAVLogFilter - Apply Filter to SKU Change Log 
            If (UBound(aryParameters)) > 31 Then
                If (Boolean.TryParse(aryParameters(32), blnChecked)) Then
                    Me.chkbxSKUAVLogFilter.Checked = blnChecked
                Else
                    Me.chkbxSKUAVLogFilter.Checked = False
                End If
            Else
                Me.chkbxSKUAVLogFilter.Checked = False
            End If


        Catch ex As Exception
            DisplayMessage("<b>ApplyProfile</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        End Try

    End Sub


    Private Function GetOwnerID(ByVal strProfileID As String, ByVal strUID As String) As String
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim strOwnerID As String = strUID

        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("SELECT DISTINCT P.EmployeeID AS OwnerID FROM dbo.ReportProfilesGroupShared AS PGS WITH (NOLOCK) RIGHT OUTER JOIN dbo.ReportProfilesShared AS PS WITH (NOLOCK) ON PGS.GroupID = PS.GroupID RIGHT OUTER JOIN dbo.ReportProfiles AS P WITH (NOLOCK) ON PS.ReportProfileID = P.ID WHERE (P.ID = " & strProfileID & ") AND (PS.EmployeeID = " & strUID & ") OR (P.ID = " & strProfileID & ") AND (PGS.EmployeeID = " & strUID & ")", Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)

        If (objDT.Rows.Count > 0) Then
            Dim objIDRow As DataRow = objDT.Rows(0)
            strOwnerID = objIDRow("OwnerID").ToString()
        End If

        objDT = Nothing
        objComm = Nothing
        objDW = Nothing

        GetOwnerID = strOwnerID

    End Function


    Private Function GetReportProfileParameters(ByVal strProfileCode As String, ByVal strProfileID As String, ByVal strUID As String) As String
        ' TO SECURE OR NOT TO SECURE...THAT IS THE QUESTION 
        Dim strParameters As String = ""
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim strOwnerID As String = strUID

        objDW = New HPQ.Data.DataWrapper()

        Select Case strProfileCode
            Case "S", "G" ' Retrieve the Owner Employee ID
                ' USER IS NOT THE OWNER OF THIS PROFILE (SHARED)
                objComm = objDW.CreateCommand("SELECT DISTINCT P.EmployeeID AS OwnerID FROM ReportProfiles P WITH(NOLOCK) LEFT JOIN ReportProfilesShared PS WITH(NOLOCK) ON P.ID=ReportProfileID LEFT JOIN ReportProfilesGroupShared PGS WITH(NOLOCK) ON P.ID=PS.ReportProfileID WHERE P.ID=" & strProfileID & " AND (PS.EmployeeID=" & strUID & " OR PGS.EmployeeID=" & strUID & ")", Data.CommandType.Text)
                objDT = objDW.ExecuteCommandTable(objComm)

                If (objDT.Rows.Count > 0) Then
                    Dim objIDRow As DataRow = objDT.Rows(0)
                    strOwnerID = objIDRow("OwnerID").ToString()

                    If (strOwnerID <> vbNull) And (strOwnerID.Length > 0) Then
                        objComm = objDW.CreateCommand("spGetReportProfile", Data.CommandType.StoredProcedure)
                        objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
                        objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, strOwnerID, 8)
                    End If

                End If
            Case Else
                ' USER IS THE OWNER OF THIS PROFILE
                objComm = objDW.CreateCommand("spGetReportProfile", Data.CommandType.StoredProcedure)
                objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
                objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, strUID, 8)
        End Select

        objDT = objDW.ExecuteCommandTable(objComm)

        If (objDT.Rows.Count > 0) Then
            Dim objRow As DataRow = objDT.Rows(0)
            strParameters = objRow("SelectedFilters")
        End If

        objDT = Nothing
        objDW = Nothing

        GetReportProfileParameters = strParameters

    End Function


    Private Function GetProfileCode() As String
        '=============================================================================================================================
        ' ADD LOGIC TO DETERMINE IF THE SELECTED PROFILE IS (1) THE USER'S PROFILE (2) A SHARED PROFILE (3) A GROUP SHARED PROFILE
        '=============================================================================================================================
        
        Dim strProfileCode As String = Me.ddlRptProfile.SelectedValue.ToString.Substring(0, 1).ToUpper()
        If IsNumeric(strProfileCode) Then
            strProfileCode = ""
        End If

        GetProfileCode = strProfileCode
        '=============================================================================================================================
    End Function


    Private Function CanEditSharedProfile(ByVal strProfileID As String, ByVal strUID As String) As Boolean
        Dim blnResult As Boolean = False
        Dim intFlags As Integer = 0
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand

        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("SELECT SUM(CONVERT(TINYINT, ISNULL(PS.CanEdit, 0))) AS AllowEdits, P.EmployeeID AS OwnerID FROM dbo.ReportProfilesGroupShared PGS WITH(NOLOCK) RIGHT OUTER JOIN dbo.ReportProfilesShared PS WITH(NOLOCK) ON PGS.GroupID = PS.GroupID RIGHT OUTER JOIN dbo.ReportProfiles P WITH(NOLOCK) ON PS.ReportProfileID = P.ID WHERE ((P.ID = " & strProfileID & " AND PS.EmployeeID = " & strUID & ") OR (P.ID = " & strProfileID & " AND PGS.EmployeeID = " & strUID & ")) GROUP BY P.EmployeeID", Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)

        If (objDT.Rows.Count > 0) Then
            Dim objPermRow As DataRow = objDT.Rows(0)

            If (Integer.TryParse(objPermRow(0), intFlags)) Then
                If (intFlags > 0) Then
                    blnResult = True
                End If
            End If

        End If

        objDT = Nothing
        objDW = Nothing

        CanEditSharedProfile = blnResult

    End Function

    Private Function CanDeleteSharedProfile(ByVal strProfileID As String, ByVal strUID As String) As Boolean
        Dim blnResult As Boolean = False
        Dim intFlags As Integer = 0
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand

        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("SELECT SUM(CONVERT(TINYINT, ISNULL(PS.CanDelete, 0))) AS AllowDeletes, P.EmployeeID AS OwnerID FROM dbo.ReportProfilesGroupShared PGS WITH(NOLOCK) RIGHT OUTER JOIN dbo.ReportProfilesShared PS WITH(NOLOCK) ON PGS.GroupID = PS.GroupID RIGHT OUTER JOIN dbo.ReportProfiles P WITH(NOLOCK) ON PS.ReportProfileID = P.ID WHERE ((P.ID = " & strProfileID & " AND PS.EmployeeID = " & strUID & ") OR (P.ID = " & strProfileID & " AND PGS.EmployeeID = " & strUID & ")) GROUP BY P.EmployeeID", Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)

        If (objDT.Rows.Count > 0) Then
            Dim objPermRow As DataRow = objDT.Rows(0)

            If (Integer.TryParse(objPermRow(0), intFlags)) Then
                If (intFlags > 0) Then
                    blnResult = True
                End If
            End If

        End If

        objDT = Nothing
        objDW = Nothing

        CanDeleteSharedProfile = blnResult

    End Function

    Private Sub ProcessProfileAction(ByVal strActionType As String)
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim intUID As Integer
        Dim strParameters As String = ""
        Dim intParamLen As Integer = 0
        Dim strProfileCode As String = GetProfileCode()
        Dim strProfileID As String = Me.ddlRptProfile.SelectedValue.ToString.Replace("S", "").Replace("G", "")
        Dim blnHasPerm As Boolean
        Dim strOwnerID As String = ""
        Dim strProfileName As String = Me.ddlRptProfile.SelectedItem.Text.Replace("|", "")
        Dim strSelProfileName As String = ""

        If (strProfileCode = "S") Then
            strProfileName = strProfileName.Replace("(Shared)", "")
        End If

        If (strProfileCode = "G") Then
            strProfileName = strProfileName.Replace("(Group Shared)", "")
        End If

        Try

            ' Validate user requesting this action
            intUID = GetUserID()

            If ((ViewState("UserID") = intUID)) Then

                objDW = New HPQ.Data.DataWrapper()

                Select Case strActionType ' ADD LOGIC TO ACCOMMODATE DIFFERENT PROFILES (i.e. - Group Shared/Shared/Etc) USE EXISTING SOURCE AS REFERENCED IN WJs TASK DESC
                    Case "S" ' Retrieve and Apply Profile (S)

                        strParameters = GetReportProfileParameters(strProfileCode, strProfileID, intUID.ToString())

                        If (strParameters.Length > 0) Then

                            ApplyProfile(strParameters)

                            DisplayMessage("Successfully loaded Profile.  Click the <b>Submit</b> button to apply.", "lblProfileStatus", Drawing.Color.Black)
                            Me.resetFlag.Value = "FALSE"
                            Me.CurrProfileName = Me.ddlRptProfile.SelectedItem.Text
                            Me.profileFlag.Value = "true"

                        Else
                            DisplayMessage("<b>Failed to apply Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
                        End If


                    Case "I" ' Add Profile (I)
                        strParameters = GetParameters()
                        intParamLen = strParameters.Length

                        strProfileName = Me.profileName.Value.Replace("|", "")

                        objComm = objDW.CreateCommand("spAddReportProfile", Data.CommandType.StoredProcedure)

                        objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, intUID.ToString, 8)
                        objDW.CreateParameter(objComm, "@ProfileName", Data.SqlDbType.VarChar, strProfileName, 120)
                        objDW.CreateParameter(objComm, "@ProfileType", Data.SqlDbType.Int, "8", 8)
                        objDW.CreateParameter(objComm, "@PageLayout", Data.SqlDbType.VarChar, "")
                        objDW.CreateParameter(objComm, "@SelectedFilters", Data.SqlDbType.VarChar, strParameters, intParamLen)
                        objDW.CreateParameter(objComm, "@TodayPageLink", Data.SqlDbType.Int, "0", 8)
                        objDW.CreateParameter(objComm, "@DefaultReport", Data.SqlDbType.Int, "0", 8)
                        objDW.CreateParameter(objComm, "@NewID", Data.SqlDbType.Int, "0", 8, ParameterDirection.InputOutput)

                        objDW.ExecuteCommandNonQuery(objComm)

                        DisplayMessage("<b>Successfully added Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
                        PopulateProfileList(True, strProfileName)
                        Me.CurrProfileName = Me.profileName.Value
                        Me.resetFlag.Value = "FALSE"

                        'Session("ASProfileID") = objComm.Parameters("@NewID").Value.ToString 'Me.ddlRptProfile.SelectedValue.ToString
                        'EnableAllElements()

                    Case "U" ' Update Profile (U)
                        ' ADD LOGIC TO CHECK FOR SHARED PROFILES AND IF SHARED, CHECK PERMISSIONS (DELETE, EDIT)

                        If (strProfileCode.Length > 0) Then
                            ' Determine if the user can Edit this Shared Profile
                            blnHasPerm = CanEditSharedProfile(strProfileID, intUID.ToString)
                            strOwnerID = GetOwnerID(strProfileID, intUID.ToString())

                            If (strOwnerID.Trim().Length = 0) Or (strOwnerID = intUID.ToString()) Then
                                blnHasPerm = False
                            End If
                        Else
                            blnHasPerm = True
                            strOwnerID = intUID.ToString
                        End If

                        If (blnHasPerm) Then

                            strParameters = GetParameters()
                            intParamLen = strParameters.Length

                            objComm = objDW.CreateCommand("spUpdateReportProfile", Data.CommandType.StoredProcedure)

                            objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
                            objDW.CreateParameter(objComm, "@ProfileName", Data.SqlDbType.VarChar, strProfileName, 120)
                            objDW.CreateParameter(objComm, "@PageLayout", Data.SqlDbType.VarChar, "")
                            objDW.CreateParameter(objComm, "@SelectedFilters", Data.SqlDbType.VarChar, strParameters, intParamLen)
                            objDW.CreateParameter(objComm, "@TodayPageLink", Data.SqlDbType.Int, "0", 8)
                            objDW.CreateParameter(objComm, "@DefaultReport", Data.SqlDbType.Int, "0", 8)
                            objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, strOwnerID, 8)

                            objDW.ExecuteCommandNonQuery(objComm)

                            DisplayMessage("<b>Successfully updated Profile.</b>", "lblProfileStatus", Drawing.Color.Black)

                            Me.CurrProfileName = Me.ddlRptProfile.SelectedItem.Text
                            Me.resetFlag.Value = "FALSE"

                        Else
                            DisplayMessage("<b>User does not have permission to Edit this Shared profile.</b>", "lblProfileStatus", Drawing.Color.Red)
                        End If


                    Case "D" ' Delete Profile (D)
                        ' ADD LOGIC TO CHECK FOR SHARED PROFILES AND IF SHARED, CHECK PERMISSIONS (DELETE, EDIT)

                        If (strProfileCode.Length > 0) Then
                            ' Determine if the user can Edit this Shared Profile
                            blnHasPerm = CanDeleteSharedProfile(strProfileID, intUID.ToString)
                            strOwnerID = GetOwnerID(strProfileID, intUID.ToString())

                            If (strOwnerID.Trim().Length = 0) Or (strOwnerID = intUID.ToString()) Then
                                blnHasPerm = False
                            End If
                        Else
                            blnHasPerm = True
                            strOwnerID = intUID.ToString
                        End If

                        If (blnHasPerm) Then

                            objComm = objDW.CreateCommand("spDeleteProfile", Data.CommandType.StoredProcedure)
                            objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
                            objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, strOwnerID, 8)

                            objDW.ExecuteCommandNonQuery(objComm)

                            DisplayMessage("<b>Successfully deleted Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
                            PopulateProfileList(True, Nothing)
                            Me.CurrProfileName = ""
                            Me.resetFlag.Value = "TRUE"

                        Else
                            DisplayMessage("<b>User does not have permission to Delete this Shared profile.</b>", "lblProfileStatus", Drawing.Color.Red)
                        End If

                    Case "R" ' Rename Profile (R)

                        If (strProfileCode.Length > 0) Then
                            ' Determine if the user can Edit (Rename) this Shared Profile
                            blnHasPerm = CanEditSharedProfile(strProfileID, intUID.ToString)
                            strOwnerID = GetOwnerID(strProfileID, intUID.ToString())

                            If (strOwnerID.Trim().Length = 0) Or (strOwnerID = intUID.ToString()) Then
                                blnHasPerm = False
                            End If
                        Else
                            blnHasPerm = True
                            strOwnerID = intUID.ToString
                        End If

                        If (blnHasPerm) Then

                            strProfileName = Me.profileName.Value.Replace("|", "").Replace("(Shared)", "").Replace("(Group Shared)", "")
                            strSelProfileName = strProfileName

                            If (strProfileCode = "S") Then
                                strSelProfileName = strSelProfileName & " (Shared)"
                            End If

                            If (strProfileCode = "G") Then
                                strSelProfileName = strSelProfileName & " (Group Shared)"
                            End If

                            objComm = objDW.CreateCommand("spRenameProfile", Data.CommandType.StoredProcedure)

                            objDW.CreateParameter(objComm, "@ID", Data.SqlDbType.Int, strProfileID, 8)
                            objDW.CreateParameter(objComm, "@Name", Data.SqlDbType.VarChar, strProfileName, 120)
                            objDW.CreateParameter(objComm, "@EmployeeID", Data.SqlDbType.Int, strOwnerID, 8)

                            objDW.ExecuteCommandNonQuery(objComm)

                            DisplayMessage("<b>Successfully renamed Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
                            PopulateProfileList(True, strSelProfileName)

                            Me.CurrProfileName = Me.ddlRptProfile.SelectedItem.Text
                            Me.resetFlag.Value = "FALSE"
                        Else
                            DisplayMessage("<b>User does not have permission to Rename this Shared profile.</b>", "lblProfileStatus", Drawing.Color.Red)
                        End If

                End Select

            Else
                DisplayMessage("<b>User credentials not verified for this action.</b>", "lblProfileStatus", Drawing.Color.Red)
            End If

        Catch ex As Exception
            DisplayMessage("<b>ProcessProfileAction</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDT = Nothing
            objComm = Nothing
            objDW = Nothing
        End Try

    End Sub

    Private Function ProfileExists(ByVal strName As String) As Boolean
        Dim blnResult As Boolean = False
        Dim objListItem As ListItem

        For Each objListItem In Me.ddlRptProfile.Items
            If (objListItem.Text.ToUpper.Trim = strName.ToUpper.Trim) Then
                blnResult = True
                Exit For
            End If
        Next

        ProfileExists = blnResult

    End Function

    Protected Sub lnkBtnApplyProfile_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkBtnApplyProfile.Click

        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to load and apply.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ProcessProfileAction("S")
    End Sub

    Protected Sub lnkBtnAddProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnAddProfile.Click

        ' Determine if a new name was specified
        If (Me.continueAction.Value.ToUpper() <> "TRUE") Then
            Exit Sub
        End If

        ' Require the user to specify at least 1 filter parameter (may require more or specific ones in the future, e.g. - dates)
        If (NumParameters() = 0) Then
            DisplayMessage("<b>No filter parameters specified.  Please specify as many filter parameters as possible in order to limit the scope of the results.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ' Prohibit whitespace only and blank Profile Names
        If (Me.profileName.Value.Trim().Length = 0) Then
            DisplayMessage("<b>Invalid Profile Name specified.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        If (ProfileExists(Me.profileName.Value)) Then
            DisplayMessage("<b>Profile Name, '" & Me.profileName.Value & "' already exists.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ProcessProfileAction("I")

    End Sub


    Protected Sub lnkBtnUpdateProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnUpdateProfile.Click

        ' Make certain that the user selected a profile
        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to Update.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ' Require the user to specify at least 1 filter parameter (may require more or specific ones in the future, e.g. - dates)
        If (NumParameters() = 0) Then
            DisplayMessage("<b>No filter parameters specified.  Please specify as many filter parameters as possible in order to limit the scope of the results.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ProcessProfileAction("U")
    End Sub


    Protected Sub lnkBtnRenameProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnRenameProfile.Click

        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to Rename.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ' Determine if a user confirmed action
        If (Me.continueAction.Value.ToUpper() <> "TRUE") Then
            DisplayMessage("<b>User cancelled Renaming of Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
            Exit Sub
        End If

        ' Prohibit whitespace only and blank Profile Names
        If (Me.profileName.Value.Trim().Length = 0) Then
            DisplayMessage("<b>Invalid Profile Name specified.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        If (ProfileExists(Me.profileName.Value)) Then
            DisplayMessage("<b>Profile Name, '" & Me.profileName.Value & "' already exists.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ProcessProfileAction("R")
    End Sub

    Protected Sub lnkBtnDeleteProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnDeleteProfile.Click


        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to Delete.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        ' Determine if a user confirmed action
        If (Me.continueAction.Value.ToUpper() <> "TRUE") Then
            DisplayMessage("<b>User cancelled Profile Deletion.</b>", "lblProfileStatus", Drawing.Color.Black)
            Exit Sub
        End If

        ProcessProfileAction("D")

    End Sub

    Protected Sub lnkBtnShareProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnShareProfile.Click

        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to Share.</b>", "lblProfileStatus", Drawing.Color.Red)
            Exit Sub
        End If

        Dim strProfileID As String = Me.ddlRptProfile.SelectedValue.Replace("G", "").Replace("S", "")
        Dim strUID As String = GetUserID().ToString
        Dim strOwnerID As String = GetOwnerID(strProfileID, strUID)
        Dim strScript As String = "<script type='text/javascript' language='javascript'>var strResult = window.showModalDialog('../Query/ProfileShare.asp?ID=" & strProfileID & "', '', 'dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No');</script>"

        If (strOwnerID = strUID) Then

            ' Define the name and type of the client script on the page. 
            Dim csName As [String] = "SHARESCRIPT"
            Dim csType As Type = Me.[GetType]()

            ' Get a ClientScriptManager reference from the Page class. 
            Dim cs As ClientScriptManager = Page.ClientScript

            ' Check to see if the client script is already registered. 
            If Not cs.IsClientScriptBlockRegistered(csType, csName) Then
                cs.RegisterClientScriptBlock(csType, csName, strScript)
            End If

        Else
            DisplayMessage("<b>Only the Owner of this Profile can Share it.</b>", "lblProfileStatus", Drawing.Color.Red)
        End If

    End Sub

    Function GetListSelections(ByVal objList As ListBox) As String
        Dim strSelections As String = ""

        For Each intIdx As Integer In objList.GetSelectedIndices
            If (Len(strSelections) = 0) Then
                strSelections = objList.Items(intIdx).Value.ToString()

            Else
                strSelections += "," & objList.Items(intIdx).Value.ToString()
            End If
        Next

        GetListSelections = strSelections

    End Function

    Function GetListItems(ByVal objList As ListBox) As String
        Dim strItems As String = ""
        Dim objListItem As ListItem

        For Each objListItem In objList.Items
            If (Len(strItems) = 0) Then
                strItems = objListItem.Value.ToString()
            Else
                strItems += "," & objListItem.Value.ToString()
            End If
        Next

        GetListItems = strItems

    End Function

    Function GetDDListItems(ByVal objList As DropDownList) As String
        Dim strItems As String = ""
        Dim objListItem As ListItem

        For Each objListItem In objList.Items
            If (Len(strItems) = 0) Then
                strItems = objListItem.Value.ToString()
            Else
                strItems += "," & objListItem.Value.ToString()
            End If
        Next

        GetDDListItems = strItems

    End Function


    Function NumParameters() As Integer
        Dim intTotalParameters As Integer = 0

        If (Me.chkbxGeoNA.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxGeoLA.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxGeoAPJ.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxGeoEMEA.Checked) Then intTotalParameters = intTotalParameters + 1

        If (Me.chkbxSKGeoNA.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxSKGeoLA.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxSKGeoAPJ.Checked) Then intTotalParameters = intTotalParameters + 1
        If (Me.chkbxSKGeoEMEA.Checked) Then intTotalParameters = intTotalParameters + 1

        intTotalParameters = intTotalParameters + Me.lstbxBrands.GetSelectedIndices.Length
        intTotalParameters = intTotalParameters + Me.lstbxCats.GetSelectedIndices.Length
        intTotalParameters = intTotalParameters + Me.lstbxOSSPs.GetSelectedIndices.Length

        If (Len(Trim(Me.txtbxKMATS.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxSKUS.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxAVS.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxSKs.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxSFPNS.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxSKUAVS.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxSAs.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxComps.Text)) > 0) Then intTotalParameters = intTotalParameters + 1

        If (Me.chkbxAdded.Checked) Then intTotalParameters = intTotalParameters + 1 ' May want to suggest date range if none specified
        If (Me.chkbxUpdated.Checked) Then intTotalParameters = intTotalParameters + 1 ' May want to suggest date range if none specified

        If (Len(Trim(Me.txtbxFromDate.Text)) > 0) Then intTotalParameters = intTotalParameters + 1
        If (Len(Trim(Me.txtbxToDate.Text)) > 0) Then intTotalParameters = intTotalParameters + 1

        NumParameters = intTotalParameters

    End Function


    Protected Sub ddlRowsPerPage_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlRowsPerPage.SelectedIndexChanged

        ' Disabled per WJ
    End Sub

    Protected Sub EnableAllElements()

        ' Enable Update link
        Me.lnkBtnUpdateProfile.Visible = True
        Me.lnkBtnUpdateProfile.Enabled = True

        ' Enable Delete link
        Me.lnkBtnDeleteProfile.Visible = True
        Me.lnkBtnDeleteProfile.Enabled = True

        ' Enable Rename link
        Me.lnkBtnRenameProfile.Visible = True
        Me.lnkBtnRenameProfile.Enabled = True

        ' Enable/Disable Page Elements
        Me.txtbxSFPNS.Enabled = True
        Me.txtbxKMATS.Enabled = True
        Me.lstbxBrands.Enabled = True
        Me.rdoProdDiv.Enabled = True
        Me.txtbxSKUS.Enabled = True
        Me.txtbxSKUAVS.Enabled = True
        Me.chkbxAllSKUAVs.Enabled = True
        Me.chkbxGeoNA.Enabled = True
        Me.chkbxGeoLA.Enabled = True
        Me.chkbxGeoAPJ.Enabled = True
        Me.chkbxGeoEMEA.Enabled = True
        Me.lstbxOSSPs.Enabled = True
        Me.txtbxSKs.Enabled = True
        Me.chkbxSKGeoNA.Enabled = True
        Me.chkbxSKGeoLA.Enabled = True
        Me.chkbxSKGeoAPJ.Enabled = True
        Me.chkbxSKGeoEMEA.Enabled = True
        Me.lstbxCats.Enabled = True
        Me.txtbxAVS.Enabled = True
        Me.chkbxAllAVs.Enabled = True
        Me.txtbxSAs.Enabled = True
        Me.txtbxComps.Enabled = True
        Me.chkbxAdded.Enabled = True
        Me.chkbxUpdated.Enabled = True
        Me.txtbxFromDate.Enabled = True
        Me.lnkBtnFromDate.Enabled = True
        Me.lnkBtnFromDate.Visible = True
        Me.txtbxToDate.Enabled = True
        Me.lnkBtnToDate.Enabled = True
        Me.lnkBtnToDate.Visible = True

        Me.ddlActionDateColumn.Enabled = True
        Me.ddlActionDateColumn.Visible = True

        'Me.selectedColumns.Enabled = True
        Me.allColumns.Enabled = True

        Me.chkbxSPBLogFilter.Enabled = True

        Me.chkbxSKUAVLogFilter.Enabled = True

        ' Add code for HTML elements

    End Sub

    Protected Sub ddlRptProfile_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlRptProfile.SelectedIndexChanged

        ' ADD LOGIC TO DETERMINE IF THE SELECTED PROFILE IS (1) THE USER'S PROFILE (2) A SHARED PROFILE (3) A GROUP SHARED PROFILE
        If Me.ddlRptProfile.SelectedValue = "0" Then
            DisplayMessage("<b>Please select a Profile to load.</b>", "lblProfileStatus", Drawing.Color.Red)

            Me.lnkBtnRemoveSharedProfile.Enabled = False
            Me.lnkBtnRemoveSharedProfile.Visible = False
            Me.lnkBtnShareProfile.Enabled = False
            Me.lnkBtnShareProfile.Visible = False
            EnableAllElements()
            Exit Sub
        End If

        ' Check Selection and Enable/Disable applicable elements
        Dim strProfileCode As String = GetProfileCode()
        Dim strProfileID As String = ""
        Dim intUID As Integer = GetUserID()
        Dim blnIsRemovedSharedProfile As Boolean = False
        Dim strRemovedSharedProfileList As String = ""

        ' ******************************** ACCOMMODATE CLIENT SIDE VALIDATION (OR LACK THEREOF) ******************************** 

        If (strProfileCode.Length > 0) Then

            strProfileID = Me.ddlRptProfile.SelectedValue.ToString().Replace(strProfileCode, "")
            strRemovedSharedProfileList = GetRemovedSharedProfilesList(intUID.ToString)

            ' Disable Share link
            Me.lnkBtnShareProfile.Visible = False
            Me.lnkBtnShareProfile.Enabled = False

            ' Enable/Disable Update link
            Me.lnkBtnUpdateProfile.Visible = CanEditSharedProfile(strProfileID, intUID.ToString())
            Me.lnkBtnUpdateProfile.Enabled = Me.lnkBtnUpdateProfile.Visible

            ' Enable/Disable Delete link
            Me.lnkBtnDeleteProfile.Visible = CanDeleteSharedProfile(strProfileID, intUID.ToString())
            Me.lnkBtnDeleteProfile.Enabled = Me.lnkBtnDeleteProfile.Visible

            ' Enable/Disable Rename link
            Me.lnkBtnRenameProfile.Visible = Me.lnkBtnUpdateProfile.Visible
            Me.lnkBtnRenameProfile.Enabled = Me.lnkBtnUpdateProfile.Visible

            ' Determine Text value for Remove Share Profile link
            blnIsRemovedSharedProfile = IsRemovedSharedProfile(strRemovedSharedProfileList, Integer.Parse(strProfileID))
            If (blnIsRemovedSharedProfile) Then
                Me.lnkBtnRemoveSharedProfile.Text = "Restore"
                Me.lnkBtnRemoveSharedProfile.ToolTip = "Restore the selected Removed Shared Profile"
            Else
                Me.lnkBtnRemoveSharedProfile.Text = "Remove"
                Me.lnkBtnRemoveSharedProfile.ToolTip = "Remove the selected Shared Profile"
            End If

            ' Enable Remove Shared Profile link
            Me.lnkBtnRemoveSharedProfile.Visible = True
            Me.lnkBtnRemoveSharedProfile.Enabled = True


            ' Enable/Disable Page Elements
            Me.txtbxSFPNS.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxKMATS.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lstbxBrands.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.rdoProdDiv.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxSKUS.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxSKUAVS.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxAllSKUAVs.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxGeoNA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxGeoLA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxGeoAPJ.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxGeoEMEA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lstbxOSSPs.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxSKs.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxSKGeoNA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxSKGeoLA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxSKGeoAPJ.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxSKGeoEMEA.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lstbxCats.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxAVS.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxAllAVs.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxSAs.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxComps.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxAdded.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.chkbxUpdated.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxFromDate.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lnkBtnFromDate.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lnkBtnFromDate.Visible = Me.lnkBtnUpdateProfile.Visible
            Me.txtbxToDate.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lnkBtnToDate.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.lnkBtnToDate.Visible = Me.lnkBtnUpdateProfile.Visible

            Me.ddlActionDateColumn.Enabled = Me.lnkBtnUpdateProfile.Visible
            'Me.ddlActionDateColumn.Visible = Me.lnkBtnUpdateProfile.Visible

            'Me.selectedColumns.Enabled = Me.lnkBtnUpdateProfile.Visible
            Me.allColumns.Enabled = Me.lnkBtnUpdateProfile.Visible

            Me.chkbxSPBLogFilter.Enabled = Me.lnkBtnUpdateProfile.Visible

            Me.chkbxSKUAVLogFilter.Enabled = Me.lnkBtnUpdateProfile.Visible

            ' Add code for HTML elements

        Else
            strProfileID = "0"

            ' Enable Share link
            Me.lnkBtnShareProfile.Visible = True
            Me.lnkBtnShareProfile.Enabled = True

            ' Disable Remove Shared Profile link
            Me.lnkBtnRemoveSharedProfile.Visible = False
            Me.lnkBtnRemoveSharedProfile.Enabled = False
            Me.lnkBtnRemoveSharedProfile.Text = "Remove"
            Me.lnkBtnRemoveSharedProfile.ToolTip = "Remove the selected Shared Profile"

            EnableAllElements()

        End If


        ProcessProfileAction("S")

    End Sub


    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        EnableAllElements()

        Me.ddlRptProfile.SelectedIndex = 0

        Me.chkbxGeoNA.Checked = False
        Me.chkbxGeoLA.Checked = False
        Me.chkbxGeoAPJ.Checked = False
        Me.chkbxGeoEMEA.Checked = False

        Me.chkbxSKGeoNA.Checked = False
        Me.chkbxSKGeoLA.Checked = False
        Me.chkbxSKGeoAPJ.Checked = False
        Me.chkbxSKGeoEMEA.Checked = False

        Me.lstbxBrands.SelectedIndex = -1
        Me.lstbxCats.SelectedIndex = -1
        Me.lstbxOSSPs.SelectedIndex = -1
        Me.ddlActionDateColumn.SelectedValue = 0 '.SelectedIndex = -1

        Me.txtbxKMATS.Text = ""
        Me.txtbxSKUS.Text = ""
        Me.txtbxAVS.Text = ""
        Me.txtbxSKs.Text = ""
        Me.txtbxSFPNS.Text = ""
        Me.txtbxSKUAVS.Text = ""
        Me.txtbxSAs.Text = ""
        Me.txtbxComps.Text = ""

        Me.chkbxAllSKUAVs.Checked = False

        Me.chkbxAdded.Checked = False
        Me.chkbxUpdated.Checked = False

        Me.txtbxFromDate.Text = ""
        Me.txtbxToDate.Text = ""

        Me.chkbxAllAVs.Checked = False

        Me.rdoProdDiv.SelectedValue = 0

        Me.chkbxSPBLogFilter.Checked = False

        Me.chkbxSKUAVLogFilter.Checked = False

        DisplayMessage("<b>Select a Profile to load.</b>", "lblProfileStatus", Drawing.Color.Black)
        DisplayMessage("Set the desired filter options and click the <b>Submit</b> button to retrieve applicable data.", "lblStatus", Drawing.Color.Black)


    End Sub


    Protected Sub SetProductFilter(ByVal strDevCenterID As String)
        Dim strFilter As String = ""

        If (strDevCenterID <> "0") Then
            If (strDevCenterID = "1") Then
                strFilter = "DevCenter<>2"
            Else
                strFilter = "DevCenter=2"
            End If
        End If

        Me.dsBrands.FilterExpression = strFilter

        Me.lstbxBrands.DataBind()

        ' Me.lstbxBrands.ClearSelection()

    End Sub

    Protected Sub rdoProdDiv_SelectedIndexedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoProdDiv.SelectedIndexChanged
        ' CONSIDER FLAG TO DETERMINE CALLING EVENT
        SetProductFilter(Me.rdoProdDiv.SelectedValue)

    End Sub

    Protected Function RemoveSharedProfile(ByVal strProfileID As String, ByVal strUID As String) As Boolean
        Dim blnResult As Boolean = False
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim intResult As Integer

        Try
            objDW = New HPQ.Data.DataWrapper()
            If (RestoreSharedProfile(strProfileID, strUID)) Then
                intResult = objDW.ExecuteSqlNonQuery("INSERT INTO dbo.BTOSSREMOVEDPROFILES(PROFILE_ID, USER_ID) VALUES(" & strProfileID & ", " & strUID & ")")
                blnResult = True
                DisplayMessage("<b>Successfully Removed Shared Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
            End If
        Catch ex As Exception
            DisplayMessage("<b>Failed to remove Shared Profile.</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDT = Nothing
            objComm = Nothing
            objDW = Nothing
        End Try

        RemoveSharedProfile = blnResult

    End Function


    Protected Function RestoreSharedProfile(ByVal strProfileID As String, ByVal strUID As String) As Boolean
        Dim blnResult As Boolean = False
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim intResult As Integer

        Try
            objDW = New HPQ.Data.DataWrapper()
            intResult = objDW.ExecuteSqlNonQuery("DELETE FROM dbo.BTOSSREMOVEDPROFILES WHERE PROFILE_ID=" & strProfileID & " AND USER_ID=" & strUID)
            blnResult = True
            DisplayMessage("<b>Successfully Restored Removed Shared Profile.</b>", "lblProfileStatus", Drawing.Color.Black)
        Catch ex As Exception
            DisplayMessage("<b>Failed to restore Removed Shared Profile.</b> - " & ex.Message, "lblProfileStatus", Drawing.Color.Red)
        Finally
            objDT = Nothing
            objComm = Nothing
            objDW = Nothing
        End Try

        RestoreSharedProfile = blnResult

    End Function

    Protected Sub lnkBtnRemoveSharedProfile_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnRemoveSharedProfile.Click
        ' Add logic to determine whether this is REMOVAL OR RESTORE ACTION
        Dim blnRestore As Boolean = False

        If (Me.lnkBtnRemoveSharedProfile.Text.ToUpper.Trim = "RESTORE") Then
            blnRestore = True
        End If

        If Me.ddlRptProfile.SelectedValue = "0" Then
            If (Not blnRestore) Then
                DisplayMessage("<b>Please select a Shared Profile to Remove.</b>", "lblProfileStatus", Drawing.Color.Red)
            Else
                DisplayMessage("<b>Please select a Removed Shared Profile to Restore.</b>", "lblProfileStatus", Drawing.Color.Red)
            End If

            Exit Sub
        End If

        ' Determine if a user confirmed action
        If (Me.continueAction.Value.ToUpper() <> "TRUE") Then

            If (Not blnRestore) Then
                DisplayMessage("<b>User cancelled Shared Profile Removal.</b>", "lblProfileStatus", Drawing.Color.Black)
            Else
                DisplayMessage("<b>User cancelled Removed Shared Profile Restoration.</b>", "lblProfileStatus", Drawing.Color.Black)
            End If

            Exit Sub
        End If

        Dim strProfileID As String = Me.ddlRptProfile.SelectedValue.Replace("G", "").Replace("S", "")
        Dim strUID As String = GetUserID().ToString

        If (Not blnRestore) Then
            Dim strOwnerID As String = GetOwnerID(strProfileID, strUID)

            If (strOwnerID <> strUID) Then
                If (RemoveSharedProfile(strProfileID, strUID)) Then
                    PopulateProfileList(False, Nothing)
                    Me.lnkBtnRemoveSharedProfile.Enabled = False
                    Me.lnkBtnRemoveSharedProfile.Visible = False
                    Me.lnkBtnShareProfile.Enabled = False
                    Me.lnkBtnShareProfile.Visible = False
                    EnableAllElements()
                End If
            Else
                DisplayMessage("<b>Only Profiles Shared by another user may be removed.</b>", "lblProfileStatus", Drawing.Color.Red)
            End If
        Else
            If (RestoreSharedProfile(strProfileID, strUID)) Then
                PopulateProfileList(False, Me.ddlRptProfile.SelectedItem.Text)
                Me.lnkBtnRemoveSharedProfile.Text = "Remove"
                Me.lnkBtnRemoveSharedProfile.ToolTip = "Remove the selected Shared Profile"
            End If
        End If

    End Sub


    Protected Sub chkBxIncludeRemSProfs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkBxIncludeRemSProfs.CheckedChanged

        PopulateProfileList(False, Nothing)

        If (Me.chkBxIncludeRemSProfs.Checked) Then
            DisplayMessage("<b>Successfully retrieved Removed Shared Profile(s).</b>", "lblProfileStatus", Drawing.Color.Black)
        End If

    End Sub

End Class
