<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Xml" %>


<script runat="server">

    Protected Sub OnSelectedIndexChangedMethod(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim rBtnList As RadioButton = CType(sender, RadioButton)
        Dim BusUnit As String

        BusUnit = rBtnList.Text

        '  Response.Write(BusUnit & "<br>")
        If BusUnit = "Commercial" Then
            BusUnit = "1"
        Else
            BusUnit = "2"
        End If

        Dim txtBU As TextBox = dvSpbDetails.FindControl("txtBusinessUnit")
        txtBU.Text = BusUnit

    End Sub

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)

        If dvSpbDetails.CurrentMode = DetailsViewMode.Edit Then

            Dim dt As DataTable = HPQ.Excalibur.Employee.ListEmployees()
            Dim ddlGplm As DropDownList = dvSpbDetails.FindControl("ddlGplmUser")
            Dim ddlBomAnalyst As DropDownList = dvSpbDetails.FindControl("ddlSpdmUser")

            Try
                For Each row As DataRow In dt.Rows
                    Page.ClientScript.RegisterForEventValidation(ddlGplm.UniqueID, row("ID").ToString())
                    Page.ClientScript.RegisterForEventValidation(ddlBomAnalyst.UniqueID, row("ID").ToString())
                Next
            Catch ex As Exception
                Response.Write(ex.Message)
            End Try

        End If

        MyBase.Render(writer)
    End Sub

    Enum UserType
        Spdm
        RPLM
    End Enum

    ReadOnly Property ProductVersionId() As String
        Get
            Return Request.QueryString("PVID")
        End Get
    End Property

    Private strSFPN As String = vbNull

    Property SFPN() As String
        Get
            Return strSFPN
        End Get
        Set(ByVal value As String)
            strSFPN = value
        End Set
    End Property

    Private strDeletedIDs As String

    Property DeletedIDs() As String
        Get
            Return strDeletedIDs
        End Get
        Set(ByVal value As String)
            strDeletedIDs = value
        End Set
    End Property


    '******************************************************************************************************************************************
    ' PERMISSION PROPERTIES BASED UPON ROLE(S)
    '******************************************************************************************************************************************

    Private blnIsSpdmUser As Boolean = False

    Property IsSpdmUser As Boolean
        Get
            Return blnIsSpdmUser
        End Get
        Set(ByVal value As Boolean)
            blnIsSpdmUser = value
        End Set
    End Property

    Private blnIsRPLMUser As Boolean = False

    Property IsRPLMUser As Boolean
        Get
            Return blnIsRPLMUser
        End Get
        Set(ByVal value As Boolean)
            blnIsRPLMUser = value
        End Set
    End Property

    ' Flag for Edit Permission of Product Version's Service Family Part Number
    Private blnCanEditSFPN As Boolean = False

    Property CanEditSFPN As Boolean
        Get
            Return blnCanEditSFPN
        End Get
        Set(ByVal value As Boolean)
            blnCanEditSFPN = value
        End Set
    End Property

    ' Flag for Edit Permission of Service Family Part Number's Detail
    Private blnCanEditSFPNDetails As Boolean = False

    Property CanEditSFPNDetails As Boolean
        Get
            Return blnCanEditSFPNDetails
        End Get
        Set(ByVal value As Boolean)
            blnCanEditSFPNDetails = value
        End Set
    End Property

    ' Flag for Edit Permission of Service Family Part Number's OSSP Assignments
    Private blnCanEditOSSPDetails As Boolean = False

    Property CanEditOSSPDetails As Boolean
        Get
            Return blnCanEditOSSPDetails
        End Get
        Set(ByVal value As Boolean)
            blnCanEditOSSPDetails = value
        End Set
    End Property

    '******************************************************************************************************************************************

    Private strReturnCode As String = "0"
    Property ReturnCode As String
        Get
            Return strReturnCode
        End Get
        Set(ByVal value As String)
            strReturnCode = value
        End Set
    End Property

    Private strReturnDesc As String = ""
    Property ReturnDesc As String
        Get
            Return strReturnDesc
        End Get
        Set(ByVal value As String)
            strReturnDesc = value
        End Set
    End Property

    '******************************************************************************************************************************************

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        Response.Cache.SetExpires(DateTime.Now())
        Response.Cache.SetCacheability(HttpCacheability.NoCache)


        If Not Page.IsPostBack Then

            Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

            Dim strSFPN As String = hpqData.GetServiceFamilyPn(ProductVersionId)
            Me.SFPN = strSFPN

            txtFamilyPn.Text = strSFPN

            ' Initialize original value of FamilyPn 
            ViewState("FamilyPn") = txtFamilyPn.Text

            If txtFamilyPn.Text.Trim <> String.Empty Then
                lblFamilyPn.Text = txtFamilyPn.Text
                txtFamilyPn.Visible = False
                lblFamilyPn.Visible = True
                btnSaveFamilyPn.Visible = False
                btnEditFamilyPn.Visible = True
                dvSpbDetails.Visible = True

                InitializeOSSPAssignGrid()
            Else
                grdSFPartners.Visible = False
                btnSaveSFP.Visible = False
                btnCancelSFP.Visible = False

                DisplayMessage("<b>Service Family not specified.</b>", "lblStatus", Drawing.Color.Red)
            End If


            ' Initialize User Roles/Permissions
            Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(HttpContext.Current.User.Identity.Name.ToString())

            Me.IsSpdmUser = (objSec.IsSysAdmin Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.GPLM) Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.ServiceBomAnalyst))
            Me.IsRPLMUser = objSec.UserInRole("RPLM")

            ' Set Edit Permissions
            Me.CanEditSFPN = Me.IsSpdmUser
            Me.CanEditSFPNDetails = Me.IsSpdmUser
            Me.CanEditOSSPDetails = Me.IsSpdmUser Or Me.IsRPLMUser

            'Response.Write("SpdmUser=" & CStr(Me.IsSpdmUser) & ", ")
            'Response.Write("RPLM=" & CStr(Me.IsRPLMUser) & ", ")
            'Response.Write("CanEditSFPN=" & CStr(Me.CanEditSFPN) & ", ")
            'Response.Write("CanEditSFPNDetails=" & CStr(Me.CanEditSFPNDetails) & ", ")
            'Response.Write("CanEditOSSPDetails=" & CStr(Me.CanEditOSSPDetails))

            ViewState("IsSpdmUser") = Me.IsSpdmUser
            ViewState("IsRPLMUser") = Me.IsRPLMUser
            ViewState("SFPN") = Me.SFPN


            objSec = Nothing
            hpqData = Nothing

            ' Determine if the user can edit First Section of Service Family Details
            Dim objDetView As DetailsView = CType(Me.FindControl("dvSpbDetails"), DetailsView)

            If Me.CanEditSFPNDetails Then

                Dim objComField As CommandField = New CommandField

                objComField.ShowEditButton = True
                objDetView.Fields.Add(objComField)
                objComField = Nothing

            End If


        Else ' Subsequent load

            'Set the Deletion IDs
            Me.DeletedIDs = ViewState("DeletedIDs")

            ' Set the Service Family Part Number
            Me.SFPN = ViewState("SFPN")

            ' Set User Roles   ---> Consider calling object methods each time for security reasons
            Me.IsSpdmUser = IsAuthorizedUser(UserType.Spdm)
            Me.IsRPLMUser = IsAuthorizedUser(UserType.RPLM)

            ' Set Edit Permissions
            Me.CanEditSFPN = Me.IsSpdmUser
            Me.CanEditSFPNDetails = Me.IsSpdmUser
            Me.CanEditOSSPDetails = Me.IsSpdmUser Or Me.IsRPLMUser

        End If

    End Sub

    Protected Sub setDirtyFlag()
        Dim txtDirtyFlag As HiddenField = CType(FindControl("dirtyFlag"), HiddenField)
        txtDirtyFlag.Value = "true"
    End Sub

    Protected Sub InitializeOSSPAssignGrid()

        ' Initialize [OSSP Partners] Grid and ViewState
        Dim objDT As DataTable

        objDT = HPQ.Excalibur.Service.SelectServiceFamilyPartnerDetails("", Me.SFPN, "", "", "A", "", strReturnCode, strReturnDesc)

        ViewState("Partners") = objDT
        grdSFPartners.DataSource = objDT
        grdSFPartners.DataBind()

        DisplayMessage(strReturnDesc, "lblStatus", Drawing.Color.Black)

        objDT = Nothing

        ' Initialize Deletion IDs
        Me.DeletedIDs = ""
        ViewState("DeletedIDs") = Me.DeletedIDs


        Me.SFPDirtyFlag.Value = "false"
        Me.SFPSaved.Value = "null" ' Reset the flag for prompting to save the OSSP Assignment data changes
    End Sub


    Protected Sub btnSaveFamilyPn_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        ' Verify Permissions, again
        Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(HttpContext.Current.User.Identity.Name.ToString())

        Me.IsSpdmUser = (objSec.IsSysAdmin Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.GPLM) Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.ServiceBomAnalyst))
        Me.IsRPLMUser = objSec.UserInRole("RPLM")

        ViewState("IsSpdmUser") = Me.IsSpdmUser
        ViewState("IsRPLMUser") = Me.IsRPLMUser

        objSec = Nothing

        If (Not Me.CanEditSFPN) Then ' Should not really reach this, since controls/element will not exist on rendered page unless user has proper permissions
            DisplayMessage("<b>Current User is not authorized to alter this data.</b>", "lblStatusTop", System.Drawing.Color.Red)
            Exit Sub
        End If

        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

        hpqData.SetServiceFamilyPn(ProductVersionId, txtFamilyPn.Text)

        Me.SFPN = txtFamilyPn.Text
        ViewState("SFPN") = Me.SFPN

        If txtFamilyPn.Text.Trim() <> String.Empty Then
            lblFamilyPn.Text = txtFamilyPn.Text
            lblFamilyPn.Visible = True
            txtFamilyPn.Visible = False
            btnSaveFamilyPn.Visible = False
            btnEditFamilyPn.Visible = True
            dvSpbDetails.Visible = True

            ' MAY NEED TO CHECK PERMISSIONS AGAIN BEFORE PROCESSING CODE IMMEDIATELY BELOW, BUT PROBABLY NOT SINCE PERMISSIONS THAT ALLOW UPDATING OF SFPN ALSO ALLOW OSSP DETAIL UPDATES
            grdSFPartners.Visible = True
            btnSaveSFP.Visible = True
            btnCancelSFP.Visible = True

            ' Restore grid to it's original state
            InitializeOSSPAssignGrid()

        Else
            dvSpbDetails.Visible = False
            grdSFPartners.Visible = False
            btnSaveSFP.Visible = False
            btnCancelSFP.Visible = False

            DisplayMessage("<b>Service Family not specified.</b>", "lblStatus", Drawing.Color.Red)

        End If

        ' Update dirty flag, if applicable
        If (ViewState("FamilyPn").ToString <> txtFamilyPn.Text) Then ' FOR REFRESHING OF PARENT PAGE ONLY
            setDirtyFlag()
        End If

        hpqData = Nothing

    End Sub

    Protected Sub btnEditFamilyPn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (Not Me.CanEditSFPN) Then ' Should not really reach this, since controls/element will not exist on rendered page unless user has proper permissions
            DisplayMessage("<b>Current User is not authorized to alter this data.</b>", "lblStatusTop", System.Drawing.Color.Red)
            Exit Sub
        End If

        btnSaveFamilyPn.Visible = True
        btnEditFamilyPn.Visible = False
        lblFamilyPn.Visible = False
        txtFamilyPn.Visible = True
    End Sub

    Protected Sub dvSpbDetails_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles dvSpbDetails.DataBound
        If dvSpbDetails.DataItemCount = 0 Then
            dvSpbDetails.AutoGenerateInsertButton = True
            dvSpbDetails.ChangeMode(DetailsViewMode.Insert)
        End If
    End Sub

    Protected Sub dvSpbDetails_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewInsertedEventArgs) Handles dvSpbDetails.ItemInserted
        dvSpbDetails.AutoGenerateInsertButton = False

        ' Update dirty flag
        setDirtyFlag()
    End Sub

    Protected Sub dvSpbDetails_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdatedEventArgs)
        odsGplmUsers.Select()
        odsSpdmUsers.Select()
        odsSvcManagers.Select()

        ' Update dirty flag
        setDirtyFlag()
    End Sub

    Protected Sub dvSpbDetails_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdateEventArgs)
        e.NewValues(0) = Request.Form("dvSpbDetails$ddlGplmUser")
        e.NewValues(1) = Request.Form("dvSpbDetails$ddlSpdmUser")
    End Sub



    '*******************************************************************************************************************************************************************************************************************
    ' ADDED LOGIC TO IMPLEMENT MAINTENANCE OF [ProductVer_Partners] TABLE
    '*******************************************************************************************************************************************************************************************************************

    Protected Sub AddPartnerRecord(ByVal strPartnerID As String, ByVal strPartnerName As String, ByVal strServiceGeoID As String, ByVal strServiceGeoShortName As String, ByVal strServicePartnerTypeCode As String)

        If Not IsUniqueRecord(strPartnerID, strServiceGeoID, strServicePartnerTypeCode) Then
            DisplayMessage("The <b>OSSP</b>, <b>Geo</b>, and <b>Type</b> MUST be unique to this <b>Service Family</b>.", "lblStatus", System.Drawing.Color.Red)
            Exit Sub
        End If

        Dim objDtaGrd As DataGrid = DirectCast(Me.grdSFPartners, DataGrid)
        Dim objDT As DataTable = ViewState("Partners")
        Dim objRow As DataRow

        Dim intID As Integer = -(objDT.Rows.Count)

        Try

            objRow = objDT.NewRow()

            objRow("ID") = intID.ToString
            objRow("ServiceFamilyPn") = Me.ProductVersionId
            objRow("PartnerID") = strPartnerID
            objRow("Name") = strPartnerName
            objRow("ServiceGeoID") = strServiceGeoID
            objRow("ServiceGeoShortName") = strServiceGeoShortName
            objRow("ServicePartnerTypeCode") = strServicePartnerTypeCode

            objDT.Rows.Add(objRow)

            ' UPDATE THE GRID
            ViewState("Partners") = objDT
            objDtaGrd.DataSource = objDT
            objDtaGrd.DataBind()

            objDT = Nothing

            ' Update dirty flag
            'setDirtyFlag()

            Me.SFPDirtyFlag.Value = "true"
            UpdateSFPSavedField(False)

            DisplayMessage("Successfully added new record.", "lblStatus", System.Drawing.Color.Black)

        Catch ex As Exception
            DisplayMessage("Failed to add new record.  Error Description: " & ex.Message, "lblStatus", System.Drawing.Color.Red)
        End Try

    End Sub


    Protected Sub DeletePartnerRecord(ByVal strID As String)
        Dim objDtaGrd As DataGrid = DirectCast(Me.grdSFPartners, DataGrid)

        Dim objDT As DataTable = ViewState("Partners")
        Dim objRow As DataRow

        Dim intIdx As Integer

        Try

            For intIdx = objDT.Rows.Count - 1 To 0 Step -1

                objRow = objDT.Rows(intIdx)

                ' NEED TO ACCOMMODATE ROWS IN THIS SESSION
                If (objRow.RowState <> DataRowState.Deleted) Then
                    If (objRow("ID").ToString().Trim() = strID) Then
                        objRow.Delete()

                        If (Me.DeletedIDs.Length = 0) Then
                            Me.DeletedIDs = strID
                        Else
                            Me.DeletedIDs += "," & strID
                        End If

                        ViewState("DeletedIDs") = Me.DeletedIDs

                    End If
                End If

            Next

            ' UPDATE THE GRID
            ViewState("Partners") = objDT
            objDtaGrd.DataSource = objDT
            objDtaGrd.DataBind()

            'setDirtyFlag()
            Me.SFPDirtyFlag.Value = "true"
            UpdateSFPSavedField(False)

            DisplayMessage("Successfully deleted record.", "lblStatus", System.Drawing.Color.Black)

        Catch ex As Exception
            DisplayMessage("Failed to delete record.  Error Description: " & ex.Message, "lblStatus", System.Drawing.Color.Red)
        End Try

    End Sub


    Protected Sub grdSFPartners_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles grdSFPartners.ItemCommand

        Select Case e.CommandName.ToUpper
            Case "INSERT"  ' NEED TO ADD VALIDATION SCRIPTS ON CLIENT to require explicit selection of ALL lists AND CONSTRAINT ENFORCEMENT
                Dim objDDLOSSP As DropDownList = DirectCast(e.Item.FindControl("add_ddlOSSP"), DropDownList)
                Dim objDDLServiceGeo As DropDownList = DirectCast(e.Item.FindControl("add_ddlServiceGeo"), DropDownList)
                Dim objDDLSPTC As DropDownList = DirectCast(e.Item.FindControl("add_ddlSPTC"), DropDownList)

                If (objDDLOSSP.SelectedItem.Value <> "" And objDDLServiceGeo.SelectedItem.Value <> "" And objDDLSPTC.SelectedItem.Value <> "") Then
                    AddPartnerRecord(objDDLOSSP.SelectedItem.Value, objDDLOSSP.SelectedItem.Text, objDDLServiceGeo.SelectedItem.Value, objDDLServiceGeo.SelectedItem.Text, objDDLSPTC.SelectedItem.Value)
                ElseIf (objDDLOSSP.SelectedItem.Value = "") Then
                    objDDLOSSP.Focus()
                    DisplayMessage("<b>OSSP</b> not specified.", "lblStatus", System.Drawing.Color.Red)
                ElseIf (objDDLServiceGeo.SelectedItem.Value = "") Then
                    objDDLServiceGeo.Focus()
                    DisplayMessage("Service <b>Geo</b> not specified.", "lblStatus", System.Drawing.Color.Red)
                ElseIf (objDDLSPTC.SelectedItem.Value = "") Then
                    objDDLSPTC.Focus()
                    DisplayMessage("<b>Type</b> not specified.", "lblStatus", System.Drawing.Color.Red)
                End If

                'Case "UPDATE"

            Case "DELETE" ' MAY WANT TO ADD CONFIRMATION BY USER
                Dim btnLink As WebControls.LinkButton = DirectCast(e.Item.FindControl("btnDeletePartnerRecord"), LinkButton)
                Dim strID As String = btnLink.Attributes("RecordID")

                DeletePartnerRecord(strID)
        End Select

    End Sub


    Private Function IsUniqueRecord(ByVal strPartnerID As String, ByVal strServiceGeoID As String, ByVal strServicePartnerTypeCode As String) As Boolean
        Dim blnResult As Boolean = True
        Dim objDT As DataTable = ViewState("Partners")

        For Each objRow As DataRow In objDT.Rows
            If (objRow.RowState <> DataRowState.Deleted) Then
                If (objRow("PartnerID") = strPartnerID And objRow("ServiceGeoID") = strServiceGeoID And objRow("ServicePartnerTypeCode") = strServicePartnerTypeCode) Then
                    blnResult = False
                    Exit For
                End If
            End If
        Next

        IsUniqueRecord = blnResult

    End Function

    Private Sub UpdateSFPSavedField(ByVal blnValue As Boolean)
        Dim txtSaveSFPFlag As HiddenField = CType(FindControl("SFPSaved"), HiddenField)
        Dim txtSFPWarnFlag As HiddenField = CType(FindControl("Warned"), HiddenField)

        txtSaveSFPFlag.Value = blnValue.ToString.ToLower
        txtSFPWarnFlag.Value = "false"

    End Sub

    Private Function IsAuthorizedUser(ByVal enUserType As UserType) As Boolean
        Dim blnResult As Boolean = False
        ' Consider replacing with object method call each time
        Select Case enUserType
            Case UserType.Spdm : blnResult = ViewState("IsSpdmUser")
            Case UserType.RPLM : blnResult = ViewState("IsRPLMUser")
        End Select

        IsAuthorizedUser = blnResult

    End Function


    Private Function SaveToDatabase() As Boolean

        Dim strUserName As String = HttpContext.Current.User.Identity.Name.ToString()

        ' Verify Permissions, again
        Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(HttpContext.Current.User.Identity.Name.ToString())

        Me.IsSpdmUser = (objSec.IsSysAdmin Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.GPLM) Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.ServiceBomAnalyst))
        Me.IsRPLMUser = objSec.UserInRole("RPLM")

        ViewState("IsSpdmUser") = Me.IsSpdmUser
        ViewState("IsRPLMUser") = Me.IsRPLMUser

        ' Set Edit Permissions
        Me.CanEditSFPN = Me.IsSpdmUser
        Me.CanEditSFPNDetails = Me.IsSpdmUser
        Me.CanEditOSSPDetails = Me.IsSpdmUser Or Me.IsRPLMUser


        objSec = Nothing

        If (Not Me.CanEditOSSPDetails) Then ' Should not really reach this, since controls/element will not exist on rendered page unless user has proper permissions
            DisplayMessage("<b>Current User is not authorized to alter this data.</b>", "lblStatus", System.Drawing.Color.Red)
            SaveToDatabase = False

            Exit Function
        End If

        Dim blnResult As Boolean = False
        Dim objDT As DataTable = ViewState("Partners")
        Dim strID As String
        Dim strPVID As String = Me.ProductVersionId
        Dim strPartnerID As String
        Dim strSGeoID As String
        Dim strSPTCode As String
        Dim lngResult As Long
        Dim strReturnCd As String = "0"
        Dim strReturnDesc As String = ""

        ' Iterate through Data Table, Adding and Deleting (Changing the [Status] to 'D') database records where applicable
        ' CONSIDER WRAPPING IN A TRANSACTION BLOCK AND TRAPPING FOR ERRORS IN ORDER TO ROLLBACK, ETC.
        ' NOTE: Original ROWS will ALWAYS BE PRESENT WITH STATE FLAGS
        Try

            ' PROCESS ADDED RECORDS ---> INSERT INTO DATABASE TABLE
            For Each objRow As DataRow In objDT.Rows

                If objRow.RowState = DataRowState.Added Then
                    strID = ""
                    strPartnerID = objRow("PartnerID")
                    strSGeoID = objRow("ServiceGeoID")
                    strSPTCode = objRow("ServicePartnerTypeCode")
                    lngResult = HPQ.Excalibur.Service.InsertServiceFamilyPartnerDetails(Me.SFPN, strPartnerID, strSGeoID, "A", strSPTCode, strUserName, strReturnCode, strReturnDesc)
                    strID = lngResult.ToString()
                End If

            Next


            ' PROCESS DELETED RECORDS 
            If (Me.DeletedIDs.Length > 0) Then
                Dim aryIDs As String() = Split(Me.DeletedIDs, ",")
                Dim intIdx As Integer

                For intIdx = LBound(aryIDs) To UBound(aryIDs)
                    strID = aryIDs(intIdx)

                    If (IsNumeric(strID)) Then
                        If (Integer.Parse(strID) > 0) Then
                            '(MAKE IT A STATUS CHANGE RATHER THAN A DELETION)
                            lngResult = HPQ.Excalibur.Service.UpdateServiceFamilyPartnerDetails(strID, "", "", "", "D", "", strUserName, strReturnCode, strReturnDesc)
                        End If
                    End If

                Next
            End If

            blnResult = True
            setDirtyFlag()

        Catch ex As Exception
            blnResult = False

            DisplayMessage("Failed to Save Records to the database.  Error Description: " & ex.Message, "lblStatus", System.Drawing.Color.Red)
        End Try

        SaveToDatabase = blnResult

    End Function

    Protected Sub btnSaveSFP_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not Boolean.Parse(Me.SFPDirtyFlag.Value) Then
            Exit Sub
        End If

        Dim blnSaved As Boolean = SaveToDatabase()

        UpdateSFPSavedField(blnSaved)

        If blnSaved Then
            DisplayMessage("Successfully saved Service Family OSSP Assignment records.", "lblStatus", Drawing.Color.Black)
            Me.SFPDirtyFlag.Value = "false"
            Me.SFPSaved.Value = "null" ' Reset the flag for prompting to save the OSSP Assignment data changes

        Else
            If (Me.IsRPLMUser) Then
                'DisplayMessage("Failed to saved Service Family OSSP Assignment records.", "lblStatus", Drawing.Color.Red)
            End If
        End If

    End Sub

    Protected Sub btnCancelSFP_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If Not Boolean.Parse(Me.SFPDirtyFlag.Value) Or Not Boolean.Parse(Me.Continue.Value) Then
            Exit Sub
        End If

        ' Restore grid to it's initial state
        InitializeOSSPAssignGrid()

        DisplayMessage("Successfully reloaded original Service Family OSSP Assignment records.", "lblStatus", Drawing.Color.Black)

        Me.SFPDirtyFlag.Value = "false"
        Me.SFPSaved.Value = "null" ' Reset the flag for prompting to save the OSSP Assignment data changes

    End Sub

    Private Sub DisplayMessage(ByVal strMsg As String, ByVal strControlName As String, ByVal objColor As System.Drawing.Color)
        Dim objStatusLabel As Label = DirectCast(Me.FindControl(strControlName), Label)
        objStatusLabel.ForeColor = objColor
        objStatusLabel.Text = strMsg
    End Sub
    '*******************************************************************************************************************************************************************************************************************


</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Untitled Page</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="../_ScriptLibrary/jsrsClient.js"></script>
    <script type="text/javascript">
       
        var bRefreshCaller = false;
        var bPostBack = false;
        var bUnloadingVerified = false;
        var oDirtyValue = false;

        function ChooseEmployee(myControl) {

            var ResultArray;
            control = document.getElementById(myControl);

        ResultArray = window.showModalDialog("../MobileSE/Today/ChooseEmployee.asp", "", "dialogWidth:400px;dialogHeight:150px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")

            if (typeof (ResultArray) != "undefined") {
                if (ResultArray[0] != 0) {
                    control.options[control.length] = new Option(ResultArray[1], ResultArray[0]);
                    control.selectedIndex = control.length - 1;
                }
            }
        }

        var bWarned = false;

        function setReturnValue() {
            
            var oDirtyElement = document.getElementById("dirtyFlag");
                 oDirtyValue = oDirtyElement.getAttribute("value");
            var oSFPSavedElement = document.getElementById("SFPSaved");
            var sSFPValue=oSFPSavedElement.getAttribute("value").toString();
            var oSFPWarnedElement = document.getElementById("Warned");
            var sSFPWarned = oSFPWarnedElement.getAttribute("value").toString();

            //window.returnValue = oDirtyElement.getAttribute("value");

            if ((sSFPValue == "false") && (sSFPWarned == "false") && (!bPostBack))
            {
               // return "YOU HAVE NOT SAVED CHANGES TO THE Service Family OSSP Assignments data.\nClick 'Cancel' and then the 'Submit' button on the page to save your changes to the OSSP Assignments.";
                if (window.confirm("YOU HAVE NOT SAVED CHANGES TO THE Service Family OSSP Assignments data.\nClick 'Cancel' and then the 'Submit' button on the page to save your changes to the OSSP Assignments."))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            else
            {
                return true;
            }

        }

        function window_onload() {

            //add beforeclose to modalparent 
            parent.window.parent.$("#modal_dialog").dialog({
                beforeClose: function (ev, ui) {
                    if (setReturnValue() !== true)
                    {
                        return false;
                    }
                    else
                    {
                        parent.window.parent.SetServiceFamilyPn_return(oDirtyValue);
                        return true;
                    }
                }
            });
        }


        function confirmCancel() {
            var oSFPSavedElement = document.getElementById("SFPSaved");
            var sSFPValue = oSFPSavedElement.getAttribute("value").toString();

            if (sSFPValue != "null") {
                var oContinue = document.getElementById("Continue");

                oContinue.setAttribute("value", window.confirm("Undo changes?").toString());
            }
        }
    </script>

</head>
<body onload="window_onload();" style="align: center;">


    <form id="form1" runat="server">
        <div>
            <p><b>Service Family Details</b></p>
            <br />
            <asp:Label ID="lblStatusTop" runat="server" ForeColor="Red"></asp:Label><br />
            <asp:Label ID="lblServiceFamilyPn" runat="server" Text="Family SPS Pn"></asp:Label>
            <asp:Label ID="lblFamilyPn" runat="server" Visible="False"></asp:Label>
            <%If (Me.CanEditSFPN) Then%>
            <asp:TextBox ID="txtFamilyPn" runat="server">
            </asp:TextBox>
            <asp:Button ID="btnSaveFamilyPn" runat="server" OnClick="btnSaveFamilyPn_Click" Text="Save" OnClientClick="bPostBack=true;" />

            <asp:LinkButton ID="btnEditFamilyPn" runat="server" OnClick="btnEditFamilyPn_Click" OnClientClick="bPostBack=true;"
                Visible="False">Edit</asp:LinkButton>
            <%End If%>
            <br />
            <br />
            <asp:DetailsView ID="dvSpbDetails" runat="server" AutoGenerateRows="False" CssClass="FormTable"
                DataSourceID="odsSpbDetails" Visible="False" Width="95%" OnItemUpdated="dvSpbDetails_ItemUpdated"
                OnItemUpdating="dvSpbDetails_ItemUpdating">

                <Fields>
                    <asp:BoundField DataField="SvcMgrContact" HeaderText="Service Manager" ReadOnly="True"
                        InsertVisible="False" />
                    <asp:TemplateField HeaderText="GPLM">
                        <EditItemTemplate>
                            <asp:DropDownList ID="ddlGplmUser" runat="server" DataSourceID="odsGplmUsers" DataTextField="Name"
                                DataValueField="ID" SelectedValue='<%# Bind("GplmContactID") %>' AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <button onclick="ChooseEmployee('dvSpbDetails_ddlGplmUser');">
                                Add</button>
                        </EditItemTemplate>
                        <InsertItemTemplate>
                            <asp:DropDownList ID="ddlGplmUser" runat="server" DataSourceID="odsGplmUsers" DataTextField="Name"
                                DataValueField="ID" SelectedValue='<%# Bind("GplmContactID") %>' AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <button onclick="ChooseEmployee('dvSpbDetails_ddlGplmUser');">
                                Add</button>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="lblGplmContact" runat="server" Text='<%# Bind("GplmContact") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Bom Analyst">
                        <EditItemTemplate>
                            <asp:DropDownList ID="ddlSpdmUser" runat="server" DataSourceID="odsSpdmUsers" DataTextField="Name"
                                DataValueField="ID" SelectedValue='<%# Bind("SpdmContactID") %>' AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <button onclick="ChooseEmployee('dvSpbDetails_ddlSpdmUser');">
                                Add</button>
                        </EditItemTemplate>
                        <InsertItemTemplate>
                            <asp:DropDownList ID="ddlSpdmUser" runat="server" DataSourceID="odsSpdmUsers" DataTextField="Name"
                                DataValueField="ID" SelectedValue='<%# Bind("SpdmContactID") %>' AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value=""></asp:ListItem>
                            </asp:DropDownList>
                            <button onclick="ChooseEmployee('dvSpbDetails_ddlSpdmUser');">
                                Add</button>
                        </InsertItemTemplate>
                        <ItemTemplate>
                            <asp:Label ID="lblSpdmContact" runat="server" Text='<%# Bind("SpdmContact") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:CheckBoxField DataField="Active" HeaderText="SPB Auto Publish" />
                    <asp:CheckBoxField DataField="AutoPublishRsl" HeaderText="RSL Auto Publish" />
                    <asp:BoundField DataField="SeriesName" HeaderText="Series Name" ReadOnly="True" InsertVisible="False" />


                    <asp:TemplateField HeaderText="Business Unit">
                        <EditItemTemplate>
                            <asp:RadioButton ID="commercial" Text="Commercial" GroupName="busunit" runat="server" Checked='<%# IIf((Eval("BusinessUnit") = 2) Or (Eval("BusinessUnit") = 0 And Eval("BusinessId") = 2), False, True)%>' OnCheckedChanged="OnSelectedIndexChangedMethod" />
                            <asp:RadioButton ID="consumer" Text="Consumer" GroupName="busunit" runat="server" Checked='<%# IIf((Eval("BusinessUnit")=1) or (Eval("BusinessUnit")=0 and Eval("BusinessId")=1) ,false,true) %>' OnCheckedChanged="OnSelectedIndexChangedMethod" />
                            <asp:TextBox ID="txtBusinessUnit" runat="server"
                                Text='<%# Bind("BusinessUnit") %>' Visible="false"></asp:TextBox>
                        </EditItemTemplate>
                        <ItemTemplate>
                            <asp:RadioButton ID="commercial" Text="Commercial" Enabled="false" GroupName="busunit" runat="server" Checked='<%# IIf((Eval("BusinessUnit") = 2) Or (Eval("BusinessUnit") = 0 And Eval("BusinessId") = 2), False, True)%>' OnCheckedChanged="OnSelectedIndexChangedMethod" />
                            <asp:RadioButton ID="consumer" Text="Consumer" Enabled="false" GroupName="busunit" runat="server" Checked='<%# IIf((Eval("BusinessUnit") = 1) or (Eval("BusinessUnit") = 0 and Eval("BusinessId") = 1), False, True)%>' OnCheckedChanged="OnSelectedIndexChangedMethod" />
                        </ItemTemplate>

                    </asp:TemplateField>

                    <asp:BoundField DataField="ProjectCd" HeaderText="Project Code" ReadOnly="True" InsertVisible="False" />
                    <asp:BoundField DataField="SelfRepairDoc" HeaderText="Self Repair Doc" />
                    <asp:BoundField DataField="PartnerName" HeaderText="Partner (ODM)" ReadOnly="True"
                        InsertVisible="False" />
                    <asp:BoundField DataField="SharePointPath" HeaderText="Share Point Path" Visible="false">
                        <ControlStyle Width="95%" />
                    </asp:BoundField>
                    <asp:BoundField DataField="SharedDrivePath" HeaderText="Shared Drive Path" Visible="false">
                        <ControlStyle Width="95%" />
                    </asp:BoundField>
                    <asp:BoundField DataField="SpdmContactID" HeaderText="SpdmContactID" Visible="False" />

                </Fields>
            </asp:DetailsView>


            <br />
            <hr />
            <p><b>OSSP Partner Assignments</b><span style="color: Red; font-weight: bold;">&nbsp;(Beta)</span></p>
            <asp:Label ID="lblStatus" runat="server" ForeColor="Red"></asp:Label>
            <br />
            <asp:DataGrid ID="grdSFPartners" runat="server" ShowFooter="True"
                AutoGenerateColumns="False" CssClass="FormTable" Width="95%" ItemCommand="grdSFPartners_ItemCommand">

                <Columns>

                    <asp:TemplateColumn HeaderText="OSSP">
                        <FooterTemplate>
                            <%If (Me.CanEditOSSPDetails) Then%>
                            <asp:DropDownList ID="add_ddlOSSP" runat="server" DataSourceID="odsListPartners"
                                DataValueField="ID" DataTextField="Name" AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value="" Selected="True"></asp:ListItem>
                            </asp:DropDownList>
                            <%End If%>
                        </FooterTemplate>
                        <ItemTemplate>
                            <%#Container.DataItem("Name")%>
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn HeaderText="Geo">
                        <FooterTemplate>
                            <%If (Me.CanEditOSSPDetails) Then%>
                            <asp:DropDownList ID="add_ddlServiceGeo" runat="server" DataSourceID="odsListServiceGeos"
                                DataValueField="ServiceGeoID" DataTextField="ServiceGeoShortName" AppendDataBoundItems="True">
                                <asp:ListItem Text="-- Select One --" Value="" Selected="True"></asp:ListItem>
                            </asp:DropDownList>
                            <%End If%>
                        </FooterTemplate>
                        <ItemTemplate>
                            <%#Container.DataItem("ServiceGeoShortName")%>
                        </ItemTemplate>
                    </asp:TemplateColumn>


                    <asp:TemplateColumn HeaderText="Type">
                        <FooterTemplate>
                            <%If (Me.CanEditOSSPDetails) Then%>
                            <asp:DropDownList ID="add_ddlSPTC" runat="server">
                                <asp:ListItem Text="-- Select One --" Value="" Selected="True"></asp:ListItem>
                                <asp:ListItem Text="WUR" Value="WUR"></asp:ListItem>
                                <asp:ListItem Text="PFH" Value="PFH"></asp:ListItem>
                            </asp:DropDownList>
                            <%End If%>
                        </FooterTemplate>
                        <ItemTemplate>
                            <%#Container.DataItem("ServicePartnerTypeCode")%>
                        </ItemTemplate>
                    </asp:TemplateColumn>


                    <asp:TemplateColumn>
                        <FooterTemplate>
                            <%If (Me.CanEditOSSPDetails) Then%>
                            <asp:LinkButton runat="server" CommandName="Insert" Text="Add" ID="btnAddPartnerRecord" OnClientClick="javascript:bPostBack=true;" />
                            <%End If%>
                        </FooterTemplate>
                        <ItemTemplate>
                            <%If (Me.CanEditOSSPDetails) Then%>
                            <asp:LinkButton ID="btnDeletePartnerRecord" runat="server" CommandName="Delete" Text="Delete" RecordID="<%#Container.DataItem(0)%>" OnClientClick="javascript:bPostBack=true;" />
                            <%End If%>
                        </ItemTemplate>
                    </asp:TemplateColumn>

                </Columns>

                <HeaderStyle CssClass="TableHeader" />
            </asp:DataGrid>

            <%If (Me.CanEditOSSPDetails) Then%>
            <div>
                <p style="text-align: left">
                    <asp:Button ID="btnSaveSFP" runat="server" Title="Save Changes" Text="Submit" OnClick="btnSaveSFP_Click" OnClientClick="bPostBack=true;" />
                    &nbsp;&nbsp;
                    <asp:Button ID="btnCancelSFP" runat="server" Title="Cancel Changes and Restore Original Records" Text="Cancel"
                        OnClientClick="bPostBack=true;confirmCancel();" OnClick="btnCancelSFP_Click" />
                </p>
            </div>
            <%End If%>
            <hr />


            <asp:ObjectDataSource ID="odsSpbDetails" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectSpbDetails" TypeName="HPQ.Excalibur.Data" UpdateMethod="UpdateServiceFamilyDetails"
                InsertMethod="UpdateServiceFamilyDetails">
                <UpdateParameters>
                    <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text"
                        Type="String" />
                    <asp:Parameter Name="SpdmContactID" Type="String" />
                    <asp:Parameter Name="GplmContactID" Type="String" />
                    <asp:Parameter Name="Active" Type="Boolean" />
                    <asp:Parameter Name="SharePointPath" Type="String" />
                    <asp:Parameter Name="SharedDrivePath" Type="String" />
                    <asp:Parameter Name="SelfRepairDoc" Type="String" />
                    <asp:Parameter Name="AutoPublishRsl" Type="Boolean" />
                    <asp:Parameter Name="BusinessUnit" Type="String" />
                </UpdateParameters>
                <InsertParameters>
                    <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text"
                        Type="String" />
                    <asp:Parameter Name="SpdmContactID" Type="String" />
                    <asp:Parameter Name="GplmContactID" Type="String" />
                    <asp:Parameter Name="Active" Type="Boolean" />
                    <asp:Parameter Name="SharePointPath" Type="String" />
                    <asp:Parameter Name="SharedDrivePath" Type="String" />
                    <asp:Parameter Name="SelfRepairDoc" Type="String" />
                    <asp:Parameter Name="AutoPublishRsl" Type="Boolean" />
                    <asp:Parameter Name="BusinessUnit" Type="String" />
                </InsertParameters>
                <SelectParameters>
                    <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text"
                        Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <asp:ObjectDataSource ID="odsSpdmUsers" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="ListSpdms" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionId" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <asp:ObjectDataSource ID="odsGplmUsers" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="ListGplms" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionId" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <asp:ObjectDataSource ID="odsSvcManagers" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="ListSvcManagers" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionId" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <!--*********************************************************************************************************-->
            <!-- ADD LOGIC BELOW TO IMPLEMENT THE MAINTENANCE OF THE [ProductVer_Partner] TABLE                              -->
            <!--*********************************************************************************************************-->
            <asp:ObjectDataSource ID="odsServiceFamilyPartnerDetails" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectServiceFamilyPartnerDetails"
                TypeName="HPQ.Excalibur.Service">
                <SelectParameters>
                    <asp:Parameter Name="ID" Type="String" />
                    <asp:Parameter DefaultValue="491039-001" Name="ServiceFamilyPn" Type="String" />
                    <asp:Parameter Name="PartnerID" Type="String" />
                    <asp:Parameter Name="ServiceGeoID" Type="String" />
                    <asp:Parameter Name="Status" Type="String" />
                    <asp:Parameter Name="ServicePartnerTypeCode" Type="String" />
                    <asp:Parameter Direction="Output" Name="ReturnCd" Type="String" />
                    <asp:Parameter Direction="Output" Name="ReturnDesc" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <asp:ObjectDataSource ID="odsListPartners" runat="server"
                OldValuesParameterFormatString="original_{0}" SelectMethod="ListPartners"
                TypeName="HPQ.Excalibur.Data" FilterExpression="active=1">
                <SelectParameters>
                    <asp:Parameter DefaultValue="1" Name="ReportType" Type="String" />
                    <asp:Parameter DefaultValue="2" Name="PartnerTypeID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>

            <asp:ObjectDataSource ID="odsListServiceGeos" runat="server"
                OldValuesParameterFormatString="original_{0}" SelectMethod="ListServiceGeos"
                TypeName="HPQ.Excalibur.Service"></asp:ObjectDataSource>
            <!--*********************************************************************************************************-->

        </div>

        <asp:HiddenField ID="dirtyFlag" runat="server" Value="false" />
        <asp:HiddenField ID="SFPDirtyFlag" runat="server" Value="false" />
        <asp:HiddenField ID="SFPSaved" runat="server" Value="null" />
        <asp:HiddenField ID="Warned" runat="server" Value="false" />
        <asp:HiddenField ID="Continue" runat="server" Value="false" />
    </form>

</body>
</html>
