<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/EmailWrapper.asp" --> 
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<%

Response.AddHeader "Pragma", "No-Cache"

Dim m_IsSysAdmin
Dim m_IsDeliverableOwner
Dim m_IsPlatformDevMgr
Dim m_EditModeOn
Dim m_FormSave
Dim m_FormClose
Dim m_FormDisplay
Dim m_CurrentUserEmail
Dim m_ProductID
Dim m_DcrID
Dim m_IsSystemTeamLead
Dim m_IsHWPM
Dim m_ReleaseCertification
Dim m_ReleaseName
Dim m_ReleaseID
Dim m_CountryID
Dim m_CurrentData
Dim m_CertType
Dim m_CertData
Dim m_ReleaseBatchUpdate
Dim m_ReleaseRecordSet
Dim m_ReleaseCount
Dim m_from_where
Dim m_Platform
Dim m_CountryName

m_FormSave = Trim(Request.Form("hidSave"))
If Len(Trim(m_FormSave)) = 0 Then
	m_FormSave = False
End If

m_FormClose = Trim(Request.Form("hidClose"))
If Len(Trim(m_FormClose)) = 0 Then
	m_FormClose = False
End If

m_IsSysAdmin = False
m_IsDeliverableOwner = False
m_EditModeOn = False
m_FormDisplay = False
m_IsSystemTeamLead = False
m_IsPlatformDevMgr = False
m_IsHWPM = False
m_ReleaseCertification = False
m_ReleaseID = 0
m_CountryID = 0
m_CertType = 0
m_CertData = ""
m_ReleaseBatchUpdate = False
m_from_where = ""
m_Platform = ""
m_CountryName = ""

Sub Main()
	If m_FormSave Then
		Call SaveData()
	Else
		Call DisplayData()
	End If
End Sub

Sub SaveLANData()
    'Can this user edit?
	Trim(Request("deliverable_root_id"))
	
    Dim Security, sUserFullName
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsDeliverableOwner = Security.IsDeliverableOwner("",Trim(Request("deliverable_root_id")))
	m_IsSystemTeamLead = Security.IsSystemTeamLead(Request.Form("product_version_id"))
	m_IsPlatformDevMgr = Security.IsPlatformDevMgr(Request.Form("product_version_id"))
	m_IsHwPM = Security.IsHardwarePm(Request.Form("product_version_id"))
	sUserFullName = Security.CurrentUser()
	
	If m_IsSysAdmin Or m_IsDeliverableOwner Or m_IsSystemTeamLead Or m_IsPlatformDevMgr Or m_IsHWPM Then
		m_EditModeOn = True
	End If
	
	m_CurrentUserEmail = Security.CurrentUserEmail()
	
	Set Security = Nothing

	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3><H4>Unable to save data changes</H4>"
		m_FormDisplay = False
		m_FormClose = False
	Else
	
		Dim dw
		Dim cn
		Dim cmd
		Dim RecordCount

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAgencyStatus")

		dw.CreateParameter cmd, "@p_AgencyStatusID", adInteger, adParamInput, 8, Trim(Request("agency_status_id"))
        dw.CreateParameter cmd, "@p_SelectedProducts", adVarChar, adParamInput, 5000, ""
        dw.CreateParameter cmd, "@p_SelectedCountries", adVarChar, adParamInput, 5000, ""
        dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, ""
		dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(sUserFullName)
		dw.CreateParameter cmd, "@p_StatusCd", adChar, adParamInput, 5, Trim(Request("cboStatus"))
		dw.CreateParameter cmd, "@p_ProjectedDate", adDate, adParamInput, 8, Trim(Request("txtProjectedDate"))
		dw.CreateParameter cmd, "@p_ActualDate", adDate, adParamInput, 8, Trim(Request("txtActualDate"))
		dw.CreateParameter cmd, "@p_CertificationNo", adVarChar, adParamInput, 50, Trim(Request("txtCertificationNo"))
		dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, Trim(Request("hidLeveragedID"))
		dw.CreateParameter cmd, "@p_Notes", adVarChar, adParamInput, 5000, Trim(Request("txtNotes"))
		dw.CreateParameter cmd, "@p_TestOrganizer", adInteger, adParamInput, 8, Trim(Request("txtTestOrganizer"))
		dw.CreateParameter cmd, "@p_TestBudget", adInteger, adParamInput, 8, Trim(Request("txtTestBudget"))
		dw.CreateParameter cmd, "@p_POR_DCR", adChar, adParamInput, 3, Trim(Request("cboPorDcr"))
		dw.CreateParameter cmd, "@p_Dcr_Id", adInteger, adParamInput, 8, Trim(Request("cboDcr"))
        dw.CreateParameter cmd, "@p_ModifiedBy", adVarChar, adParamInput, 15, Trim(Request("hidfromWhere"))

        RecordCount = dw.ExecuteNonQuery(cmd)
	
		Set cmd = Nothing
	
		If RecordCount = 0 Then
			Response.Write "<H3>Error Updating Record</H3>"
			m_FormDisplay = False
			m_FormClose = False
			Exit Sub
		End If

		If Request.Form("txtProjectedDate") <> Request.Form("projected_date") And Application("SendAgencyEmail") Then
			Dim sRecipient, sSubject, sBody

			Dim oMessage
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")		


			'sRecipient = "kenneth.berntsen@hp.com"
			sRecipient = GetEmailNotificationList(Request.Form("agency_status_id")) ' Get a list of names from each program that is using this deliverable.
			sSubject = "The Agency availability date for " & Request.Form("deliverable_name") & " has changed."
			sBody = "The availability date for " & Request.Form("deliverable_name") & " has been moved from " & _
				Request.Form("projected_date") & " to " & Request.Form("txtProjectedDate") & "." & vbcrlf ' & GetEmailNotificationList(Request.Form("agency_status_id"))
			
			oMessage.To = sRecipient
			oMessage.From = m_CurrentUserEmail
			oMessage.Subject = sSubject
			oMessage.TextBody = sBody
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send
			
			
			Set oMessage = Nothing

		End If
	End If 'Not m_EditModeOn

	If Not m_FormClose Then
		m_FormDisplay = True
	End If
End Sub

Sub SaveReleaseData()
    Dim dw
    Dim cn
	Dim cmd
	Dim RecordCount
    Dim InputCountry
    Dim InputRelease

    If Trim(Request.Form("hidBatchUpdate")) = "True" Then
       InputCountry = Trim(Request.Form("hidBatchUpdateCountry"))
       InputRelease = Trim(Request.Form("hidBatchUpdateRelease"))
    Else
       InputCountry = Trim(Request.Form("hidCountryID"))
       InputRelease = Trim(Request.Form("hidReleaseID"))
    End If

    Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAngecyReleaseStatus")

    dw.CreateParameter cmd, "@ProductVersionID", adInteger, adParamInput, 8, Trim(Request.Form("hidProductID"))
    dw.CreateParameter cmd, "@CountryID", adVarChar, adParamInput, 16, InputCountry
    dw.CreateParameter cmd, "@ReleaseID", adVarChar, adParamInput, 16, InputRelease
    dw.CreateParameter cmd, "@Type", adInteger, adParamInput, 8, Trim(Request.Form("cboStatus"))
    dw.CreateParameter cmd, "@Date", adVarChar, adParamInput, 16, Trim(Request("txtProjectedDate"))
    RecordCount = dw.ExecuteNonQuery(cmd)

    Set cmd = Nothing

    If RecordCount = 0 Then
		Response.Write "<H3>Error Updating Record</H3>"
		m_FormDisplay = False
		m_FormClose = False
		Exit Sub        
	End If

End Sub

Sub SaveData()
    If Request.Form("hidBatchUpdate") = "True" Or (Request.Form("hidReleaseID") > 0 And Request.Form("hidCountryID") > 0) Then
       call SaveReleaseData()
    Else 
       call SaveLANData()
    End If
End Sub

Function GetEmailNotificationList(AgencyStatusID)
	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	Dim sRecipients
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyProjectDistribution")
	dw.CreateParameter cmd, "@p_AgencyStatusID", adInteger, adParamInput, 8, Trim(AgencyStatusID)
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		sRecipients = sRecipients & ";" & rs("distribution")
		rs.movenext
	Loop

	sRecipients = Replace(sRecipients, ";;", ";")

	GetEmailNotificationList = sRecipients
End Function

Sub DisplayWLANData()
    Dim dw
	Dim cn
	Dim cmd
	Dim rsStatus

    'Get Status Information
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAgencyStatus")
	dw.CreateParameter cmd, "@p_StatusID", adInteger, adParamInput, 8, Request("ID")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_DeliverableVersionID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_DeliverableCategoryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_MappingID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_StatusCd", adVarChar, adParamInput, 10, ""
	dw.CreateParameter cmd, "@p_LeveragedID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AgencyID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_CountryID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_Region", adVarChar, adParamInput, 15, ""
    dw.CreateParameter cmd, "@p_Region", adVarChar, adParamInput, 15, m_from_where 

	Set rsStatus = dw.ExecuteCommandReturnRS(cmd)	
	
	If rsStatus.EOF and rsStatus.BOF Then
		Response.Write "<H3>No Data Returned for StatusID = " & Request("ID") & "</H3>"
		Response.End
	End If
	
	m_ProductID = rsStatus("product_version_id")
	m_DcrID = rsStatus("dcr_id")

	'Can this user edit?
	Dim sDeliverableRootID
	sDeliverableRootID = Trim(rsStatus("deliverable_root_id"))
	
	Dim Security
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsSystemTeamLead = Security.IsSystemTeamLead(m_ProductID)
	m_IsDeliverableOwner = Security.IsDeliverableOwner("", sDeliverableRootID)
    m_IsPlatformDevMgr = Security.IsPlatformDevMgr(m_ProductID)
    
	If m_IsSysAdmin Or m_IsDeliverableOwner Or m_IsSystemTeamLead Or m_IsPlatformDevMgr Then
		m_EditModeOn = True
	End If
	
	Set Security = Nothing
	
	Dim field
	For Each field in rsStatus.Fields
		Response.Write "<input type=""hidden"" id=""" & Trim(field.name) & """ name=""" & Trim(field.name) & """ value=""" & Replace(Trim(field.value)&"", vbCr, "<br />") & """ >" & vbcrlf
	Next

	'Response.Write "<input type=""hidden"" id=""EditModeOn"" name=""EditModeOn"" value=""" & m_EditModeOn & """ >" & vbcrlf
	
	m_FormDisplay = True
End Sub

Sub DisplayDataForRelease()
    m_FormDisplay = True

    m_ProductID = Trim(Request("ProductID"))
    m_ReleaseID = Trim(Request("ReleaseID"))
    m_CountryID = Trim(Request("CountryID"))
    m_ReleaseCertification = True
    m_ReleaseName = Trim(Request("ReleaseName"))
    m_CurrentData = Trim(Request("currentData"))

    Dim Security
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	m_IsSystemTeamLead = Security.IsSystemTeamLead(m_ProductID)
    m_IsPlatformDevMgr = Security.IsPlatformDevMgr(m_ProductID)
    
	If m_IsSysAdmin Or m_IsSystemTeamLead Or m_IsPlatformDevMgr Then
		m_EditModeOn = True
	End If
	
	Set Security = Nothing

    If m_ReleaseID > 0 And m_CountryID > 0 Then
        If m_CurrentData <> "" And InStr(m_CurrentData & "","-") > 0 Then
            dim args 
            args = Split(Trim(m_CurrentData),"-")
            m_CertType = args(0)
            m_CertData = args(1)
        End If
    Else   
        m_ReleaseBatchUpdate = True

        Dim cn
	    Dim cmd
	    Dim dw
        Dim rs
    
        Set rs = Server.CreateObject("ADODB.RecordSet")
	    Set dw = New DataWrapper
	    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	    Set cmd = dw.CreateCommandSP(cn, "rpt_AgencyPMView")
	    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, m_ProductID
	    Set m_ReleaseRecordSet = dw.ExecuteCommandReturnRS(cmd)

        Set rs = m_ReleaseRecordSet.NextRecordset
        m_ReleaseCount = rs.Fields(1).value
        m_Platform = rs.Fields(2).value
        rs.Close
    End If
End Sub

Sub DisplayData()

    m_from_where = Trim(Request("from_where"))
    m_Platform = Trim(Request("Platform"))
    m_CountryName = Trim(Request("CountryName"))

    If Len(Trim(Request("ID"))) > 0 Then
        call DisplayWLANData()
    ElseIf Len(Trim(Request("CountryID"))) > 0 And Len(Trim(Request("ReleaseID"))) > 0 Then
        call DisplayDataForRelease()
    Else
        Response.Write "<H3>No StatusID Provided Unable to Process Your Request</H3>"
	    Response.End
    End If

End Sub

Sub FillDcrStatus(ProductID)
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spListApprovedDCRs")
	dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 8, Trim(ProductID)
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Response.Write m_dcrid
	
	Do until rs.eof
		if trim(m_DcrId) = trim(rs("ID")) then
			Response.Write "<option selected value=""" & rs("ID") & """>" & rs("ID") & ":" & server.HTMLEncode(rs("Summary")) & "</option>"					
		else
			Response.Write "<option value=""" & rs("ID") & """>" & rs("ID") & ":" & server.HTMLEncode(rs("Summary")) & "</option>"					
		end if
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

%>

<HTML>
<HEAD>
<title>Agency Main</title>
<!-- #include file="../includes/bundleConfig.inc" -->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    var cboStatus_LastIndex;

    function cmdDate_onclick(FieldID) {
        var strID;
        var oldValue = window.frmStatus.elements(FieldID).value;

        strID = window.showModalDialog("../mobilese/today/caldraw1.asp", FieldID, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) == "undefined")
            return

        window.frmStatus.elements(FieldID).value = strID;

        var newDate = Date.parse(strID);
        var oldDate = Date.parse(oldValue);

        var dcrID;
        var programID = window.frmStatus.product_version_id.value;
        var statusID = window.frmStatus.agency_status_id.value;

        if (newDate > oldDate) {
            dcrID = window.showModalDialog("ChooseDCR.asp?ID=" + programID + "&StatusID=" + statusID, FieldID, "dialogWidth:700px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        }
    }

    function ChooseDCR() {
        $("#txtProjectedDate").change(function () {
            var dcrID;
            var strID = $("#txtProjectedDate").val();
            var oldValue = $("#inpProjectedDate").val();

            if (oldValue == "") {
                $("#inpProjectedDate").val(strID);
            }

            if (oldValue != "") {
                var newDate = Date.parse(strID);
                var oldDate = Date.parse(oldValue);

                var programID = window.frmStatus.product_version_id.value;
                var statusID = window.frmStatus.agency_status_id.value;

                if (newDate > oldDate) {
                    modalDialog.open({ dialogTitle: 'Choose DCR', dialogURL: 'ChooseDCR.asp?ID=' + programID + '&StatusID=' + statusID + '', dialogHeight: 200, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
                    /*dcrID = window.showModalDialog("ChooseDCR.asp?ID=" + programID + "&StatusID=" + statusID, FieldID, "dialogWidth:700px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No");*/
                }

                $("#inpProjectedDate").val(strID);
            }
        });
    }

    function Left(str, n) {
        if (n <= 0)     // Invalid bound, return blank string
            return "";
        else if (n > String(str).length)   // Invalid bound, return
            return str;                // entire string
        else // Valid bound, return appropriate substring
            return String(str).substring(0, n);
    }

    function window_onload(pulsarplusDivId) {
        if (window.frmStatus.hidClose) {
            if (window.frmStatus.hidClose.value.toLowerCase() == 'true') {
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    parent.window.parent.certificationPageCallBack();
                    //parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                    // For Closing current popup
                    parent.window.parent.closeExternalPopup();
                }
                else {

                    var iframeName = parent.window.name;

                    if (iframeName != '') {
                        if (parent.window.parent.document.getElementById('modal_dialog')) {
                            parent.window.parent.modalDialog.cancel();
                        } else {
                            parent.window.parent.CloseIframeDialog();
                        }
                    } else {
                        this.close();
                    }
                }
            }
        }

        if (window.frmStatus.hidDisplay) {
            if (window.frmStatus.hidDisplay.value.toLowerCase() == 'true') { 
                populateForm();
                cboStatus_onchange();
                window.DisplayForm.style.display = "";
            }
        }

        if (typeof (window.parent.frames["LowerWindow"].frmButtons) == 'object') {
            if (window.frmStatus.hidEdit.value.toLowerCase() == 'false')
                window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = true;
        }

        //Instantiate modalDialog load
        modalDialog.load();

        //add datepicker
        load_datePicker();

        //Choose DCR
        ChooseDCR();
    }

    function populateForm() {
        if (window.frmStatus.agency_status_id) {
            window.lblPlatform.innerHTML = window.frmStatus.product_name.value;
            window.lblCountry.innerHTML = window.frmStatus.country_name.value;
            window.lblCertification.innerHTML = window.frmStatus.agency_name.value;
            setComboBoxValue(window.frmStatus.cboStatus, window.frmStatus.status_cd.value);
            //cboStatus_LastIndex = window.frmStatus.cboStatus.selectedIndex;
            window.frmStatus.txtProjectedDate.value = window.frmStatus.projected_date.value;
            window.frmStatus.txtTestOrganizer.value = '';
            window.frmStatus.txtTestBudget.value = '';
            window.frmStatus.txtLeveragedSystem.value = window.frmStatus.leveraged_name.value;
            window.frmStatus.hidLeveragedID.value = window.frmStatus.leveraged_id.value;
            window.frmStatus.txtNotes.value = window.frmStatus.status_notes.value;
            window.lblTitle.innerHTML = window.frmStatus.deliverable_name.value;
            window.lblNotices.innerHTML = window.frmStatus.mapping_notes.value;
            if (window.frmStatus.mapping_notes.value == '')
                window.MapingNotice.style.display = 'none';
            setComboBoxValue(window.frmStatus.cboPorDcr, window.frmStatus.por_dcr.value);
            cboPorDcr_onchange();
        }
    }

    function setComboBoxValue(object, value) {
        if (value == "O")
            value = "SU";
        object.value = value;
        if ((value == 'NS') && (window.frmStatus.supported_country_yn.value == 'N')) {
            window.frmStatus.cboStatus.disabled = true;
        }
    }

    function cboStatus_onchange() {

        //	if ((window.frmStatus.cboStatus.value == 'NS') && (typeof(cboStatus_LastIndex) != 'undefined'))
        //		window.frmStatus.cboStatus.selectedIndex = cboStatus_LastIndex;

        if (window.AvailDate != undefined) {
            window.AvailDate.style.display = "none";
        }
        if (window.BudgetRow != undefined) {
            window.BudgetRow.style.display = "none";
        }
        if (window.OrganizerRow != undefined) {
            window.OrganizerRow.style.display = "none";
        }
        if (window.LeveragedRow != undefined) {
            window.LeveragedRow.style.display = "none";
        }
        if (window.CertNo != undefined) {
            window.CertNo.style.display = "none";
        }
        if (window.ActualDate != undefined) {
            window.ActualDate.style.display = "none";
        }
        if (window.DcrPorSelectRow != undefined) {
            window.DcrPorSelectRow.style.display = "";
        }


        switch (window.frmStatus.cboStatus.value) {
            case 'L':
                window.LeveragedRow.style.display = "";
                window.frmStatus.txtProjectedDate.value = "";
                break;
            case 'P':
                window.AvailDate.style.display = "";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'O':
                window.AvailDate.style.display = "";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'SU':
                window.AvailDate.style.display = "";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'C':
                window.AvailDate.style.display = "";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'NS':
                if (window.DcrPorSelectRow != undefined) {
                    window.DcrPorSelectRow.style.display = "none";
                }
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'NR':
                window.DcrPorSelectRow.style.display = "none";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case 'NC':
                window.DcrPorSelectRow.style.display = "none";
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                break;
            case '0':
                if (window.AvailDate != undefined) {
                    window.AvailDate.style.display = "none";
                    $("#txtProjectedDate").val('');
                }
                break;
            case '1':
                if (window.AvailDate != undefined) {
                    window.AvailDate.style.display = "none";
                    //$("#txtProjectedDate").val('');
                }
                break;
            case '2':
                if (window.AvailDate != undefined) {
                    window.AvailDate.style.display = "";
                }
                break;
            default:
                if (window.frmStatus.hidLeveragedID != undefined) {
                    window.frmStatus.hidLeveragedID.value = "";
                }
                if (window.frmStatus.txtLeveragedSystem != undefined) {
                    window.frmStatus.txtLeveragedSystem.value = "";
                }
                window.frmStatus.txtProjectedDate.value = "";
                break;
        }
        cboStatus_LastIndex = window.frmStatus.cboStatus.selectedIndex;


    }

    function cboPorDcr_onchange() {
        if (window.frmStatus.cboPorDcr.value == 'DCR')
            window.DcrSelectRow.style.display = "";
        else
            window.DcrSelectRow.style.display = "none";
    }

    function LeverageSearch_onClick() {
        var RetVal;
        var DRID;
        var PVID;
        var CID;

        DRID = window.frmStatus.deliverable_root_id.value;
        PVID = window.frmStatus.product_version_id.value;
        CID = window.frmStatus.country_id.value;
        PULSARTABID = document.getElementById('hdnPulsarTabName').value;
        //alert("AgencyLeverage.asp?DRID="+DRID+"&PVID="+PVID+"&CID="+CID);
        modalDialog.open({ dialogTitle: 'Leverage Search', dialogURL: 'AgencyLeverage.asp?DRID=' + DRID + '&PVID=' + PVID + '&CID=' + CID + '' + '&pulsarplusDivId=' + PULSARTABID + '&AgencyPage=AgencyLeveraged', dialogHeight: 400, dialogWidth: 500, dialogResizable: true, dialogDraggable: true });
        //RetVal = window.showModalDialog("AgencyLeverage.asp?DRID="+DRID+"&PVID="+PVID+"&CID="+CID, window.self, "dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;maximize:No;resizable:No;status: Yes"); 
    }

    function LeverageSearchResult(RetVal) {
        if (typeof (RetVal) != "undefined") {
            window.frmStatus.hidLeveragedID.value = RetVal.split("|")[0];
            window.frmStatus.txtLeveragedSystem.value = RetVal.split("|")[1];
        }
    }
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<form id="frmStatus" method="post" action=AgencyMain.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>><% Call Main() %>
<div id=DisplayForm style="display:none">
<font face=verdana size=2>
    <span style="font:bold x-small verdana">
        <% If m_ReleaseCertification = False Then %>
            <label ID="lblTitle"></label>&nbsp;Agency Status
        <% ElseIf m_ReleaseBatchUpdate = True Then %>
            <label ID="lblTitle"></label>&nbsp;Batch Update
        <% Else %>
            <label ID="lblTitle"></label>&nbsp;<%=m_ReleaseName %>
        <% End If %>
    </span>
</font>
<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Platform:</span>&nbsp;</td>
		<td>
			<label style="font:x-small verdana" id="lblPlatform" name="lblPlatform">
                <%= m_Platform %>
			</label>
		</td>
	</tr>

    <% If m_ReleaseBatchUpdate = True Then
          Dim field
    %>
     <tr>
         <td valign=top width=140 nowrap><span style="font:bold x-small verdana">Select Batch Update Release(s):</span>&nbsp;</td>
         <td>
            <select id="release_select" name="release_select" style="height: 150px;" multiple>  
            <% For field = 1 To m_ReleaseCount  %>
                <option value="<%=m_ReleaseRecordSet.Fields(field).name %>"><%=m_ReleaseRecordSet.Fields(field).value %></option>
            <% Next %>
            </select>
         </td>
     </tr>
     <tr>
         <td valign=top width=140 nowrap><span style="font:bold x-small verdana">Select Batch Update Country(s):</span>&nbsp;</td>
         <td>
             <select id="country_select" name="country_select" style="height: 150px;" multiple>
             <%
                Set m_ReleaseRecordSet = m_ReleaseRecordSet.NextRecordset  
                Do Until m_ReleaseRecordSet.EOF
                   IF m_ReleaseRecordSet.Fields(3).Value & "" <> "" Then
             %>
                    <option value="<%=m_ReleaseRecordSet.Fields(0).Value %>"><%=m_ReleaseRecordSet.Fields(1).Value %></option>
             <% 
                   End If 
                m_ReleaseRecordSet.MoveNext
                Loop 
                 
                m_ReleaseRecordSet.Close
             %>
             </select>
         </td>
     </tr>
    <% End If %>

    <% IF m_ReleaseBatchUpdate = False Then %>
	<tr>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Country:</span>&nbsp;</td>
		<td>
			<label style="font:x-small verdana" id="lblCountry" name="lblCountry">
                <%= m_CountryName %>
			</label>
		</td>
	</tr>
    <% END IF %>   

	<tr>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Certification:</span>&nbsp;</td>
		<td>
			<label style="font:x-small verdana" id="lblCertification" name="lblCertification">Country Specific</label>
		</td>
	</tr>
    
    <% If m_ReleaseCertification = False Then %>
	<tr ID=MapingNotice>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Requirements:</span>&nbsp;</td>
		<td>
			<label style="font:x-small verdana" id="lblNotices" name="lblNotices">Country Specific</label>
		</td>
	</tr>
    <% End If %>

	<tr>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Status:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<SELECT id="cboStatus" name="cboStatus" language="javascript" onchange="cboStatus_onchange();" >
            <% If m_ReleaseCertification = False Then %>
                <option value=SU>Supported</option>
			    <option value=P>Partial</option>
			    <option value=C>Complete</option>
			    <option value=L>Leveraged</option>
                <option value=NS>Not Supported</option>
                <option value=NR>Not Requested</option>
			    <option value=NC>No Cert Needed</option>
            <% Else %>
                <option value=0>Not Supported</option>
                <option value=1>Supported</option>
                <option value=2>Date</option>
            <% End If %>
			</SELECT>
		</td>
	</tr>
	<tr ID=AvailDate>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Availability Date:</span></td>
		<td>
			<INPUT type="text" id=txtProjectedDate name=txtProjectedDate value="" class="dateselection">
            <input type="hidden" id="inpProjectedDate" value="" />
		</td>
	</tr>

    <% If m_ReleaseCertification = False Then %>
	<tr ID=ActualDate>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Actual Date:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtActualDate name=txtActualDate value="" readonly>
	</td>
	</tr>
	<tr ID=LeveragedRow>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Leveraged:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="hidden" id=hidLeveragedID name=hidLeveragedID value="">
			<INPUT type="text" id=txtLeveragedSystem name=txtLeveragedSystem value="" readonly>&nbsp;<a href="javascript: LeverageSearch_onClick()"><img ID="picSearch" SRC="../images/search.gif" alt="Find Platform" border="0" WIDTH="20" HEIGHT="16"></a>
		</td>
	</tr>
	<tr ID=OrganizerRow>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Test Organizer:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtTestOrganizer name=txtTestOrganizer value="">
		</td>
	</tr>
	<tr ID=BudgetRow>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Test Budget:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtTestBudget name=txtTestBudget value="">
		</td>
	</tr>
	<tr ID=CertNo>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Certification No:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<INPUT type="text" id=txtCertificationNo name=txtCertificationNo value="">
		</td>
	</tr>
	<tr ID=DcrPorSelectRow>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Added By:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<SELECT id=cboPorDcr name=cboPorDcr onchange="cboPorDcr_onchange();">
			<option value="">-- Select One --</option>
			<option value="POR">POR</option>
			<option value="DCR">DCR</option>
			</SELECT>
		</td>
	</tr>
	<tr ID=DcrSelectRow style="display:none">
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Added By DCR:</span>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<SELECT id=cboDcr name=cboDcr>
			<option value="">-- Select A DCR --</option>
			<% FillDcrStatus(m_ProductID) %>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td valign=top width=140 nowrap><span style="font:bold x-small verdana">Notes:</span></td>
		<td>
			<TEXTAREA rows=3 cols=37 id=txtNotes name=txtNotes></TEXTAREA>
		</td>
	</tr>
    <% End If %>

</table>
</div>
<input type="hidden" id="hidDisplay" name="hidDisplay" value="<%= m_FormDisplay%>">
<input type="hidden" id="hidSave" name="hidSave" value="<%= m_FormSave%>">
<input type="hidden" id="hidClose" name="hidClose" value="<%= m_FormClose%>">
<input type="hidden" id="hidEdit" name="hidEdit" value="<%= m_EditModeOn%>">
<input type="hidden" id="hdnPulsarTabName" name="hdnPulsarTabName" value="<%=Request("pulsarplusDivId")%>">
<input type="hidden" id="hidReleaseID" name="hidReleaseID" value="<%= m_ReleaseID%>">
<input type="hidden" id="hidCountryID" name="hidCountryID" value="<%= m_CountryID%>">
<input type="hidden" id="hidProductID" name="hidProductID" value="<%= m_ProductID%>">
<input type="hidden" id="hidBatchUpdate" name="hidBatchUpdate" value="<%= m_ReleaseBatchUpdate%>" />
<input type="hidden" id="hidBatchUpdateRelease" name="hidBatchUpdateRelease" />
<input type="hidden" id="hidBatchUpdateCountry" name="hidBatchUpdateCountry" />
<input type="hidden" id="hidfromWhere" name="hidfromWhere" value="<%=m_from_where %>" />
</form>

<script type="text/javascript">
    function FillReleaseStatus(isReleaseCertification, type, content) {
        if (isReleaseCertification == 'True' && type > -1) {
            if (type == 2) {
                $("#txtProjectedDate").val(content);
            }
            $("#cboStatus").val(type);
        }
    }

    $(document).ready(function () {
        FillReleaseStatus('<%=m_ReleaseCertification %>', '<%=m_CertType %>', '<%=m_CertData %>');
    });
</script>
</BODY>
</HTML>


