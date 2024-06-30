<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var CurrentState;
//var States = new Array(2);
var FormLoading = true;
//var DisplayedID;

function ProcessState() {
	var steptext;
	
	switch (CurrentState)
	{
		case "Actions":
			steptext = "";

			tabActions.style.display="";
			//tabPreinstall.style.display="none";
			tabTest.style.display="none";
			tabOTS.style.display="none";
			tabMisc.style.display="none";

			window.scrollTo(0,0);		
		break;

		case "OTS":
			steptext = "";
	
			tabActions.style.display="none";
			//tabPreinstall.style.display="none";
			tabTest.style.display="none";
			tabOTS.style.display="";
			tabMisc.style.display="none";

			window.scrollTo(0,0);		
		break;
/*
		case "Preinstall":
			steptext = "";
	
			tabActions.style.display="none";
			tabPreinstall.style.display="";
			tabOTS.style.display="none";
			tabMisc.style.display="none";

			window.scrollTo(0,0);		
		break;
*/
		case "Misc":
			steptext = "";
	
			tabActions.style.display="none";
			//tabPreinstall.style.display="none";
			tabTest.style.display="none";
			tabOTS.style.display="none";
			tabMisc.style.display="";

			window.scrollTo(0,0);		
		break;
		case "Test":
			steptext = "";
	
			tabActions.style.display="none";
			//tabPreinstall.style.display="none";
			tabTest.style.display="";
			tabOTS.style.display="none";
			tabMisc.style.display="none";

			window.scrollTo(0,0);		
		break;

	}
}


function SelectTab(strStep) {
	var i;

	//Reset all tabs
	document.all("CellActionsb").style.display="none";
	document.all("CellActions").style.display="";
	//document.all("CellPreinstallb").style.display="none";
	//document.all("CellPreinstall").style.display="";
	document.all("CellOTSb").style.display="none";
	document.all("CellOTS").style.display="";
	document.all("CellTestb").style.display="none";
	document.all("CellTest").style.display="";
	document.all("CellMiscb").style.display="none";
	document.all("CellMisc").style.display="";

	//Highight the selected tab
	document.all("Cell"+strStep).style.display="none";
	document.all("Cell"+strStep+"b").style.display="";

	
	CurrentState = strStep;
	ProcessState();

}

function ApprovalText_onclick() {
	if (Configure.chkApproval.checked)
		Configure.chkApproval.checked = false;
	else
		Configure.chkApproval.checked = true;
}

function ApprovalText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function PastDueText_onclick() {
	if (Configure.chkPastDue.checked)
		Configure.chkPastDue.checked = false;
	else
		Configure.chkPastDue.checked = true;
}

function PastDueText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function DueThisWeekText_onclick() {
	if (Configure.chkDueThisWeek.checked)
		Configure.chkDueThisWeek.checked = false;
	else
		Configure.chkDueThisWeek.checked = true;
}

function DueThisWeekText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function SubmittedText_onclick() {
	if (Configure.chkSubmitted.checked)
		Configure.chkSubmitted.checked = false;
	else
		Configure.chkSubmitted.checked = true;
}

function SubmittedText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function OpenText_onclick() {
	if (Configure.chkIOwn.checked)
		Configure.chkIOwn.checked = false;
	else
		Configure.chkIOwn.checked = true;
}

function OpenText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ClosedText_onclick() {
	if (Configure.chkClosed.checked)
		Configure.chkClosed.checked = false;
	else
		Configure.chkClosed.checked = true;
}

function ClosedText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ProposedText_onclick() {
	if (Configure.chkProposed.checked)
		Configure.chkProposed.checked = false;
	else
		Configure.chkProposed.checked = true;
}

function ProposedText_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function OTSOwnerText_onclick() {
	if (Configure.chkOTSOwner.checked)
		Configure.chkOTSOwner.checked = false;
	else
		Configure.chkOTSOwner.checked = true;
}

function OTSDeliverableText_onclick() {
	if (Configure.chkOTSDeliverable.checked)
		Configure.chkOTSDeliverable.checked = false;
	else
		Configure.chkOTSDeliverable.checked = true;
}

function OTSSubmittedText_onclick() {
	if (Configure.chkOTSSubmitted.checked)
		Configure.chkOTSSubmitted.checked = false;
	else
		Configure.chkOTSSubmitted.checked = true;
}

function PostRTMText_onclick() {
	if (Configure.chkPostRTM.checked)
		Configure.chkPostRTM.checked = false;
	else
		Configure.chkPostRTM.checked = true;
}


function FTDeveloperText_onclick() {
	if (Configure.chkFunctionalTestDeveloper.checked)
		Configure.chkFunctionalTestDeveloper.checked = false;
	else
		Configure.chkFunctionalTestDeveloper.checked = true;
}

function FTOtherText_onclick() {
    if (Configure.chkFunctionalTestOther.checked)
        Configure.chkFunctionalTestOther.checked = false;
    else
        Configure.chkFunctionalTestOther.checked = true;
}

function FTBIOSText_onclick() {
    if (Configure.chkFunctionalTestBIOS.checked)
        Configure.chkFunctionalTestBIOS.checked = false;
    else
        Configure.chkFunctionalTestBIOS.checked = true;
}

function FT3rdPartyInternalText_onclick() {
    if (Configure.chkFunctionalTest3rdPartyInternal.checked)
        Configure.chkFunctionalTest3rdPartyInternal.checked = false;
    else
        Configure.chkFunctionalTest3rdPartyInternal.checked = true;
}

function FTIntelTechnologiesText_onclick() {
	if (Configure.chkFunctionalTestIntelTechnologies.checked)
		Configure.chkFunctionalTestIntelTechnologies.checked = false;
	else
		Configure.chkFunctionalTestIntelTechnologies.checked = true;
}

function FTHWEnablingText_onclick() {
	if (Configure.chkFunctionalTestHWEnabling.checked)
		Configure.chkFunctionalTestHWEnabling.checked = false;
	else
		Configure.chkFunctionalTestHWEnabling.checked = true;
}

function FTMultimediaAppsText_onclick() {
    if (Configure.chkFunctionalTestMultimediaApps.checked)
        Configure.chkFunctionalTestMultimediaApps.checked = false;
    else
        Configure.chkFunctionalTestMultimediaApps.checked = true;
}
function FTVirtualizationText_onclick() {
    if (Configure.chkFunctionalTestVirtualization.checked)
        Configure.chkFunctionalTestVirtualization.checked = false;
    else
        Configure.chkFunctionalTestVirtualization.checked = true;
}

function FTSecurityText_onclick() {
    if (Configure.chkFunctionalTestSecurity.checked)
        Configure.chkFunctionalTestSecurity.checked = false;
    else
        Configure.chkFunctionalTestSecurity.checked = true;
}
function FTThinClientText_onclick() {
    if (Configure.chkFunctionalTestThinClient.checked)
        Configure.chkFunctionalTestThinClient.checked = false;
    else
        Configure.chkFunctionalTestThinClient.checked = true;
}


function FTToolsText_onclick() {
	if (Configure.chkFunctionalTestTools.checked)
		Configure.chkFunctionalTestTools.checked = false;
	else
		Configure.chkFunctionalTestTools.checked = true;
}

function FTHelpText_onclick() {
	if (Configure.chkFunctionalTestHelpAndSupport.checked)
		Configure.chkFunctionalTestHelpAndSupport.checked = false;
	else
		Configure.chkFunctionalTestHelpAndSupport.checked = true;
}
function FT3rdPartyText_onclick() {
	if (Configure.chkFunctionalTest3rdParty.checked)
		Configure.chkFunctionalTest3rdParty.checked = false;
	else
		Configure.chkFunctionalTest3rdParty.checked = true;
}

function FT3rdPartyConsText_onclick() {
	if (Configure.chkFunctionalTest3rdPartyCons.checked)
		Configure.chkFunctionalTest3rdPartyCons.checked = false;
	else
		Configure.chkFunctionalTest3rdPartyCons.checked = true;
}

function FTMADTText_onclick() {
	if (Configure.chkFunctionalTestMADT.checked)
		Configure.chkFunctionalTestMADT.checked = false;
	else
		Configure.chkFunctionalTestMADT.checked = true;
}

function FTMultimediaText_onclick() {
	if (Configure.chkFunctionalTestMultimedia.checked)
		Configure.chkFunctionalTestMultimedia.checked = false;
	else
		Configure.chkFunctionalTestMultimedia.checked = true;
}

function FTUserGuidesText_onclick() {
	if (Configure.chkFunctionalTestUserGuides.checked)
		Configure.chkFunctionalTestUserGuides.checked = false;
	else
		Configure.chkFunctionalTestUserGuides.checked = true;
}

function window_onload() {
	//CurrentState =  Configure.txtStartTab.value;//"General";
	//ProcessState();

	SelectTab(Configure.txtStartTab.value);
	FormLoading = false;
	self.focus();
}

//-->
</SCRIPT>
</HEAD>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<body bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">


<%
	dim CurrentUser 
	dim CurrentUserID
	dim CurrentUserName
	dim CurrentUserGroupID
	dim i
	dim blnApprovalSection
	dim blnWorkingSection
	dim blnPastDueSection
	dim blnDueThisWeekSection
	dim blnAllOpenSection
	dim blnISubmittedSection
	dim blnClosedSection
	dim blnProposedSection
	dim blnFunctionalTestMADTSection
	dim blnFunctionalTestMultimediaSection
	dim blnFunctionalTestUserGuidesSection
	dim blnFunctionalTestOtherSection
	dim blnFunctionalTestBIOSSection
    dim blnFunctionalTest3rdPartyInternalSection
	dim blnFunctionalTestIntelTechnologiesSection
	dim blnFunctionalTestHWEnablingSection
	dim blnFunctionalTestMultimediaAppsSection
    dim blnFunctionalTestSecuritySection
    dim blnFunctionalTestThinClientSection
	dim blnFunctionalTestToolsSection
	dim blnFunctionalTest3rdPartySection 
	dim blnFunctionalTestVirtualizationSection 
	dim blnFunctionalTest3rdPartyConsSection
	dim blnFunctionalTestHelpAndSupportSection
	dim blnFunctionalTestDeveloperSection
	
	dim strPIPMAlertSection
	dim strPINewRequestSection
	dim strScripterSection
	dim strPIDBTeamSection
	dim strPIReassignedSection
	dim strAllScripters
	dim strMyScriptsOnly
	dim strOTSSumbitted
	dim strOTSOwner
	dim strOTSDeliverable
	dim strDefaultWorkingProject
	dim strWorkingList
	
	dim cn
	dim rs	
	dim cm
	dim p
	

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open



	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	

	CurrentUserID = 0
	CurrentUserGroupID = 0
	blnApprovalSection = true
	blnWorkingSection = false
	blnPastDueSection = true
	blnDueThisWeekSection = true
	blnAllOpenSection = true
	blnISubmittedSection = true
	blnClosedSection = true
	ProposedSection = true
	strPIPMAlertSection = ""
	strPINewRequestSection = ""
	strPIDBTeamSection = ""
	strScripterSection = ""
	strPIReassignedSection = ""
	strAllScripters = "checked"
	strMyScriptsOnly = ""
	strOTSSumbittedSection = ""
	strOTSOwnerSection = ""
	strOTSDeliverableSection = ""
	strPostRTMSection = ""
	strDefaultWorkingProject = ""
	strWorkingList = ""
	
'	rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserName = rs("Name") & ""
		CurrentUserPartner = rs("PartnerID") & ""
		CurrentUserDivision = rs("Division") & ""
		CurrentUserDivisionID = CurrentUserDivision
		if CurrentUserDivision = "1" then
			CurrentUserDivision = " - Mobile"
		elseif CurrentUserDivision = "2" then
			CurrentUserDivision = " - Desktops"
		elseif CurrentUserDivision = "3" then
			CurrentUserDivision = " - PDC"
		end if
		CurrentUserGroupID = rs("workgroupID")
		blnApprovalSection = rs("ApprovalSection")
		blnWorkingSection = rs("WorkingSection")
		blnPastDueSection = rs("PastDueSection")
		blnDueThisWeekSection = rs("DueThisWeekSection")
		blnAllOpenSection = rs("AllOpenSection")
		blnISubmittedSection = rs("ISubmittedSection")
		blnClosedSection = rs("ClosedSection")
		blnProposedSection = rs("ProposedSection")
		strPIPMAlertSection = rs("PIPMAlertSection") & ""
		strPINewRequestSection = rs("PINewRequestSection") & ""
		strScripterSection = rs("PIScripterSection") & ""
		strPIDBTeamSection = rs("PIDBTeamSection") & ""
		strPIReassignedSection = rs("PIReassignedSection") & ""
		strOTSSumbittedSection = rs("OTSSubmittedSection") & ""
		strOTSOwnerSection = rs("OTSOwnerSection") & ""
		strOTSDeliverableSection = rs("OTSDeliverableSection") & ""
		strPostRTMSection = rs("PostRTMSection") & ""
		blnFunctionalTestUserGuidesSection = rs("FunctionalTestUserGuidesSection")
		blnFunctionalTestMADTSection = rs("FunctionalTestMADTSection")
		blnFunctionalTestMultimediaSection = rs("FunctionalTestMultimediaSection")
		blnFunctionalTestOtherSection = rs("FunctionalTestOtherSection")
		blnFunctionalTestBIOSSection = rs("FTBIOSSection")
		blnFunctionalTest3rdPartyInternalSection = rs("FunctionalTest3rdPartyInternalSection")
		blnFunctionalTestIntelTechnologiesSection = rs("FTIntelTechnologiesSection")
		blnFunctionalTestHWEnablingSection = rs("FTHWENablingSection")
		blnFunctionalTestMultimediaAppsSection = rs("FTMultimediaAppsSection")
        blnFunctionalTestSecuritySection = rs("FTSecuritySection")
        blnFunctionalTestThinClientSection = rs("FTThinClientSection")
		blnFunctionalTestToolsSection = rs("FunctionalTestToolsSection")
		blnFunctionalTest3rdPartySection = rs("FunctionalTest3rdPartySection")
        blnFunctionalTestVirtualizationSection = rs("FTVirtualizationSection")
		blnFunctionalTest3rdPartyConsSection = rs("FunctionalTest3rdPartyConsSection")
		blnFunctionalTestHelpAndSupportSection = rs("FunctionalTestHelpAndSupportSection")
		blnFunctionalTestDeveloperSection = rs("FunctionalTestDeveloperSection")
		strDefaultWorkingProject = trim(rs("DefaultWorkingListProduct") & "")
	end if
	rs.Close

	rs.open "spGetProducts 2",cn,adOpenForwardOnly	strWorkingList = "<SELECT style=""width: 200"" id=cboWorkingProject name=cboWorkingProject>"	strWorkingList = strWorkingList & "<OPTION selected value=""0""></OPTION>"
	do while not rs.eof
		if trim(rs("ID")) = strDefaultWorkingProject then			strWorkingList = strWorkingList & "<OPTION selected value=""" & rs("ID") & """>" & rs("name")  & " " & rs("Version") & "</OPTION>"
		else
			strWorkingList = strWorkingList & "<OPTION value=""" & rs("ID") & """>" & rs("name")   & " " & rs("Version") & "</OPTION>"
		end if
		rs.movenext	
	loop
	rs.close	
	strWorkingList = strWorkingList & "</SELECT>"
	
	if isnull(blnFunctionalTestMADTSection) then
		blnFunctionalTestMADTSection = ""
	else
		blnFunctionalTestMADTSection = replace(replace(blnFunctionalTestMADTSection,true,"checked"),false,"")
	end if

	if isnull(blnFunctionalTestMultimediaSection) then
		blnFunctionalTestMultimediaSection = ""
	else
		blnFunctionalTestMultimediaSection = replace(replace(blnFunctionalTestMultimediaSection,true,"checked"),false,"")
	end if

	if isnull(blnFunctionalTestUserGuidesSection) then
		blnFunctionalTestUserGuidesSection = ""
	else
		blnFunctionalTestUserGuidesSection = replace(replace(blnFunctionalTestUserGuidesSection,true,"checked"),false,"")
	end if

	if isnull(blnFunctionalTestOtherSection) then
		blnFunctionalTestOtherSection = ""
	else
		blnFunctionalTestOtherSection = replace(replace(blnFunctionalTestOtherSection,true,"checked"),false,"") 
	end if

	if isnull(blnFunctionalTestBIOSSection) then
		blnFunctionalTestBIOSSection = ""
	else
		blnFunctionalTestBIOSSection = replace(replace(blnFunctionalTestBIOSSection,true,"checked"),false,"") 
	end if

	if isnull(blnFunctionalTestIntelTechnologiesSection) then
		blnFunctionalTestIntelTechnologiesSection = ""
	else
		blnFunctionalTestIntelTechnologiesSection = replace(replace(blnFunctionalTestIntelTechnologiesSection,true,"checked"),false,"") 
	end if

	if isnull(blnFunctionalTestHWEnablingSection) then
		blnFunctionalTestHWEnablingSection = ""
	else
		blnFunctionalTestHWEnablingSection = replace(replace(blnFunctionalTestHWEnablingSection,true,"checked"),false,"") 
	end if
    
	if isnull(blnFunctionalTestMultimediaAppsSection) then
		blnFunctionalTestMultimediaAppsSection = ""
	else
		blnFunctionalTestMultimediaAppsSection = replace(replace(blnFunctionalTestMultimediaAppsSection,true,"checked"),false,"") 
	end if

	if isnull(blnFunctionalTestSecuritySection) then
		blnFunctionalTestSecuritySection = ""
	else
		blnFunctionalTestSecuritySection = replace(replace(blnFunctionalTestSecuritySection,true,"checked"),false,"") 
	end if

	if isnull(blnFunctionalTestThinClientSection) then
		blnFunctionalTestThinClientSection = ""
	else
		blnFunctionalTestThinClientSection = replace(replace(blnFunctionalTestThinClientSection,true,"checked"),false,"") 
	end if
    
	if isnull(blnFunctionalTestToolsSection) then
		blnFunctionalTestToolsSection = ""
	else
		blnFunctionalTestToolsSection = replace(replace(blnFunctionalTestToolsSection,true,"checked"),false,"") 
	end if
    
	if isnull(blnFunctionalTestVirtualizationSection) then
		blnFunctionalTestVirtualizationSection = ""
	else
		blnFunctionalTestVirtualizationSection = replace(replace(blnFunctionalTestVirtualizationSection,true,"checked"),false,"")
	end if

	if isnull(blnFunctionalTest3rdPartySection) then
		blnFunctionalTest3rdPartySection = ""
	else
		blnFunctionalTest3rdPartySection = replace(replace(blnFunctionalTest3rdPartySection,true,"checked"),false,"")
	end if
	if isnull(blnFunctionalTest3rdPartyConsSection) then
		blnFunctionalTest3rdPartyConsSection=""
	else
		blnFunctionalTest3rdPartyConsSection = replace(replace(blnFunctionalTest3rdPartyConsSection,true,"checked"),false,"")
	end if

	if isnull(blnFunctionalTest3rdPartyInternalSection) then
		blnFunctionalTest3rdPartyInternalSection=""
	else
		blnFunctionalTest3rdPartyInternalSection = replace(replace(blnFunctionalTest3rdPartyInternalSection,true,"checked"),false,"")
	end if	
	if isnull(blnFunctionalTestHelpAndSupportSection) then
		blnFunctionalTestHelpAndSupportSection = ""
	else
		blnFunctionalTestHelpAndSupportSection = replace(replace(blnFunctionalTestHelpAndSupportSection,true,"checked"),false,"")
	end if
	if isnull(blnFunctionalTestDeveloperSection) then
		blnFunctionalTestDeveloperSection = ""
	else
		blnFunctionalTestDeveloperSection = replace(replace(blnFunctionalTestDeveloperSection,true,"checked"),false,"")
	end if 
	blnApprovalSection = replace(replace(blnApprovalSection,true,"checked"),false,"")
	blnWorkingSection = replace(replace(blnWorkingSection,true,"checked"),false,"")
	blnPastDueSection = replace(replace(blnPastDueSection,true,"checked"),false,"")
	blnDueThisWeekSection = replace(replace(blnDueThisWeekSection,true,"checked"),false,"")
	blnAllOpenSection = replace(replace(blnAllOpenSection,true,"checked"),false,"")
	blnISubmittedSection = replace(replace(blnISubmittedSection,true,"checked"),false,"")
	blnClosedSection = replace(replace(blnClosedSection,true,"checked"),false,"")
	blnProposedSection = replace(replace(blnProposedSection,true,"checked"),false,"")
	strPIPMAlertSection = replace(replace(strPIPMAlertSection,"1","checked"),"0","")
	strPINewRequestSection = replace(replace(strPINewRequestSection,"1","checked"),"0","")
	strPIDBTeamSection = replace(replace(strPIDBTeamSection,"1","checked"),"0","")
	strPIReassignedSection = replace(replace(strPIReassignedSection,"1","checked"),"0","")
	strOTSSumbittedSection = replace(replace(strOTSSumbittedSection,true,"checked"),false,"")
	strOTSOwnerSection = replace(replace(strOTSOwnerSection,true,"checked"),false,"")
	strOTSDeliverableSection = replace(replace(strOTSDeliverableSection,true,"checked"),false,"")
	strPostRTMSection = replace(replace(strPostRTMSection,true,"checked"),false,"")

	if strScripterSection = "0" then
		'No
		strScripterSection = ""
		strAllScripters = "checked"
		strMyScriptsOnly = ""
	elseif strScripterSection = "1" or trim(strScripterSection) = "-3" then
		'Yes All
		strScripterSection = "checked"
		strAllScripters = "checked"
		strMyScriptsOnly = ""
	elseif strScripterSection > 1 then
		'Yes Mine
		strScripterSection = "checked"
		strAllScripters = ""
		strMyScriptsOnly = "checked"
	end if

	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0

%>


<FORM ACTION="ConfigureSave.asp" METHOD="post" NAME="Configure">
<font face=verdana>
<table Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0">
	<tr bgcolor="<%=strTitleColor%>">
		<td id="CellActions" style="Display:none" width="10"><font size="2" color="black"><b>&nbsp;<a href="javascript:SelectTab('Actions')">Actions</a>&nbsp;</b></font></td>
		<td id="CellActionsb" style="Display:" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Actions&nbsp;</b></font></td>
		<td id="CellOTS" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('OTS')">OTS</a>&nbsp;</b></font></td>
		<td id="CellOTSb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;OTS&nbsp;</b></font></td>
		<!--
		<td id="CellPreinstall" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Preinstall')">Preinstall</a>&nbsp;</b></font></td>
		<td id="CellPreinstallb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Preinstall&nbsp;</b></font></td>
		-->
<%if trim(CurrentUserDivision) = "1" or trim(CurrentUserPartner) = "1" then%>
		<td id="CellTest" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Test')">Test</a>&nbsp;</b></font></td>
<%else%>
		<dir id="CellTest"></div>
<%end if%>
		<td id="CellTestb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Test&nbsp;</b></font></td>
<%if trim(CurrentUserDivision) = "1" or trim(CurrentUserPartner) = "1" then%>
		<td id="CellMisc" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Misc')">Misc</a>&nbsp;</b></font></td>
<%else%>
		<dir id="CellMisc"></div>
<%end if%>
		<td id="CellMiscb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Misc&nbsp;</b></font></td>
	</tr>
</table>

<hr color="Tan">
<strong><font size=2>Sections To Display</font></strong>
<table ID=tabActions border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr><td><INPUT id=chkApproval <%=blnApprovalSection%> type=checkbox name=chkApproval><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ApprovalText_onmouseover()" onclick="return ApprovalText_onclick()"> Items Awaiting My Approval</font></td></TR>
  <tr><td><INPUT id=chkPastDue <%=blnPastDueSection%> type=checkbox name=chkPastDue><font face=verdana size=2  LANGUAGE=javascript onmouseover="return PastDueText_onmouseover()" onclick="return PastDueText_onclick()"> Items Past Due</font></td></TR>
  <tr><td><INPUT id=chkDueThisWeek <%=blnDueThisWeekSection%> type=checkbox name=chkDueThisWeek><font face=verdana size=2  LANGUAGE=javascript onmouseover="return DueThisWeekText_onmouseover()" onclick="return DueThisWeekText_onclick()"> Items Due This Week</font></td></TR>
  <tr><td><INPUT id=chkSubmitted <%=blnISubmittedSection%> type=checkbox name=chkSubmitted><font face=verdana size=2  LANGUAGE=javascript onmouseover="return SubmittedText_onmouseover()" onclick="return SubmittedText_onclick()"> All Open Items I Submitted</font></td></TR>
  <tr><td><INPUT id=chkIOwn <%=blnAllOpenSection%> type=checkbox name=chkIOwn><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return OpenText_onclick()"> All Open Items I Own</font></td></TR>
  <tr><td><INPUT id=chkWorking <%=blnWorkingSection%> type=checkbox name=chkWorking><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return OpenText_onclick()"> All Open Items On My Working List</font><font color=green size=1 face=verdana> (For Tool Projects)</font></td></TR>
  <tr><td><INPUT id=chkClosed <%=blnClosedSection%> type=checkbox name=chkClosed><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ClosedText_onmouseover()" onclick="return ClosedText_onclick()"> Items Closed this Week</font> <font color=green size=1 face=verdana> (For Products in Favorites)</font></td></TR>
  <tr><td><INPUT id=chkProposed <%=blnProposedSection%> type=checkbox name=chkProposed><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> All Proposed Items</font> <font color=green size=1 face=verdana> (For Products in Favorites)</font></td></TR>

</table>
<table style="display:none" ID=tabTest border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr><td colspan=2><b>Deliverables in Functional Test</b></td></TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTest3rdParty <%=blnFunctionalTest3rdPartySection%> type=checkbox name=chkFunctionalTest3rdParty><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FT3rdPartyText_onclick()"> 3rd Party SW - Commercial</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestMADT <%=blnFunctionalTestMADTSection%> type=checkbox name=chkFunctionalTestMADT><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTMADTText_onclick()"> MADT</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTest3rdPartyCons <%=blnFunctionalTest3rdPartyConsSection%> type=checkbox name=chkFunctionalTest3rdPartyCons><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FT3rdPartyConsText_onclick()"> 3rd Party SW - Consumer</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestUserGuides <%=blnFunctionalTestUserGuidesSection%> type=checkbox name=chkFunctionalTestUserGuides><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTUserGuidesText_onclick()"> User Guides</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTest3rdPartyInternal <%=blnFunctionalTest3rdPartyInternalSection%> type=checkbox name=chkFunctionalTest3rdPartyInternal><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FT3rdPartyInternalText_onclick()"> 3rd Party SW - Internal</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestTools <%=blnFunctionalTestToolsSection%> type=checkbox name=chkFunctionalTestTools><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTToolsText_onclick()"> Tools</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTestHelpAndSupport <%=blnFunctionalTestHelpAndSupportSection%> type=checkbox name=chkFunctionalTestHelpAndSupport><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTHelpText_onclick()"> Help and Support</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalHWEnabling <%=blnFunctionalTestHWEnablingSection%> type=checkbox name=chkFunctionalTestHWEnabling><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTHWEnablingText_onclick()"> HW Enabling</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTestIntelTechnologies <%=blnFunctionalTestIntelTechnologiesSection%> type=checkbox name=chkFunctionalTestIntelTechnologies><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTIntelTechnologiesText_onclick()"> Intel Technologies</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestMultimedia <%=blnFunctionalTestMultimediaSection%> type=checkbox name=chkFunctionalTestMultimedia><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTMultimediaText_onclick()"> Multimedia</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTestMultimediaApps <%=blnFunctionalTestMultimediaAppsSection%> type=checkbox name=chkFunctionalTestMultimediaApps><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTMultimediaAppsText_onclick()"> Multimedia Apps</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestBIOS <%=blnFunctionalTestBIOSSection%> type=checkbox name=chkFunctionalTestBIOS><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTBIOSText_onclick()"> BIOS</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalThinClient <%=blnFunctionalTestThinClientSection%> type=checkbox name=chkFunctionalTestThinClient><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTThinClientText_onclick()"> Thin Client Software</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestSecurity <%=blnFunctionalTestSecuritySection%> type=checkbox name=chkFunctionalTestSecurity><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTSecurityText_onclick()"> Security</font></td>
  </TR>
  <tr>
    <td>&nbsp;<INPUT id=chkFunctionalTestVirtualization <%=blnFunctionalTestVirtualizationSection%> type=checkbox name=chkFunctionalTestVirtualization><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTVirtualizationText_onclick()"> Virtualization Software</font></td>
    <td>&nbsp;<INPUT id=chkFunctionalTestOther <%=blnFunctionalTestOtherSection%> type=checkbox name=chkFunctionalTestOther><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTOtherText_onclick()"> Other</font></td>
  </TR>
  
  <tr><td colspan=2>&nbsp;<INPUT id=chkFunctionalTestDeveloper <%=blnFunctionalTestDeveloperSection%> type=checkbox name=chkFunctionalTestDeveloper><font face=verdana size=2  LANGUAGE=javascript onmouseover="return OpenText_onmouseover()" onclick="return FTDeveloperText_onclick()"> Deliverables I Released (Developer) - Any Test Team</font></td></TR>

</table>

	<%if CurrentUserGroupID <> 15 and CurrentUserGroupID <> 22 then
		strDisplayPI = "none"
	else
		strDisplayPI = ""
	end if%>
<table style="Display:none" ID=tabPreinstall border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr style="display:<%=strDisplayPI%>"><td><INPUT id=chkNewRequests <%=strPINewRequestSection%> type=checkbox name=chkNewRequests><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> New Requests</font> <font color=blue size=1 face=verdana> (Recommended for Scripters)</font></td></TR>
  <tr style="display:<%=strDisplayPI%>"><td><INPUT id=chkScripter <%=strScripterSection%> type=checkbox name=chkScripter><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> Scripting In Progress</font> <font color=blue size=1 face=verdana> (Recommended for Scripters)<BR></font>
  <TABLE style="display:<%=strDisplayPI%>"><TR><TD width=20>&nbsp;</td><td><INPUT <%=strAllScripters%> type="radio" id=optAllScripts name=optScripters value="1"> Show All Scripters<BR><INPUT <%=strMyScriptsOnly%> type="radio" id=optMyScripts name=optScripters value="2"> Show Mine Only<BR></td></tr></table>
  </td></TR>
  <tr style="display:<%=strDisplayPI%>"><td><INPUT id=chkDBTeam <%=strPIDBTeamSection%> type=checkbox name=chkDBTeam><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> Database Team Alerts</font> <font color=blue size=1 face=verdana> (Recommended for Database Team)</font></td></TR>
  <tr style="display:<%=strDisplayPI%>"><td><INPUT id=chkPIPMAlert <%=strPIPMAlertSection%> type=checkbox name=chkPIPMAlert><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> Upcoming Requests</font></td></TR>
  <tr style="display:<%=strDisplayPI%>"><td><INPUT id=chkPIReassigned <%=strPIReassignedSection%> type=checkbox name=chkPIReassigned><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ProposedText_onmouseover()" onclick="return ProposedText_onclick()"> Deliverables Assigned to Other Teams</font></td></TR>
	<%if CurrentUserGroupID <> 15 and CurrentUserGroupID <> 22 then%>
		<tr><td>The Preinstall section will only be displayed for people in one of the Preinstall Workgroups.
	<%end if%>
</table>
<table style="Display:none" ID=tabOTS border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr><td><INPUT id=chkOTSOwner <%=strOTSOwnerSection%> type=checkbox name=chkOTSOwner><font face=verdana size=2  LANGUAGE=javascript onmouseover="return PastDueText_onmouseover()" onclick="return OTSOwnerText_onclick()"> Observations Assigned To Me</font></td></TR>
  <tr><td><INPUT id=chkOTSDeliverable <%=strOTSDeliverableSection%> type=checkbox name=chkOTSDeliverable><font face=verdana size=2  LANGUAGE=javascript onmouseover="return DueThisWeekText_onmouseover()" onclick="return OTSDeliverableText_onclick()"> Observations On My Deliverables</font></td></TR>
  <tr><td><INPUT id=chkOTSSubmitted <%=strOTSSumbittedSection%> type=checkbox name=chkOTSSubmitted><font face=verdana size=2  LANGUAGE=javascript onmouseover="return ApprovalText_onmouseover()" onclick="return OTSSubmittedText_onclick()"> Observations I Submitted</font></td></TR>
</table>

<table style="Display:none" ID=tabMisc border="1" cellPadding="2" cellSpacing="0" width="400" bgcolor=cornsilk bordercolor=tan>
  <tr><td><INPUT id=chkPostRTM <%=strPostRTMSection%> type=checkbox name=chkPostRTM><font face=verdana size=2  LANGUAGE=javascript onmouseover="return PastDueText_onmouseover()" onclick="return PostRTMText_onclick()"> Sustaining Product Support</font></td></TR>
  <tr><td><font face=verdana size=2>Default Working Project:&nbsp;</font><%=strWorkingList%></td></TR>
</table>

<INPUT type="hidden" id=txtUserID name=txtUserID value="<%=CurrentUserID%>">
<%if trim(CurrentUserDivision) = "1" or trim(CurrentUserPartner) = "1" then%>
<%
	dim strTab
	strTab = request("Tab")
	if not (strTab = "Actions" or strTab = "Test" or strTab = "OTS" or strTab = "Misc") then
		strTab="Actions"
	end if

%>


	<INPUT type="hidden" id=txtStartTab name=txtStartTab value="<%=strTab%>">
<%else%>
	<INPUT type="hidden" id=txtStartTab name=txtStartTab value="Actions">
<%end if%>
</FORM>
</font>
</body>

</HTML>
