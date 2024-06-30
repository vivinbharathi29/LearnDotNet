<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value!="0")
		{
		window.returnValue = "1";
		window.close();
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
Saving.  Please wait...

<%


	dim strConnect
	dim cn
	dim cm
	dim rowseffected
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.Open


	set cm = server.CreateObject("ADODB.command")
		
	cm.ActiveConnection = cn
	cm.CommandText = "spUpdateTodayConfig"
	cm.CommandType =  &H0004


	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(request("txtUserID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Approvals", 11, &H0001)
	if request("chkApproval") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PastDue", 11, &H0001)
	if request("chkPastDue") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Working", 11, &H0001)
	if request("chkWorking") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DueThisWeek", 11, &H0001)
	if request("chkDueThisWeek") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Open", 11, &H0001)
	if request("chkIOwn") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Submitted", 11, &H0001)
	if request("chkSubmitted") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Closed", 11, &H0001)
	if request("chkClosed") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Proposed", 11, &H0001)
	if request("chkProposed") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@NewRequestSection", adTinyInt, &H0001)
	if request("chkNewRequests") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PIPMAlertSection", adTinyInt, &H0001)
	if request("chkPIPMAlert") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PIDBTeamSection", adTinyInt, &H0001)
	if request("chkDBTeam") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PIReassignedSection", adTinyInt, &H0001)
	if request("chkPIReassigned") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PIScripterSection",adInteger, &H0001)
	if request("chkScripter") = "on" then
		if request("optScripters") = "2" then
			p.Value = request("txtUserID")
		else
			p.Value = -3
		end if
	else
		p.Value = 0
	end if
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@FunctionalTestMADTSection", 11, &H0001)
	if request("chkFunctionalTestMADT") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestMultimediaSection", 11, &H0001)
	if request("chkFunctionalTestMultimedia") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@FunctionalTest3rdPartySection", 11, &H0001)
	if request("chkFunctionalTest3rdParty") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTest3rdPartyConsSection", 11, &H0001)
	if request("chkFunctionalTest3rdPartyCons") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTest3rdPartyInternalSection", 11, &H0001)
	if request("chkFunctionalTest3rdPartyInternal") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestHelpAndSupportSection", 11, &H0001)
	if request("chkFunctionalTestHelpAndSupport") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestUserGuidesSection", 11, &H0001)
	if request("chkFunctionalTestUserGuides") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestToolsSection", 11, &H0001)
	if request("chkFunctionalTestTools") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestOtherSection", 11, &H0001)
	if request("chkFunctionalTestOther") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestHWEnablingSection", 11, &H0001)
	if request("chkFunctionalTestHWEnabling") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestMultimediaAppsSection", 11, &H0001)
	if request("chkFunctionalTestMultimediaApps") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestSecuritySection", 11, &H0001)
	if request("chkFunctionalTestSecurity") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestThinClientSection", 11, &H0001)
	if request("chkFunctionalTestThinClient") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestIntelTechnologiesSection", 11, &H0001)
	if request("chkFunctionalTestIntelTechnologies") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestDeveloperSection", 11, &H0001)
	if request("chkFunctionalTestDeveloper") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestBIOSSection", 11, &H0001)
	if request("chkFunctionalTestBIOS") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@FunctionalTestVirtualizationSection", 11, &H0001)
	if request("chkFunctionalTestVirtualization") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PostRTMSection", 11, &H0001)
	if request("chkPostRTM") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OTSOwnerSection", 11, &H0001)
	if request("chkOTSOwner") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OTSSubmittedSection", 11, &H0001)
	if request("chkOTSSubmitted") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OTSDeliverableSection", 11, &H0001)
	if request("chkOTSDeliverable") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DefaultWorkingProject", 3, &H0001)
	p.Value = cint(request("cboWorkingProject"))
	cm.Parameters.Append p



	cm.Execute RowsEffected
	
	if cn.Errors.Count > 1 or Rowseffected <> 1 then
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""0"">"
		Response.Write "<font size=2 face=verdana><b>Unable to save this configuration.</b></font>"
	else
		Response.Write "<INPUT style=""Display:none"" type=""text"" id=txtSuccess name=txtSuccess value=""1"">"
	end if
	

	set cm = nothing
	set cn = nothing



%>

</BODY>
</HTML>
