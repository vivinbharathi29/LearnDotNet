<%@ Language=VBScript %>
<% Option Explicit
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>
	
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<!-- #include file = "../includes/EmailWrapper.asp" -->
<!-- #include file = "clsSchedule.asp" -->

<html>
<head>
<title></title>
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../Scripts/jquery-1.10.2.js"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
<script type="text/javascript">
<!--

    $(function () {
        var OutArray = new Array();
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Reload PulsarPlusPmView Tab
            parent.window.parent.reloadFromPopUp(pulsarplusDivId.value);

            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if ($("#txtSuccess").val() == "1") {
                if (IsFromPulsarPlus()) {
                    window.parent.parent.parent.popupCallBack(1);
                    ClosePulsarPlusPopup();
                }
                else {
                    OutArray[0] = $("#txtOut0").val();
                    OutArray[1] = $("#txtOut1").val();
                    OutArray[2] = $("#txtOut2").val();
                    OutArray[3] = $("#txtOut3").val();
                    OutArray[4] = $("#txtOut4").val();
                    OutArray[5] = $("#txtOut5").val();

                    if (window.parent.frames["UpperWindow"]) {
                        parent.window.parent.modalDialog.cancel(true);
                    } else {
                        var iframeName = parent.window.name;
                        if (iframeName != '') {
                            parent.window.parent.ClosePropertiesDialog(OutArray);
                        } else {
                            window.returnValue = OutArray;
                            if (IsFromPulsarPlus()) {
                                ClosePulsarPlusPopup();
                            }
                            else {
                                window.parent.close();
                            }
                        }
                    }
                }
            }
            else {
                document.write(document.body.innerHTML + "<BR><BR>Unable to update schedule.  An unexpected error occurred.");
            }
        }
    });

//-->
</script>
</head>
<body >

<%
	dim cn
	Dim dw
	dim cmd
	dim iRowsChanged
	dim FoundErrors
	Dim sShowOnReports_YN
	Dim strSuccess
	Dim bNotifyOfPddLock
	bNotifyOfPddLock = False
	
'##############################################################################	
'
' Create Security Object to get User Info
'
	Dim m_IsSysAdmin
	Dim m_IsProgramManager
	Dim m_IsSysEngProgramManager
	Dim m_IsSysTeamLead
    Dim m_IsPlatformPm
    Dim m_IsSEPMProductsEditor
	Dim m_EditModeOn
	Dim m_UserEmail
	Dim m_IsSupplyChain
	
	m_EditModeOn = False
	
	Dim Security
	Dim sUserFullName
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

    m_IsSupplyChain = Security.UserInRole(Request("PVID"), "SupplyChain")
	m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
	m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
    m_IsPlatformPm = Security.IsPlatformDevMgr(Request("PVID"))
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	sUserFullName = Security.CurrentUser()
	m_UserEmail = Security.CurrentUserEmail()
	
	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsPlatformPm Or m_IsSupplyChain Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Unable to save data changes</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

	If Request.Form("cbxShowOnReports") = "on" Then
		sShowOnReports_YN = "Y"
	Else
		sShowOnReports_YN = "N"
	End If

	FoundErrors = false	

	'Create Database Connection
	Set dw = New DataWrapper
	set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans

    Dim obj
    Set obj = New Schedule
    CALL obj.UpdateMilestone(cn, Trim(Request.Form("hidPorEndDt")), Request.Form("hidScheduleDataID"), Trim(sUserFullName), m_UserEmail, Request.Form("txtComments"), Request.Form("txtItemNotes"), Request("hidProjectedStartDt"), Request("hidProjectedEndDt"), Request.Form("txtProjectedStartDt"), Request.Form("txtProjectedEndDt"), Request.Form("txtActualStartDt"), Request.Form("txtActualEndDt"), sShowOnReports_YN, Request.Form("selItemPhase"), Request.Form("hidScheduleDefinitionDataID"), Request.Form("selItemOwner"))

    cn.CommitTrans

    strSuccess = "1"

	Set dw = nothing
	Set cmd = nothing	
	set cn = nothing	

%>

<input type="text" id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>" /><br />
<input type="text" id="txtOut0" name="txtOut0" value="<%=Request.Form("hidScheduleDataID")%>" />
<input type="text" id="txtOut1" name="txtOut1" value="<%=Request.Form("txtProjectedStartDt")%>" />
<input type="text" id="txtOut2" name="txtOut2" value="<%=Request.Form("txtProjectedEndDt")%>" />
<input type="text" id="txtOut3" name="txtOut3" value="<%=Request.Form("txtActualStartDt")%>" />
<input type="text" id="txtOut4" name="txtOut4" value="<%=Request.Form("txtActualEndDt")%>" />
<input type="text" id="txtOut5" name="txtOut5" value="<%=Request.Form("txtItemNotes")%>" />
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=Request.Form("pulsarplusDivId")%>">
</body>
</html>
