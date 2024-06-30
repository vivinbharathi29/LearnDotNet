<%@ Language=VBScript %>
<%Option Explicit%>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 

<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim m_IsSysAdmin
Dim m_IsProgramManager 'CM/POPM
Dim m_IsSysEngProgramManager 'SEPM
Dim m_IsSysTeamLead 'SM/STL
Dim m_IsSEPMProductsEditor 'SEPM Products
Dim m_IsPOPManager
Dim m_UserFullName
Dim m_EditModeOn
Dim m_ScheduleID
Dim m_ScheduleDescription
Sub Main()
'##############################################################################	
'
' Create Security Object to get User Info
'

	Dim Security
	
	Set Security = New ExcaliburSecurity
	
' Debug Section
'
'	If Security.CurrentUserID = 1396 Then
'		m_IsSysAdmin = False
'		Security.CurrentUserID = 1288
'		Response.Write "PVID:" & Request.QueryString("PVID")
'		Response.Write "<BR>"
'		Response.Write "SID:" & Request.QueryString("SID")
'		Response.Write "<BR>"
'		Response.Write "UID:" & Security.CurrentUserID
'		Response.Write "<BR>"
'		Response.Write "PM:" & Security.IsProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write "SEPM:" & Security.IsSysEngProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write "SM:" & Security.IsSystemTeamLead(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If
	
	Select Case CLng(Security.CurrentUserID)
		Case 8
			m_IsSysAdmin = True
		Case 31
			m_IsSysAdmin = True
		Case 1396
			m_IsSysAdmin = True
		Case Else
			m_IsSysAdmin = False
	End Select
	
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(Request("PVID"))
	m_IsSysTeamLead = Security.IsSystemTeamLead(Request("PVID"))
	m_IsProgramManager = Security.IsProgramManager(Request("PVID"))
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	
	Set Security = Nothing

	If Not (m_IsSysAdmin Or m_IsSysTeamLead Or m_IsSysEngProgramManager Or m_IsProgramManager Or m_IsSEPMProductsEditor) Then
		Response.Write "<H3>Insuficient User Privileges</H3><H4>Access Denied</H4><p>This operation is reserved for the System Manager & SE PM</p>"
		Response.End
	End If
'##############################################################################	

	m_ScheduleID = Request("SID")
	
	If m_ScheduleID = "" Then
		Response.Write "Insufficient Data to process request"
		Response.End
	End If 
	
	Call DeactivateSchedule()
End Sub

Sub DeactivateSchedule()
	Dim dw, cn, cmd, iRowsChanged
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans
	Set cmd = server.CreateObject("adodb.command")
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateScheduleStatus")
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, m_ScheduleID
	dw.CreateParameter cmd, "@p_ActiveYN", adChar, adParamInput, 1, "N"
	iRowsChanged = dw.ExecuteNonQuery(cmd)
	
	cn.CommitTrans
	Response.Write "<input type=hidden id=CloseOnLoad value=true>"

	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
	
End Sub
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Delete Schedule</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

    function window_onLoad(pulsarplusDivId) {
    if (window.CloseOnLoad) {
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Reload PulsarPlusPmView Tab
            parent.window.parent.RemoveScheduleResultPulsarPlus(PVID.value);
            parent.window.parent.reloadFromPopUp(pulsarplusDivId);
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (window.location != window.parent.location) {
                window.parent.RemoveScheduleResult(1);
                setTimeout(function () {
                    window.parent.modalDialog.cancel();
                }, 150);
            } else {
                window.close();
            }
        }
    }
}
//-->
</script>
</head>
<body bgcolor="ivory" leftMargin="9" topMargin="9" OnLoad="window_onLoad('<%= Request("pulsarplusDivId")%>')">
    
<% Call Main() %>
    <span><strong>Deactivating Schedule...</strong></span>
    <input type="hidden" id="PVID" name="PVID" value="<%=Request("PVID")%>">
</body>
</html>
