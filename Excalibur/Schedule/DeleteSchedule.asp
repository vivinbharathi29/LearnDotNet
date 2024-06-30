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
    m_IsSysAdmin = False
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_UserFullName
Dim m_EditModeOn
Dim m_ScheduleID
Dim m_ScheduleDescription
Dim p
dim CurrentDomain
dim Currentuser
dim CurrentuserID
dim rs

Sub Main()
    'Get User
	
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	m_ScheduleID = Request("SID")	
	Call DeleteSchedule()
End Sub

Sub DeleteSchedule()
	Dim dw, cn, cmd, iRowsChanged
	
    	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    set rs = server.CreateObject("ADODB.recordset")
    set cmd = server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = cn
	cmd.CommandType = 4
	cmd.CommandText = "spGetUserInfo"
	

	Set p = cmd.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cmd.Parameters.Append p

	Set p = cmd.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cmd.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cmd.Execute 
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		if rs("PulsarSystemAdmin") = 1 then
            m_IsSysAdmin = True
        end if
        CurrentuserID = rs("ID")
	end if
	rs.Close

    If Not m_IsSysAdmin Then
        set rs = server.CreateObject("ADODB.recordset")
        set cmd = server.CreateObject("ADODB.Command")
	    Set cmd.ActiveConnection = cn
	    cmd.CommandType = 4
	    cmd.CommandText = "usp_USR_GetPermissions"
	

	    Set p = cmd.CreateParameter("@p_intUserID", adInteger, &H0001)
	    p.Value = CurrentuserID
	    cmd.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cmd.Execute 
	
        do while not rs.EOF
		    if trim("Schedule.Delete") = trim(rs("PermissionName")) then
			    m_IsSysAdmin = True
		    end if
		    rs.MoveNext
        loop
	    rs.Close

        If Not m_IsSysAdmin Then
		    Response.Write "<H3>Insuficient User Privileges</H3><H4>Access Denied</H4>"
		    Response.End
        End If
	End If

	cn.BeginTrans
	Set cmd = server.CreateObject("adodb.command")
	Set cmd = dw.CreateCommandSP(cn, "usp_DeleteSchedule")
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, m_ScheduleID
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
    <span><strong>Deleting Schedule...</strong></span>
    <input type="hidden" id="PVID" name="PVID" value="<%=Request("PVID")%>">
</body>
</html>