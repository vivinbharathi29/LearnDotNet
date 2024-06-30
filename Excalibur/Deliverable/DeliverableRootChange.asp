
<%@ Language=VBScript %>
<%
    Server.ScriptTimeout = 600
%>

<html>
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
<STYLE>

.Details TABLE
{
    BORDER-RIGHT-STYLE: thin;
    BORDER-TOP-STYLE: thin;
    BORDER-LEFT-STYLE: thin;
    BORDER-BOTTOM-STYLE: thin;
}
.Details THEAD
{
    BACKGROUND-COLOR: gainsboro;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
	FONT-WEIGHT: bold;
}
.Details TBody
{
    BACKGROUND-COLOR: white;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
}
.Details TD
{
    BORDER-RIGHT-STYLE: thin;
    BORDER-TOP-STYLE: thin;
    BORDER-LEFT-STYLE: thin;
    BORDER-BOTTOM-STYLE: thin;
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
}



</STYLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
}

function cboVersion_onchange() {
	window.location.href = "DeliverableRootChange.asp?ID=" + txtRootID.value + "&VersionID=" + cboVersion.options[cboVersion.selectedIndex].value + "&ActionID=" + txtActionID.value;
}

function cboAction_onchange() {
	window.location.href = "DeliverableRootChange.asp?ID=" + txtRootID.value + "&VersionID=" + txtVersionID.value + "&ActionID=" + cboAction.options[cboAction.selectedIndex].value ;
}

//-->
</SCRIPT>
<link href="../style/Excalibur.css" type="text/css" rel="stylesheet">
</head>
<body LANGUAGE=javascript onload="return window_onload()" bgColor=white>
<%

	dim cn
	dim cm
	dim p
	dim rs
	dim strError
	dim strFileame
	dim strDeliverableName
	dim strVersion
	dim strType
	dim strPOR
	dim CurrentUserPartner
	dim RowCount
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	set rs2 = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.IsolationLevel=256
	cn.Open

	if request("ID") = "" or not isnumeric(request("ID")) then
		strError = "No Deliverable Specified."
	else
		'Get User
		dim CurrentDomain
		CurrentUser = lcase(Session("LoggedInUser"))
	
		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
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

		if (rs.EOF and rs.BOF) then
			set rs = nothing
			set cn=nothing
			Response.Redirect "../NoAccess.asp?Level=0"
		else
			CurrentUserPartner = rs("PartnerID")
		end if 
		rs.Close	
	end if
	
'	if trim(strPOR) = "" then
'		strError = "No changes are logged before Product POR."
'	else
'		strPOR = formatdatetime(strPOR ,vbshortdate)	
'	end if
	if strError <> "" then
		Response.Write "<font size=2 face=verdana><b>" & strError & "</b></font>"
	end if

	strSQL = "spGetDelPropSummary " & clng(request("ID"))
	rs.Open strSQL,cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "Unable to find the selected deliverable."
	else
		strFilename = rs("Filename") & ""
		strType = rs("TypeID") & ""
		strDeliverableName = rs("Name") & ""		
	end if
	rs.Close

%>

<center><font face=verdana size=4><b><%=strDeliverableName%> Change History</b></font>
	<font face=verdana size=2><br><br><%=now()%><BR><BR></font>
</center>

<%
	dim strVersions
	dim strActions
	strVersions = "<option selected value=0>All Versions</option>"
	strActions = "<option selected value="""">All Changes</option>"
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListDeliverableVersions"
	

	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	p.Value = clng(request("ID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ActiveOnly", 3, &H0001)
	p.Value = 0
	cm.Parameters.Append p
	
	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	
	do while not rs.EOF
		strVersion = rs("VersionID") & ": " & rs("Version")
		if rs("Revision") & "" <> "" then
			strVersion = strVersion & "," & rs("Revision")
		end if
		if rs("Pass") & "" <> "" then
			strVersion = strVersion & "," & rs("Pass")
		end if

		if trim(strType) = "1" and  strFilename <> "HFCN" then
			if trim(rs("VersionID")) & "" = trim(request("VersionID")) & "" then	
				strVersions = strVersions & "<option selected value=" & rs("VersionID") & ">" & strVersion & "[" & rs("Partnumber") & "]</option>"
			else
				strVersions = strVersions & "<option value=" & rs("VersionID") & ">" & strVersion & "&nbsp;&nbsp;[" & rs("Partnumber") & "]</option>"
			end if				
		elseif strFilename = "HFCN" then
			if trim(rs("VersionID")) & "" = trim(request("VersionID")) & "" then		
				strVersions = strVersions & "<option selected value=" & rs("VersionID") & ">" & strVersion & "</option>"
			else
				strVersions = strVersions & "<option value=" & rs("VersionID") & ">" & strVersion & "</option>"
			end if				
		else
			if trim(rs("VersionID")) & "" = trim(request("VersionID")) & "" then		
				strVersions = strVersions & "<option selected value=" & rs("VersionID") & ">" & strVersion & "</option>"
			else
				strVersions = strVersions & "<option value=" & rs("VersionID") & ">" & strVersion & "</option>"
			end if				
		end if
			
		rs.Movenext
	loop
	rs.Close
	
	rs.Open "spListHistoryChangeTypes",cn,adOpenStatic
	do while not rs.EOF
		if trim(request("ActionID")) = trim(rs("ID")) then
			strActions = strActions & "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>"
		else
			strActions = strActions & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
		end if
		rs.MoveNext
	loop
	rs.Close

%>

<TABLE class="DisplayBar" width=100%>
	<TR>
	<TD class="DisplayTitle"><font size=2 face=verdana>Display:&nbsp;&nbsp;</font></TD>
	<TD><font size=2 face=verdana><b>Version:&nbsp;</b><SELECT id=cboVersion name=cboVersion LANGUAGE=javascript onchange="return cboVersion_onchange()"><%=strVersions%></SELECT></font>&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<TD width=100%><font size=2 face=verdana><b>Action:&nbsp;</b><SELECT id=cboAction name=cboAction LANGUAGE=javascript onchange="return cboAction_onchange()"><%=strActions%></SELECT></font></td>
</TR></TABLE> <BR>
<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%
	dim strDeliverable
	dim strRow

'	Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></td><TD><font size=1 face=verdana><b>Deliverable Version</b></font></td><TD><font size=1 face=verdana><b>Vendor</b></font></td><TD><font size=1 face=verdana><b>Product</b></font></td><TD><font size=1 face=verdana><b>Action</b></font></td><TD><font size=1 face=verdana><b>Details</b></font></td><TD nowrap><font size=1 face=verdana><b>Updated By</b></font></td><TD><font size=1 face=verdana><b>Updated</b></font></td></tr>"
	if trim(strType) = "1" then
		Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></td><TD><font size=1 face=verdana><b>Deliverable</b></font></td><TD><font size=1 face=verdana><b>HW/FW/REV</b></font></td><TD><font size=1 face=verdana><b>Vendor</b></font></td><TD><font size=1 face=verdana><b>Part</b></font></td><TD><font size=1 face=verdana><b>Model</b></font></td><TD><font size=1 face=verdana><b>Product</b></font></td><TD><font size=1 face=verdana><b>Action</b></font></td><TD><font size=1 face=verdana><b>Details</b></font></td><TD nowrap><font size=1 face=verdana><b>Updated By</b></font></td><TD><font size=1 face=verdana><b>Updated</b></font></td></tr>"
	else
		Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></td><TD><font size=1 face=verdana><b>Deliverable</b></font></td><TD><font size=1 face=verdana><b>Version</b></font></td><TD><font size=1 face=verdana><b>Vendor</b></font></td><TD><font size=1 face=verdana><b>Product</b></font></td><TD><font size=1 face=verdana><b>Action</b></font></td><TD><font size=1 face=verdana><b>Details</b></font></td><TD nowrap><font size=1 face=verdana><b>Updated By</b></font></td><TD><font size=1 face=verdana><b>Updated</b></font></td></tr>"
	end if

	strRow = ""
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListDeliverableHistory"
	
	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	p.Value = clng(request("ID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID", 3, &H0001)
	p.Value = clng(request("VersionID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ActionID", 3, &H0001)
	if trim(request("ActionID")) = "" then
		p.value = null
	else
		p.Value = clng(request("ActionID"))
	end if
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	RowCount = 0
	do while not rs.eof
		RowCount = RowCount + 1
		if trim(rs("VersionID") & "") = "0" or trim(rs("VersionID") & "") = "" then
			Response.Write "<TR><TD>" & rs("RootID") & "</TD>" 
		else
			Response.Write "<TR><TD>" & rs("VersionID") & "</TD>" 
		end if
		Response.Write "<TD>" & rs("Deliverable") & "</TD>"
		if trim(rs("VersionID") & "") = "0" or trim(rs("VersionID") & "") = "" then
			Response.Write "<TD>[ROOT]</TD>"
		else
			if rs("Typeid") = 1 then
				strDeliverable =  rs("Version")
				if rs("revision") <> "" then
					strDeliverable = strDeliverable & "," & rs("Revision")
				end if
				if rs("pass") <> "" then
					strDeliverable = strDeliverable & "," & rs("Pass")
				end if
				Response.Write "<TD>" & strDeliverable & "</TD>"
			else
				strDeliverable = rs("Version")
				if rs("revision") <> "" then
					strDeliverable = strDeliverable & "," & rs("Revision")
				end if
				if rs("pass") <> "" then
					strDeliverable = strDeliverable & "," & rs("Pass")
				end if
				Response.Write "<TD>" & strDeliverable & "</TD>"
			end if
		end if
		Response.Write "<TD>" & rs("Vendor") & "&nbsp;</TD>"
		if rs("Typeid") = 1 then
			Response.Write "<TD nowrap>" & rs("PartNumber") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
		end if
		Response.Write "<TD>" & rs("Product") & "&nbsp</TD>"
		if rs("Action") = "Deliverables Updated" and (trim(rs("VersionID") & "") = "0" or trim(rs("VersionID") & "") = "") then
		    Response.Write "<TD>Root Deliverable Updated</TD>"
        else
		    Response.Write "<TD>" & rs("Action") & "</TD>"
		end if
		Response.Write "<TD>" & rs("Details") & "&nbsp;</TD>"
		Response.Write "<TD>" & rs("Username") & "</TD>"
		Response.Write "<TD>" & rs("Updated") & "</TD></TR>"

		rs.MoveNext
	loop
	rs.close
	
	Response.Write "</table>"

	if RowCount > 0 then
		Response.Write "<BR><BR><BR>Rows Displayed: " & RowCount
	end if

	cn.Close
	set rs = nothing
	set cn = nothing

    dim strVersionID
    if request("VersionID") = "" then
        strVersionID = "0"
    else
        strVersionID = request("VersionID")
    end if

%>
</TABLE>

<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=strVersionID%>">
<INPUT type="hidden" id=txtActionID name=txtActionID value="<%=request("ActionID")%>">
</body>
</html>
