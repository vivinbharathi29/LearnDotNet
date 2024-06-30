<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
}

function cboRoot_onchange() {
	if (cboRoot.selectedIndex == 0)
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&ActionID=" + txtActionID.value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value ;
	else
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + cboRoot.options[cboRoot.selectedIndex].value  + "&ActionID=" + txtActionID.value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value;
}

function cboVersion_onchange() {
	if (cboVersion.selectedIndex == 0)
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&ActionID=" + txtActionID.value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value ;
	else
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&VersionID=" + cboVersion.options[cboVersion.selectedIndex].value  + "&ActionID=" + txtActionID.value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value;
}

function cboAction_onchange() {
	if (cboAction.selectedIndex == 0)
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&VersionID=" +  txtVersionID.value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value ;
	else
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&VersionID=" +  txtVersionID.value + "&ActionID=" + cboAction.options[cboAction.selectedIndex].value + "&UserID=" + txtUserID.value + "&TypeID=" + txtTypeID.value;
}
function cboUser_onchange() {
	if (cboUser.selectedIndex == 0)
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&VersionID=" +  txtVersionID.value + "&ActionID=" + txtActionID.value + "&TypeID=" + txtTypeID.value ;
	else
		window.location.href = "DeliverableChanges.asp?ProductID=" + txtProductID.value + "&RootID=" + txtRootID.value + "&VersionID=" +  txtVersionID.value + "&ActionID=" + txtActionID.value + "&UserID=" + cboUser.options[cboUser.selectedIndex].value + "&TypeID=" + txtTypeID.value;
}
//-->
</SCRIPT>
<html>

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
<link href="../style/Excalibur.css" type="text/css" rel="stylesheet">

<body LANGUAGE=javascript onload="return window_onload()" bgColor=white>
<%

	dim cn
	dim cm
	dim p
	dim rs
	dim strError
	dim strProdName
	dim strPOR
	dim CurrentUserPartner
	dim strVersions
	dim RowCount
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	if request("ProductID") = "" or not isnumeric(request("ProductID")) then
		strError = "No Product Specified."
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
	
	
		strError = ""
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersionName"
		
	
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProductID")
		cm.Parameters.Append p
	
	
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
		
		'rs.Open "spGetProductVersionName " & request("ProductID"),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strError = "Could not find the selected product."
		else
			strProdName = rs("Name") & ""
			strPOR = rs("PDDReleased") & ""


			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					set rs = nothing
					set cn=nothing
					
					Response.Redirect "../NoAccess.asp?Level=0"
				end if
			end if			
		
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
	else
%>


<center><font face=verdana size=4><b><%=strProdName%> Deliverable Matrix Changes</b></font>
	<font face=verdana size=2><br><br><%=now()%><BR><BR></font>
</center>

<%
	dim strRoots
	strRoots = "<option selected value="""">All Deliverables</option>"

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListDeliverableMatrixUpdateRoots"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(request("ProductID"))
	cm.Parameters.Append p

	if request("TypeID") <> "" then
		Set p = cm.CreateParameter("@TypeID", 3, &H0001)
		p.Value = clng(request("TypeID"))
		cm.Parameters.Append p
	end if

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	do while not rs.EOF
		if trim(request("RootID")) = trim(rs("ID")) then
			strRoots = strRoots & "<option selected value=" & rs("ID") & ">" & rs("name") & "</option>"
		else
			strRoots = strRoots & "<option value=" & rs("ID") & ">" & rs("name") & "</option>"
		end if
		rs.Movenext
	loop
	rs.Close
	

	strVersions = "<option selected value=0>All Versions</option>"
	strActions = "<option selected value="""">All Changes</option>"
	strUsers = "<option selected value=0>All People</option>"

	rs.Open "spListHistoryChangeTypes " & clng(request("ProductID")),cn,adOpenStatic
	do while not rs.EOF
		if trim(request("ActionID")) = trim(rs("ID")) then
			strActions = strActions & "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>"
		else
			strActions = strActions & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
		end if
		rs.MoveNext
	loop
	rs.Close
	
'	rs.Open "spListProductDeliverableMatrixUpdaters " & clng(request("ProductID")),cn,adOpenStatic
'	do while not rs.EOF
'		if trim(request("UserID")) = trim(rs("ID")) then
'			strUsers = strUsers & "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>"
'		else
'			strUsers = strUsers & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"
'		end if
'		rs.MoveNext
'	loop
'	rs.Close
'
	if trim(request("RootID"))<>"" then
		rs.Open "spListDeliverableVersions4Product " & clng(request("RootID")) & "," & clng(request("ProductID")),cn,adOpenStatic
		do while not rs.EOF
			strVersion = rs("VersionID") & ": " & rs("Version")
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")
			end if

			if trim(rs("Partnumber") &"") <> "" then
				if trim(rs("VersionID")) & "" = trim(request("VersionID")) & "" then	
					strVersions = strVersions & "<option selected value=" & rs("VersionID") & ">" & strVersion & "[" & rs("Partnumber") & "]</option>"
				else
					strVersions = strVersions & "<option value=" & rs("VersionID") & ">" & strVersion & "&nbsp;&nbsp;[" & rs("Partnumber") & "]</option>"
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
	end if	
	
	
%>
<TABLE class="DisplayBar" width=100%>
	<TR>
	<TD valign=top><table><tr><td valign=top class="DisplayTitle"><font size=2 face=verdana>Display:&nbsp;&nbsp;</font></td></tr></table></TD>
	<TD  width=100%>
		<TABLE width=100%>
			<tr>
				<td><font size=2 face=verdana><b>Deliverables:&nbsp;</b></font></td><td colspan=3><SELECT style="width:100%" id=cboRoot name=cboRoot LANGUAGE=javascript onchange="return cboRoot_onchange()"><%=strRoots%></SELECT></td>
			</tr>			
			<TR>
			<TD><font size=2 face=verdana><b>Version:&nbsp;</b></font></td><td width=100%><SELECT id=cboVersion name=cboVersion style="width:100%" LANGUAGE=javascript onchange="return cboVersion_onchange()"><%=strVersions%></SELECT></td>
			<TD><font size=2 face=verdana><b>Action:&nbsp;</b></font></td><td><SELECT id=cboAction name=cboAction LANGUAGE=javascript onchange="return cboAction_onchange()"><%=strActions%></SELECT></td>
			<TD style="display:none"><font size=2 face=verdana><b>Updated&nbsp;By:&nbsp;</b></font></td><td style="display:none"><SELECT id=cboUser name=cboUser LANGUAGE=javascript onchange="return cboUser_onchange()"><%=strUsers%></SELECT></td>
			</TR>
		</TABLE>
	</TD>
</TR></TABLE> <BR>


<TABLE width="100%" border=1 bgColor=ivory style="FONT-SIZE: xx-small; FONT-FAMILY: Verdana" bordercolor=tan cellpadding=1 cellspacing=0>
<%
	dim strDeliverable
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListProductDeliverableHistory"
	
	Set p = cm.CreateParameter("@ProductID", 3, &H0001)
	p.Value = clng(request("ProductID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@RootID", 3, &H0001)
	if trim(request("RootID")) = "" or trim(request("RootID")) = "0" then 
		p.Value = 0
	else
		p.Value = clng(request("RootID"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID", 3, &H0001)
	if trim(request("VersionID")) = "" or trim(request("VersionID")) = "0" then 
		p.Value = 0
	else
		p.Value = clng(request("VersionID"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TypeID", 3, &H0001)
	if trim(request("TypeID")) = "" then
		p.value = null
	else
		p.Value = clng(request("TypeID"))
	end if
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


	if rs.EOF and rs.BOF then
			Response.Write "<TR><TD><font size=2 face=verdana>No Deliverable Matrix changes have been made which match the selected filters.</font></td></tr>"
	else
		if trim(request("TypeID")) = "1" or trim(request("TypeID")) = "" then
			Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></td><TD><font size=1 face=verdana><b>Deliverable</b></font></td><TD><font size=1 face=verdana><b>Model</b></font></td><TD><font size=1 face=verdana><b>Part</b></font></td><TD><font size=1 face=verdana><b>Vendor</b></font></td><TD><font size=1 face=verdana><b>Action</b></font></td><TD><font size=1 face=verdana><b>Details</b></font></td><TD nowrap><font size=1 face=verdana><b>Updated By</b></font></td><TD><font size=1 face=verdana><b>Updated</b></font></td></tr>"
		else
			Response.Write "<TR bgcolor=cornsilk><TD><font size=1 face=verdana><b>ID</b></font></td><TD><font size=1 face=verdana><b>Deliverable</b></font></td><TD><font size=1 face=verdana><b>Action</b></font></td><TD><font size=1 face=verdana><b>Details</b></font></td><TD nowrap><font size=1 face=verdana><b>Updated By</b></font></td><TD><font size=1 face=verdana><b>Updated</b></font></td></tr>"
		end if
		RowCount=0
		do while not rs.eof
			if trim(request("RootID")) = trim(rs("RootID")) or request("RootID") = "" then
    			RowCount = RowCount + 1
    			strDeliverable = rs("Deliverable") & " " & rs("Version")
	    		if rs("revision") <> "" then
		    		strDeliverable = strDeliverable & "," & rs("Revision")
			    end if
			    if rs("pass") <> "" then
    				strDeliverable = strDeliverable & "," & rs("Pass")
			    end if
			    if rs("VersionID") <> 0 then
    				Response.Write "<TR><TD>" & rs("VersionID") & "</TD>"
			    else 
    				Response.Write "<TR><TD>" & rs("RootID") & "</TD>"
			    end if
			    Response.Write "<TD>" & strDeliverable & "</TD>"
			    if trim(request("TypeID")) = "1" or trim(request("TypeID")) = "" then
    				Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
				    Response.Write "<TD>" & rs("PartNumber") & "&nbsp;</TD>"
				    Response.Write "<TD>" & rs("Vendor") & "&nbsp;</TD>"
			    end if
			    Response.Write "<TD>" & rs("Action") & "</TD>"
			    Response.Write "<TD>" & rs("Details") & "&nbsp;</TD>"
			    Response.Write "<TD>" & rs("Username")&"" & "</TD>"
			    Response.Write "<TD>" & rs("Updated") & "</TD></TR>"
			end if
			rs.MoveNext
		loop
	end if
	rs.close

	
	end if

	cn.Close
	set rs = nothing
	set cn = nothing
	
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
		
%>
</TABLE>
<%
	if RowCount > 0 then
		Response.Write "<BR><BR><BR>Rows Displayed: " & RowCount
	end if

%>
	<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
	<INPUT type="hidden" id=txtTypeID name=txtTypeID value="<%=request("TypeID")%>">
	<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
	<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
	<INPUT type="hidden" id=txtActionID name=txtActionID value="<%=request("ActionID")%>">
	<INPUT type="hidden" id=txtUserID name=txtUserID value="<%=request("UserID")%>">
</body>
</html>
 
