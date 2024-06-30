<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<TITLE>Reassign Preinstall Ownership</TITLE>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.parent.close();
}

function cmdOK_onclick() {

	if (Reassign.txtNotes.value == "" && Reassign.chkUrgent.checked)
		{
			alert("You must enter notes for Urgent requests.");
			Reassign.txtNotes.focus();
			return;
		}
	Reassign.txtName.value = Reassign.cboOwner.options[Reassign.cboOwner.selectedIndex].text;
	Reassign.submit();
}

function chkUrgent_onclick() {
	if (Reassign.chkUrgent.checked)
		{
		ReqNotes.style.display = "";
		Reassign.txtNotes.focus();
		}
	else
		ReqNotes.style.display = "none";
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<%


if (request("ProdID") = "" or request("VersionID") = "") and request("PIAlerts") = "" then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
	dim strPreinstallStatus
	dim cn
	dim rs
	dim cm
	dim p
	dim strEmployeeList
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserGroup
	dim strNotes
	dim strUrgent
	dim strShowNotes
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
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
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserGroup = rs("WorkgroupID") & ""
	end if
	rs.Close
	
	strPart = ""
	strNotes = ""
	strUrgent = ""
	strShowNotes = "none"
	
	if instr(request("PIAlerts"),",") > 0 then 
		'Multi-select w/ multi-items
		strPreinstallStatus = 0
		strID = request("PIAlerts")
	else
		if request("PIAlerts") <> "" then 
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetPreinstallStatus"
		
			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			p.Value = left( request("PIAlerts"),instr(request("PIAlerts"),":")-1)
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			p.Value = mid( request("PIAlerts"),instr(request("PIAlerts"),":")+1)
			cm.Parameters.Append p
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
		
'			rs.Open "spGetPreinstallStatus " & left( request("PIAlerts"),instr(request("PIAlerts"),":")-1) & "," & mid( request("PIAlerts"),instr(request("PIAlerts"),":")+1),cn,adOpenForwardOnly
		else
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetPreinstallStatus"
		
			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			p.Value = request("ProdID")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@ProdID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
					
			'rs.Open "spGetPreinstallStatus " & request("ProdID") & "," & request("VersionID"),cn,adOpenForwardOnly
		end if
		strPreinstallStatus =  ""
		if not (rs.EOF and rs.BOF) then
			strPreinstallStatus = rs("PreinstallStatus") & ""
			strPart = rs("PartNumber") & ""
			strID = rs("ID") & ""
			strNotes = rs("PreinstallNotes") & ""
			strUrgent = replace(replace(rs("Urgent") & "","True","checked"),"False","")
			if strUrgent="checked" then
				strShowNotes = ""
			end if
		end if
		rs.Close
	end if
	
	
	if trim(request("FunctionID")) = "4" then
		'strEmployeeList = "<option selected value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option selected value=-1>Database Team</option>"
	elseif trim(request("FunctionID")) = "5" then
		'strEmployeeList = "<option selected value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option selected value=-4>Imported</option>"
	elseif trim(request("FunctionID")) = "6" then
		'strEmployeeList = "<option selected value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option selected value=0>Cancel Request</option>"
	elseif trim(strPreinstallStatus) = "-1" then
		strEmployeeList = "<option value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option selected value=-1>Database Team</option>"
	elseif trim(strPreinstallStatus) = "-4"  then
		strEmployeeList = "<option value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option value=-1>DB Team - Import</option>"	
		strEmployeeList = strEmployeeList & "<option selected value=-4>Database Team</option>"	
	elseif (trim(request("FunctionID")) = "2") then 'or trim(request("FunctionID")) = "1")  then
		strEmployeeList = "<option value=0>New Request</option>"
	else
		strEmployeeList = "<option selected value=0>New Request</option>"
		strEmployeeList = strEmployeeList & "<option value=-1>Database Team</option>"	
	end if	
	
'	if trim(request("FunctionID")) = "1" then
'		if trim(CurrentUserGroup) = "15" then
'			strEmployeeList = strEmployeeList & "<option value=-2>TDC Preinstall Team</option>"	
'		elseif trim(CurrentUserGroup) = "22" then
'			strEmployeeList = strEmployeeList & "<option value=-3>Houston Preinstall Team</option>"	
'		end if
'	end if 
		
	if request("FunctionID") <> "4" and request("FunctionID") <> "5"  and request("FunctionID") <> "6" then 
		rs.Open "spGetEmployees",cn,adOpenForwardOnly
			do while not rs.EOF
				if trim(rs("WorkgroupID")) = trim(CurrentUserGroup) and rs("Active") and rs("Division") = 1 and rs("ID") <> 646 and rs("ID") <> 791 and rs("ID") <> 903 then
					if trim(strPreinstallStatus) = trim(rs("ID")) then
						strEmployeeList = strEmployeeList & "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>"		
					elseif rs("Active") = 1 then
						strEmployeeList = strEmployeeList & "<option value=" & rs("ID") & ">" & rs("Name") & "</option>"		
					end if
				end if
				rs.MoveNext
			loop
		rs.Close
	end if	
%>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<form ID=Reassign method=post action="PreinstallReassignSave.asp">
<%if CurrentUserGroup = 22 then%>
	<INPUT type="hidden" id=txtUpdateGroup name=txtUpdateGroup value=2>
<%else%>
	<INPUT type="hidden" id=txtUpdateGroup name=txtUpdateGroup value=1>
<%end if%>
<INPUT type="hidden" id=txtID name=txtID value="<%=strID%>">
<INPUT type="hidden" id=txtName name=txtName value="">
<table>
<%if request("FunctionID") = "1"  then%>
	<TR><TD><font size=3 face=verdana><B>Assign Owner:</b></font>&nbsp;
<%elseif request("FunctionID") = "2" then%>
	<TR><TD><font size=3 face=verdana><B>Reassign:</b></font>&nbsp;
<%elseif request("FunctionID") = "4" then%>
	<TR><TD><font size=3 face=verdana><B>Release to DB Team:</b></font>&nbsp;
<%elseif trim(request("FunctionID")) = "5" then%>
	<TR><TD><font size=3 face=verdana><B>Update Status:</b></font>&nbsp;
<%elseif trim(request("FunctionID")) = "6" then%>
	<TR><TD><font size=3 face=verdana><B>Cancel Request:</b></font>&nbsp;
<%else%>
	<TR><TD><font size=3 face=verdana><B>Reassign:</b></font>&nbsp;
<%end if%>

<%if request("VersionID") = "" and instr(request("PIAlerts"),",")>0 then%>
	<font size=1 color=blue>(All Selected Deliverables)</font>
<%end if%>
</td></tr><TR><TD>
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<%if trim(request("FunctionID")) = "5"  or trim(request("FunctionID")) = "6" then%>
			<td width="150" nowrap><b>Preinstall Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<%else%>
			<td width="150" nowrap><b>Preinstall Owner:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<%end if%>
		<td width="100%">
			<SELECT id=cboOwner name=cboOwner style="WIDTH:180">
				<%=strEmployeeList%>
			</SELECT>
		</td>
	</tr>

<%	if instr(request("PIAlerts"),",") = 0  then'request("PIAlerts")="" then%>
	
	<tr>
		<td width="150" nowrap><b>Part&nbsp;Number:</b>&nbsp;</td>
		<td width="100%">
			<INPUT type="text" id=txtPart style="WIDTH:100%" name=txtPart value="<%=strPart%>" maxlength=50>
		</td>
	</tr>
	<tr>
		<td width="150" nowrap><b>Priority:</b>&nbsp;</td>
		<td width="100%"><INPUT <%=strUrgent%> type="checkbox" id=chkUrgent name=chkUrgent LANGUAGE=javascript onclick="return chkUrgent_onclick()">&nbsp;Urgent
		</td>
	</tr>
	<tr>
		<td width="150" nowrap valign=top><b>Notes:</b><span style="Display:<%=strShowNotes%>" ID=ReqNotes>&nbsp;<font color="red" size="1">*</font></span>&nbsp;</td>
		<td width="100%"><TEXTAREA rows=3 cols=20 id=txtNotes name=txtNotes><%=strNotes%></TEXTAREA>
		</td>
	</tr>
<%else%>
<TEXTAREA style="Display:none" rows=3 cols=20 id=txtNotes name=txtNotes><%=strNotes%></TEXTAREA>
<INPUT style="Display:none"  <%=strUrgent%> type="checkbox" id=chkUrgent name=chkUrgent LANGUAGE=javascript onclick="return chkUrgent_onclick()">
<INPUT style="Display:none" type="text" id=txtPart style="WIDTH:100%" name=txtPart value="<%=strPart%>">
<%end if%>
	
</table>
</td></tr>
<TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>

</table>
</form>

<%	if instr(request("PIAlerts"),",") > 0  then ' then%>
	<font face=verdana size=1>Note: You must select only one request to assign a priority, comments, or a part number.</font>
<%	end if
	
	set rs = nothing
	set cn = nothing
end if


%>

</BODY>
</HTML>
