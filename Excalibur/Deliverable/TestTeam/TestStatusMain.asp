<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function CheckTextSize(field, maxLength) {
	if (field.value.length > maxLength + 1)
		{
		field.value = field.value.substring(0, maxLength);
		alert("The maximum size of this field in 200 characters. You input has been truncated.");
		}
	else if (field.value.length >= maxLength)
		{
		window.event.keyCode=0;
		field.value = field.value.substring(0, maxLength);
		}
} 

function cboStatus_onclick() {
	if (frmMain.cboStatus.selectedIndex==2 || frmMain.cboStatus.selectedIndex==3)
		RequireNotes.style.display="";
	else
		RequireNotes.style.display="none";
}

function window_onload() {
	frmMain.cboStatus.focus();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<LINK href="../../style/wizard style.css" type=text/css rel=stylesheet >
<font size=3 face=verdana><b>
<%
	if request("FieldID") = "2" then
		response.write "Update ODM HW Test Status</b><BR><BR></font>"
	elseif request("FieldID") = "3" then
		response.write "Update COMM Test Status</b><BR><BR></font>"
    elseif request("FieldID") = "4" then
		response.write "Update DEV Test Status</b><BR><BR></font>"
	else
		response.write "Update SE Test Status</b><BR><BR></font>"
	end if

	dim cn 
	dim rs
	dim strName
	dim strVersion
	dim strRevision
	dim strPass
	dim strTypeID
	dim strModelNumber
	dim strPartNumber
	dim strEOLDate
	dim strVendor
	dim strPMEmail
	dim strStatus
	dim strUnitsReceived
	dim strTestNotes
	dim DisplayNotesRequired
	dim OTSRootText
	dim OTSVersionText
	dim OTSRootCount
	dim OTSVersionCount
	dim strRootID
	dim strOTSNumbers
  	dim CurrentUserPartner
  	dim strProductPartner

	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	
	
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("email") & ""
		CurrentUserPartner = rs("PartnerID") & ""
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
		CurrentUserPartner = 0
	end if
	rs.Close

	rs.Open "spGetProductVersion " & clng(request("ProductID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
  		strProductPartner = "0"
	else
  		strProductPartner = rs("PartnerID") & ""
	end if
	rs.Close


	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(strProductPartner) <> trim(CurrentUserPartner) or trim(CurrentUserPartner) = "0" then
			set rs = nothing
			set cn=nothing
				
			'Response.Redirect "../../NoAccess.asp?Level=1"
		end if
	end if
	
	
	
	
	rs.Open "spGetVersionProperties4Web " & clng(request("VersionID")),cn,adOpenForwardOnly
	if not(rs.EOF and rs.BOF) then
		strName = trim(rs("DeliverableName") & "")
		strDelID = "<a target=_blank href=""../../Query/DeliverableVersionDetails.asp?Type=1&RootID=" & rs("RootID") & "&ID=" & rs("VersionID") & """>" & rs("versionID") & "</a>"
		strVersion = trim(rs("Version") & "")
		strRootID = rs("RootID") & ""
		strRevision = trim(rs("Revision") & "")
		strPass = trim(rs("Pass") & "")
		strTypeID = trim(rs("TypeID") & "")
		strModelNumber = trim(rs("ModelNumber") & "")
		strPartNumber = trim(rs("PartNumber") & "")
		strEOLDate = rs("EOLDate") & ""
		if trim(rs("VersionVendor") & "") <> "" then
			strVendor = rs("VersionVendor") & ""
		else
			strVendor = rs("Vendor") & ""
		end if
	end if
	rs.Close
	
	rs.Open "spGetCategoryPM " & clng(request("VersionID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strPMEmail=""
	else
		strPMEmail = rs("Email") & ""
	end if
	rs.Close
	



	rs.Open "spGetTestLeadStatus " & clng(request("ProductID")) & "," & clng(request("VersionID")) & "," & clng(request("FieldID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strProductName =""
		strStatus = ""
		strUnitsReceived = ""
		strTestNotes = ""
	else
		strProductName =rs("Product") & ""
		strStatus = rs("TestStatus") & ""
		strUnitsReceived = rs("UnitsReceived") & ""
		strTestNotes = rs("TestNotes") & ""
	end if
	rs.Close
	
	if strStatus = "2" or strStatus = "3" then
		DisplayNotesRequired = ""
	else
		DisplayNotesRequired = "none"
	end if 
	
	set rs = nothing
	cn.Close
	set cn = nothing

	if strProductName = "" or strName = "" then
		Response.write "<font size=2 face=verdana>Deliverable not found. (" & request("ID") & ")</font>"	
	else

%>
<form ID=frmMain action="TestStatusSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Deliverable:</b></TD>
	<TD width="100%" colspan=3><%=strName%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD valign=top><b>Product:</b></TD>
	<TD><%=strProductName%>&nbsp;</TD>
	<TD valign=top><b>Deliverable&nbsp;ID:</b></TD>
	<TD width="100%" colspan=3><%=strDelID%></TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>HW Version:&nbsp;&nbsp;&nbsp;&nbsp;</b></TD>
	<TD width=40%><%=strVersion%>&nbsp;</TD>
	<TD nowrap valign=top><b>Vendor:</b></TD>
	<TD width=60%><%=strVendor%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>FW Version:</b></TD>
	<TD><%=strRevision%>&nbsp;</TD>
	<TD nowrap valign=top><b>Model&nbsp;Number:</b></TD>
	<TD><%=strModelNumber%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Revision:</b></TD>
	<TD><%=strPass%>&nbsp;</TD>
	<TD nowrap valign=top><b>Part&nbsp;Number:</b></TD>
	<TD><%=strPartNumber%>&nbsp;</TD>
</TR>
<TR bgcolor=cornsilk style=display:none>
	<TD nowrap valign=top><b>OTS - Root:</b></TD>
	<TD><%=OTSRootText%></TD>
	<TD nowrap valign=top><b>OTS - Version:</b></TD>
	<TD>4 Open Observations</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Test&nbsp;Status:</b></TD>
	<TD>
		<SELECT style="width:100%" id=cboStatus name=cboStatus LANGUAGE=javascript onclick="return cboStatus_onclick()">
			<OPTION value=0 selected></OPTION>
			<%if trim(strStatus) = "1" then%>
				<OPTION value=1 selected >Passed</OPTION>
			<%else%>
				<OPTION value=1>Passed</OPTION>
			<%end if%>
			<%if trim(strStatus) = "2" then%>
				<OPTION value=2 selected>Failed</OPTION>
			<%else%>
				<OPTION value=2>Failed</OPTION>
			<%end if%>
			<%if trim(strStatus) = "3" then%>
				<OPTION value=3 selected>Blocked</OPTION>
			<%else%>
				<OPTION value=3>Blocked</OPTION>
			<%end if%>
			<%if trim(strStatus) = "4" then%>
				<OPTION value=4 selected>Watch</OPTION>
			<%elseif clng(request("FieldID")) = 3 then%>
				<OPTION value=4>Watch</OPTION>
			<%end if%>
			<%if trim(strStatus) = "5" then%>
				<OPTION value=5 selected>N/A</OPTION>
			<%elseif clng(request("FieldID")) = 3 then%>
				<OPTION value=5>N/A</OPTION>
			<%end if%>
		</SELECT>
	</TD>
	<TD nowrap valign=top><b>Samples&nbsp;Available:&nbsp;&nbsp;</b></TD>
	<TD ><INPUT style="width:60" maxlength=3 type="text" id=txtReceived name=txtReceived value="<%=strUnitsReceived%>"> <font size=1 color=green face=verdana>Total for your group.</font>
	</TD>
</TR>
<TR bgcolor=cornsilk>
	<TD nowrap valign=top><b>Test&nbsp;Notes:</b>&nbsp;<span style="Display:<%=DisplayNotesRequired%>" ID=RequireNotes><font color="#ff0000" size="1">*</font></a></span><BR><font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font></TD>
	<TD colspan=3>
	<TEXTAREA rows=4 style="width:100%" id=txtNotes name=txtNotes onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"><%=strTestNotes%></TEXTAREA>
	</TD>
</TR>
</table>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtFieldID name=txtFieldID value="<%=request("FieldID")%>">
<INPUT type="hidden" id=txtPMEmail name=txtPMEmail value="<%=strPMEmail%>">

</form>

	<%end if%>
</BODY>
</HTML>
