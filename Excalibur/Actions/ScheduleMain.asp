<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<HEAD>
<script language="JavaScript" src="../_ScriptLibrary/jsrsClient.js"></script>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}


function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


function myCallback( returnstring ){
	CellPriority.innerHTML = returnstring; 
}


function window_onload() {
	frmUpdate.txtNotes.value=frmUpdate.txtNotes.mytag;
	frmUpdate.txtSummary.value=frmUpdate.txtSummary.mytag;
	frmUpdate.txtSummary.focus();
}


function cboProject_onchange() {
	strID = event.srcElement.value;
	if (event.srcElement.value !="")
		{
	      jsrsExecute("ScheduleRSget.asp", myCallback, "getItem", strID);
		}
}

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">


<%

	dim cn
	dim rs
	dim i
	dim cm
	dim p
	dim CurrentUser
	dim CurrentUserID
	dim strID
	dim strSummary
	dim strProduct
	dim strOriginalTimeframe
	dim strTimeframe
	dim strStatus
	dim strMilestone
	dim strNotes
	dim strDetails
	dim blnFound
	dim strDisplayOrder
	dim strPriority
	dim strOwnerID
	dim strTimeFrameNotes
	dim strOnReport
	
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
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	strID = trim(Request("ID"))
	strProduct = trim(Request("ProductID"))
	strSummary = ""
	strOriginalTimeframe = ""
	strTimeframe = ""
	strStatus = ""
	strMilestone = ""
	strNotes = ""
	strDetails = ""
	strDisplayOrder="1"
	strPriority = 0
	strOwnerID = 0
	strTimeFrameNotes = ""
	strOnReport = ""

	if strID <> "" then
		rs.Open "spGetActionRoadmapItemProperties " & clng(strID),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			blnFound = false
		else
			blnFound = true
			strProduct = rs("ProductVersionID") & ""
			strOwnerID = rs("OwnerID") & ""
			strStatus = rs("ActionStatusID") & ""
			strSummary = rs("Summary") & ""
			strTimeframe = rs("Timeframe") & ""
			strOriginalTimeframe = rs("OriginalTimeframe") & ""
			strNotes = rs("Notes") & ""
			strTimeFrameNotes = rs("TimeframeNotes") & ""
			strDisplayOrder = trim(rs("DisplayOrder") & "")
			strOnReport = replace(replace(rs("StatusReport"),"True","checked"),"False","")
			strDetails = rs("Details") & ""
'			rs.Close
'			
'			rs.open "spGetActionRoadmapTaskCounts " & clng(strID),cn,adOpenForwardOnly
'			if rs.EOF and rs.BOF then
'				strTasks = "&nbsp;"
'			else
'				strTasks = trim(rs("TaskCount"))
'				if strTasks =0 then
'					strTasks = "None Defined"
'				else
'					strTasks = strTasks & " (" & int(((rs("CompleteCount")/strTasks)*100)) & "% Complete)"
'				end if
'			end if
		end if
		rs.Close
	else
		blnFound = true
	end if


if not blnFound then
	Response.Write "Unable to find the requested roadmap item."
else
%>



<font face=verdana size=><b>
<label ID="lblTitle">
<%if strID = "" then%>
	Add Roadmap Item
<%else%>
	Update Roadmap Item
<%end if%>
</label></b></font>

<form id="frmUpdate" method="post" action="ScheduleSave.asp">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<td valign=top width=120 nowrap><b>Summary:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD colspan=3><INPUT style="width:100%" type="text" id=txtSummary name=txtSummary maxlength=256 mytag="<%=replace(strSummary,"""","&quot;")%>" value=""></TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Project:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD><SELECT style="width=150" id=cboProject name=cboProject LANGUAGE=javascript onchange="return cboProject_onchange()">
				<OPTION selected value=0></OPTION>
				<%					rs.Open "spGetProducts 2",cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(rs("ID")) = trim(strProduct) then
							Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
						else
							Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
						end if						rs.MoveNext
					loop
					rs.Close				%>
			</SELECT>
			</SELECT></TD>
		<td valign=top width=110 nowrap><b>Owner:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD><SELECT style="Width:100%" id=cboOwner name=cboOwner Language=javascript onkeydown="return combo_onkeydown()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeypress="return combo_onkeypress()">
					<Option value=0 selected></option>
					<%					rs.Open "spGetEmployees",cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(rs("ID")) = trim(currentuserid) and strID = "" then
							Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						elseif trim(rs("ID")) = trim(strOwnerID) and strID <> "" then
							Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						elseif rs("Active") then
							Response.Write "<OPTION value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						end if						rs.MoveNext
					loop
					rs.Close					%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Status:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD><SELECT style="width=150" id=cboStatus name=cboStatus>				<%					rs.Open "spListActionStatuses 3",cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(rs("ID")) = trim(strStatus) or (strID="" and rs("ID")=1) then
							Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						else
							Response.Write "<OPTION value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						end if						rs.MoveNext
					loop
					rs.Close				%>
			</SELECT>
		</TD>
		<td valign=top width=110 nowrap><b>Target:</b></td>
		<TD width=100%>
			<INPUT style="Width:100%" type="text" id=txtTimeframe name=txtTimeframe mytag="<%=strTimeframe%>" value="<%=strTimeframe%>">
			<INPUT type="hidden" id=txtOriginalTimeframe name=txtOriginalTimeframe value="<%=strOriginalTimeframe%>">
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Status&nbsp;Notes:</b></td>
		<TD colspan=3><INPUT style="width:100%" type="text" id=txtNotes name=txtNotes maxlength=80 mytag="<%=strNotes%>" value=""></TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Target Notes:</b></td>
		<TD width=100% colspan=3>
		<%
		if trim(strOriginalTimeframe) <> "" then
			Response.Write "<font size=2 face=verdana>Original Target: " & strOriginalTimeframe & "<BR></font>"
		end if
		%>
		<INPUT style="Width:100%" type="text" id=txtTimeframeNotes name=txtTimeframeNotes  maxlength=500 mytag="<%=strTimeframeNotes%>" value="<%=strTimeframeNotes%>">
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Follows Item:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD colspan=3 ID=CellPriority>			<SELECT style="width=100%" id=cboPriority name=cboPriority>				<option value="" selected>[ Beginning of Roadmap ]</Option>
				<%					rs.Open "spListActionRoadmap " & clng(strProduct),cn,adOpenForwardOnly
					i=0					do while not rs.EOF
						i=i+1
						if trim(rs("ID")) <> trim(strID) then
							if clng(rs("DisplayOrder")) = (clng(strDisplayOrder) -1) then								Response.Write "<OPTION selected value=" & rs("ID") & ">" & i & ". " &  rs("summary") & "</OPTION>"
								strPriority = rs("ID")
							else								Response.Write "<OPTION value=" & rs("ID") & ">" & i & ". " &  rs("summary") & "</OPTION>"
							end if
						end if						rs.MoveNext
					loop
					rs.Close				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Details:</b></td>
		<TD colspan=3><TEXTAREA rows=6 style="Width=100%" id=txtDetails name=txtDetails><%=strDetails%></TEXTAREA>
	</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Report:</b></td>
		<TD width=100% colspan=3><INPUT <%=strOnReport%> type="checkbox" id=chkReport name=chkReport> Include on the Roadmap Summary Report</TD>
	</TR>
</table>


<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=strID%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=tagDisplayOrder name=tagDisplayOrder value="<%=strPriority%>">
</form>
<%

end if

	set rs = nothing
	cn.Close
	set cn = nothing


%>


</BODY>
</HTML>


