<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file = "DataWrapper.asp" --> 
<%
	Response.Buffer = True
	Response.ExpiresAbsolute = Now() - 1
	Response.Expires = 0
	Response.CacheControl = "no-cache"

	Dim dw
	Dim cn
	Dim cmd
	Dim rs
	Dim sMilestone, sProgram, IsMilestone, sPorStartDt, sPorEndDt, sOldStartDt, sOldEndDt, sNewStartDt, sNewEndDt
	
	If Trim(Request("ID")) = "" Then
		Response.End
	End If
	
	If UCase(Request.ServerVariables("SERVER_NAME")) <> "LOCALHOST" Then
		'Response.End
	End If
	
	'Get Status Information
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData")
	dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, Request("ID")
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_Active_YN", adInteger, adParamInput, 8, ""
	Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
	If rs.EOF and rs.BOF Then
'		Response.Write "<H3>No Data Returned for ScheduleDataID = " & Request("ID") & "</H3>"
		Response.End
	End If
	
	sMilestone = rs("item_description")
	sProgram = rs("family_name") & " " & rs("schedule_name")
	IsMilestone = rs("milestone_yn")
	sPorStartDt = rs("por_start_dt")
	If sPorStartDt = "" Then sPorStartDt = "&nbsp;"
	sPorEndDt = rs("por_end_dt")
	If sPorEndDt = "" Then sPorEndDt = "&nbsp;"
	sOldStartDt = Request("OldStartDt")
	If sOldStartDt = "" Then sOldStartDt = "&nbsp;"
	sOldEndDt = Request("OldEndDt")
	If sOldEndDt = "" Then sOldEndDt = "&nbsp;"
	sNewStartDt = rs("projected_start_dt")
	If sNewStartDt = "" Then sNewStartDt = "&nbsp;"
	sNewEndDt = rs("projected_end_dt")
	If sNewEndDt = "" Then sNewEndDt = "&nbsp;"

	rs.close
	set rs = nothing
	set cmd = nothing
	set cn = nothing
	

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../Reports/Reports.css">
</HEAD>
<BODY>

<P class=tabletitle>The <%= sMilestone%> Date on the <%= sProgram%> schedule has changed.</P>
<P>
<%
IF Ucase(IsMilestone) = "Y" Then
%>
<TABLE class=grid cellSpacing=1 cellPadding=3 >
  <THEAD>
  <TR>
    <TH width=33%>POR</TH>
    <TH width=33%>Planned</TH>
    <TH width=33%>Actual</TH></TR></THEAD>
  <TBODY>
  <TR>
    <TD><%=sPorEndDt%></TD>
    <TD><%=sOldStartDt%></TD>
    <TD><%=sNewEndDt%></TD>
    </TR></TBODY></TABLE></P>
<%
Else
%>
<TABLE class=grid cellSpacing=1 cellPadding=3 >
  <THEAD>
  <TR>
    <TH colspan=2 width=33%>POR</TH>
    <TH colspan=2 width=33%>Planned</TH>
    <TH colspan=2 width=33%>Actual</TH></TR>
  <TR>
    <TH width=16%>Start</TH>
    <TH width=16%>End</TH>
    <TH width=16%>Start</TH>
    <TH width=16%>End</TH>
    <TH width=16%>Start</TH>
    <TH width=16%>End</TH></TR></THEAD>
  <TBODY>
  <TR>
    <TD><%=sPorStartDt%></TD>
    <TD><%=sPorEndDt%></TD>
    <TD><%=sOldStartDt%></TD>
    <TD><%=sOldEndDt%></TD>
    <TD><%=sNewStartDt%></TD>
    <TD><%=sNewEndDt%></TD>
    </TR></TBODY></TABLE></P>
<%
End If
%>

</BODY>
</HTML>
