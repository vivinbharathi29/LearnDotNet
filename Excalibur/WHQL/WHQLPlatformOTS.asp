<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</HEAD>

<BODY>

<%
'Response.Flush()
%>

<%Request("Platform")%>
<%if request("Platform") = "Aspen" then%>
	<b>Aspen 1.0 WHQL OTS List</b>
<%elseif request("Platform") = "Davos" then%>
	<b>Davos WHQL OTS List</b>
<%elseif request("Platform") = "Heavenly" then%>
	<b>Heavenly WHQL OTS List</b>
<%elseif request("Platform") = "Vail" then%>
	<b>Vail WHQL OTS List</b>	
<%end if%>
<br><br>
<%
	dim cn
	dim rs
	dim rowcount

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
%>
	<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<thead>
	<tr>
	<td valign=top width=50><b><font size=1 face=verdana>&nbsp;ID</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;author</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;Owner</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;Developer</b></font></td>
	<td valign=top width=50><b><font size=1 face=verdana>&nbsp;priority</b></font></td>
	<td valign=top width=150><b><font size=1 face=verdana>&nbsp;ShortDescription</b></font></td>
	<td valign=top width=50><b><font size=1 face=verdana>&nbsp;State</b></font></td>
	<td valign=top width=50><b><font size=1 face=verdana>&nbsp;ErrorType</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;Frequency</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;testPlatform</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;systemboard</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;systemROM</b></font></td>
	<td valign=top width=90><b><font size=1 face=verdana>&nbsp;Category</b></font></td>
	</tr></thead>
<%
	rowcount = 0	if request("Platform") <> "" then
		rs.Open "spListOTSSearch " & request("Platform") & ", '%WHQL%'" & ", '%signature%'" ,cn,adOpenForwardOnly
		do while not rs.EOF
			rowcount = rowcount + 1
			Response.Write "<tr>"
			Response.Write "<TD valign=top width=50><a target=OTS href=""http://" & Application("Excalibur_ServerName") & "/search/ots/Report.asp?txtReportSections=1&txtObservationID=" & rs("ObservationID") & """><font size=1 face=verdana>" & rs("ObservationID") & "</font></a></TD>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>"& rs("author") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>"& rs("Owner") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>"& rs("Developer") & "</td></font>"
			Response.Write "<td valign=top width=50><font size=1 face=verdana>"& rs("priority") & "</td></font>"
			Response.Write "<td valign=top width=150><font size=1 face=verdana>" & rs("ShortDescription") & "</td></font>"
			Response.Write "<td valign=top width=50><font size=1 face=verdana>" & rs("State") & "</td></font>"
			Response.Write "<td valign=top width=50><font size=1 face=verdana>" & rs("ErrorType") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>" & rs("Frequency") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>" & rs("testPlatform") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>" & rs("systemboard") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>" & rs("systemROM") & "</td></font>"
			Response.Write "<td valign=top width=90><font size=1 face=verdana>" & rs("Category") & "</td></font>"
			Response.Write "</tr>"
			rs.MoveNext
		loop
			Response.Write "<font size=1 face=verdana>Total Observations: " & rowcount & "</font>"
		rs.Close	
	end if
%>
	</table>
<%
	set rs=nothing
	cn.Close
	set cn=nothing
%>


</BODY>
</HTML>
