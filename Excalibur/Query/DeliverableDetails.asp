<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD
{
	font-size: xx-small;
}
</STYLE>
<BODY>

<font size=2 face=verdana><b>Deliverables Failed By Functional Test</b><BR></font>
<%
	set cn = server.CreateObject("ADODB.Connection")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")

	rs.Open "spListFunctionalTestFailures",cn,adOpenForwardOnly
	do while not rs.EOF
		strVersion = rs("version") & ""
		if trim(rs("revision") & "") <> "" then
			strVersion = strVersion & "," & rs("version")
		end if
		if trim(rs("pass") & "") <> "" then
			strVersion = strVersion & "," & rs("pass")
		end if
		Response.Write "<Table bgcolor=ivory cellpadding=2 cellspacing=0 border =1 width=""100%"">"
		Response.Write "<TR><td><b>Name:</b></td><TD>" & rs("Name") & "</td><td><b>Version:</b></td><TD>" & strVersion & "</td></tr>"
		Response.Write "<TR><td><b>Released:</b></td><TD>" & rs("ReleaseDate") & "</td><td><b>Vendor&nbsp;Version:</b></td><TD>" & rs("vendorversion") & "&nbsp;</td></tr>"
		Response.Write "<TR><td><b>Failed:</b></td><TD>" & rs("FailDate") & "</td><td><b>Days&nbsp;In&nbsp;Test:</b></td><TD>" & rs("DaysInTest") & "</td></tr>"
		Response.Write "<TR><td colspan=4><b>Comments:</b><BR>" & replace(rs("Comments") & "",vbcrlf,"<BR>") & "</td></tr>"

		Response.Write "</Table><BR>"
		rs.MoveNext
	loop
	rs.Close

	set rs = nothing
	cn.Close
	set cn = nothing
%>

</BODY>
</HTML>
