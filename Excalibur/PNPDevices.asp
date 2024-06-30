<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
td{
	font-family: verdana;
	font-size:xx-small;
}
</STYLE>
<BODY>

<%if request("Report") = "1" then%>
	<Font face=verdana size=2><b>PNP Device ID Numbers by Version</b></font><BR><BR>
	<Table border=1 cellspacing=0 cellpadding=2>
	<TR bgcolor=gainsboro><TD>ID</TD><TD>Name</TD><TD>Version</TD><TD>PNP Devices</TD></TR>
	<%

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.CommandTimeout =120
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")

		rs.Open "Select ID, deliverablename, version, revision, pass, PNPDevices from deliverableversion with (NOLOCK) where pnpdevices <> '' and pnpdevices is not null order by deliverablename, id",cn,adOpenStatic
        if rs.eof and rs.bof then
            response.write "<tr><td colspan=4>none</td></tr>"
        end if
		do while not rs.EOF
			strVersion = rs("Version")
			if trim(rs("Revision") & "") <> "" then
				strversion = strversion & "," & trim(rs("Revision") & "")
			end if
			if trim(rs("Pass") & "") <> "" then
				strversion = strversion & "," & trim(rs("pass") & "")
			end if
		
			Response.write "<TR>"
			Response.write "<TD valign=top>" & rs("ID") & "</TD>"
			Response.write "<TD valign=top>" & rs("deliverablename") & "</TD>"
			Response.write "<TD valign=top>" & strVersion & "</TD>"
			Response.write "<TD valign=top>" & rs("pnpdevices") & "</TD>"
			Response.write "</TR>"
			rs.MoveNext
		loop
		rs.Close


		set rs = nothing
		cn.Close
		set cn = nothing


	%>
	</Table>
	
	<%elseif request("Report") = "2" then%>
	
	<%
	
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.CommandTimeout =120
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
	
		rs.open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenStatic
		if rs.eof and rs.bof then
			strProductName = "this Product"
		else
			strProductName = rs("name")
		end if
		rs.close	
	%>
	<Font face=verdana size=2><b>PNP Device ID Numbers for targeted deliverables on <%=strProductname%></b></font><BR><BR>
	<Table border=1 cellspacing=0 cellpadding=2>
	<TR bgcolor=gainsboro><TD>ID</TD><TD>Name</TD><TD>Version</TD><TD>PNP Devices</TD></TR>
	<%


		rs.Open "Select v.ID, v.deliverablename, v.version, v.revision, v.pass, v.PNPDevices from deliverableversion v with (NOLOCK), product_deliverable pd with (NOLOCK) where pd.deliverableversionid = v.id and pd.productversionid=" & clng(request("ProductID")) & " and pd.targeted=1 and pnpdevices <> '' and pnpdevices is not null order by v.deliverablename, v.id",cn,adOpenStatic
        if rs.eof and rs.bof then
            response.write "<tr><td colspan=4>none</td></tr>"
        end if
		do while not rs.EOF
			strVersion = rs("Version")
			if trim(rs("Revision") & "") <> "" then
				strversion = strversion & "," & trim(rs("Revision") & "")
			end if
			if trim(rs("Pass") & "") <> "" then
				strversion = strversion & "," & trim(rs("pass") & "")
			end if
		
			Response.write "<TR>"
			Response.write "<TD valign=top>" & rs("ID") & "</TD>"
			Response.write "<TD valign=top>" & rs("deliverablename") & "</TD>"
			Response.write "<TD valign=top>" & strVersion & "</TD>"
			Response.write "<TD valign=top>" & rs("pnpdevices") & "</TD>"
			Response.write "</TR>"
			rs.MoveNext
		loop
		rs.Close


		set rs = nothing
		cn.Close
		set cn = nothing


	%>
	</Table>


<%else%>
	<Font face=verdana size=2><b>PNP Device ID Numbers grouped by distinct device and deliverable</b></font><BR><BR>

	<Table border=1 cellspacing=0 cellpadding=2>
	<TR bgcolor=gainsboro><TD>Name</TD><TD>PNP Devices</TD><TD>Versions</TD></TR>
	<%

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
		cn.CommandTimeout =120
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")

		rs.Open "Select r.name, v.PNPDevices, Count(1) as Versions from deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK) where v.pnpdevices <> '' and r.id = v.deliverablerootid and v.pnpdevices is not null group by r.name, v.PNPDevices order by r.name",cn,adOpenStatic
        if rs.eof and rs.bof then
            response.write "<tr><td colspan=4>none</td></tr>"
        end if
		do while not rs.EOF
		
			Response.write "<TR>"
			Response.write "<TD valign=top>" & rs("name") & "</TD>"
			Response.write "<TD valign=top>" & replace(rs("pnpdevices") & "",vbcrlf,"<BR>") & "</TD>"
			Response.write "<TD valign=top>" & rs("Versions") & "</TD>"
			Response.write "</TR>"
			rs.MoveNext
		loop
		rs.Close


		set rs = nothing
		cn.Close
		set cn = nothing


	%>
	</Table>

<%end if%>

</BODY>
</HTML>
