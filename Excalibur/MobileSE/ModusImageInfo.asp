<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<TABLE ID=TableLocal width=100%>
<TR><TD>
<%
	dim cn, rs
	
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")

	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

rs.Open "spListImageDefinitionsByProduct " & request("ID"),cn,adOpenForwardOnly 

if rs.EOF and rs.BOF then
	Response.Write "<font size=2 face=verdana>No Images definitions or RTM dates available for this product.</font>"	
else

%>

<table cellpadding=2 border=1 bordercolor=tan bgcolor=ivory>
<TR bgcolor=cornsilk>
	<!--<TD><INPUT style="WIDTH:12;HEIGHT:12;" type="checkbox" id=chkImagesAll name=chkImagesAll></TD>
	--><TD><b><font size=1 face=verdana>SKU Number</font></b></TD>
	<TD><b><font size=1 face=verdana>Brand</font></b></TD>
	<TD><b><font size=1 face=verdana>Software</font></b></TD>
	<TD><b><font size=1 face=verdana>OS</font></b></TD>
	<TD><b><font size=1 face=verdana>Type</font></b></TD>
	<TD><b><font size=1 face=verdana>Status</font></b></TD>
	<TD><b><font size=1 face=verdana>RTM Date</font></b></TD>
	<TD><b><font size=1 face=verdana>Comments</font></b></TD>
</TR>

<%
	if request("ID") = 268 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>SMB images are for Thurman 1.1 only</b></td></TR>"
	elseif request("ID") = 267 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>These images are shared with Thurman 1.1</b></td></TR>"
	end if

	do while not rs.EOF
	%>
		<TR class="ID=<%=rs("ID")%>" id="Imagerows" LANGUAGE=javascript onmouseover="return imagerows_onmouseover()" onmouseout="return imagerows_onmouseout()" oncontextmenu="localizationMenu(<%=rs("ID")%>);return false;" onclick="return localizationMenu(<%=rs("ID")%>)">
			<!--<TD class="check"><INPUT class="check" style="WIDTH:12;HEIGHT:12;" type="checkbox" id=chkImages<%=rs("ID")%> name=chkImages<%=rs("ID")%>></TD>
			--><TD class="cell"><font size=1 face=verdana class="text"><%=rs("SKUNumber") & "&nbsp;" %></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Brand") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("SWType") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("OS") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("ImageType") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Status") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("RTMDate") & "&nbsp;"%></font></TD>
			<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Comments") & "&nbsp;"%></font></TD>
		</TR>
	<%	
		rs.Movenext
	loop

	if request("ID") = 268 then
		response.write "<TR bgcolor=MediumAquamarine><TD colspan=8><font size=1 face=verdana><b>Consumer images are shared with Ford 1.1</b></td></TR>"
		rs.Close
		rs.Open "spListImageDefinitionsByProduct " & 267,cn,adOpenForwardOnly 

		do while not rs.EOF
		%>
			<TR class="ID=<%=rs("ID")%>" id="Imagerows" LANGUAGE=javascript onmouseover="return imagerows_onmouseover()" onmouseout="return imagerows_onmouseout()" oncontextmenu="localizationMenu(<%=rs("ID")%>);return false;" onclick="return localizationMenu(<%=rs("ID")%>)">
				<!--<TD class="check"><INPUT class="check" style="WIDTH:12;HEIGHT:12;" type="checkbox" id=chkImages<%=rs("ID")%> name=chkImages<%=rs("ID")%>></TD>
				--><TD class="cell"><font size=1 face=verdana class="text"><%=rs("SKUNumber") & "&nbsp;" %></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Brand") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("SWType") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("OS") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("ImageType") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Status") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("RTMDate") & "&nbsp;"%></font></TD>
				<TD class="cell"><font size=1 face=verdana class="text"><%=rs("Comments") & "&nbsp;"%></font></TD>
			</TR>
		<%	
			rs.Movenext
		loop
	end if

%>
</table>
</td></tr>
</TABLE>
<%
rs.Close
end if
%>

</BODY>
</HTML>


