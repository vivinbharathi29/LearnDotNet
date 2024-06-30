<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--



//-->
</SCRIPT>
</HEAD>
<STYLE>
BODY
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}
TD
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: xx-small;
}

H1
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: small;
}
</STYLE>
<BODY bgColor=Ivory>
<H1>Lead Product Deliverable Usage Exceptions</H1>
	<form ID=frmMain action="ViewExceptionsSave.asp" method=post>
	<%
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
	
		
		rs.Open "spListLeadproductExceptions " &  clng(request("PMID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			Response.Write "No Exclusions are currently defined on active products."
			rs.Close
		else
			Response.Write "<font size=2 face=verdana>Select exceptions to remove</font><BR><BR>"
			Response.write "<table bgcolor=ivory border=1 cellpadding=2 cellspacing=0 width=""100%"">"
			Response.Write "<TR bgcolor=beige><TD>&nbsp;</TD><TD><b>Product</b></TD><TD width=""60%""><b>Deliverable</b></TD><TD><b>Versions</b></TD><TD><b>Distribution</b></TD><TD><b>Notes</b></TD><TD><b>Images</b></TD><TD width=""40%""><b>Comments</b></TD></tr>"
			do while not rs.EOF
				strVersion = rs("Version") & ""
				if trim(rs("Revision")&"") <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if trim(rs("Pass")&"") <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				if rs("TypeID") = 1 then
					strIDPrefix="R"
				else
					strIDPrefix="V"
				end if
				Response.write "<TR>"
				Response.write "<TD width=10><INPUT type=""checkbox"" id=chkRemove name=chkRemove value=""" & strIDPrefix & rs("ID") & ":" & rs("ReleaseID") & """></TD>"
				Response.write "<TD nowrap>" & rs("Product") & "</TD>"
				Response.write "<TD>" & rs("Deliverable") & "</TD>"
				Response.write "<TD>" & strVersion & "</TD>"
				Response.write "<TD>" & replace(replace(rs("SyncDistribution")&"","False","Ignore"),"True","&nbsp;") & "</TD>"
				Response.write "<TD>" & replace(replace(rs("SyncNotes")&"","False","Ignore"),"True","&nbsp;") & "</TD>"
				Response.write "<TD>" & replace(replace(rs("SyncImages")&"","False","Ignore"),"True","&nbsp;") & "</TD>"
				Response.write "<TD>" & rs("Comments") & "&nbsp;</TD>"
				Response.write "</TR>"
				rs.MoveNext
			loop
			Response.write "</table>"
			rs.Close
		end if
	
	%>

	</form>
	</TD>
	</TR>


	<%
		set rs = nothing
		cn.Close
		set cn = nothing
	%>

</BODY>
</HTML>