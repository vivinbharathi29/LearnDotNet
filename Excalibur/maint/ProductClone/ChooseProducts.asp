<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdNext_onclick() {
	if (frmMain.lstSource.selectedIndex == -1)
		alert("You must select a source product to continue.");	
	else if (frmMain.lstTarget.selectedIndex == -1) 
		alert("You must select a destination product to continue.");	
	else if (frmMain.lstTarget.selectedIndex == frmMain.lstSource.selectedIndex)
		alert("You must select different source and destination products to continue.");	
	else
	frmMain.submit();
}

//-->
</SCRIPT>
<STYLE>
body
{
	font-family: Verdana;
	font-size: xx-small;
}

</STYLE>
</HEAD>


<BODY>


<H3>Product Clone Wizard</H3>
<font size=2><b>Step 1: Choose Products to Clone</font></b>

<%
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")

%>

<form ID=frmMain method=post action="FindRoots.asp">
<TABLE border=0><TR>

<TD>
	<font size=2>Source Product:<BR></font>
	<SELECT id=lstSource style="WIDTH: 171px; HEIGHT: 135px" size=2 name=lstSource> 
		<%
			rs.Open "Select ID, DotsName from productversion with (NOLOCK) where typeid in (1,3) and productstatusid < 4 order by dotsname",cn,adOpenForwardOnly
			do while not rs.EOF
				Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("DotsName") & "</OPTION>"
		
				rs.MoveNext
			loop
			rs.Close
		%>
		</SELECT>
</TD>
<TD width=10>&nbsp;</TD>

<TD>
	<font size=2>Destination Product:<BR></font>
	<SELECT id=lstTarget style="WIDTH: 171px; HEIGHT: 135px" size=2 name=lstTarget> 
		<%
			rs.Open "Select ID,DotsName from productversion with (NOLOCK) where typeid in (1,3) and productstatusid < 4 order by dotsname",cn,adOpenForwardOnly
			do while not rs.EOF
				Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("DotsName") & "</OPTION>"
		
				rs.MoveNext
			loop
			rs.Close
		%>
		</SELECT>
</TD>
</TR></TABLE><BR>
</form>
<INPUT type="button" value="Next" id=cmdNext name=cmdNext LANGUAGE=javascript onclick="return cmdNext_onclick()">

<%
	set rs = nothing
	cn.Close
	set cn = nothing

%>

</P>

</BODY>
</HTML>
