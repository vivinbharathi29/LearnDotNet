<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<STYLE>
TABLE
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana, Tahoma, Arial
}
A:link
{
    COLOR: Blue;
}
A:visited
{
    COLOR: Blue;
}
A:hover
{
    COLOR: red;
}
</STYLE>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
	function row_onmouseover() {
		window.event.srcElement.parentElement.style.cursor = "hand"
		window.event.srcElement.parentElement.style.color = "red"
	
	}
	function row_onmouseout() {
		window.event.srcElement.parentElement.style.color = "black"
	
	}

	function row_onclick(){
	window.alert("Not implimented yet.");
	return;
	
	strResult = window.showModalDialog("../mobilese/today/action.asp?" + window.event.srcElement.parentElement.className,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
		if (typeof(strResult) != "undefined")
		{
			window.location.reload(true);
		}
	}

//-->
</SCRIPT>
</HEAD>
<BODY>

<FONT face=verdana>
<h3>Deliverable Matrix Changes - Mobile Products</h3>
<font size=1 face=verdana>Report Period: (<%=Date() - 7%> - <%=Date()  %>)<BR><BR></font>
<P>
<font size=2 face=verdana><b><u>Deliverables Targeted This Week:</u></b></font><BR>

<%
	dim strVendor
	dim strVersion
	dim strCurrent
	dim strLastProduct	
	dim strDeliverable
	dim strDate
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spListMatrixUpdates",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<TABLE cellSpacing=1 cellPadding=1 width=""100%"" border=1 borderColor=tan bgColor=ivory>"
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR>"
	else
		do while not rs.EOF
			if strLastproduct <> rs("Product")  then
				if strlastproduct <> "" then
					Response.Write "</table>"	
				end if
				Response.write "<BR><font size=2 face=verdana><b>" & rs("Product") & "</b></font>"
				Response.Write "<TABLE cellSpacing=1 cellPadding=1 width=450 border=1 borderColor=tan bgColor=ivory>"
				Response.Write  "<TR><TD width=370><b>Deliverable</b></TD><TD width=80><b>Targeted</b></TD></TR>"
				strLastproduct = rs("Product") 
			end if
	
			strDeliverable = rs("Deliverable") & " " & rs("version")
			if rs("revision") <> "" then
				strDeliverable = strDeliverable & "," & rs("Revision")
			end if
			if rs("Pass") <> "" then
				strDeliverable = strDeliverable & "," & rs("Pass")
			end if
			
			if rs("Updated") & "" = "&nbsp;" then
				strDate = ""
			else
				strDate = formatdatetime(rs("Updated"), vbshortdate)
			end if
			
			Response.Write  "<TR class=""ID=" & rs("ID")& "&Type=3"" LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""><TD>" & strDeliverable & "</TD><TD>" & strDate & "</TD></TR>"
			rs.MoveNext
		loop
	end if
	rs.close	
%>	  
   </Table>
<BR>
  
</font>
</BODY>
</HTML>