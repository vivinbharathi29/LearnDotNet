<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<STYLE>
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana, Tahoma, Arial
}
TH
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
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript" />
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />

<script  id="clientEventHandlersJS"  language="javascript" type="text/javascript">
<!--
	function row_onmouseover() {
		window.event.srcElement.parentElement.style.cursor = "hand"
		window.event.srcElement.parentElement.style.color = "red"
	
	}
	function row_onmouseout() {
		window.event.srcElement.parentElement.style.color = "black"
	
	}

	function row_onclick(){
	strResult = window.showModalDialog("../mobilese/today/action.asp?" + window.event.srcElement.parentElement.className,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
		if (typeof(strResult) != "undefined")
		{
			window.location.reload(true);
		}
	}

//-->
</script>
</head>
<body>

<font face=verdana>
<h3>Change Requests</h3>
<font size=1 face=verdana>Report Period: (<%=Date() - 7%> - <%=Date() -1 %>)<BR><BR></font>
<P>
<font size=2 face=verdana><b>Change Requests Added or Closed This Week:</b></font>

<%
	dim strVendor
	dim strVersion
	dim strCurrent
	dim strLastProduct	
	dim strStatus
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spListDCRThisWeek",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<TABLE cellSpacing=1 cellPadding=1 width=""100%"" border=0 bgColor=ivory>"
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR></table>"
	else
		Response.Write "<TABLE cellSpacing=0 cellPadding=2 width=""100%"" border=0 bgColor=ivory>"
		do while not rs.EOF
			if strLastproduct <> rs("Product")  then
				Response.Write "<TR><TD bgcolor=white colspan=6><b><font size=2><BR>" & rs("Product") & "</font><b></td></TR>"
				Response.Write  "<TR bgcolor=beige><TD style=""BORDER-TOP: lightgrey 1px solid;""><b>Number</b></TD><TD style=""BORDER-TOP: lightgrey 1px solid;""><b>Status</b></TD><TD style=""BORDER-TOP: lightgrey 1px solid;""><b>Submitter</b></TD><TD style=""BORDER-TOP: lightgrey 1px solid;""><b>Summary</b></TD></TR>"
				strLastproduct = rs("Product") 
			end if
	
			select case rs("Status")
			case 1
				strStatus = "Open"			
			case 2
				strStatus = "Need More Input"			
			case 3
				strStatus = "Closed"			
			case 4
				strStatus = "Approved"			
			case 5
				strStatus = "Disapproved"			
			case 6
				strStatus = "Investigating"			
			case else
				strStatus = "N/A"
			end select
	
			Response.Write  "<TR class=""ID=" & rs("ID")& "&Type=3"" LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""><TD style=""BORDER-TOP: lightgrey 1px solid;"">" & rs("ID") & "</TD><TD style=""BORDER-TOP: lightgrey 1px solid;"">" & strStatus & "&nbsp;&nbsp;&nbsp;</TD><TD style=""BORDER-TOP: lightgrey 1px solid;"" nowrap>" & rs("Submitter") & "&nbsp;</TD><TD style=""BORDER-TOP: lightgrey 1px solid;"">" & rs("Summary") & "</TD></TR>"
			rs.MoveNext
		loop
		Response.Write "</table>"	
	end if
	rs.close	
%>	  
   </table>
<br />
  
</font>
</body>
</HTML>