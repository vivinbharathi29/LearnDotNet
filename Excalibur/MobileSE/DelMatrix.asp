<%@ Language=VBScript %>
<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<title><%= request("Product")%> Deliverable Matrix - HP Restricted</title>
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--

Sub window_onload
	generating.style.display="none"
	reporttitle.style.display=""
End Sub

-->
</SCRIPT>
</head>
<body>
<% if request("Product")= "" or request("ID") = "" then
	Response.Write "<label id=generating></label><label id=reporttitle></label>Can not determine product requested."
	
else

%>

<Label id="generating">Generating Matrix.  Please wait...</label>
<TABLE border=0 width="100%"><TR><TD align=middle>
<h2>
<%
  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

  'Create a recordset
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn

  Dim PreRow
  Dim MidRow
  Dim PostRow
  Dim PreRowA
  Dim MidRowA
  Dim PostRowA
  dim PrimaryColor 
  dim strComments
  dim strMilestone
  dim strPOR
  dim strTarget
  dim strActual
  dim strproduct
  dim strLastMilestone
  dim rsCount   
  dim rowcount
  dim strExpired
  dim Milestonecount
  dim rs
  dim strDate
  dim strTargetDate
  dim strDays

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=#e6f7ff valign=top><TD>"
  MidRow = "</TD><TD valign=top ><FONT size=1>"
  PostRow = "</FONT></TD></TR>"
  
  
  %><label id=reporttitle style="Display: none"><%
  Response.Write request("Product") & " Deliverable Matrix</h2>" 

%></label>

</TD></TR></TABLE><!-- Start Objective Section -->

	<table ID=ObjectiveTable bgColor=#e6f7ff border=1 borderColor=skyblue cellPadding=2 cellSpacing=1 width="100%">
		<tr bgcolor=steelblue> <TD align=middle><font color=white size=1><b>Category</b></font></TD><TD  align=middle><font color=white size=1><b>Name</b></font></TD><TD align=middle><font color=white size=1><b>Ver</b></font></TD><TD  align=middle><font color=white size=1><b>Rev</b></font></TD><TD  align=middle><font color=white size=1><b>Pass</b></font></TD><TD  align=middle><font color=white size=1><b>Vendor Version</b></font></TD><TD  align=middle><font color=white size=1><b>Developer</b></font></TD><TD  align=middle><font color=white size=1><b>Workflow Step</b></font></TD></tr>
<%
	rs.Open "spListTargetedDel4Product " & clng(request("ID")) & ";",cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		 response.write "<TR bgcolor=#e6f7ff valign=top><TD colspan=9><font size=1>No deliverables targeted for this product." & "</font></TD>" &  postrow
	else
		do while not rs.EOF
			response.write prerow & "<font size=1 face=verdana>" & rs("Category") & midrow & "<font size=1 face=verdana>" & rs("Name") & "</font>"  & midrow & "<font size=1 face=verdana>" & rs("Version") & "</font>"  & midrow & "<font size=1 face=verdana>" & rs("Revision")& "&nbsp;" & "</font>"  &  "</font>"  & midrow & "<font size=1 face=verdana>" & rs("Pass") & "&nbsp;" & "</font>" & midrow & "<font size=1 face=verdana>" & rs("VendorVersion")& "&nbsp;" &  midrow & "<font size=1 face=verdana>" & rs("Developer") & midrow & "<font size=1 face=verdana>" & rs("location") & postrow
			rs.MoveNext
		loop
	end if
%>	

</table>

<br>
<br>
<font size="1">Report Generated <%=formatdatetime(date(),vblongdate) %></font>
<br>
<br>
<font Size="1" Color="red"><p><strong>HP Restricted</strong></p></font>
<%end if%>
</body>
</html>
