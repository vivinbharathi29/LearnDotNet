<%@ Language=VBScript %>
<html>
<head>
<title>Requirement Exception Report - HP Restricted</title>
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<table border="0" cellPadding="1" cellSpacing="1" width="100%">
  
  <tr>
    <td Width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
    <td><img SRC="images/RequirementExceptionReport.gif" WIDTH="337" HEIGHT="50"></td>
    <td Align="Right"><img SRC="images/programoffice.gif" WIDTH="283" HEIGHT="50"></td></tr></table>

<%
  dim rs
  dim ProjectFound
  dim strDescription
  dim strDeliverables

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
  
  
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn
  rs.Open "SELECT PDDReleased,Description FROM productVersion with (NOLOCK) WHERE id = " & clng(request("ID")) & ";"

  if rs.eof and rs.bof then
%>
<h2><%= request("Product")%> Product</h2>
<%	ProjectFound = false
  else
 %>
<h2><%= request("Product")%> Product</h2>
<%
	if isnull(rs("PDDReleased")) then
%>
<p><font color="red"><strong>WARNING: This information has not been finalized or released.</strong></font></p>
<%
	else
%>
<p>Released <%= formatdatetime (rs("PDDReleased"),vbshortdate)%></p>
<%
	end if
	projectfound = true	
  end if
  	  
  rs.close
	
%> 

<br>
<br>
The following is a list of all requirements that are not satisifed by at least one deliverable for this product.
<hr>



<%
  Dim PreRow
  Dim MidRow
  Dim PostRow
  Dim PreRowA
  Dim MidRowA
  Dim PostRowA
  dim PrimaryColor 
  Dim strDetails
  dim strName
  dim strSummary
  Dim strSubsystem
  dim LastCategory
  dim strCategory
  
  'Setup table row HTML for primary color cells
  PreRowA = "<TR  bgcolor=white valign=top><TD>"
  MidRowA = "</TD><TD valign=top>"
  PostRowA = "</TD></TR>"

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=#e6f7ff valign=top><TD>"
  MidRow = "</TD><TD valign=top>"
  PostRow = "</TD></TR>"

  'Create a recordset
  'set rs = server.CreateObject("ADODB.recordset")
  rs.Open "SELECT pr.Specification as Specification, req.Name as Reqname, cat.name as Category FROM Requirement req with (NOLOCK), Product_requirement pr with (NOLOCK), Category cat with (NOLOCK) WHERE pr.deliverables Like '' AND cat.id = req.categoryid and pr.requirementid = req.id and pr.productid = " & clng(request("ID")) & " ORDER By cat.DisplayOrder,Req.name;"
  'Initializae the cell backgroud color selection
  primarycolor = false
  
  'Display requirements
  lastcategory = ""
  do while not rs.EOF
	strDetails =  rs("Specification").GetChunk(9999999) & ""
	strname =  rs("reqname")
	if trim(strDetails) = "" then
		strDetails = "&nbsp;"
	else
		strDetails = replace(replace(replace(strDetails & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>")
	end if
	if lastcategory <> rs("category") then
		if lastcategory <> "" then
			'Finish off previous table
		    Response.Write "</TBODY></table></p>"
		end if
		Response.Write "<H3>" & rs("Category") & "</H3>"
		Response.Write "<table  bgColor=#e6f7ff border=1 borderColor=skyblue cellPadding=2 cellSpacing=1 width=""95%"">"
		Response.Write "<tr bgColor=#006697><td width=""20%""><strong><font color=white>"
		Response.write "Requirement</font></strong></td><td width=""80%"">"
		Response.Write "<font color=white><strong>Specification</strong></font></td></tr>"
		lastcategory = rs("Category")
    end if
	primarycolor = not primarycolor
	strCategory = rs("Category")
	
	if primarycolor then
		Response.Write prerow & "<FONT size=1>" & strname  & "</FONT>" &  midrow & " <FONT size=1>" & strDetails & "</font>" & postrow
	else
		Response.Write prerowA & "<FONT size=1>" & strname & "</FONT>" &  midrowA & " <FONT size=1>" & strDetails & " </font>" & postrowA
	end if
	rs.MoveNext
  loop
  
  if rs.EOF and rs.BOF then
	if projectfound then
		'Say there were no requirements
		Response.Write "</table><strong>Deliverables have been mapped to all requirements for this product.</strong>"	
	else
		'Say the project doesn't exist
		Response.Write "</table><strong>Product is not available.</strong>"
	end if  
  else
   'Finish off table 
    Response.Write "</table></p>"
  end if
  
  'Cleanup
  rs.Close
%>
<br>
<br>
<font Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>

</body>
</html>
