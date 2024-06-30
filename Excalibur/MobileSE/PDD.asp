<%@ Language=VBScript %>
<html>
<head>
<title>Product Definition Document - HP Restricted</title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
</head>

<body>
<table border="0" cellPadding="1" cellSpacing="1" width="100%">
  
  <tr>
    <td Width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
    <td><img SRC="images/ProductDefinitionDocument.gif" WIDTH="325" HEIGHT="50"></td>
    <td Align="Right"><img SRC="images/programoffice.gif" WIDTH="283" HEIGHT="50"></td></tr></table>

<h2>

<%
	dim cn
    dim rs
	dim ProjectFound
	dim strDescription
	dim strDeliverables

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
  
	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn  
  
  	'Get User
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPartner
	
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing	
	
	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=0"
	else
		CurrentUserPartner = rs("PartnerID")
	end if 
	rs.Close

  
  rs.Open "SELECT PartnerID, PDDReleased,Description FROM ProductVersion with (NOLOCK) WHERE id = " & clng(request("ID")) & ";"

  if rs.eof and rs.bof then
	Response.Write request("Product") & " Product"   
	ProjectFound = false
	PartnerID=0
  else
	Response.Write request("Product") & " Product" 
	PartnerID=rs("PartnerID")
	
	
	'Verify Access is OK
	if trim(CurrentUserPartner) <> "1" then
		if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
			set rs = nothing
			set cn=nothing
			
			Response.Redirect "../NoAccess.asp?Level=0"
		end if
	end if	
	
%>
</h2>
<%
	if isnull(rs("PDDReleased")) then
		Response.Write "<BR><font color=red><Strong>WARNING: This information has not been finalized or released.</Strong></font>"
	else
		Response.Write "Released " & formatdatetime (rs("PDDReleased"),vbshortdate)
	end if
	strDescription = trim(rs("description").GetChunk(9999999))
	if strDescription <> "" then
		Response.Write "<BR><BR><u>Product Description:</u><BR><ul> " & replace(replace( replace(strDescription,"<","&lt;"),">","&gt;"),vbcrlf,"<BR>") & "</ul>"
	end if
	projectfound = true	
  end if
  	  
  rs.close
	
%> 


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
  MidRowA = "</TD><TD valign=top><FONT size=1>"
  PostRowA = "</FONT></TD></TR>"

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=#e6f7ff valign=top><TD>"
  MidRow = "</TD><TD valign=top><FONT size=1>"
  PostRow = "</FONT></TD></TR>"

  'Create a recordset
  rs.Open "SELECT pr.Specification as Specification, pr.Deliverables as Deliverables, req.Name as Reqname, cat.name as Category FROM Requirement req with (NOLOCK), Product_requirement pr with (NOLOCK), Category cat with (NOLOCK) WHERE cat.id = req.categoryid and pr.requirementid = req.id and pr.Productid = " & clng(request("ID")) & " ORDER By cat.DisplayOrder,Req.name;"
  'Initializae the cell backgroud color selection
  primarycolor = false
  
  'Display requirements
  lastcategory = ""
  do while not rs.EOF
	strDetails =  rs("Specification").GetChunk(9999999) & ""
	strDeliverables = rs("Deliverables").getChunk(9999999) & ""
	if trim(strDeliverables) = "" then
		strDeliverables = "&nbsp;"
	else
		strDeliverables = replace(replace( replace(strDeliverables,"<","&lt;"),">","&gt;"),vbcrlf,"<BR>")
	end if
	strname =  rs("reqname")
	if trim(strDetails) = "" then
		strDetails = "&nbsp;"
	else
		strDetails = replace(replace(replace(strDetails & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>")
	end if
	if lastcategory <> rs("category") then
		if lastcategory <> "" then
			'Finish off previous table
		    Response.Write "</table></p>"
		end if
		Response.Write "<h3>" & rs("Category") & "</H3>"
		Response.Write "<table id=myTable bgColor=#e6f7ff border=1 borderColor=skyblue cellPadding=2 cellSpacing=1 width=""95%"">"
		Response.Write "<tr  bgColor=#006697><td width=""20%""><strong><font color=white>"
		Response.write "Requirement</font></strong></td><td width=""40%"">"
		Response.Write "<font color=white><strong>Description</strong></font></td>"
		Response.Write "<td width=""40%""><font color=white><strong>Deliverables</strong></font></td>"
		Response.Write "</tr>"
		lastcategory = rs("Category")
    end if
	primarycolor = not primarycolor
	'strSummary = replace(replace(replace(rs("Summary") & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>")
	strCategory = rs("Category")
	
	if primarycolor then
		Response.Write prerow & "<FONT size=1>" &  strname & "</FONT>" &  midrow & " <FONT size=1>" & strDetails & "</font>" & midrow & "<FONT size=1>" & strDeliverables & "</font>" & postrow
	else
		Response.Write prerowA & "<FONT size=1>" & strname & "</FONT>" &  midrowA & " <FONT size=1>" & strDetails & " </font>" &  midrowA & "<FONT size=1>" & strDeliverables & " </font>" & postrowA
	end if
	rs.MoveNext
  loop
  
  if rs.EOF and rs.BOF then
	if projectfound then
		'Say there were no requirements
		Response.Write "</table><strong>No Requirements are defined for this Product.</strong>"	
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
  cn.Close
  set cn=nothing
  set rs=nothing
%>
<br>
<br>
<font Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>

</body>
</html>
