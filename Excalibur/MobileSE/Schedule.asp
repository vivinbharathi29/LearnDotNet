<%@ Language=VBScript %>
<html>
<head>
<%if instr(request("ID"),"_") > 0 then%>
<title>Compare Product Schedules - HP Restricted</title>
<%else%>
<title><%= request("Product")%> Schedule - HP Restricted</title>
<%end if%>
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<body>
<table border="0" cellPadding="1" cellSpacing="1" width="100%">
  
  <tr>
    <td Width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
    <td><img SRC="images/ProductScheduleReport.gif" WIDTH="283" HEIGHT="50"></td>
    <td Align="Right"><img SRC="images/programoffice.gif" WIDTH="283" HEIGHT="50"></td></tr></table>

<hr>
<h1>

<%
	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 



  Dim blnCompare
  dim strIDs
  if instr(request("ID"),"_") > 0 then
	Response.Write "Compare Product Schedules" 
	strIDs = "(" & replace(scrubsql(request("ID")),"_",",") & ")"
	blncompare = true
  else
	Response.Write request("Product") & " Schedule" 
	blncompare = false
  end if
%>
</h1>

<%

  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

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
  
  'Setup table row HTML for primary color cells
  PreRowA = "<TR  bgcolor=white valign=top><TD>"
  MidRowA = "</TD><TD valign=top><FONT size=1>"
  PostRowA = "</FONT></TD></TR>"

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=#e6f7ff valign=top><TD>"
  MidRow = "</TD><TD valign=top ><FONT size=1>"
  PostRow = "</FONT></TD></TR>"

  'Create a recordset
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn
  if blncompare then
	rs.Open "SELECT ml.ID as MilestoneID, f.Name + ' ' + v.version as Product, m.comments as COmments,ml.Milestone as Milestone, m.POR as POR, m.ActualDate as ActualDate, m.TargetDate as TargetDate FROM ProductFamily f with (NOLOCK), ProductVersion v with (NOLOCK), Milestone m with (NOLOCK), MilestoneLookup ml with (NOLOCK) WHERE f.ID = v.productfamilyID and ml.id = m.milestoneid and m.Productversionid = v.id and v.id in " & strids & " ORDER By ml.orderid, f.Name, v.version;",,adOpenStatic
  else
	if request("active") then
		rs.Open "SELECT m.comments as COmments,ml.Milestone as Milestone, m.POR as POR, m.ActualDate as ActualDate, m.TargetDate as TargetDate FROM ProductVersion pv with (NOLOCK), Milestone m with (NOLOCK), MilestoneLookup ml with (NOLOCK) WHERE ml.id = m.milestoneid and m.Productversionid = pv.id and pv.id = " & clng(request("ID")) & "and pv.Active = 1 and m.Active = 1 ORDER By ml.orderid;",,adOpenStatic
	else
		rs.Open "SELECT m.comments as COmments,ml.Milestone as Milestone, m.POR as POR, m.ActualDate as ActualDate, m.TargetDate as TargetDate FROM ProductVersion pv with (NOLOCK), Milestone m with (NOLOCK), MilestoneLookup ml with (NOLOCK) WHERE ml.id = m.milestoneid and m.Productversionid = pv.id and pv.id = " & clng(request("ID")) & " and m.Active = 1 ORDER By ml.orderid;",,adOpenStatic
	end if
  end if
  'Initializae the cell backgroud color selection
  primarycolor = false
  
  'Start Table
  if not(rs.EOF and rs.BOF) then%>	  
	<table bgColor="#e6f7ff" border="1" borderColor="skyblue" cellPadding="2" cellSpacing="1" width="95%">
	  <tr bgColor="#006697">
	    <td nowrap width="45%"><strong><font color="white">Milestone</font></strong></td>
<%  if blncompare then%>
		<td nowrap width="100"><font color="white"><strong>Product</strong></td>
<%	end if%>
	<td nowrap width="50"><font color="white"><strong>POR</strong></td>
	<td nowrap width="50"><font color="white"><strong>Target</strong></td>
	<td nowrap width="50"><font color="white"><strong>Actual</strong></td>
	<td width="100%"><font color="white"><strong>Comments</strong></td></tr>
<%end if  
  
  'Display requirements
  LastMilestone = ""
  do while not rs.EOF
	primarycolor = not primarycolor
	strComments = replace(replace(replace(rs("Comments") & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>") & "&nbsp;"
	strMilestone = replace(replace(replace(rs("Milestone") & "","<","&lt;"),">","&gt;"),vbcrlf,"<BR>")& "&nbsp;"
	strPOR = rs("POR")& "&nbsp;"
	strActual = rs("ActualDate")& "&nbsp;"
	strTarget = rs("TargetDate")
	strexpired = ""
	if stractual = "&nbsp;" then
		if isdate(strtarget)then
			if datediff("d",strtarget,now)>= 0 then
				strExpired = " <IMG SRC=""images/alert.gif"">"
			end if
		end if
	end if
	strtarget = strtarget & "&nbsp;"
	if blncompare then
		strproduct = midrow  & "<FONT size=1>" & rs("Product") & "</FONT>" 
		if primarycolor then%>
			<tr bgcolor="#e6f7ff" valign="top">
<%		else%>
			<tr bgcolor="white" valign="top">
<%		end if			
			if lastmilestone <> strMilestone then
				lastmilestone = strmilestone
				set rscount = server.CreateObject("ADODB.recordset")
				rsCount.ActiveConnection = cn
				rsCount.Open "SELECT Count(*) as Rows FROM ProductVersion pv with (NOLOCK), Milestone m with (NOLOCK), MilestoneLookup ml with (NOLOCK) WHERE ml.id = m.milestoneid and m.Productversionid = pv.id and pv.id in " & strids & " and m.milestoneid = " & rs("MilestoneID") & " ;",,adOpenStatic
				rowcount = rsCount("rows")
				rsCount.Close
				set rscount = nothing%>
				<td rowspan="<%= rowcount%>" bgcolor="<%if primarycolor then Response.Write "#e6f7ff" else Response.Write "white" end if%>"><font size="1"><%= strMilestone%></font><%= strproduct &  strexpired &  midrow%>
<%			else%>
				<%= strproduct & strexpired &  midrow%>
<%			end if%>
			<font size="1"><%= strPOR%></font><%= midrow%><font size="1"><%= strTarget%></font><%= midrow%><font size="1"><%= strActual%></font><%= midrow%><font size="1"><%= strComments%></font>
<%			Response.Write postrow
	else
		strproduct = ""
		if primarycolor then
			Response.Write prerow
		else
			Response.Write prerowa
		end if%>
		<font size="1"><%= strMilestone%> <%= strexpired%></font><%= strproduct &  midrow%><font size="1"><%= strPOR%></font><%= midrow%><font size="1"><%= strTarget%></font><%= midrow%><font size="1"><%= strActual%></font><%= midrow%><font size="1"><%= strComments%></font>
<%			Response.Write postrow
	end if
	rs.MoveNext
  loop
  
  if rs.EOF and rs.BOF then
	Response.Write "</TBODY></table></p><FONT face=verdana size = 2><strong>No Schedule is defined for this product.</strong></font>"	
  else
   'Finish off table 
    Response.Write "</TBODY></table></p>"
  end if
  
  'Cleanup
  rs.Close
  cn.Close
  set cn=nothing
  set rs=nothing
  Response.Write "</TBODY></table></p>"
%>
<br>
<br>
<font size="1">Report Generated <%=formatdatetime(date(),vblongdate) %></font>
<br>
<br>
<font Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>
</body>
</html>
