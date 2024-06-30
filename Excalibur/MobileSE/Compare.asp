<%@ Language=VBScript %>
<html>
<head>
<title>Product Comparison - HP Restricted</title>
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</head>
<script ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

sub export_onclick

'	if msgbox("Do you have Excel 2000 installed?", 36, "Confirmation") = 6 Then
'		myrows = mytable.rows.length
'		for i = 0 to (myrows - 1)
'			mycols = mytable.rows(i).cells.length
'			for j = 0 to (mycols - 1)
'				myValue = "'" & mytable.rows(i).cells(j).innertext
'				myValue = replace(myValue, vbcrlf, vblf)
'				Excel.Cells((i+1),(j+1)).Value = myValue
'			next
'		next
'
'		with excel
'		  .Rows(1).Select
'		  .Selection.Font.Bold = True
'		  .Columns(1).Select
'		  .Selection.Font.Bold = True
'		  .ActiveSheet.Calculate
'		  .ActiveSheet.Export
'		end with
'	End If
	
End Sub

Sub Export_onmouseover
'	window.event.srcElement.style.cursor = "hand"
End Sub

-->
</script>
<body>
<table border="0" cellPadding="1" cellSpacing="1" width="100%">
  <tr>
    <td Width="180"><img src="images/LOGO.gif" WIDTH="67" HEIGHT="58"></td>
    <td><img SRC="images/ProduceComparisonDocument.gif" WIDTH="361" HEIGHT="50"></td>
    <td Align="Right"><img SRC="images/programoffice.gif" WIDTH="283" HEIGHT="50"></td></tr></table>
<table width="100%">
  <tr>
    <td><h2><%
    
	If request("Title") = "" Then%>
      <%if request("Display") = "Requirements" then%>
	    Compare Requirements
      <%elseif request("Display") = "Deliverables" then%>
	    Compare Deliverables
	  <%else%>
		Compare Requirements and Deliverables
	  <%end if
	Else
		Response.Write request("Title")
	End If %></h2>
	<%
	if request("Delta") = "Y" then
	Response.Write "<h3>Display Differences Only</h3>"
	end if
	
	%>
	</td>
<%'    <td valign="top"><p align="right">Printable Version <img SRC="../images/print.gif" align="absmiddle" WIDTH="24" HEIGHT="24"></p></td></table>%>

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


	dim IDList
	dim ID
	dim i
	dim rs
	DIM strSql
	Dim HeaderCount
	Dim LastPointer
	dim headerlist()
	dim IDinList
	dim strOutput()
	dim strStatus
	Dim PreRow
	Dim MidRow
	Dim PostRow
	Dim PreRowA
	Dim MidRowA
	Dim PostRowA
	Dim DarkColor
	Dim MedColor
	Dim LightColor
	
	If request("Print") Then
		DarkColor = "Black"
		MedColor = "Gray"
		LightColor = "Silver"
	Else
		DarkColor = "#006697"
		MedColor = "#e6f7ff" '"SkyBlue"
		LightColor = "white" '"#e6f7ff"
	end if


  'Setup table row HTML for primary color cells
  PreRowA = "<TR  bgcolor=" & lightcolor & " valign=top> <TD><FONT size=2>"
  MidRowA = "</FONT></TD><TD valign=top style=""TEXT-ALIGN: left""><FONT size=1>"
  PostRowA = "</FONT></TD></TR>"

  'Setup table row HTML for seconday color cells
  PreRow = "<TR bgcolor=" & lightcolor & " valign=top><TD><FONT size=2>"
  MidRow = "</FONT></TD><TD valign=top style=""TEXT-ALIGN: left""><FONT size=1>"
  PostRow = "</FONT></TD></TR>"
		
	IDList = request("ID") & "_"
	
%>
<!---<h3>Export to Excel 2000 <img id="export" SRC="images/ICON-DOC-EXCEL.GIF" WIDTH="16" HEIGHT="16"></h3>--->

<table id="mytable" bgColor="<%= lightcolor%>" border="1" borderColor="<%= darkcolor%>" cellPadding="2" cellSpacing="1" width="95%">
  <tr bgcolor="<%= darkcolor%>"><td width="20%">&nbsp;</td>

<%	lastPointer = 1
	do while instr(lastPointer,idlist,"_") > 0
		LastPointer = instr(lastPointer,idlist,"_") + 1
		HeaderCount = Headercount + 1
	loop
	
	redim HeaderList(headercount)
	redim strOutput (headercount)

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	headercount = 0
	Do while len(IDList) > 0
		ID = left(IDList,instr(idlist,"_") - 1)
		
		strSQl = "SELECT v.id, f.name + ' ' + v.version as Name, v.prdreleased, v.pddreleased FROM ProductVersion v with (NOLOCK),ProductFamily f with (NOLOCK) Where f.id = v.ProductFamilyID and v.ID = " & clng(ID)
		rs.ActiveConnection = cn
		rs.Open strSQL

		headerlist(headercount) = rs("Name")
		if isnull(rs("PRDReleased")) and  isnull (rs("PDDReleased")) then
			strStatus = "PRD: In Progress<BR>PDD: In Progress"
		elseif not isnull(rs("PDDReleased")) then
			strStatus = "PRD: Released<BR>PDD: Released"
		else
			strstatus  = "PRD: Released<BR>PDD: In Progress"
		end if
		headercount = headercount + 1
		IDinlist = IDinList & "," & rs("ID")
%>
   <td width="300"><font color="white"><strong><%= rs("name")%></strong><br><font size="1"><%= strstatus%></font></font></td>
<%
    	rs.Close

		idlist = mid(idlist,instr(idlist,"_") + 1)	
	loop
%>
	</tr>
<%
	if len(idinlist) > 0 then
		idinlist = mid(idinlist,2)
		idinlist = scrubsql(idinlist)
	end if
  dim PrimaryColor 
  dim LastRequirement
  dim lastcategory
  dim strRequirement
  dim strSpecification
  dim strDeliverable
  dim ReqStart
  dim PreviousReq
  dim blnDifferent
   dim temp
	
'  strSQl = "Select Platform, Deliverables, Requirement, Specification, Category from DeliverablesByPlatformID with (NOLOCK) where Platformid in (" & idinlist & ") Order By DisplayOrder, Requirement;"
 ' strSQl = "Select Platform, Deliverables, Requirement, Specification, Category from DeliverablesByPlatformID with (NOLOCK) where Platformid in (" & idinlist & ") Order By DisplayOrder, Requirement;"
  
  strSQL = "select f.name + ' ' + v.version as product, v.id as ProductID, pr.Deliverables, r.name as Requirement, r.id as ReqID, pr.Specification, c.name as Category FROM ProductFamily f with (NOLOCK), ProductVersion v with (NOLOCK), Product_Requirement pr with (NOLOCK), Requirement r with (NOLOCK), Category c with (NOLOCK) WHERE v.productfamilyid = f.id and v.ID = pr.Productid and pr.requirementid = r.id	and r.categoryid = c.id and v.id in (" & idinlist & ") Order By DisplayOrder, Requirement;"
  
  rs.ActiveConnection = cn
  rs.Open strSQL
  LastRequirement = ""
  Lastcategory = ""
  PrimaryColor = true
  do while not rs.eof

	set rs2 = server.CreateObject("ADODB.recordset")
	rs2.Open "spListDeliverablesByRequirement " & clng(rs("ReqID")) & "," & clng(rs("ProductID")),cn,adOpenForwardOnly
	strDeliverable = ""
	do while not rs2.EOF
		strDist = ""
		if rs2("Preinstall") then
			strDist = strDist & ",Preinstall"
		end if
		if rs2("DropInBox") then
			strDist = strDist & ",DIB"
		end if
		if rs2("Preload") then
			strDist = strDist & ",Preload"
		end if
		if rs2("Web") then
			strDist = strDist & ",Web"
		end if
		if rs2("ARCD") then
			strDist = strDist & ",DRCD"
		end if
		if rs2("DRDVD") then
			strDist = strDist & ",DRDVD"
		end if
		if rs2("RACD_Americas") then
			strDist = strDist & ",RACD-Americas"
		end if
		if rs2("RACD_EMEA") then
			strDist = strDist & ",RACD-EMEA"
		end if
		if rs2("RACD_APD") then
			strDist = strDist & ",RACD-APD"
		end if
		
		if rs2("OSCD") then
			strDist = strDist & ",OSCD"
		end if
		if rs2("DocCD") then
			strDist = strDist & ",DocCD"
		end if

		if trim(rs2("Patch")&"") <> "0" then
			strDist = strDist & ",Patch"
		end if
		
		if rs2("SelectiveRestore") then
			strDist = strDist & ",Selective Restore"
		end if
				
		if strDist <> "" then
			strDist = mid(strDist,2)
			strDeliverable = strDeliverable & "- " & replace(replace(rs2("Name"),"<","&lt;"),">","&gt;") & " <font color=green>[" & strDist & "]</font>" & vbcrlf
		else
			strDeliverable = strDeliverable & "- " & replace(replace(rs2("Name"),"<","&lt;"),">","&gt;") & vbcrlf
		end if
		rs2.MoveNext	
	loop
	rs2.close
	set rs2 = nothing	
	'strDeliverable = replace(strDeliverable,"<","&lt;")
	'strDeliverable = replace(strDeliverable,">","&gt;")
	strDeliverable = replace(strDeliverable,chr(13) + chr(10),"<br>")
	if trim(strDeliverable) = "" then
		strDeliverable = "&nbsp;"
	end if
	do while right(strDeliverable,4) = "<br>" and len(strDeliverable) > 4
		strDeliverable = left(strDeliverable,len(strDeliverable)-4)
	loop

	strRequirement = replace(rs("Requirement") & "","<","&lt;")
	strRequirement = replace(strrequirement,">","&gt;")
	strRequirement = replace(strrequirement,chr(13) + chr(10),"<BR>")

	strspecification = rs("Specification") & ""
	'strSpecification = replace(strspecification,"<","&lt;")
	'strSpecification = replace(strSpecification,">","&gt;")
	strSpecification = replace(strSpecification,chr(13) + chr(10),"<br>")
	if trim(strspecification) = "" then
		strspecification = "&nbsp;"
	end if
	do while right(strspecification,4) = "<br>" and len(strspecification) > 4
		strspecification = left(strspecification,len(strspecification)-4)
	loop


	
	if lastrequirement <> strRequirement then
		if primarycolor then
			if lastrequirement <> "" then
				blnDifferent = false
				for i = 0 to headercount -2
					if ucase(trim(stroutput(i))) <> ucase(trim(stroutput(i+1))) then
						blndifferent = true
						exit for
					end if
				next
				if blnDifferent or request("Delta") <> "Y" then
					Response.write strReqStart 
					for i = 0 to headercount -1
						Response.write  midrowA  & stroutput(i)
					next
					Response.Write PostrowA
				else
					strReqStart = ""					
				end if
			end if
			for i = 0 to headercount -1
				stroutput(i) = "&nbsp;"
			next
			if lastCategory <> rs("Category") then%>
					<tr bgcolor="<%= medcolor%>" valign="top"><td><strong><%= rs("Category")%> </strong></td>
<%				for i = 1 to headercount%>
					<td valign="top" style="TEXT-ALIGN: left"><font size="1">&nbsp;</font></td>
<%				next%>
				</tr>
			
<%				lastcategory = rs("Category")
			end if
			strReqStart = PreRow & " &nbsp; &nbsp;<font size=1>" & strRequirement & " </font>"
			lastrequirement = strRequirement
			primarycolor = false
		else
			if lastrequirement <> "" then
				blnDifferent = false
				for i = 0 to headercount -2
					if ucase(trim(stroutput(i))) <> ucase(trim(stroutput(i+1))) then
						blndifferent = true
						exit for
					end if
				next
				if blnDifferent or request("Delta") <> "Y" then
			
					Response.Write strReqStart
					for i = 0 to headercount -1
						Response.write  midrow   & stroutput(i)
					next
					Response.Write Postrow
				else
					strReqStart = ""
				end if
			end if
			for i = 0 to headercount -1
				stroutput(i) = "&nbsp;"
			next
			if lastCategory <> rs("Category") then%>
				<tr bgcolor="<%=medcolor%>" valign="top"><td><strong><%= rs("Category")%> </strong>
<%				for i = 1 to headercount%>
					</td><td valign="top" style="TEXT-ALIGN: left"><font size="1">&nbsp;
<%				next%>
				</font></td></tr>
<%				lastcategory = rs("Category")
			end if
			
			
			strReqStart =  PreRowA & " &nbsp; &nbsp;<font size=1>" & strRequirement & "</font>"
			lastrequirement = strRequirement
			primarycolor = true		
			
		end if
	end if



'	if primarycolor then
'		Response.write midrow &  strSpecification
'	else
'		Response.write midrowA &  strSpecification
'	end if
	for i = 0 to headercount -1
		if headerlist(i) = rs("Product") then
			if request("Display") = "Requirements" then
				stroutput(i) = strSpecification
			elseif request("Display") = "Deliverables" then
				stroutput(i) = strDeliverable				
			else
				if strspecification = "&nbsp;" then
					strSpecification = "[Specification To Be Defined]"
				end if
				if strdeliverable = "&nbsp;" then
					strdeliverable = "[Deliverable List To Be Defined]"
				end if
				stroutput(i) = strSpecification & "<BR><HR color=Skyblue>" & strDeliverable
			end if
		end if
	next
	rs.movenext
  loop
  if primarycolor then
	blnDifferent = false
	for i = 0 to headercount -2
		if ucase(trim(stroutput(i))) <> ucase(trim(stroutput(i+1))) then
			blndifferent = true
			exit for
		end if
	next
	if blnDifferent or request("Delta") <> "Y" then

		Response.Write strreqstart
		for i = 0 to headercount -1
			Response.write midrow &  stroutput(i)
		next
		Response.Write Postrow
	else
		strreqstart = ""
	end if
  else
	blnDifferent = false
	for i = 0 to headercount -2
		if ucase(trim(stroutput(i))) <> ucase(trim(stroutput(i+1))) then
			blndifferent = true
			exit for
		end if
	next
	if blnDifferent or request("Delta") <> "Y" then
  
	Response.Write strreqstart
	for i = 0 to headercount -1
		Response.write midrowA &  stroutput(i)
	next
  
	Response.Write PostrowA
	else
		strreqstart = ""
	
	end if
  end if
  cn.Close
%>
</table>
<br>
<br>
<br>
<font Size="2" Color="red"><p><strong>HP Restricted</strong></p></font>


</body>
</html>
