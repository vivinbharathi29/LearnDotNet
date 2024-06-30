<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<META NAME="Generator" CONTENT="Microsoft Word 97">
<TITLE>Deliverable Comparison</TITLE>
</HEAD>
<BODY>
<H2><FONT size=5><STRONG>Deliverable Comparison</STRONG></FONT></H2>
<%
	Dim cn
	Dim rs
	dim rs2
	dim strDeliverables
	dim rowcount
	dim ReportName
	dim lastassembly
	dim ID
	dim strChanges
	dim strPart
	dim strOTS
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn  
	set rs2 = server.CreateObject("ADODB.recordset")
	rs2.ActiveConnection = cn  

	ID = request("ID")
	if not isnumeric(trim(ID)) then
		ID = 0
	end if
	rs.Open "SELECT Name From DeliverableRoot with (NOLOCK) Where id = " & clng(ID) & ";",cn,adOpenForwardOnly
	if rs.eof and rs.bof then
		Response.Write "Deliverable Not Found"
	else
		reportName = rs("Name")
		rs.close
%>
	

<H4><FONT face=Verdana><%=ReportName%></FONT><FONT size=1>
<HR color=#006697>
</H4>
<P><FONT face=Verdana>
<TABLE borderColor=skyblue cellSpacing=1 cellPadding=1 width="100%" 
bgColor=#e6f7ff border=1>
  
  <TR>
    <TD bgColor=#006697><STRONG><FONT color=white size=2>Version</FONT></STRONG></TD>
    <TD bgColor=#006697><STRONG><FONT color=white size=2>Developer</FONT></STRONG></TD>
    <TD bgColor=#006697><STRONG><FONT color=white size=2>Problem Reports</FONT></STRONG></TD>
    <TD bgColor=#006697><STRONG><FONT color=white size=2>Changes</FONT></STRONG></TD></TR>

    
    <%
		
  
	rs.Open "Select v.ID, Version, Revision, Pass, PartNumber,Changes, e.Name as Developer FROM DeliverableVersion v with (NOLOCK), Employee e with (NOLOCK) WHERE e.id = v.DeveloperID and DeliverableRootID =  " & clng(ID) & " Order by Version, Revision, Pass;",cn,adOpenForwardOnly
		
	do while not rs.EOF
		strChanges = rs("Changes") & ""
		if trim(strchanges) = "" then
			strchanges = "&nbsp;"
		end if
		strpart = rs("partNumber") & ""
		if trim(strpart) = "" then
			strpart = "&nbsp;" 
		else
			strpart = " (" & strpart & ")"
		end if

        rs2.Open "spGetOTSByDelVersion "  & clng(rs("ID")), cn, adOpenForwardOnly
		strots = ""
		do while not rs2.EOF
			strots = strots & rs2("OTSNumber") & " - " & rs2("shortdescription") & " (Priority: " & rs2("Priority") & ")" &  "<br>"
			rs2.MoveNext		
		loop
		rs2.Close

		if trim(strots)  ="" then
			strots = "&nbsp;"
		end if
		
		Response.Write "<TR><TD><font face=verdana size = 1>" & rs("Version") & "," & rs("Revision") & "," & rs("Pass")  & strpart  & "</font></TD><TD><font face=verdana size = 1>" & rs("Developer") & "</FONT></TD><TD><font face=verdana size = 1>" & strots & " </FONT></TD><TD><font face=verdana size = 1>" & strChanges & "</FONT></TD></TR>"
		
		rs.movenext
	loop
	
    rs.close
    cn.Close
    %>

    </TABLE>
</FONT></P>
<%
	end if
%>
</BODY>
</HTML>
