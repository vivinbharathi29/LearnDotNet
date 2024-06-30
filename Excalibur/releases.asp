<%@ Language=VBScript %>
<HTML>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<BODY>

<FONT face=verdana>
<h3>Excalibur Releases</h3>
<font size=1 face=verdana>Report Period: (<%=Date() - 7%> - <%=Date() -1 %>)<BR><BR></font>
<P>

<font size=3 face=verdana><b>Mobile Application Development</b></font><hr>
<font size=2 face=verdana><b>Deliverables Pre-Released:</b></font>
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%
	dim strVendor
	dim strVersion
	dim i
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spListPreReleasesThisWeek",cn,adOpenForwardOnly
	i=0
	do while not rs.EOF
		if rs("workflowid") = 15 then 'rs("DevManagerID") = 855 or rs("DevManagerID") = 548 or rs("DevManagerID") = 437 or rs("DevManagerID") = 719 or rs("DevManagerID") = 463 or rs("DevManagerID") = 467 or rs("DevManagerID") = 433 then	
			i=i+1
			if rs("VendorID") =  200 or rs("VendorID") = 219 then
				strVendor = ""
			else
				strVendor = " (3rd party)"
			end if
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("revision")		
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")		
			end if
			Response.Write "<TR><TD><FONT face=verdana size=1>" & rs("Name") & " " & strVersion & strVendor & "</FONT></TD></TR>"
		end if
		rs.MoveNext
	loop
	Response.Write ""
	rs.close	
	if i=0 then
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR>"
	end if
%>	  
   </Table>
<BR>
<font size=2 face=verdana><b>Deliverables Released To Functional Test (but not pre-released yet):</b></font>
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%
	
	rs.Open "spListReleasesInTestThisWeek",cn,adOpenForwardOnly
	i=0
	do while not rs.EOF
		if rs("workflowid") = 15 then 'or rs("DevManagerID") = 855 or rs("DevManagerID") = 548 or rs("DevManagerID") = 437 or rs("DevManagerID") = 719 or rs("DevManagerID") = 463 or rs("DevManagerID") = 467 or rs("DevManagerID") = 433 then	
			i=i+1
			if rs("VendorID") =  200 or rs("VendorID") = 219 then
				strVendor = ""
			else
				strVendor = " (3rd party)"
			end if
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("revision")		
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")		
			end if
			Response.Write "<TR><TD><FONT face=verdana size=1>" & rs("Name") & " " & strVersion & strVendor & "</FONT></TD></TR>"
		end if
		rs.MoveNext
	loop
	Response.Write ""
	rs.Close
	if i=0 then
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR>"
	end if

%>	  
   </Table>
  
  <BR><BR><BR><font size=3 face=verdana><b>Other Groups</b></font><hr>
<font size=2 face=verdana><b>Deliverables Pre-Released:</b></font>
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open "spListPreReleasesThisWeek",cn,adOpenForwardOnly
	i=0
	do while not rs.EOF
		if rs("workflowid") <> 15  then 'not (rs("DevManagerID") = 855 or rs("DevManagerID") = 548 or rs("DevManagerID") = 437 or rs("DevManagerID") = 719 or rs("DevManagerID") = 463 or rs("DevManagerID") = 467or rs("DevManagerID") = 433) then
			i=i+1
			if rs("VendorID") =  200 or rs("VendorID") = 219 then
				strVendor = ""
			else
				strVendor = " (3rd party)"
			end if
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("revision")		
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")		
			end if
			Response.Write "<TR><TD><FONT face=verdana size=1>" & rs("Name") & " " & strVersion & strVendor & "</FONT></TD></TR>"
		end if
		rs.MoveNext
	loop
	Response.Write ""
	rs.close	
	if i=0 then
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR>"
	end if
%>	  
   </Table>
<BR>
<font size=2 face=verdana><b>Deliverables Released To Functional Test (but not pre-released yet):</b></font>
<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%
	
	rs.Open "spListReleasesInTestThisWeek",cn,adOpenForwardOnly
	i=0
	do while not rs.EOF
		if rs("workflowid") <> 15  then 'not (rs("DevManagerID") = 855 or rs("DevManagerID") = 548 or rs("DevManagerID") = 437 or rs("DevManagerID") = 719 or rs("DevManagerID") = 463 or rs("DevManagerID") = 467 or rs("DevManagerID") = 433) then
			i=i+1
			if rs("VendorID") =  200 or rs("VendorID") = 219 then
				strVendor = ""
			else
				strVendor = " (3rd party)"
			end if
			strVersion = rs("version") & ""
			if rs("Revision") & "" <> "" then
				strVersion = strVersion & "," & rs("revision")		
			end if
			if rs("Pass") & "" <> "" then
				strVersion = strVersion & "," & rs("Pass")		
			end if
			Response.Write "<TR><TD><FONT face=verdana size=1>" & rs("Name") & " " & strVersion & strVendor & "</FONT></TD></TR>"
		end if
		rs.MoveNext
	loop
	Response.Write ""
	rs.Close
	set rs = nothing
	set cn  = nothing
	if i=0 then
		Response.Write "<TR><TD><font size=1 face=verdana>none</font></TD></TR>"
	end if
%>	  
   </Table>
  
</font>
</BODY>
</HTML>