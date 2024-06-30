<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
<TITLE>Excalibur Usage Summary</TITLE>
<meta name="VI60_defaultClientScript" content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

</head><link rel="stylesheet" type="text/css" href="Style/programoffice.css">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {

}



//-->
</SCRIPT>
<body>
<%
	dim cn, rs, cmd, prm

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.IsolationLevel=256
	cn.Open
    response.write "<font size=3 face=verdana><b>Product1 PRL</b></font><br><br>"

    response.write "<font size=2 face=verdana><b><u>Hard Drives</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"
    response.write "<font size=2 face=verdana><b><u>Memory</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"

    response.write "<font size=2 face=verdana><b><u>Application Software</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"



    Response.Write "<br><br>" & "---------------------------------------------------------------------------" & "<br><br>"

    response.write "<font size=3 face=verdana><b>Select Hard Drives</b></font><br><br>"

    response.write "<fieldset><legend>Filters</legend>"
    response.write "<table>"
    response.write "<tr><td><font size=2 face=verdana><b>Capicity:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>1TB</option><option>2TB</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Speed:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>5400 RPM</option><option>7200 RPM</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Thickness:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>7mm</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Interface:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>SATA</option></select></td></tr>"
    response.write "</table></fieldset><br>"
    
    response.write "<fieldset><legend>Available Hard Drives</legend>"
    response.write "<br>&nbsp;&nbsp;<a href="""">Request New Hard Drive</a> | <a href="""">Create Hard Drive Combo</a><br><br>"
    response.write "<table cellpadding=2 cellspacing=0  border=1 width=""600px"">"
    response.write "<tr>"
    response.write "<td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td>"
    response.write "<td><font size=2 face=verdana><b>Name</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Capicity</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Speed</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Thickness</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Interface</b></font></td>"
    response.write "</tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 500GB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>500GB</font></td><td><font size=2 face=verdana>5400 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 500GB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>500GB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>5400 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB (2x500GB) 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "</table></fieldset>"


    Response.Write "<br><br>" & "---------------------------------------------------------------------------" & "<br><br>"

    response.write "<font size=3 face=verdana><b>Request New Hard Drive</b></font><br><br>"

    
    response.write "<table>"
    response.write "<tr><td><font size=2 face=verdana><b>Request Name:</b></font></td><td><input style=""width:200px;"" id=""Text1"" type=""text"" /></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Capicity:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>1TB</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Speed:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>5400 RPM</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Thickness:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>7mm</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Interface:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>SATA</option></select></td></tr>"
    response.write "</table><br>"
    response.write "<input id=""Button1"" type=""button"" value=""Save"" />"



    Response.Write "<br><br>" & "---------------------------------------------------------------------------" & "<br><br>"

    response.write "<font size=3 face=verdana><b>Select Hard Drives</b></font><br><br>"

    response.write "<fieldset><legend>Filters</legend>"
    response.write "<table>"
    response.write "<tr><td><font size=2 face=verdana><b>Capicity:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>1TB</option><option>2TB</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Speed:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>5400 RPM</option><option>7200 RPM</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Thickness:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>7mm</option></select></td></tr>"
    response.write "<tr><td><font size=2 face=verdana><b>Interface:</b></font></td><td><select style=""width:200px;"" id=""Select1""><option></option><option>SATA</option></select></td></tr>"
    response.write "</table></fieldset><br>"
    
    response.write "<fieldset><legend>Available Hard Drives</legend>"
    response.write "<br>&nbsp;&nbsp;<a href="""">Request New Hard Drive</a> | <a href="""">Create Hard Drive Combo</a><br><br>"
    response.write "<table cellpadding=2 cellspacing=0  border=1 width=""600px"">"
    response.write "<tr>"
    response.write "<td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td>"
    response.write "<td><font size=2 face=verdana><b>Name</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Capicity</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Speed</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Thickness</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Interface</b></font></td>"
    response.write "</tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 500GB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>500GB</font></td><td><font size=2 face=verdana>5400 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 500GB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>500GB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>5400 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana color=red>HDD 2TB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana color=red>1TB</font></td><td><font size=2 face=verdana  color=red>7200 RPM</font></td><td><font size=2 face=verdana color=red>7mm</font></td><td><font size=2 face=verdana color=red>SATA</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB (2x500GB) 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>1TB</font></td><td><font size=2 face=verdana>7200 RPM</font></td><td><font size=2 face=verdana>7mm</font></td><td><font size=2 face=verdana>SATA</font></td></tr>"
    response.write "</table></fieldset>"


    Response.Write "<br><br>" & "---------------------------------------------------------------------------" & "<br><br>"
    response.write "<font size=3 face=verdana><b>Product1 PRL</b></font><br><br>"

    response.write "<font size=2 face=verdana><b><u>Hard Drives</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"

    response.write "<table style=""margin-left: 15px"" cellpadding=2 cellspacing=0  border=1 width=""600px"">"
    response.write "<tr>"
    response.write "<td><font size=2 face=verdana><b>Requests</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Roots</b></font></td>"
    response.write "</tr>"
    response.write "<tr><td><font size=2 face=verdana>HDD 500GB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>HDD 500GB 5400RPM 7mm SATA</font></td></tr>"
    response.write "<tr><td><font size=2 face=verdana color=red>HDD 2TB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana color=red>Requested</font></td></tr>"
    response.write "<tr><td><font size=2 face=verdana>HDD 1TB (2x500GB) 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>HDD 250GB 7200RPM 7mm SATA</font></td></tr>"

    response.write "</table><br><br>"

    response.write "<font size=2 face=verdana><b><u>Memory</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"

    response.write "<font size=2 face=verdana><b><u>Application Software</u></b></font><br><br>"
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;<a href="""">Choose</a><br><br>"

    Response.Write "<br><br>" & "---------------------------------------------------------------------------" & "<br><br>"
    response.write "<font size=3 face=verdana><b>Today Page</b></font><br><br>"

    response.write "<div style=""background-color:gainsboro;width:100%""><font size=2 face=verdana><b>Marketing Requests</b></font></div><br>"

    response.write "<font size=2 face=verdana><b><u>Roots Requested for Products</u></b></font><br><br>"
    response.write "<a href="""">Approve</a> | <a href="""">Reject</a><br><br>"

    response.write "<table style=""margin-left: 15px"" cellpadding=2 cellspacing=0  border=1 width=""600px"">"
    response.write "<tr>"
    response.write "<td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td>"
    response.write "<td><font size=2 face=verdana><b>Request</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Product</b></font></td>"
    response.write "</tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 500GB 5400RPM 7mm SATA</font></td><td><font size=2 face=verdana>Product1</font></td></tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana>HDD 1TB (2x500GB) 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana>Product1</font></td></tr>"

    response.write "</table><br><br>"
    response.write "<font size=2 face=verdana><b><u>New Roots Requested</u></b></font><br><br>"
    response.write "<a href="""">Link to New Root</a> | <a href="""">Link to Existing Root</a> | <a href="""">Reject</a><br><br>"
    response.write "<table style=""margin-left: 15px"" cellpadding=2 cellspacing=0  border=1 width=""600px"">"
    response.write "<tr>"
    response.write "<td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td>"
    response.write "<td><font size=2 face=verdana><b>Request</b></font></td>"
    response.write "<td><font size=2 face=verdana><b>Product</b></font></td>"
    response.write "</tr>"
    response.write "<tr><td width=""10px""><input id=""Checkbox1"" type=""checkbox"" /></td><td><font size=2 face=verdana color=red>HDD 2TB 7200RPM 7mm SATA</font></td><td><font size=2 face=verdana color=red>Product1</font></td></tr>"

    response.write "</table><br><br>"
    

    rs.open "Select * from requirement",cn
    do while not rs.eof
       ' response.write rs("name") &  "<br>"
        rs.movenext
    loop
    rs.close
    cn.close
    set cn = nothing
	
%>


</body>
</html>
