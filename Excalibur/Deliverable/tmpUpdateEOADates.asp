<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<font size=3 face=verdana><b>Update EOAD ates</b><BR><BR></font>
<%
	dim cn 
	dim rs
	dim strSQl
	
	strSQl = "Select v.id as versionID, v.endoflifedate as versionEOA, v.active as VersionActive, s.* " & _
                "from datawarehouse.dbo.sheet1$ s with (NOLOCK), deliverableversion v with (NOLOCK) " & _
                "where s.DeliverableversionID = v.id " & _
                "and (v.endoflifedate is null or v.endoflifedate <> s.eoadate or s.active <> v.active) " 
	
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.eof
	    response.write "Update DeliverableVersion set active=" & replace(replace(rs("Active"),"True","1"),"False","0") & ", EndOfLifeDate='" & rs("EOADate") & "' where id=" &  rs("VersionID") & "<BR>"
	    
	    rs.movenext
	loop
	rs.Close
	set rs = nothing
	cn.Close
	set cn = nothing

%>
</BODY>
</HTML>


