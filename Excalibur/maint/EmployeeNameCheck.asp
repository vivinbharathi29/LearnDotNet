<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%
	dim cn
	dim rs
	dim strNameLetter
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	rs.Open "spGetEmployees",cn,adOpenStatic
	do while not rs.EOF
		strNameLetter = left(rs("Name"),1)
		if strNameLetter <> ucase(strNameLetter) then
			response.write rs("ID") & ":" & rs("Name") & "<BR>"
		end if	
		rs.MoveNext
	loop

	set rs = nothing
	set cn = nothing
%>

</BODY>
</HTML>
