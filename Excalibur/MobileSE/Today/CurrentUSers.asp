<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=ivory>

<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<tr>
<td>
<font size=2 face=verdana>	
<%
	dim strSQL
	dim UserArray
	dim ItemArray
	dim strItem
	dim strEmployeeDomain
	dim strEmployeeName
	dim EmployeeArray
	dim strEmployee
	dim strEmailList
	
	strSQL = ""
	strEmailList = ""
	UserArray = split(Application("ActiveUserNames"),",")
	for each strItem in UserArray
		if trim(strItem) <> "" then
			ItemArray = split(strItem,"-")
            if instr(ItemArray(0),"\")> 0 then
			EmployeeArray = split(ItemArray(0),"\")
			strSQL = strSQL & " or (domain = '" & EmployeeArray(0) & "' and ntname = '" & EmployeeArray(1) & "') " 
		    else
			    strSQL = strSQL & " or (email = '" & ItemArray(0) & "') " 
            end if
		end if
	next
	
	if strSQl <> "" then
		strSQL	= mid(strSQL,5)

		strConnect = Session("PDPIMS_ConnectionString")
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.Recordset")
		cn.ConnectionString = strConnect
		cn.Open
		rs.Open "Select distinct email from employee with (NOLOCK) where " & strSQL,cn,adOpenForwardOnly
		do while not rs.EOF
			strEmailList = strEmailList & ";" & rs("Email")
			rs.MoveNext
		loop
		rs.Close
		
		cn.Close
		
		if strEmailList <> "" then
			strEmailList = mid(strEmailList,2)
		end if
	end if
	
	Response.Write "<B>&nbsp;&nbsp;&nbsp;Active Sessions</b>"
	if trim(strEmailList) <> "" then
		Response.write "&nbsp;&nbsp;<a href=""mailto:" & strEmailList & """>Email These Users</a>"
	end if
	 
	Response.write Replace(Application("ActiveUserNames"),",","<BR>&nbsp;&nbsp;&nbsp;")

	

%>
</font>
</td>
</tr>
</table>
</BODY>
</HTML>
