<HTML>
<HEAD>
<%if request("ID") <> "" then%>
	<title>Update Employee</title>
<%else%>
	<title>Add Employee</title>
<%end if%>


</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="employeemain.asp?Source=<%=request("Source")%>&ID=<%=Request("ID")%>&NTName=<%=request("NTName")%>" >
</FRAMESET>
</HTML>
