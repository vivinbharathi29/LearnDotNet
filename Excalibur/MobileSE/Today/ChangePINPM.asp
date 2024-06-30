	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<title>Switch PIN PM</title>


</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="ChangePINPMMain.asp?EmployeeID=<%=request("EmployeeID")%>&PINPMImpersonateID=<%=request("PINPMImpersonateID")%>">
</FRAMESET>
</HTML>
