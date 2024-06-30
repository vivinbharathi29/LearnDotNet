	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<title>Switch PM</title>


</HEAD>
<FRAMESET ROWS="*" ID=TopWindow>
	<FRAME ID="MyWindow" Name="MyWindow" SRC="ChangePMMain.asp?EmployeeID=<%=request("EmployeeID")%>&PMImpersonateID=<%=request("PMImpersonateID")%>&PMTypeID=<%=request("PMTypeID")%>">
</FRAMESET>
</HTML>
