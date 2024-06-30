<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Share Profile</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProfileSharePropertiesMain.asp?ProfileID=<%=request("ProfileID")%>&EmployeeID=<%=Request("EmployeeID")%>&CanEdit=<%=Request("CanEdit")%>&CanDelete=<%=Request("CanDelete")%>&AddType=<%=Request("AddType")%>">
</FRAMESET>

</HTML>