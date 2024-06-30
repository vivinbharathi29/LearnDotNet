<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	dim CurrentUser
	CurrentUser = lcase(Session("LoggedInUser"))

	%>


<HTML>
<TITLE>Part Numbers</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,50" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="DisplayPartNumberMain.asp?RootID=<%=request("RootID")%>&VersionID=<%=request("VersionID")%>&app=<%=Request("app")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="DisplayPartNumberbutton.asp?app=<%=Request("app")%>" scrolling=no>
</FRAMESET>

</HTML>
