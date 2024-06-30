<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<title>Configure Today Page</title>
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="configuremain.asp?Tab=<%=request("Tab")%>">
	<FRAME ID="LowerWindow" Name="LowerWindow" SRC="configurebuttons.asp">
</FRAMESET>
</HTML>
