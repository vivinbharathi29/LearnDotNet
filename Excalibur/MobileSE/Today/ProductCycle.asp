<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Choose Product Group</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProductCycleMain.asp?ID=<%=request("ID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProductCycleButtons.asp">
</FRAMESET>

</HTML>
