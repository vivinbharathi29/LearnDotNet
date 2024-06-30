<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Update Supported Products</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow>
	<FRAME ID="UpperWindow" Name="UpperWindow" SRC="ProductsSupportedMain.asp?VersionID=<%=request("VersionID")%>">
	<FRAME  noresize ID="LowerWindow" Name="LowerWindow" SRC="ProductsSupportedButtons.asp" scrolling=no>
</FRAMESET>

</HTML>
