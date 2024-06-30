<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<TITLE>Generate Files</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="GenerateMain.asp?VersionID=<%=Request("VersionID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="GenerateButtons.asp">
</FRAMESET>

</HTML>