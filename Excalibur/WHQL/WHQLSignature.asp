<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
	<TITLE>WHQL Signature Request</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="120,30" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="WHQLSignatureMain.asp">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="WHQLSignatureButton.asp">
	</FRAMESET>
</HTML>