<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
<HEAD>
	<TITLE>Choose Versions</TITLE>
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ChooseVersionsMain.asp">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ChooseVersionsButtons.asp">
	</FRAMESET>
</HTML>




