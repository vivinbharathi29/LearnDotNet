<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Edit Master Sku Components</TITLE>
<HEAD>

</HEAD>
<FRAMESET ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ImageDriveDefinitionMain.aspx?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>