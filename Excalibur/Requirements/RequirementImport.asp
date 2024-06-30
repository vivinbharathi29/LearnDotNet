<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<TITLE>Import Requirements</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="RequirementImportMain.asp?ProductID=<%=Request("ProductID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="RequirementButtons.asp?ProductID=<%=Request("ProductID")%>&pulsarplusDivId=<%=request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>