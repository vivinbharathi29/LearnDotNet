<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Change Report</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,65" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProductStatusOptionsMain.asp?ID=<%=request("ID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProductStatusOptionsButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>