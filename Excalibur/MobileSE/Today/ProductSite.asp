<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Choose RCTO Sites</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ProductSiteMain.asp?ID=<%=request("ID")%>&Sites=<%=request("Sites")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ProductSiteButtons.asp">
</FRAMESET>

</HTML>
