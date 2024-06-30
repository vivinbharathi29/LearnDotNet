<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>



<HTML>
<HEAD>
	<TITLE>Custom Reports</TITLE>
</HEAD>

	<FRAMESET ROWS="*" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="CustomReportMain1.asp?ProfileID=<%=request("ProfileID")%>&RunReportOK=<%=request("RunReportOK")%>">
	</FRAMESET>
</HTML>