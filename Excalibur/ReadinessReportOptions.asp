<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Readiness Report Report</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ReadinessReportOptionsMain.asp?ProdID=<%=request("ProdID")%>&ReportType=<%=request("ReportType")%>&TeamID=<%=request("TeamID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ReadinessReportOptionsButtons.asp">
</FRAMESET>

</HTML>