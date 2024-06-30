<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<TITLE>Advanced Deliverable Targeting</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,65" ID=TopWindow >
	<%if Not(isnull(Request("ExcludeFunComp"))) and Request("ExcludeFunComp")<> "" then%>
	    <FRAME noresize ID="MainWindow" Name="MainWindow" SRC="TargetAdvancedMain.asp?ProductID=<%=Request("ProductID")%>&VersionID=<%=Request("VersionID")%>&RootID=<%=Request("RootID")%>&ExcludeFunComp=<%=Request("ExcludeFunComp")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<%else%>
        <FRAME noresize ID="MainWindow" Name="MainWindow" SRC="TargetAdvancedMain.asp?ProductID=<%=Request("ProductID")%>&VersionID=<%=Request("VersionID")%>&RootID=<%=Request("RootID")%>&ExcludeFunComp=false&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<%end if%>
    <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="TargetAdvancedButtons.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
</FRAMESET>

</HTML>