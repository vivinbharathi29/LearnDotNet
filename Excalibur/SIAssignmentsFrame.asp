	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
	<TITLE>SI Assignments</TITLE>
</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME  noresize ID="UpperWindow" Name="UpperWindow" SRC="MobileSE/Today/SIAssignments.asp?PVID=<%=request("PVID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="MobileSE/Today/SIAssignmentsButtons.asp">
</FRAMESET>
</HTML>