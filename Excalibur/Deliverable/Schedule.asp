<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<HTML>
<TITLE>Update Schedule</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ScheduleMain.asp?ID=<%=Request("ID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ScheduleButtons.asp">
</FRAMESET>

</HTML>