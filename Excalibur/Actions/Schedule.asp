<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<% if Request("ID") <> "" then %>
	<TITLE>Update Roadmap Item</TITLE>
<% else%>
	<TITLE>Add Roadmap Item</TITLE>
<% end if%>
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
</HEAD>

	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ScheduleMain.asp?ID=<%=Request("ID")%>&ProductID=<%=Request("ProductID")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ScheduleButtons.asp">
	</FRAMESET>
</HTML>