<%@ Language=VBScript %>


	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<% if Request("ID") <> "" and Request("ID") <> "0" then %>
	<TITLE>Update
<% else%>
	<TITLE>Add
<% end if%>
<% if trim(Request("Type")) = "1" then %>
Issue
<%else%>
Task
<%end if%>
</TITLE>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,65" ID=TopWindow >
	    <FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ActionMain.asp?TicketID=<%=request("TicketID")%>&AppErrorID=<%=Request("AppErrorID")%>&ID=<%=Request("ID")%>&RoadmapID=<%=Request("RoadmapID")%>&ProdID=<%=Request("ProdID")%>&Working=<%=Request("Working")%>&Type=<%=request("Type")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ActionButtons.asp">
	</FRAMESET>
</HTML>