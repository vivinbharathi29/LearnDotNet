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
	<TITLE>Edit Deliverable</TITLE>
<% else%>
	<TITLE>Add New Deliverable</TITLE>
<% end if%>
<HEAD>

</HEAD>
<FRAMESET ROWS="*" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="HFCNAddMain.asp?ID=<%=Request("ID")%>&CatID=<%=Request("CatID")%>">
</FRAMESET>

</HTML>