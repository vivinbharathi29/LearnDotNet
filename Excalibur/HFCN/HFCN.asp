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
	<TITLE>HFCN Properties</TITLE>
<% else%>
	<TITLE>Add New HFCN</TITLE>
<% end if%>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,55" ID=TopWindow >
	<FRAME noresize ID="MainWindow" Name="MainWindow" SRC="HFCNMain.asp?ID=<%=Request("ID")%>&RootID=<%=Request("RootID")%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="HFCNButtons.asp">
</FRAMESET>

</HTML>