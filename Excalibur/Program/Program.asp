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
	<TITLE>Update Product Group</TITLE>
<% else%>
	<TITLE>Add Product Group</TITLE>
<% end if%>
<HEAD>
</HEAD>

	<FRAMESET ROWS="*,65" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="EditProgram.asp?ID=<%=Request("ID")%>&pulsarplus=<%=Request("pulsarplus")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="Buttons.asp?pulsarplus=<%=Request("pulsarplus")%>">
	</FRAMESET>
</HTML>