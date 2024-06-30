<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<!-- PBI 19513/ Task 20989 - Add Copy with Targeting to context menu -->
<% if Request("ID") <> "" then %>
	<TITLE>Edit Image Definition </TITLE>
<% elseif request("CopyID") and Request("CopyTarget") = "0" then%>
	<TITLE>Copy Image Definition </TITLE>
<% elseif request("CopyID") and Request("CopyTarget") = "1" then%>
	<TITLE>Copy Image Definition with Targeting </TITLE>
<% else%>
	<TITLE>Add New Image Definition </TITLE>
<% end if%>
<HEAD>

</HEAD>
	<FRAMESET ROWS="*,50" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="imageMain.asp?CopyID=<%=Request("CopyID")%>&ID=<%=Request("ID")%>&ProdID=<%=Request("ProdID")%>&CopyTarget=<%=Request("CopyTarget")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="Imagebuttons.asp?CopyID=<%=Request("CopyID")%>&ID=<%=Request("ID")%>&ProdID=<%=Request("ProdID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>" scrolling=no>
	</FRAMESET>
</HTML>