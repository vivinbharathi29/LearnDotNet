<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<% if Request("ID") <> "" then %>
	<TITLE>Edit Image Definition (Pulsar)</TITLE>
<% elseif request("CopyID") and Request("CopyTarget") = "0" then%>
	<TITLE>Copy Image Definition (Pulsar)</TITLE>
<% elseif request("CopyID") and Request("CopyTarget") = "1" then%>
	<TITLE>Copy Image Definition with Targeting (Pulsar)</TITLE>
<% else%>
	<TITLE>Add New Image Definition (Pulsar)</TITLE>
<% end if%>
<HEAD>

</HEAD>
	<FRAMESET ROWS="*,50" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ImageMain_Pulsar.asp?CopyID=<%=Request("CopyID")%>&ID=<%=Request("ID")%>&ProdID=<%=Request("ProdID")%>&CopyTarget=<%=Request("CopyTarget")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="Imagebuttons.asp?CopyID=<%=Request("CopyID")%>&ID=<%=Request("ID")%>&ProdID=<%=Request("ProdID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>" scrolling=no>
	</FRAMESET>
</HTML>