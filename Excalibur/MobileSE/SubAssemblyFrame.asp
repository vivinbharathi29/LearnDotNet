<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<% if Request("Existing") <> "" then %>
	<TITLE>Edit Exisiting Subassembly Definition</TITLE>
<% else%>
	<TITLE>Add New Subassembly Definition</TITLE>
<% end if%>
<HEAD>

</HEAD>
	<FRAMESET ROWS="*,50" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="SubAssemblyMain.asp?SAType=<%=Request("SAType")%>&Step=<%=Request("Step")%>&FeatureCategoryID=<%=Request("FeatureCategoryID")%>&SubassemblyID=<%=Request("SubassemblyID")%>&BusinessID=<%=Request("BusinessID")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
		<%If Request("Existing") <> "" then %>
		    <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="SubAssemblyButtons.asp?SAType=<%=Request("SAType")%>&Existing=1&pulsarplusDivId=<%=Request("pulsarplusDivId")%>"scrolling=no>
		<% else%>
		    <FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="SubAssemblyButtons.asp?SAType=<%=Request("SAType")%>&pulsarplusDivId=<%=Request("pulsarplusDivId")%>"scrolling=no>
		<% end if%>
	</FRAMESET>
</HTML>