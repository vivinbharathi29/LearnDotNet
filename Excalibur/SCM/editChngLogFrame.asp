<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsPC

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	
	m_IsPC = Security.UserInRole(m_ProductVersionID, "PC")

	Set Security = Nothing
'##############################################################################	
%>
<HTML>
<HEAD>
<TITLE>Edit Change Log Details</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="editChngLogMain.asp?<%=Request.QueryString%>">
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="editChngLogButtons.asp?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>