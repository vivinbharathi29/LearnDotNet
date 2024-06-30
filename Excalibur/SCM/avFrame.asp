<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
Dim m_ProductVersionID : m_ProductVersionID = Request("PVID")
Dim m_UserID : m_ProductVersionID = Request("UserID")
Dim m_IsMarketingUser

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	Dim Security
	
	Set Security = New ExcaliburSecurity
	
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If

    Dim isAdmin 
    isAdmin = Security.IsSysAdmin()
	Set Security = Nothing
'##############################################################################	
%>
<HTML>
<HEAD>
<TITLE>SCM AV Details</TITLE>
</HEAD>
	<FRAMESET ROWS="*,60" ID=TopWindow >
<% If m_IsMarketingUser Then %>
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="avMarketingDetail.asp?<%=Request.QueryString%>">
<% Else %>
		<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="avDetail.asp?<%=Request.QueryString%>">
<% End If %>
		<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="avButtons.asp?<%=Request.QueryString%>">
	</FRAMESET>
</HTML>