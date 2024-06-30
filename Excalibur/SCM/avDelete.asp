<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
Dim rs, dw, cn, cmd

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName


'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "usp_DeleteAvDetail_ProductBrand")
cmd.NamedParameters = True
dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request("AVID")
dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
dw.CreateParameter cmd, "@p_UserName", adVarChar, adParamInput, 50, m_UserFullName
cmd.Execute

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY onload="window.close();">
<P>&nbsp;</P>
</BODY>
</HTML>
