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

'Dim AvPcActionTypeID
'If Request("Status") = "O" Then
'   AvPcActionTypeID = 2 'Delete or Obsolete an AV
'Elseif Request("Status") = "A" Then
'   AvPcActionTypeID = 1 'Activate an Obsolete AV
'End If

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

'
' Debug Section
'
'	If Security.CurrentUserID = 1396 Then
'		m_IsSysAdmin = False
'		Security.CurrentUserID = 1288
'		Response.Write Security.CurrentUserID
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSysEngProgramManager(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsSystemTeamLead(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write Security.IsProgramCoordinator(Request("PVID"))
'		Response.Write "<BR>"
'		Response.Write m_ProductVersionID
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If


	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If
	
    
	Set Security = Nothing

	If m_EditModeOn Then
        On Error Resume Next
        Err.Clear
        Set rs = Server.CreateObject("ADODB.RecordSet")
        Set cn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        Set dw = New DataWrapper
        Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

        Set cmd = dw.CreateCommandSP(cn, "usp_SetAvStatus")
        cmd.NamedParameters = True
        dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Request("AVID")
        dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
        dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, Request("Status")
        dw.CreateParameter cmd, "@p_UserName", adVarChar, adParamInput, 50, m_UserFullName
        'dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
        cmd.Execute

        If Request("pulsarplusDivId") = "" Then
            If Err.Number <> 0 or cn.Errors.count > 0 Then         
                Response.Write("<script language=javascript>alert('" + Err.Description + "'); window.returnValue = 0; window.close(); </script>")
            else
                Response.Write("<script language=javascript>window.returnValue = " + Request("AVID") + "; window.close(); </script>")                    
            End If
        Else        
            If Err.Number <> 0 or cn.Errors.count > 0 Then         
                Response.Write("0")
            else
                Response.Write(Request("AVID"))                    
            End If
            Response.End()    
         End if
    End If



%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY >
    <form method="post" action="" id="frmAvSetStatus">
        <table style="width:98%; height: 59px;" >
            <tr align="center">
                <td >
                    <p  id="lblPrivilege" name="lblPrivilege"><H3>Insufficient User Privileges</H3><H4>Access Denied</H4></p>
                </td>
            </tr>
            <tr`>
                <td align="center" >
                  <input type="button" id="btnCancel" name="btnCancel"  value="Cancel" onclick="javascript: window.returnValue = 0; window.close();" />
                </td>
            </tr>
        </table>
    </form>
</BODY>
</HTML>
