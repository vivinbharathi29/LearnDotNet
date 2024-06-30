<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
'Response.Write Request.QueryString
'Response.End

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim m_ChangeLogID		: m_ChangeLogID = Request("CLID")
Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_IsMarketingUser
Dim m_EditModeOn
Dim m_UserFullName
Dim m_Function			: m_Function = Request.Form("hidFunction")
Dim m_Mode				: m_Mode = Request.Form("hidMode") : If m_Mode = "" Then m_Mode = Request.QueryString("Mode")
Dim m_UserID
Dim sAvNo
Dim sChangeDt
Dim sChangeBy
Dim sField
Dim sChangeType
Dim sChangeFrom
Dim sChangeTo
Dim sChangeReason
Dim bShowOnScm
Dim bShowOnPM

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()

	'm_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	'm_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	'm_UserFullName = Security.CurrentUserFullName()
	
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If
	
    'response.Write m_IsProgramCoordinator
	'If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Or (m_IsMarketingUser And m_Mode = "add") Then
	'	m_EditModeOn = True
	'End If
	
	'If Not m_EditModeOn Then
	'	Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
	'	Response.End
	'End If  
  

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
    m_UserID = Security.CurrentUserId()
	
    
    'Response.Write m_UserFullName & "<br />"
	'Response.Write m_IsConfigurationManager  & "<br />"
	'Response.Write m_IsProgramCoordinator  & "<br />"
	'Response.Write m_IsMarketingUser  & "<br />"

	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Or (m_IsMarketingUser And m_Mode = "add") Then
		m_EditModeOn = True
	End If

	Set Security = Nothing
'##############################################################################	

Function GetProductVersion( ProductVersionID )
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	GetProductVersion = rs("version") & ""
	
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Function

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( replace(value, Chr(10), "<br />" ) )
	End If

End Function

Sub Main()

	bShowOnScm = True
	bShowOnPM = False
	
	If m_Mode <> "add" Then
		Set cmd = dw.CreateCommAndSP(cn, "usp_SelectAvHistory")
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ""
		dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, Request("CLID")
		Set rs = dw.ExecuteCommandReturnRS(cmd)

		sAvNo = rs("AvNo")
		sChangeDt = rs("Last_Upd_Date")
		sChangeBy = rs("Last_Upd_User")
		sField = rs("ColumnChanged")
		sChangeType = rs("AvChangeTypeDesc")
		sChangeFrom = rs("OldValue")
		sChangeTo = rs("NewValue")
		sChangeReason = rs("Comments")
		If rs("ShowOnScm") = 0 Then
			bShowOnScm = False
		End If
		If rs("ShowOnPM") = 1 Then
		    bShowOnPM = True
		End If

		rs.Close
	End If

End Sub

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	
	cn.BeginTrans
	
'Save AvDetail data
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvHistory")
	dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, Request.QueryString("CLID")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request.QueryString("PBID")
	dw.CreateParameter cmd, "@p_Comments", adVarchar, adParamInput, 500, Request.Form("txtComments")
	dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, Request.Form("chkShowOnScm")
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
	dw.CreateParameter cmd, "@p_OldValue", adVarchar, adParamInput, 800, Request.Form("txtChangeFrom")
	dw.CreateParameter cmd, "@p_NewValue", adVarchar, adParamInput, 800, Request.Form("txtChangeTo")
	dw.CreateParameter cmd, "@p_ShowOnPM", adBoolean, adParamInput, 1, Request.Form("chkShowOnPM")
	returnValue = dw.ExecuteNonQuery(cmd)

	If returnValue <> 1 Then
		' Abort Transaction
		Response.Write returnValue
		cn.RollbackTrans()
		Exit Sub
	End If

    m_ChangeLogID = returnValue

	m_Function = "save"
	cn.CommitTrans
	
End Sub

If LCase(m_Function) = "save" Then
	Call Save()
Else
	Call Main()
End If
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/excalibur.css">
<SCRIPT type="text/javascript">
function Body_OnLoad()
{
	switch (frmMain.hidFunction.value)
	{
	    case "close":
	        var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	        }
	        else {
	            window.close();
	        }
			break;
	    case "save":
	        var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	        }
	        else {
	            window.returnValue = window.frmMain.hidCLID.value;
	            window.close();
	        }
            break;
	}		
	
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
	    if ((window.frmMain.hidMode.value.toLowerCase() == 'edit' || window.frmMain.hidMode.value.toLowerCase() == 'add') && window.frmMain.HasAccess.value.toLowerCase() == 'true') {
	        window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = false;
	    }
	    else {
	        window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = true;
	    }
	}
	
}

</SCRIPT>
</HEAD>
<BODY OnLoad="Body_OnLoad()">
<FORM method=post id=frmMain>
<INPUT id="hidMode" name="hidMode" type=HIDDEN value="<%= LCase(m_Mode)%>">
<INPUT id="hidCLID" name="hidCLID" type=HIDDEN value="<%= m_ChangeLogID%>">
<input id="HasAccess" name="HasAccess" type="hidden" value="<%= m_EditModeOn %>" />
<INPUT id="hidFunction" name="hidFunction" type=HIDDEN value="<%= m_Function%>">
<input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
<TABLE class="FormTable" bgcolor=cornsilk WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=1 bordercolor=tan>
<% If m_Mode <> "add" Then %>
	<TR>
		<TH>Change Date</TH>
		<TD><%= PrepForWeb(sChangeDt)%></TD>
	</TR>
	<TR>
		<TH>Change By</TH>
		<TD><%= PrepForWeb(sChangeBy)%></TD>
	</TR>
	<TR>
		<TH>Av No.</TH>
		<TD><%= PrepForWeb(sAvNo)%></TD>
	</TR>
	<TR>
		<TH>Field</TH>
		<TD><%= PrepForWeb(sField)%></TD>
	</TR>
	<TR>
		<TH>Change Type</TH>
		<TD><%= PrepForWeb(sChangeType)%></TD>
	</TR>
	<TR>
		<TH>Change From</TH>
		<TD><%= PrepForWeb(sChangeFrom)%></TD>
	</TR>
	<TR>
		<TH>Change To</TH>
		<TD><%= PrepForWeb(sChangeTo)%></TD>
	</TR>
<% Else %>
	<TR>
		<TH>Change From</TH>
		<TD>
            <input id="txtChangeFrom" name="txtChangeFrom" type="text" maxlength="800"  /></TD>
	</TR>
	<TR>
		<TH>Change To</TH>
		<TD>
            <input id="txtChangeTo" name="txtChangeTo" type="text" maxlength="800" /></TD>
	</TR>
<% End If %>	
	<TR>
		<TH>Comment</TH>
		<TD><TEXTAREA rows=2 id=txtComments name=txtComments style="width:300px"><%= sChangeReason%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Show On SCM Change Log</TH>
		<TD><INPUT type="checkbox" id="chkShowOnScm" name="chkShowOnScm" <% If bShowOnScm Then Response.Write "CHECKED" End If %>>
	<TR>
		<TH>Show On Program Matrix Change Log</TH>
		<TD><INPUT type="checkbox" id="chkShowOnPM" name="chkShowOnPM" <% If bShowOnPM Then Response.Write "CHECKED" End If %>>
</TABLE>
</FORM>
</BODY>
</HTML>
