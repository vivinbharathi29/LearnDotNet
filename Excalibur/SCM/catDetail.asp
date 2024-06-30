<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sMode				    : sMode = Request.QueryString("Mode")
Dim sFunction			    : sFunction = Request.Form("hidFunction")
Dim sFeatureCat			    : sFeatureCat = ""
Dim sManufacturingNotes	    : sManufacturingNotes = ""
Dim sMarketingDescription   : sMarketingDescription = ""
Dim sConfigRules		    : sConfigRules = ""
Dim iBrandID			    : iBrandID = Request.QueryString("BID")

Dim m_ProductVersionID	    : m_ProductVersionID = Request("PVID")
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
'
'PC Can do any thing
'Marketing can change description

Function PrepForWeb( value )
	
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		PrepForWeb = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		PrepForWeb = "X"
	Else
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Sub Main()
'
'TODO: Get CatDetail Data
'
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvFeatureDetail")
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Trim(Request("FCID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		If Not rs.EOF Then
			sFeatureCat = rs("AvFeatureCategory") & ""
			sConfigRules = rs("ConfigRules") & ""
			sManufacturingNotes = rs("ManufacturingNotes") & ""
			sMarketingDescription = rs("MarketingDescription") & ""
		End If
		
		rs.Close
		
End Sub

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	Dim iAvId

If Request.Form("txtConfigRules") <> Request.Form("hidConfigRulesDefault") _
    Or Request.Form("txtManufacturingNotes") <> Request.Form("hidManufacturingNotesDefault") _
    Or Request.Form("txtMarketingDescription") <> Request.Form("hidMarketingDescriptionDefault") Then
	
	cn.BeginTrans
'Save AvDetail data
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvFeatureDetail")
	cmd.NamedParameters = True
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Request("FCID")
	dw.CreateParameter cmd, "@p_ConfigRules", adVarchar, adParamInput, 2000, Request.Form("txtConfigRules")
	dw.CreateParameter cmd, "@p_ManufacturingNotes", adVarchar, adParamInput, 2000, Request.Form("txtManufacturingNotes")
	dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 2000, Request.Form("txtMarketingDescription")
	dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 2000, ""
	dw.CreateParameter cmd, "@p_LastUpdUser", adVarchar, adParamInput, 50, m_UserFullName
	returnValue = dw.ExecuteNonQuery(cmd)

	'If returnValue <> 1 Then
	'	' Abort Transaction
	'	Response.Write returnValue
	'	cn.RollbackTrans()
	'	Exit Sub
	'End If
	
'	If Request.Form("txtConfigRules") <> Request.Form("hidConfigRulesDefault") Then
'		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvHistory")
'		cmd.NamedParameters = True
'		dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, ""
'		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
'		dw.CreateParameter cmd, "@p_Comments", adVarchar, adParamInput, 500, "Changing Configuration Rules for the " & Request.Form("hidCatName") & " category."
'		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, 1
'		dw.CreateParameter cmd, "@p_LastUpdUser", adVarchar, adParamInput, 50, m_UserFullName
'		dw.CreateParameter cmd, "@p_OldValue", adVarchar, adParamInput, 800, Request.Form("hidConfigRulesDefault")
'		dw.CreateParameter cmd, "@p_NewValue", adVarchar, adParamInput, 800, Request.Form("txtConfigRules")
'		returnValue = dw.ExecuteNonQuery(cmd)

'		If returnValue <> 1 Then
'			' Abort Transaction
'			Response.Write returnValue
'			cn.RollbackTrans()
'			Exit Sub
'		End If
'	End If

'	If Request.Form("txtMarketingDescription") <> Request.Form("hidMarketingDescriptionDefault") Then
'		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvHistory")
'		cmd.NamedParameters = True
'		dw.CreateParameter cmd, "@p_AvHistoryID", adInteger, adParamInput, 8, ""
'		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
'		dw.CreateParameter cmd, "@p_Comments", adVarchar, adParamInput, 500, "Changing Marketing Description for the " & Request.Form("hidCatName") & " category."
'		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, 1
'		dw.CreateParameter cmd, "@p_LastUpdUser", adVarchar, adParamInput, 50, m_UserFullName
'		dw.CreateParameter cmd, "@p_OldValue", adVarchar, adParamInput, 800, Request.Form("hidMarketingDescriptionDefault")
'		dw.CreateParameter cmd, "@p_NewValue", adVarchar, adParamInput, 800, Request.Form("txtMarketingDescription")
'		returnValue = dw.ExecuteNonQuery(cmd)
'
'		If returnValue <> 1 Then
'			' Abort Transaction
'			Response.Write returnValue
'			cn.RollbackTrans()
'			Exit Sub
'		End If
'	End If

	sFunction = "close"
	cn.CommitTrans
End If
End Sub

If LCase(sFunction) = "save" Then
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
			window.close();
			break;
		case "save":
			window.close();
			break;
	}		
	
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
	}
	
}
</SCRIPT>
</HEAD>
<BODY OnLoad="Body_OnLoad()">
<%'= response.Write(request.QueryString) %>
<FORM method=post id=frmMain>
<INPUT id="hidMode" name="hidMode" type="hidden" value=<%= LCase(sMode)%>>
<INPUT id="hidFunction" name="hidFunction" type="hidden" value=<%= LCase(sFunction)%>>
<INPUT id="BID" name="BID" type="hidden" value="<%= iBrandID%>">
<INPUT id="hidCatName" name="hidCatName" type="hidden" value="<%= sFeatureCat%>">
<TEXTAREA rows=5 id=hidConfigRulesDefault name=hidConfigRulesDefault style="display:none;"><%= sConfigRules %></TEXTAREA>
<TEXTAREA rows=5 id=hidManufacturingNotesDefault name=hidManufacturingNotesDefault style="display:none;"><%= sManufacturingNotes %></TEXTAREA>
<TEXTAREA rows=5 id=hidMarketingDescriptionDefault name=hidMarketingDescriptionDefault style="display:none;"><%= sMarketingDescription %></TEXTAREA>
<TABLE class="FormTable" bgcolor=cornsilk WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=1 bordercolor=tan>
	<TR>
		<TH>Feature Category</TH>
		<TD><%= PrepForWeb(sFeatureCat)%></TD>
	</TR>
	<TR>
		<TH>Configuration Rules</TH>
		<TD><TEXTAREA rows=5 id=txtConfigRules name=txtConfigRules style="width:300px"><%= sConfigRules%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Manufacturing Notes</TH>
		<TD><TEXTAREA rows=5 id=txtManufacturingNotes name=txtManufacturingNotes style="width:300px"><%= sManufacturingNotes%></TEXTAREA></TD>
	</TR>
	<tr>
	    <th>Marketing Description</th>
	    <td><textarea rows=5 id=txtMarketingDescription name=txtMarketingDescription style="width:300px"><%= sMarketingDescription %></textarea></td>
	</tr>
</TABLE>
</FORM>
</BODY>
</HTML>
