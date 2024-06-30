<%@ Language=VBScript %>
<%Option Explicit%>

<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 

<%
Response.Buffer = True
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim m_ScheduleID
Dim m_ProductVersionID
Dim m_SelectedProduct
Dim m_SelectedRelease
Dim m_IsSysAdmin
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_IsSEPMProductsEditor
Dim m_EditModeOn
Dim m_UserFullName
Dim m_CopyDates

Sub Main()

	m_ProductVersionID = Request("PVID")

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
'		Response.Write Security.IsProgramManager(m_ProductVersionID)
'		Response.Write "<BR>"
'		Response.Write Security.IsSysEngProgramManager(m_ProductVersionID)
'		Response.Write "<BR>"
'		Response.Write Security.IsSystemTeamLead(m_ProductVersionID)
'		Response.Write "<BR>"
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If

	m_IsProgramManager = Security.IsProgramManager(m_ProductVersionID)
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(m_ProductVersionID)
	m_IsSysTeamLead = Security.IsSystemTeamLead(m_ProductVersionID)
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	m_UserFullName = Security.CurrentUser()
	
	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager or m_IsSysTeamLead Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	

	m_ProductVersionID = Request("ID")
	If m_ProductVersionID = "" Then m_ProductVersionID = Request("PVID")
	m_ScheduleID = Request("ScheduleID")
	m_CopyDates = (Request("cbCopyDates") = "on")
	
	If Request.Form("selProduct") = "" Then
		m_SelectedProduct = Request.QueryString("PVID")
	Else
		m_SelectedProduct = Request.Form("selProduct")
	End If
	
	m_SelectedRelease = Request.Form("selRelease")	

	If m_ScheduleID = "" Then
		Response.Write "Insufficient Data to process request"
		Response.End
	End If 
	
	If Request.Form("FormSave") = "True" Then
		Call SaveData()
	End If
End Sub

Sub SaveData()
	Dim dw, cn, cmd, iRowsChanged
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans
	
	If m_CopyDates Then
	    m_CopyDates = 1
	Else
	    m_CopyDates = 0
    End If
	
	Set cmd = dw.CreateCommandSP(cn, "usp_CopyScheduleData")
	dw.CreateParameter cmd, "@p_SrcScheduleID", adInteger, adParamInput, 8, m_SelectedRelease
	dw.CreateParameter cmd, "@p_DstScheduleID", adInteger, adParamInput, 8, m_ScheduleID
	dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(m_UserFullName)
	dw.CreateParameter cmd, "@p_CopyDates", adBoolean, adParamInput, 1, m_CopyDates
	iRowsChanged = dw.ExecuteNonQuery(cmd)

	If iRowsChanged < 1 Then
		cn.RollbackTrans	
		Response.Write "Error Saving Schedule Item"
		Response.End
	Else
		cn.CommitTrans
		Response.Write "<input type=hidden id=CloseOnLoad value=true>"
	End If

	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
		
End Sub

Sub FillProductList()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spGetProducts")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		if trim(m_SelectedProduct) = trim(rs("ID")) then
			Response.Write "<option selected value=""" & rs("ID") & """>" & server.HTMLEncode(rs("Name")) & " " & server.HTMLEncode(rs("Version")) & "</option>"					
		else
			Response.Write "<option value=""" & rs("ID") & """>" & server.HTMLEncode(rs("Name")) & " " & server.HTMLEncode(rs("Version")) & "</option>"					
		end if
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

Sub FillReleaseList()
	Dim dw, cn, cmd, rs
	
	If Trim(m_SelectedProduct) <> "" Then
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_ListSchedules")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, m_SelectedProduct
		dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, "Y"
		Set rs = dw.ExecuteCommandReturnRS(cmd)
	
		Do until rs.eof
			if trim(m_SelectedRelease) = trim(rs("schedule_id")) then
				Response.Write "<option selected value=""" & rs("schedule_id") & """>" & server.HTMLEncode(rs("schedule_name")) & "</option>"					
			else
				Response.Write "<option value=""" & rs("schedule_id") & """>" & server.HTMLEncode(rs("schedule_name")) & "</option>"					
			end if
			rs.movenext
		Loop

		rs.close
	
		set dw = nothing
		set cn = nothing
		set cmd = nothing
		set rs = nothing
	End If
End Sub

%>
	
	
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Add Custom Item</title>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    }
    else {
        if (window.location != window.parent.location) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.close();
        }
    }
}

function VerifySave()
{
	with (window.frmMain)
	{
	if (!validateTextInput(selProduct, 'Product')){	return false; }
	if (!validateTextInput(selRelease, 'Release')){	return false; }
	}
	return true;
}

function cmdOK_onclick() 
{
	if (VerifySave())
	{
		
		window.frmMain.cmdCancel.disabled = true;
		window.frmMain.cmdOK.disabled = true;
		window.frmMain.FormSave.value = "True";
		window.frmMain.submit();
	}
}


function window_onLoad(pulsarplusDivId) {
    if (window.CloseOnLoad) {
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Reload PulsarPlusPmView Tab
            parent.window.parent.reloadFromPopUp(pulsarplusDivId);
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (window.location != window.parent.location) {
                parent.window.parent.modalDialog.cancel(true);
            } else {
                window.returnValue = 1;
                window.close();
            }
        }
    }
}

function selProduct_onChange()
{
	window.frmMain.submit();
}
//-->
</script>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
</head>
<body bgcolor="ivory" leftMargin="9" topMargin="9" OnLoad="window_onLoad('<%= Request("pulsarplusDivId")%>')">
<% Call Main() %>
	<h3>Copy Schedule Items</h3>
	<form ID="frmMain" method="post">
	<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr>
			<td nowrap valign="top"><b>Product:</b>&nbsp;<font color="red" size="1">*</font></td>
			<td>
			<select id="selProduct" name="selProduct" onchange="selProduct_onChange()">
				<option value>--- Select Product ---</option>
				<% Call FillProductList() %>
			</select>
			</td></tr>
		<tr>
			<td width="150" nowrap><b>Release:</b>&nbsp;<font color="red" size="1">*</font></td>
			<td width="100%">
			<select id="selRelease" name="selRelease">
				<option value>--- Select Release ---</option>
				<% Call FillReleaseList() %>
			</select>
			</td></tr>
		<tr>
			<td width="150" nowrap><b>Copy Dates:</b></td>
			<td width="100%">
            <input id="cbCopyDates" name="cbCopyDates" type="checkbox" />
			</td></tr>
	</table>
    <br />
<table width="100%" border="0">
  <tr><td align="right">
<input type="button" value="OK" id="cmdOK" name="cmdOK" LANGUAGE="javascript" onclick="return cmdOK_onclick()">
<input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%= Request("pulsarplusDivId")%>')">
  </td></tr>
</table>
<input type="hidden" id="PostBack" name="PostBack" value="True">
<input type="hidden" id="FormSave" name="FormSave" value="False">
<input type="hidden" id="ScheduleID" name="ScheduleID" value="<%= m_ScheduleID%>">
<input type="hidden" id="PVID" name="PVID" value="<%= m_ProductVersionID%>">
</form>
</body>
</html>
