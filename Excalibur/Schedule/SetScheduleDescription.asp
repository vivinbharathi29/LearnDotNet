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

Dim m_IsSysAdmin
Dim m_IsProgramManager
Dim m_IsSysEngProgramManager
Dim m_IsSysTeamLead
Dim m_IsSEPMProductsEditor
Dim m_UserFullName
Dim m_EditModeOn
Dim m_ScheduleID
Dim m_ScheduleDescription
Dim m_IsPulsarProduct
Dim m_ProductVersionID

Sub Main()
'##############################################################################	
'
' Create Security Object to get User Info
'

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
'		Response.Write Request.QueryString
'		Response.Write "<BR>"
'		Response.Write Request.Form
'		Response.Write "<BR>"
'		Response.End
'	End If
	
    m_ProductVersionID = Request.QueryString("PVID")
	m_IsProgramManager = Security.IsProgramManager(m_ProductVersionID)
	m_IsSysEngProgramManager = Security.IsSysEngProgramManager(m_ProductVersionID)
	m_IsSysTeamLead = Security.IsSystemTeamLead(m_ProductVersionID)
    m_IsSEPMProductsEditor = Security.IsSEPMProductsPermissions()
	m_UserFullName = Security.CurrentUser()
	m_IsPulsarProduct = Request.QueryString("IsPulsarProduct")

	If m_IsSysAdmin Or m_IsProgramManager Or m_IsSysEngProgramManager Or m_IsSysTeamLead Or m_IsSEPMProductsEditor Then
		m_EditModeOn = True
	End If
	
	Set Security = Nothing

	If Not m_EditModeOn Then
		Response.Write "<H3>Insuficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If
'##############################################################################	

	m_ScheduleID = Request.QueryString("ScheduleID")
	
	If m_ScheduleID = "" Then
		Response.Write "Insufficient Data to process request"
		Response.End
	End If 
	
	If Request.Form("FormSave") = "True" Then
		If Trim(Request.Form("txtScheduleDescription")) <> Trim(Request.Form("hidScheduleDescription")) Then
			Call SaveData()
		End If
	End If
	
	Call GetScheduleDescription()
End Sub

Sub SaveData()
	Dim dw, cn, cmd, iRowsChanged, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	cn.BeginTrans
	
    Set cmd = dw.CreateCommandSP(cn, "usp_ScheduleNameLookUp")
	dw.CreateParameter cmd, "@p_Name", adVarChar, adParamInput, 500, Left(Trim(Request.Form("txtScheduleDescription")), 500)
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	If not rs.EOF and not rs.BOF and m_IsPulsarProduct = 1 Then
		cn.RollbackTrans
		response.Write "<span style='color:red; font-size:80%'>Custom schedule name can not be the same as release name. Please use a name that does not match a Product Release name.</span>"
	else	

	    Set cmd = dw.CreateCommandSP(cn, "usp_UpdateSchedule")
	    dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, m_ScheduleID
	    dw.CreateParameter cmd, "@p_Description", adVarChar, adParamInput, 500, Left(Trim(Request.Form("txtScheduleDescription")), 500)
	    dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 500, Left(Trim(m_UserFullName), 500)
	    iRowsChanged = dw.ExecuteNonQuery(cmd)

	    If iRowsChanged < 1 Then
		    cn.RollbackTrans	
		    Response.Write "Error Saving Schedule Item"
		    Response.End
	    Else
		    cn.CommitTrans
		    Response.Write "<input type=hidden id=CloseOnLoad value=true>"
	    End If

    End IF

	Set cmd = Nothing
	Set cn = Nothing
	Set dw = Nothing
		
End Sub

Sub GetScheduleDescription()
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectSchedule")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, NULL
	dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, NULL
	dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, m_ScheduleID
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	m_ScheduleDescription = Trim(rs("description") & "")

	rs.close
	
	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub


%>
	
	
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Edit Schedule Description</title>
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
	if (!validateTextInput(txtScheduleDescription, 'Schedule Description')){	return false; }
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
	if (window.CloseOnLoad)
	{
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
	else {
	    document.getElementById("txtScheduleDescription").focus();
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
	<h3>Edit Schedule Description</h3>
	<form ID="frmMain" method="post">
	<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
		<tr>
			<td nowrap valign="top"><b>Description:</b>&nbsp;<font color="red" size="1">*</font></td>
			<td width="100%"><INPUT type="text" id=txtScheduleDescription name=txtScheduleDescription size=50 value="<%= m_ScheduleDescription%>"><INPUT type="hidden" id=hidScheduleDescription name=hidScheduleDescription value="<%= m_ScheduleDescription%>">
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
</form>
</body>
</html>
