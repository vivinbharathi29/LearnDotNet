<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%

Dim regEx
Set regEx = New RegExp
regEx.Global = True

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim i


regEx.Pattern = "[^0-9]"
Dim sBID            : sBID = RegEx.Replace(Request.QueryString("BID"), "")

regEx.Pattern = "[^0-9a-zA-Z-]"
Dim sMode			: sMode = regEx.Replace(Request.Form("hidMode"), "") : If sMode = "" Then sMode = regEx.Replace(Request.QueryString("Mode"), "")
Dim sKmat
Dim sEzcKmat
Dim sPlantCd
Dim sProjectCd
Dim sSalesOrg
Dim sConfigCd
Dim bShowOnPm
Dim bShowPhWebActionItems
Dim sAVSeriesName

Dim m_ProductVersionID	: m_ProductVersionID = regEx.Replace(Request("PVID"),"")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName

If m_ProductVersionID = "" Then
    response.Write "<h3>Missing Product Version Information Unable to Continue</h3>"
    response.End
End If


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
'	    response.Write m_UserFullName & "<br>"
'	    response.Write m_EditModeOn & "<br>"
'	    response.Write m_ProductVersionID & "<br>"
'	    response.Write sMode & "<br>"
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If

	Set Security = Nothing
'##############################################################################	
'
'PC Can do any thing
'Marketing can change description


Function GetProductVersion( ProductVersionID )
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
	dw.CreateParameter cmd, "@ID", adInteger, adParaminput, 8, Trim(Request("PVID"))
	
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
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Sub Main()
'
' Get KMAT Data
'
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectKmat")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParaminput, 8, ""
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParaminput, 8, sBID
	Set rs = dw.ExecuteCommandReturnRS(cmd)

    regEx.Pattern = "[^0-9a-zA-Z,=/-]"
	
	sKmat = regEx.Replace(rs("KMAT")&"","")
	sEzcKmat = regEx.Replace(rs("EZCKMAT")&"","")
	sPlantCd = regEx.Replace(rs("PlantCd")&"","")
	sProjectCd = regEx.Replace(rs("ProjectCd")&"","")
	sSalesOrg = regEx.Replace(rs("SalesOrg")&"","")
	sConfigCd = regEx.Replace(rs("ConfigCd")&"","")
	bShowOnPm = rs("ShowOnPm")
	bShowPhWebActionItems = rs("ShowPhWebActionItems")
	sAVSeriesName = regEx.Replace(rs("AVSeriesName")&"","")
	rs.Close
	
End Sub

Sub Save()
	Dim returnValue
'
'TODO: Save KMAT Data
'
    regEx.Pattern = "[^0-9a-zA-Z,=/-]"

    sKmat = regEx.Replace(Request.Form("txtKMAT"),"")
	sEzcKmat = regEx.Replace(Request.Form("txtEZCKMAT"),"")
	sPlantCd = regEx.Replace(Request.Form("txtPlantCd"),"")
	sProjectCd = regEx.Replace(Request.Form("txtProjectCd"),"")
	sSalesOrg = regEx.Replace(Request.Form("txtSalesOrg"),"")
	sConfigCd = regEx.Replace(Request.Form("txtConfigCd"),"")
	bShowOnPm = (Request.Form("cbShowOnPm") = "on")
	bShowPhWebActionItems = (Request.Form("cbShowPhWebActionItems") = "on")
	sAVSeriesName = regEx.Replace(Request.Form("txtAVSeriesName"),"")

	Dim sShowOnPm
	If bShowOnPm Then
	    sShowOnPm = 1
	Else
	    sShowOnPm = 0
	End If
	Dim sShowPhWebActionItems
	If bShowPhWebActionItems Then
	    sShowPhWebActionItems = 1
	Else
	    sShowPhWebActionItems = 0
	End If

	cn.BeginTrans()

	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateKmat")
	dw.CreateParameter cmd, "@p_ProductBrandID", adVarchar, adParaminput, 8, Request("BID")
	dw.CreateParameter cmd, "@p_Kmat", adVarchar, adParaminput, 10, sKmat
	dw.CreateParameter cmd, "@p_EzcKmat", adVarchar, adParaminput, 10, sEzcKmat
	dw.CreateParameter cmd, "@p_ProjectCd", adVarchar, adParaminput, 15, sProjectCd
	dw.CreateParameter cmd, "@p_PlantCd", adVarchar, adParaminput, 15, sPlantCd
	dw.CreateParameter cmd, "@p_SalesOrg", adVarchar, adParaminput, 50, sSalesOrg
	dw.CreateParameter cmd, "@p_ConfigCd", adVarchar, adParaminput, 15, sConfigCd
	dw.CreateParameter cmd, "@p_ShowOnPm", adBoolean, adParaminput, 1, sShowOnPm
	dw.CreateParameter cmd, "@p_ShowPhWebActionItems", adBoolean, adParaminput, 1, sShowPhWebActionItems
	dw.CreateParameter cmd, "@p_AVSeriesName", adVarchar, adParaminput, 50, sAVSeriesName
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParaminput, 50, m_UserFullName
	returnValue = dw.ExecuteNonQuery(cmd)

	If returnValue <> 1 Then
		' Abort Transaction
		cn.RollbackTrans()
		Exit Sub
	End If

	sMode = "close"
	cn.CommitTrans

End Sub

If LCase(Request.Form("hidMode")) = "save" Then
	Call Save()
Else
	Call Main()
End If
%>
<html>
<head>
<title>KMAT Details</title>
<link rel="stylesheet" type="text/css" href="../style/excalibur.css" />
<script type="text/javascript">
function Body_OnLoad()
{

    if (frmMain.hidMode.value == "close")
    {
        parent.window.parent.modalDialog.cancel(true);
        //window.close();

    }
	

	if (frmMain.txtKmat.value == "")
		EditKmat();
		
	if (frmMain.txtEzcKmat.value == "")
		EditEzcKmat();
		
	if (frmMain.txtProjectCd.value == "")
		EditProjectCd();
		
	if (frmMain.txtPlantCd.value == "")
		EditPlantCd();
		
	if (frmMain.txtSalesOrg.value == "")
		EditSalesOrg();
		
    if (frmMain.txtConfigCd.value == "")
        EditConfigCd();

    if (frmMain.txtAvSeriesName.value == "")
        EditAVSeriesName();    
	
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
	}
	
}

function EditKmat()
{
	kmat.style.display="";
	kmatText.style.display="none";
}

function EditEzcKmat()
{
	ezcKmat.style.display="";
	ezcKmatText.style.display="none";
}

function EditPlantCd()
{
	plantCd.style.display="";
	plantCdText.style.display="none";
}

function EditProjectCd()
{
	projectCd.style.display="";
	projectCdText.style.display="none";
}

function EditSalesOrg()
{
	salesOrg.style.display="";
	salesOrgText.style.display="none";
}

function EditConfigCd()
{
	configCd.style.display="";
	configCdText.style.display="none";
}

function EditAVSeriesName() {
    divAVSeriesName.style.display = "";
    divAVSeriesNameText.style.display = "none";
}

</script>
</head>
<body onload="Body_OnLoad()">
<form method="post" id="frmMain">
<input id="hidMode" name="hidMode" type="hidden" value="<%= LCase(sMode)%>" />
<table class="FormTable" style="background-color:cornsilk; width:100%; border-width:1px; border-spacing: 0px; padding:1px; border-color:tan;">
	<tr>
		<th>KMAT</th>
		<td><div id="kmat" style="display:none"><input type="text" id="txtKmat" name="txtKmat" maxlength="10" value="<%= sKmat%>" /></div><div id="kmatText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sKmat)%></td><td style="border:none; text-align:right"><a href="javascript:EditKmat();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>EZCKMAT</th>
		<td><div id="ezcKmat" style="display:none"><input type="text" id="txtEzcKmat" name="txtEzcKmat" maxlength="10" value="<%= sEzcKmat%>" /></div><div id="ezcKmatText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sEzcKmat)%></td><td style="border:none; text-align:right"><a href="javascript:EditEzcKmat();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>Project Code</th>
		<td><div id="projectCd" style="display:none"><input type="text" id="txtProjectCd" name="txtProjectCd" value="<%= sProjectCd%>" style="width:300px" /></div><div id="projectCdText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sProjectCd)%></td><td style="border:none; text-align:right"><a href="javascript:EditProjectCd();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>Plant Code</th>
		<td><div id="plantCd" style="display:none"><input type="text" id="txtPlantCd" name="txtPlantCd" value="<%= sPlantCd%>" style="width:300px" /></div><div id="plantCdText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sPlantCd)%></td><td style="border:none; text-align:right"><a href="javascript:EditPlantCd();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>Config Code</th>
		<td><div id="configCd" style="display:none"><input type="text" id="txtConfigCd" name="txtConfigCd" value="<%= sConfigCd%>" style="width:300px" /></div><div id="configCdText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sConfigCd)%></td><td style="border:none; text-align:right"><a href="javascript:EditConfigCd();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>Sales Org</th>
		<td><div id="salesOrg" style="display:none"><input type="text" id="txtSalesOrg" name="txtSalesOrg" value="<%= sSalesOrg%>" style="width:300px" /></div><div id="salesOrgText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sSalesOrg)%></td><td style="border:none; text-align:right"><a href="javascript:EditSalesOrg();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>Show On PM</th>
		<td align="left"><input type="checkbox" id="cbShowOnPm" name="cbShowOnPm" <% If bShowOnPm Then %>Checked<% End If %> /></td>
	</tr>
	<tr>
		<th>Show PhWeb</br>Action Items</th>
		<td align="left"><input type="checkbox" id="cbShowPhWebActionItems" name="cbShowPhWebActionItems" <% If bShowPhWebActionItems Then %>Checked<% End If %> /></td>
	</tr>
	<tr>
		<th>AV Series Name</th>
		<td><div id="divAVSeriesName" style="display:none"><input type="text" id="txtAvSeriesName" name="txtAvSeriesName" value="<%= sAVSeriesName%>" /></div><div id="divAVSeriesNameText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sAVSeriesName)%></td><td style="border:none; text-align:right"><a href="javascript:EditAVSeriesName();">Edit</a></td></tr></table></div></td>
	</tr>
</table>
</form>
</body>
</html>
