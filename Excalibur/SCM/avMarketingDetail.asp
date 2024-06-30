<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
'Response.Write Request.QueryString
'Response.End
response.Buffer = true

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set cn = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.Command")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim i
Dim sMode				: sMode = Request.QueryString("Mode")
Dim sAvNo				: sAvNo = ""
Dim sCategoryOpt		: sCategoryOpt = ""
Dim	iCategoryOpt		: iCategoryOpt = ""
Dim sFeatureCat			: sFeatureCat = ""
Dim sGpgDesc			: sGpgDesc = ""
Dim sMarketingDesc		: sMarketingDesc = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sProgramVersion		: sProgramVersion = GetProductVersion(Request("PVID"))
Dim sConfigRules		: sConfigRules = ""
Dim bIdsSkus			: bIdsSkus = false
Dim bIdsCto				: bIdsCto = false
Dim bRctoSkus			: bRctoSkus = false
Dim bRctoCto			: bRctoCto = false
Dim sUpc				: sUpc = "&nbsp;"
Dim saBrands			: saBrands= Split(Request.Form("chkBrand"), ",")
Dim sCbxBrand			: sCbxBrand = ""
Dim sStatus				: sStatus = "A"
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim iBrandID			: iBrandID = ""
Dim sCplBlindDt			: sCplBlindDt = ""
Dim sRasDiscDt			: sRasDiscDt = ""
Dim sWeight				: sWeight = ""
Dim sRTPDt		        : sRTPDt = ""
Dim sPhWebInstruction	: sPhWebInstruction = ""
Dim sSortOrder			: sSortOrder = ""
Dim sSDFFlag            : sSDFFlag = 0

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_IsMarketingUser
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
	m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "COMMERCIALMARKETING")
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "SMBMARKETING")
	End If
	If Not m_IsMarketingUser Then
		m_IsMarketingUser = Security.UserInRole(m_ProductVersionID, "CONSUMERMARKETING")
	End If

	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsMarketingUser Then
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
		PrepForWeb = Server.HTMLEncode( value )
	End If

End Function

Function GetBoolValue( value )

	Select Case UCase( value )
		Case "Y"
			GetBoolValue = true
		Case "N"
			GetBoolValue = false
		Case 1
			GetBoolValue = true
		Case 0
			GetBoolValue = false
		Case "T"
			GetBoolValue = true
		Case "F"
			GetBoolValue = false
		Case Else
			GetBoolValue = false
	End Select

End Function

Sub Main()
'
'TODO: Get AvDetail Data
'
	Set cmd = dw.CreateCommAndSP(cn, "spListBrAnds4Product")
	dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID")
	dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
	
	Do Until rs.EOF
		sCbxBrand = sCbxBrand & "<input type=checkbox id=chkBrand name=chkBrand value=" & rs("ProductBrandID") 
		If Trim(rs("ProductBrandID")) = Trim(Request("BID")) Then
			sCbxBrand = sCbxBrand & " CHECKED "
		End If
		sCbxBrand = sCbxBrand & ">" & rs("Name") & "<BR>"
		rs.MoveNext
	Loop
	
	sCbxBrand = Left(sCbxBrand, Len(sCbxBrand) - 4)

	If Request("AVID") <> "" Then	'Get the values for the request AV

		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)

		sAvNo = rs("AvNo")
		iCategoryOpt = rs("FeatureCategoryID")
		sFeatureCat = rs("AvFeatureCategory")
		sGpgDesc = rs("GPGDescription")
		sMarketingDesc = rs("MarketingDescription")
		sConfigRules = rs("ConfigRules")
		sManufacturingNotes = rs("ManufacturingNotes")
		bIdsSkus = GetBoolValue(rs("IdsSkus_YN"))
		bIdsCto = GetBoolValue(rs("IdsCto_YN"))
		bRctoSkus = GetBoolValue(rs("RctoSkus_YN"))
		bRctoCto = GetBoolValue(rs("RctoCto_YN"))
		sUpc = rs("UPC")
		sStatus = rs("Status")
		iBrandID = rs("ProductBrandID")
		sCplBlindDt = rs("CplBlindDt")
		sRasDiscDt = rs("RasDiscontinueDt")
		sWeight = rs("Weight")
		sRTPDt = rs("RTPDate")
		sPhWebInstruction = rs("PhWebInstruction")
		sSDFFlag = rs("SDFFlag")
		sSortOrder = rs("SortOrder")

		rs.Close

	End If

	Set cmd = dw.CreateCommandSP(cn, "usp_ListAvFeatureCategories")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do Until rs.EOF
		sCategoryOpt = sCategoryOpt & "<OPTION Value='" & rs("AvFeatureCategoryID") & "'"
		If iCategoryOpt = rs("AvFeatureCategoryID") Then
			sCategoryOpt = sCategoryOpt & " SELECTED "
		End If
		sCategoryOpt = sCategoryOpt & ">" & rs("AvFeatureCategory") & "</OPTION>" & VbCrLf
		rs.MoveNext
	Loop

	rs.Close
	
End Sub

Function GetSkuCtoValue( value )
	If lcase(Value) = "y" Then
		GetSkuCtoValue = "X"
	Else
		GetSkuCtoValue = "&nbsp;"
	End If
End Function

Sub Save()
On Error Goto 0
'
' Save AV Entry
'
	Dim returnValue
	Dim iAvId
	
	cn.BeginTrans
	iAvId = Request("AVID")
		
	'Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand_RTPDate")
	'dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
    'dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request("BID")
	'dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("txtRTPDate")
	'dw.ExecuteNonQuery(cmd)
	
    'Save AvDetail data
	Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail")
	dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
	dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Request.Form("hidFeatureCategoryID")
	dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtGPGDescription")
	dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
	dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 40, ""
	dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 50, Request.Form("txtMarketingDiscDate")
	dw.CreateParameter cmd, "@p_UPC", adVarchar, adParamInput, 12, Request.Form("hidUpc")
	dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, ""
	dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
	dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("txtRTPDate")
	dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50, ""

	returnValue = dw.ExecuteNonQuery(cmd)

	If Not (returnValue = 1 Or returnValue = -1) Then
		' Abort Transaction
		Response.Write returnValue
		cn.RollbackTrans()
		Exit Sub
	End If

	sFunction = "close"
	cn.CommitTrans()

	If Request.Form("hidAVID") <> Request.QueryString("AVID") And sMode <> "add" Then
		Response.Redirect "avMarketingDetail.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request("hidAVID")
	End If

End Sub

Function GetCbxValue( value )
	If lcase(value) = "on" Or lcase(value) = "yes" Then
		GetCbxValue = 1
	Else
		GetCbxValue = 0
	End If
End Function

If LCase(sFunction) = "save" Then
	Call Save()
Else
	Call Main()
End If
%>
<html>
<head>
<title>Marketing Detail</title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<link rel="stylesheet" type="text/css" href="../style/excalibur.css" />
<script type="text/javascript">
    function Body_OnLoad() {

	switch (frmMain.hidFunction.value)
	{
		case "close":
			window.close();
			break;
	}		
	
	if (frmMain.txtMarketingDesc.value == "")
	    EditMktgDesc();
	    
	if (frmMain.txtMarketingDiscDate.value == "")
	    EditMktgDiscDate();

	if (frmMain.txtRTPDate.value == "")
	    EditMktgRTPDate();

	if (frmMain.txtMarketingDesc.value == "")
	    EditMktgDesc();
	    
	if (typeof(window.parent.frames["LowerWindow"].frmButtons) == 'object')
	{
		if (window.frmMain.hidMode.value.toLowerCase() == 'edit'||window.frmMain.hidMode.value.toLowerCase() == 'add')
			window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled =false;
	}
	
}

function EditMktgDesc()
{
	mktgDesc.style.display="";
	mktgText.style.display="none";
}

function EditMktgBlindDate() {
    mktgBlindDate.style.display = "";
    mktgBlindDateText.style.display = "none";
}

function EditMktgDiscDate() {
    mktgDiscDate.style.display = "";
    mktgDiscDateText.style.display = "none";
}

function EditMktgRTPDate() {
    mktgRTPDate.style.display = "";
    mktgRTPDateText.style.display = "none";
}

</script>
</head>
<body onload="Body_OnLoad()">
<form method="post" id="frmMain">
<input id="hidMode" name="hidMode" type="hidden" value="<%= LCase(sMode)%>" />
<input id="hidFunction" name="hidFunction" type="hidden" value="<%= LCase(sFunction)%>" />
<input id="hidStatus" name="hidStatus" type="hidden" value="<%= UCase(sStatus)%>" />
<input id="hidAVID" name="hidAVID" type="hidden" value="<%= Request("AVID")%>" />
<input id="hidFeatureCategoryID" name="hidFeatureCategoryID" type="hidden" value="<%= iCategoryOpt%>" />
<input id="txtGPGDescription" name="txtGPGDescription" type="hidden" value="<%= sGpgDesc%>" />
<input id="hidCplBlindDt" name="hidCplBlindDt" type="hidden" value="<%= sCplBlindDt%>" />
<input id="hidRasDidcontinueDt" name="hidRasDidcontinueDt" type="hidden" value="<%= sRasDiscDt%>" />
<input id="hidRTPDt" name="hidRTPDt" type="hidden" value="<%= sRTPDt%>" />
<input id="hidPhWebInstruction" name="hidPhWebInstruction" type="hidden" value="<%= sPhWebInstruction%>" />
<input id="hidUpc" name="hidUpc" type="hidden" value="<%= sUpc%>" />
<input id="BID" name="BID" type="hidden" value="<%= iBrandID%>" />
<input id="PVID" name="PVID" type="hidden" value="<%= m_ProductVersionID %>" />
<table class="FormTable" width="100%" border="1" cellspacing="0" cellpadding="1" style="background-color:cornsilk; border-color:tan;">
	<tr>
		<th>AV#</th>
		<td><%= PrepForWeb(sAvNo)%></td>
	</tr>
	<tr>
		<th>Feature Category</th>
		<td><%= PrepForWeb(sFeatureCat)%></td>
	</tr>
	<tr>
		<th>GPG Description</th>
		<td><%= PrepForWeb(sGpgDesc)%></td>
	</tr>
	<tr>
		<th>Marketing Description</th>
		<td><div id="mktgDesc" style="display:none"><input type="text" id="txtMarketingDesc" name="txtMarketingDesc" value="<%= sMarketingDesc%>" style="width:300px" /></div><div id="mktgText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sMarketingDesc)%></td><td align=Right style="border:none"><a href="javascript:EditMktgDesc();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>AV RTP Date</th>
		<td><div id="mktgRTPDate" style="display:none"><input type="text" id="txtRTPDate" name="txtRTPDate" value="<%= sRTPDt%>" style="width:300px" /></div><div id="mktgRTPDateText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sRTPDt)%></td><td align=Right style="border:none"><a href="javascript:EditMktgRTPDate();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>CPL Blind Date</th>
		<td><div id="mktgBlindDateText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sCplBlindDt)%></td></tr></table></div></td>
	</tr>
	<tr>
		<th>End of Manufacturing</th>
		<td><div id="mktgDiscDate" style="display:none"><input type="text" id="txtMarketingDiscDate" name="txtMarketingDiscDate" value="<%= sRasDiscDt%>" style="width:300px" /></div><div id="mktgDiscDateText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sRasDiscDt)%></td><td align=Right style="border:none"><a href="javascript:EditMktgDiscDate();">Edit</a></td></tr></table></div></td>
	</tr>
	<tr>
		<th>PhWeb Instructions</th>
		<td><div id="mktgPhWebInstructionText"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sPhWebInstruction)%></td></tr></table></div></td>
	</tr>
</table>
</form>
</body>
</html>
