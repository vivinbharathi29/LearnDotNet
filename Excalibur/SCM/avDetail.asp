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

Dim AppRoot
AppRoot = Session("ApplicationRoot")

Dim i
Dim sMode				: sMode = Request.QueryString("Mode")
Dim sAvNo				: sAvNo = ""
Dim sCategoryOpt		: sCategoryOpt = ""
Dim	iCategoryOpt		: iCategoryOpt = ""
Dim sFeatureCat			: sFeatureCat = ""
Dim sGpgDesc			: sGpgDesc = ""
Dim sSortOrder			: sSortOrder = ""
Dim sMarketingDesc		: sMarketingDesc = ""
Dim sMarketingDescPMG	: sMarketingDescPMG = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sProgramVersion		: sProgramVersion = GetProductVersion(Request("PVID"))
Dim sConfigRules		: sConfigRules = ""
Dim bIdsSkus			: bIdsSkus = false
Dim bIdsCto				: bIdsCto = false
Dim bRctoSkus			: bRctoSkus = false
Dim bRctoCto			: bRctoCto = false
Dim bBSAMSkus           : bBSAMSkus = false
Dim bBSAMBparts         : bBSAMBparts = false
Dim sUpc				: sUpc = "&nbsp;"
Dim saBrands			: saBrands= Split(Request.Form("chkBrand"), ",")
Dim sCbxBrand			: sCbxBrand = ""
Dim sStatus				: sStatus = "A"
Dim sFunction			: sFunction = Request.Form("hidFunction")
Dim iBrandID			: iBrandID = ""
Dim sCplBlindDt			: sCplBlindDt = ""
Dim sGeneralAvailDt     : sGeneralAvailDt = ""
Dim sRasDiscDt			: sRasDiscDt = ""
Dim sWeight				: sWeight = ""
Dim sChangeNote         : sChangeNote= ""
Dim sGSEndDt            : sGSEndDt = ""
Dim sRTPDt		        : sRTPDt = ""
Dim sPhWebInstruction	: sPhWebInstruction = ""
'Dim sPDMFeedback    	: sPDMFeedback = ""
Dim sSDFFlag            : sSDFFlag = "False"
Dim sAvId       		: sAvId = ""
Dim sGroup1     		: sGroup1 = ""
Dim sGroup2		        : sGroup2 = ""
Dim sGroup3		        : sGroup3 = ""
Dim sGroup4		        : sGroup4 = ""
Dim sGroup5		        : sGroup5 = ""
Dim iDeliverableRootID  : iDeliverableRootID = ""
Dim iViaAvCreate
dim strAVName
dim strAVName2
dim strAVName3
dim strAVName4
dim strAVName5
dim strAVName6
dim strAVName7
dim strAVName8
dim strAVNameOld
dim strSQL
dim strPCList
dim strCatID
dim	bOriginatedByDCR
dim	iDCRNo
Dim sProdVersionBSAMFlag

strAVNameOld = ""
strAVName = ""
strAVName2 = ""
strAVName3 = ""
strAVName4 = ""
strAVName5 = ""
strAVName6 = ""
strAVName7 = ""
strAVName8 = ""

dim strDeliverableValues
strDeliverableValues = ""

Dim sDeliverableOpt : sDeliverableOpt = ""

dim strElementValues
strElementValues = ""

dim strExistingNameElements
strExistingNameElements = ""

dim strRequiresFormattedName
strRequiresFormattedName = ""

dim IsNameFormatted
IsNameFormatted = "False"

dim strAvPrefixValues
strAvPrefixValues = ""

dim strNameFormats
strNameFormats=""

'Dim sOriginalGpgDesc			: sOriginalGpgDesc = ""
'Dim sOriginalPhWebInstruction	: sOriginalPhWebInstruction = ""
'Dim sOriginalSDFFlag            : sOriginalSDFFlag = "False"
'Dim sOriginalAvNo               : sOriginalAvNo = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim avParentID : avParentID = 0

'##############################################################################	
'
' Create Security Object to get User Info
'
	
	m_EditModeOn = False
	
	Dim Security
	
	
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
	'm_IsSysAdmin = false 'Security.IsSysAdmin()

	m_IsProgramCoordinator = Security.IsProgramCoordinator(m_ProductVersionID)
	m_IsConfigurationManager = Security.IsProgramManager(m_ProductVersionID)
	m_UserFullName = Security.CurrentUserFullName()
	'Response.Write m_UserFullName & "<br>"
	'Response.Write m_IsConfigurationManager  & "<br>"
	'Response.Write m_IsProgramCoordinator  & "<br>"
	
	If m_IsSysAdmin Or m_IsProgramCoordinator Or m_IsConfigurationManager Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		sMode = "view"
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
    sProdVersionBSAMFlag = rs("BSAMFlag") & ""
	
	rs.Close
	Set rs = Nothing
	Set cmd = Nothing
End Function

Function PrepForWeb( value )
	Dim newValue
	If Trim( value ) = "" Or IsNull(value) Or Trim(UCase( value )) = "N" Then
		newValue = "&nbsp;"
	ElseIf Trim(UCase( value )) = "Y" Then
		newValue = "X"
	Else
		newValue = Server.HTMLEncode( value )
		newValue = Replace(newValue, vbCrLf, "<br /><br />")
	End If
	PrepForWeb = newValue
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
		sCbxBrand = sCbxBrand & "<INPUT type=checkbox id=chkBrand name=chkBrand value=" & rs("ProductBrandID") 
		If Trim(rs("ProductBrandID")) = Trim(Request("BID")) Then
			sCbxBrand = sCbxBrand & " CHECKED "
		End If
		sCbxBrand = sCbxBrand & ">" & rs("Name") & "<BR>"
		rs.MoveNext
	Loop
	
	sCbxBrand = Left(sCbxBrand, Len(sCbxBrand) - 4)
	sSDFFlag = "New"
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
		sMarketingDescPMG = rs("MarketingDescriptionPMG")
		sConfigRules = rs("ConfigRules")
		sManufacturingNotes = rs("ManufacturingNotes")
		bIdsSkus = GetBoolValue(rs("IdsSkus_YN"))
		bIdsCto = GetBoolValue(rs("IdsCto_YN"))
		bRctoSkus = GetBoolValue(rs("RctoSkus_YN"))
		bRctoCto = GetBoolValue(rs("RctoCto_YN"))
		bBSAMSkus = GetBoolValue(rs("BSAMSkus_YN"))
		bBSAMBparts = GetBoolValue(rs("BSAMBparts_YN"))
		sUpc = rs("UPC")
		sStatus = rs("Status")
		iBrandID = rs("ProductBrandID")
		if rs("CplBlindDt") <> "1/1/1900" then
		    sCplBlindDt = rs("CplBlindDt")
		end if
		if rs("GeneralAvailDt") <> "1/1/1900" then
		    sGeneralAvailDt = rs("GeneralAvailDt")
		end if
		if rs("RasDiscontinueDt") <> "1/1/1900" then
		    sRasDiscDt = rs("RasDiscontinueDt")
		end if
		sWeight = rs("Weight")
        sChangeNote= rs("ChangeNote")
		sSortOrder = rs("SortOrder")
		if rs("GSEndDt") <> "1/1/1900" then
		    sGSEndDt = rs("GSEndDt")
		end if
		if rs("RTPDate") <> "1/1/1900" then
		    sRTPDt = rs("RTPDate")
		end if
		sPhWebInstruction = rs("PhWebInstruction")
		sSDFFlag = rs("SDFFlag")
		sAvId = rs("AvId")
		sGroup1 = rs("Group1")
		sGroup2 = rs("Group2")
		sGroup3 = rs("Group3")
		sGroup4 = rs("Group4")
		sGroup5 = rs("Group5")
		iDeliverableRootID = rs("DeliverableRootID")
		iViaAvCreate = rs("ViaAvCreate") & ""
		bOriginatedByDCR = rs("OriginatedByDCR")
		iDCRNo = rs("DCRNo") & ""
		
		strAVName3 = rs("GPGDescription")
		strAVName5 = rs("MarketingDescription")
		strAVName7 = rs("MarketingDescriptionPMG")
		strExistingNameElements = rs("NameElements")		
		avParentID = rs("ParentID")

		'sPDMFeedback = rs("PDMFeedback")
		
		'if trim(sPDMFeedback) & "" <> "" then
		'    sPDMFeedback = Replace(rs("PDMFeedback"),"][", "]" & vbCrLf & "[")
		'end if
		
		'sOriginalGpgDesc = rs("GPGDescription")		
		'sOriginalPhWebInstruction = rs("PhWebInstruction")
		'sOriginalSDFFlag = rs("SDFFlag")    
		'sOriginalAvNo = rs("AvNo") 

		rs.Close

	End If
	
	If sMode = "clone" Then
		sAvNo = ""
		'sMode = "add"
	End If

	Set cmd = dw.CreateCommandSP(cn, "usp_ListAvFeatureCategories")
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request("BID")
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	Do Until rs.EOF
	    'If rs("AvFeatureCategoryID") <> 1 and rs("AvFeatureCategoryID") <> 86 Then
		    strNameFormats = strNameFormats & "<option value=""" & rs("AvFeatureCategoryID") &   """>" & rs("NameFormat") & "</option>"
		    sCategoryOpt = sCategoryOpt & "<OPTION Value='" & rs("AvFeatureCategoryID") & "'"
		    If iCategoryOpt = rs("AvFeatureCategoryID") Then
			    sCategoryOpt = sCategoryOpt & " SELECTED "
		    End If
		    sCategoryOpt = sCategoryOpt & ">" & rs("AvFeatureCategory") & "</OPTION>" & VbCrLf
		'End If
		rs.MoveNext
	Loop

	rs.Close
	
	strSQL = "usp_SelectAvFeatureCategoryAvPrefixValues"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		if strAvPrefixValues = "" then
			strAvPrefixValues = "'" & rs("AvFeatureCategoryID") & "|" & rs("AvPrefix")
		else
			strAvPrefixValues =  strAvPrefixValues & ";" & rs("AvFeatureCategoryID") & "|" & rs("AvPrefix")
		end if
		rs.MoveNext
	loop
	rs.Close
	strAvPrefixValues =  strAvPrefixValues & "'"


	strSQL = "usp_SelectAvElementDDLValues"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
	if strElementValues = "" then
		strElementValues = "'" & rs("ID") & "|" & rs("ElementID") & "|" & rs("ElementValue") & "|" & rs("Value3") & "|" & rs("Value5") & "|" & rs("Value7")
	else
		strElementValues =  strElementValues & ";" & rs("ID") & "|" & rs("ElementID") & "|" & rs("ElementValue") & "|" & rs("Value3") & "|" & rs("Value5") & "|" & rs("Value7")
	end if
	rs.MoveNext
	loop
	rs.Close
	strElementValues =  strElementValues & "'"
	
	'strSQL = "usp_SelectDeliverablesByAvFeatureCategory"
	'rs.Open strSQL,cn,adOpenForwardOnly
	'do while not rs.EOF
	'if strDeliverableValues = "" then
	'    strDeliverableValues = "'" & rs("Name") & "|" & rs("ID") & "|" & rs("AvFeatureCategoryID") 
	'else
	'    strDeliverableValues =  strDeliverableValues & ";" & rs("Name") & "|" & rs("ID") & "|" & rs("AvFeatureCategoryID")
	'end if
	'rs.MoveNext
	'loop
	'rs.Close
	'strDeliverableValues =  strDeliverableValues & "'"
	
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
	If Not m_EditModeOn Then
		Response.Write "<H3>Insufficient User Privileges</H3><H4>Access Denied</H4>"
		Response.End
	End If


	Dim returnValue
	Dim iAvId
	
	cn.BeginTrans
	
	Dim AvPcActionTypeID
	
'Save AvDetail data
	If LCase(sMode) = "add" Or LCase(sMode) = "clone" Then
		
		'AvPcActionTypeID = NULL
		'If GetCbxBlnValue(Request.Form("chkSDFFlag")) = 1 Then
		'   AvPcActionTypeID = 3 'Add new AV to SCM (with SDF flag ON)
		'Else
		'   AvPcActionTypeID = 4 'Add new AV to SCM (with SDF flag OFF)
		'End If
		
		Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvDetail")
		cmd.NamedParameters = True
		dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
		dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Request.Form("cboCategory")
		dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtAvGpgDescription")
		dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
		dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 100, Request.Form("txtMarketingDescPMG")
		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkShowOnScm"))
		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
		dw.CreateParameter cmd, "@p_NameElements", adVarchar, adParamInput, 500, Request.Form("strNameElements")
		dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamOutput, 8, ""
		'dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		returnValue = dw.ExecuteNonQuery(cmd)
		
		iAvId = cmd("@p_AvDetailID")
	ElseIf Request("AVID") <> "" Then
		iAvId = Request("AVID")
		
		'AvPcActionTypeID = NULL
		'If Request.Form("txtAvGpgDescription") <> Request.Form("hidOriginalGpgDescription") Then
		'   AvPcActionTypeID = 7 'Change GPG Description
		'End If
		
		'If ISNULL(AvPcActionTypeID) = False Then
		'    Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvActionItem")
		'    dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		'    dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
		'    returnValue = dw.ExecuteNonQuery(cmd)
		'End If
		
		'AvPcActionTypeID = NULL
		'AvNumber = NULL
		'If (Request.Form("AvNo") <> Request.Form("hidOriginalAvNo")) And Request.Form("hidSDFFlag") = "True" Then
		'   AvPcActionTypeID = 11 'Add AV# to existing record (with SDF Flag ON)
		'ElseIf (Request.Form("AvNo") <> Request.Form("hidOriginalAvNo")) And Request.Form("hidSDFFlag") = "False" Then
		'  AvPcActionTypeID = 12 'Add AV# to existing record (with SDF Flag OFF)   
		'End If
		
		'If ISNULL(AvPcActionTypeID) = False Then
		'    Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvActionItem")
		'    dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		'    dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
		'    returnValue = dw.ExecuteNonQuery(cmd)
		'End If
		   
		
		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail")
		dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
		dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
		dw.CreateParameter cmd, "@p_FeatureCategoryID", adInteger, adParamInput, 8, Request.Form("cboCategory")
		dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtAvGpgDescription")
		dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
		dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 100, Request.Form("txtMarketingDescPMG")
		dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 8, Request.Form("hidCplBlindDt")
		dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 50, Request.Form("hidRasDiscDt")
		dw.CreateParameter cmd, "@p_UPC", adVarchar, adParamInput, 12, Request.Form("hidUpc")
		dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, Request.Form("txtChangeReason")
		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkShowOnScm"))
		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
		dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("hidRTPDt")
		dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 8, Request.Form("hidGeneralAvailDt")
		dw.CreateParameter cmd, "@p_NameElements", adVarchar, adParamInput, 500, Request.Form("strNameElements")
		'dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		dw.CreateParameter cmd, "@p_weight", adInteger, adParamInput, 8, Request.Form("txtWeight")
        If Request.Form("cboDeliverables") > 0 Then
		   dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request.Form("cboDeliverables")
		End If  
      
     	returnValue = dw.ExecuteNonQuery(cmd)

	End If
	
	If returnValue = 0 Then
		' Abort Transaction
		'Response.Write "Error while saving Av Detail"
		'cn.RollbackTrans()
		'Exit Sub
	
	End If

' Link AV to Product_Brands
	If Request("AVID") <> "" Then
		saBrands = Split(Request("BID"), ",")
	End If
	
	For i = LBound(saBrands) To UBound(saBrands)  
	  'If LCase(sMode) <> "add" AND LCase(sMode) <> "clone" Then
		'AvPcActionTypeID = NULL
		'If Request.Form("hidSDFFlag") = "True" And Request.Form("hidOriginalSDFFlag") = "False" Then
		'   AvPcActionTypeID = 5 'Update AV SDF Flag from OFF to ON
		'Elseif Request.Form("hidSDFFlag") = "False" And Request.Form("hidOriginalSDFFlag") = "True" Then
		'   AvPcActionTypeID = 6 'Update AV SDF Flag from ON to OFF
		'End If
		
		'If ISNULL(AvPcActionTypeID) = False Then
		'    Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvActionItem")
		'    dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		'    dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
		'    returnValue = dw.ExecuteNonQuery(cmd)
		'End If
		
		'If Request.Form("txtPhWebInstruction") <> Request.Form("hidOriginalPhWebInstruction") Then
		'    Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvActionItem")
		'    dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, 8 'Any change to PhWeb Instructions
		'    dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
		'    returnValue = dw.ExecuteNonQuery(cmd)
		'End If  
	  'End If
	    If Not IsNull(iAvId) Then
		    ' Add Records to AvDetail_ProductBrand
		    Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand")
		    dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, iAvId
		    dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, saBrands(i)
		    dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, Request.Form("hidStatus")
		    dw.CreateParameter cmd, "@p_ProgramVersion", adVarchar, adParamInput, 5, ""
		    dw.CreateParameter cmd, "@p_ConfigRules", adVarchar, adParamInput, 2000, Request.Form("txtConfigRules")
		    dw.CreateParameter cmd, "@p_ManufacturingNotes", adVarchar, adParamInput, 2000, Request.Form("txtManufacturingNotes")
		    dw.CreateParameter cmd, "@p_IdsSkus_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkIdsSkus"))
		    dw.CreateParameter cmd, "@p_IdsCto_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkIdsCto"))
		    dw.CreateParameter cmd, "@p_RctoSkus_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkRctoSkus"))
		    dw.CreateParameter cmd, "@p_RctoCto_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkRctoCto"))
		    dw.CreateParameter cmd, "@p_SortOrder", adInteger, adParamInput, 4, Request.Form("txtSortOrder")
		    dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, Request.Form("txtChangeReason")
		    dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 1, GetCbxBlnValue(Request.Form("chkShowOnScm"))
		    dw.CreateParameter cmd, "@p_GSEndDt", adDate, adParamInput, 8, Request.Form("txtGSEndDt")
		    dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
		    dw.CreateParameter cmd, "@p_SDFFlag", adBoolean, adParamInput, 1,  GetCbxBlnValue(Request.Form("chkSDFFlag"))
		    dw.CreateParameter cmd, "@p_PhWebInstruction", adVarchar, adParamInput, 500, Request.Form("txtPhWebInstruction")
		    dw.CreateParameter cmd, "@p_AvId", adVarchar, adParamInput, 2000, Request.Form("txtAvId")
		    dw.CreateParameter cmd, "@p_Group1", adVarchar, adParamInput, 2000, Request.Form("txtGroup1")
		    dw.CreateParameter cmd, "@p_Group2", adVarchar, adParamInput, 2000, Request.Form("txtGroup2")
		    dw.CreateParameter cmd, "@p_Group3", adVarchar, adParamInput, 2000, Request.Form("txtGroup3")
		    dw.CreateParameter cmd, "@p_Group4", adVarchar, adParamInput, 2000, Request.Form("txtGroup4")
		    dw.CreateParameter cmd, "@p_Group5", adVarchar, adParamInput, 2000, Request.Form("txtGroup5")
		    dw.CreateParameter cmd, "@p_OriginatedByDCR", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkDCR"))
		    dw.CreateParameter cmd, "@p_DCRNo", adInteger, adParamInput, 8, Request.Form("txtDCRNo") 
            If (sProdVersionBSAMFlag = "True") Then
		        dw.CreateParameter cmd, "@p_BSAMSkus_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkBSAMSkus"))
		        dw.CreateParameter cmd, "@p_BSAMBparts_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkBSAMBparts"))
		    End If
		    returnValue = dw.ExecuteNonQuery(cmd)
		End If
	Next
	
	sFunction = "close"
	cn.CommitTrans
	
	If Request.Form("hidAVID") <> Request.QueryString("AVID") And sMode <> "add" Then
		Response.Redirect "avDetail.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request("hidAVID")
	End If

End Sub

Sub NewPage()
	If Request.Form("hidAVID") <> Request.QueryString("AVID") And sMode <> "add" Then
		Response.Redirect "avDetail.asp?Mode=" & Request("MODE") & "&PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request("hidAVID")
	End If
End Sub

Function GetCbxValue( value )
	If lcase(value) = "on" Or lcase(value) = "yes" Then
		GetCbxValue = "Y"
	Else
		GetCbxValue = "N"
	End If
End Function

Function GetCbxBlnValue( value )
	If lcase(value) = "on" Or lcase(value) = "yes" Then
		GetCbxBlnValue = 1
	Else
		GetCbxBlnValue = 0
	End If
End Function

If LCase(sFunction) = "save" And LCase(sMode) = "view" Then
	Call NewPage()
ElseIf LCase(sFunction) = "save" Then
	Call Save()
Else
	Call Main()
End If

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="../style/excalibur.css">
<script src="../includes/client/jquery.min.js" type="text/javascript"></script>
<SCRIPT type="text/javascript">
    $(function () {
            var CheckIllegalChars = function (e) {
		    var pastedData;

		    if (e.originalEvent.clipboardData === undefined)//IE
			    pastedData = clipboardData.getData('text');
		    else
			    pastedData = e.originalEvent.clipboardData.getData('text');

		    var clean = pastedData.replace(/[^\x20-\x7E \r\n]/g, '_');//replace every non printable character with '_'

		    if (clean != pastedData) {
			    var msg = 'Invalid characters detected, they might be displayed incorrectly after saving (location marked with _):\n'
			    alert(msg + clean);
		    }
	    }

	    var ua = window.navigator.userAgent;
	    var msie = ua.indexOf("MSIE ");
	    if (msie > 0) // If Internet Explorer
	    {
	    	$("body").off('paste');
	    	$("body").on('paste', function (e) { CheckIllegalChars(e); });
	    }
	    else
	    {
	    	$(document).on('paste', function (e) { CheckIllegalChars(e); });
	    }
    });

    function Body_OnLoad() {
        switch (frmMain.hidMode.value) {
            case "add":
                EditAv();
                EditGpgDesc();
                EditMktgDesc();
                EditMktgDescPMG();
                EditFeatureCat();
                EditMktgPhWebInstruction();
                chkSDFFlag_onclick();
                break;
            case "clone":
                EditAv();
                EditGpgDesc()
                EditMktgDesc();
                EditMktgDescPMG();
                EditFeatureCat();
                EditMktgPhWebInstruction();
                chkSDFFlag_onclick();
                break;
        }

        switch (frmMain.hidFunction.value) {
            case "close":
                window.close();
                break;
        }

        if (typeof (window.parent.frames["LowerWindow"].frmButtons) == 'object') {
            if (window.frmMain.hidMode.value.toLowerCase() == 'edit' || window.frmMain.hidMode.value.toLowerCase() == 'add' || window.frmMain.hidMode.value.toLowerCase() == 'clone')
                window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = false;
        }

        LoadExistingDDLValues(frmMain.ElementValues.value);

        //var ID = frmMain.ID.value;
        var ExistingValues = frmMain.ExistingNameElements.value;
        if (frmMain.CategoryOpt.value != null && (frmMain.ViaAvCreate.value == 0 || window.frmMain.hidMode.value.toLowerCase() == 'add')) {
            LoadDelRootValues(frmMain.CategoryOpt.value, frmMain.DelRootID.value);
        }
        //	else {
        //		delRootCbo.style.display = "none";
        //		delRootText.style.display = "";
        //	}
    }

    function LoadDelRootValues(AvFeatureCategoryID, DelRootID) {
        var parameters = "function=GetDelRootValues&AvFeatureCategoryID=" + AvFeatureCategoryID + "&PVID=" + frmMain.ProductVersionID.value + "&DelRootID=" + DelRootID;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {        //Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else {                            //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        }
        //End setup Ajax.
        request.open("POST", "<%= AppRoot %>/SCM/AvDetailDelRootValues.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
        if (request.responseText == "<select id=\"cboDeliverables\" name=\"cboDeliverables\" style=\"WIDTH: 70%\"><OPTION VALUE=0>--- Please Make a Selection ---</OPTION></select>") {
            //disable div delRootCbo
            delRootCbo.style.display = "none";
            delRootText.style.display = "";
            var devEdit = document.getElementById("divEditDeliverableRoot");
            if (devEdit!=null)
                devEdit.style.display = "";
            frmMain.cboDeliverables.options(frmMain.cboDeliverables.selectedIndex).value = 0;
        } else {
            document.getElementById("delRootCbo").innerHTML = request.responseText;
            var devEdit = document.getElementById("divEditDeliverableRoot");
            if (devEdit != null)
            devEdit.style.display = "";
        }
    }

    function EditAv() {
        avinput.style.display = "";
        avtext.style.display = "none";
    }

    function EditGpgDesc() {
        gpgDesc.style.display = "";
        gpgText.style.display = "none";
    }

    function EditMktgDesc() {
        mktgDesc.style.display = "";
        mktgText.style.display = "none";
    }

    function EditMktgDescPMG() {
        mktgDescPMG.style.display = "";
        mktgTextPMG.style.display = "none";
    }

    function EditDeliverableRoot() {
        delRootCbo.style.display = "";
        delRootText.style.display = "none";
    }

    function EditFeatureCat() {
        if (frmMain.IsNameFormatted.value == "True") {
            trGPGDesc.style.display = "none";
            trMarketingDesc40.style.display = "none";
            trMarketingDesc100.style.display = "none";

            if (frmMain.m_EditModeOn == "True") {
                divEditGpgDesc.style.display = ""
                divEditMktgDesc.style.display = ""
                divEditMktgDescPMG.style.display = ""
            }
        }
        featureCatSelect.style.display = "";
        featureCatText.style.display = "none";
    }

    function EditWeight() {
        WeightInput.style.display = "";
        WeightTxt.style.display = "none";
    }

    function EditGSEndDt() {
        GSEndDtInput.style.display = "";
        GSEndDtTxt.style.display = "none";
    }

    function EditMktgPhWebInstruction() {
        mktgPhWebInstruction.style.display = "";
        mktgPhWebInstructionText.style.display = "none";
    }

    function chkSDFFlag_onclick() {
        var chkSDFFlag = document.getElementById("chkSDFFlag");
        var hidSDFFlag = document.getElementById("hidSDFFlag");
        if (chkSDFFlag.checked == true) {
            hidSDFFlag.value = "True";
        } else {
            hidSDFFlag.value = "False";
        }
    }

    function chkDCR_onclick() {
        var chkDCR = document.getElementById("chkDCR");
        var divDVR = document.getElementById("divDVR");
        if (chkDCR.checked == true) {
            divDVR.style.display = "inline";
        } else {
            divDVR.style.display = "none";
        }
    }

    function ViewAvActionItems(AvId) {
        var strID;
        strID = window.parent.showModalDialog("<%= AppRoot %>/SCM/PDMFeedbackFrame.asp?AvId=" + AvId, "", "dialogWidth:1095px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
    }

    function cboCategory_onchange(strElementValues) {
        var i;
        var j;
        var TypeID;
        var strNameFormat = "";
        var strNewRow = "";
        frmMain.IsNameFormatted.value = "False";
        var strDeliverables = frmMain.DeliverableValues.value;

        if (frmMain.ViaAvCreate.value == 0) {
            LoadDelRootValues(frmMain.cboCategory.value, frmMain.DelRootID.value);
        }

        for (i = 0; i < frmMain.cboNameFormat.length; i++)
            if (frmMain.cboNameFormat.options[i].value == frmMain.cboCategory.value)
                if (frmMain.cboNameFormat[i].text != "") {
                    strNameFormat = frmMain.cboNameFormat[i].text;
                }

        frmMain.tagCategory.value = frmMain.cboCategory.value;

        if (strNameFormat != "") //&& CloningID == ""
        {
            frmMain.IsNameFormatted.value = "True";

            trGPGDesc.style.display = "none";
            trMarketingDesc40.style.display = "none";
            trMarketingDesc100.style.display = "none";

            divEditGpgDesc.style.display = "";
            divEditMktgDesc.style.display = "";
            divEditMktgDescPMG.style.display = "";

            var NewRows = strNameFormat.split(";");
            var FormatParts;
            var j;

            var Elements = strElementValues.split(";");
            var ElementValues;
            var k;
            var cboValues;
            cboValues = "";

            strNewRow = "<table id=tbName oncontextmenu=displayMenu()>"
            var ExistingValues = frmMain.ExistingNameElements.value;
            //if (ExistingValues != "") {
            ExistingValues = ExistingValues.split("|");
            strNewRow = strNewRow + "<tr><td><font size=1><b>GPG Description:&nbsp;&nbsp;&nbsp;<br />(40-char PhWeb)</b></font></td><td><label ID=lblFinishedName3><font color=black>" + frmMain.txtAVName3.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Description:&nbsp;&nbsp;&nbsp;<br />(40 Char GPSy)</b></font></td><td><label ID=lblFinishedName5><font color=black>" + frmMain.txtAVName5.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Description:&nbsp;&nbsp;&nbsp;<br />(100 Char PMG)</b></font></td><td><label ID=lblFinishedName7><font color=black>" + frmMain.txtAVName7.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "</table><table id=tbEdit>"

            for (i = 0; i < NewRows.length; i++)
                if (NewRows[i] != "") {
                    FormatParts = NewRows[i].split("|");
                    if (FormatParts.length == 7) {
                        if (FormatParts[6] == 0) {
                            //alert(ExistingValues[i]);
                            if (typeof (ExistingValues[i]) != "undefined") {
                                strNewRow = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onkeyup=\"return cboElement_onchange()\" value=" + ExistingValues[i] + "><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                            } else {
                                strNewRow = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onkeyup=\"return cboElement_onchange()\"><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                            }
                        }
                        else if (FormatParts[6] == 1) {
                            cboValues = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><select id=cboElement name=cboElement class=cbo style=WIDTH:150px onchange=\"return cboElement_onchange();\"><option></option>";
                            for (k = 0; k < Elements.length; k++) {
                                ElementValues = Elements[k].split("|");
                                if (ElementValues[1] == FormatParts[5]) {
                                    if (ElementValues[0] == ExistingValues[i]) {
                                        cboValues = cboValues + "<Option Selected Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                                    } else {
                                        cboValues = cboValues + "<Option Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                                    }
                                }
                            }
                            if (cboValues != "") {
                                strNewRow = cboValues + "</select><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>";
                                cboValues = "";
                            }
                        }
                    }
                }
            //if (lblDisplayedID.innerText != "") {
            strNewRow = strNewRow + "</table>";
            NameRowUpdate.innerHTML = strNewRow;
            NameRowUpdate.style.display = "";
            //frmMain.txtAVName.style.display = "none";
            //}
        }
        else {
            trGPGDesc.style.display = "";
            trMarketingDesc40.style.display = "";
            trMarketingDesc100.style.display = "";

            divEditGpgDesc.style.display = ""
            divEditMktgDesc.style.display = ""
            divEditMktgDescPMG.style.display = ""

            //if (lblDisplayedID.innerText != "") {
            NameRowUpdate.style.display = "none";
            //frmMain.txtAVName.style.display = "";
            //}

        }

    }

    function cboElement_onchange() {
        //var strBuild = "";
        //var strName2 = "";
        var strName3 = "";
        //var strName4 = "";
        var strName5 = "";
        //var strName6 = "";
        var strName7 = "";
        //var strName8 = "";

        var strElementValues = frmMain.ElementValues.value;
        var strAvPrefixValues = frmMain.AvPrefixValues.value;

        strElementValues = strElementValues.replace("\'", "");
        strElementValues = strElementValues.replace("'", "");

        strAvPrefixValues = strAvPrefixValues.replace("\'", "");
        strAvPrefixValues = strAvPrefixValues.replace("'", "");

        var Elements = strElementValues.split(";");
        var Prefixes = strAvPrefixValues.split(";");

        var ElementValues;
        var PrefixValues;

        var strComments = "";
        var i;
        var Element;

        var Prefix = "";

        if (typeof (frmMain.cboElement.length) == "undefined") {
            //strBuild = strBuild + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            //strName2 = strName2 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            strName3 = strName3 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            //strName4 = strName4 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            strName5 = strName5 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            //strName6 = strName6 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            strName7 = strName7 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            //strName8 = strName8 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
        }
        else {
            for (i = 0; i < frmMain.cboElement.length; i++) {
                //alert(frmMain.cboElement(i).tagName);
                if (frmMain.cboElement(i).tagName == "INPUT") {
                    //alert(Elements);
                    //alert(frmMain.cboElement(i).className);
                    if (frmMain.cboElement(i).value != "" && frmMain.cboElement(i).className == "name=cboElement") {
                        //strBuild = strBuild + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        //strName2 = strName2 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        strName3 = strName3 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        //strName4 = strName4 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        strName5 = strName5 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        //strName6 = strName6 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        strName7 = strName7 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                        //strName8 = strName8 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                    } else if (frmMain.cboElement(i).value != "") {
                        for (k = 0; k < Elements.length; k++) {
                            ElementValues = Elements[k].split("|");
                            if (ElementValues[1] == frmMain.cboElement(i).className) {
                                //alert(frmMain.cboElement(i).className);
                                if (trim(ElementValues[3]) == "[text]") {
                                    //alert(frmMain.cboElement(i).value);
                                    strName3 = trim(strName3) + " " + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                                if (trim(ElementValues[4]) == "[text]") {
                                    strName5 = trim(strName5) + " " + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                                if (trim(ElementValues[5]) == "[text]") {
                                    strName7 = trim(strName7) + " " + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                            }
                        }
                    }
                }
                else {
                    if (frmMain.cboElement(i).options(frmMain.cboElement(i).selectedIndex).value != "") {
                        for (k = 0; k < Elements.length; k++) {
                            ElementValues = Elements[k].split("|");
                            if (ElementValues[0] == frmMain.cboElement(i).options(frmMain.cboElement(i).selectedIndex).value) {
                                if (trim(ElementValues[3]) != "") {
                                    strName3 = trim(strName3) + " " + lblPreNamePart(i).innerText + ElementValues[3] + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                                if (trim(ElementValues[4]) != "") {
                                    strName5 = trim(strName5) + " " + lblPreNamePart(i).innerText + ElementValues[4] + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                                if (trim(ElementValues[5]) != "") {
                                    strName7 = trim(strName7) + " " + lblPreNamePart(i).innerText + ElementValues[5] + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                                }
                            }
                        }
                    }
                }
            }
        }
        //frmMain.txtAVName.value = strBuild;
        //lblFinishedName.innerText = strBuild;

        //frmMain.txtAVName2.value = strName2;
        //lblFinishedName2.innerText = strName2;

        frmMain.txtAVName3.value = strName3;
        lblFinishedName3.innerText = strName3;

        //frmMain.txtAVName4.value = strName4;
        //lblFinishedName4.innerText = strName4;

        frmMain.txtAVName5.value = strName5;
        lblFinishedName5.innerText = strName5;

        //frmMain.txtAVName6.value = strName6;
        //lblFinishedName6.innerText = strName6;

        frmMain.txtAVName7.value = strName7;
        lblFinishedName7.innerText = strName7;

        //frmMain.txtAVName8.value = strName8;
        //lblFinishedName8.innerText = strName8;

        //alert(Prefixes.length);
        //alert(Prefixes);
        for (l = 0; l < Prefixes.length; l++) {
            PrefixValues = Prefixes[l].split("|");
            //alert(PrefixValues[0]);
            if (PrefixValues[0] == frmMain.tagCategory.value) {
                //alert(PrefixValues[1]);
                if (trim(PrefixValues[1]) != "") {
                    frmMain.txtAVName3.value = trim(PrefixValues[1]) + " " + strName3;
                    lblFinishedName3.innerText = trim(PrefixValues[1]) + " " + strName3;
                }
                if (trim(PrefixValues[2]) != "") {
                    frmMain.txtAVName5.value = trim(PrefixValues[2]) + " " + strName5;
                    lblFinishedName5.innerText = trim(PrefixValues[2]) + " " + strName5;
                }
                if (trim(PrefixValues[3]) != "") {
                    frmMain.txtAVName7.value = trim(PrefixValues[3]) + " " + strName7;
                    lblFinishedName7.innerText = trim(PrefixValues[3]) + " " + strName7;
                }
            }
        }
        //        alert(frmMain.txtAVName2.value);
        //        alert(frmMain.txtAVName3.value);
        //        alert(frmMain.txtAVName4.value);
        //        alert(frmMain.txtAVName5.value);
        //        alert(frmMain.txtAVName6.value);
        //        alert(frmMain.txtAVName7.value);
        //        alert(frmMain.txtAVName8.value);
    }

    function trim(stringToTrim) {
        return stringToTrim.replace(/^\s+|\s+$/g, "");
    }

    function SaveNameElements() {
        var strBuild = "";
        var strComments = "";
        var i;
        //var DelName;
        var cboElement = document.getElementById("cboElement");
        //var txtAVName = document.getElementById("txtAVName").value;
        //alert(txtAVName);
        //DelName = txtAVName.split(" ");
        var Elements = "";
        var Elements2 = "";
        var Elements3 = "";
        var Elements4 = "";
        var Elements5 = "";
        var Elements6 = "";
        var Elements7 = "";
        var Elements8 = "";
        if (cboElement != null) {
            if (typeof (frmMain.cboElement.length) == "undefined") {
                Elements = frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text;
            }
            else {
                for (i = 0; i < frmMain.cboElement.length; i++) {
                    if (frmMain.cboElement(i).tagName == "INPUT") {
                        if (Elements == "") {
                            Elements = frmMain.cboElement(i).value;
                            //Elements2 = replace(DelName[i + 1], frmMain.cboElement(i).text, "");
                        } else {
                            Elements = Elements + "|" + frmMain.cboElement(i).value;
                        }
                    }
                    else {
                        if (Elements == "") {
                            Elements = frmMain.cboElement(i).options(frmMain.cboElement(i).selectedIndex).value;
                        } else {
                            Elements = Elements + "|" + frmMain.cboElement(i).options(frmMain.cboElement(i).selectedIndex).value;
                        }
                    }
                }
            }
        }
        //alert(frmMain.strNameElements.value);
        frmMain.strNameElements.value = trim(Elements);

        //        alert(frmMain.txtAVName2.value);
        //        alert(frmMain.txtAVName3.value);
        //        alert(frmMain.txtAVName4.value);
        //        alert(frmMain.txtAVName5.value);
        //        alert(frmMain.txtAVName6.value);
        //        alert(frmMain.txtAVName7.value);
        //        alert(frmMain.txtAVName8.value);
    }

    function LoadExistingDDLValues(strElementValues) {
        var i;
        var j;
        var TypeID;
        var strNameFormat = "";
        var strNewRow = "";
        frmMain.IsNameFormatted.value = "False"

        for (i = 0; i < frmMain.cboNameFormat.length; i++)
            if (frmMain.cboNameFormat.options[i].value == frmMain.cboCategory.value)
                if (frmMain.cboNameFormat[i].text != "") {
                    strNameFormat = frmMain.cboNameFormat[i].text;
                }

        frmMain.tagCategory.value = frmMain.cboCategory.value;

        if (strNameFormat != "") { //&& CloningID == ""
            frmMain.IsNameFormatted.value = "True"
            var NewRows = strNameFormat.split(";");
            var FormatParts;
            var j;

            if (frmMain.m_EditModeOn == "True") {
                divEditGpgDesc.style.display = "none"
                divEditMktgDesc.style.display = "none"
                divEditMktgDescPMG.style.display = "none"
            }

            strElementValues = strElementValues.replace("\'", "");
            strElementValues = strElementValues.replace("'", "");

            var Elements = strElementValues.split(";");
            //alert(Elements);
            var ElementValues;
            var k;
            var cboValues;
            cboValues = "";

            strNewRow = "<table id=tbName oncontextmenu=displayMenu()>"
            var ExistingValues = frmMain.ExistingNameElements.value;
            //if (ExistingValues != "") {
            ExistingValues = ExistingValues.split("|");
            strNewRow = strNewRow + "<tr><td><font size=1><b>GPG Description:&nbsp;&nbsp;&nbsp;<br />(40-char PhWeb)</b></font></td><td><label ID=lblFinishedName3><font color=black>" + frmMain.txtAVName3.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Description:&nbsp;&nbsp;&nbsp;<br />(40 Char GPSy)</b></font></td><td><label ID=lblFinishedName5><font color=black>" + frmMain.txtAVName5.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Description:&nbsp;&nbsp;&nbsp;<br />(100 Char PMG)</b></font></td><td><label ID=lblFinishedName7><font color=black>" + frmMain.txtAVName7.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
            strNewRow = strNewRow + "</table><table id=tbEdit>"

            for (i = 0; i < NewRows.length; i++)
                if (NewRows[i] != "") {
                    FormatParts = NewRows[i].split("|");
                    if (FormatParts.length == 7) {
                        if (FormatParts[6] == 0) {
                            //alert(ExistingValues[i]);
                            if (typeof (ExistingValues[i]) != "undefined") {
                                strNewRow = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onkeyup=\"return cboElement_onchange()\" value=" + ExistingValues[i] + "><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                            } else {
                                strNewRow = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onkeyup=\"return cboElement_onchange()\"><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                            }
                        }
                        else if (FormatParts[6] == 1) {
                            cboValues = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><select id=cboElement name=cboElement class=cbo style=WIDTH:150px onchange=\"return cboElement_onchange();\"><option></option>";
                            for (k = 0; k < Elements.length; k++) {
                                ElementValues = Elements[k].split("|");
                                if (ElementValues[1] == FormatParts[5]) {
                                    if (ElementValues[0] == ExistingValues[i]) {
                                        cboValues = cboValues + "<Option Selected Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                                    } else {
                                        cboValues = cboValues + "<Option Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                                    }
                                }
                            }
                            if (cboValues != "") {
                                strNewRow = cboValues + "</select><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>";
                                cboValues = "";
                            }
                        }
                    }
                }
            strNewRow = strNewRow + "</table>";
            NameRowUpdate.innerHTML = strNewRow;
            NameRowUpdate.style.display = "";
            //frmMain.txtAVName.style.display = "none";
        }
    }


</SCRIPT>
</HEAD>
<BODY OnLoad="Body_OnLoad()">
<FORM method=post id=frmMain>
<INPUT id="hidMode" name="hidMode" type=HIDDEN value=<%= LCase(sMode)%>>
<INPUT id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<INPUT id="hidStatus" name="hidStatus" type=HIDDEN value=<%= UCase(sStatus)%>>
<INPUT id="hidGpgDescription" id="hidGpgDescription" type=HIDDEN value="<%= sGpgDesc%>">
<INPUT id="hidAVID" name="hidAVID" type=HIDDEN value=<%= Request("AVID")%>>
<INPUT id="hidCplBlindDt" name="hidCplBlindDt" type=HIDDEN value="<%= sCplBlindDt%>">
<INPUT id="hidRasDiscDt" name="hidRasDiscDt" type=HIDDEN value="<%= sRasDiscDt%>">
<INPUT id="hidUpc" name="hidUpc" type=hidden value="<%= sUpc%>">
<INPUT id="BID" name="BID" type=HIDDEN value="<%= iBrandID%>">
<INPUT id="hidRTPDt" name="hidRTPDt" type=HIDDEN value="<%= sRTPDt%>">
<INPUT id="hidPhWebInstruction" name="hidPhWebInstruction" type=HIDDEN value="<%= sPhWebInstruction%>">
<INPUT id="hidSDFFlag" name="hidSDFFlag" type=HIDDEN value="<%= sSDFFlag%>">
<INPUT id="hidGeneralAvailDt" name="hidGeneralAvailDt" type=HIDDEN value="<%= sGeneralAvailDt%>">
<input id="strNameElements" name="strNameElements" type="hidden">
<input style="Display:none" type="text" id="ExistingNameElements" name="ExistingNameElements" value="<%=strExistingNameElements%>">
<input style="Display:none" type="text" id="RequiresFormattedName" name="RequiresFormattedName" value="<%=strRequiresFormattedName%>">
<input style="Display:none" type="text" id="IsNameFormatted" name="IsNameFormatted" value="<%=IsNameFormatted%>">
<input style="Display:none" type="text" id="ElementValues" name="ElementValues" value="<%=strElementValues%>">
<input style="Display:none" type="text" id="DeliverableValues" name="DeliverableValues" value="<%=strDeliverableValues%>">
<input style="Display:none" type="text" id="txtPCListEmails" name="txtPCListEmails" value="<%=strPCList%>">
<input style="Display:none" type="text" id="txtAVNameOld" name="txtAVNameOld" value="<%=strAVNameOld%>">
<input style="Display:none" type="text" id="txtAVName3" name="txtAVName3" value="<%=strAVName3%>">
<input style="Display:none" type="text" id="txtAVName5" name="txtAVName5" value="<%=strAVName5%>">
<input style="Display:none" type="text" id="txtAVName7" name="txtAVName7" value="<%=strAVName7%>">
<input style="Display:none" type="text" id="AvPrefixValues" name="AvPrefixValues" value="<%=strAvPrefixValues%>">
<input style="Display:none" type="text" id="m_EditModeOn" name="m_EditModeOn" value="<%=m_EditModeOn%>">
<input style="Display:none" type="text" id="DelRootID" name="DelRootID" value="<%=iDeliverableRootID%>">
<input style="Display:none" type="text" id="CategoryOpt" name="CategoryOpt" value="<%=iCategoryOpt%>">
<input style="Display:none" type="text" id="DeliverableOpt" name="DeliverableOpt" value="<%=sDeliverableOpt%>">
<input style="Display:none" type="text" id="ViaAvCreate" name="ViaAvCreate" value="<%=iViaAvCreate%>">
<input style="Display:none" type="text" id="ProductVersionID" name="ProductVersionID" value="<%=m_ProductVersionID%>">
<input id="hidParentID" name="hidParentID" type="hidden" value="<%=avParentID%>" />

<label id="lblDisplayedID" style="Display:none"><%= request("AVID")%></label>
  <table width=100% border=0>
	<tr>
	  <td width=100% align=right><font size=1 face=verdana><a href="#" onclick="ViewAvActionItems('<%= Request("AVID")%>')">AV Action Items</a></font></td>
	</tr>
  </table>

<TABLE class="FormTable" style="display:none;" bgcolor=cornsilk WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=1 bordercolor=tan>
<TR><TH>DCR:</TH>
	<TD><SELECT id=selDCR name=selDCR>
			<OPTION VALUE=0>--- Please Make a Selection ---</OPTION>
			<% %>
		</SELECT></TD></TR>
			
<TABLE class="FormTable" bgcolor=cornsilk WIDTH=100% BORDER=1 CELLSPACING=0 CELLPADDING=1 bordercolor=tan>
	<%	If LCase(sMode) = "add" Then %>
	<TR><TH>Brands</TH>
	<TD>
		<%= sCbxBrand%>
	</TD></TR>
	<%	End If %>   
	<TR>
		<TH>AV#</TH>
		<TD><div id=avinput style="display:none"><INPUT type="text"  id=AvNo name=AvNo maxlength="18" value="<%= sAvNo%>"></div><div id=avtext><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sAvNo)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><a href="javascript:EditAv();">Edit</a></td><%End If%></tr></table></div></TD>
	</TR>
	<tr>
		<th>Feature Category</th>
		<td><div id=featureCatText><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sFeatureCat)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><a href="javascript:EditFeatureCat();">Edit</a></td><%End If%></tr></table></div>
		<div id=featureCatSelect style="display:none">
		<div ID=NameRowUpdate style="Display:none"></div>
			
		<SELECT style="Display:none" id=cboNameFormat name=cboNameFormat><%=strNameFormats%></SELECT>	
		<input id="tagCategory" name="tagCategory" type="hidden" value="<%=trim(strCatID)%>">
		<select id="cboCategory" name="cboCategory" style="WIDTH: 80%" LANGUAGE="javascript" onchange="return cboCategory_onchange(<%=strElementValues%>)">			
		  <OPTION VALUE=0>--- Please Make a Selection ---</OPTION>
			<%=sCategoryOpt%>
		</select>			
		</td>
	</tr>
	<TR>
		<TH>Deliverable Root ID</TH>
		<%	If iViaAvCreate = "1" Then %>
		<TD><%= PrepForWeb(iDeliverableRootID)%>&nbsp;&nbsp;<font size=1>(Cannot Edit - Automatically Assigned)</font></TD>
		<%	Else %> 
		<TD><div id=delRootCbo style="display:none">
				<select id="cboDeliverables" name="cboDeliverables" style="WIDTH: 70%">			
					<OPTION VALUE=0>--- Please Make a Selection ---</OPTION>
					<%=sDeliverableOpt%>
				</select>
			</div>
			<div id=delRootText>
			 <%Dim iDelRootID
			   If iDeliverableRootID = "0" Then iDelRootID = "" Else iDelRootID = PrepForWeb(iDeliverableRootID)
			 %>
				<table width="100%" border=0><tr><td style="border:none"><%= iDelRootID%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><div id="divEditDeliverableRoot"><a href="javascript:EditDeliverableRoot();">Edit</a></div></td><%End If%></tr></table></div></TD>
		<%	End If %> 
	</TR>
	<TR>
		<TH>Originated By DCR</TH>
		<%	If bOriginatedByDCR = "True" Then %>
		    <TD><input type="checkbox" id="chkDCR" name="chkDCR" checked style="WIDTH:16;HEIGHT:16" onclick="return chkDCR_onclick()"><div id=divDVR style="display:inline">&nbsp;&nbsp;&nbsp;DCR Number:&nbsp;<INPUT type="text" style="width:50px" id=txtDCRNo name=txtDCRNo value="<%= iDCRNo%>"></div></TD>
		<%	Else %> 
		    <TD><input type="checkbox" id="chkDCR" name="chkDCR" style="WIDTH:16;HEIGHT:16" onclick="return chkDCR_onclick()"><div id="divDVR" style="display:none;">&nbsp;&nbsp;&nbsp;DCR Number:&nbsp;<INPUT type="text" style="width:50px" id="txtDCRNo" name="txtDCRNo" value="<%= iDCRNo%>"></div></TD>
		<%	End If %> 
	</TR>
	<TR id=trGPGDesc style="">
		<TH>GPG Description</TH>
		<TD><div id=gpgDesc style="display:none"><INPUT type="text" id=txtAvGpgDescription name=txtAvGpgDescription maxlength="50" value="<%= sGpgDesc%>" style="width:300px"></div>
        <div id=gpgText><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sGpgDesc)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><div id="divEditGpgDesc"><a href="javascript:EditGpgDesc();">Edit</a></div></td><%End If%></tr></table></div></TD>
	</TR>
	<TR id=trMarketingDesc40 style="">
		<TH>Marketing Description<br />(40 Char GPSy)</TH>
		<TD><div id=mktgDesc style="display:none"><INPUT type="text" id=txtMarketingDesc name=txtMarketingDesc value="<%= sMarketingDesc%>" style="width:300px"></div><div id=mktgText><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sMarketingDesc)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><div id="divEditMktgDesc"><a href="javascript:EditMktgDesc();">Edit</a></div></td><%End If%></tr></table></div></TD>
	</TR>
	<TR id=trMarketingDesc100 style="">
		<TH>Marketing Description<br />(100 Char PMG)</TH>
		<TD><div id=mktgDescPMG style="display:none"><INPUT type="text" id=txtMarketingDescPMG name=txtMarketingDescPMG value="<%= sMarketingDescPMG%>" style="width:300px"></div><div id=mktgTextPMG><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sMarketingDescPMG)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><div id="divEditMktgDescPMG"><a href="javascript:EditMktgDescPMG();">Edit</a></div></td><%End If%></tr></table></div></TD>
	</TR>
	<TR>
		<TH>Program Version</TH>
		<TD><%= sProgramVersion%></TD>
	</TR>
	<TR>
		<TH>AV RTP Date</TH>
		<TD><%= PrepForWeb(sRTPDt)%></TD>
	</TR>
	<TR>
		<TH>Select Availability<br />(SA) Date</TH>
		<TD><%= PrepForWeb(sCplBlindDt)%></TD>
	</TR>
	<TR>
		<TH>General Availability<br />(GA) Date</TH>
		<TD><%= PrepForWeb(sGeneralAvailDt)%></TD>
	</TR>
	<TR>
		<TH>End of Manufacturing<br />(EM) Date</TH>
		<TD><%= PrepForWeb(sRasDiscDt)%></TD>
	</TR>
	<TR>
		<TH>PhWeb Instructions</TH>
		<TD><div id=mktgPhWebInstruction style="display:none"><TEXTAREA rows=5 maxlength="100" id=txtPhWebInstruction name=txtPhWebInstruction  style="width:300px"><%= sPhWebInstruction%></TEXTAREA></div><div id=mktgPhWebInstructionText><table width="100%" border=0><tr><td style="border:none; width:300px; WORD-BREAK:BREAK-ALL"><%= PrepForWeb(sPhWebInstruction)%></td><td align=Right style="border:none"><a href="javascript:EditMktgPhWebInstruction();">Edit</a></td></tr></table></div></TD>
	</TR>
	<TR>
		<TH>SDF Flag</TH>
		<%If sSDFFlag = "False" or sSDFFlag = "" or ISNULL(sSDFFlag) Then %>
			<TD><input type="checkbox" id="chkSDFFlag" name="chkSDFFlag" style="WIDTH:16;HEIGHT:16" onclick="return chkSDFFlag_onclick()"></TD>
		<%Else%>
			<TD><input type="checkbox" id="chkSDFFlag" name="chkSDFFlag" checked style="WIDTH:16;HEIGHT:16" onclick="return chkSDFFlag_onclick()"></TD>
		<%End If%>
	</TR>
	<TR>
		<TH>Configuration Rules</TH>
		<TD><TEXTAREA rows=5 id=txtConfigRules name=txtConfigRules style="width:300px"><%= sConfigRules%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>AVID</TH>
		<TD><TEXTAREA rows=5 id=txtAvId name=txtAvId style="width:300px"><%= sAvId%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Group 1</TH>
		<TD><TEXTAREA rows=5 id=txtGroup1 name=txtGroup1 style="width:300px"><%= sGroup1%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Group 2</TH>
		<TD><TEXTAREA rows=5 id=txtGroup2 name=txtGroup2 style="width:300px"><%= sGroup2%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Group 3</TH>
		<TD><TEXTAREA rows=5 id=txtGroup3 name=txtGroup3 style="width:300px"><%= sGroup3%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Group 4</TH>
		<TD><TEXTAREA rows=5 id=txtGroup4 name=txtGroup4 style="width:300px"><%= sGroup4%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Group 5</TH>
		<TD><TEXTAREA rows=5 id=txtGroup5 name=txtGroup5 style="width:300px"><%= sGroup5%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Manufacturing Notes</TH>
		<TD><TEXTAREA rows=5 id=txtManufacturingNotes name=txtManufacturingNotes style="width:300px"><%= sManufacturingNotes%></TEXTAREA></TD>
	</TR>
	<TR>
		<TH>Weight</TH>
		<TD><div id=WeightInput style="display:none"><INPUT type="text" id=txtWeight name=txtWeight maxlength="9" value="<%= sWeight%>"></div><div id=WeightTxt><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sWeight)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><a href="javascript:EditWeight();">Edit</a></td><%End If%></tr></table></div></TD>
	</TR>
	<TR>
		<TH>Global Series Config<br />Planfor End of<br />Manufacturing<br />(PE) Date</TH>
		<TD><div id=GSEndDtInput style="display:none"><INPUT type="text" id=txtGSEndDt name=txtGSEndDt value="<%= sGSEndDt%>"></div><div id=GSEndDtTxt><table width="100%" border=0><tr><td style="border:none"><%= PrepForWeb(sGSEndDt)%></td><%If m_EditModeOn Then%><td align=Right style="border:none"><a href="javascript:EditGSEndDt();">Edit</a></td><%End If%></tr></table></div></TD>
	</TR>
	<TR>
		<TH>IDS-SKUS</TH>
		<TD><INPUT type="checkbox" id="chkIdsSkus" name="chkIdsSkus" <%IF bIdsSkus Then%>CHECKED<%End If%>></TD>
	</TR>
	<TR>
		<TH>IDS-CTO</TH>
		<TD><INPUT type="checkbox" id="chkIdsCto" name="chkIdsCto" <%IF bIdsCto Then%>CHECKED<%End If%>></TD>
	</TR>
	<TR>
		<TH>RCTO-SKUS</TH>
		<TD><INPUT type="checkbox" id="chkRctoSkus" name="chkRctoSkus" <%IF bRctoSkus Then%>CHECKED<%End If%>></TD>
	</TR>
	<TR>
		<TH>RCTO-CTO</TH>
		<TD><INPUT type="checkbox" id="chkRctoCto" name="chkRctoCto" <%IF bRctoCto Then%>CHECKED<%End If%>></TD>
	</TR>
    <% If (sProdVersionBSAMFlag = "True") Then %>
	    <TR>
	        <TH>BSAM SKUS</TH>
		    <TD><INPUT type="checkbox" id="chkBSAMSkus" name="chkBSAMSkus" <%IF bBSAMSkus Then%>CHECKED<%End If%>></TD>
	    </TR>
	    <TR>
	        <TH>BSAM -B parts</TH>
		    <TD><INPUT type="checkbox" id="chkBSAMBparts" name="chkBSAMBparts" <%IF bBSAMBparts Then%>CHECKED<%End If%>></TD>
	    </TR>
	<% End If %>
	<TR>
		<TH>UPC</TH>
		<TD><%= sUpc%></TD>
	</TR>
<%	If LCase(sMode) <> "add" Then %>
	<TR>
		<TH>Reason for Change:</TH>
		<TD><TEXTAREA rows="2" id="txtChangeReason" name="txtChangeReason" style="width:300px"><%= sChangeNote%></TEXTAREA></TD></TR>
<%	End If %>
	<TR>
		<TH>Sorting Weight</TH>
		<TD><INPUT type="text" id="txtSortOrder" name="txtSortOrder" size="3" maxlength="3" value="<%= sSortOrder%>"></TD></TR>
	<TR>
		<TH>Show Change on SCM</TH>
		<TD><INPUT type="checkbox" id="chkShowOnScm" name="chkShowOnScm" checked="CHECKED"></TD>
	</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
