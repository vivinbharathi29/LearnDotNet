<%@  language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" -->
<!-- #include file = "../includes/lib_debug.inc" -->
<%
'printrequest
'response.end
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
Dim iAliasID            : iAliasID = ""
Dim sCategoryOpt		: sCategoryOpt = ""
Dim	iCategoryOpt		: iCategoryOpt = ""
Dim sCategoryAbbr       : sCategoryAbbr = ""
Dim sPlatformName       : sPlatformName = ""
Dim sSCMCat			    : sSCMCat = ""
Dim sGpgDesc			: sGpgDesc = ""
Dim sSortOrder			: sSortOrder = ""
Dim sMarketingDesc		: sMarketingDesc = ""
Dim sMarketingDescPMG	: sMarketingDescPMG = ""
Dim sManufacturingNotes	: sManufacturingNotes = ""
Dim sProgramVersion		: sProgramVersion = GetProductVersion(Request("PVID"))
Dim sPRLOffConstraints  : sPRLOffConstraints = ""
Dim sConfigRules		: sConfigRules = ""
Dim sRulesSyntax		: sRulesSyntax = ""
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
Dim iBrandID			: iBrandID = Request("BID")
Dim sCplBlindDt			: sCplBlindDt = ""
Dim sGeneralAvailDt     : sGeneralAvailDt = ""
Dim sRasDiscDt			: sRasDiscDt = Request.QueryString("EMDate")
Dim sPAADDate			: sPAADDate = ""    
Dim sWeight				: sWeight = ""
Dim sChangeNote         : sChangeNote= ""
Dim sGSEndDt            : sGSEndDt = ""
Dim sRTPDt		        : sRTPDt = Request.QueryString("RTPDate")
Dim sPhWebInstruction	: sPhWebInstruction = ""
'Dim sPDMFeedback    	: sPDMFeedback = ""
Dim sSDFFlag            : sSDFFlag = "False"
Dim sAvId       		: sAvId = ""
Dim sGroup1     		: sGroup1 = ""
Dim sGroup2		        : sGroup2 = ""
Dim sGroup3		        : sGroup3 = ""
Dim sGroup4		        : sGroup4 = ""
Dim sGroup5		        : sGroup5 = ""
Dim sGroup6		        : sGroup6 = ""
Dim sGroup7		        : sGroup7 = ""
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
dim sSCMParent              'New dimension to get the parent of the AV
dim	bOriginatedByDCR
dim	iDCRNo
Dim sProdVersionBSAMFlag
Dim strProductLines     : strProductLines = ""
Dim strProductLineID	: strProductLineID = ""
Dim strProductLineName	: strProductLineName = ""
Dim sIsDesktop  
Dim strFeatureName      : strFeatureName = ""
Dim iFeatureID          : iFeatureID = "" 
Dim iParentID           : iParentID = "0"
Dim Platform            : Platform = ""
Dim sCreated 
Dim sCreatedBy
Dim sUpdated
Dim sUpdatedBy
Dim sConfigCode         : sConfigCode = ""
'Dim sGeneralAvailSysUpdate : sGeneralAvailSysUpdate  = ""
Dim sEOM            : sEOM = "" 
Dim sRTP            : sRTP = ""   

Dim sSCMCategoriesProductLines : sSCMCategoriesProductLines = ""
Dim sProductProductLine : sProductProductLine = ""
Dim sSCMCategoriesInheritProductLine    : sSCMCategoriesInheritProductLine = ""

Dim sProductLinesAll     : sProductLinesAll = ""
Dim strDemandRegionID    : strDemandRegionID = ""
Dim bParentAV
Dim bBaseParent          : bBaseParent = false

Dim sFeatureID		    : sFeatureID = ""
Dim sFeatureName		: sFeatureName = ""
Dim bSharedAV			: bSharedAV = false
Dim sComments           : sComments =""
Dim strSCMOldCategory   : strSCMOldCategory = Request("SCMCat")
Dim iFromTodayPage      : iFromTodayPage = Request("FromTodayPage")
Dim bFeaturechanged     : bFeaturechanged =""
Dim bDescriptionChanged : bDescriptionChanged ="" 'need to know if descriptions are changed after save
Dim sBaseAvNo           : sBaseAvNo = ""

Dim scmlist             : scmlist = ""

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

Dim sOriginalAvNo               : sOriginalAvNo = ""

Dim m_ProductVersionID	: m_ProductVersionID = Request("PVID")
Dim m_IsSysAdmin
Dim m_IsProgramCoordinator
Dim m_IsConfigurationManager
Dim m_EditModeOn
Dim m_UserFullName
Dim m_UserID

'------------------------- COMBINED AV FUNCTIONALITY------(santodip)------------------------- 
Dim FID	: FID = Request.QueryString("FID")
Dim FName : FName = Request.QueryString("FName")
Dim FRequiresRoot : FRequiresRoot = Request.QueryString("FRequiresRoot")
Dim FComponentLinkage : FComponentLinkage = Request.QueryString("FComponentLinkage")
Dim FComponentRootID : FComponentRootID = Request.QueryString("FComponentRootID")
Dim FGPGDescription	: FGPGDescription = Request.QueryString("FGPGDescription")
Dim FMarketingDescriptionPMG : FMarketingDescriptionPMG = Request.QueryString("FMarketingDescriptionPMG")
Dim FMarketingDescription : FMarketingDescription = Request.QueryString("FMarketingDescription")
Dim FSCMCategoryID_singlefeature : FSCMCategoryID_singlefeature = Request.QueryString("FSCMCategoryID_singlefeature")

Dim m_CanEditDates : m_CanEditDates = Request.QueryString("IsMarketingUser")
    
'task 16227
Dim sProductRelease		: sProductRelease = Request.QueryString("ReleaseName")
Dim sProductReleaseIDs : sProductReleaseIDs = Request.QueryString("Release")
Dim strLoaded : strLoaded = ""

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
    m_UserID = Security.CurrentUserId()
	
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
	Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
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
    if(len(trim(value)) > 0) then
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
    else
		GetBoolValue = false
    end if


End Function

Sub Main()
'
'TODO: Get AvDetail Data
'
	 Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sIsDesktop = rs("IsDesktop") 
	sProductProductLine=rs("ProductLineId") 
	rs.Close

    Set cmd = dw.CreateCommAndSP(cn, "usp_GetBrands4Product")
	dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, Request("PVID")
	dw.CreateParameter cmd, "@SelectedOnly", adTinyInt, adParamInput, 1, "1"
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
    rs.Sort = "CombinedName, Name"
    dim m_BrandName, m_BrandID, sBrandDisplayed
    sBrandDisplayed =""        	
	Do Until rs.EOF

        if rs("CombinedName") <> "" and not isnull(rs("CombinedName")) then	'see if there is a combined name first
			m_BrandName = rs("CombinedName")
			m_BrandID = rs("combinedProductBrandId")			
		else
			m_BrandName = rs("Name")	'no combined name so display the name from the Brand table
            m_BrandID =  rs("ProductBrandID") 
		end if
        if sBrandDisplayed = "" or sBrandDisplayed <> m_BrandName then 'have to handle if it is a combined Brand name because we don't want to display the same name multiple times
			sCbxBrand = sCbxBrand & "<INPUT type=checkbox id=chkBrand name=chkBrand value=" & m_BrandID 
		    If Trim(m_BrandID) = Trim(Request("BID")) Then
			    sCbxBrand = sCbxBrand & " CHECKED "
		    End If
		    sCbxBrand = sCbxBrand & ">" & m_BrandName & "<BR>"
		end if
		sBrandDisplayed = m_BrandName
		
		rs.MoveNext
	Loop
	
	sCbxBrand = Left(sCbxBrand, Len(sCbxBrand) - 4)
	sSDFFlag = "New"

	If Request("AVID") <> "" Then	'Get the values for the request AV

		Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail_Pulsar")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
		dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
		dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
		Set rs = dw.ExecuteCommandReturnRS(cmd)

        iParentID = rs("ParentId")
        sFeatureID = rs("FeatureID")
        sFeatureName = rs("FeatureName")
        iAliasID = rs("AliasID")
		sAvNo = rs("AvNo")
        sOriginalAvNo = rs("AvNo")
		iCategoryOpt = rs("SCMCategoryID")
        sCategoryAbbr = rs("SCMCategoryAbbr")
        sPlatformName = rs("PlatformName")
		sSCMCat = rs("Name")
        sSCMParent = rs("HasParent")                'Returns scm category id if it is null it should be 0
		sGpgDesc = rs("GPGDescription")
		sMarketingDesc = rs("MarketingDescription")
		sMarketingDescPMG = rs("MarketingDescriptionPMG")
        sPRLOffConstraints = rs("PRLOfferingConstraints")
        'sPRLOffConstraints = Trim(Replace(sPRLOffConstraints,"\n", vbCrLf))
        if InStr(1,sPRLOffConstraints,"\n") = 1 then
            sPRLOffConstraints = Mid(sPRLOffConstraints,3)
        end if
        sPRLOffConstraints = Replace(sPRLOffConstraints,"\n", "<br /><br />")
		sConfigRules = rs("ConfigRules")
		sManufacturingNotes = rs("ManufacturingNotes")
		bIdsSkus = GetBoolValue(rs("IdsSkus_YN"))
		bIdsCto = GetBoolValue(rs("IdsCto_YN"))
		bRctoSkus = GetBoolValue(rs("RctoSkus_YN"))
		bRctoCto = GetBoolValue(rs("RctoCto_YN"))
		bBSAMSkus = GetBoolValue(rs("BSAMSkus_YN"))
		bBSAMBparts = GetBoolValue(rs("BSAMBparts_YN"))
        strProductLineID = rs("ProductLineID")
        strProductLineName = rs("ProductLine")
        if LCase(rs("bSharedAV")) = "true" then
            bSharedAV = true
        end if
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
		sGroup6 = rs("Group6")
		sGroup7 = rs("Group7")
		iDeliverableRootID = rs("DeliverableRootID")
		iViaAvCreate = rs("ViaAvCreate") & ""
		bOriginatedByDCR = rs("OriginatedByDCR")
		iDCRNo = rs("DCRNo") & ""
		
		strAVName3 = rs("GPGDescription")
		strAVName5 = rs("MarketingDescription")
		strAVName7 = rs("MarketingDescriptionPMG")
		strExistingNameElements = rs("NameElements")		
		
        sPAADDate=rs("PHWebDate")
        
        sComments =  rs("Comments")
		sRulesSyntax = rs("RulesSyntax")
		if rs("ParentID") > 0 then
			bParentAV = False
		else
			bParentAV = True
		end if
		strDemandRegionID = rs("DemandRegionID")
		if rs("BaseParent") = 1 then
			bBaseParent = True
		end if

        sCreated = rs("Created")
        sCreatedBy = rs("CreatedBy")
        sUpdated = rs("Updated")
        sUpdatedBy = rs("UpdatedBy")
        'sGeneralAvailSysUpdate = rs("GeneralAvailSysUpdate")
        sConfigCode = rs("ConfigCode")
		rs.Close 
    
        'Get the Releases for the AV
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_ProductRelease_OnAV")
	    dw.CreateParameter cmd, "@PBID", adInteger, adParamInput, 8, Trim(Request("BID")) 
        dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, Trim(Request("AVID"))
	    Set rs = dw.ExecuteCommandReturnRS(cmd)
        sProductRelease = rs("PR")
        sProductReleaseIDs = rs("PRIDs")
        rs.Close 
      
	End If   

    if sProductReleaseIDs <> "" then
        'Get End of manufaturing date
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_EOMDate")
		    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
            dw.CreateParameter cmd, "@p_ReleaseID", adVarchar, adParamInput, 200, sProductReleaseIDs
		    Set rs = dw.ExecuteCommandReturnRS(cmd)  
            if not rs.BOF and not rs.EOF then
                if not IsNull(rs("EOM")) then
                    sEOM = rs("EOM")
                end if
            end if
            rs.Close 
        
        'Get RTP date
        Set cmd = dw.CreateCommandSP(cn, "usp_Get_RTPDate")
		    dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
            dw.CreateParameter cmd, "@p_ReleaseID", adVarchar, adParamInput, 200, sProductReleaseIDs
		    Set rs = dw.ExecuteCommandReturnRS(cmd)
            if not rs.BOF and not rs.EOF then
                if not IsNull(rs("RTP")) then
                    sRTP = rs("RTP")
                end if
            end if
            rs.Close 
   end if

	If sMode = "clone" Then
		sAvNo = ""
		'sMode = "add"
	End If

    if iParentID > 0 then 
        'Get Base Av PartNo if ParentID is not 0     
        
        Set rs = server.createobject("ADODB.Recordset")
        Set cmd = server.createobject("ADODB.Command")        
        
        cmd.ActiveConnection = cn
        cmd.CommandText = "Select isnull(AvNo,'') as AvNo From AvDetail Where AvDetailID = ?"
        cmd.CommandType = adCmdText
        'cmd.CommandTimeout = 900 
            
        cmd.Parameters.Append cmd.CreateParameter("@AvDetaiID", 3, 1, , iParentID)

        ' Execute the query for readonly
        rs.CursorLocation = adUseClient
        rs.Open cmd, , adOpenForwardOnly, adLockReadOnly  
       
        sBaseAvNo = rs("AvNo")                                  
        rs.Close
    end if

    'use the SCM Category list - task 9900
	Set cmd = dw.CreateCommandSP(cn, "usp_SCM_GetSCMCategories") 
	dw.CreateParameter cmd, "@p_CurrentCatID", adInteger, adParamInput, 8, iCategoryOpt
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	Do Until rs.EOF
		    strNameFormats = strNameFormats & "<option value=""" & rs("SCMCategoryID") &   """></option>"
		    sCategoryOpt = sCategoryOpt & "<OPTION Value='" & rs("SCMCategoryID") & "'"
		    If iCategoryOpt = rs("SCMCategoryID") Then
			    sCategoryOpt = sCategoryOpt & " SELECTED "
		    End If
		    sCategoryOpt = sCategoryOpt & ">" & rs("Name") & "</OPTION>" & VbCrLf
		rs.MoveNext
	Loop
     rs.Close

    ' read SCM Categories asociated with ProductLines 
    Set cmd = dw.CreateCommandSP(cn, "usp_SCM_GetSCMCategories_ProductLines") 
	Set rs = dw.ExecuteCommandReturnRS(cmd)

    Do Until rs.EOF
        If sSCMCategoriesProductLines="" then
            sSCMCategoriesProductLines= rs("SCMCategoryID") & "#" & rs("AssignedProductLine") & "#" & rs("NoOfProductLine") 
        else
           sSCMCategoriesProductLines= sSCMCategoriesProductLines & "," & rs("SCMCategoryID") & "#" & rs("AssignedProductLine") & "#" & rs("NoOfProductLine") 
        End If
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
	
    ' ADD PRODUCT LINE TEXT FIELD TO FORM
    rs.Open "spGetProductLinesAll",cn,adOpenForwardOnly
	
    If Request("AVID") <> "" Then	'Get the values for the request AV
    do while not rs.EOF
	    if trim(strProductLineID) = trim(rs("ID") & "" ) then
		    strProductLines = strProductLines & "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & " - " & rs("Description") & "</OPTION>"
	    else
		    strProductLines = strProductLines & "<OPTION value=" & rs("ID") & ">" & rs("Name") & " - " & rs("Description") & "</OPTION>"
	    end if
        if sProductLinesAll="" then
            sProductLinesAll =   rs("ID")  & "#"  & rs("Name") & " - " & rs("Description") 
        else
            sProductLinesAll = sProductLinesAll  & "," & rs("ID")  & "#"  & rs("Name") & " - " & rs("Description") 
        end if
	    rs.MoveNext
	loop
    else 'new av  - select the default product line - strAVDefaultProductLine
        do while not rs.EOF
		        strProductLines = strProductLines & "<OPTION value=" & rs("ID") & ">" & rs("Name") & " - " & rs("Description") & "</OPTION>"
            if sProductLinesAll="" then
                sProductLinesAll =   rs("ID")  & "#"  & rs("Name") & " - " & rs("Description") 
            else
                sProductLinesAll = sProductLinesAll  & "," & rs("ID")  & "#"  & rs("Name") & " - " & rs("Description") 
	        end if
	    rs.MoveNext
	    loop        
    end if
    
    
    
	rs.Close
  ' ADD PRODUCT LINE TEXT FIELD TO FORM

    If bSharedAV = true Then
        'Response.Write("<script language=VBScript>MsgBox """ + scmlist + """</script>") 
         Dim rsSCM
         Set cmd = dw.CreateCommAndSP(cn, "usp_GetSCMNames")
         dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, sAvNo
         Set rsSCM = dw.ExecuteCommAndReturnRS(cmd)

         If rsSCM.EOF Then
              scmlist = ""
         Else
             Do While Not rsSCM.EOF
             'Response.Write("<script language=VBScript>MsgBox """ + rsSCM.Fields("DOTSName").value + """</script>") 
             scmlist = rsSCM.Fields("DOTSName").value + ","  + scmlist   
             rsSCM.MoveNext()
             Loop
             'Response.Write("<script language=VBScript>MsgBox """ + scmlist + """</script>") 
         End If
        rsSCM.Close  

   End If

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

    Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion_Pulsar")
	dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	sIsDesktop = rs("IsDesktop") 
	


	Dim returnValue
	Dim iAvId
	Dim errDesc

	cn.BeginTrans
	
	Dim AvPcActionTypeID

    'Save AvDetail data
	If LCase(sMode) = "add" Or LCase(sMode) = "clone" Then
		Set cmd = dw.CreateCommandSP(cn, "usp_InsertAvDetail_Pulsar")
		cmd.NamedParameters = True
        'dw.CreateParameter cmd, "@p_FeatureID", adVarchar, adParamInput, 50, Request.Form("txtfeatureIDDesc") 
        dw.CreateParameter cmd, "@p_FeatureID", adInteger, adParamInput, 8, Request.Form("txtfeatureIDDesc") 
        dw.CreateParameter cmd, "@p_DeliverableRootID", adVarchar, adParamInput, 50, Request.Form("DelRootID")
		dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
		dw.CreateParameter cmd, "@p_SCMCategoryID", adInteger, adParamInput, 8, Request.Form("cboCategory")
		dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtAvGpgDescription")
		dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
		dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 100, Request.Form("txtMarketingDescPMG")
		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkShowOnScm"))
		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
		dw.CreateParameter cmd, "@p_NameElements", adVarchar, adParamInput, 500, Request.Form("strNameElements")
		dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamOutput, 8, ""
		'dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
	    dw.CreateParameter cmd, "@p_ProductLineID", adInteger, adParamInput, 8, Request.Form("cboProductLine")
        dw.CreateParameter cmd, "@p_SharedAV", adInteger, adParamInput, 8, Request.Form("hdnSharedValue")
        dw.CreateParameter cmd, "@p_BID", adInteger, adParamInput, 8, Request.Form("BID")
        '-----------------additional parameter to get RTP and EOM dates------------------------
        dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
        dw.CreateParameter cmd, "@p_Releases", adVarchar, adParamInput, 2000, Request.Form("txtReleaseList") 'task 16227   
       dw.CreateParameter cmd, "@p_RTPDate", adDate, adParamInput, 50, Request.Form("txtRTPDate")
       dw.CreateParameter cmd, "@p_EMDate", adDate, adParamInput, 50, Request.Form("txtMarketingDiscDate")
        dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 50, Request.Form("txtBlindDate1")
        dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50, Request.Form("txtGeneralAvailDt1")
       dw.CreateParameter cmd, "@PhwebDate", adDate, adParamInput, 50,  Request.Form("txtPAADDate1")
       dw.CreateParameter cmd, "@p_ErrDesc", adVarchar, adParamOutput, 500, 0

        returnValue = dw.ExecuteNonQuery(cmd)
        errDesc = cmd("@p_ErrDesc")		
		iAvId = cmd("@p_AvDetailID")
	ElseIf Request("AVID") <> "" Then
		iAvId = Request("AVID")
		'if featureID changed,create a new avdetail record 
		'and set the original one with empty avno and obsolete the original one used in scms; then the
        'following update will be applied to the new avdetailID
        'add new parameter for scm category to usp_SCM_ChangeFeatureForAV so we can set all av under the category that user selected or its subcategory if any
        if len(Request.Form("txtfeatureIDDesc")) > 0 and len(Request.Form("txtPrevFeatureID")) > 0  and Request.Form("txtfeatureIDDesc") <> Request.Form("txtPrevFeatureID") then
    
            Set cmd = dw.CreateCommandSP(cn, "usp_SCM_ChangeFeatureForAV")
		    cmd.NamedParameters = True
            dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
            dw.CreateParameter cmd, "@p_FeatureID", adInteger, adParamInput, 8, Request.Form("txtfeatureIDDesc") 
            dw.CreateParameter cmd, "@p_SCMCategoryID", adVarchar, adParamInput, 8, Request.Form("cboCategory")
            dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName		    		   
            returnValue = dw.ExecuteNonQuery(cmd)
		               
            bFeaturechanged="1"
            
            If returnValue = -1 Then                  
                'Abort Transaction  
                'sometime share av will cause error if it share by two different production that one is desktop and the other is notebook
	            Response.Write("<script language=javascript> alert('There is error and nothing is saved.  The AV could belong to multiple Products that is not the same type.'); window.location.href = 'avDetail.asp?Mode=" + Request("MODE") + "&PVID=" + Request("PVID") + "&BID=" + Request("BID") + "&AVID=" + Request("hidAVID") + "';</script>")
                cn.RollbackTrans()
                response.End
            End If

        end if 
		       
		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_Pulsar")
		dw.CreateParameter cmd, "@p_AvDetailID", adVarchar, adParamInput, 8, iAvId
		dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, Request.Form("AvNo")
		dw.CreateParameter cmd, "@p_SCMCategoryID", adInteger, adParamInput, 8, Request.Form("cboCategory")
		dw.CreateParameter cmd, "@p_GPGDescription", adVarchar, adParamInput, 50, Request.Form("txtAvGpgDescription")
		dw.CreateParameter cmd, "@p_MarketingDescription", adVarchar, adParamInput, 40, Request.Form("txtMarketingDesc")
		dw.CreateParameter cmd, "@p_MarketingDescriptionPMG", adVarchar, adParamInput, 100, Request.Form("txtMarketingDescPMG")
		
        '----------------check if SA value is from edit or from calculation-----------------------------------
        if Request.Form("hdnflag1") = "E" then
            'Response.Write("<script language=VBScript>MsgBox """ + "if loop cplblinddate" + """</script>")
            dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 50, Request.Form("txtBlindDate1")          
        else
            'Response.Write("<script language=VBScript>MsgBox """ + "ELSE if loop sa date" + """</script>")
            dw.CreateParameter cmd, "@p_CPLBlindDt", adDate, adParamInput, 8, Request.Form("hidCplBlindDt")
        end if
        
        if Request.Form("hdnCanEditDates") = "True" then
            dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 50, Request.Form("txtMarketingDiscDate")        
    	else
            dw.CreateParameter cmd, "@p_RASDiscontinueDt", adDate, adParamInput, 50, Request.Form("hidRasDiscDt")
        end if

		dw.CreateParameter cmd, "@p_UPC", adVarchar, adParamInput, 12, Request.Form("hidUpc")
		dw.CreateParameter cmd, "@p_ChangeNotes", adVarchar, adParamInput, 500, Request.Form("txtChangeReason")
		dw.CreateParameter cmd, "@p_ShowOnScm", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkShowOnScm"))
		dw.CreateParameter cmd, "@p_Last_Upd_User", adVarchar, adParamInput, 50, m_UserFullName
        
		
        if Request.Form("hdnCanEditDates") = "True" then
            dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("txtRTPDate")
        else
            dw.CreateParameter cmd, "@p_RTPDt", adDate, adParamInput, 50, Request.Form("hidRTPDt")
        end if
        
        '----------------check if GA value is from edit or from calculation----------------------------------- 
        if Request.Form("hdnflag2") = "E" then
           'Response.Write("<script language=VBScript>MsgBox """ + "if loop GA date" + """</script>")
            dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50, Request.Form("txtGeneralAvailDt1")       
        else
            dw.CreateParameter cmd, "@p_GeneralAvailDt", adDate, adParamInput, 50,  Request.Form("hidGeneralAvailDt") 
        end if 

        dw.CreateParameter cmd, "@p_NameElements", adVarchar, adParamInput, 500, Request.Form("strNameElements")
		'dw.CreateParameter cmd, "@p_AvPcActionTypeID", adInteger, adParamInput, 8, AvPcActionTypeID
		dw.CreateParameter cmd, "@p_weight", adInteger, adParamInput, 8, Request.Form("txtWeight")
        If Request.Form("cboDeliverables") > 0 Then
		   dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request.Form("cboDeliverables")
        else
            dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 8, Request.Form("DelRootID")
        End If  
        dw.CreateParameter cmd, "@p_ProductLineID", adInteger, adParamInput, 8, Request.Form("cboProductLine")
       
        '----------------check if PAAD value is from edit or from calculation----------------------------------- 
        if Request.Form("hdnflag") = "E" then
            'Response.Write("<script language=VBScript>MsgBox """ + "if loop" + """</script>")
            dw.CreateParameter cmd, "@PhwebDate", adDate, adParamInput, 50,  Request.Form("txtPAADDate1")       
        else
            dw.CreateParameter cmd, "@PhwebDate", adDate, adParamInput, 50,  Request.Form("txtPAADDate")
        end if 

        dw.CreateParameter cmd, "@p_FeatureID", adVarchar, adParamInput, 50, Request.Form("txtfeatureIDDesc")
        dw.CreateParameter cmd, "@p_BUAvailList", adVarchar, adParamInput, 2000, Request.Form("BUAvailList")
        dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Request.Form("BID")
        dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, m_UserID
        dw.CreateParameter cmd, "@p_ErrDesc", adVarchar, adParamOutput, 500, 0

        returnValue = dw.ExecuteNonQuery(cmd)
        errDesc = cmd("@p_ErrDesc")

        'compare description, if is changed we set the bit to refresh the whole grid
        if  Request.Form("txtAvGpgDescription") <> Request.Form("hidOldGpgDescription") or Request.Form("txtMarketingDesc") <> Request.Form("hidOldMS40Description") or Request.Form("txtMarketingDescPMG") <> Request.Form("hidOldML100Description") then
            bDescriptionChanged="1"
        end if

    End If
	
    If errDesc <> "" Then
        Response.Write("<script language=javascript> alert('" & errDesc & "'); </script>")        
    else   
        ' Link AV to Product_Brands
	    If Request("AVID") <> "" Then
		    saBrands = Split(Request("BID"), ",")
	    End If
	
	    For i = LBound(saBrands) To UBound(saBrands)  
	        If Not IsNull(iAvId) Then
		        ' Add Records to AvDetail_ProductBrand
		        Set cmd = dw.CreateCommandSP(cn, "usp_UpdateAvDetail_ProductBrand_Pulsar")
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
		      '  dw.CreateParameter cmd, "@p_PhWebInstruction", adVarchar, adParamInput, 500, Request.Form("txtPhWebInstruction")
                'this field is not needed any more in pulsar
                dw.CreateParameter cmd, "@p_PhWebInstruction", adVarchar, adParamInput, 500, ""
		        dw.CreateParameter cmd, "@p_AvId", adVarchar, adParamInput, 2000, Request.Form("txtAvId")
		        dw.CreateParameter cmd, "@p_Group1", adVarchar, adParamInput, 2000, Request.Form("txtGroup1")
		        dw.CreateParameter cmd, "@p_Group2", adVarchar, adParamInput, 2000, Request.Form("txtGroup2")
		        dw.CreateParameter cmd, "@p_Group3", adVarchar, adParamInput, 2000, Request.Form("txtGroup3")
		        dw.CreateParameter cmd, "@p_Group4", adVarchar, adParamInput, 2000, Request.Form("txtGroup4")

                'If sIsDesktop=False Then
		            dw.CreateParameter cmd, "@p_Group5", adVarchar, adParamInput, 2000, Request.Form("txtGroup5")
                'Else
                '    dw.CreateParameter cmd, "@p_Group5", adVarchar, adParamInput, 2000, ""
                'End If

		        dw.CreateParameter cmd, "@p_OriginatedByDCR", adBoolean, adParamInput, 8, GetCbxBlnValue(Request.Form("chkDCR"))
		        dw.CreateParameter cmd, "@p_DCRNo", adInteger, adParamInput, 8, Request.Form("txtDCRNo")
                dw.CreateParameter cmd, "@p_Comments", adVarchar, adParamInput, 500, Request.Form("txtComments") 
                If (sProdVersionBSAMFlag = "True") Then
		            dw.CreateParameter cmd, "@p_BSAMSkus_YN", adChar, adParamInput, 1, "" 'GetCbxValue(Request.Form("chkBSAMSkus"))
		            dw.CreateParameter cmd, "@p_BSAMBparts_YN", adChar, adParamInput, 1, GetCbxValue(Request.Form("chkBSAMBparts"))
				else
		            dw.CreateParameter cmd, "@p_BSAMSkus_YN", adChar, adParamInput, 1, ""
		            dw.CreateParameter cmd, "@p_BSAMBparts_YN", adChar, adParamInput, 1, ""
		        End If
				
				dw.CreateParameter cmd, "@p_RulesSyntax", adVarchar, adParamInput, 512, Request.Form("txtRulesSyntax") 
				
				dw.CreateParameter cmd, "@p_DemandRegionID", adInteger, adParamInput, 4, NULL
				
                dw.CreateParameter cmd, "@p_Group6", adVarchar, adParamInput, 2000, Request.Form("txtGroup6")
                dw.CreateParameter cmd, "@p_Group7", adVarchar, adParamInput, 2000, Request.Form("txtGroup7")
                ' Pass releases to usp_UpdateAvDetail_ProductBrand_Pulsar sp instead of usp_UpdateAvDetail_Pulsar
                dw.CreateParameter cmd, "@p_Releases", adVarchar, adParamInput, 2000, Request.Form("txtReleaseList") 'task 16227

		        returnValue = dw.ExecuteNonQuery(cmd)
                If returnValue = -1 Then
	                Response.Write("<script language=javascript> alert('Base unit AV with the same Feature already exists in the SCM.  Please select a different Feature.'); </script>")
                end if

		    End If
	    Next
	
	    sFunction = "close"
	    cn.CommitTrans
	

    End if

 

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
    Call Main()
Else
	Call Main()
End If

%>
<html>
<head>
    <title></title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../style/excalibur.css">
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">


        $(document).ready(function () {

            if (window.parent.frames["LowerWindow"].frmButtons.cmdCancel) {
                window.parent.frames["LowerWindow"].frmButtons.cmdCancel.disabled = false;
            }

            if ($("#hidMode").val() == 'edit' && $("#hdnSharedValue1").val().toLowerCase() != 'true')
            {
                if ($("#SA1").html() == '&nbsp;' && $("#GA1").html() == '&nbsp;' && $("#PAAD1").html() == '&nbsp;')
                {
                    //alert("edit mode - blank dates");
                    calcDates();
                    document.getElementById('hdnAlgorithm').value = 'Y';
                }
            }
                
            if ($("#hidMode").val() == 'add' || $("#hidMode").val() == 'clone')    
            {
                document.getElementById('hdnAlgorithm').value = 'Y';
                calcDates();
            }
            $("#txtDCRNo").off('keyup').on('keyup',function(e)
            {
                if($(this).val().indexOf('0')==0)
                $(this).val('');
            });


            $("#txtDCRNo").off('keydown').on('keydown',function(e)
            {

             if (!e) e = window.event;
                   var key = e.keyCode ? e.keyCode : e.which;

                //allow delete and backspace
                if(e.keyCode==8|| e.keyCode==46)
                return;
                //do not allow control c, control v
                if(e.ctrlKey)
                {
                e.preventDefault();
                return;
                }
            var begin1=47, begin2=96;
            var selectedlength= document.getElementById("txtDCRNo").selectionEnd - document.getElementById
            ("txtDCRNo").selectionStart;
            //if part of the text is selected and tried to enter 0 at the beginning
            if($(this).val().length>0 && document.getElementById("txtDCRNo").selectionStart==0)
            {
            if(key==48 || key==96)
            {
            e.preventDefault();
            return;
            }
            }
            //check if this is first character or if the wholetext is selected and replaced with 0
            if($(this).val().length==0)
            {
            begin1=48;
            begin2=97;
            }
            if(selectedlength>0 && $(this).val().length== selectedlength)
            {
            begin1=48;
            begin2=97;
            }
                   // Was key that was pressed a numeric character (0-9) or backspace?
                   if (( key > begin1 && key < 58 ) || key == 8 ||(e.keyCode >= begin2 && e.keyCode <= 105)) 
                           return; // if so, do nothing
                   else // otherwise, discard character                
                                   e.preventDefault(); 

 
            });            

            $("#txtWeight").off('keydown').on('keydown', function (e) {

                if (!e) e = window.event;
                var key = e.keyCode ? e.keyCode : e.which;
                if (e.shiftKey == 1) {
                    e.preventDefault();
                    return;
                }
                //allow delete and backspace
                if (e.keyCode == 8 || e.keyCode == 46)
                    return;
                //do not allow control c, control v
                if (e.ctrlKey) {
                    e.preventDefault();
                    return;
                }


                var decimalExists=$("#txtWeight").val().indexOf('.') >=0;

                // Was key that was pressed a numeric character (0-9) or backspace?
                if ((key >= 48 && key <= 57) || key == 8 || (key >= 96 && key <= 106) || ((key==110 || key==190) && !decimalExists))
                {
                    //if there is a decimal then allow only 2 values
                    if(decimalExists)
                    {
                        if($("#txtWeight").val().length - ($("#txtWeight").val().indexOf('.')) >2)
                            e.preventDefault();
                        return;
                    }
                    else{
                        return; // if so, do nothing
                    }
                }
                else // otherwise, discard character 
                    e.preventDefault();
           
            });

            $("#txtRTPDate").focusout(function () {
                if (checkRTPDate()) {
                    if (document.getElementById('hdnAlgorithm').value == 'Y') {
                        calcDates();
                    }
                }
            });

            $("#txtMarketingDiscDate").focusout(function () {
                checkEMDate();
            });


            if ($("#txtPAADDate1").val() == "" && $("#txtBlindDate1").val() == "" && $("#txtGeneralAvailDt1").val() == "") {
                if ($("#hdnSharedValue1").val().toLowerCase() != 'true') {
                    $("#txtRTPDate").val($("#hidRTP").val());
                    $("#txtMarketingDiscDate").val($("#hidEOM").val());
                    document.getElementById('hdnAlgorithm_blankdates').value = 'Y';
                    calcDates_whendatesareblank();
                }
            }
       
        });

        function calcDates() {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (isDate($("#txtRTPDate").val())) {

                var RTPDate = new Date($("#txtRTPDate").val());
                
                var GeneralAvailDt = RTPDate;
                   
                var monday = getMonday(RTPDate);
                var firstDay = new Date(GeneralAvailDt.getFullYear(), GeneralAvailDt.getMonth(), 1);
                var PAADDate = new Date((monday.getMonth() + 1) + '/' + monday.getDate() + '/' + monday.getFullYear());

                var BlindDate;
                BlindDate = new Date(PAADDate.getFullYear(), (PAADDate.getMonth() - 1), 1);

                if (BlindDate < Today)
                    BlindDate = new Date(Today.getFullYear(), Today.getMonth(), Today.getDate() + 7);

                if ($("#hidCplBlindDt").val() != "") {
                    var existingsadate = new Date($("#hidCplBlindDt").val());                  
                    if (existingsadate < Today)
                        BlindDate = existingsadate;                    
                }               

                $("#mktgBlindDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td></tr></table>');
                $("#hidCplBlindDt").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtPAADDate").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                $("#mktgPAADDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear() + '</span></td><td align=Right style="border:none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td></tr></table>');
                $("#mktgGeneralAvailDtText").html('<table width="100%" border="0"><tr><td style="border:none">' + (GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td></tr></table>');
                $("#hidGeneralAvailDt").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());

                if (document.getElementById('hdnAlgorithm').value == 'Y') {
                    $("#txtBlindDate1").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                    $("#txtGeneralAvailDt1").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                    $("#txtPAADDate1").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                }                
            }
        }

        function calcDates_whendatesareblank() {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (isDate($("#txtRTPDate").val())) {

                var RTPDate = new Date($("#txtRTPDate").val());

                var GeneralAvailDt = RTPDate;

                var monday = getMonday(RTPDate);
                
                var firstDay = new Date(GeneralAvailDt.getFullYear(), GeneralAvailDt.getMonth(), 1);
                var PAADDate = new Date((monday.getMonth() + 1) + '/' + monday.getDate() + '/' + monday.getFullYear());

                var BlindDate;
                BlindDate = new Date(PAADDate.getFullYear(), (PAADDate.getMonth() - 1), 1);

                if (BlindDate < Today)
                    BlindDate = new Date(Today.getFullYear(), Today.getMonth(), Today.getDate() + 7);

                if ($("#hidCplBlindDt").val() != "") {
                    var existingsadate = new Date($("#hidCplBlindDt").val());
                    if (existingsadate < Today)
                        BlindDate = existingsadate;
                }

                $("#mktgBlindDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td></tr></table>');
                $("#hidCplBlindDt").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtPAADDate").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                $("#mktgPAADDateText").html('<table width="100%" border="0"><tr><td style="border:none">' + (PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear() + '</span></td><td align=Right style="border:none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td></tr></table>');
                $("#mktgGeneralAvailDtText").html('<table width="100%" border="0"><tr><td style="border:none">' + (GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear() + '</td><td align=Right style="border:none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td></tr></table>');
                $("#hidGeneralAvailDt").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());

                $("#txtBlindDate1").val((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                $("#txtGeneralAvailDt1").val((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                $("#txtPAADDate1").val((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());


                if (document.getElementById('hdnAlgorithm_blankdates').value == 'Y') {
                    $("#RTP1").html($("#hidRTP").val());
                    $("#SA1").html((BlindDate.getMonth() + 1) + '/' + BlindDate.getDate() + '/' + BlindDate.getFullYear());
                    $("#GA1").html((GeneralAvailDt.getMonth() + 1) + '/' + GeneralAvailDt.getDate() + '/' + GeneralAvailDt.getFullYear());
                    $("#PAAD1").html((PAADDate.getMonth() + 1) + '/' + PAADDate.getDate() + '/' + PAADDate.getFullYear());
                    $("#EOM1").html($("#hidEOM").val());
                }
            }
        }

        function checkRTPDate()
        {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (! isDate($("#txtRTPDate").val())) {
                alert("The RTP/MR Date is not in a correct format.");
                return false;
            }
            return true;
        }

        function checkEMDate() {
            var q = new Date();
            var m = q.getMonth();
            var d = q.getDate();
            var y = q.getFullYear();
            var Today = new Date(y, m, d);

            if (! isDate($("#txtMarketingDiscDate").val())) {
                alert("The End of Manufaturing (EM) Date is not in a correct format.");
                return false;
            }
        }

        function getMonday(d) {
            d = new Date(d);
            
            var day = d.getDay(),
                diff = d.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday

            return new Date(d.setDate(diff));
        }

        function isDate(txtDate) {            
            var currVal = txtDate;
            if (currVal == '')
                return false;

            var rxDatePattern = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
            var dtArray = currVal.match(rxDatePattern); // is format OK?

            if (dtArray == null)
                return false;

            dtMonth = dtArray[1];
            dtDay = dtArray[3];
            dtYear = dtArray[5];

            if (dtMonth < 1 || dtMonth > 12)
                return false;
            else if (dtDay < 1 || dtDay > 31)
                return false;
            else if ((dtMonth == 4 || dtMonth == 6 || dtMonth == 9 || dtMonth == 11) && dtDay == 31)
                return false;
            else if (dtMonth == 2) {
                var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
                if (dtDay > 29 || (dtDay == 29 && !isleap))
                    return false;
            }

            return true;
        }


        function Body_OnLoad() {
            
            //------------------------- COMBINED AV FUNCTIONALITY------(santodip)------------------------ 
            var FFID = frmMain.hidFID.value;
            var FFName = frmMain.hidFName.value;
            var FFRequiresRoot = frmMain.hidFRequiresRoot.value;
            var FFComponentLinkage = frmMain.hidFComponentLinkage.value;
            var FFComponentRootID = frmMain.hidFComponentRootID.value;
            var FFGPGDescription = frmMain.hidFGPGDescription.value;
            var FFMarketingDescriptionPMG = frmMain.hidFMarketingDescriptionPMG.value;
            var FFMarketingDescription = frmMain.hidFMarketingDescription.value;
            var FFSCMCategoryID_singlefeature = frmMain.hidFSCMCategoryID_singlefeature.value;
            var FFAliasID = frmMain.hidFAliasID.value;
            var FFAbbreviation = frmMain.hidFAbbreviation.value;
            var FFPlatform = frmMain.hidFPlatform.value;
            var FFPulsarPlusDivId = frmMain.hidPulsarPlusDivId.value;

            switch (frmMain.hidMode.value) {
                case "add":
                    EditAv();
                    EditFeatureCat();
                    chkSDFFlag_onclick();
                    //------------------------- COMBINED AV FUNCTIONALITY------(santodip)------------------------- 
                    SetAVDescriptions('true', FFID, FFName, FFGPGDescription, FFMarketingDescription, FFMarketingDescriptionPMG, FFRequiresRoot, FFComponentLinkage, FFComponentRootID, FFSCMCategoryID_singlefeature, FFAliasID, FFAbbreviation, FFPlatform);
                    SelectDefaultSCMCategoryProductLine(FFSCMCategoryID_singlefeature,false);
                    break;
                case "clone":
                    EditAv();
                    EditFeatureCat();
                    EditProductLine();
                    chkSDFFlag_onclick();
                    break;
            }

            switch (frmMain.hidFunction.value) {
                case "close":
                    var Ids_Skus = '', Ids_Cto = '', Rcto_Skus = '', Rcto_Cto = '', BSAMB = '';
                    var AvNo = '', GpgDescription = '', MarketingDesc = '', MarketingDescPMG = '', ConfigRules = '', TextAvId = '', Group1 = '', Group2 = '', Group3 = '', Group4 = '', Group5 = '',  Group6 = '', Group7 = '', Weight = '', GSEndDt = '', ProductLine = '', RTP = '', PAAD = '', SA = '', GA = '', EOM = '', EDITMODE='';
                    var PBID = ''
                    var Releases = ''; //task 16227 - Releases

                    if (frmMain.chkIdsSkus != undefined) {
                        if (frmMain.chkIdsSkus.checked)
                            Ids_Skus = 'X';
                    }
                    if (frmMain.chkIdsCto != undefined) {
                        if (frmMain.chkIdsCto.checked)
                            Ids_Cto = 'X';
                    }
                    if (frmMain.chkRctoSkus != undefined) {
                        if (frmMain.chkRctoSkus.checked)
                            Rcto_Skus = 'X';
                    }
                    if (frmMain.chkRctoCto != undefined) {
                        if (frmMain.chkRctoCto.checked)
                            Rcto_Cto = 'X';
                    }
                    if (frmMain.chkBSAMBparts != undefined) {
                        if (frmMain.chkBSAMBparts.checked)
                            BSAMB = 'X';
                    }
                        
                    AvID = frmMain.hidAVID.value;
                    PBID = frmMain.BID.value;
                    AvNo = frmMain.AvNo.value;
                    FCID = frmMain.cboCategory.value;
                    GpgDescription = frmMain.txtAvGpgDescription.value;
                    MarketingDesc = frmMain.txtMarketingDesc.value;
                    MarketingDescPMG = frmMain.txtMarketingDescPMG.value;
                    ConfigRules = frmMain.txtConfigRules.value;
                    RulesSyntax = frmMain.txtRulesSyntax.value;
                    TextAvId = frmMain.txtAvId.value;
                    Group1 = frmMain.txtGroup1.value;
                    Group2 = frmMain.txtGroup2.value;
                    Group3 = frmMain.txtGroup3.value;
                    Group4 = frmMain.txtGroup4.value;

                    RTP = frmMain.txtRTPDate.value;
                    PAAD = frmMain.txtPAADDate1.value;
                    SA = frmMain.hidCplBlindDt.value;
                    GA = frmMain.hidGeneralAvailDt.value;
                    EOM = frmMain.txtMarketingDiscDate.value;
                   
                    //if (frmMain.txtGroup5)
                    Group5 = frmMain.txtGroup5.value;
                    Group6 = frmMain.txtGroup6.value;
                    Group7 = frmMain.txtGroup7.value;

                    Weight = frmMain.txtWeight.value;
                    GSEndDt = frmMain.txtGSEndDt.value;
                    //task 16227 - - Releases
                    Releases = frmMain.txtReleaseAdded.value;
                    var ProductlineSelectedText = frmMain.cboProductLine.options[frmMain.cboProductLine.selectedIndex].text;
                    var res = ProductlineSelectedText.split("-");
                    ProductLine = res[0];
                    if (window.frmMain.hidMode.value.toLowerCase() == 'edit') {
                        if ((window.frmMain.txtOldCategory.value == frmMain.cboCategory.value || window.frmMain.txtFromTodayPage.value == "1")
                            //if featureID changed, need to refresh the grid 
                            //if description or feature changed and is a base av and scm category is not changed then we only refresh a row
                            && ((window.frmMain.txtFeatureChanged.value == "" && window.frmMain.txtDescriptionChanged.value == "" && $("#hidParentID").val() == 0)
                            || (window.frmMain.txtFeatureChanged.value == "" && window.frmMain.txtDescriptionChanged.value == "" && $("#hidParentID").val() > 0)
                            || (window.frmMain.txtDescriptionChanged.value == "1" && $("#hidParentID").val() == 0 && window.frmMain.txtFeatureChanged.value == ""))
                            ) {
                            if (IsFromPulsarPlus()) {
                                ClosePulsarPlusPopup();
                                window.parent.parent.parent.popupCallBack(1);
                            }
                            else if (FFPulsarPlusDivId != undefined && FFPulsarPlusDivId != "") {
                                // For Reload PulsarPlusPmView Tab
                                parent.window.parent.reloadFromPopUp(FFPulsarPlusDivId);
                                // For Closing current popup
                                parent.window.parent.closeExternalPopup();
                            }
                            else {
                                parent.window.parent.ReloadAVData(AvID, AvNo, GpgDescription, MarketingDesc, MarketingDescPMG, ConfigRules, RulesSyntax, TextAvId, Group1, Group2, Group3, Group4, Group5, Group6, Group7, Ids_Skus, Ids_Cto, Rcto_Skus, Rcto_Cto, Weight, GSEndDt, ProductLine, PBID, RTP, PAAD, SA, GA, EOM, BSAMB, Releases);
                                parent.window.parent.ClosePropertiesDialog();
                            }
                        }
                        else {
                            if (IsFromPulsarPlus()) {
                                ClosePulsarPlusPopup();
                            }
                            else if (FFPulsarPlusDivId != undefined && FFPulsarPlusDivId != "") {
                                // For Reload PulsarPlusPmView Tab
                                parent.window.parent.reloadFromPopUp(FFPulsarPlusDivId);
                                // For Closing current popup
                                parent.window.parent.closeExternalPopup();
                            }
                            else {
                                parent.window.parent.ClosePropertiesDialog(1);
                            }
                        }
                    }
                    else if (IsFromPulsarPlus()) {
                        ClosePulsarPlusPopup();
                    }
                    else if (FFPulsarPlusDivId != undefined && FFPulsarPlusDivId != "") {
                        // For Reload PulsarPlusPmView Tab
                        parent.window.parent.reloadFromPopUp(FFPulsarPlusDivId);
                        // For Closing current popup
                        parent.window.parent.closeExternalPopup();
                    }
                    else if (window.frmMain.hidMode.value.toLowerCase() == 'add' || window.frmMain.hidMode.value.toLowerCase() == 'clone')
                        parent.window.parent.ClosePropertiesDialog(1);
                    else
                        window.close();
                    break;
            }		
	
            if (typeof (window.parent.frames["LowerWindow"].frmButtons) == 'object') {
                if (window.frmMain.hidMode.value.toLowerCase() == 'add' || window.frmMain.hidMode.value.toLowerCase() == 'clone') {
                    window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = false;
                    if (window.frmMain.hidMode.value.toLowerCase() == 'add' && window.frmMain.hidSCMCat.value.toUpperCase() == 'BASE UNIT')//hidSCMCat
                    {
                        window.parent.frames["LowerWindow"].frmButtons.cmdClone.disabled = true;
                    }
                }            
                else if (window.frmMain.hidMode.value.toLowerCase() == 'edit') {
                    if (window.frmMain.hidStatus.value.toLowerCase() != "o" && window.frmMain.hidStatus.value.toLowerCase() != "d") {
                        window.parent.frames["LowerWindow"].frmButtons.cmdOK.disabled = false;
                    }
                    if (window.frmMain.hidSCMCat.value.toUpperCase() == 'BASE UNIT')//hidSCMCat
                    {
                        if (window.parent.frames["LowerWindow"].frmButtons.cmdClone != null)
                            window.parent.frames["LowerWindow"].frmButtons.cmdClone.disabled = true;
                    }
                }
            }

        LoadExistingDDLValues(frmMain.ElementValues.value);

        //var ID = frmMain.ID.value;
        var ExistingValues = frmMain.ExistingNameElements.value;

        if (frmMain.CategoryOpt.value != null && (frmMain.ViaAvCreate.value == 0 || window.frmMain.hidMode.value.toLowerCase() == 'add')) {
            LoadDelRootValues(frmMain.CategoryOpt.value, frmMain.DelRootID.value);
        }
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
                //var devEdit = document.getElementById("divEditDeliverableRoot");
                //devEdit.style.display = "";
                frmMain.cboDeliverables.options(frmMain.cboDeliverables.selectedIndex).value = 0;
            } else {
                document.getElementById("delRootCbo").innerHTML = request.responseText;
                //var devEdit = document.getElementById("divEditDeliverableRoot");
                //devEdit.style.display = "";
            }
        }
        //------------------------- COMBINED AV FUNCTIONALITY------(santodip)------------------------- 
        function SetAVDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform)
        {
            if (Refresh) {             

                if (!SetSCMCategory(SCMCategoryID, Abbreviation, FeatureID, $("#hidParentID").val()))
                    return;

                if (RequiresRoot == "True" && ComponentLinkage == "Linked") {
                    var str = "<a href=" + "javascript:OpenRoot('" + ComponentRootID + "')" + " id=linkroot>" + ComponentRootID + "</a>"
                    $("#delRootText").html(str);
                    document.getElementById('DelRootID').value = ComponentRootID;
                }
                else {
                    var str = "";
                    $("#delRootText").html(str);
                    document.getElementById('DelRootID').value = 0;
                }
               
                if (Abbreviation == "BUNIT") {
                    $("#trPlatform").show();
                }
                else
                    $("#trPlatform").hide();
                
                $("#hidAbbreviation").val(Abbreviation);
                $("#divPlatform").append(Platform);
                              
                document.getElementById('txtfeatureIDDesc').value = FeatureID;
                $("#assignfeatureIDText").html(FeatureID);
                document.getElementById('txtFeatureID').value = FeatureID;

                $("#divfeatureIDDesc").html("");

                document.getElementById('txtfeatureNameDesc').value = FeatureName;
                $("#assignfeatureNameText").html(FeatureName);

                document.getElementById('txtAvGpgDescription').value = GPGDescription;
                $("#gpgText").html(GPGDescription);
                document.getElementById('txtMarketingDesc').value = MarketingDescription;
                $("#mktgText").html(MarketingDescription);
                document.getElementById('txtMarketingDescPMG').value = MarketingDescriptionPMG;
                $("#mktgTextPMG").html(MarketingDescriptionPMG);

            }

            $("#txtAvGpgDescription").prop('readonly', true);
            $("#txtMarketingDesc").prop('readonly', true);
            $("#txtMarketingDescPMG").prop('readonly', true);
        }

        function SetSCMCategory(SCMCategoryID, Abbreviation, NewFeatureID, AvParentID) {           
            var GoodtoGo = true;
            var c = document.getElementById("cboCategory");
            var opts = c.options.length;
            var intCategoryID = c.options[c.selectedIndex].value;
            
            if (intCategoryID != SCMCategoryID && intCategoryID > 0 && SCMCategoryID > 0 && Abbreviation != "BUNIT") {
                if ($("#hidParentID").val() == 0) {
                    $("#Promptdialog").dialog({
                        resizable: false,
                        height: 140,
                        modal: true,
                        buttons: {
                            "Yes": function () {                                    
                                for (var i = 0; i < opts; i++) {
                                    if (c.options[i].value == SCMCategoryID) {
                                        c.options[i].selected = true;                                        
                                        document.getElementById("tagCategory").value = SCMCategoryID;
                                        document.getElementById("scmCategoryTxt").innerHTML = c.options[i].text;
                                        break;
                                    }
                                }
                                $(this).dialog("close");
                            },
                            "No": function () {
                                $(this).dialog("close");
                            }
                        }
                    });
                } else {
                    var GEO = FindMissingSubCategories(AvParentID, NewFeatureID);
                    if(GEO != "")
                    {
                        alert("Pulsar can not find the SCM Subcategories for all the localizations of the new Feature being selected.  \n\nThe new Feature can not be used to replace the current Feature because it does not have all the Subcategories.  \n\nPlease select another Feature or ask the Category Admin user to create the subcategories for the new Feature being selected, GEO: " + GEO + ".");
                        GoodtoGo = false;
                    }
                    else {
                        for (var i = 0; i < opts; i++) {
                            if (c.options[i].value == SCMCategoryID) {
                                var index = c.options[i].text.indexOf('>>>');
                                var subCategory = c.options[i].text;
                                if (index > 0) {
                                    subCategory = subCategory.replace(subCategory.substr(0, index + 3), "");
                                }
                                var oldSubCategory = document.getElementById("scmCategoryTxt").innerHTML.replace("&gt;&gt;&gt;", "");
                                if (confirm("After clicking OK, Pulsar will replace the SCM subcategory of \"" + oldSubCategory + "\" with \"" + subCategory + "\" for \"" + $("#AvNo").val() + "\".  \n\nKeep in mind replacing the subcategory for one localized AV will replace all the subcategories for the other localizations of this AV. \n\nClick OK to continue or Cancel to stop Change Feature.")) {
                                    c.options[i].selected = true;
                                    document.getElementById("tagCategory").value = SCMCategoryID;
                                    document.getElementById("scmCategoryTxt").innerHTML = ">>> " + subCategory;
                                    break;
                                }
                                else {
                                    GoodtoGo = false;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            else if (intCategoryID == 0 && SCMCategoryID > 0) {
                for (var i = 0; i < opts; i++) {
                    if (c.options[i].value == SCMCategoryID) {
                        c.options[i].selected = true;   
                        var subCategory = c.options[i].text;
                        if (AvParentID > 0) {
                            var GEO = FindMissingSubCategories(AvParentID, NewFeatureID);
                            if(GEO != "")
                            {
                                alert("Pulsar can not find the SCM Subcategories for all the localizations of the new Feature being selected.  \n\nThe new Feature can not be used to replace the current Feature because it does not have all the Subcategories.  \n\nPlease select another Feature or ask the Category Admin user to create the subcategories for the new Feature being selected, GEO: " + GEO + ".");
                                GoodtoGo = false;
                            }
                            else {
                                var index = subCategory.indexOf('>>>');
                                if (index > 0) {
                                    subCategory = subCategory.replace(subCategory.substr(0, index), "");
                                }
                            }
                        }
                        document.getElementById("tagCategory").value = SCMCategoryID;
                        document.getElementById("scmCategoryTxt").innerHTML = subCategory;
                        break;
                    }
                }
            }
            else if (intCategoryID != SCMCategoryID && intCategoryID > 0 && SCMCategoryID > 0 && Abbreviation == "BUNIT") {
                for (var i = 0; i < opts; i++) {
                    if (c.options[i].value == SCMCategoryID) {
                        c.options[i].selected = true;                        
                        document.getElementById("tagCategory").value = SCMCategoryID;
                        document.getElementById("scmCategoryTxt").innerHTML = c.options[i].text;
                        break;
                    }
                }
            }
            else if (SCMCategoryID == 0) {
                if ($("#hidAVID").val() > 0 && $("#hidConfigCode").val() != "" && AvParentID > 0) {
                    var GEO = FindMissingSubCategories(AvParentID, NewFeatureID);
                    alert("Pulsar can not find the SCM Subcategories for all the localizations of the new Feature being selected.  \n\nThe new Feature can not be used to replace the current Feature because it does not have all the Subcategories.  \n\nPlease select another Feature or ask the Category Admin user to create the subcategories for the new Feature being selected, GEO: " + GEO + ".");
                    GoodtoGo = false;
                } else {
                    c.options[0].selected = true;
                    document.getElementById("tagCategory").value = 0;
                    document.getElementById("scmCategoryTxt").innerHTML = "";
                    $("#featureCatText").hide();
                    $("#featureCatSelect").show();
                } 
            }

            return GoodtoGo;
        }

        function EditAv() {            
            avinput.style.display = "";
            avtext.style.display = "none";

            if ($("#hidParentID").val() > 0) {
                if ($("#AvNo").val() == "" && $("#hidBaseAvNo").val() != "") {
                    if (confirm("Would you like to use parent AV# '" + $("#hidBaseAvNo").val() + "' as your base?")) {
                        $("#AvNo").val($("#hidBaseAvNo").val() + "#" + $("#hidConfigCode").val());
						$("#AvNo").focus().val();
                    }
                    else {
                        $("#AvNo").val("#" + $("#hidConfigCode").val());
                    	$("#AvNo").focus();
                    }
                }
                else {
                    $("#AvNo").focus();
                }
            }
            else {
                $("#AvNo").focus();
            }
        }

        function EditFeatureID() {
            featureIDDesc.style.display = "";
            featureIDText.style.display = "none";
        }

        function EditFeatureName() {
            featureNameDesc.style.display = "";
            featureNameText.style.display = "none";
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

        function EditProductLine() {
            ProductLineSelect.style.display = "";
            ProductLineEdit.style.display = "none";
            SelectDefaultSCMCategoryProductLine(frmMain.cboCategory.options[frmMain.cboCategory.selectedIndex].value,false);
        }

        function EditProductRelease() {
            ProductRelease.style.display = "";
            ProductReleaseEdit.style.display = "none";
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
                    //divEditGpgDesc.style.display = ""
                    //divEditMktgDesc.style.display = ""
                    //divEditMktgDescPMG.style.display = ""
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
            strID = window.parent.showModalDialog("<%= AppRoot %>/SupplyChain/PDMFeedbackFrame.asp?AvId=" + AvId, "", "dialogWidth:1095px;dialogHeight:510px;edge: Sunken;center:Yes; help: No;resizable: No;status: No;scrollbars: No;");
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
                strNewRow = strNewRow + "<tr><td><font size=1><b>GPG Description:&nbsp;&nbsp;&nbsp;<br />(40-char PHweb)</b></font></td><td><label ID=lblFinishedName3><font color=black>" + frmMain.txtAVName3.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
                strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Short Description:&nbsp;&nbsp;&nbsp;<br />(40 Char)</b></font></td><td><label ID=lblFinishedName5><font color=black>" + frmMain.txtAVName5.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
                strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Long Description:&nbsp;&nbsp;&nbsp;<br />(100 Char)</b></font></td><td><label ID=lblFinishedName7><font color=black>" + frmMain.txtAVName7.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
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
		
            }
            else {
                trGPGDesc.style.display = "";
                trMarketingDesc40.style.display = "";
                trMarketingDesc100.style.display = "";

                NameRowUpdate.style.display = "none";
		

            }

	
            // update the default ProductLine for the Selected SCM
            var sSCMCategorySelected = frmMain.cboCategory.options[frmMain.cboCategory.selectedIndex].value;
            
            SelectDefaultSCMCategoryProductLine(sSCMCategorySelected,true);

        }

        function SelectDefaultSCMCategoryProductLine(SCMCategorySelected, ShowMessage) {
            var sAddedline;
            var sSCMCatInheritProductLine;

            sSCMCatInheritProductLine = "";
            sAddedline = "0";

            //find out if the scmcategory has inherit equals to null or zero
            var bInheritFromSCM = false;

            var sProductProductLine = frmMain.hidProductProductLine.value;
            var element = document.getElementById('cboProductLine'); //product line drop down

            //Id , Description
            var sProductLinesAll = frmMain.hidProductLinesAll.value;
            var arrProductLinesAll = sProductLinesAll.split(",")

            // delete existing ProductLines
            var m;
            for (m = element.options.length - 1; m >= 0; m--) {
                element.remove(m);
            }
            var intSCMCatIndex = 0;
            var sAllSCMcategoryProductLines = frmMain.hidSCMCategoryProductLines.value;
            var arrSCMPCategoriesProducLinesAll = sAllSCMcategoryProductLines.split(",");
            for (var i = 0; i < arrSCMPCategoriesProducLinesAll.length; i++) {
                var arrSCMProductLine = arrSCMPCategoriesProducLinesAll[i].split("#");                
                if (arrSCMProductLine[0] == SCMCategorySelected) {
                    bInheritFromSCM = true;
                    intSCMCatIndex = i;
                    break;
                }
            }            
            if (bInheritFromSCM) {
                if (ShowMessage)
                    alert('Changing the SCM Category has changed the list of Product Line, click the Edit link (if exists) and select a new product line.');                                
                //get the product line from the scm list
                option = document.createElement("OPTION");
                option.text = "Please select a product line";
                option.value = "";
                element.add(option);
                var arrSCMProductLine = arrSCMPCategoriesProducLinesAll[intSCMCatIndex].split("#");
                var NoOfProductLines = 0;
                NoOfProductLines = Number(arrSCMProductLine[2]);
                if (NoOfProductLines > 0) {                        
                    //add Product Lines                   
                    if (arrSCMProductLine[1].length > 0) {
                        for (var n = 0; n < arrSCMProductLine[1].split(";").length; n++) {
                            var productlineitem = arrSCMProductLine[1].split(";")[n];
                            option = document.createElement("OPTION");
                            option.value = productlineitem.split("|")[0];
                            option.text = productlineitem.split("|")[1];
                            element.add(option);
                        }
                    }
                }
                else {
                    for (var n = 0; n < arrProductLinesAll.length; n++) {
                        var arrProductLine = arrProductLinesAll[n].split("#")
                        //add Product Lines
                        option = document.createElement("OPTION");
                        option.value = arrProductLine[0];
                        option.text = arrProductLine[1];
                        element.add(option);
                    }
                }                  
                document.getElementById("tdProductLine").innerHTML = "";
            }
            else {
                //get all the product lines and defatult to the selected product line in the product version table
                //load all the ProductLines
                option = document.createElement("OPTION");
                option.text = "Please select a product line";
                option.value = "";
                element.add(option);

                for (var n = 0; n < arrProductLinesAll.length; n++) {
                    var arrProductLine = arrProductLinesAll[n].split("#")           
                    //add Product Lines
                    option = document.createElement("OPTION");
                    option.value = arrProductLine[0];
                    option.text = arrProductLine[1];
                    element.add(option);
                }        
                //select the required one.
                element.value = sProductProductLine;
                document.getElementById("tdProductLine").innerHTML = frmMain.cboProductLine.options[frmMain.cboProductLine.selectedIndex].text; //sProductProductLine;
            }    
        }

        function cboElement_onchange() {
            var strName3 = "";
            var strName5 = "";
            var strName7 = "";

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
                strName3 = strName3 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
                strName5 = strName5 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
                strName7 = strName7 + lblPreNamePart.innerText + frmMain.cboElement.options(frmMain.cboElement.selectedIndex).text + lblPostNamePart.innerText + lblNameFieldDiv.innerText;
            }
            else {
                for (i = 0; i < frmMain.cboElement.length; i++) {
			
                    if (frmMain.cboElement(i).tagName == "INPUT") {
				
                        if (frmMain.cboElement(i).value != "" && frmMain.cboElement(i).className == "name=cboElement") {
                            strName3 = strName3 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                            strName5 = strName5 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
                            strName7 = strName7 + lblPreNamePart(i).innerText + frmMain.cboElement(i).value + lblPostNamePart(i).innerText + lblNameFieldDiv(i).innerText;
					
                        } else if (frmMain.cboElement(i).value != "") {
                            for (k = 0; k < Elements.length; k++) {
                                ElementValues = Elements[k].split("|");
                                if (ElementValues[1] == frmMain.cboElement(i).className) {
                                    if (trim(ElementValues[3]) == "[text]") {
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

            frmMain.txtAVName3.value = strName3;
            lblFinishedName3.innerText = strName3;

            frmMain.txtAVName5.value = strName5;
            lblFinishedName5.innerText = strName5;

            frmMain.txtAVName7.value = strName7;
            lblFinishedName7.innerText = strName7;


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
        }

        function trim(stringToTrim) {
            return stringToTrim.replace(/^\s+|\s+$/g, "");
        }

        function textCounter(field, countfield, maxlimit) {
            if (field.value.length > maxlimit)
                field.value = field.value.substring(0, maxlimit);
            else {
                if (countfield != null)
                    countfield.innerHTML = maxlimit - field.value.length;
            }
        }

        function SaveNameElements() {
            var strBuild = "";
            var strComments = "";
            var i;
	
            var cboElement = document.getElementById("cboElement");
	
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
            frmMain.strNameElements.value = trim(Elements);
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
                    //divEditGpgDesc.style.display = "none"
                    //divEditMktgDesc.style.display = "none"
                    //divEditMktgDescPMG.style.display = "none"
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
                strNewRow = strNewRow + "<tr><td><font size=1><b>GPG Description:&nbsp;&nbsp;&nbsp;<br />(40-char PHweb)</b></font></td><td><label ID=lblFinishedName3><font color=black>" + frmMain.txtAVName3.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
                strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Short Description:&nbsp;&nbsp;&nbsp;<br />(40 Char)</b></font></td><td><label ID=lblFinishedName5><font color=black>" + frmMain.txtAVName5.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
                strNewRow = strNewRow + "<tr><td><font size=1><b>Marketing Long Description:&nbsp;&nbsp;&nbsp;<br />(100 Char)</b></font></td><td><label ID=lblFinishedName7><font color=black>" + frmMain.txtAVName7.value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
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

    </script>

    <script type="text/javascript">
        $(function () {
            $("#iframeDialog").dialog({
                modal: true,
                autoOpen: false,
                width: 900,
                height: 800,
                close: function () {
                    $("#modalDialog").attr("src", "about:blank");
                },
                resizable: false

            });    
            

            if ($("#hidAbbreviation").val() == "BUNIT")
                $("#trPlatform").show();
        });

        function ShowIframeDialog(QueryString, Title, DlgWidth, DlgHeight) {
            if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
            if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
            $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });
            $("#modalDialog").attr("width", "98%");
            $("#modalDialog").attr("height", "98%");
            $("#modalDialog").attr("src", QueryString);
            $("#iframeDialog").dialog("option", "title", Title);
            $("#iframeDialog").dialog("open");
        }

        function ShowSelectFeatureDialog() {
            if ($("#hidParentID").val() > 0) {
                if (confirm("Changing the Feature for the localized AV will change the Feature for all the other localizations too.")) {
                    //we want add single feature to know that this av is a localized av so please only show me the feature that is belong to localized categories
                    var url = "../../IPulsar/SCM/SCM_AddSingleFeature.aspx?IsDesktop=" + "<%=LCase(sIsDesktop) %>&IsLocalized=1&CurrentUserID= <%=m_UserID%> &ProductBrandID=" + frmMain.BID.value + "&OptionCode=" + $("#hidConfigCode").val();
                    if (IsFromPulsarPlus()) {
                        url = "../../IPulsar/SCM/SCM_AddSingleFeature.aspx?IsDesktop=" + "<%=LCase(sIsDesktop) %>&IsLocalized=1&CurrentUserID= <%=m_UserID%> &ProductBrandID=" + frmMain.BID.value + "&app=PulsarPlus&OptionCode=" + $("#hidConfigCode").val();
                        strID = window.showModalDialog(url, "Select Feature", "dialogWidth:980px;dialogHeight:800px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                        if (typeof (strID) != "undefined") {
                            var retValues = strID;
                            ResetAvDescriptions(retValues.refreshgrid, retValues.FeatureID, retValues.FeatureName, retValues.GPGDescription, retValues.MarketingDescription, retValues.MarketingDescriptionPMG, retValues.RequiresRoot, retValues.ComponentLinkage, retValues.ComponentRootID, retValues.SCMCategoryID, retValues.AliasID, retValues.Abbreviation, retValues.Platform);
                        }
                    }
                    else {
                        //parent.window.parent.ShowFeatureSelectDialog(url, "Select Feature", 980, 800);
                        //ShowFeatureSelectDialog(url, "Select Feature");
                        window.showModalDialog(url, window, "dialogWidth:900px;dialogHeight:800px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
                    }
                }
            }
            else {
                var url = "../../IPulsar/SCM/SCM_AddSingleFeature.aspx?IsDesktop=" + "<%=LCase(sIsDesktop) %>&IsLocalized=0&CurrentUserID= <%=m_UserID%> &ProductBrandID=" + frmMain.BID.value;
                if (IsFromPulsarPlus()) {
                    url = "../../IPulsar/SCM/SCM_AddSingleFeature.aspx?IsDesktop=" + "<%=LCase(sIsDesktop) %>&IsLocalized=0&CurrentUserID= <%=m_UserID%> &ProductBrandID=" + frmMain.BID.value + "&app=PulsarPlus";
                    strID = window.showModalDialog(url, "Select Feature", "dialogWidth:980px;dialogHeight:800px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                    if (typeof (strID) != "undefined") {
                        var retValues = strID;
                        ResetAvDescriptions(retValues.refreshgrid, retValues.FeatureID, retValues.FeatureName, retValues.GPGDescription, retValues.MarketingDescription, retValues.MarketingDescriptionPMG, retValues.RequiresRoot, retValues.ComponentLinkage, retValues.ComponentRootID, retValues.SCMCategoryID, retValues.AliasID, retValues.Abbreviation, retValues.Platform);
                    }
                }
                else {
                    //parent.window.parent.ShowFeatureSelectDialog(url, "Select Feature", 980, 800);
                    //ShowFeatureSelectDialog(url, "Select Feature");
                    window.showModalDialog(url, window, "dialogWidth:900px;dialogHeight:800px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
                }
            }
            
        }
        function OpenFeatureProperties(FeatureID, DeliveryType) {
            var url = "../../IPulsar/Features/FeatureProperties.aspx?FeatureID=" + FeatureID + "&DeliveryType=" + DeliveryType + "&ViewFrom=AvDetail&IsDesktop=" + "<%=LCase(sIsDesktop) %>";
            if (IsFromPulsarPlus()) {
                url = "../../IPulsar/Features/FeatureProperties.aspx?FeatureID=" + FeatureID + "&DeliveryType=" + DeliveryType + "&app=PulsarPlus&ViewFrom=AvDetail&IsDesktop=" + "<%=LCase(sIsDesktop) %>";
                strID = window.showModalDialog(url, "Feature Properties", "dialogWidth:980px;dialogHeight:800px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
                if (typeof (strID) != "undefined") {
                    var retValues = strID;
                    ResetAvDescriptions(retValues.refreshgrid, retValues.FeatureID, retValues.FeatureName, retValues.GPGDescription, retValues.MarketingDescription, retValues.MarketingDescriptionPMG, retValues.RequiresRoot, retValues.ComponentLinkage, retValues.ComponentRootID, 0, 0, "", "");
                }
            } else {
                // parent.window.parent.ShowFeatureSelectDialog(url, "Feature Properties", 980, 800);
                ShowFeatureSelectDialog(url, "Feature Properties");
            }
        }

        function ShowFeatureSelectDialog(QueryString, Title) {
            var DlgWidth = adjustWidth(97);
            var DlgHeight = adjustHeight(100);            
            //the dialog of feature do not reload after we close it. the code below is doing all that and we are all should call one function to open diag in this page.
            OpenPopUp(QueryString, DlgHeight, DlgWidth, Title, false, false, true, "divMultipleAVs", "ifMultipleAVs")
            /*window.showModalDialog(QueryString, "", "dialogWidth:"+DlgWidth+"px;dialogHeight:"+DlgHeight+"px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
        }

        function adjustWidth(percent) {
            return document.documentElement.offsetWidth * (percent / 100);
        }

        function adjustHeight(percent) {
            return (document.documentElement.offsetHeight * (percent / 100));
        }

        function OpenPopUp(link, newHeight, newWidth, title, noScrollBar, hideCloseButton, Resizable, divID, ifrID) {
            var $divPopup = $('#' + divID);
            $divPopup.dialog({
                height: newHeight,
                width: newWidth,
                modal: true,
                title: title,
                resizable: Resizable,
                draggable: true,
                open: function (event, ui) {
                    if (hideCloseButton)
                        $(this).parent().children().children('.ui-dialog-titlebar-close').hide();
                    else
                        $(this).parent().children().children('.ui-dialog-titlebar-close').show();

                    if (noScrollBar)
                        $divPopup.css('overflow', 'hidden');
                },
                close: function (event, ui) {
                    //everytime the jquery dialog is closed trigger this event to clear the iframe so when dialogue is called again it will show blank first then load with the url
                    $("#" + ifrID).attr("src", "");
                }
            });

            loadIframe(ifrID, link);
        }

        function loadIframe(iframeName, url) {
            var $iframe = $('#' + iframeName);
            $iframe.attr("width", "100%");
            $iframe.attr("height", "100%");
            if ($iframe.length) {
                $iframe.attr('src', url);
                return false;
            }
            return true;
        }

        function ClosePopUpViewFromAvDetail_Features(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform) {
            $("#divMultipleAVs").dialog("close");
            if (document.getElementById('modalDialog') != null) {
                ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, AliasID, Abbreviation, Platform);
            }            
        }

        function ClosePopUpViewFromAvDetail(refreshgrid, intFeatureID) {
            $("#divMultipleAVs").dialog("close");
            window.location.reload(refreshgrid);
        }
        function OpenAddExistingSharedAV() {
            var url = "../../IPulsar/Admin/SCM/AddExistingSharedAV.aspx?PVID=" + frmMain.ProductVersionID.value + "&BID=" + frmMain.BID.value;
            parent.window.parent.OpenAddExistingSharedAV(url, 890, 830);
            //window.showModalDialog(url, "", "dialogWidth:890px;dialogHeight:830px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
        }   

        function OpenRoot(CRID) {
            var url = "../../Pulsar/Component/Root/" + CRID;
            //alert(url);
            //ShowIframeDialog(url, "Root", 980, 800);
            //Root
            window.showModalDialog(url, "","dialogWidth:980px;dialogHeight:800px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	
        }


        function ClosePopUpViewFromAvDetail(refreshgrid, intFeatureID) {
            $("#iframeDialog").dialog("close");
            window.location.reload(true);
        }

        function ResetAvDescriptions(Refresh, FeatureID, FeatureName, GPGDescription, MarketingDescription, MarketingDescriptionPMG, RequiresRoot, ComponentLinkage, ComponentRootID, SCMCategoryID, aliasid, abbreviation, platform) {
            if (Refresh) {                
                if (!SetSCMCategory(SCMCategoryID, abbreviation, FeatureID, $("#hidParentID").val()))
                    return;

                if (RequiresRoot == "True" && ComponentLinkage == "Linked") {
                    var str = "<a href=" + "javascript:OpenRoot('" + ComponentRootID + "')" + " id=linkroot>" + ComponentRootID + "</a>"
                    $("#delRootText").html(str);
                    document.getElementById('DelRootID').value = ComponentRootID;
                }
                else {
                    var str = "";
                    $("#delRootText").html(str);
                    document.getElementById('DelRootID').value = 0;
                }

                if (abbreviation == "BUNIT")
                    $("#trPlatform").show();
                else
                    $("#trPlatform").hide();
                                
                $("#hidAbbreviation").val(abbreviation);
                $("#divPlatform").append(platform);
            
                document.getElementById('txtfeatureIDDesc').value = FeatureID;
                $("#assignfeatureIDText").html(FeatureID);
                document.getElementById('txtFeatureID').value = FeatureID;
            
                $("#divfeatureIDDesc").html("");

                document.getElementById('txtfeatureNameDesc').value = FeatureName;
                $("#assignfeatureNameText").html(FeatureName);

                if ($("#hidParentID").val() > 0) {
                    var regionCodes = document.getElementById('txtAvGpgDescription').value.split(' ');
                    var regionCode = regionCodes[regionCodes.length - 1];
                    GPGDescription = GPGDescription + " " + regionCode;
                    MarketingDescriptionPMG = MarketingDescriptionPMG + " " + regionCode;
                }

                document.getElementById('txtAvGpgDescription').value = GPGDescription;
                $("#gpgText").html(document.getElementById('txtAvGpgDescription').value);
                document.getElementById('txtMarketingDesc').value = MarketingDescription;
                $("#mktgText").html(document.getElementById('txtMarketingDesc').value);
                document.getElementById('txtMarketingDescPMG').value = MarketingDescriptionPMG;
                $("#mktgTextPMG").html(document.getElementById('txtMarketingDescPMG').value);
            }

            $("#txtAvGpgDescription").prop('readonly', true);
            $("#txtMarketingDesc").prop('readonly', true);
            $("#txtMarketingDescPMG").prop('readonly', true);
        }

        function NonSharedAV_onclick() {
            document.getElementById('hdnSharedValue').value = "0";            
        }

        function SharedAV_onclick() {
            document.getElementById('hdnSharedValue').value = "1";            
        }

        function ChangeStatus(BUFeatureID) {
            if ($("#BUAvail-" + BUFeatureID).text() == 'A')
                value = 'N/A';
            else
                value = 'A';

            $("#BUAvail-" + BUFeatureID).html("<a href='javascript:void(0)' onclick='ChangeStatus(" + BUFeatureID + ")'>" + value + "</a>");

            var Xml = '<?xml version="1.0" encoding="utf-8" ?>';
            Xml += '<BUs>';

            $("td[id^=BUAvail-]").each(function () {
                var IDs = $(this).attr('id').split('-');
                if ($(this).attr("OldStatus") != $(this).text()) {
                    Xml += '<BU BID="' + $("#BID").val() + '" BUFeatureID="' + IDs[1] + '" AvailStatus="' + $(this).text() + '" />';
                }
            });

            Xml += '</BUs>';
            $("#BUAvailList").val(Xml);
        }

        function onlyNumbers(evt) {
            var charCode;
            if (window.event)
                charCode = window.event.keyCode;   //if IE
            else
                charCode = evt.which; //if firefox
            if (charCode > 31 && (charCode < 48 || charCode > 57))
                return false;
            return true;
        }

        function onlyAlphabets(evt) {
            var charCode;
            if (window.event)
                charCode = window.event.keyCode;  //for IE
            else
                charCode = evt.which;  //for firefox
            if (charCode == 32) //for &lt;space&gt; symbol
                return true;
            if (charCode > 31 && charCode < 65) //for characters before 'A' in ASCII Table
                return false;
            if (charCode > 90 && charCode < 97) //for characters between 'Z' and 'a' in ASCII Table
                return false;
            if (charCode > 122) //for characters beyond 'z' in ASCII Table
                return false;
            return true;
        }

        function onlyAlphaNumbers(evt) {
            var charCode;
            if (window.event)
                charCode = window.event.keyCode;   //if IE
            else
                charCode = evt.which; //if firefox
       
            var avno = $("#AvNo").val();
            if (charCode == 35 && avno.length == 0)
                return false;
            else
                return true;

            if(!onlyAlphabets(evt) && !onlyNumbers())
                return false;               

            return true;
        }

        function EditMktgRTPDate(isAuto, RTPDate) {

            if ($('#hdnSharedValue1').val() == "True") {

                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    // remove FCS from all areas - task 20243
                    $("#dialog").attr("title", "Please Confirm");
                    $("#dialog").html("<p>Apply this new RTP/MR Date to reset the Select Availability (SA) Date, General Availability (GA) Date, and PA:AD (Intro Date) for this AV in this Product? </p> <p><ul><li> General Availability (GA) Date = RTP/MR Date</li> <li> PA:AD (Intro) Date = Monday of the week of RTP/MR Date </li> <li> Select Availability (SA) Date = one month prior to PA:AD and always the first of the month </li></p>");
                    $("#dialog").dialog({
                        resizable: false,
                        width: 600,
                        height: 400,
                        modal: true,
                        buttons: {
                            "Yes": function () {
                                $(this).dialog("close");                               
                                $("#txtRTPDate").val(RTPDate);
                                $("#RTP1").html(RTPDate);                                
                                mktgRTPDate.style.display = "";
                                mktgRTPDateText.style.display = "none";                          
                            },

                            "No": function () {
                                $(this).dialog("close");
                                $("#txtRTPDate").unbind('keyup');
                                mktgRTPDate.style.display = "";
                                mktgRTPDateText.style.display = "none";
                            }
                        }
                    });
                }

                else {
                    mktgRTPDate.style.display = "none";
                    mktgRTPDateText.style.display = "";
                }

            }
            else {
                $("#dialog").attr("title", "Please Confirm");
                $("#dialog").html("<p>Apply this new RTP/MR Date to reset the Select Availability (SA) Date, General Availability (GA) Date, and PA:AD (Intro Date) for this AV in this Product? </p> <p><ul><li> General Availability (GA) Date = RTP/MR Date</li> <li> PA:AD (Intro) Date = Monday of the week of RTP/MR Date </li> <li> Select Availability (SA) Date = one month prior to PA:AD and always the first of the month </li></p>");
                $("#dialog").dialog({
                    resizable: false,
                    width: 600,
                    height: 400,
                    modal: true,
                    buttons: {
                        "Yes": function () {
                            $(this).dialog("close");
                            document.getElementById('hdnAlgorithm').value = 'Y';                            
                                                   
                            $("#txtRTPDate").val(RTPDate);
                            $("#RTP1").html(RTPDate);
                            if (checkRTPDate()) {                                    
                               calcDates();
                            }
                            
                            mktgRTPDate.style.display = "";
                            mktgRTPDateText.style.display = "none";
                                
                        },

                        "No": function () {
                            $(this).dialog("close");
                            $("#txtRTPDate").unbind('keyup');
                            mktgRTPDate.style.display = "";
                            mktgRTPDateText.style.display = "none";
                            document.getElementById('hdnAlgorithm').value = '';
                        }
                    }
                });
            }
        }

        function EditMktgPAADDate1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgPAADDate.style.display = "";
                    document.getElementById('hdnflag').value = 'E';
                    mktgPAADDateText.style.display = "none";
                }
                else {
                    mktgPAADDate.style.display = "none";
                    mktgPAADDateText.style.display = "";
                }
            }
            else {
                mktgPAADDate.style.display = "";
                document.getElementById('hdnflag').value = 'E';
                mktgPAADDateText.style.display = "none";
            }
        }

        function EditMktgBlindDate1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgBlindDate.style.display = "";
                    document.getElementById('hdnflag1').value = 'E';
                    mktgBlindDateText.style.display = "none";
                }
                else {
                    mktgBlindDate.style.display = "none";
                    mktgBlindDateText.style.display = "";
                }
            }
            else {
                mktgBlindDate.style.display = "";
                document.getElementById('hdnflag1').value = 'E';
                mktgBlindDateText.style.display = "none";
            }
        }

        function EditMktgGeneralAvailDt1() {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    mktgGeneralAvailDt.style.display = "";
                    document.getElementById('hdnflag2').value = 'E';
                    mktgGeneralAvailDtText.style.display = "none";
                }
                else {
                    mktgGeneralAvailDt.style.display = "none";
                    mktgGeneralAvailDtText.style.display = "";
                }
            }
            else {
                mktgGeneralAvailDt.style.display = "";
                document.getElementById('hdnflag2').value = 'E';
                mktgGeneralAvailDtText.style.display = "none";
            }
        }

        function EditMktgDiscDate(isAuto, sEMDate) {

            if ($('#hdnSharedValue1').val() == "True") {
                var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                if (r == true) {
                    if (isAuto) {
                        $("#txtMarketingDiscDate").val(sEMDate);
                        $("#EOM1").html(sEMDate);
                    }

                    mktgDiscDate.style.display = "";
                    mktgDiscDateText.style.display = "none";                    
                }
                else {
                    mktgDiscDate.style.display = "none";
                    mktgDiscDateText.style.display = "";
                }
            }
            else {
                if (isAuto) {
                    $("#txtMarketingDiscDate").val(sEMDate);
                    $("#EOM1").html(sEMDate);
                }

                mktgDiscDate.style.display = "";
                mktgDiscDateText.style.display = "none";
            }
        }

        function getScheduleDates()
        {
            var sRTPDate = "";
            var sEMDate = "";

            $('input[name=chkRelease]:checkbox:checked').each(function () {
                if (sRTPDate == "") {
                    sRTPDate = $(this).attr('RTPDate');
                }
                else {
                    if ((new Date(sRTPDate).getTime() > new Date($(this).attr('RTPDate')).getTime())) {
                        sRTPDate = $(this).attr('RTPDate');
                    }
                }

                if (sEMDate == "") {
                    sEMDate = $(this).attr('EMDate');
                }
                else {
                    if ((new Date(sEMDate).getTime() < new Date($(this).attr('EMDate')).getTime())) {
                        sEMDate = $(this).attr('EMDate');
                    }
                }
            });
            
            if ((new Date($("#txtRTPDate").val())).getTime() != (new Date(sRTPDate)).getTime() && sRTPDate != "" && (new Date($("#txtMarketingDiscDate").val())).getTime() != (new Date(sEMDate)).getTime() && sEMDate != "")
            {
                if (confirm("There is new RTP Date 1 " + sRTPDate + " and new OEM Date " + sEMDate + " from releases that you have selected/unselected.  Would you like to replace them? ")) {
                    EditMktgRTPDate(true, sRTPDate);
                    EditMktgDiscDate(true, sEMDate);                    
                }
            }
            else if ((new Date($("#txtRTPDate").val())).getTime() != (new Date(sRTPDate)).getTime() && sRTPDate != "" && (new Date($("#txtMarketingDiscDate").val())).getTime() == (new Date(sEMDate)).getTime() && sEMDate != "")
            {
                if (confirm("There is new RTP Date 2 " + sRTPDate + " from releases that you have selected/unselected.  Would you like to replace it? ")) {
                    EditMktgRTPDate(true, sRTPDate);
                }
            }
            else if ((new Date($("#txtRTPDate").val())).getTime() == (new Date(sRTPDate)).getTime() && sRTPDate != "" && (new Date($("#txtMarketingDiscDate").val())).getTime() != (new Date(sEMDate)).getTime() && sEMDate != "") {
                if (confirm("There is new OEM Date " + sEMDate + " from releases that you have selected/unselected.  Would you like to replace current OEM Date with new OEM Date? ")) {
                    EditMktgDiscDate(true, sEMDate);
                }
            }
            else {
                if ($("#txtRTPDate").val() != "" && sRTPDate == "" )
                {                    
                    if (confirm("There is not a RTP Date from releases that you have selected/unselected.  Would you like to clear current RTP Date? ")) {
                        if ($('#hdnSharedValue1').val() == "True") {
                            var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                            if (r == true) {
                                $("#txtRTPDate").val("");
                                mktgRTPDate.style.display = "";
                                mktgRTPDateText.style.display = "none";

                                $("#txtPAADDate1").val("");
                                mktgPAADDate.style.display = "";
                                mktgPAADDateText.style.display = "none";

                                $("#txtBlindDate1").val("");
                                mktgBlindDate.style.display = "";
                                mktgBlindDateText.style.display = "none";

                                $("#txtGeneralAvailDt1").val("");
                                mktgGeneralAvailDt.style.display = "";
                                mktgGeneralAvailDtText.style.display = "none";

                                $("#RTP1").html("");
                                $("#SA1").html("");
                                $("#GA1").html("");
                                $("#PAAD1").html("");
                            }
                        }
                        else {
                            $("#txtRTPDate").val("");
                            mktgRTPDate.style.display = "";
                            mktgRTPDateText.style.display = "none";

                            $("#txtPAADDate1").val("");
                            mktgPAADDate.style.display = "";
                            mktgPAADDateText.style.display = "none";

                            $("#txtBlindDate1").val("");
                            mktgBlindDate.style.display = "";
                            mktgBlindDateText.style.display = "none";

                            $("#txtGeneralAvailDt1").val("");
                            mktgGeneralAvailDt.style.display = "";
                            mktgGeneralAvailDtText.style.display = "none";

                            $("#RTP1").html("");
                            $("#SA1").html("");
                            $("#GA1").html("");
                            $("#PAAD1").html("");
                        }
                    }
                }

                if ($("#txtMarketingDiscDate").val() != "" && sEMDate == "") {
                    if (confirm("There is not a OEM Date from releases that you have selected/unselected.  Would you like to clear current OEM Date? ")) {
                        if ($('#hdnSharedValue1').val() == "True") {
                            var r = confirm("This AV is shared with " + $('#hdnscmlist').val() + " etc.  Changes to dates will be applied across every SCM where this AV is shared.");
                            if (r == true) {
                                $("#txtMarketingDiscDate").val("");
                                $("#EOM1").html("");
                                mktgDiscDate.style.display = "";
                                mktgDiscDateText.style.display = "none";
                            }
                        }
                        else {
                            $("#txtMarketingDiscDate").val("");
                            $("#EOM1").html("");
                            mktgDiscDate.style.display = "";
                            mktgDiscDateText.style.display = "none";
                        }
                    }
                }
            }
        }

        function FindMissingSubCategories(AvParentID, NewFeatureID)
        {   
            var missingGeo = "";
            var ajaxurl = "FindMissingSubCategories.asp?AvParentID=" + AvParentID + "&NewFeatureID=" + NewFeatureID;
            $.ajax({
                url: ajaxurl,
                type: "GET",
                async: false,
                success: function (data) {
                    if (data != "")
                    {
                        missingGeo = data;
                    }
                },
                error: function (xhr, status, error) {
                    alert(error);
                },
                cache: false
            });   

            return missingGeo
        }
    </script>

</head>
<body onload="Body_OnLoad()">
    <form method="post" id="frmMain">
        <input type='hidden' name="BUAvailList" id="BUAvailList" value="" />
        <input id="hidMode" name="hidMode" type="HIDDEN" value="<%= LCase(sMode)%>" />
        <input id="hidFunction" name="hidFunction" type="HIDDEN" value="<%= LCase(sFunction)%>" />
        <input id="hidStatus" name="hidStatus" type="HIDDEN" value="<%= UCase(sStatus)%>" />
        <input id="hidOldGpgDescription" name="hidOldGpgDescription" type="HIDDEN" value="<%= sGpgDesc%>" />
        <input id="hidOldMS40Description" name="hidOldMS40Description" type="HIDDEN" value="<%= sMarketingDesc%>" />
        <input id="hidOldML100Description" name="hidOldML100Description" type="HIDDEN" value="<%= sMarketingDescPMG%>" />
        <input id="hidAvNo" name="hidAvNo" type="hidden" value="<%= sOriginalAvNo%>" />
        <input id="hidSCMCat" name="hidSCMCat" type="hidden" value="<%= sSCMCat%>" />
        <input id="hidIsDesktop" name="hidIsDesktop" type="HIDDEN" value="<%= sIsDesktop%>" />
        <input id="hidAVID" name="hidAVID" type="HIDDEN" value='<%= Request("AVID")%>' />
        <input id="hidCplBlindDt" name="hidCplBlindDt" type="HIDDEN" value="<%= sCplBlindDt%>" />
        <input id="hidRasDiscDt" name="hidRasDiscDt" type="HIDDEN" value="<%= sRasDiscDt%>" />
        <input id="hidUpc" name="hidUpc" type="hidden" value="<%= sUpc%>" />
        <input id="hidAbbreviation" name="hidAbbreviation" type="hidden" value="<%= sCategoryAbbr %>" />
        <input id="BID" name="BID" type="HIDDEN" value="<%= iBrandID%>" />
        <input id="hidRTPDt" name="hidRTPDt" type="HIDDEN" value="<%= sRTPDt%>" />
        <input id="hidPhWebInstruction" name="hidPhWebInstruction" type="HIDDEN" value="<%= sPhWebInstruction%>" />
        <input id="hidSDFFlag" name="hidSDFFlag" type="HIDDEN" value="<%= sSDFFlag%>" />
        <input id="hidGeneralAvailDt" name="hidGeneralAvailDt" type="HIDDEN" value="<%= sGeneralAvailDt%>" />
        <input id="txtPAADDate" name="txtPAADDate" type="HIDDEN" value="<%= sPAADDate%>" />
        <input id="hidParentID" name="hidParentID" type="hidden" value="<%= iParentID %>" />
        <input id="strNameElements" name="strNameElements" type="hidden" />
        <input type="hidden" id="hdnGroup1" name="hdnGroup1" value="" />
        <input type="hidden" id="hdnSharedValue" name="hdnSharedValue" value="0" />
        <input type="hidden" id="hidSCMCategoryProductLines" name="hidSCMCategoryProductLines" value="<%= sSCMCategoriesProductLines %>" />
        <input type="hidden" id="hidProductProductLine" name="hidProductProductLine" value="<%= sProductProductLine %>" />
        <input type="hidden" id="hidSCMCategoriesInheritProductLine" name="hidSCMCategoriesInheritProductLine" value="<%= sSCMCategoriesInheritProductLine %>" />
        <input type="hidden" id="hidProductLinesAll" name="hidProductLinesAll" value="<%= sProductLinesAll %>" />
        <input type="hidden" id="hidBaseParent" name="hidBaseParent" value="<% if bBaseParent then response.write "1" else response.write "0" %>" />
        <input id="hidEOM" name="hidEOM" type="HIDDEN" value="<%= sEOM%>" />
        <input id="hidRTP" name="hidRTP" type="HIDDEN" value="<%= sRTP%>" />               
        <input type="hidden" id="hdnSharedValue1" name="hdnSharedValue1" value="<%= bSharedAV %>" />
        <input type="hidden" id="hdnflag" name="hdnflag" value="" />
        <input type="hidden" id="hdnflag1" name="hdnflag1" value="" />
        <input type="hidden" id="hdnflag2" name="hdnflag2" value="" />
        <input type="hidden" id="hdnAlgorithm" name="hdnAlgorithm" value="" />
        <input type="hidden" id="hdnAlgorithm_blankdates" name="hdnAlgorithm_blankdates" value="" />
        <input type="hidden" id="hidBaseAvNo" name="hidBaseAvNo" value="<%= sBaseAvNo %>" />
        <input type="hidden" id="hdnscmlist" name="hdnscmlist" value="<%= scmlist %>" />
        <input type="hidden" id="hdnCanEditDates" name="hdnCanEditDates" value="<%= m_CanEditDates %>" />
             
        <input id="hidProductRelease" name="hidProductRelease" type="HIDDEN" value="<%= sProductRelease%>" />

        <input style="display: none" type="text" id="ExistingNameElements" name="ExistingNameElements" value="<%=strExistingNameElements%>">
        <input style="display: none" type="text" id="RequiresFormattedName" name="RequiresFormattedName" value="<%=strRequiresFormattedName%>">
        <input style="display: none" type="text" id="IsNameFormatted" name="IsNameFormatted" value="<%=IsNameFormatted%>">
        <input style="display: none" type="text" id="ElementValues" name="ElementValues" value="<%=strElementValues%>">
        <input style="display: none" type="text" id="DeliverableValues" name="DeliverableValues" value="<%=strDeliverableValues%>">
        <input style="display: none" type="text" id="txtPCListEmails" name="txtPCListEmails" value="<%=strPCList%>">
        <input style="display: none" type="text" id="txtAVNameOld" name="txtAVNameOld" value="<%=strAVNameOld%>">
        <input style="display: none" type="text" id="txtAVName3" name="txtAVName3" value="<%=strAVName3%>">
        <input style="display: none" type="text" id="txtAVName5" name="txtAVName5" value="<%=strAVName5%>">
        <input style="display: none" type="text" id="txtAVName7" name="txtAVName7" value="<%=strAVName7%>">
        <input style="display: none" type="text" id="AvPrefixValues" name="AvPrefixValues" value="<%=strAvPrefixValues%>">
        <input style="display: none" type="text" id="m_EditModeOn" name="m_EditModeOn" value="<%=m_EditModeOn%>">
        <input style="display: none" type="text" id="DelRootID" name="DelRootID" value="<%=iDeliverableRootID%>">
        <input style="display: none" type="text" id="CategoryOpt" name="CategoryOpt" value="<%=iCategoryOpt%>">
        <input style="display: none" type="text" id="DeliverableOpt" name="DeliverableOpt" value="<%=sDeliverableOpt%>">
        <input style="display: none" type="text" id="ViaAvCreate" name="ViaAvCreate" value="<%=iViaAvCreate%>">
        <input style="display: none" type="text" id="ProductVersionID" name="ProductVersionID" value="<%=m_ProductVersionID%>">
        <input style="display: none" type="text" id="txtOldCategory" name="txtOldCategory" value="<%=strSCMOldCategory%>" />
        <input style="display: none" type="text" id="txtFromTodayPage" name="txtFromTodayPage" value="<%=iFromTodayPage%>">
        <input style="display: none" type="text" id="txtFeatureID" name="txtFeatureID" value="<%=sFeatureID%>">
        <input style="display: none" type="text" id="txtIsMarketingScreen" name="txtIsMarketingScreen" value="0">
        <input style="display: none" type="text" id="txtFeatureChanged" name="txtFeatureChanged" value="<%=bFeaturechanged%>">
        <input style="display: none" type="text" id="txtDescriptionChanged" name="txtDescriptionChanged" value="<%=bDescriptionChanged%>">

        <!-- ------------------------- COMBINED AV FUNCTIONALITY------(santodip)------------------------- -->
        <input id="hidFID" name="hidFID" type="HIDDEN" value='<%= Request("FID")%>' />
        <input id="hidFName" name="hidFName" type="HIDDEN" value='<%= Request("FName")%>' />
        <input id="hidFRequiresRoot" name="hidFRequiresRoot" type="HIDDEN" value='<%= Request("FRequiresRoot")%>' />
        <input id="hidFComponentLinkage" name="hidFComponentLinkage" type="HIDDEN" value='<%= Request("FComponentLinkage")%>' />
        <input id="hidFComponentRootID" name="hidFComponentRootID" type="HIDDEN" value='<%= Request("FComponentRootID")%>' />
        <input id="hidFGPGDescription" name="hidFGPGDescription" type="HIDDEN" value='<%= Request("FGPGDescription")%>' />
        <input id="hidFMarketingDescriptionPMG" name="hidFMarketingDescriptionPMG" type="HIDDEN" value='<%= Request("FMarketingDescriptionPMG")%>' />
        <input id="hidFMarketingDescription" name="hidFMarketingDescription" type="HIDDEN" value='<%= Request("FMarketingDescription")%>' />
        <input id="hidFSCMCategoryID_singlefeature" name="hidFSCMCategoryID_singlefeature" type="HIDDEN" value='<%= Request("FSCMCategoryID_singlefeature")%>' />
        <input id="hidFAliasID" name="hidFAliasID" type="HIDDEN" value='<%= Request("AliasID")%>' />
        <input id="hidFAbbreviation" name="hidFAbbreviation" type="HIDDEN" value='<%= Request("Abbreviation")%>' />
        <input id="hidFPlatform" name="hidFPlatform" type="HIDDEN" value='<%= Request("Platform")%>' />
        <input id="hidConfigCode" name="hidConfigCode" type="hidden" value='<%=sConfigCode%>' />
        <input id="hidPulsarPlusDivId" name="hidPulsarPlusDivId" type="HIDDEN" value='<%= Request("pulsarplusDivId")%>' />
        <label id="lblDisplayedID" style="display: none"><%= request("AVID")%></label>
        <table width="100%" border="0">
            <tr>
                <td width="100%" align="right"><font size="1" face="verdana"><a href="#" onclick="ViewAvActionItems('<%= Request("AVID")%>')">AV Action Items</a></font></td>
            </tr>
        </table>

        <table class="FormTable" style="display: none;" bgcolor="cornsilk" width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="tan">
            <tr>
                <th>DCR:</th>
                <td>
                    <select id="selDCR" name="selDCR">
                        <option value="0">--- Please Make a Selection ---</option>
                        <% %>
                    </select></td>
            </tr>
        </table>

        <table class="FormTable" bgcolor="cornsilk" width="100%" border="1" cellspacing="0" cellpadding="1" bordercolor="tan">
            <tr>
                <th>Feature Full Name</th>

                <td>
                    <div id="featureNameDesc" style="display: none">
                        <table width="100%" border="0">
                            <tr>
                                <td>
                                    <input type="text" id="txtfeatureNameDesc" name="txtfeatureNameDesc" value="<%= sFeatureName%>" style="width: 300px" maxlength="40"></td>
                                <!--<td style="border: none; text-align: right; font-size: x-small"><a href="javascript:ShowSelectFeatureDialog()" id="selectfeature">Select Feature</a></td>-->
                            </tr>
                        </table>
                    </div>
                    <div id="featureNameText">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none">
                                    <div id="assignfeatureNameText"><%= PrepForWeb(sFeatureName)%></div>
                                </td>
                                <td style="border: none; text-align: right; font-size: x-small">
                                    <div id="divfeatureNameDesc">
                                        <%If sStatus <> "O" and sStatus <> "D" Then%>
                                            <!-- will not show hyperlink if av is shared av -->
                                            <%If bSharedAV = false Then%>
                                                <a href="javascript:ShowSelectFeatureDialog()" id="selectfeature1">                                           
                                                <%If sFeatureID = "" Then %>Select Feature
                                                <%Else%>Change Feature<%End If%>
                                                </a>
                                            <%End If%>
                                        <%End If%>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>



            </tr>
            <tr>
                <th>Feature ID</th>
                <td>
                    <div id="featureIDDesc" style="display: none">
                        <table width="100%" border="0">
                            <tr>
                                <td>
                                    <input type="text" id="txtfeatureIDDesc" name="txtfeatureIDDesc" value="<%= sFeatureID%>" style="width: 300px" maxlength="40">
                                    <input type="text" id="txtPrevFeatureID" name="txtPrevFeatureID" value="<%= sFeatureID%>">
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div id="featureIDText">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none">
                                    <div id="assignfeatureIDText"><%= PrepForWeb(sFeatureID)%></div>
                                </td>
                                <td style="border: none; text-align: right; font-size: x-small">
                                    <%If sFeatureID = "" or sFeatureID = "0"  or ISNULL(sFeatureID) Then%>
                                    <%Else%>
                                    <div id="divfeatureIDDesc"><a href="javascript:OpenFeatureProperties(<%=sFeatureID%>, 'SRP')" id="lnkShowFeature1">View Feature</a></div>
                                    <%End If%>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>

            </tr>
            <tr>
                <th>Shared AV?</th>
                <td>
                    <%	If ((LCase(sMode) = "add") or (LCase(sMode) = "clone")) Then %>
                    <input type="radio" name="rdSharedAV" id="rdNonSharedAv" value="0" checked="checked" onclick="return NonSharedAV_onclick()" />Non-Shared AV &nbsp&nbsp
            <input type="radio" name="rdSharedAV" id="rdSelectSharedAV" value="2" onclick="return OpenAddExistingSharedAV();" />Select existing Shared AV &nbsp&nbsp
            <input type="radio" name="rdSharedAV" id="rdSetUpSharedAv" value="1" onclick="return SharedAV_onclick()" />Set up New AV as Shared &nbsp&nbsp      
        <%else 'add new av or clone situation, will add radio buttons here later  %>
                    <%If bSharedAV Then %>
                Yes (SCM Category and product Line for Shared AVs may only be changed in Shared AV Admin)
		    <%else%>
                No            
            <%	End If %>
                    <%	End If %>   
          
                
                </td>

            </tr>
            <%	If LCase(sMode) = "add" Then %>
            <tr>
                <th>Brands</th>
                <td>
                    <%= sCbxBrand%>
                </td>
            </tr>
            <%	End If %>
            <tr>
                <th>SCM Category</th>
                <td>
                    <div id="featureCatText">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none" id="scmCategoryTxt"><%= PrepForWeb(sSCMCat)%></td>
                                <%If m_EditModeOn and not bSharedAV and sStatus <> "O" and sStatus <> "D" and sSCMParent = 0 Then%><td align="Right" style="border: none"><a href="javascript:EditFeatureCat();">Edit</a></td>  
                                <%End If%>
                            </tr>
                        </table>
                    </div>
                    <div id="featureCatSelect" style="display: none">
                        <div id="NameRowUpdate" style="display: none"></div>

                        <select style="display: none" id="cboNameFormat" name="cboNameFormat"><%=strNameFormats%></select>
                        <input id="tagCategory" name="tagCategory" type="hidden" value="<%=trim(strCatID)%>">
                        <select id="cboCategory" name="cboCategory" style="opacity:.2; background-color: transparent; width: 80%" language="javascript" onchange="return cboCategory_onchange(<%=strElementValues%>)">
                            <option value="0">--- Please Make a Selection ---</option>
                            <%=sCategoryOpt%>
                        </select>
                    </div>
                </td>
            </tr>
            
            <tr id="trPlatform" style="display:none">
                <th>Base Unit Group</th>
                <td>
                    <div id="divPlatform" style="margin-left:.5em; margin-left:5px"><%=PrepForWeb(sPlatformName)%></div>
                </td>
            </tr>
            
            <tr>
                <th>Product&nbsp;Line:</th>
                <td>
                    <div id="ProductLineSelect" style="display: none">
                        <select id="cboProductLine" name="cboProductLine" style="width: 200px;">
                            <option></option>
                            <%=strProductLines%>
                        </select>&nbsp;
                    </div>
                    <div id="ProductLineEdit">
                        <table width="100%" border="0">
                            <tr>
                                <td id="tdProductLine" style="border: none"><%= PrepForWeb(strProductLineName)%></td>
                                <%If m_EditModeOn and not bSharedAV and sStatus <> "O" and sStatus <> "D" Then%>
                                <td align="right" style="border: none">
                                    <div id="divProdLine"><a href="javascript:EditProductLine();">Edit</a></div>
                                </td>
                                <%End If%>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <th>AV#</th>
                <td>
                    <div id="avinput" style="display: none">
                        <input type="text" id="AvNo" name="AvNo" value="<%= sAvNo%>" maxlength="18" onkeypress="return onlyAlphaNumbers(event);">
                    </div>
                    <div id="avtext">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sAvNo)%></td>
                                <%If m_EditModeOn and sStatus <> "O" and sStatus <> "D" Then%><td align="Right" style="border: none"><a href="javascript:EditAv();">Edit</a></td>
                                <%End If%>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <th>Status</th>
                <td>
                    <div id="avStatus">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sStatus)%></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <th>Root ID</th>
                <%	If iViaAvCreate = "1" Then %>
                <td><%= PrepForWeb(iDeliverableRootID)%>&nbsp;&nbsp;<font size="1">(Cannot Edit - Automatically Assigned)</font></td>
                <%	Else %>
                <td>
                    <div id="delRootCbo" style="display: none">
                        <select id="cboDeliverables" name="cboDeliverables" style="width: 70%">
                            <option value="0">--- Please Make a Selection ---</option>
                            <%=sDeliverableOpt%>
                        </select>
                    </div>
                    <div id="delRootText">
                        <%Dim iDelRootID
			   If iDeliverableRootID = "0" Then iDelRootID = "" Else iDelRootID = PrepForWeb(iDeliverableRootID)
                        %>
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= iDelRootID%></td>
                            </tr>
                        </table>
                    </div>
                </td>
                <%	End If %>
            </tr>



            <tr>
                <th>Originated By DCR</th>
                <%	If bOriginatedByDCR = "True" Then %>
                <td>
                    <input type="checkbox" id="chkDCR" name="chkDCR" checked style="width: 16; height: 16" onclick="return chkDCR_onclick()"><div id="divDVR" style="display: inline">&nbsp;&nbsp;&nbsp;DCR Number:&nbsp;<input type="text" style="width: 100px" id="txtDCRNo" name="txtDCRNo" value="<%= iDCRNo%>" maxlength="9"></div>
                </td>
                <%	Else %>
                <td>
                    <input type="checkbox" id="chkDCR" name="chkDCR" style="width: 16; height: 16" onclick="return chkDCR_onclick()"><div id="divDVR" style="display: none;">&nbsp;&nbsp;&nbsp;DCR Number:&nbsp;<input type="text" style="width: 100px" id="txtDCRNo" name="txtDCRNo" value="<%= iDCRNo%>" maxlength="9"></div>
                </td>
                <%	End If %>
            </tr>


            <tr id="trGPGDesc" style="">
                <th>GPG Description</th>
                <td>
                    <div id="gpgDesc" style="display: none">
                        <input type="text" id="txtAvGpgDescription" name="txtAvGpgDescription" value="<%= sGpgDesc%>" style="width: 300px" maxlength="50" />
                    </div>
                    <div id="gpgText">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sGpgDesc)%></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr id="trMarketingDesc40" style="">
                <th>Marketing Short Description<br />
                    (40 Char)</th>
                <td>
                    <div id="mktgDesc" style="display: none">
                        <input type="text" id="txtMarketingDesc" name="txtMarketingDesc" value="<%= sMarketingDesc%>" style="width: 300px" maxlength="40" />
                    </div>
                    <div id="mktgText">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sMarketingDesc)%></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr id="trMarketingDesc100" style="">
                <th>Marketing Long Description<br />
                    (100 Char)</th>
                <td>
                    <div id="mktgDescPMG" style="display: none">
                        <input type="text" id="txtMarketingDescPMG" name="txtMarketingDescPMG" value="<%= sMarketingDescPMG%>" style="width: 300px" maxlength="100" />
                    </div>
                    <div id="mktgTextPMG">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sMarketingDescPMG)%></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <!--<tr>
                <th>Program Version</th>
                <td><%= sProgramVersion%></td>
            </tr>-->


            <!--task 16227 Releases--->
            <tr>
		        <th>Release(s)</th>
            	<td><div id="ProductRelease" style="display:none">
                            <%                                 
                                dim strName
                                rs.open "usp_Get_ProductRelease " & clng(request("PVID")) & "," & clng(request("AVID")) & "," & clng(Trim(Request("BID"))), cn, adOpenForwardOnly
                                do while not rs.eof
                                    strname = rs("Name")   

                                    if trim(rs("OnAV")) = "1" or InStr(sProductReleaseIDs , "," & trim(rs("ReleaseID")) & ",") > 0 then
                                        response.write "<input checked id=""chkRelease"" RTPDate=""" & rs("RTPDate") & """ EMDate=""" & rs("EMDate") & """ onclick=""getScheduleDates();"" name=""chkRelease"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """  > " & strname & "&nbsp;"
                                        strLoaded = strLoaded & "," & rs("ReleaseID")
                                    else
                                        response.write "<input id=""chkRelease"" RTPDate=""" & rs("RTPDate") & """ EMDate=""" & rs("EMDate") & """ onclick=""getScheduleDates();"" name=""chkRelease"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """   > " & strname & "&nbsp;"
                                    end if

                                    rs.movenext
                                loop

                                if trim(strLoaded) <> "" then
                                    strLoaded = mid(strLoaded,2)
                                end if                              
                                rs.close                            
                             %>
            	    </div>
                <div id="ProductReleaseEdit"><table width="100%" border="0"><tr><td style="border:none"><%= PrepForWeb(sProductRelease)%></td>
                <td align=Right style="border:none"><a href="javascript:EditProductRelease();">Edit</a></td>
                </tr></table></div></td>
            </tr>           
            <tr>
		        <th>RTP/MR Date</th>
		        <td><div id="mktgRTPDate" style="display:none"><input type="text" id="txtRTPDate" name="txtRTPDate" value="<%= sRTPDt%>" style="width:300px" autocomplete='off' /></div>
                    <div id="mktgRTPDateText"><table width="100%" border="0"><tr><td id="RTP1" style="border:none"><%= PrepForWeb(sRTPDt)%></td>
                        <%If not bSharedAV and m_CanEditDates = "True" Then%>
                        <td align=Right style="border:none"><a href="javascript:EditMktgRTPDate(false, '');">Edit</a></td>
                        <% end if %>
                        </tr></table></div></td>
	        </tr>
	        <tr>
		        <th>PA:AD (Intro Date)<sup style="color: green;"> [1]</sup></th>
		        <td>
                    <div id="mktgPAADDate" style="display:none"><input type="text" id="txtPAADDate1" name="txtPAADDate1" value="<%= sPAADDate%>" style="width:300px" autocomplete='off' /></div>
                    <div id="mktgPAADDateText"><table width="100%" border="0"><tr><td id="PAAD1" style="border:none"><%= PrepForWeb(sPAADDate)%></td>
                        <%If not bSharedAV and m_CanEditDates = "True" Then%>
                        <td align=Right style="border:none"><a href="javascript:EditMktgPAADDate1();">Edit</a></td>
                        <% end if %>
                        </tr></table></div></td>
	        </tr>

	        <tr>
		        <th>Select Availability (SA) Date<sup style="color: green;"> [1]</sup></th>
                <td>
                <div id="mktgBlindDate" style="display:none"><input type="text" id="txtBlindDate1" name="txtBlindDate1" value="<%= sCplBlindDt%>" style="width:300px" autocomplete='off' /></div>
		        <div id="mktgBlindDateText"><table width="100%" border="0"><tr><td id="SA1" style="border:none"><%= PrepForWeb(sCplBlindDt)%></td>
                    <%If not bSharedAV and m_CanEditDates = "True" Then%>
                    <td align=Right style="border:none"><a href="javascript:EditMktgBlindDate1();">Edit</a></td>
                    <% end if %>
		            </tr></table></div></td>
	        </tr>

	        <tr>
		        <th>General Availability (GA) Date<sup style="color: green;"> [1]</sup></th>
                <td>
                <div id="mktgGeneralAvailDt" style="display:none"><input type="text" id="txtGeneralAvailDt1" name="txtGeneralAvailDt1" value="<%= sGeneralAvailDt%>" style="width:300px" autocomplete='off' /></div>
		        <div id="mktgGeneralAvailDtText"><table width="100%" border="0"><tr><td id="GA1" style="border:none"><%= PrepForWeb(sGeneralAvailDt)%></td>
                    <%If not bSharedAV and m_CanEditDates = "True" Then%>
                    <td align=Right style="border:none"><a href="javascript:EditMktgGeneralAvailDt1();">Edit</a></td>
                    <% end if %>
                    </tr></table></div></td>
	        </tr>

            <tr>
		        <th>End of Manufacturing (EM) Date</th>
		        <td><div id="mktgDiscDate" style="display:none"><input type="text" id="txtMarketingDiscDate" name="txtMarketingDiscDate" value="<%= sRasDiscDt%>" style="width:300px" autocomplete='off' /></div>
                    <div id="mktgDiscDateText"><table width="100%" border="0"><tr><td id="EOM1" style="border:none"><%= PrepForWeb(sRasDiscDt)%></td>
                        <%If not bSharedAV and m_CanEditDates = "True" Then%>
                        <td align=Right style="border:none"><a href="javascript:EditMktgDiscDate(false, '');">Edit</a></td>
                        <% end if %>
                        </tr></table></div></td>
	        </tr>  
            <tr>
                <th>Global Series Config<br />
                        Planned End 
                <td>
                    <div id="GSEndDtInput" style="display: none">
                        <input type="text" id="txtGSEndDt" name="txtGSEndDt" value="<%= sGSEndDt%>">
                    </div>
                    <div id="GSEndDtTxt">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sGSEndDt)%></td>
                                <%If m_EditModeOn and sStatus <> "O" and sStatus <> "D" Then%><td align="Right" style="border: none"><a href="javascript:EditGSEndDt();">Edit</a></td>
                                <%End If%>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <!-- Do not need this PHweb Instructions any more for pulsar
            -->
            <tr>
                <th>SDF Flag</th>
                <%If sSDFFlag = "False" or sSDFFlag = "" or ISNULL(sSDFFlag) Then %>
                <td>
                    <input type="checkbox" id="chkSDFFlag" name="chkSDFFlag" style="width: 16; height: 16" onclick="return chkSDFFlag_onclick()"></td>
                <%Else%>
                <td>
                    <input type="checkbox" id="chkSDFFlag" name="chkSDFFlag" checked style="width: 16; height: 16" onclick="return chkSDFFlag_onclick()"></td>
                <%End If%>
            </tr>
            <tr>
                <th>PRL Offering Constraints</th>
                <td><div  id="txtPRLOffConstraints" style="height:120px;overflow:auto;"><%= sPRLOffConstraints%></div></td>
            </tr>
            <tr>
                <th>Configuration Rules</th>
                <td>
                    <textarea rows="5" id="txtConfigRules" name="txtConfigRules" style="width: 300px"><%= sConfigRules%></textarea></td>
            </tr>
            <tr>
                <th>Rules Syntax</th>
                <td>
                    <textarea rows="5" id="txtRulesSyntax" name="txtRulesSyntax" maxlength="512" style="width: 300px; text-transform: uppercase;"
                        onkeydown="textCounter(this.form.txtRulesSyntax, document.getElementById('tLen2'), 512);"
                        onkeyup="textCounter(this.form.txtRulesSyntax, document.getElementById('tLen2'), 512);"
                        onchange="this.value = this.value.toUpperCase();"><%= sRulesSyntax%></textarea>
                    &nbsp; <span class="Label">Remaining characters: </span>
                    <span class="LabelHeader"><span id="tLen2"><% if len(sRulesSyntax)>0 then response.write 512-len(sRulesSyntax) else response.write "512" %></span></span>
                </td>
            </tr>
            <tr>
                <th>AVID</th>
                <td>
                    <textarea rows="5" id="txtAvId" name="txtAvId" style="width: 300px"><%= sAvId%></textarea></td>
            </tr>
            <tr>
                <th>Group 1</th>
                <td>
                    <textarea rows="5" id="txtGroup1" name="txtGroup1" style="width: 300px"><%= sGroup1%></textarea></td>
            </tr>
            <tr>
                <th>Group 2</th>
                <td>
                    <textarea rows="5" id="txtGroup2" name="txtGroup2" style="width: 300px"><%= sGroup2%></textarea></td>
            </tr>
            <tr>
                <th>Group 3</th>
                <td>
                    <textarea rows="5" id="txtGroup3" name="txtGroup3" style="width: 300px"><%= sGroup3%></textarea></td>
            </tr>
            <tr>
                <th>Group 4</th>
                <td>
                    <textarea rows="5" id="txtGroup4" name="txtGroup4" style="width: 300px"><%= sGroup4%></textarea></td>
            </tr>
            <%'if sIsDesktop=False Then 'notebooks %>
            <tr>
                <th>Group 5</th>
                <td>
                    <textarea rows="5" id="txtGroup5" name="txtGroup5" style="width: 300px"><%= sGroup5%></textarea></td>
            </tr>
            <%'End If%>
            <tr>
                <th>Group 6</th>
                <td>
                    <textarea rows="5" id="txtGroup6" name="txtGroup6" style="width: 300px"><%= sGroup6%></textarea></td>
            </tr>
            <tr>
                <th>Group 7</th>
                <td>
                    <textarea rows="5" id="txtGroup7" name="txtGroup7" style="width: 300px"><%= sGroup7%></textarea></td>
            </tr>
            <tr>
                <th>Manufacturing Notes</th>
                <td>
                    <textarea rows="5" id="txtManufacturingNotes" name="txtManufacturingNotes" style="width: 300px"><%= sManufacturingNotes%></textarea></td>
            </tr>
            <tr>
                <th>Weight</th>
                <td>
                    <div id="WeightInput" style="display: none">
                        <input type="text" id="txtWeight" name="txtWeight" maxlength="9" value="<%= sWeight%>">
                    </div>
                    <div id="WeightTxt">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sWeight)%></td>
                                <%If m_EditModeOn and sStatus <> "O" and sStatus <> "D" Then%>
                                <td align="Right" style="border: none"><a href="javascript:EditWeight();">Edit</a></td>
                                <%End If%>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <!--            <tr>
                <th>Global Series Config<br />
                    Plan for End of<br />
                    Manufacturing<br />
                    (PE) Date</th>
                <td>
                    <div id="GSEndDtInput" style="display: none">
                        <input type="text" id="txtGSEndDt" name="txtGSEndDt" value="<%= sGSEndDt%>"></div>
                    <div id="GSEndDtTxt">
                        <table width="100%" border="0">
                            <tr>
                                <td style="border: none"><%= PrepForWeb(sGSEndDt)%></td>
                                <%If m_EditModeOn and sStatus <> "O" and sStatus <> "D" Then%><td align="Right" style="border: none"><a href="javascript:EditGSEndDt();">Edit</a></td>
                                <%End If%></tr>
                </table>
            </div>
                </td>
            </tr>-->
            <%'if sIsDesktop=False Then 'notebooks %>
            <tr>
                <th>IDS</th>
                <td>
                    <input type="checkbox" id="chkIdsSkus" name="chkIdsSkus" <%IF bIdsSkus Then%>CHECKED<%End If%>></td>
            </tr>
            <tr>
                <th>IDS-CTO</th>
                <td>
                    <input type="checkbox" id="chkIdsCto" name="chkIdsCto" <%IF bIdsCto Then%>CHECKED<%End If%>></td>
            </tr>
            <tr>
                <th>RCTO</th>
                <td>
                    <input type="checkbox" id="chkRctoSkus" name="chkRctoSkus" <%IF bRctoSkus Then%>CHECKED<%End If%>></td>
            </tr>
            <tr>
                <th>RCTO-CTO</th>
                <td>
                    <input type="checkbox" id="chkRctoCto" name="chkRctoCto" <%IF bRctoCto Then%>CHECKED<%End If%>></td>
            </tr>
            <%'end if %>

            <% If (sProdVersionBSAMFlag = "True") Then %>
            <!--<tr>
                <th>BSAM SKUS</th>
                <td>
                    <input type="checkbox" id="Checkbox1" name="chkBSAMSkus" <%IF bBSAMSkus Then%>CHECKED<%End If%>></td>
            </tr>-->
            <tr>
                <th>BSAM -B</th>
                <td>
                    <input type="checkbox" id="chkBSAMBparts" name="chkBSAMBparts" <%IF bBSAMBparts Then%>CHECKED<%End If%>></td>
            </tr>
            <% End If %>
            <tr>
                <th>UPC</th>
                <td><%= sUpc%></td>
            </tr>
            <%	If LCase(sMode) <> "add" Then %>
            <tr>
                <th>Reason for Change</th>
                <td>
                    <textarea rows="2" id="txtChangeReason" name="txtChangeReason" style="width: 300px" maxlength="300"><%= sChangeNote%></textarea></td>
            </tr>
            <%	End If %>
            <tr>
                <th>Comments</th>
                <td>
                    <textarea rows="4" id="txtComments" name="txtComments" style="width: 300px" maxlength="500"><%= sComments%></textarea>
                </td>
            </tr>
            <tr>
                <th>Sorting Weight</th>
                <td>
                    <input type="text" id="txtSortOrder" name="txtSortOrder" size="3" maxlength="4" value="<%= sSortOrder%>"></td>
            </tr>
            <tr>
                <th>Show Change on SCM</th>
                <td>
                    <input type="checkbox" id="chkShowOnScm" name="chkShowOnScm" checked="CHECKED"></td>
            </tr>
            <%If IsNumeric(Request("AVID")) and Request("AVID") <> "" Then%>
            <tr>
                <th>Base Unit Availability</th>
                <td>
                    <div>
                        <table width="100%" border="0" id="BUAvail">
                            <thead>
                                <tr>
                                    <td>Base Unit Group</td>
                                    <td>Base Unit</td>
                                    <td>Status (click on status to change)</td>
                                </tr>
                            </thead>
                            <%
                                ' 07/26/2016 - ADao - Change IRS_Platform_Alias, IRS_Platform, IRS_Alias synonyms to use actual tables 
                                rs.Open "select GPGDescription as BUDescription, isnull(a.FeatureID, 0) as BUFeatureID, aa.Availstatus, [BUStatus] = apb.Status, [AVStatus] = apb1.Status, GenericName = pl.MarketingName " &_
                                        "from Product_Brand pb " &_
                                        "join AvDetail_ProductBrand apb WITH (NOLOCK) on apb.ProductBrandID = pb.ID " &_
                                        "join AvDetail a WITH (NOLOCK) on a.AvDetailID = apb.AvDetailID " &_
                                        "join SCMCategory c WITH (NOLOCK) on c.ScmCategoryID = a.ScmCategoryID " &_
                                        "join Feature f WITH (NOLOCK) on f.FeatureID = a.featureid " &_
                                        "join Platform_Alias PA WITH (NOLOCK) ON PA.AliasID = F.AliasID " &_
                                        "join Platform pl with (nolock) on pa.platformid = pl.platformid " &_
                                        "join avdetail_avail aa with (nolock) on aa.BUFeatureID = a.FeatureID and aa.productbrandid = apb.ProductBrandID " &_
                                        "join AvDetail_ProductBrand apb1 WITH (NOLOCK) on apb1.AvDetailID = aa.AvDetailID and apb1.ProductBrandID = aa.ProductBrandID " &_
                                        "where pb.ProductVersionID = " & Request("PVID") & " and pb.ID = " & Request("BID") &_
                                        "and (c.abbreviation = 'BUNIT' or c.Name like 'Base Unit') AND apb.Status not in ('K', 'D', 'O') " &_
                                        "and aa.avdetailid = " & Trim(Request("AVID")), cn, adOpenForwardOnly
	                            do while not rs.EOF 
                                    If m_EditModeOn and rs("BUStatus") <> "D" and rs("BUStatus") <> "O" and rs("AVStatus") <> "D" and rs("AVStatus") <> "O" Then                 
                                        Response.Write "<tr ><td>" & rs("GenericName") & "</td><td>" & rs("BUDescription") & "</td><td id='BUAvail-" & rs("BUFeatureID") & "' OldStatus='" & rs("Availstatus") & "'><a href='javascript:void(0)' onclick='ChangeStatus(" & rs("BUFeatureID") & ")' >" & rs("Availstatus") & "</a></td></tr>"		                           
                                    ELSE
                                        Response.Write "<tr ><td>" & rs("GenericName") & "</td><td>" & rs("BUDescription") & "</td><td id='BUAvail-" & rs("BUFeatureID") & "' OldStatus='" & rs("Availstatus") & "'>" & rs("Availstatus") & "</td></tr>"
                                    End IF
                                        
                                    rs.MoveNext
                                    
	                            loop
	                            rs.Close                                
                            %>
                        </table>
                    </div>
                </td>
            </tr>
            <%End If%>
            <%'if sIsDesktop and bBaseParent then 'Only display if it is a Base Parent AV (of localized AVs)
                'Since we are no longer showing base parents, this code should never be run (PBI 8363) but we'll comment out just in case %>
            <%' <tr>
              '  <th>Demand Region</th>
              '  <td>
              '      <select id="cboDemandRegion" name="cboDemandRegion"> %>
                        <%
						'if trim(strDemandRegionID) = "0" then
			    		'	Response.Write "<option value=""0""></option>"
						'end if
						'rs.Open "usp_GetDemandRegion",cn,adOpenForwardOnly
						'rs.Sort = "DemandRegion ASC"
						'do while not rs.EOF
						'	if trim(strDemandRegionID) = trim(rs("DemandRegionID") & "") then
						'		Response.Write "<option selected value=" & rs("DemandRegionID") & ">" & rs("DemandRegion") & "</option>"
						'	else
						'		Response.Write "<option value=" & rs("DemandRegionID") & ">" & rs("DemandRegion") & "</option>"
						'	end if
						'	rs.Movenext
						'loop
						'rs.Close
                        %>
                    <% '</select>
                '</td>
            '</tr>
            %>
            <% 'end if 
               if not Request("AVID") = "" then
            %>
            <tr>
                <th>Created</th>
                <td><%= PrepForWeb(sCreated)%></td>
            </tr>
            <tr>
                <th>Created By</th>
                <td><%= PrepForWeb(sCreatedBy)%></td>
            </tr>
            <tr>
                <th>Updated</th>
                <td><%= PrepForWeb(sUpdated)%></td>
            </tr>
            <tr>
                <th>Updated By</th>
                <td><%= PrepForWeb(sUpdatedBy)%></td>
            </tr>
            <% end if %>
        </table>
        <!--task 16227 - Releases-->
        <input id="txtReleaseLoaded" name="txtReleaseLoaded" type="hidden" value="<%=strLoaded%>">
        <input id="txtReleaseList" name="txtReleaseList" type="hidden" value="">
        <input id="txtReleaseAdded" name="txtReleaseAdded" type="hidden" value="<%=sProductRelease%>">
    </form>

    <div style="display: none;">
        <div id="iframeDialog" title="Coolbeans">
            <iframe style="border: none; width: 100%; height: 100%" name="modalDialog" id="modalDialog"></iframe>
        </div>
    </div>

    <div id="Promptdialog" title="SCM Prompt" style="display: none;">
        <p>Would you like to change current SCM Category to the new selected feature's SCM Category?</p>
    </div>

    <div style="font-size:xx-small;color: green;font-style: italic;"><p></p>1. PA:AD (Intro Date), Selected Availability (SA) Date, and General Availability (GA) Date are calculated based on the RTP/MR date and automatically updated</div>

    <div id="dialog" title="Confirmation" ></div>
     <div id="divOpenFeaturePopUp" title="Coolbeans" style="display: none;">
        <iframe frameborder="0" name="ifOpenFeaturePopUp" id="ifOpenFeaturePopUp" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
    <div id="divMultipleAVs" title="Coolbeans" style="display: none;">
        <iframe frameborder="0" name="ifMultipleAVs" id="ifMultipleAVs" style="height: 100%; width: 100%" marginheight="0" marginwidth="0"></iframe>
    </div>
</body>
</html>