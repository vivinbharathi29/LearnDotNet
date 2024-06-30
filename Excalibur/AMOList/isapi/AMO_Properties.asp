 <%@ Language=VBScript %>
<% Option Explicit %>
<%response.buffer = false %>

<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/Groups.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->
<!-- #include file="../library/includes/DualListboxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../includes/AMO.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<%
'printrequest
Call ValidateSession 

dim sDesc, sHeader, sModuleTypeHTML, sModuleDivisionHTML, sCtrlStyle, sHelpFile, strCheckDivisionIDs
dim sCreator, sUpdater, sCreatedDate, sUpdatedDate, sErr, sCategoryDesc, sAMOStatus, sGroupByDivision
dim sRasDisconDate, sCPLBlindDate, sBOMRevADate, sBluePN, sRedPN, sPlUserDivisions
dim sNetWeight, sExportWeight, sAirPackedWeight, sAirPackedCubic, sExportCubic, sReplacement, sAlternative
dim sDivisionIDs, sShortDesc, sModuleType, sActualCost, sHWModuleCategoryHTML, sSWModuleCategoryHTML
dim sGroupName, sTmp, sDivisionID, strFrom, strOwnerHTML, strFilter, strDiv, strNotes, sCostCtrlStyle, sRasCtrlStyle
dim sOriginalDivisionIDs, sDivisionTarget, sGroupId, strRuleDescription
dim nID, nMode, nTypeID, nModuleTypeID, nAMOStatusID, sProductLineHTML
dim lngChangeMask, lngGroupID, lngMOLHide, bIDP, strCloneRegionIds, strLongDescription, strReplacementDescription, strOrderInstructions
dim oSvr, oErr, sRasObsoleteDate
dim oRs, oRsPlatforms, oSelectedRs, oRsCheckedRegion, oRsGroups, oRsHWCategory, oRsSWCategory, oRsUserRights, oRsCreateGroups
dim bUpdate, bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim bCostCreate, bCostView, bCostUpdate, bCostDelete, bRasCreate, bRasView, bRasUpdate, bRasDelete
dim bFromAMO, bEdit
dim arrDivisions
dim sAMOPartNoRe, intTargetNA, intTargetLA, intTargetEMEA, intTargetAPJ, intTargetAllReg, intBurdenPer
dim intContraPer, intBurden, intContra, intNetRevenue, intDealerMargin, intTargetCostBurden, intGrossMarginPer
dim intGrossMargin, intLifetimeGrossMargin, sJustificationnotes
dim bPORCreate, bPORView, bPORUpdate, bPORDelete, nUsedInMOL, sNameComment, sNameCtrlStyle
dim sVisibility_NA, sVisibility_EM, sVisibility_AP, sVisibility_LA
dim sRuleID, sRuleIDCtrlStyle, bAVLKECreate, bAVLKEView, bAVLKEUpdate, bAVLKEDelete
dim sCloneGroupName, sUserDivisions, bSave, oRsProductLine, nProductLineId
dim sAMOCost, intCount, bEditPlatform, sDisabledCtrlStyle, sHubCheckboxlist, sHideHubCheckboxlist
dim sAMOWWPrice, sOtherSelectedAliasIDs, oRsBusSeg, oRsBusSegSelected, nBusSelected
dim sManufactureCountry, sWarrantyCode, sObsoleteDate, bBusEnabled, sComparatibilitySelected, sRasDisCtrlStyle
dim lngSCMHide, lngSCLHide
dim nID_output
dim sNewCPLLocaleString, sNewBOMLocaleString, sNewRASLocaleString, sNewOBSLocaleString, sLocalized
dim sProductline_Ori, sDisabledIDP, strKeyWord

sNewCPLLocaleString = ""
sNewBOMLocaleString = ""
sNewRASLocaleString = ""
sNewOBSLocaleString = ""
sHubCheckboxlist = ""
sProductLineHTML = ""
sHideHubCheckboxlist = ""

sCtrlStyle = ""
strCloneRegionIds = ""
sUserDivisions = ""	
sDisabledCtrlStyle = ""
sDisabledIDP = ""
sComparatibilitySelected = ""
sLocalized = 0
bSave = False

sHelpFile = "../help/HELP_AMO_Properties.asp"
strOwnerHTML = ""
sErr = ""
bUpdate = False

if Request.QueryString("TreePath") <> "" then	'instr(lcase(Request.ServerVariables("HTTP_REFERER")), "deeptree.asp") > 0 and 
	'we came from the tree so save the cookie of this module
	Call SaveDBCookie( "Modules tree_obout", Request.QueryString("TreePath"))
end if

if Request.QueryString("Mode") = "" then
	nMode = 1	'create
else
	nMode = clng(Request.QueryString("Mode"))
end if

if Request.QueryString("ID") <> "" then
	nID = clng(Request.QueryString("ID"))
else
	nID = 0
end if

strFrom = Request.QueryString("from")
if Request.QueryString("from") = "1" then
	'need to know if this is a popup window from AMO or not
	bFromAMO = True
else
	bFromAMO = False
end if

'set rsRoles and IRSUserID Session: ----
'Call SetPermission()

set oRsCreateGroups = GetGroupsForRole2(cstr(Application("AMOLIST")), true, false, false, false, true, false)
if (oRsCreateGroups is nothing) then
    Response.Write("Empty Recordset: oRsCreateGroups")
    Response.End()
else
	if nMode = 3 then
		if oRsCreateGroups.RecordCount > 1 then
			'more than one group so we have to be coming from the select user group page
			lngGroupID = Request.Querystring("nGroupID")
			sDivisionIDs = ""
			if oRsCreateGroups.RecordCount > 0 then
				if Len(lngGroupID) > 0 then
					oRsCreateGroups.Filter = "GroupID = " & cstr(lngGroupID)
				end if
				if oRsCreateGroups.RecordCount > 0  then
					sDivisionIDs = replace(oRsCreateGroups("DivisionIDs"), "|", ",")
				end if
			end if
		else
			'use the only groupid and divisionids		
			lngGroupID = oRsCreateGroups("GroupID").value
			sCloneGroupName = oRsCreateGroups("GroupName").value
			sDivisionIDs = replace(oRsCreateGroups("DivisionIDs"), "|", ",")
		end if
	end if
	
	For intCount = 0 To oRsCreateGroups.RecordCount-1
		sUserDivisions = sUserDivisions & oRsCreateGroups("DivisionIDs")
		oRsCreateGroups.MoveNext		
	Next
end if


''clone from CTO-srp module
'moved the following line here since it will be shared in later codes
 'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO
	
if strFrom = "CTO" then
	if sErr = "" then
		sErr = oSvr.AMO_ClonefromCTO(Application("REPOSITORY"), nID, session("FullName"),  Session("AMOUserID"))
		if IsNumeric(sErr) = False then
			sErr = "Missing required parameters.  Unable to complete your request."
		    Response.Write(sErr)
		    Response.End()
		else
            nID_output = sErr
			if nID_output > 0 then
				nID = nID_output
			end if 
			
		end if

	end if

end if 

set oRsBusSeg = Nothing
set oRsBusSegSelected = Nothing
set oRsProductLine = Nothing

nBusSelected = ""
'set oErr = GetMOLCategory(oRsBusSeg, 28)
set oRsBusSeg = GetMOLCategory(34)	
if oRsBusSeg is Nothing then
	Response.Write("Recordset error: oRsBusSeg")
	Response.End()
end if

''end of clone from CTO-srp module

if sErr = "" then
	select case nMode
		'1=create, 2=modify, 3=clone
		case 1
			'this is coming from the New page where the user enters the Owner group
			sGroupName = Request.QueryString("sGroupName")
			lngChangeMask = 0
			'get this from the New form
			lngGroupID = Request.Querystring("nGroupID")
			if lngGroupID = "" then
				sErr = "Missing required parameters.  Unable to complete your request."
		        Response.Write(sErr)
		        Response.End()
			end if
			'Determine default target business segment from Owner Group
			sDivisionIDs = ""
			if oRsCreateGroups.RecordCount > 0 then
				oRsCreateGroups.Filter = "GroupID = " & cstr(lngGroupID)
				if oRsCreateGroups.RecordCount > 0 then
					sDivisionIDs = replace(oRsCreateGroups("DivisionIDs"), "|", ",")
				end if
			end if
			
			if sErr = "" then
				if right(sDivisionIDs, 1) = "," then
					strCheckDivisionIDs = left(sDivisionIDs, len(sDivisionIDs)-1)
				else
					strCheckDivisionIDs = sDivisionIDs
				end if
				
			
				set oRsUserRights = GetAllRolesByDivision( Session("AMOUserID"), strCheckDivisionIDs, oRsUserRights)
				if oRsUserRights is nothing then
					Response.Write("Empty Recordset: adoRs")
		            Response.End()
				else
					GetRightsByRecordset Application("AMOList"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, oRsUserRights
					GetRightsByRecordset Application("AMOCost"), bCostCreate, bCostView, bCostUpdate, bCostDelete, oRsUserRights
					GetRightsByRecordset Application("AMOPOR"), bPORCreate, bPORView, bPORUpdate, bPORDelete, oRsUserRights	
								
					oRsUserRights.Close
	
					sHeader = "Create After Market Option"
					if bAMOCreate then
						bUpdate = True
					end if
					bEdit = True
					sAMOStatus = "New"
					bIDP = false
					lngMOLHide = 0
					lngSCMHide = 0
					lngSCLHide = 0
					set oSelectedRs = nothing
				end if
				set oRsUserRights = nothing
			end if
	
		case 2, 3, 4	'2=edit, 3=clone, 4=view only, 5=clone from CTO module
		'	set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
			'get option data
			
			strKeyWord = ""
			set oRs = oSvr.AMOModule_Search(Application("REPOSITORY"), " and O.ModuleID=" & cstr(nID) & " and O.SCMID=1 and R.SCMID = 1 and R.RegionID = "&Application("AMO_GLOBAL_REGIONID")&"",strKeyWord, null, null)
			if oRs is nothing then
				sErr = "Missing required parameters.  Unable to complete your request."
		        Response.Write(sErr)
		        Response.End()
			else
				if oRs.RecordCount = 0 then
                    sErr = "The After Market Option was not found in the system., AMO_Properties.asp"
		            Response.Write(sErr)
		            Response.End()
				else
					'fields returned but not used: Category, RAS, GPSy, RegionComment

					if nMode = 3 then
						nAMOStatusID = Application("AMO_NEW")
					else
						nAMOStatusID = oRs.Fields("AMOStatusID").Value
					end if
							
					if nAMOStatusID = Application("AMO_COMPLETE") and Request.QueryString("Edit") = "" then
						bEdit = False
					else
						bEdit= True
					end if
										
					nTypeID = oRs.Fields("CategoryID").Value
					sDesc = oRs.Fields("Description").Value
					strLongDescription = oRs.Fields("LongDescription").Value
					sShortDesc = oRs.Fields("ShortDescription").Value
					sCreator = oRs.Fields("Creator").Value
					sUpdater = oRs.Fields("Updater").Value
					sCreatedDate = oRs.Fields("TimeCreated").Value
					sUpdatedDate = oRs.Fields("TimeChanged").Value
					nModuleTypeID = oRs.Fields("ModuleTypeID").Value
					sModuleType = oRs.Fields("ModuleType").Value
					sAMOStatus = oRs.Fields("AMOStatus").Value
					sGroupName = oRs.Fields("GroupName").Value
					sGroupId = oRs.Fields("GroupId").Value
					
					if nMode = 3 then
						sBluePN = ""
					else
						sBluePN = oRs.Fields("BluePN").Value
					end if
					
					sRedPN = oRs.Fields("RedPN").Value
					sBOMRevADate = oRs.Fields("BOMRevADate").Value
					sRasDisconDate = oRs.Fields("RASDiscontinueDate").Value
					sRasObsoleteDate = ""				
					sCPLBlindDate = oRs.Fields("CPLBlindDate").Value
					sAMOCost = oRs.Fields("AMOCost").Value
					sAMOWWPrice = oRs.Fields("AMOWWPrice").Value
					sActualCost = oRs.Fields("ActualCost").Value
					sNetWeight = oRs.Fields("NetWeight").Value
					sAirPackedWeight = oRs.Fields("AirPackedWeight").Value
					sExportWeight = oRs.Fields("ExportWeight").Value
					sAirPackedCubic = oRs.Fields("AirPackedCubic").Value
					sExportCubic = oRs.Fields("ExportCubic").Value
					sRuleID = oRs.Fields("RuleID").Value
					sReplacement = oRs.Fields("Replacement").Value
					sAlternative = oRs.Fields("Alternative").Value
					sCategoryDesc = oRs.Fields("CategoryDesc").Value
					lngChangeMask = oRs.Fields("ChangeMask").Value
					
					sAMOPartNoRe = oRs.Fields("AMOPN_Replacement").Value
					intTargetNA = oRs.Fields("TargetVolumn_NA").Value
					intTargetLA = oRs.Fields("TargetVolumn_LA").Value
					intTargetEMEA = oRs.Fields("TargetVolumn_EM").Value
					intTargetAPJ = oRs.Fields("TargetVolumn_AP").Value
					intBurdenPer = oRs.Fields("Burden").Value
					intContraPer = oRs.Fields("Contra").Value
					sJustificationnotes = oRs.Fields("VolumnMargin_Justification").Value
					nUsedInMOL = oRs.Fields("UsedInMOL").Value
					sManufactureCountry = oRs.Fields("ManufactureCountry").Value
					sWarrantyCode = oRs.Fields("WarrantyCode").Value
					sObsoleteDate = oRs.Fields("ObsoleteDate").Value
					sVisibility_NA = oRs.Fields("Visibility_NA").Value
					sVisibility_EM = oRs.Fields("Visibility_EM").Value
					sVisibility_AP = oRs.Fields("Visibility_AP").Value
					sVisibility_LA = oRs.Fields("Visibility_LA").Value
					strReplacementDescription = oRs.Fields("ReplacementAVDescription").Value
					strOrderInstructions = oRs.Fields("OrderInstruction").Value
					sGroupByDivision = oRs.Fields("GroupToDivisions").Value
					strRuleDescription = oRs.Fields("RuleDescription").Value
					sComparatibilitySelected = oRs.Fields("ComparatibilityDivisions").Value
					nProductLineId = oRs.Fields("ProductLineID").Value
					sProductline_Ori = ""
					if oRs.Fields("Localized").Value = 1 Or cbool(oRs.Fields("SCM_Localized").Value) = True then
						sLocalized = 1
					else
						sLocalized = 0
					end if		
							
									
					if isnumeric(intTargetNA) and isnumeric(intTargetLA) and isnumeric(intTargetEMEA) and isnumeric(intTargetAPJ) then
						intTargetAllReg	= CDbl(intTargetNA) + CDbl(intTargetLA) + CDbl(intTargetEMEA) + CDbl(intTargetAPJ)
					end if					
					
					if isnumeric(sAMOCost) and isnumeric(sAMOWWPrice) then					
						if isnumeric(intBurdenPer) and isnumeric(intContraPer) then
						  intBurden	= Round(CDbl(sAMOCost) * (intBurdenPer/100),2)				
						  intDealerMargin =  Round(CDbl(sAMOWWPrice) * 0.94,2)											 
						  intContra	= Round((intContraPer/100) * intDealerMargin,2)
						  intNetRevenue = Round(intDealerMargin - intContra,2)
						  intTargetCostBurden	= Round(intBurden + CDbl(sAMOCost),2)
						  intGrossMargin	= Round(intNetRevenue - intTargetCostBurden,2)
						  if intNetRevenue > 0 then		
							intGrossMarginPer = Round((intGrossMargin / intNetRevenue)*100,2)
							intLifetimeGrossMargin  = Round(intGrossMargin * CDbl(intTargetAllReg),2)
						  else
							intGrossMarginPer = ""
						  end if						  
						end if			    
					end if
									
					if sUserDivisions = "" Then
					  bEditPlatform = False
					else
						bEditPlatform = isIdBelong(sUserDivisions,oRs.Fields("DivisionIDs").Value,"|")
					end if				    
					
					if nMode = 3 then
						sOriginalDivisionIDs = oRs.Fields("DivisionIDs").Value
						if right(sOriginalDivisionIDs, 1) <> "," then
							sOriginalDivisionIDs = sOriginalDivisionIDs & ","
						end if
					else
						lngGroupID = oRs.Fields("GroupID").Value
						sDivisionIDs = oRs.Fields("DivisionIDs").Value
					end if
					
					strNotes = oRs.Fields("Notes").value
					lngMOLHide = oRs.Fields("MOLHide").value
					lngSCMHide = oRs.Fields("SCMHide").value
					'lngSCLHide = oRs.Fields("SCLHide").value
					bIDP = oRs.Fields("NoSCLDeploy").value
										
					set oRs = nothing	
				
					if nMode = 3 then
						if right(sDivisionIDs, 1) = "," then
							sDivisionIDs = left(sDivisionIDs, len(sDivisionIDs)-1)
						end if
						set oRsUserRights = GetAllRolesByDivision( Session("AMOUserID"), sDivisionIDs)
					else
						set oRsUserRights= GetAllRolesByDivision( Session("AMOUserID"), sGroupByDivision)
					end if
					
					if oRsUserRights is nothing then
						sErr = "Missing required parameters.  Unable to complete your request."
		                Response.Write(sErr)
		                Response.End()
					else
						GetRightsByRecordset Application("AMOList"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, oRsUserRights
						GetRightsByRecordset Application("AMOCost"), bCostCreate, bCostView, bCostUpdate, bCostDelete, oRsUserRights
						GetRightsByRecordset Application("AMOPOR"), bPORCreate, bPORView, bPORUpdate, bPORDelete, oRsUserRights	
					
						oRsUserRights.Close
						select case nMode
							case 2
								sHeader = "View/Modify After Market Option"
							case 3
								sHeader = "Clone After Market Option"
								if bAMOCreate then
									bUpdate = True
								end if
							case 4
								sHeader = "View After Market Option"
						end select
	
						'make just one variable to check throughout
						'check to make sure option is not in RAS Review, the user has update rights
						if nMode = 2 then
							if  clng(nAMOStatusID) <> clng(Application("AMO_DISABLED")) _
								and (bAMOUpdate or (bCostUpdate and not bEditPlatform))  then
								bUpdate = True
								
							end if
							
							'if user can update but we're not in edit mode, don't let update happen
							if not bEdit and bUpdate then
								bUpdate = false
							end if
		
							if clng(nAMOStatusID) = clng(Application("AMO_DISABLED")) then
								sHeader = "View After Market Option"
								bEdit = False
							end if
							
							
						end if
	
						if nMode = 3 then
							'suggest new Marketing Description if total length <= 64 characters
							sDesc = "Clone of " & sDesc
							if len(sDesc) >= 200 then
								sDesc = left(sDesc, 200)
							end if
							sAMOStatus = "New"
						end if
	
						'if user is in more than one user group that can update options,
						'give drop down to change the ownership of an option
						if bUpdate and bAMOUpdate and bEdit then
							set oRsGroups = GetGroupsForRole2(cstr(Application("AMOLIST")), false, false, true, false, true, false)
							if oRsGroups is nothing then
								sErr = "Missing required parameters.  Unable to complete your request."
		                        Response.Write(sErr)
		                        Response.End()
							else
								if (not oRsGroups is nothing) then
									if oRsGroups.RecordCount > 1 then
										strOwnerHTML = Lbx_GetHTML2("lbxGroupID", false, 1, 300, oRsGroups, "GroupName", "GroupID", clng(lngGroupID))
									end if
								end if
							end if
						end if 'if bUpdate then
					end if 'no user rights found
					set oRsUserRights = nothing
				end if	'no recordset found
			end if
			
	
			if sErr = "" then
				'get selected platforms
				set oSelectedRs = oSvr.AMOPlatforms_Search(Application("REPOSITORY"), nID & "|", 1, 0)

				if oSelectedRs is nothing then
					sErr = "Missing required parameters.  Unable to complete your request."
		            Response.Write(sErr)
		            Response.End()
				end if
				'don't show ChangeMask = 2 because those are already unchecked by user
				oSelectedRs.Filter = "ChangeMask <> 2"
				if nMode = 3 and sOriginalDivisionIDs <> sDivisionIDs then
					'make no selected platforms
					set oSelectedRs = CopyRS(oSelectedRs, True)
				end if
			end if
	end select
end if

if sErr = "" then

	sCtrlStyle = " readonly " & Application("DISABLED_CTRL_STYLE")
	sCostCtrlStyle = " disabled " & Application("DISABLED_CTRL_STYLE")
	sRasCtrlStyle = " readonly " & Application("DISABLED_CTRL_STYLE")
	sRasDisCtrlStyle = " disabled " & Application("DISABLED_CTRL_STYLE")
	sDisabledCtrlStyle = " disabled " & Application("DISABLED_CTRL_STYLE")
	sDisabledIDP = sDisabledCtrlStyle
	
	if nMode <> 4 then
		if bUpdate then
			if bAMOUpdate or bAMOCreate then
				sCtrlStyle = ""
				sDisabledCtrlStyle = ""
				sDisabledIDP = ""
			end if
			if (bAMOUpdate or bCostUpdate or bAMOCreate) and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then
				sCostCtrlStyle = ""
			end if
			
			if (bAMOUpdate or bAMOCreate) and (clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW"))) and (clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")))then
				sRasCtrlStyle = ""
				sRasDisCtrlStyle = ""
			end if
			
		end if
	end if
	
	if clng(nAMOStatusID) <> clng(Application("AMO_NEW")) then
		if clng(nAMOStatusID) = 0 then
			sDisabledIDP = ""
		else
			sDisabledIDP = " disabled " & Application("DISABLED_CTRL_STYLE")
		end if
	end if
	
	
	'make mktg desc. readonly if module is used in a POR'd MOL - quoc wants to remove this lock but thinks might come back
'	if nUsedInMOL = 2 and bUpdate then
'		sNameCtrlStyle = " disabled " & Application("DISABLED_CTRL_STYLE")
'		sNameComment = " (Cannot change this field as the module is used in PORed MOL. Please contact IRS admin to change)"
'	else
		sNameCtrlStyle = sCtrlStyle
		sNameComment = ""
'	end if 

	'Check if the user is an PAL admin (or has KE rights)
	GetRights2 Application("AVLKE"), bAVLKECreate, bAVLKEView, bAVLKEUpdate, bAVLKEDelete
	'Only PAL admins can edit the RuleID filed, disable when the module is being created, unless the owner has PAL admin access
	if bAVLKECreate and bEdit then 
			sRuleIDCtrlStyle = ""
	else
		sRuleIDCtrlStyle = " disabled " & Application("DISABLED_CTRL_STYLE")
	end if 

	'get available Platforms
	
	'set oErr = GetCategory(oRsPlatforms, 61, 0)

	'only show platforms that are in the target business segments
	'strip trailing comma
	if right(sDivisionIDs, 1) = "," then
		sDivisionIDs = left(sDivisionIDs, len(sDivisionIDs)-1)
	end if
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	if not bAMOCreate and bEditPlatform then	
		sPlUserDivisions = sUserDivisions
		sPlUserDivisions = replace(sPlUserDivisions, "|", ",")
		if right(sPlUserDivisions, 1) = "," then
			sPlUserDivisions = left(sPlUserDivisions, len(sPlUserDivisions)-1)
		end if	
		set oRsPlatforms = oSvr.AMO_AllPlatforms_Search(Application("REPOSITORY"), sPlUserDivisions)
	else
		set oRsPlatforms = oSvr.AMO_AllPlatforms_Search(Application("REPOSITORY"), sDivisionIDs)
	end if

		
	if oRsPlatforms is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
	end if
'	arrDivisions = split(sDivisionIDs, ",")
'	strFilter = ""
'	for each strDiv in arrDivisions
'		if strFilter <> "" then
'			strFilter = strFilter & " or "
'		end if
'		strFilter = strFilter & "DivisionID = " & strDiv
'	next
'	oRsPlatforms.Sort = "Description DESC"
'	oRsPlatforms.Filter = strFilter
	' need to create the recordset again because the duallistbox function will change the filter
	set oRsPlatforms = CopyRS(oRsPlatforms, True)
end if


if sErr = "" then
	'get module HW Categories
	'set oErr = GetMOLCategory(oRsHWCategory, 1)
	set oRsHWCategory = GetMOLCategory(1)	
	if oRsHWCategory is Nothing then
		Response.Write("Recordset error: oRsHWCategory")
		Response.End()
	else
		select case nMode
			case 1
				'Show only active categories
				if not oRsHWCategory.EOF and not oRsHWCategory.BOF then
					oRsHWCategory.Filter = "State = 1"
				end if
			case 2, 3, 4
				if not oRsHWCategory.EOF and not oRsHWCategory.BOF then
					oRsHWCategory.Filter = "State = 1 OR CategoryID = " & nTypeID
				end if
		end select

		oRsHWCategory.Sort = "FullCatDescription"
		
		if bUpdate and nMode <> 4 then
			sHWModuleCategoryHTML = Lbx_GetHTML5("lbxTempCategory", false, 1, 0, _
					oRsHWCategory, "FullCatDescription", "CategoryID", nTypeID, true, "", not ((bAMOUpdate or bAMOCreate) and bEdit))
		else
			sHWModuleCategoryHTML = Lbx_GetHTML5("lbxTempCategory", false, 1, 0, _
					oRsHWCategory, "FullCatDescription", "CategoryID", nTypeID, true, "", true)
		end if
	end if
end if

if sErr = "" then
	'get module SW Categories
	'set oErr = GetMOLCategory(oRsSWCategory, 12)
	set oRsSWCategory = GetMOLCategory(12)	
	if oRsSWCategory is Nothing then
		Response.Write("Recordset error: oRsSWCategory")
		Response.End()	 
	else
		select case nMode
			case 1
				'Show only active categories
				if not oRsSWCategory.EOF and not oRsSWCategory.BOF then
					oRsSWCategory.Filter = "State = 1"
				end if
			case 2, 3, 4
				if not oRsSWCategory.EOF and not oRsSWCategory.BOF then
					oRsSWCategory.Filter = "State = 1 OR CategoryID = " & nTypeID
				end if
		end select

		oRsSWCategory.Sort = "FullCatDescription"

		if bUpdate and nMode <> 4 then
			sSWModuleCategoryHTML = Lbx_GetHTML5("lbxTempCategory", false, 1, 0, _
					oRsSWCategory, "FullCatDescription", "CategoryID", nTypeID, true, "", not ((bAMOUpdate or bAMOCreate) and bEdit))
		else
			sSWModuleCategoryHTML = Lbx_GetHTML5("lbxTempCategory", false, 1, 0, _
					oRsSWCategory, "FullCatDescription", "CategoryID", nTypeID, true, "", true)
		end if
	end if
end if

if sErr = "" then
	'Get Module types
	'set oErr = GetMOLCategory(oRs, 10)
	set oRs = GetMOLCategory(10)	
	if oRs is Nothing then
		Response.Write("Recordset error: oRs")
		Response.End() 
	else
		select case nMode
			case 1
				'Show only active Types
				if not oRs.EOF and not oRs.BOF then
					oRs.Filter = "State = 1"
				end if
			case 2, 3, 4
				if not oRs.EOF and not oRs.BOF then
					oRs.Filter = "State = 1 OR CategoryTypeID = " & nModuleTypeID
				end if
		end select

		sModuleTypeHTML = ""
		Do until oRs.EOF
		
			if sModuleTypeHTML <> "" then
				sModuleTypeHTML = sModuleTypeHTML & "&nbsp;&nbsp;"
			end if
			if (oRs("CategoryTypeID").Value = nModuleTypeID) or (nModuleTypeID=0 and oRs("CategoryTypeID").Value = Application("MD_TYPE_HW")) then
				sModuleTypeHTML = sModuleTypeHTML & "<INPUT type='radio' id=rdType name=rdType onclick='javascript:rdType_onclick()' value=" & oRs("CategoryTypeID").Value & " checked" & sDisabledCtrlStyle & ">" & oRs("Description").Value
			else
				sModuleTypeHTML = sModuleTypeHTML & "<INPUT type='radio' id=rdType name=rdType onclick='javascript:rdType_onclick()' value=" & oRs("CategoryTypeID").Value & " " & sDisabledCtrlStyle & ">" & oRs("Description").Value
			end if
			oRS.MoveNext
		Loop
		oRs.Close
	end if
end if

if sErr = "" then
	'Get Module Divisions
	'set oErr = GetMOLCategory(oRs, 9)
	set oRs = GetMOLCategory(9)	
	if oRs is Nothing then
		Response.Write("Recordset error: oRs")
		Response.End() 
	else
		select case nMode
			case 1
				'Show only active Divisions
				if not oRs.EOF and not oRs.BOF then
					oRs.Filter = "State = 1"
				end if
			case 2, 3, 4
				if not oRs.EOF and not oRs.BOF then
					strFilter = "State = 1"
					'strFilter = ""
					if sDivisionIDs <> "" Then
						sTmp = sDivisionIDs & ","
						While Trim(sTmp) <> ""
							sDivisionID = Left(sTmp, InStr(1, sTmp, ",") - 1)
							if strFilter = "" then
								strFilter = strFilter & "CategoryID = " & sDivisionID
							else
								strFilter = strFilter & " OR CategoryID = " & sDivisionID
							end if
							sTmp = Right(sTmp, Len(sTmp) - InStr(1, sTmp, ","))
						Wend
					End If
					oRs.Filter = strFilter
				end if
		end select

		sModuleDivisionHTML = ""
		sDivisionTarget = ""
	
		Do until oRs.EOF
			if sModuleDivisionHTML <> "" then
				sModuleDivisionHTML = sModuleDivisionHTML & "&nbsp;&nbsp;"
			end if
			
			if instr("," & sDivisionIDs & ",", "," & cstr(oRs("CategoryID").Value) & ",") > 0 then
			
			  if (instr("|" & sUserDivisions & "|", "|" & cstr(oRs("CategoryID").Value) & "|") = 0) and not bAMOCreate and ( (nAMOStatusID <> Application("AMO_DISABLED") )  and ( clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) ) and ( clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) ))  then
				'steven
				'Response.Write "if status is not DISABLE, RAS-review, RAS-update"
				sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " checked disabled>" & oRs("Description").Value
				sDivisionTarget = sDivisionTarget & oRs("Description").Value & " ; "
				
			  elseif isIdBelong(sUserDivisions, cstr(Trim(oRs("CategoryID").Value)),"|") and nAMOStatusID <> Application("AMO_DISABLED")  and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE"))  then
				if bEdit then
					sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " checked>" & oRs("Description").Value
				else
					'steven
					'Response.Write "elseif bEdit =false, status is not DISABLE, RAS-review, RAS-update, disable chkbox"
					'OLD:disabled chkbox when status = complete
					'sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " checked disabled>" & oRs("Description").Value
	
					sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " checked >" & oRs("Description").Value
				end if
				sDivisionTarget = sDivisionTarget & oRs("Description").Value & " ; "
				bBusEnabled = True

			  else
				'steven
				'Response.Write "else other status...leave chkbox" & sDisabledCtrlStyle
			  
				sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " checked " & sDisabledCtrlStyle & ">" & oRs("Description").Value
				sDivisionTarget = sDivisionTarget & oRs("Description").Value & " ; "
				
			  end if
			else
				if isIdBelong(sUserDivisions,cstr(Trim(oRs("CategoryID").Value)),"|") and nAMOStatusID <> Application("AMO_DISABLED")  and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE"))  then
					sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & ">" & oRs("Description").Value
					bBusEnabled = True
				else
					sModuleDivisionHTML = sModuleDivisionHTML & "<INPUT type='checkbox' id=chkDivision name=chkDivision value=" & oRs("CategoryID").Value & " " & sDisabledCtrlStyle & ">" & oRs("Description").Value
				end if
			end if
			oRS.MoveNext
			
		Loop
		oRs.Close
	end if
	
	'set oErr = GetMOLCategory(oRsProductLine, 29)
	set oRsProductLine = GetMOLCategory(29)	
	if oRsProductLine is Nothing then
		Response.Write("Recordset error: oRsProductLine")
		Response.End()
    end if
	
end if 

function isIdBelong(byval dIds, byval sIds, byVal charSplit)
	dim i, arrIds, bFlag
	
	if Trim(dIds) <> "" then
		arrIds = Split(Trim(dIds), charSplit)
		bFlag = False
		for i = 0 to UBound(arrIds)
			if Trim(arrIds(i)) <> "" and instr(Trim(cstr(sIds)), Trim(cstr(arrIds(i)))) > 0 then
				bFlag = True
				Exit For
			end if
		Next
	end if
	
	isIdBelong = bFlag	 
end function

function GetPlatformByDivision(byval oRs, byVal sModuleDivIds)
	dim oDupRs, oFld, i, arrModuleDivId, sDivId
	set oDupRs = Server.CreateObject ("ADODB.Recordset")

	for i = 0 to oRs.Fields.Count - 1
		set oFld =  oRs.Fields(i)
		oDupRs.Fields.Append oFld.Name, oFld.Type, oFld.DefinedSize, oFld.Attributes
	next
	oDupRs.CursorLocation = 3	'Use client-side cursors
	oDupRs.Open

	If oRs.RecordCount > 0 Then
		oRs.MoveFirst
	End If
	
	sOtherSelectedAliasIDs = ""
	
	While Not oRs.EOF
		 
		if isIdBelong(oRs.Fields("DivisionIDs").Value, sModuleDivIds, ",") then
			oDupRs.AddNew
			For i = 0 To oRs.Fields.Count - 1
				oDupRs.Fields(oDupRs.Fields(i).Name).Value = oRs.Fields(i).Value
			Next
		else
		   if sOtherSelectedAliasIDs = "" then
			sOtherSelectedAliasIDs = oRs.Fields("AliasID").Value
		   else
			 sOtherSelectedAliasIDs = sOtherSelectedAliasIDs & ", " & oRs.Fields("AliasID").Value 
		   end if
		end if
	
		oRs.MoveNext
	Wend
	
	If oRs.RecordCount > 0 Then
		oRs.MoveFirst
	End If
	
	set GetPlatformByDivision = oDupRs

end function

set oRs = nothing
%>
<!DOCTYPE html>
<HTML>
<head>
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sHeader%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sHeader%> - AMO Properties</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT type="text/javascript">
<!-- 
//var oPopup = window.createPopup();
var textdecoration = '';

var ChangeMask;

function ClickEvent(evt) {
	var id;
	var ModuleID, CategoryID, PlatformID, RegionID, GEOID;
	if (!evt) evt = window.event;
	var objUnknown = evt.srcElement? evt.srcElement : evt.target;	
	id = objUnknown.id;
			
	ModuleID = <%=nID%>;  
	
	if ((objUnknown.tagName.toUpperCase() != "TD") && (objUnknown.tagName.toUpperCase() != "IMG"))
		return;

	if (objUnknown.tagName.toUpperCase() == "IMG") {
		// make object the parent TD
		objUnknown = objUnknown.parentNode; //parentElement
	}

	//Extract ModuleID and CategoryID from 'm123c123'
	
	CategoryID = objUnknown.parentNode.getAttribute("cid");

	switch (id){
		case 'prop': //Properties column
			ShowModuleProperties(ModuleID)
			break;
		case 'o': //Short Description column
			break;
		case 'rc': //Region Comment column
			enterComment(ModuleID, 'rc')
			break;
		case 'gc': // GEO column
			GEOID = objUnknown.gid
			editGEODate(objUnknown, ModuleID, GEOID);
			break;
		case 'reg': // Region column
			//Extract RegionID
			RegionID = objUnknown.rid;
			
			if (objUnknown.innerHTML == "&nbsp;") {
				// add checkmark				
				cM_ChangeRegion(ModuleID, RegionID, 1, objUnknown);
			} else {
				// remove checkmark								
				cM_ChangeRegion(ModuleID, RegionID, 0, objUnknown);
			}
			break;
	}
}

function checkEnter(theitem, e) {
	if (!e) e = window.event;
	var charCode = e.keyCode ? e.keyCode : e.which;
	if (charCode == 13) {
		theitem.blur()
		return false;
	}
}

function editGEODate(evtobj, ModuleID, GEOID) {
	var sHTML
	var objUnknown = evtobj;
	if (objUnknown.innerHTML.indexOf("editcell" + ModuleID + GEOID)<0) {
		sHTML = "<input onKeyPress='return checkEnter(this, event)' "
		sHTML += "onBlur='javascript:getGEODate(event," + ModuleID + "," + GEOID + ",\"" + objUnknown.innerHTML.replace(/&nbsp;/g, "") + "\")' "
		sHTML += "type=text maxlength=10 size=10 value=\"" + objUnknown.innerHTML.replace(/&nbsp;/g, "") + "\"' "
		sHTML += "id='editcell" + ModuleID + GEOID + "' NAME='editcell" + ModuleID + GEOID + "'>";
		objUnknown.innerHTML = sHTML
		document.getElementById("editcell" + ModuleID + GEOID).focus()
	}
}

//*****************************************************************
//Function:      getGEODate();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function getGEODate(evt, ModuleID, GEOID, OldValue) {
	if (!evt) evt = window.event;
	var objUnknown = evt.srcElement? evt.srcElement : evt.target;	
	var Parent = objUnknown.parentNode; //parentElement
	var ajaxurl = "";
	var strValue = "";

	if (objUnknown.id == "editcell" + ModuleID + GEOID) {
	    if (!checkDate (objUnknown, "GEO Date", true)){
	        return false;
	    }

	    var NewValue = objUnknown.value;
		if (OldValue == NewValue) { // nothing changed
				if (NewValue == '') {
				    Parent.innerHTML = '&nbsp;';
				} else {
					Parent.innerHTML = objUnknown.value;
				}
		} else { // something changed, save the data
		    //var objRS = RSGetASPObject("AMO_RS.asp");
		    strValue = NewValue;
		    
            var fullname = "<%= session("FullName") %>";
            ajaxurl = "AMO_SetGEODate.asp?RGS=1&ModuleID=" + ModuleID + "&GEOID=" + GEOID + "&Value=" + strValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
		    //var objResult = objRS.setGEODate("<%=Application("REPOSITORY")%>", ModuleID, GEOID, NewValue, "<%= session("FullName") %>");
	
			$.ajax({
			    url: ajaxurl,
			    type: "GET",
			    async: false,
			    success: function (data) {
			        errormsg = data;                     
			    },
			    error: function (xhr, status, error) {
			        errormsg = error; 
			        erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";                  
			    }
			})

			if (errormsg == "success") {
				// highlight the field
				Parent.className = "clsAMO_ChangedCell";
				if (NewValue == '') {
					Parent.innerHTML = '&nbsp;'
				} else {
					Parent.innerHTML = objUnknown.value;
				}
			}
		}
	}
}

function enterComment(ModuleID, Field) {
	thisform.ID.value = ModuleID
	thisform.Field.value = Field
	thisform.action = "AMO_AddComment.asp";
	thisform.submit ();
}

function cM_ChangeRegion(ModuleID, RegionID, SetStatus, objUnknown) {
	// The variables "lefter" and "topper" store the X and Y coordinates
	// to use as parameter values for the following show method. In this
	// way, the popup displays near the location the user clicks. 
	var lefter = event.clientX;
	var topper = event.clientY;
	var popupBody;
	
	popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative;TOP:0px\">"; 

	popupBody= popupBody+"<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='pointer';this.style.color='white'\"onmouseout=\"this.style.background='white';this.style.color='black'\">";
	popupBody = popupBody + "<FONT face=Arial size=2>";
	
	if (SetStatus == 1) {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeRegionStatus(" + ModuleID + "," + RegionID + "," + SetStatus + "," + objUnknown.sourceIndex + ")'\" >&nbsp;&nbsp;Add to Region <img src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0><\/SPAN><\/FONT><\/DIV>";
	} else {
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:ChangeRegionStatus(" + ModuleID + "," + RegionID + "," + SetStatus + "," + objUnknown.sourceIndex + ")'\" >&nbsp;&nbsp;Remove from Region<\/SPAN><\/FONT><\/DIV>";
	}

	popupBody = popupBody + "</DIV>";
	
	oPopup.document.body.innerHTML = popupBody; 

	if (SetStatus == 1) {
		oPopup.show(lefter, topper, 130, 22, document.body);
	} else {
		oPopup.show(lefter, topper, 150, 18, document.body);
	}
}

function ShowModuleProperties(ModuleID) {
    window.open("/IPulsar/Features/AMOFeatureProperties.aspx?FromModule=1&FeatureID="+ModuleID, "_blank", "resizable=yes,menubar=yes,scrollbars=yes,toolbar=yes");
}

function ChangeRegionStatus(ModuleID, RegionID, SetStatus, srcIndex) {
	var strCheckmark, err;

	if (SetStatus == 1) {
		// add the checkmark
		strCheckmark = "<img onclick='javascript:ClickEvent();return true;' "
		strCheckmark += "id='" + document.all(srcIndex).id + "' title='" + document.all(srcIndex).title + "' "
		strCheckmark += "rid='" + document.all(srcIndex).rid + "' "
		strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>"
		document.all(srcIndex).innerHTML = strCheckmark
		// highlight the field
		document.all(srcIndex).className = "clsAMO_ChangedCell";
		document.all("strRegionIdMaskIds").value = document.all("strRegionIdMaskIds").value + RegionID + "," + SetStatus + ";";
	} else {
		// clear the checkmark
		document.all(srcIndex).innerHTML = "&nbsp;"
		// highlight the field
		document.all(srcIndex).className = "clsAMO_ChangedBlankCell";
		document.all("strRegionIdMaskIds").value = document.all("strRegionIdMaskIds").value + RegionID + "," + SetStatus + ";";
	}
	
	oPopup.hide();
}

function lbxGo_onchange() {
	// if no status check boxes checked, then check all
	var bNoneChecked = 0;
	var oObject = document.getElementsByName("chkStatus")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoneChecked = 1;
				break;
			}
		}
		if (bNoneChecked == 0) {
			// none were checked, go through and check them all
			for (i = 0; i < oObject.length; i++) {
				oObject[i].checked = true;
			}
		}
	}
	// if no business segment check boxes checked, then check all
	bNoneChecked = 0;
	oObject = document.getElementsByName("chkBusSeg")
	if (oObject) {
		for (i = 0; i < oObject.length; i++) {
			if (oObject[i].checked) {
				bNoneChecked = 1;
				break;
			}
		}
		if (bNoneChecked == 0) {
			// none were checked, go through and check them all
			for (i = 0; i < oObject.length; i++) {
				oObject[i].checked = true;
			}
		}
	}
	if (!checkDate (thisform.txtEOLDate, "Show options with End of Manufacturing (EM)", true))
		return false;

	thisform.action = "AMO_Localization.asp";	
	thisform.submit();
}

<%
if nMode <> 4 then

'Part of the ChangeMask table:
'POS MASK
'0	Blue Part Number										BluePartNo		1
'1	Red Part Number											RedPartNo		2
'2	PHweb (General) Availability (GA) (RAS Availability Date)		BOMRevADate		4
'Pass in the MASK to the ThisChanged function.
%>
function ThisChanged(intAdd) {
	// OR the mask to the existing ChangeMask
	ChangeMask = ChangeMask | intAdd
}

function ValidateInputPOR() {
	<% if IsODM = 0 then %>
	if (!thisform.txtAMOCost.disabled) {
		if (isWhitespace(thisform.txtAMOCost.value) ) {
			warnInvalid(thisform.txtAMOCost, "Please enter AMO Cost.");
			return false;
		}
	}
	
	if (!thisform.txtAMOWWPrice.disabled) {
		if (isWhitespace(thisform.txtAMOWWPrice.value) ) {
			warnInvalid(thisform.txtAMOWWPrice, "Please enter AMO target price.");
			return false;
		}	
	}

	if (isWhitespace(thisform.txtTargetNA.value) ) {
		warnInvalid(thisform.txtTargetNA, "Please enter Target Lifetime Volume - North America.");
		return false;
	}
	
	if (isWhitespace(thisform.txtTargetLA.value) ) {
		warnInvalid(thisform.txtTargetLA, "Please enter Target Lifetime Volume - Latin America.");
		return false;
	}
	
	if (isWhitespace(thisform.txtTargetEMEA.value) ) {
		warnInvalid(thisform.txtTargetEMEA, "Please enter Target Lifetime Volume - EMEA .");
		return false;
	}
	
	if (isWhitespace(thisform.txtTargetAPJ.value) ) {
		warnInvalid(thisform.txtTargetAPJ, "Please enter Target Lifetime Volume -APJ .");
		return false;
	}
	
	if (isWhitespace(thisform.txtBurdenPer.value) ) {
		warnInvalid(thisform.txtBurdenPer, "Please enter burden percentage.");
		return false;
	}	
	
	//make sure there is only one decimal place
	var floatBurden = parseFloat(thisform.txtBurdenPer.value);
	var intBurden = parseInt(floatBurden * 10);
	if (intBurden != (floatBurden * 10)){
		alert("Burden can have only one decimal place. Ex: 12.3");
		return false;
	}

	if (isWhitespace(thisform.txtContraPer.value) ) {
		warnInvalid(thisform.txtContraPer, "Please enter contra percentage .");
		return false;
	}
	
	//make sure there is only one decimal place
	var floatContra = parseFloat(thisform.txtContraPer.value);
	var intContra = parseInt(floatContra * 10);
	if (intContra != (floatContra * 10)){
		alert("Contra can have only one decimal place. Ex: 12.3");
		return false;
	}
	<% end if %>

	return true;
}

<% if IsODM = 0 then %>
function btnCalculate_onclick() {
	var intTargetNA, intTargetLA, intTargetEMEA, intTargetAPJ, 
	intTargetAllReg, intBurdenPer, intContraPer, intBurden, 
	intContra, intNetRevenue, intDealerMargin, intTargetCostBurden, 
	intGrossMarginPer, intGrossMargin, intLifetimeGrossMargin, 
	intAMOWWPrice, intAMOCost, i;
	
	if (ValidateInputPOR()) {
		intAMOCost		= thisform.txtAMOCost.value;
		
		if(intAMOCost == "") {				
			alert("Cannot calculate POR Detail invalid AMO Cost");
			return false;
		}
		
		intAMOWWPrice	= thisform.txtAMOWWPrice.value;
		
		if(intAMOWWPrice == "") {				
			alert("Cannot calculate POR Detail invalid AMO Price");
			return false;
		}
		
		intTargetNA		= parseFloat(thisform.txtTargetNA.value);
		intTargetLA		= parseFloat(thisform.txtTargetLA.value);
		intTargetEMEA	= parseFloat(thisform.txtTargetEMEA.value);
		intTargetAPJ	= parseFloat(thisform.txtTargetAPJ.value);
		intBurdenPer	= parseFloat(thisform.txtBurdenPer.value);
		intContraPer	= parseFloat(thisform.txtContraPer.value);
	
		intAMOWWPrice = intAMOWWPrice.replace(",","");
		intAMOWWPrice = intAMOWWPrice.replace(",","");
		intAMOWWPrice = intAMOWWPrice.replace(",","");
		intAMOWWPrice = intAMOWWPrice.replace(",","");
		intAMOCost = intAMOCost.replace(",","");
		intAMOCost = intAMOCost.replace(",","");
		intAMOCost = intAMOCost.replace(",","");
		intAMOCost = intAMOCost.replace(",","");
				
		intTargetAllReg	= intTargetNA + intTargetLA + intTargetEMEA + intTargetAPJ;		
		intBurden	= round(intAMOCost * (intBurdenPer/100),2);			
		intDealerMargin = round((intAMOWWPrice * 0.94),2);							 
		intContra	= round((intContraPer/100) * intDealerMargin,2);		
		intNetRevenue = round(intDealerMargin - intContra,2);
	
		intTargetCostBurden	= round((parseFloat(intBurden) + parseFloat(intAMOCost)),2);
		
		intGrossMargin	= round(intNetRevenue - intTargetCostBurden,2);
	
		if (intNetRevenue > 0) { 	
			intGrossMarginPer = round((intGrossMargin / intNetRevenue) * 100,2);
			intLifetimeGrossMargin  = round(intGrossMargin * intTargetAllReg,2);
		}
		else
			intGrossMarginPer = "";
	
		document.getElementById("txtTargetAllReg").value = intTargetAllReg; 
		document.getElementById("txtBurden").value = intBurden;
		document.getElementById("txtContra").value = intContra;
		document.getElementById("txtNetRevenue").value = intNetRevenue;
		document.getElementById("txtDealerMargin").value = intDealerMargin;
		document.getElementById("txtTargetCost").value = intTargetCostBurden;
		document.getElementById("txtGrossMarginPer").value = intGrossMarginPer;
		document.getElementById("txtGrossMargin").value = intGrossMargin;
		document.getElementById("txtLifetimeGrossMargin").value = intLifetimeGrossMargin;
	}
}

function round(number,X) {
	// rounds number to X decimal places, defaults to 2
	X = (!X ? 2 : X);
	return Math.round(number*Math.pow(10,X))/Math.pow(10,X);
}

function btnClear_onclick() {
	document.getElementById("txtTargetNA").value = "";
	document.getElementById("txtTargetLA").value = "";
	document.getElementById("txtTargetEMEA").value = "";
	document.getElementById("txtTargetAPJ").value = "";
	document.getElementById("txtBurdenPer").value = "";
	document.getElementById("txtContraPer").value = "";
	document.getElementById("txtTargetAllReg").value = "";
	document.getElementById("txtBurden").value = "";
	document.getElementById("txtContra").value = "";
	document.getElementById("txtNetRevenue").value = "";
	document.getElementById("txtDealerMargin").value = "";
	document.getElementById("txtTargetCost").value = "";
	document.getElementById("txtGrossMarginPer").value = "";
	document.getElementById("txtGrossMargin").value = "";
	document.getElementById("txtLifetimeGrossMargin").value = "";
	document.getElementById("txtLifetimeGrossMargin").value = "";
}
<% end if %>

function checkNumeric(e) {
	// Get ASCII value of key that user pressed
	if (!e) e = window.event;
	var key = e.keyCode ? e.keyCode : e.which;

	// Was key that was pressed a numeric character (0-9) or backspace or decimal point?
	if (( key > 47 && key < 58 ) || key == 8 || key == 46)
		return; // if so, do nothing
	else // otherwise, discard character 
		if (window.event)
			e.returnValue = null; // IE
		else
			e.preventDefault(); // Firefox
}

function checkSymbol(object, name)
{
	var varValue = object.value;
	var aPosition = varValue.indexOf("\"");
	if (aPosition > -1) {
		 alert( "Please remove symbol in " + name + "!" ); 
		 object.focus();    
		 return false; }
	else {
		return true;
		}
}

function ValidateInput() {
	// make sure all required fields are there
	var tmp;
	var chekedTypeID;

	if (isWhitespace(thisform.txtDesc.value) ) {
		warnInvalid(thisform.txtDesc, "Please enter a Marketing Description.");
		return false;
	} else
	{
		if (!checkSymbol(thisform.txtDesc, "Marketing Description")) 
			return false;	
	}

	if (isWhitespace(thisform.txtShortDesc.value) ) {
		warnInvalid(thisform.txtShortDesc, "Please enter a Short Description.");
		return false;
	}else
	{
		if (!checkSymbol(thisform.txtShortDesc, "Short Description")) 
			return false;	
	}
	
//	if (thisform.txtShortDesc.value.length > 30) {
//		warnInvalid(thisform.txtShortDesc, "Maximum length for Short Description is 30 characters");
//		return false;
//	}
	
	if (!checkSymbol(thisform.txtLongDescription, "Long Description")) {
		return false;
	}
	
	if (isWhitespace(thisform.txtBluePN.value) ) {
		warnInvalid(thisform.txtBluePN, "Please enter a HP PartNo.");
		return false;
	}

	var bChecked = false;
	for (var i=0; i<thisform.rdType.length; i++) {
		if (thisform.rdType[i].checked) {
			bChecked = true;
			chekedTypeID = thisform.rdType[i].value
		}
	}
	if (!bChecked) {
		alert("Please select a Option Type.");
		return false;}
	else
		{
			if (chekedTypeID == <%=Application("MD_TYPE_SW")%>)
			{
				if (!thisform.chkMOLHide.checked)
				{	alert("Please check 'Hide from Module and Option List' for AMO SW module");
					return false;
				}
			}
		}	

	if (document.getElementById("lbxCategory").value == 0) {
		alert("Please select an Option Category.");
		return false;
	}

	bChecked = false;
	for (i=0; i<thisform.chkDivision.length; i++) {
		if (thisform.chkDivision[i].checked) {
			bChecked = true;
		}
	}
	if (!bChecked) {
		alert("Please select at least one Business Segment.");
		return false;
	}
	<% if IsODM = 0 then %>
	if (thisform.txtAMOCost.value.length > 20) {
		warnInvalid(thisform.txtAMOCost, "Maximum length for the AMO Cost field is 20 characters");
		return false;
	}

	if (thisform.txtJustification.value.length > 300) {
		warnInvalid(thisform.txtJustification, "Maximum length for the Margin Justification notes field is 300 characters");
		return false;
	}

	tmp = stripCharsInBag(thisform.txtAMOCost.value, ',')
	if (!isFloat(tmp, true)) {
		warnInvalid(thisform.txtAMOCost, "Please enter only numbers in the AMO Cost field");
		return false;
	}

	if (thisform.txtAMOWWPrice.value.length > 20) {
		warnInvalid(thisform.txtAMOWWPrice, "Maximum length for the AMO Price field is 20 characters");
		return false;
	}
	
	tmp = stripCharsInBag(thisform.txtAMOWWPrice.value, ',')
	if (!isFloat(tmp, true)) {
		warnInvalid(thisform.txtAMOWWPrice, "Please enter only numbers in the AMO Price field");
		return false;
	}

	if (thisform.txtActualCost.value.length > 20) {
		warnInvalid(thisform.txtActualCost, "Maximum length for the Actual Cost field is 20 characters");
		return false;
	}
	
	tmp = stripCharsInBag(thisform.txtActualCost.value, ',')
	if (!isFloat(tmp, true)) {
		warnInvalid(thisform.txtActualCost, "Please enter only numbers in the Actual Cost field");
		return false;
	}
	<% end if %>
	if (!isInteger(thisform.txtNetWeight.value, true)) {
		warnInvalid(thisform.txtNetWeight, "Please enter only whole numbers in the Net Weight field");
		return false;
	}

	if (!isInteger(thisform.txtExportWeight.value, true)) {
		warnInvalid(thisform.txtExportWeight, "Please enter only whole numbers in the Export Weight field");
		return false;
	}

	if (!isInteger(thisform.txtAirPackedWeight.value, true)) {
		warnInvalid(thisform.txtAirPackedWeight, "Please enter only whole numbers in the Air Packed Weight field");
		return false;
	}

	if (!isInteger(thisform.txtAirPackedCubic.value, true)) {
		warnInvalid(thisform.txtAirPackedCubic, "Please enter only whole numbers in the Air Packed Cubic field");
		return false;
	}

	if (!isInteger(thisform.txtExportCubic.value, true)) {
		warnInvalid(thisform.txtExportCubic, "Please enter only whole numbers in the Export Cubic field");
		return false;
	}

	if (thisform.txtNotes.value.length > 1024) {
		warnInvalid(thisform.txtNotes, "Maximum length for the Notes field is 1024 characters");
		return false;
	}
	
	if (thisform.txtLongDescription.value.length > 160) {
		warnInvalid(thisform.txtLongDescription, "Maximum length for the long description field is 160 characters");
		return false;
	}
	
	if (thisform.txtReplacementDescription.value.length > 80) {
		warnInvalid(thisform.txtReplacementDescription, "Maximum length for the Replacement Description field is 80 characters");
		return false;
	}
	
	if (thisform.txtRuleDescription.value.length > 1024) {
		warnInvalid(thisform.txtRuleDescription, "Maximum length for the Rule Description field is 1024 characters");
		return false;
	}
	
	if (thisform.txtOrderInstructions.value.length > 600) {
		warnInvalid(thisform.txtOrderInstructions, "Maximum length for the Order Instructions field is 600 characters");
		return false;
	}
	<% if IsODM = 0 then %>
	//if Burden and contra are entered make sure there is only one decimal place
	if (thisform.txtBurdenPer.value != "") {
		var floatBurden = parseFloat(thisform.txtBurdenPer.value);
		var intBurden = parseInt(floatBurden * 10);
		if (intBurden != (floatBurden * 10)){
			alert("Burden (in POR Details section) can have only one decimal place. Ex: 12.3");
			return false;
		}
	}

	if (thisform.txtContraPer.value != "") {
		var floatContra = parseFloat(thisform.txtContraPer.value);
		var intContra = parseInt(floatContra * 10);
		if (intContra != (floatContra * 10)){
			alert("Contra (in POR Details section) can have only one decimal place. Ex: 12.3");
			return false;
		}
	}
	<% end if %>
	return true;
}

function currencyFormat(fld, milSep, decSep, e, maxlen) {
	var sep = 0;
	var key = '';
	var i = j = 0;
	var len = len2 = 0;
	var strCheck = '0123456789';
	var aux = aux2 = '';
	if (!e) e = window.event;
	var whichCode = e.keyCode ? e.keyCode : e.which;
	if (whichCode == 13) { // Enter
		fld.blur()
		return false;
	}
	key = String.fromCharCode(whichCode);  // Get key value from key code
	if (strCheck.indexOf(key) == -1) return false;  // Not a valid key
	len = fld.value.length;
	if (len >= maxlen)
		return false
	for(i = 0; i < len; i++)
		if ((fld.value.charAt(i) != '0') && (fld.value.charAt(i) != decSep)) break;
	aux = '';
	for(; i < len; i++)
		if (strCheck.indexOf(fld.value.charAt(i))!=-1) aux += fld.value.charAt(i);
	aux += key;
	len = aux.length;
	if (len == 0) fld.value = '';
	if (len == 1) fld.value = '0'+ decSep + '0' + aux;
	if (len == 2) fld.value = '0'+ decSep + aux;
	if (len > 2) {
		aux2 = '';
		for (j = 0, i = len - 3; i >= 0; i--) {
			if (j == 3) {
				aux2 += milSep;
				j = 0;
			}
			aux2 += aux.charAt(i);
			j++;
		}
		fld.value = '';
		len2 = aux2.length;
		for (i = len2 - 1; i >= 0; i--)
			fld.value += aux2.charAt(i);
		fld.value += decSep + aux.substr(len - 2, len);
	}
	return false;
}

<% if bFromAMO then %>
function btnCancel_onclick() {
	var ret;
	<% if bUpdate then %>
	ret = window.confirm("Are you sure you want to cancel all the changes on this page?\nClick OK to continue or Cancel to stay on this page.");
	<% else %>
	ret = window.confirm("Are you sure?");
	<% end if %>
	if (ret)
		window.close();
}
<% end if %>

function btnSave_onclick() {
	var tempstrchecked = "";
	
	if (!ValidateInput())
		return false;
		
	SelectAll(document.getElementById("lbxSelectedAliasID"));
	SelectAll(document.getElementById("lbxSelectedDivision"));
	
	thisform.ChangeMask.value = ChangeMask
	
	<% if Request.QueryString("nEditLocalization") <> "" Then%>
		thisform.action = "AMO_Save.asp?nEditLocalization=1";
	<% else %>
		thisform.action = "AMO_Save.asp?nEditLocalization=0";
	<% end if %>
		
	thisform.target = "";

	//need to enable the controls so the processing page can get the values
	if (document.getElementById("chkidp").disabled)
		document.getElementById("chkidp").disabled = false;
			
	if (document.getElementById("lbxCategory").disabled) {
		document.getElementById("lbxCategory").disabled = false; }

	// for some reason Firefox doesn't like always passing the category
	// Just put it in a hidden variable and the save page can use it if necessary
	if (document.getElementById("lbxCategory"))
		document.getElementById("lbxCategorySelection").value = document.getElementById("lbxCategory").value
		
	if (thisform.rdType != null) {
		for (var i=0; i<thisform.rdType.length; i++) {
			if (thisform.rdType[i].disabled) {
				thisform.rdType[i].disabled = false;
			}
		}
	}

	if (thisform.chkDivision != null) {
		for (i=0; i<thisform.chkDivision.length; i++) {
			if (thisform.chkDivision[i].disabled) {
				thisform.chkDivision[i].disabled = false;
			}
		}
	}
		
	if (thisform.txtDesc.disabled)
		thisform.txtDesc.disabled = false;

	if (thisform.txtShortDesc.disabled)
		thisform.txtShortDesc.disabled = false;

	if (thisform.txtBluePN.disabled)
		thisform.txtBluePN.disabled = false;

	if (thisform.txtReplacement.disabled)
		thisform.txtReplacement.disabled = false;

	if (thisform.txtAlternative.disabled)
		thisform.txtAlternative.disabled = false;

	if (thisform.txtNetWeight.disabled)
		thisform.txtNetWeight.disabled = false;

	if (thisform.txtAirPackedWeight.disabled)
		thisform.txtAirPackedWeight.disabled = false;

	if (thisform.txtAirPackedCubic.disabled)
		thisform.txtAirPackedCubic.disabled = false;

	if (thisform.txtExportCubic.disabled)
		thisform.txtExportCubic.disabled = false;

	if (thisform.lbxSelectedAliasID.disabled)
		thisform.lbxSelectedAliasID.disabled = false;
		
	if (thisform.txtNotes.disabled)
		thisform.txtNotes.disabled = false;

	if (thisform.chkMOLHide.disabled)
		thisform.chkMOLHide.disabled = false;
	<% if IsODM = 0 then %>
	if (thisform.txtAMOCost.disabled)
		thisform.txtAMOCost.disabled = false;

	if (thisform.txtAMOWWPrice.disabled)
		thisform.txtAMOWWPrice.disabled = false;

	if (thisform.txtActualCost.disabled)
		thisform.txtActualCost.disabled = false;
	<% end if %>
	if (thisform.oVisibilityNA.disabled)
		thisform.oVisibilityNA.disabled = false;
		
	if (thisform.oVisibilityEM.disabled)
		thisform.oVisibilityEM.disabled = false;
		
	if (thisform.oVisibilityAP.disabled)
		thisform.oVisibilityAP.disabled = false;
		
	if (thisform.oVisibilityLA.disabled)
		thisform.oVisibilityLA.disabled = false;
					
	<% if Request.QueryString("nEditLocalization") <> "" Then %>
		var coll = document.getElementsByName("chkHub");
		for (i=0;i< coll.length; i++) {
			if (coll[i].checked) {
				if(tempstrchecked == "") 
					tempstrchecked = coll[i].value;
				else
					tempstrchecked = tempstrchecked + "," + coll[i].value;
			}
		}
		
		document.getElementById("txtHubCheckboxlist").value = tempstrchecked;
		
		tempstrchecked = "";

		coll = document.getElementsByName("chkHideHub");
		for (i=0;i< coll.length; i++) {
			if (coll[i].checked) {
				if(tempstrchecked == "") 
					tempstrchecked = coll[i].value;
				else
					tempstrchecked = tempstrchecked + "," + coll[i].value;
			}
		}
		
		document.getElementById("txtHideHubCheckboxlist").value = tempstrchecked;
	<%end if%>
	
	return thisform.submit();
}

function formatCurrency(num) {
	num = num.toString().replace(/\$|\,/g,'');
	if(isNaN(num))
		num = "0";
	sign = (num == (num = Math.abs(num)));
	num = Math.floor(num*100+0.50000000001);
	cents = num%100;
	num = Math.floor(num/100).toString();
	if(cents<10)
		cents = "0" + cents;
	for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
		num = num.substring(0,num.length-(4*i+3))+','+
	num.substring(num.length-(4*i+3));
	return (((sign)?'':'-') + num + '.' + cents);
}

function calculateCPLBlindDate(RASDate) {
	var somedate = new Date(RASDate)
	var themonth = somedate.getMonth()
	var theday = somedate.getDate()
	var theyear = somedate.getFullYear()
	//1 day prior to GA date : SUG 9763,Vinutha
	somedate = new Date(theyear, themonth, theday-1)	
	return (somedate.getMonth() + 1) + '/' + somedate.getDate() + '/' + somedate.getFullYear();
}

function calculateObsoleteDate(RASDate) {
	var somedate = new Date(RASDate);
	var themonth = somedate.getMonth();
	var theday = somedate.getDate();
	var theyear = somedate.getFullYear();		
	
	// add 3 month to date	
	themonth = themonth + 4;
	var timeA = new Date(theyear,themonth,1)
	var timeB = new Date(timeA - (60*60*24*1000)); // subtract 1 day
	return (timeB.getMonth()+1) + '/' + timeB.getDate() + '/' + timeB.getFullYear();
}

function BOMRevADate_Update(thefield, strRegionID) {
		if (!checkDate(thefield, "PHweb (General) Availability (GA)", true)) {
			return false;
		}

		if(strRegionID == 334) {
		    alert("Modifying PHweb (General) Availability (GA) may cause discrepancy against the spec (Global Series Config EOL has to be in the range of PHweb (General) Availability (GA) and End of Manufacturing (EM))");
		}
		
		if (thefield.value.length > 0) {
			// calculate Select Availability (SA) only if not an empty date
			var cplObject = document.all.item("txtCPLBlindDate" + strRegionID);	
			cplObject.value = calculateCPLBlindDate(thefield.value)
			changeSum(cplObject, cplObject.name);
		}
		changeSum(thefield, thefield.name);
}

function RASDiscontinueDate_Update(thefield, strRegionID) {
		if (!checkDate (thefield, "End of Manufacturing (EM)", true))
			return false;

		if(strRegionID == 334) {
		    alert("Modifying End of Manufacturing (EM) may cause discrepancy against the spec (Global Series Config EOL has to be in the range of PHweb (General) Availability (GA) and End of Manufacturing (EM))");
		}
		
		if (thefield.value.length > 0) {
			// calculate Obsolete only if not an empty date
			var obdObject = document.all.item("txtObsoleteDate" + strRegionID); 
			obdObject.value = calculateObsoleteDate(thefield.value)
			changeSum(obdObject, obdObject.name);
		}
		changeSum(thefield, thefield.name);
}

function CPLBlindDate_Update(thefield, strRegionID) {
	if (!checkDate (thefield, "Select Availability (SA)", true))
		return false;
		
   changeSum(thefield, thefield.name);
}

function ObsoleteDate_Update(thefield, strRegionID) {
	if (!checkDate (thefield, "Obsolete Date", true))
		return false;
		
	changeSum(thefield, thefield.name);
}
<% if IsODM = 0 then %>
function AMOCost_Update(thefield) {
	// need to update AMO Price field now too by multiplying by 2
	var wwpObject = document.all.item("txtAMOWWPrice")
	var newValue = formatCurrency( stripCharsInBag(thefield.value, ',') * 2)
	if (newValue.length <= 20) {
		wwpObject.value = newValue
		ThisChanged(64)	// AMO Price changed
	}
}
<% end if %>
function btnEdit_onclick() {
	thisform.action = "AMO_Properties.asp?Edit=1&Mode=2&from=<%= strFrom %>&ID=<%= nID %>"
	return thisform.submit();
}

function btnClone_onclick() {
    <% if oRsCreateGroups.RecordCount > 1 then %>
    var MOLLink = <%=Application("IRSWebServer")%>;
	thisform.action = "Module/isapi/GetGroupsForModuleRole.asp?Edit=1&Mode=3&from=<%= strFrom %>&ID=<%= nID %>"
	<% else %>
	thisform.action = "../library/AMO_Properties.asp?nEditLocalization=1&Edit=1&Mode=3&from=<%= strFrom %>&ID=<%= nID %>"
	<% end if %>
	return thisform.submit();
}

function btnDelete_onclick() {
	if (!confirm("Are you sure you want to delete the After Market Option?"))
		return false;
	thisform.action = "AMO_Save.asp"
	thisform.nMode.value = "5"
	return thisform.submit();
}

function btnDisable_onclick() {
	var msg = "You are disabling this Option for future use and no further modifications will be allowed.\n"
	<% if bEdit then %>
	msg += "Any changes that may have been made on this page will be lost.\n"
	<% end if %>
	msg += "Click OK to continue or Cancel to stop the operation"
	if (!confirm(msg))
		return false;

	thisform.action = "AMO_Save.asp";
	thisform.StatusID.value = <%= Application("AMO_DISABLED") %>;
	return thisform.submit();
}

function btnEnable_onclick() {
	thisform.action = "AMO_Save.asp";
	thisform.StatusID.value = <%= Application("AMO_RE-ENABLED") %>;
	return thisform.submit();
}
<%
end if 'if nMode <> 4
%>
function rdType_onclick() {
	var bHW=true, bSW=true;

	if (thisform.rdType != null) {
		for (var i=0; i<thisform.rdType.length; i++) {
			if (thisform.rdType[i].checked) {
				if (thisform.rdType[i].value == <%=Application("MD_TYPE_HW")%>)
					bSW=false;
				else if (thisform.rdType[i].value == <%=Application("MD_TYPE_SW")%>)
					bHW=false;
			}
		}
		if (bHW)
			strInnerHTML = document.getElementById("divhwcateg").innerHTML;

		if (bSW)
			strInnerHTML = document.getElementById("divswcateg").innerHTML;

		var re = /lbxTempCategory/g;
		divcateg.innerHTML = strInnerHTML.replace(re, "lbxCategory");
	}
	return true;
}

function window_onload() {
	ChangeMask = <%= lngChangeMask %>;

	<% if sErr = "" then %>
	rdType_onclick();
	<% end if %>
}

function btnLocalization_onclick() {
  if (document.getElementById("btnLocalization").value == "+") {
		document.getElementById("oTable1").style.display="";
	document.getElementById("oTable").style.display="";
	document.getElementById("btnLocalization").value="-";
	//window.scroll(0, 700);
  }
  else {
		document.getElementById("oTable1").style.display="none";
	document.getElementById("oTable").style.display="none";
	document.getElementById("btnLocalization").value="+";
  }
  return true;
}

function editLocalization_onclick() {
	if (!confirm("Please be sure to save your data before edit the localization. Proceed ?"))
		return false;
		
	thisform.action = "AMO_Properties.asp?Edit=1&Mode=2&from=<%=strFrom%>&ID=<%=nID%>&nEditLocalization=1"
	return thisform.submit();
}

function btnPlatforms_onclick() {
  if (document.getElementById("btnPlatforms").value == "+") {
	document.getElementById("tbPlatforms").style.display="";
	document.getElementById("btnPlatforms").value="-";
  }
  else {
	document.getElementById("tbPlatforms").style.display="none";
	document.getElementById("btnPlatforms").value="+";
  }
  return true;
}

function btnPorDetails_onclick() {
  if (document.getElementById("btnPorDetails").value == "+") {
	document.getElementById("tbPorDetails").style.display="";
	document.getElementById("btnPorDetails").value="-";
  }
  else {
	document.getElementById("tbPorDetails").style.display="none";
	document.getElementById("btnPorDetails").value="+";
  }
  return true;
}

function btnCompatibility_onclick() {
  if (document.getElementById("btnCompatibility").value == "+") {
	document.getElementById("tbCompatibility").style.display="";
	document.getElementById("btnCompatibility").value="-";
  }
  else {
	document.getElementById("tbCompatibility").style.display="none";
	document.getElementById("btnCompatibility").value="+";
  }
  return true;
}

function btnRegion_onclick() {
  if (document.getElementById("btnRegion").value == "+") {
	document.getElementById("tbRegion").style.display="";
	document.getElementById("btnRegion").value="-";
  }
  else {
	document.getElementById("tbRegion").style.display="none";
	document.getElementById("btnRegion").value="+";
  }
  return true;
}

function OptionNA_Change() {
	if (document.getElementById("oVisibilityNA").value == 1) {
		document.getElementById("oVisibilityEM").value = 1;
		document.getElementById("oVisibilityAP").value = 1;
		document.getElementById("oVisibilityLA").value = 1;
	}
}

function OptionEM_Change() {
	if (document.getElementById("oVisibilityEM").value == 1) {
		document.getElementById("oVisibilityNA").value = 1;
		document.getElementById("oVisibilityAP").value = 1;
		document.getElementById("oVisibilityLA").value = 1;
	}
}

function OptionAP_Change() {
	if (document.getElementById("oVisibilityAP").value == 1) {
		document.getElementById("oVisibilityEM").value = 1;
		document.getElementById("oVisibilityNA").value = 1;
		document.getElementById("oVisibilityLA").value = 1;
	}
}

function OptionLA_Change() {
	if (document.getElementById("oVisibilityLA").value == 1) {
		document.getElementById("oVisibilityEM").value = 1;
		document.getElementById("oVisibilityAP").value = 1;
		document.getElementById("oVisibilityNA").value = 1;
	}
}

function GlobalseriesDate_Update(thefield, strRegionID) {
	var discontinueDate, mrAvailableDate, globalSeriesDate

	if (!checkDate (thefield, "Globalseries Date", true))
		return false;
		
	if(thefield.value != "" && thisform.txtBOMRevADate43.value == "") {
	    alert("Please enter PHweb (General) Availability (GA) before proceeding");	
		thefield.value = "";
		return false;
	}
	if(thefield.value != "" && thisform.txtRASDiscontinueDate43.value == "") {
		alert("Please enter End of Manufacturing (EM) before proceeding");
		thefield.value = "";
		return false;
	}
		
	mrAvailableDate = new Date(thisform.txtBOMRevADate43.value)
	discontinueDate = new Date(thisform.txtRASDiscontinueDate43.value)
	
	if(thefield.value != "") {
		globalSeriesDate = new Date(thefield.value)
			
		if(globalSeriesDate < mrAvailableDate || globalSeriesDate > discontinueDate) {
		    alert("Global Series Config EOL has to be in the range of PHeb (General) Availability (GA) and End of Manufacturing (EM)");	
			thefield.value = "";
			return false;
		}
	}
	
 changeSum(thefield, thefield.name);
}

function changeSum(thefield, sAMORegionID) {
	var sNewCPLLocaleString, sNewBOMLocaleString;
	var sNewRASLocaleString, sNewOBSLocaleString;
	var sGlobalSeriesLocaleString;
	var nAMORegionID;

	if (sAMORegionID.indexOf("txtCPLBlindDate") != -1)	
	{
		sNewCPLLocaleString = thisform.newCPLLocaleString.value;
		
		nAMORegionID = sAMORegionID.substring(15, sAMORegionID.length);
		// if the region is not in the string then add it
		if (sNewCPLLocaleString == "" || sNewCPLLocaleString.indexOf("," + nAMORegionID + ",") == -1) {
			sNewCPLLocaleString += "," + nAMORegionID + "," 
			thisform.newCPLLocaleString.value = sNewCPLLocaleString;
		}
	}
	else if (sAMORegionID.indexOf("txtBOMRevADate") != -1)	
	{
		sNewBOMLocaleString = thisform.newBOMLocaleString.value;
		nAMORegionID = sAMORegionID.substring(14, sAMORegionID.length);
		// if the region is not in the string then add it
		if (sNewBOMLocaleString == "" || sNewBOMLocaleString.indexOf("," + nAMORegionID + ",") == -1) {
			sNewBOMLocaleString += "," + nAMORegionID + "," 
			thisform.newBOMLocaleString.value = sNewBOMLocaleString;
		}
	}
	else if (sAMORegionID.indexOf("txtRASDiscontinueDate") != -1)	
	{
		sNewRASLocaleString = thisform.newRASLocaleString.value;
		nAMORegionID = sAMORegionID.substring(21, sAMORegionID.length);
		// if the region is not in the string then add it
		if (sNewRASLocaleString == "" || sNewRASLocaleString.indexOf("," + nAMORegionID + ",") == -1) {
			sNewRASLocaleString += "," + nAMORegionID + "," 
			thisform.newRASLocaleString.value = sNewRASLocaleString;
		}
	}
	else if (sAMORegionID.indexOf("txtObsoleteDate") != -1)
	{
		sNewOBSLocaleString = thisform.newOBSLocaleString.value;
		nAMORegionID = sAMORegionID.substring(15, sAMORegionID.length);
		// if the region is not in the string then add it
		if (sNewOBSLocaleString == "" || sNewOBSLocaleString.indexOf("," + nAMORegionID + ",") == -1) {
			sNewOBSLocaleString += "," + nAMORegionID + "," 
			thisform.newOBSLocaleString.value = sNewOBSLocaleString;
		}
	}
	else if (sAMORegionID.indexOf("txtGlobalseriesDate") != -1)
	{
		thisform.newGBLLocaleString.value = thefield.value;
	}	
}
//-->
</SCRIPT>
</HEAD>
<BODY bgcolor="#FFFFFF" onLoad="return window_onload()">
<% 
if nMode = 4 then
	'InsertGlobalNavigationBar_HomeParm(True)
	'Response.write BuildHelp(sHeader, "")
else
	
	'Response.write BuildHelp(sHeader, sHelpfile)
end if

if sErr <> "" then
	Response.Write sErr
else
	%>
    <h1 class="page-title"><%=sHeader%></h1>
	<TABLE border=0 cellPadding=0 cellSpacing=5 width="100%">
	<TR><TD colspan='2'>
		<FORM name=thisform method=post>
		<% if nMode = 2 or nMode = 3 or nMode = 4 then %>
			<% if not bEdit and bAMOUpdate and nAMOStatusID <> Application("AMO_DISABLED") and nMode <> 4  and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then %>
				<INPUT id=btnEdit name=btnEdit type=button value="Modify" LANGUAGE=javascript onClick="return btnEdit_onclick()">
			<% elseif not bEdit and bCostUpdate and nAMOStatusID <> Application("AMO_DISABLED") and nMode <> 4 and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE"))  then%>
				<INPUT id=btnEdit name=btnEdit type=button value="Modify" LANGUAGE=javascript onClick="return btnEdit_onclick()">
		<% end if

			if nMode <> 4 then

				if bEdit then
					if bUpdate then
						response.write "<INPUT id=btnSave name=btnSave type=button value=""Save"" LANGUAGE=javascript onclick=""return btnSave_onclick()"">" & vbCrLf
					end if
				end if
				
				if bAMOUpdate and nMode = 2 and clng(nAMOStatusID) <> clng(Application("AMO_DISABLED"))and  clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then
					response.write "<INPUT id=btnDisable name=btnDisable type=button value=""Disable"" LANGUAGE=javascript onclick=""return btnDisable_onclick()"">" & vbCrLf
				end if
	
				if bAMOUpdate and nMode = 2 and clng(nAMOStatusID) = clng(Application("AMO_DISABLED")) then
					response.write "<INPUT id=btnEnable name=btnEnable type=button value=""Enable"" LANGUAGE=javascript onclick=""return btnEnable_onclick()"">" & vbCrLf
				end if
	
				if bAMODelete and nMode = 2 and clng(nAMOStatusID) = clng(Application("AMO_NEW")) then
					response.write "<INPUT id=btnDelete name=btnDelete type=button value=""Delete"" LANGUAGE=javascript onclick=""return btnDelete_onclick()"">" & vbCrLf
				end if
	
				if oRsCreateGroups.RecordCount > 0 and nMode = 2 then
					response.write "<INPUT id=btnClone name=btnClone type=button value=""Clone"" LANGUAGE=javascript onclick=""return btnClone_onclick()"">" & vbCrLf
				end if
	
				if bFromAMO then
					response.write "<INPUT id=btnCancel name=btnCancel type=button value=""Close"" LANGUAGE=javascript onclick=""return btnCancel_onclick()"">" & vbCrLf
				end if
			end if
			%>	
		  <% if (bAMOUpdate or bCostUpdate or bPORCreate) then%>
			</td>
			<td align=right>
				<A align=right style="color:blue" href="AMO_ExcelPropertiesReport.asp?ID=<%= nID %>&TypeId=<%= nTypeId %>&DivisionTarget=<%= sDivisionTarget %>">Report</A>
			</td>
		  <%end if %>
			<td align=right width="10%">
				<A align=right style="color:blue" href="<%=Application("IRSWebServer") %>MOL/Module/isapi/WhereUsed.asp?ID=<%= nID %>">Where Used</A>
		<% end if %>
		</TD></TR>

	<% 
	if nMode <> 4 then 
	
		if nAMOStatusID = Application("AMO_RASREVIEW") or nAMOStatusID = Application("AMO_RASUPDATE") then %>
		<TR>
			<TD colspan=2>
				<img src="../library/images/gifs/lock.gif" align="top" ID="img1" name="img1" border="0"  WIDTH="10" HEIGHT="10">
				<%
				if nAMOStatusID = Application("AMO_RASREVIEW") then 
				%>
				  <font color="blue"><b>Some of the After Market Option is locked and cannot be modified because the status is RAS Review </font>					
				<%else%>
				  <font color="blue"><b>Some of the After Market Option is locked and cannot be modified because the status is RAS Update </font>					
				<%end if%>
			</TD></TR>
		<% 
		end if 
	end if
	%>

	<TR>
		<TD width="20%">Marketing Description<font color=red>*</font></TD>
		<TD><INPUT onChange="ThisChanged(32768)" id=txtDesc name=txtDesc value="<%=server.htmlencode(sDesc)%>" size=60 maxlength=100 <%=sNameCtrlStyle%>>
			<font size=1><i>100 maximum characters. <%=sNameComment%></i></font></td></TR>

	<TR>
		<TD>Short Description<font color=red>*</font></TD>
		<TD><INPUT onChange="ThisChanged(65536)" id=txtShortDesc name=txtShortDesc value="<%=server.htmlencode(sShortDesc)%>" size=64 maxlength=40 <%=sRasCtrlStyle%>>
			<font size=1><i>40 maximum characters</i></font></td></tr>
	
	<TR>
		<TD>Long Description</TD>
		<TD colspan=2 valign="middle">
			<table border=0>
				<tr>
					<td><textarea cols="50" rows="4" onChange="ThisChanged(134217728)" id="txtLongDescription" name="txtLongDescription" <%=sRasCtrlStyle%>><%=strLongDescription%></textarea></td>
					<td><font size=1><i>160 maximum characters</i></font></td>
				</tr>
			</table>
			</TD></TR>  
	<TR>
		<TD>Option Type<font color=red>*</font></TD>
		<TD><b><%=sModuleTypeHTML%></b></TD></TR>

	<TR>
		<TD>Option Category<font color="red">*</font></TD>
		<TD>
			<DIV id=divhwcateg name=divhwcateg style="display:none"><%=sHWModuleCategoryHTML%></DIV>
			<DIV id=divswcateg name=divswcateg style="display:none"><%=sSWModuleCategoryHTML%></DIV>
			<DIV id=divcateg name=divcateg></DIV></TD></TR>

	<TR>
		<TD>Target Business Segment<font color="red">*</font></TD>
		<TD><b><%=sModuleDivisionHTML%></b></TD></TR>

	<TR>
		<TD>HP Part Number<font color=red>*</font></TD>
		<TD><INPUT onChange="ThisChanged(1)" id=txtBluePN name=txtBluePN value="<%=server.htmlencode(sBluePN)%>" size=20 maxlength=20 <%=sRasCtrlStyle%>>
			<font size=1><i>20 maximum characters</i></font></td></TR>

	
	<% if IsODM = 0 then %>
	<TR> 
		<TD>AMO Cost</TD>
		<!-- OnChange isn't kicked off so using onBlur instead -->
		<td>
			<table border=0 cellpadding=0 cellspacing=0>
			<tr>
				<td>
					<INPUT onBlur="ThisChanged(32);this.value=formatCurrency(this.value);AMOCost_Update(this)"
					onKeyPress="return(currencyFormat(this, ',', '.', event, 20))"
					id=txtAMOCost name=txtAMOCost value="<%=sAMOCost%>" size=20 maxlength=20 <%=sCostCtrlStyle%>></td>
				<td style="padding-left: 4px;">
					<font size=1><i>20 maximum characters. If changed, AMO Price is set equal to AMO Cost times 2 unless greater than 20 characters.</i></font></td>
			</tr>
			</table>
			</td></TR>

	<TR>
		<TD>AMO Price</TD>
		<TD><INPUT onBlur="ThisChanged(64);this.value=formatCurrency(this.value)"
			onKeyPress="return(currencyFormat(this, ',', '.', event, 20))"
			id=txtAMOWWPrice name=txtAMOWWPrice value="<%=sAMOWWPrice%>" size=20 maxlength=20 <%=sCostCtrlStyle%>>
			<font size=1><i>20 maximum characters</i></font></td></TR>

	<TR>
		<TD>Actual Cost</TD>
		<TD><INPUT onBlur="ThisChanged(131072);this.value=formatCurrency(this.value)"
			onKeyPress="return(currencyFormat(this, ',', '.', event, 20))"
			id=txtActualCost name=txtActualCost value="<%=sActualCost%>" size=20 maxlength=20 <%=sCostCtrlStyle%>>
			<font size=1><i>20 maximum characters</i></font></td></TR>
	<% end if %>
	<TR>
		<TD>Replacement</TD>
		<TD><INPUT onChange="ThisChanged(128)" id=txtReplacement name=txtReplacement value="<%=server.htmlencode(sReplacement)%>" size=30 maxlength=30 <%=sRasCtrlStyle%>>
			<font size=1><i>30 maximum characters</i></font></td></TR>

	<TR>
		<TD>Alternative</TD>
		<TD><INPUT onChange="ThisChanged(256)" id=txtAlternative name=txtAlternative value="<%=server.htmlencode(sAlternative)%>" size=30 maxlength=30 <%=sRasCtrlStyle%>>
			<font size=1><i>30 maximum characters</i></font></td></TR>

	<TR>
		<TD>Net Weight</TD>
		<TD><INPUT onChange="ThisChanged(512)" id=txtNetWeight name=txtNetWeight value="<%=sNetWeight%>" size=9 maxlength=9 <%=sRasCtrlStyle%>>
			<font size=1><i>9 maximum characters</i></font></td></TR>

	<TR>
		<TD>Export Weight</TD>
		<TD><INPUT onChange="ThisChanged(1024)" id=txtExportWeight name=txtExportWeight value="<%=sExportWeight%>" size=9 maxlength=9 <%=sRasCtrlStyle%>>
			<font size=1><i>9 maximum characters</i></font></td></TR>

	<TR>
		<TD>Air Packed Weight</TD>
		<TD><INPUT onChange="ThisChanged(2048)" id=txtAirPackedWeight name=txtAirPackedWeight value="<%=sAirPackedWeight%>" size=9 maxlength=9 <%=sRasCtrlStyle%>>
			<font size=1><i>9 maximum characters</i></font></td></TR>

	<TR>
		<TD>Air Packed Cubic</TD>
		<TD><INPUT onChange="ThisChanged(4096)" id=txtAirPackedCubic name=txtAirPackedCubic value="<%=sAirPackedCubic%>" size=9 maxlength=9 <%=sRasCtrlStyle%>>
			<font size=1><i>9 maximum characters</i></font></td></TR>

	<TR>
		<TD>Export Cubic</TD>
		<TD><INPUT onChange="ThisChanged(8192)" id=txtExportCubic name=txtExportCubic value="<%=sExportCubic%>" size=9 maxlength=9 <%=sRasCtrlStyle%>>
			<font size=1><i>9 maximum characters</i></font></td></TR>
			
	<TR>
		<TD>Warranty Code</TD>
		<TD><INPUT onChange="ThisChanged(16777216)" id=txtWarrantyCode name=txtWarrantyCode value="<%=server.htmlencode(Trim(sWarrantyCode))%>" size=5 maxlength=5 <%=sRasCtrlStyle%>>
			<font size=1><i>5 maximum characters</i></font></td></TR>
			
	<TR>
		<TD>Country of Manufacture</TD>
		<TD><INPUT onChange="ThisChanged(67108864)" id=txtManufactureCountry name=txtManufactureCountry value="<%=server.htmlencode(Trim(sManufactureCountry))%>" size=2 maxlength=2 <%=sRasCtrlStyle%>>
			<font size=1><i>2 maximum characters</i></font></td></TR>
			
		<%	dim strexpand
			dim strshowradio
			if Request.QueryString("nEditLocalization") <> 1  then
				strexpand = "+"
				strshowradio = "display:none"
			else
				strexpand = "-"
				strshowradio = ""
			end if 	  %>

	<tr><td colspan=4><hr></td></tr>
	<TR>
	  <TD><a name="editlocalization"></a><font size='3'><strong>Localization&nbsp;</strong></font>
		<input id="btnLocalization" language="javascript" name="btnLocalization" type="button" style="HEIGHT:22px;WIDTH:22px" value="<%=strexpand%>" onClick="return btnLocalization_onclick()" >
	  </TD>
	</TR> 
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td><table id='oTable1' name='oTable1' align='left' style='<%=strshowradio%>' border='0'>
		<TR>
			<TD>Localized : </TD>
			<TD colspan=2>
			<% if sLocalized > 0 then %>
				<INPUT type='radio' id=rdLocalized name=rdLocalized  value="True" <%=sDisabledCtrlStyle%> checked >Yes
				<INPUT type='radio' id=rdLocalized name=rdLocalized  value="False" <%=sDisabledCtrlStyle%>>No
			<%else%>
				<INPUT type='radio' id=rdLocalized name=rdLocalized  value="True" <%=sDisabledCtrlStyle%>>Yes
				<INPUT type='radio' id=rdLocalized name=rdLocalized  value="False" <%=sDisabledCtrlStyle%> checked>No
			<%end if %>
			</TD>
		</TR>
		 <tr><td>&nbsp;</td></tr>
		
		<% if Request.QueryString("nEditLocalization") <> 1 and sCtrlStyle = "" and sRasCtrlStyle = "" then %>
		<tr>
			<td colspan=3><INPUT type='button' value='Edit Localization' id=btnEditLocale name=btnEditLocale LANGUAGE=javascript onClick="return editLocalization_onclick();" >
			</td>
		</tr>
		<%end if %>
		</table>
		</td>
	</tr>
	
   
	<TR><TD valign=middle colspan=3>
		<%
		sErr = WritePropertiesLocalizationGridHTML(Cstr(sDivisionIDs), Cstr(nID), oRsCheckedRegion, Cstr(sRasCtrlStyle), nMode, sHubCheckboxlist, sHideHubCheckboxlist)
		if nMode = 3 then
			If oRsCheckedRegion.RecordCount <> 0 Then
				oRsCheckedRegion.MoveFirst
				do while (not oRsCheckedRegion.EOF)
					if Trim(oRsCheckedRegion("CPLBlindDate").value) <> "" Then
						sNewCPLLocaleString =sNewCPLLocaleString & "," & oRsCheckedRegion("RegionID").value & ","
					end if
					  
					if Trim(oRsCheckedRegion("BOMRevADate").value) <> "" Then
						sNewBOMLocaleString = sNewBOMLocaleString  & "," & oRsCheckedRegion("RegionID").value & ","
					end if
					  
					if Trim(oRsCheckedRegion("RASDiscontinueDate").value) <> "" Then
						sNewRASLocaleString = sNewRASLocaleString  & "," & oRsCheckedRegion("RegionID").value & ","
					end if
						
					if Trim(oRsCheckedRegion("ObsoleteDate").value) <> "" Then
						sNewOBSLocaleString = sNewOBSLocaleString  & "," & oRsCheckedRegion("RegionID").value & ","
					end if
					oRsCheckedRegion.MoveNext
				loop
			End If
		end if
		%> 
		</TD>
	</TR>
	<tr><td colspan=4><hr></td></tr>
	<TR>    
			<TD><font size='3'><strong>SKU Visibility&nbsp;&nbsp;</strong></font>
				<input id="btnRegion" language="javascript" name="btnRegion" type="button" style="HEIGHT:22px;WIDTH:22px" value="+" onClick="return btnRegion_onclick()">
			</TD>
	</TR>
	<tr><td></td></tr><tr><td></td></tr>
	<tr><td>
		<table id='tbRegion' border=0 CELLSPACING=2 CELLPADDING=2 style='display:none' width="100%">
	
			<tr>
				<td>NA&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>
					<select onchange='javascript:OptionNA_Change()' id='oVisibilityNA' name='oVisibilityNA'  style = "width:80px" <%=sRasDisCtrlStyle%>>
					<% if sVisibility_NA = "Global" then %>
						<option value=0></option>
						<option value=1 selected>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_NA = "Blind" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2 selected>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_NA = "Full" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3 selected>Full</option>
					<%else%>
						<option value=0 selected> </option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%end if%>					
					</select>
				</td>
			</tr>
			<tr>
				<td>EMEA&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>
					<select onchange='javascript:OptionEM_Change()' id='oVisibilityEM' name='oVisibilityEM' style = "width:80px" <%=sRasDisCtrlStyle%>>
					<% if sVisibility_EM = "Global" then %>
						<option value=0></option>
						<option value=1 selected>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_EM = "Blind" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2 selected>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_EM = "Full" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3 selected>Full</option>
					<%else%>
						<option value=0 selected> </option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%end if%>
					</select>
				</td>
			</tr>
	
			<tr>
				<td>APJ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>
					<select onchange='javascript:OptionAP_Change()' id='oVisibilityAP' name='oVisibilityAP' style = "width:80px" <%=sRasDisCtrlStyle%>>
					<% if sVisibility_AP = "Global" then %>
						<option value=0></option>
						<option value=1 selected>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_AP = "Blind" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2 selected>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_AP = "Full" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3 selected>Full</option>
					<%else%>
						<option value=0 selected> </option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%end if%>					
					</select>
				</td>
			</tr>
	
			<tr>
				<td>LA&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
				<td>
					<select onchange='javascript:OptionLA_Change()' id='oVisibilityLA' name='oVisibilityLA' style = "width:80px" <%=sRasDisCtrlStyle%>>
					<% if sVisibility_LA = "Global" then %>
						<option value=0></option>
						<option value=1 selected>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_LA = "Blind" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2 selected>Blind</option>
						<option value=3>Full</option>
					<%elseif sVisibility_LA = "Full" then %>
						<option value=0></option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3 selected>Full</option>
					<%else%>
						<option value=0 selected> </option>
						<option value=1>Global</option>
						<option value=2>Blind</option>
						<option value=3>Full</option>
					<%end if%>					
					</select>
				</td>
			</tr>
			
		</table>
	</td></tr>
	<tr><td colspan=4><hr></td></tr>
	<TR>    
			<TD><font size='3'><strong>Platforms&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font>
				<input id="btnPlatforms" language="javascript" name="btnPlatforms" type="button" style="HEIGHT:22px;WIDTH:22px" value="+" onClick="return btnPlatforms_onclick()" >
			</TD>
	</TR> 
	<TR>
		<td valign=middle colspan=3>
			<table id='tbPlatforms' border=0 CELLSPACING=0 CELLPADDING=0 style='display:none'>
				<TR>
					<TD align=left>
						<%
						'display the dual list box
						
						
						if bUpdate and nMode <> 4 and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) _
								and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then
						
							Call DualListboxRs_GetHTML6_Write( oRsPlatforms, "Description_Alias", "AliasID", oSelectedRs, "FullName", "AliasID", _
								False, False, "", _
								"Available Platforms", "Selected Platforms", "AliasID", True, _
								250, 400, not ((bAMOUpdate or bAMOCreate) and bEdit), True, False, 500, 12 )
								
						elseif (not bAMOUpdate) and bEditPlatform and nMode <> 4 and clng(nAMOStatusID) <> clng(Application("AMO_DISABLED")) _
																and  clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) _
																and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then
													
							set oSelectedRs = GetPlatformByDivision(oSelectedRs, sUserDivisions)
										
			
							Call DualListboxRs_GetHTML6_Write( oRsPlatforms, "Description_Alias", "AliasID", oSelectedRs, "FullName", "AliasID", _
								False, False, "", _
								"Available Platforms", "Selected Platforms", "AliasID", True, _
								250, 400, False, True, False, 500, 12 )
								
								Response.Write "<p>If the module is owned by different group, the platforms for that group may exist, but they are not displayed here "
						else
						
							Call DualListboxRs_GetHTML6_Write( oRsPlatforms, "Description_Alias", "AliasID", oSelectedRs, "FullName", "AliasID", _
								False, False, "", _
								"Available Platforms", "Selected Platforms", "AliasID", True, _
								250, 400, True, True, False, 500, 12 )
						end if
						%>
					</td></tr>
			</table>
			</td>
	</tr>
	
	<tr><td colspan=4><hr></td></tr>
	<TR>    
			<TD><font size='3'><strong>Compatibility&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</strong></font>
				<input id="btnCompatibility" language="javascript" name="btnCompatibility" type="button" style="HEIGHT:22px;WIDTH:22px" value="+" onClick="return btnCompatibility_onclick()" >
			</TD>
	</TR> 
	<TR>
		<td valign=middle colspan=3>
			<table id='tbCompatibility' border=0 CELLSPACING=2 CELLPADDING=2 style='display:none'>
	
				<tr>
					<td width=20%>Business Segments</td>
					<td width=80%>
					<%	
						if bUpdate and nMode <> 4 then
							DualListboxRs_GetHTML6_Write oRsBusSeg, "Description", "CategoryID", oRsBusSegSelected, _
							"Division", "DivisionID", true, true, sComparatibilitySelected, "Available", "Selected", _
							"Division", true, 130, 250, false, true, false, 350, 13	
						else
							DualListboxRs_GetHTML6_Write oRsBusSeg, "Description", "CategoryID", oRsBusSegSelected, _
							"Division", "DivisionID", true, true, sComparatibilitySelected, "Available", "Selected", _
							"Division", true, 130, 250, True, true, false, 350, 13	
						end if
					%>
					</td>
				</tr>
			</table>
			</td>
	</tr>
	
	
	 <tr><td colspan=4><hr></td></tr>
   
	
	<% if (bAMOUpdate or bCostUpdate or bPORCreate) then %>
	<TR>
	  <TD><font size='3'><strong>POR Details&nbsp;&nbsp;</strong></font>
	  <input id="btnPorDetails" language="javascript" name="btnPorDetails" type="button" style="HEIGHT:22px;WIDTH:22px" value="+" onClick="return btnPorDetails_onclick()">
	  </TD>
	</TR>
	
	<% else %>
	 <TR>
	  <TD><font size='3'><strong>POR Details&nbsp;&nbsp;</strong></font>
	  <input id="btnPorDetails" language="javascript" name="btnPorDetails" type="button" style="HEIGHT:22px;WIDTH:22px" value="+" onClick="return btnPorDetails_onclick()" <%=sDisabledCtrlStyle%>>
	  </TD>
	</TR>
	<% end if %>
	
	<tr>
		<td valign=middle colspan=3>
			<table id='tbPorDetails' border=0 CELLSPACING=2 CELLPADDING=2 style='display:none'>
				<TR><td>&nbsp;</td></TR>
				<TR>
					<TD>This product replaces AMO Part Number&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT id=txtAMOPartNoRe name=txtAMOPartNoRe value="<%=server.htmlencode(sAMOPartNoRe)%>" size=30 maxlength=30 <%=sCtrlStyle%>>
					<font size=1><i>30 maximum characters</i></font></td>
				</TR>
							
				
				<TR><td>&nbsp;</td></TR>
				<% if IsODM = 0 then %>
				<TR>
					<TD>Target Lifetime Volume - North America <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtTargetNA name=txtTargetNA value="<%=intTargetNA %>" size=10 maxlength=10 <%=sCtrlStyle%>>
					<font size=1><i>10 maximum characters</i></font></td>
				</TR>
				<TR>
					<TD>Target Lifetime Volume - Latin America <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtTargetLA name=txtTargetLA value="<%=intTargetLA %>" size=10 maxlength=10 <%=sCtrlStyle%>>
					<font size=1><i>10 maximum characters</i></font></td>
				</TR>
				<TR>
					<TD>Target Lifetime Volume - EMEA <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtTargetEMEA name=txtTargetEMEA value="<%=intTargetEMEA %>" size=10 maxlength=10 <%=sCtrlStyle%>>
					<font size=1><i>10 maximum characters</i></font></td>
				</TR>
				<TR>
					<TD>Target Lifetime Volume - APJ <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtTargetAPJ name=txtTargetAPJ value="<%=intTargetAPJ %>" size=10 maxlength=10 <%=sCtrlStyle%>>
					<font size=1><i>10 maximum characters</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Target Lifetime Volume - Total All Regions&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT id=txtTargetAllReg name=txtTargetAllReg value="<%=intTargetAllReg %>" size=10 maxlength=10 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>				
				
				<TR><td>&nbsp;</td></TR>			
				
				<TR>
					<TD>Burden (%) <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtBurdenPer name=txtBurdenPer value="<%=intBurdenPer %>" size=4 maxlength=4 <%=sCtrlStyle%>>
					<font size=1><i>4 maximum characters (xx.x)</i></font></td>
				</TR>
	
				<TR>
					<TD>Contra (%) <font color=green>*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD><INPUT onKeyPress="return checkNumeric()" id=txtContraPer name=txtContraPer value="<%=intContraPer %>" size=4 maxlength=4 <%=sCtrlStyle%>>
					<font size=1><i>4 maximum characters (xx.x)</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Burden&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtBurden name=txtBurden value="<%=intBurden %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Contra&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtContra name=txtContra value="<%=intContra %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Net Revenue&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtNetRevenue name=txtNetRevenue value="<%=intNetRevenue %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Dealer Margin at 6%&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtDealerMargin name=txtDealerMargin value="<%=intDealerMargin %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Target Cost + Burden&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtTargetCost name=txtTargetCost value="<%=intTargetCostBurden %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Gross Margin&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtGrossMargin name=txtGrossMargin value="<%=intGrossMargin %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Gross Margin %&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtGrossMarginPer name=txtGrossMarginPer value="<%=intGrossMarginPer %>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				<TR>
					<TD bgColor="#C0C0C0">Lifetime Gross Margin&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD bgColor="#C0C0C0"><INPUT  id=txtLifetimeGrossMargin name=txtLifetimeGrossMargin value="<%=intLifetimeGrossMargin%>" size=12 maxlength=12 disabled>
					<font size=1><i>Disabled result field</i></font></td>
				</TR>
				
				<TR><td>&nbsp;</td></TR>
				
				<TR>
					<TD >Provide Low Volume or Low Margin Justification&nbsp;&nbsp</TD>
					<TD colspan=2 valign="middle">
						<table border=0>
							<tr>
								<td><textarea cols="50" rows="10" id="txtJustification" name="txtJustification" <%=sCtrlStyle%>><%=server.htmlencode(sJustificationnotes) %></textarea></td>
								<td><font size=1><i>300 maximum characters</i></font></td>
							</tr>
						</table>
					</TD>
				</TR>
				<TR><td>&nbsp;</td></TR>
				<TR>
					<TD align="left">
					<input id="btnClear" language="javascript" name="btnClear" type="button" style="HEIGHT:22px;WIDTH:50px" value="Clear" onClick="return btnClear_onclick()" <%=sDisabledCtrlStyle%>>
					<input id="btnCalculate" language="javascript" name="btnCalculate" type="button" style="HEIGHT:22px;WIDTH:100px" value="Calculate" onClick="return btnCalculate_onclick()" <%=sDisabledCtrlStyle%>>
					</TD>
				</TR>
				<TR><td>&nbsp;</td></TR>
				<TR>
					<td align="left" colspan=3>(<font color="green">*</font>) Required fields for POR Details</td></tr>
				<% end if %>
			</table>
		</td>
	</tr>
	<tr><td colspan=4><hr></td></tr>
	<TR>
		<TD>Status</TD>
		<TD colspan=2><strong><%= sAMOStatus %></strong></TD></TR>

	<TR>
		<TD>Owned by</TD>
		<TD colspan=2><%
			if strOwnerHTML = "" or nMode = 4 then
				'user doesn't have the rights to change the owner group or they are not in more than one group
				'so just list the name of the group
				%>
				<% if nMode = 3 then %>
					<b><% = sCloneGroupName %></b>
				<%else %>
					<b><% = sGroupName %></b>
				<%end if %>
				<INPUT type="hidden" id="lbxGroupID" name="lbxGroupID" value="<%= cstr(lngGroupID) %>">
				<%
			else
				response.write strOwnerHTML
			end if%></TD></TR>
			
	<TR>
		<TD>Rules Description</TD>
		<TD colspan=2 valign="middle">
			<table border=0>
				<tr>
					<td><textarea cols="50" rows="10" onChange="ThisChanged(2)" id="txtRuleDescription" name="txtRuleDescription" <%=sRasCtrlStyle%>><%=strRuleDescription%></textarea></td>
					<td><font size=1><i>1024 maximum characters</i></font></td>
				</tr>
			</table>
			</TD></TR>
			
	<TR>
		<TD>Notes</TD>
		<TD colspan=2 valign="middle">
			<table border=0>
				<tr>
					<td><textarea cols="50" rows="10" id="txtNotes" name="txtNotes" <%=sRasCtrlStyle%>><%= strNotes %></textarea></td>
					<td><font size=1><i>1024 maximum characters</i></font></td>
				</tr>
			</table>
			</TD></TR>
			
	<TR>
		<TD>Order Instructions</TD>
		<TD colspan=2 valign="middle">
			<table border=0>
				<tr>
					<td><textarea cols="50" rows="7" onChange="ThisChanged(268435456)" id="txtOrderInstructions" name="txtOrderInstructions" <%=sRasCtrlStyle%>><%=strOrderInstructions%></textarea></td>
					<td><font size=1><i>600 maximum characters</i></font></td>
				</tr>
			</table>
			</TD></TR>  
	<TR>
	
	<TR>
		<TD>Replacement AV Description</TD>
		<TD colspan=2 valign="middle">
			<table border=0>
				<tr>
					<td><textarea cols="50" rows="4" onChange="ThisChanged(536870912)" id="txtReplacementDescription" name="txtReplacementDescription" <%=sRasCtrlStyle%>><%=strReplacementDescription%></textarea></td>
					<td><font size=1><i>80 maximum characters</i></font></td>
				</tr>
			</table>
			</TD></TR>  
	<TR>
	
				
	<TR>
		<TD>Product Line</TD>
		
		<TD><select onchange="ThisChanged(1073741824)" id='lbxProductLine' name='lbxProductLine' style = "width:400px" <%=sDisabledCtrlStyle%>>
		<option value=0></option>
		<% do while (not oRsProductLine.EOF)
			if nProductLineID =  oRsProductLine("CategoryID").Value then %>
				<option value=<%=oRsProductLine("CategoryID").Value%> selected><%=oRsProductLine("Description").Value%></option>
			<%else%>
				<option value=<%=oRsProductLine("CategoryID").Value%>><%=oRsProductLine("Description").Value%></option>
			<%end if
			oRsProductLine.MoveNext
		Loop
		%>
		'fix issue 5575
		<%if sDisabledCtrlStyle<> "" then
				sProductline_Ori = nProductLineID
			end if 	
		%>
			
		
		</TD>
	</TR>

	<TR>
		<TD>Ignore SCL Deployment Plan</TD>
		<TD><input type="checkbox" name="chkidp" id="chkidp" value="1" <%=sDisabledIDP%> 
		<%if bIDP then response.write " checked " %>><b>Ignore Deployment Plan</b></TD></TR>
	<TR>
		<TD>Module and Option List (MOL)</TD>
		<TD><input type="checkbox" name="chkMOLHide" id="chkMOLHide" value="1" <%=sDisabledCtrlStyle%> 
		<%if lngMOLHide = 1  then response.write " checked " %>><b>Hide from PRL</b></TD></TR>
			
	<TR>
		<TD>Site Commit List (SCL)</TD>
		<TD><input type="checkbox" name="chkSCLHide" id="chkSCLHide" value="1" <%=sDisabledCtrlStyle%> <%
			if lngSCLHide then response.write " checked " %>><b>Hide from SCL</b></TD></TR>
			
	<TR>
		<TD>SCM Report</TD>
		<TD><input type="checkbox" name="chkSCMHide" id="chkSCMHide" value="1" <%=sDisabledCtrlStyle%> <%
			if lngSCMHide then response.write " checked " %>><b>Hide from SCM</b></TD></TR>



	<% if nMode = 2 or nMode = 4 then %>
	<TR>
		<TD>Created by </TD>
		<TD colspan=2><b><%=sCreator%> on <%=sCreatedDate%></TD></TR>
	<TR>
		<TD>Last Updated by</TD>
		<TD colspan=2><b><%=sUpdater%> on <%=sUpdatedDate%></TD></TR>
	<% end if %>
	
	<TR>
		<TD colspan=3 align=left><br>
			<%

			if nMode <> 4 then

				if bEdit then
					if bUpdate then
						response.write "<INPUT id=btnSave name=btnSave type=button value=""Save"" LANGUAGE=javascript onclick=""return btnSave_onclick()"">" & vbCrLf
						bSave = True
					end if
				else
					if (bAMOUpdate or bCostUpdate) and nAMOStatusID <> Application("AMO_DISABLED") and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) then
						response.write "<INPUT id=btnEdit name=btnEdit type=button value=""Modify"" LANGUAGE=javascript onclick=""return btnEdit_onclick()"">" & vbCrLf
						bSave = True
					end if
					
				end if
				
				if Not bSave and bEditPlatform and nAMOStatusID <> Application("AMO_DISABLED")  and clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then					
					response.write "<INPUT id=btnSave name=btnSave type=button value=""Save"" LANGUAGE=javascript onclick=""return btnSave_onclick()"">" & vbCrLf
					bSave = True
				end if
				
				if Not bSave and bBusEnabled then
					response.write "<INPUT id=btnSave name=btnSave type=button value=""Save"" LANGUAGE=javascript onclick=""return btnSave_onclick()"">" & vbCrLf
				end if
				
				if bAMOUpdate and nMode = 2 and clng(nAMOStatusID) <> clng(Application("AMO_DISABLED"))and  clng(nAMOStatusID) <> clng(Application("AMO_RASREVIEW")) and clng(nAMOStatusID) <> clng(Application("AMO_RASUPDATE")) then
					response.write "<INPUT id=btnDisable name=btnDisable type=button value=""Disable"" LANGUAGE=javascript onclick=""return btnDisable_onclick()"">" & vbCrLf
				end if
	
				if bAMOUpdate and nMode = 2 and clng(nAMOStatusID) = clng(Application("AMO_DISABLED")) then
					response.write "<INPUT id=btnEnable name=btnEnable type=button value=""Enable"" LANGUAGE=javascript onclick=""return btnEnable_onclick()"">" & vbCrLf
				end if
	
				if bAMODelete and nMode = 2 and clng(nAMOStatusID) = clng(Application("AMO_NEW")) then
					response.write "<INPUT id=btnDelete name=btnDelete type=button value=""Delete"" LANGUAGE=javascript onclick=""return btnDelete_onclick()"">" & vbCrLf
				end if
	
				if oRsCreateGroups.RecordCount > 0 and nMode = 2 then
					response.write "<INPUT id=btnClone name=btnClone type=button value=""Clone"" LANGUAGE=javascript onclick=""return btnClone_onclick()"">" & vbCrLf
				end if
	
				if bFromAMO then
					response.write "<INPUT id=btnCancel name=btnCancel type=button value=""Close"" LANGUAGE=javascript onclick=""return btnCancel_onclick()"">" & vbCrLf
				end if
			end if
			%>
		</TD></TR>
		
		<TR>
		<td align="left" colspan=3>(<font color="red">*</font>) Required fields</td></tr>

	</TABLE>

	<INPUT type="hidden" id=nID name=nID value="<%= nID %>">
	<INPUT type="hidden" id=nMode name=nMode value="<%=nMode%>">
	<input type="hidden" id=from name=from value="<%=strFrom%>">
	<input type="hidden" id=ChangeMask name=ChangeMask value="">
	<input type="hidden" id=StatusID name=StatusID value="">
	<INPUT type="hidden" id=sDesc name=sDesc value="<%=sDesc%>">
	<INPUT type="hidden" id=nUsedInMOL name=nUsedInMOL value="<%=nUsedInMOL%>">
	<INPUT type="hidden" id=nOtherSelectedAliasIDs name=nOtherSelectedAliasIDs value="<%=sOtherSelectedAliasIDs%>">
	<input type="hidden" id=strRegionIdMaskIds name=strRegionIdMaskIds value ="<%=strCloneRegionIds %>">
	
	<INPUT type="hidden" id=newCPLLocaleString name=newCPLLocaleString value="<%=sNewCPLLocaleString %>">
	<INPUT type="hidden" id=newRASLocaleString name=newRASLocaleString value="<%=sNewRASLocaleString %>">
	<INPUT type="hidden" id=newBOMLocaleString name=newBOMLocaleString value="<%=sNewBOMLocaleString %>">
	<INPUT type="hidden" id=newOBSLocaleString name=newOBSLocaleString value="<%=sNewOBSLocaleString %>">
	<INPUT type="hidden" id=newGBLLocaleString name=newGBLLocaleString value="">
	
	<INPUT type="hidden" id=cloneCPLLocaleString name=cloneCPLLocaleString value="<%=sNewCPLLocaleString %>">
	<INPUT type="hidden" id=cloneRASLocaleString name=cloneRASLocaleString value="<%=sNewRASLocaleString %>">
	<INPUT type="hidden" id=cloneBOMLocaleString name=cloneBOMLocaleString value="<%=sNewBOMLocaleString %>">
	<INPUT type="hidden" id=cloneOBSLocaleString name=cloneOBSLocaleString value="<%=sNewOBSLocaleString %>">
	<input type="hidden" id=txtHubCheckboxlist name=txtHubCheckboxlist value="<%=sHubCheckboxlist%>">
	<input type="hidden" id=txtHideHubCheckboxlist name=txtHideHubCheckboxlist value="<%=sHideHubCheckboxlist%>">
	<INPUT type="hidden" id=sPlUserDivisions name=sPlUserDivisions value="<%=sPlUserDivisions%>">
	<INPUT type="hidden" id=txtnProductLineId name=txtnProductLineId value="<%= sProductline_Ori%>">

	<INPUT type="hidden" id=lbxCategorySelection name=lbxCategorySelection value="">
    <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
	<%
end if 'no error

set oRsCreateGroups = nothing
set oRsProductLine = Nothing
set oRsBusSeg = Nothing
set oRsBusSegSelected = Nothing

%>
</FORM>
<% if Request.QueryString("nEditLocalization") = 1 and nMode = 2 then %>
<script> 
	// go to edit localization area
	self.location.hash="#editlocalization"; 
</script> 
<% end if %>
</BODY>
</HTML>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
