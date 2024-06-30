<%@ Language=VBScript %>
<% Option Explicit %>
<!------------------------------------------------------------------- 
'Description: AMO DATA
'----------------------------------------------------------------- //-->    
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataAVL.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../data/oDataNotification.asp" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/MOL_SendEmail.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
'printrequest
'Response.End
Call ValidateSession

dim sDesc, sHeader, sErr, sHelpFile, sShortDesc, sLongDesc, sActualCost, strViewLink, nHistoryMode
dim sBluePN, sRedPN, sRasDisconDate, sCPLBlindDate, sAMOCost, sAMOWWPrice, sNetWeight
dim sExportWeight, sAirPackedWeight, sAirPackedCubic, sOptionType, sOptionCategory, sBusSegIDs, sGroupID
dim sReplacement, sAlternative, sBOMRevADate, sExportCubic, sAliasIDs, strNotes
dim nID, nMode, strRegionalInfos, strRegionId, strMaskValue, arrRegionalInfo, arrRegionMaskIds
dim lngChangeMask, lngModuleID, lngNewStatusID, lngMOLHide, intCount
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim bCostCreate, bCostView, bCostUpdate, bCostDelete, sHideHubCheckboxlists
dim oSvr, oErr, oRsReturn, oRsToUsers, oRsWSToUsers, oRs, oRsOriginal
dim sAMOPartNoRe, sTargetNA, sTargetLA, sTargetEMEA, sTargetAPJ, sBurdenPer, sContraPer, sJustificationnotes 
dim sVisibility_NA, sVisibility_EM, sVisibility_AP, sVisibility_LA, sReplacementDescription, sOrderInstructions
dim sOldDesc, sRuleID, sPath, sViewLink, oNotification, arrLocales, strLocale
dim sManufactureCountry, sWarrantyCode, sObsoleteDate, sRasObsoleteDate
dim lngSCMHide, lngSCLHide, lngIDP, intClone, strRuleDescription, intLocalized, sComBusSelected, sHubCheckboxlists
dim sCloneBOMRevADate, sCloneRasDisconDate, sCloneCPLBlindDate, sCloneObsoleteDate, nProductLineID, strORIHideFromSCMIDS, tmpstrORIHideFromSCMIDS
dim rsMOL,  ndivisionID, rsMOL2, sGlobalSeriesDate, strHideFromSCMIDS_XML, strHideFromSCMIDS, tmpstrHideFromSCMIDS
dim sPlUserDivisions
dim sProductline_Ori

sErr = ""


'get permission
GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMOCost_Permission"), bCostCreate, bCostView, bCostUpdate, bCostDelete

sHelpFile = ""

if Request.Form("StatusID") = cstr(Application("AMO_DISABLED")) then
	sHeader = "Disable After Market Option - Error"
	lngNewStatusID = clng(Application("AMO_DISABLED"))
elseif Request.Form("StatusID") = cstr(Application("AMO_RE-ENABLED")) then
	sHeader = "Enable After Market Option - Error"
	lngNewStatusID = clng(Application("AMO_RE-ENABLED"))
elseif Request.Form("StatusID") = cstr(Application("AMO_OBSOLETE")) then
	sHeader = "Enable After Market Option - Error"
	lngNewStatusID = clng(Application("AMO_OBSOLETE"))
elseif Request.Form("StatusID") = cstr(Application("AMO_COMPLETE")) then
    sHeader = "Enable After Market Option - Error"
	lngNewStatusID = clng(Application("AMO_COMPLETE"))
else
	sHeader = "Save After Market Option - Error"
	lngNewStatusID = 0
end if

nMode = Request.Form("nMode")
nHistoryMode = Request.QueryString("nMode")

if nHistoryMode <> 6 then
	select case nMode
		case "1", "2", "3", "5"
			nMode = clng(nMode)
		case else
			'invalid mode
			sErr = "Invalid Mode passed. AMO_Save.asp"
			Response.Write(sErr)
		    Response.End()
	end select
end if

if sErr = "" then
	if Request.Form("nID") = "" then
		nID = 0	'create mode
	else
		if nMode = 3 then	'clone
			nID = 0 'make it like create mode
			if Request.Form("from") = "CTO" then 'if clone from CTO,
				' already created new module in databse, so update module
				nID = clng(Request.Form("nID"))
			end if 
		else
			nID = clng(Request.Form("nID"))	'update mode
		end if
	end if

	if nMode = 5 then
		'delete after market option
		strViewLink = ""
		'set oSvr = server.CreateObject("JF_S_MODULE.ISMODULE")
        set oSvr = New ISMODULE
		set oErr = oSvr.Module_Remove(Application("REPOSITORY"), cstr(nID), oRsReturn )
		if not oErr is nothing then
			sHeader = "Delete After Market Option - Error"
		else
			if not oRsReturn is nothing then
				if oRsReturn.RecordCount > 0 then
					set oErr = server.CreateObject("JF_H_Error.CErrors")
					oErr.Add 1, "The After Market Option cannot be deleted because it is used in Module and Option Lists. Please run the Where Used report for the After Market Option.", "AMO_Save.asp", ""
					strViewLink = Application("IRSWebServer") & "irs/validate.asp?link=AMO/AMO_Properties.asp?Edit=1&Mode=2&from=" & Request.form("from") & "&ID=" & cstr(nID)
					sHeader = "Delete After Market Option - Error"
				else
					sHeader = "Delete After Market Option - Confirmation"
				end if
				set oRsReturn = nothing
			end if
		end if

	elseif nHistoryMode = 6 then
		
		strHideFromSCMIDS = Request.Form("chkHideFromSCM")
		strHideFromSCMIDS = strHideFromSCMIDS & ", "
		strHideFromSCMIDS_XML = "<ROOT>"       
		While Not strHideFromSCMIDS=""
			tmpstrHideFromSCMIDS = left(strHideFromSCMIDS, instr(1,strHideFromSCMIDS, ",",1)-1)
			strHideFromSCMIDS_XML = strHideFromSCMIDS_XML & "<ChangeHistory ChangeHistoryID=""" & tmpstrHideFromSCMIDS & """/>" & vbCrLf
			strHideFromSCMIDS = right(strHideFromSCMIDS, len(strHideFromSCMIDS)- instr(1,strHideFromSCMIDS, ",",1)-1)
		Wend
    
		strORIHideFromSCMIDS = Request.Form("OriginalSelectedIDs")
		strORIHideFromSCMIDS = strORIHideFromSCMIDS & ","
		  
		While Not strORIHideFromSCMIDS=""
			tmpstrORIHideFromSCMIDS = left(strORIHideFromSCMIDS, instr(1,strORIHideFromSCMIDS, ",",1)-1)
			strHideFromSCMIDS_XML = strHideFromSCMIDS_XML & "<ORIChangeHistory ORIChangeHistoryID=""" & tmpstrORIHideFromSCMIDS & """/>"
			strORIHideFromSCMIDS = right(strORIHideFromSCMIDS, len(strORIHideFromSCMIDS)- instr(1,strORIHideFromSCMIDS, ",",1))
		Wend
		strHideFromSCMIDS_XML = strHideFromSCMIDS_XML & "</ROOT>"
		
        'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		sErr = oSvr.AMO_ChangeHistory_Update(Application("REPOSITORY"),_
				 replace(Request.Form("newAVReassonID"),",,", ","),_
				 Replace(Request.Form("newAVReasson"), "||", "|"), _
				 strHideFromSCMIDS_XML)
				 
		
		if sErr = "True" then		 
			Response.Redirect "AMO_ChangeHistory.asp?nMode=1"
		end if
		
	else	'nMode = 1,2,3
		
		 'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		if lngNewStatusID <> 0 then
			sErr = oSvr.UpdateStatus(Application("REPOSITORY"), clng(nID), lngNewStatusID, session("FullName"), False, Session("AMOUserID") )
			lngModuleID = clng(nID)
		else
			sHeader = "Save After Market Option - Confirmation"

			if IsODM = 0 then
				sAMOCost = Request.Form("txtAMOCost")
				sAMOWWPrice = Request.Form("txtAMOWWPrice")
				sActualCost = Request.Form("txtActualCost")

				sTargetNA = Request.form("txtTargetNA")
				sTargetLA = Request.form("txtTargetLA")
				sTargetEMEA = Request.form("txtTargetEMEA")
				sTargetAPJ = Request.form("txtTargetAPJ")
				sBurdenPer = Request.form("txtBurdenPer")
				sContraPer = Request.form("txtContraPer")
				sJustificationnotes = Request.form("txtJustification")

			else
				'ODM user so retrieve original values
				if nMode = 2 or nMode = 3 then	'edit or clone mode
					set oRsOriginal = oSvr.AMOModule_Search(Application("REPOSITORY"), " and O.FeatureID=" & Request.Form("nID") & " and O.SCMID=1 and R.SCMID = 1 and R.RegionID = " & Application("AMO_GLOBAL_REGIONID") & "", "", null, null)
					if oRsOriginal is nothing then
						sErr = "Missing required parameters.  Unable to complete your request."
		                Response.Write(sErr)
                        Response.End()
					else
						sAMOCost = oRsOriginal.Fields("AMOCost").Value
						sAMOWWPrice = oRsOriginal.Fields("AMOWWPrice").Value
						sActualCost = oRsOriginal.Fields("ActualCost").Value

						sTargetNA = cstr(oRsOriginal.Fields("TargetVolumn_NA").Value)
						sTargetLA = cstr(oRsOriginal.Fields("TargetVolumn_LA").Value)
						sTargetEMEA = cstr(oRsOriginal.Fields("TargetVolumn_EM").Value)
						sTargetAPJ = cstr(oRsOriginal.Fields("TargetVolumn_AP").Value)
						sBurdenPer = cstr(oRsOriginal.Fields("Burden").Value)
						sContraPer = cstr(oRsOriginal.Fields("Contra").Value)
						sJustificationnotes = oRsOriginal.Fields("VolumnMargin_Justification").Value
						
						oRsOriginal.Close
						set oRsOriginal = nothing
					end if
				else 'new 
					sAMOCost = ""
					sAMOWWPrice = ""
					sActualCost = ""

					sTargetNA = ""
					sTargetLA = ""
					sTargetEMEA = ""
					sTargetAPJ = ""
					sBurdenPer = ""
					sContraPer = ""
					sJustificationnotes = ""
				end if
			end if

			if sErr = "" then			
				sDesc = Request.Form("txtDesc")
				sShortDesc = Request.Form("txtShortDesc")
				sLongDesc = Request.Form("txtLongDescription")
				sOptionType = Request.Form("rdType")
				sOptionCategory = Request.Form("lbxCategory")
				if sOptionCategory = "" then
					sOptionCategory = Request.Form("lbxCategorySelection")	'handle Firefox not passing selection for some reason
				end if
				sBusSegIDs = Request.Form("chkDivision")
				sBluePN = Request.Form("txtBluePN")
						
				sHubCheckboxlists = Request.Form("txtHubCheckboxlist")
				sHideHubCheckboxlists = Request.Form("txtHideHubCheckboxlist")
				
				sBOMRevADate = Request.Form("newBOMLocaleString")
				sRasDisconDate = Request.Form("newRASLocaleString")
				sCPLBlindDate = Request.Form("newCPLLocaleString")
				sObsoleteDate = Request.Form("newOBSLocaleString")
				sGlobalSeriesDate = Request.Form("newGBLLocaleString")
				sPlUserDivisions = Request.Form("sPlUserDivisions") 
				if sGlobalSeriesDate = "" then
					sGlobalSeriesDate = Request.Form("txtGlobalseriesDate")
				end if
				
				intClone = 0
					
				sBOMRevADate = replace(sBOMRevADate, ",,", ",")
				if left(sBOMRevADate, 1) = "," then sBOMRevADate = mid(sBOMRevADate, 2)
				if right(sBOMRevADate, 1) = "," then sBOMRevADate = left(sBOMRevADate, len(sBOMRevADate)-1)

				arrLocales = split(sBOMRevADate, ",")
				sBOMRevADate = ""
				for each strLocale in arrLocales
					sBOMRevADate = sBOMRevADate & strLocale & "," & request.form("txtBOMRevADate" & strLocale) & "|"
				next
				
				
				sRasDisconDate = replace(sRasDisconDate, ",,", ",")
				if left(sRasDisconDate, 1) = "," then sRasDisconDate = mid(sRasDisconDate, 2)
				if right(sRasDisconDate, 1) = "," then sRasDisconDate = left(sRasDisconDate, len(sRasDisconDate)-1)

				arrLocales = split(sRasDisconDate, ",")
				sRasDisconDate = ""
				for each strLocale in arrLocales
					sRasDisconDate = sRasDisconDate & strLocale & "," & request.form("txtRASDiscontinueDate" & strLocale) & "|"
				next
				
			
				sCPLBlindDate = replace(sCPLBlindDate, ",,", ",")
				if left(sCPLBlindDate, 1) = "," then sCPLBlindDate = mid(sCPLBlindDate, 2)
				if right(sCPLBlindDate, 1) = "," then sCPLBlindDate = left(sCPLBlindDate, len(sCPLBlindDate)-1)

				arrLocales = split(sCPLBlindDate, ",")
				sCPLBlindDate = ""
				for each strLocale in arrLocales
					sCPLBlindDate = sCPLBlindDate & strLocale & "," & request.form("txtCPLBlindDate" & strLocale) & "|"
				next
							
				sObsoleteDate = replace(sObsoleteDate, ",,", ",")
				if left(sObsoleteDate, 1) = "," then sObsoleteDate = mid(sObsoleteDate, 2)
				if right(sObsoleteDate, 1) = "," then sObsoleteDate = left(sObsoleteDate, len(sObsoleteDate)-1)

				arrLocales = split(sObsoleteDate, ",")
				sObsoleteDate = ""
				for each strLocale in arrLocales
					sObsoleteDate = sObsoleteDate & strLocale & "," & request.form("txtObsoleteDate" & strLocale) & "|"
				next
							
				if nMode = 3 then
				'	sBOMRevADate = Request.Form("cloneBOMLocaleString") & sBOMRevADate 
				'	sRasDisconDate = Request.Form("cloneRASLocaleString") & sRasDisconDate 
				'	sCPLBlindDate = Request.Form("cloneCPLLocaleString") & sCPLBlindDate 
				'	sObsoleteDate = Request.Form("cloneOBSLocaleString") & sObsoleteDate 
					intClone = 1		
				End If
				
				if nMode = 1 then
					intClone = 1		
				End If
				
				sReplacement = Request.Form("txtReplacement")
				sAlternative = Request.Form("txtAlternative")
				sNetWeight = Request.Form("txtNetWeight")
				sExportWeight = Request.Form("txtExportWeight")
				sAirPackedWeight = Request.Form("txtAirPackedWeight")
				sAirPackedCubic = Request.Form("txtAirPackedCubic")
				sExportCubic = Request.Form("txtExportCubic")
				sRuleID = Request.Form("txtRuleID")
				sAliasIDs = Request.Form("lbxSelectedAliasID")
				if len(sPlUserDivisions) > 0 then
					sAliasIDs = sPlUserDivisions & "-" & sAliasIDs
				end if 
				
				sComBusSelected = Request.Form("lbxSelectedDivision")			
				sGroupID = Request.Form("lbxGroupID")
				lngChangeMask = Request.form("ChangeMask")
				strNotes = Request.Form("txtNotes")
				lngMOLHide = Request.Form("chkMOLHide")
				strRegionalInfos = Request.form("strRegionIdMaskIds")
				nProductLineID = Request.form("lbxProductLine")
				sProductline_Ori =  Request.form("txtnProductLineId")
				'steven: fix #5585: saving when lbox ProductLine is disabled, status change
				'Response.Write "product Line new: " & cstr(nProductLineId)
				if (nProductLineID = "0" or nProductLineID ="") and sProductline_Ori<> "0" and sProductline_Ori<> ""then
					'Response.Write "product Line orig: " & cstr(sProductLine_Ori) 
					nProductLineID = sProductline_Ori
				end if 
				'Response.End
			
				sAMOPartNoRe = Request.form("txtAMOPartNoRe")
				sManufactureCountry = Request.form("txtManufactureCountry")
				sWarrantyCode = Request.form("txtWarrantyCode")
				sVisibility_NA = Request.form("oVisibilityNA")
				sVisibility_EM = Request.form("oVisibilityEM")
				sVisibility_AP = Request.form("oVisibilityAP")
				sVisibility_LA = Request.form("oVisibilityLA")
				lngSCMHide = Request.Form("chkSCMHide")
				lngSCLHide = Request.Form("chkSCLHide")
				strRuleDescription = Request.form("txtRuleDescription")
				lngIDP = Request.Form("chkidp")
				sReplacementDescription = Request.form("txtReplacementDescription")
				sOrderInstructions = Request.form("txtOrderInstructions")
				
				if cbool(Request.form("rdLocalized")) = True then
					intLocalized = 1
				else
					intLocalized = 0
				end if
				
				oSvr.intIgnoreDeploy = lngIDP		
				sErr = oSvr.AMO_Update(Application("REPOSITORY"), clng(nID), sDesc, sShortDesc, clng(sOptionType), clng(sOptionCategory), _
					sBusSegIDs, sBluePN, sGlobalSeriesDate, sBOMRevADate, sRasDisconDate, sCPLBlindDate, sAMOCost, sAMOWWPrice, sActualCost, _
					sReplacement, sAlternative, sNetWeight, sExportWeight, sAirPackedWeight, sAirPackedCubic, _
					sExportCubic, sAliasIDs, clng(sGroupID), clng(lngChangeMask), session("FullName"), strNotes, _
					sObsoleteDate, lngMOLHide, lngSCLHide, sAMOPartNoRe, sTargetNA, sTargetLA, sTargetEMEA, sTargetAPJ, sBurdenPer, _
					sContraPer, sJustificationnotes, lngModuleID, cint(sVisibility_NA), cint(sVisibility_EM), cint(sVisibility_AP), cint(sVisibility_LA), _
					sRuleID, sPath, sManufactureCountry, sWarrantyCode, lngSCMHide, sLongDesc, _
					cstr(sReplacementDescription), sOrderInstructions, strRuleDescription, intClone, intLocalized, sComBusSelected, sHubCheckboxlists, _
					sHideHubCheckboxlists, cint(nProductLineID), cint(Request.QueryString("nEditLocalization"))) 				
			end if		  
		end if
	end if
	
	if sErr <> "True" then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
	else
		'if the module is used in a POR'd MOL And the description has changed, send an email to SEPM of all the POR'd MOLs using this module
		if nMode = 2 then
			sOldDesc = Request.Form("sDesc")
			if (Request.Form("nUsedInMOL") = 2 and sOldDesc <> sDesc and sDesc <> "" and sOldDesc <> "") then
				'send email
				
				sErr = SendEmailToSEPM(nID, sOldDesc, sDesc)
				if sErr <> "True" then
					sErr = "Missing required parameters.  Unable to complete your request."
		            Response.Write(sErr)
		            Response.End()
				end if
			end if
			
		    'set oSvr = Server.CreateObject("JF_S_Module.ISModule")
    		set oSvr = New ISMODULE 	
    		set rsMOL = oSvr.Module_MOLWhereUsed(application("Repository"), nID)
    		if rsMOL is nothing then
			    sErr = "Missing required parameters.  Unable to complete your request."
		        Response.Write(sErr)
		        Response.End()
			else		
    			if rsMOL.recordcount>0 then      
    				set rsMOL2 = CopyRs(rsMOL, true)
    				rsMOL.sort="DivisionID"  
    				nDivisionID=0
    				
    				rsMOL.movefirst
					while (not rsMOL.EOF) 
						'if rsMOL("StatusID") >= Application("MOL_EG_COMMITMENT") then
								
								if nDivisionID<>rsMOL("DivisionID") then
								
								
									SendMOL_ModuleChangeEmail_update "updated",rsMOL("ListID"), cstr(nID ), rsMOL("DivisionID"), rsMOL2, sDesc
								'based on experience elsewhere, put time delays here
									Delay_5Seconds	
									
								end if 				
						'end if 
						nDivisionID=rsMOL("DivisionID")
						rsMOL.MoveNext				
					wend
					set rsMOL2=nothing
				end if	
			end if
			
			
			
			
			
		end if

	end if
	set oSvr = nothing
end if

if sErr = "" and lngModuleID <> 0 then
'	if (lngModuleID <> 0) and nMode <> 5 then
'		'return to properties page
'		Response.Redirect "/AMO/isapi/AMO_Properties.asp?Edit=1&Mode=2&from=" & Request.form("from") & "&ID=" & cstr(lngModuleID)
'	end if

	if (nMode = 1 or nMode = 3) then
		'construct the link to refresh the tree to the new item
		sViewLink = Application("IRSWebServer") & "irs/validate.asp?link=MOL/ModuleTreeView/isapi/Module_Tree.asp?Path=" & sPath
	else
		'construct the link for view/modify the current module
		sViewLink = Application("IRSWebServer") & "irs/validate.asp?link=AMO/AMO_Properties.asp?Edit=1&Mode=2&from=" & Request.form("from") & "&ID=" & cstr(lngModuleID)

	end if

	if nMode <> 5 then 'not delete mode and no error -> show the module properties page
      		Response.Redirect sViewLink
	end if

end if


'save the Region checkmark
function SaveRegionStatus(sRepository, ModuleID, RegionID, SetStatus, strPerson)
	dim oSvr, oErr, sErr

	SaveRegionStatus = ""
	On Error Resume Next
	'set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
    set oSvr = New ISAMO
	if Err.Number <> 0 then
		SaveRegionStatus = Err.Description 
	else
		sErr = oSvr.SaveRegionStatus(cstr(sRepository), ModuleID, RegionID, SetStatus, strPerson)
		if Err.number <> 0 then 
			SaveRegionStatus = Err.Description 
		end if 
		if sErr <> "True" then
		  SaveRegionStatus = sErr
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function


function SendEmailToSEPM (Byval nModuleID, byval sOldName, byval sName)

	dim sToSEPM, sBody, rsSEPM, sModuleLink

	'Get a list of all SEPM emails
	'set oSvr = server.CreateObject("JF_S_Module.ISModule")
    set oSvr = New ISMODULE
	set rsSEPM = oSvr.Module_GetPORSEPM(Application("REPOSITORY"), nModuleID)	
	if rsSEPM is nothing then 
		sToSEPM = ""
		if not rsSEPM.EOF and not rsSEPM.BOF then
			rsSEPM.MoveFirst
			do while (not rsSEPM.EOF)
				if sToSEPM <> "" then sToSEPM = sToSEPM & ";"
				sToSEPM = sToSEPM & rsSEPM.Fields("Email").Value
				rsSEPM.MoveNext
			loop
		end if
	end if
	
	'Set oNotification = Server.CreateObject("JF_S_Notification.ISNotification")
	set oSvr = New ISNotifciation	
	
    'set oErr = GetMOLCategory(oRs, 9)
    set oRs = GetMOLCategory(9)	
    if oRs is Nothing then
	    Response.Write("Recordset error: oRs")
	    Response.End()
    end if
	
	oRs.Filter = "Description = 'Commercial Desktops'"
	if instr(1, sBusSegIDs, oRs("CategoryID").Value) > 0 then
		set oRsToUsers = oNotification.ViewUsersByEventRs(Application("Repository"), Application("DT_MARKETING_NAME_CHANGE"), -1, -1 )
		do while not oRsToUsers.EOF
			if sToSEPM <> "" then sToSEPM = sToSEPM & ";"
			sToSEPM = sToSEPM & trim(oRsToUsers("Email"))				
			oRsToUsers.MoveNext
		loop	
	end if
	
	oRs.Filter = "Description = 'Workstations'"
	if instr(1, sBusSegIDs, oRs("CategoryID").Value) > 0 then
		set oRsWSToUsers = oNotification.ViewUsersByEventRs(Application("Repository"), Application("WS_MARKETING_NAME_CHANGE"), -1, -1 )
		do while not oRsWSToUsers.EOF
			if sToSEPM <> "" then sToSEPM = sToSEPM & ";"
			sToSEPM = sToSEPM & trim(oRsWSToUsers("Email"))				
			oRsWSToUsers.MoveNext
		loop	
	end if
		
	if sToSEPM <> "" then
		sModuleLink = Server.URLEncode(Application("IRSWebServer") & "irs/validate.asp?link=AMO/AMO_Properties.asp?Mode=2&ID=" & nID)
		sBody = " The following Module Marketing Description has changed from <br>" & Server.HTMLEncode(sOldName) _
				& "<br> to <br>" & Server.HTMLEncode(sName) & "<br> <br>" _
				& "Please go to the following link for more details about the update.<br>" _
		    	& "<A href='" & Session("PrefixFullPath") & "default.asp?link=" & sModuleLink & "' style=color:blue target=_blank>View/Modify After Market Option</A> "
		set oErr = SendIRSEmail_To("", sToSEPM, "", "", "", "", "Module Marketing Description changed", "", sBody, false, "" )
	end if
	
	set oRsToUsers = nothing
	set oRsWSToUsers = nothing
	set oRs	= nothing
	set oNotification = nothing	

    If Err.Number <> 0 Then
	    SendEmailToSEPM = True
    Else
        sendEmailToSEPM = False
    End If

end function


%>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">

<title><%=sHeader%></title>
</HEAD>
<BODY LANGUAGE=javascript>
<FORM name=thisform method=post>
<%
'Response.write BuildHelp(sHeader, sHelpfile)
if sErr <> "" then 
	Response.Write sErr 
else 
	Response.write "<p>Your change has been recorded in the IRS system. " & vbCrLf
end if
%>
<table border="0" cellPadding="5" cellSpacing="5" width="100%">
	<% if strViewLink <> "" then %>
	<tr>
		<td><A href="<%= strViewLink %>" style="COLOR: blue">View/Modify Current After Market Option</a></td>
	</tr>
	<% end if %>
	<% if bAMOCreate then %>
	<tr>
		<td><A href="Module/isapi/GetGroupsForModuleRole.asp" style="COLOR: blue">Create a New Module</a></td>
	</tr>
	<% end if %>
</table>
<%

%>
</FORM>
</BODY>
</HTML>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->