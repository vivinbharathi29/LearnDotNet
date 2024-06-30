<%@ Language=VBScript %>
<% OPTION EXPLICIT 
	
    Response.Buffer = False
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"

    Server.ScriptTimeout = 6000
%>
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
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->

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
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
Call ValidateSession

dim sHeader, sHelpFile, sErr, sStatusIDs, strEOLDate
dim sModuleCategoryHTML, sStatusHTML, sBusSegIDs, sBusSegHTML
dim nCategoryID, nNumTotalModules, sGpsyCom, sRasCom
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete, bRAS, nMode
dim bRASCreate, bRASView, bRASUpdate, bRASDelete
dim bCostCreate, bCostView, bCostUpdate, bCostDelete, bGpsyCom, bRasCom
dim oSvr, oErr, strError, rs, dw, cn, cmd
dim oRsAMOModules, oRsCategories, sHideColumns
dim sDivisionIDs, intCount, oRsCreateGroups, sGroupIDs, sOwnerIDs, sOwnersHTML
dim bDisplayHideromMOL, bDisplayHideromSCM, sDisplayHideromMOL, sDisplayHideromSCM, sDisplayHideromSCL, bDisplayHideromSCL, bDisplayIncludeBlankGSD, sDisplayIncludeBlankGSD
dim strMRFromDate, strMRToDate, strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, strRasObsoleteToDate, strRasObsoleteFromDate
dim strchkMRBlank, strchkCPLBlank, strchkDisBlank, oRsProductLine, nCategoriesSelected, strKeyWord, ckKeyWord

sHeader = "After Market Option List"
sErr = ""

'set rsRoles and IRSUserID Session: ----
Call SetPermission()

'get permissions
GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
GetRights2 Application("AMORAS_Permission"), bRASCreate, bRASView, bRASUpdate, bRASDelete
GetRights2 Application("AMOCost_Permission"), bCostCreate, bCostView, bCostUpdate, bCostDelete

sDivisionIDs = ""
sGroupIDs = ""

set oRsProductLine = Nothing

'get user permission to Business Segement and Group: ---
'Validate User Role has Permission to either CREATE AMO List, AMO RAS, or AMO Cost; if so, creates a string of Business Segment and Group IDs the User has additional access to: ----
if bAMOUpdate then
	set oRsCreateGroups = GetGroupsForRole2(cstr(Session("AMOUserRoleNames")), cstr(Application("AMO_Permission")), true, false, false, false, true, false)
	if (oRsCreateGroups is nothing) then
		Response.Write("Empty Recordset: oRsCreateGroups")
		Response.End()
	else
		For intCount = 0 To oRsCreateGroups.RecordCount-1
			sDivisionIDs = "," & sDivisionIDs & replace(oRsCreateGroups("DivisionIDs"), "|", ",")
			sGroupIDs = "," & sGroupIDs & oRsCreateGroups("GroupID") & ","
			oRsCreateGroups.MoveNext		
		Next
	end if  
	
elseif bRASUpdate then
	set oRsCreateGroups = GetGroupsForRole2(cstr(Session("AMOUserRoleNames")), cstr(Application("AMORAS_Permission")), true, false, false, false, true, false)
	if (oRsCreateGroups is nothing) then
		Response.Write("Empty Recordset: oRsCreateGroups")
		Response.End()
	else		
		For intCount = 0 To oRsCreateGroups.RecordCount-1
			sDivisionIDs = "," & sDivisionIDs & replace(oRsCreateGroups("DivisionIDs"), "|", ",")
			sGroupIDs = "," & sGroupIDs & oRsCreateGroups("GroupID") & ","
			oRsCreateGroups.MoveNext		
		Next
	end if	
else
	set oRsCreateGroups = GetGroupsForRole2(cstr(Session("AMOUserRoleNames")), cstr(Application("AMOCost_Permission")), true, false, false, false, true, false)
	if (oRsCreateGroups is nothing) then
		Response.Write("Empty Recordset: oRsCreateGroups")
		Response.End()
	else
		For intCount = 0 To oRsCreateGroups.RecordCount-1
			sDivisionIDs = "," & sDivisionIDs & replace(oRsCreateGroups("DivisionIDs"), "|", ",")
			sGroupIDs = "," & sGroupIDs & oRsCreateGroups("GroupID") & ","
			oRsCreateGroups.MoveNext		
		Next				
	end if	
end if


bRAS = False
if bRASUpdate then
	bRAS = True
end if
'if bAMOUpdate and bRASUpdate then
	'default to non-RAS mode if both
'	bRAS = False
'end if

if bRAS then
	sHelpFile = "../help/HELP_AMO_ModuleList_RAS.asp"
else
	sHelpFile = "../help/HELP_AMO_ModuleList_Options.asp"
end if

'dim strModuleIDs
'strModuleIDs = Request.Form("txtModuleIDs")
'nmode=Request.QueryString("nMode")

'if nMode =1 then
'Response.Write("call email reject, module:" + strModuleIDs)
'
'	sendEmailReject(strModuleIDs)
'
'end if 
'Set status to RAS Review or Reject
if request.form("ID") <> "" and request.form("StatusID") <> "" then
	'set the status first
	if clng(request.form("StatusID")) = clng(Application("AMO_RASREVIEW")) and (bAMOUpdate or bCostUpdate) then
		'Set to RAS Review		
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		strError = oSvr.UpdateStatus(Application("REPOSITORY"), clng(request.form("ID")), clng(Application("AMO_RASREVIEW")), session("FullName"), False, Session("AMOUserID"))
		if strError <> "True" then
			Response.Write(strError)
            Response.End()
		end if		
	elseif clng(request.form("StatusID")) = clng(Application("AMO_RASUPDATE")) and (bAMOUpdate or bCostUpdate) then
		'Set to RAS Update
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
		set oSvr = New ISAMO
		if clng(request.form("PreviousStatusID")) = clng(Application("AMO_DISABLED")) then
			strError = oSvr.UpdateStatus(Application("REPOSITORY"), clng(request.form("ID")), clng(Application("AMO_RASUPDATE")), session("FullName"), True, Session("AMOUserID"))
		else
			strError = oSvr.UpdateStatus(Application("REPOSITORY"), clng(request.form("ID")), clng(Application("AMO_RASUPDATE")), session("FullName"), False, Session("AMOUserID"))
		end if		
		if strError <> "True" then
			Response.Write(strError)
            Response.End()
		end if
	elseif clng(request.form("StatusID")) = clng(Application("AMO_COMPLETE")) and bRASUpdate then
		'Set to Complete
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		strError = oSvr.UpdateStatus(Application("REPOSITORY"), clng(request.form("ID")), clng(Application("AMO_COMPLETE")), session("FullName"), False, Session("AMOUserID"))
		if strError <> "True" then
			Response.Write(strError)
            Response.End()
		end if
	elseif clng(request.form("StatusID")) = clng(Application("AMO_REJECT")) and bRASUpdate then
		'Set to Reject
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		strError = oSvr.UpdateStatus(Application("REPOSITORY"), clng(request.form("ID")), clng(Application("AMO_REJECT")), session("FullName"), False, Session("AMOUserID") )
		if strError <> "True" then
			Response.Write(strError)
            Response.End()
		end if
	end if
end if

if sErr = "" then

    '6/29/16 - VHarris - PBI 17487/Task 21005 - Create AMOFeatureCategory cookie for pages using Feature Category
    If GetDBCookie("AMO lbxAMOFeatureCategoryStored") = "" Then
        'If User never used Feature Category Stored, create cookie
        Call SaveDBCookie( "AMO lbxAMOFeatureCategoryStored", "1")
    End If

	if Request.Form("lbxSelectedCategory") = "" and GetDBCookie("AMO lbxAMOFeatureCategoryStored") = "1"  and Request.Form("emptycategory") <> "1" then
		'get the cookie. If we didn't get it default it
		set oRsCategories = GetDBCookieSet( "AMO lbxAMOFeatureCategoryIDs")
        
        'store the cookie
	    Call SaveDBCookieSet( "AMO lbxAMOFeatureCategoryIDs", "", oRsCategories)
	else
        'get the selected categories: ---
		nCategoriesSelected = Request.Form("lbxSelectedCategory")
        
        'store the selected category in cookie table: ---
	    Call SaveDBCookieSet( "AMO lbxAMOFeatureCategoryIDs", nCategoriesSelected, "")

       'get the cookie. If we didn't get it default it
		set oRsCategories = GetDBCookieSet( "AMO lbxAMOFeatureCategoryIDs")
	end if	
	Call SaveDBCookie( "AMO lbxAMOFeatureCategoryStored", "1")
	
	if not oRsCategories is nothing then
		if oRsCategories.RecordCount > 0 then
			nCategoriesSelected = replace(RecordsetToDelimitedString(oRsCategories, ",", "Value"), " ", "")
		else
			nCategoriesSelected = ""
		end if
	end if
	
	if Request.Form("chkStatus") = "" then
		'get the cookie. If we didn't get it default it
		sStatusIDs = GetDBCookie( "AMO chkStatus")
		if trim(sStatusIDs) = "" then
			if bAMOUpdate or bCostUpdate then
				sStatusIDs = cstr(Application("AMO_NEW")) & "," & cstr(Application("AMO_COMPLETE")) & "," & cstr(Application("AMO_RASREVIEW")) & "," & cstr(Application("AMO_RASUPDATE"))
			else
				sStatusIDs = cstr(Application("AMO_RASREVIEW")) & "," & cstr(Application("AMO_RASUPDATE"))
			end if
		end if
	else
		sStatusIDs = Request.Form("chkStatus")
	end if
	
	'store the cookie
	Call SaveDBCookie( "AMO chkStatus", sStatusIDs )
	
	if Request.Form("chkBusSeg") = "" then
		'get the cookie. If we didn't get it default it
		sBusSegIDs = GetDBCookie( "AMO chkFeatureBusSeg")
	else
		sBusSegIDs = Request.Form("chkBusSeg")
	end if
	'store the cookie
	Call SaveDBCookie( "AMO chkFeatureBusSeg", sBusSegIDs )
	
	'store the cookie
	Call SaveDBCookie( "AMO chkStatus", sStatusIDs )
	
	if Request.Form("chkGroupOwner") = "" then
		'get the cookie. If we didn't get it default it
		sOwnerIDs = GetDBCookie( "AMO chkGroupOwner")
	else
		sOwnerIDs = Request.Form("chkGroupOwner")
	end if
	'store the cookie
	Call SaveDBCookie( "AMO chkGroupOwner", sOwnerIDs )
	
	
	nMode = Request.QueryString("nMode")
	
	
	if nMode = 10 Then 
		sHideColumns = Request.Form("nColumnIDs")
		'store the cookie
		Call SaveDBCookie("AMO Hide Column", sHideColumns)
	else
		sHideColumns = GetDBCookie("AMO Hide Column")
	end if
    'So code accurately evaluates each column number, add comma to beginning and end of list
    If Len(sHideColumns) > 0 Then
        sHideColumns = "," & sHideColumns & ","
    End If
	
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strMRFromDate = GetDBCookie( "AMO txtMRFromDate")
	else
		strMRFromDate = Request.Form("txtMRFromDate")
		Call SaveDBCookie( "AMO txtMRFromDate", strMRFromDate)
	end if
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strMRToDate = GetDBCookie( "AMO txtMRToDate")
	else  
		strMRToDate = Request.Form("txtMRToDate")
		'store the cookie
	Call SaveDBCookie( "AMO txtMRToDate",strMRToDate)
	end if
	
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strCPLFromDate = GetDBCookie( "AMO txtCPLFromDate")
	else
		strCPLFromDate = Request.Form("txtCPLFromDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtCPLFromDate",strCPLFromDate)
	end if
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strCPLToDate = GetDBCookie( "AMO txtCPLToDate")
	else
		strCPLToDate = Request.Form("txtCPLToDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtCPLToDate",strCPLToDate)
	end if
		
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strDisFromDate = GetDBCookie( "AMO txtDisFromDate")
	else
		strDisFromDate = Request.Form("txtDisFromDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtDisFromDate",strDisFromDate)
	end if
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strDisToDate = GetDBCookie( "AMO txtDisToDate")
	else
		strDisToDate = Request.Form("txtDisToDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtDisToDate",strDisToDate)
	end if
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strRasObsoleteToDate = GetDBCookie( "AMO txtRasObsoleteToDate")
	else
		strRasObsoleteToDate = Request.Form("txtRasObsoleteToDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtRasObsoleteToDate",strRasObsoleteToDate)
	end if
	
	
	if nMode <> 1 Then 
		'get the cookie. If we didn't get it default it
		strRasObsoleteFromDate = GetDBCookie( "AMO txtRasObsoleteFromDate")
	else
		strRasObsoleteFromDate = Request.Form("txtRasObsoleteFromDate")
		'store the cookie
		Call SaveDBCookie( "AMO txtRasObsoleteFromDate",strRasObsoleteFromDate)
	end if
	
	
	if nMode <> 1 Then
		'get the cookie. If we didn't get it default it
		strchkCPLBlank = GetDBCookie( "AMO chkCPLBlankDate")
	else
		strchkCPLBlank = Request.Form("chkCPLBlankDate")
		Call SaveDBCookie( "AMO chkCPLBlankDate", Request.Form("chkCPLBlankDate"))	
	end if
	

	if strchkCPLBlank <> "on" then
		strchkCPLBlank = ""
	else
		strchkCPLBlank = "checked"
	end if
	
	
	if nMode <> 1 Then
		'get the cookie. If we didn't get it default it
		strchkMRBlank = GetDBCookie( "AMO chkMRBlankDate")
	else
		strchkMRBlank = Request.Form("chkMRBlankDate")
		Call SaveDBCookie( "AMO chkMRBlankDate", Request.Form("chkMRBlankDate"))	
	end if
	

	if strchkMRBlank <> "on" then
		strchkMRBlank = ""
	else
		strchkMRBlank = "checked"
	end if
	
	
	if nMode <> 1 Then
		'get the cookie. If we didn't get it default it
		strchkDisBlank = GetDBCookie( "AMO chkDisBlankDate")
	else
		strchkDisBlank = Request.Form("chkDisBlankDate")
		Call SaveDBCookie( "AMO chkDisBlankDate", Request.Form("chkDisBlankDate"))	
	end if
	

	if strchkDisBlank <> "on" then
		strchkDisBlank = ""
	else
		strchkDisBlank = "checked"
	end if
	
	
	'save cookies for Ras and GPSy incompleted filter
	
	if nMode <> 1 Then
		bGpsyCom = GetDBCookie( "AMO chkGpsyCom")
	else
		bGpsyCom = Request.Form("chkGpsyCom")
		Call SaveDBCookie( "AMO chkGpsyCom", Request.Form("chkGpsyCom"))
	
	end if
	
		
	
	if nMode <> 1 Then
		bRasCom = GetDBCookie( "AMO chkRasCom")
	else
		bRasCom = Request.Form("chkRasCom")
		Call SaveDBCookie( "AMO chkRasCom", Request.Form("chkRasCom"))
	
	end if
	
			
	if bRasCom = "on" then
		bRasCom = true
		sRasCom = "checked"
	else
		bRasCom = false
		sRasCom = ""
	end if
	
	if bGpsyCom = "on" then
		bGpsyCom = true
		sGpsyCom = "checked"
	else
		bGpsyCom = false
		sGpsyCom = ""
	end if
	
	
	if nMode <> 1 Then
		bDisplayHideromMOL = GetDBCookie( "AMO chkIncludeHidefromMOL")
		if bDisplayHideromMOL ="" then
			'default this check box to be checked
			bDisplayHideromMOL = "on"
		end if 
		
	else
		bDisplayHideromMOL = Request.Form("chkIncludeHidefromMOL")
		if bDisplayHideromMOL = "on" then
			Call SaveDBCookie( "AMO chkIncludeHidefromMOL", Request.Form("chkIncludeHidefromMOL"))
		else
			Call SaveDBCookie( "AMO chkIncludeHidefromMOL", "off")
		end if 
	end if
	
	if bDisplayHideromMOL = "on" then
		bDisplayHideromMOL = true
		sDisplayHideromMOL = "checked"
	else
		bDisplayHideromMOL = false
		sDisplayHideromMOL = ""
	end if
	
	if nMode <> 1 Then
		bDisplayHideromSCM = GetDBCookie( "AMO chkIncludeHidefromSCM")
	else
		bDisplayHideromSCM = Request.Form("chkIncludeHidefromSCM")
		Call SaveDBCookie( "AMO chkIncludeHidefromSCM", Request.Form("chkIncludeHidefromSCM"))
	end if

	if bDisplayHideromSCM = "on" then
		bDisplayHideromSCM = true
		sDisplayHideromSCM = "checked"
	else
		bDisplayHideromSCM = false
		sDisplayHideromSCM = ""
	end if
	
	'if nMode <> 1 Then
		'bDisplayHideromSCL = GetDBCookie( "AMO chkIncludeHidefromSCL")
	'else
		'bDisplayHideromSCL = Request.Form("chkIncludeHidefromSCL")
		'Call SaveDBCookie( "AMO chkIncludeHidefromSCL", Request.Form("chkIncludeHidefromSCL"))
	'end if
	
	'if bDisplayHideromSCL = "on" then
		'bDisplayHideromSCL = true
		'sDisplayHideromSCL = "checked"
	'else
		bDisplayHideromSCL = false
		sDisplayHideromSCL = ""
	'end if
	
	if nMode <> 1 Then
		bDisplayIncludeBlankGSD = GetDBCookie( "AMO chkIncludeBlankGSD")
		if bDisplayIncludeBlankGSD ="" then
			'default this check box to be checked
			bDisplayIncludeBlankGSD = "off"
		end if 		
	else
		bDisplayIncludeBlankGSD = Request.Form("chkIncludeBlankGSD")
		if bDisplayIncludeBlankGSD = "on" then
			Call SaveDBCookie( "AMO chkIncludeBlankGSD", Request.Form("chkIncludeBlankGSD"))
		else
			Call SaveDBCookie( "AMO chkIncludeBlankGSD", "off")
		end if 
	end if
	
	
	if bDisplayIncludeBlankGSD = "on" then
		bDisplayIncludeBlankGSD = true
		sDisplayIncludeBlankGSD = "checked"
	else
		bDisplayIncludeBlankGSD = false
		sDisplayIncludeBlankGSD = ""
	end if
	'walter add keyword search	
	if nMode = 1 then
		Call SaveDBCookie("AMO Keyword", "")
		Call SaveDBCookie("AMO SelCat", "")
		'Call SaveDBCookie("AMO ckKeyword", "")
	end if
	
	if nMode = 12 then
		Call SaveDBCookie("AMO ckKeyword", ""&Request.Form("txtKeyWord")&"")
		strKeyWord = Session("AMOUserID")& "|" & "'%"&Request.Form("txtKeyWord")&"%'"
		Call SaveDBCookie("AMO SelCat", "")
	else
		strKeyWord = GetDBCookie("AMO Keyword")
	end if
	
	Call SaveDBCookie("AMO Keyword", strKeyWord)
	
	ckKeyword = GetDBCookie("AMO ckKeyword")	
	' walter end
	sModuleCategoryHTML = ""
	sStatusHTML = ""
	sBusSegHTML = ""
	sOwnersHTML = ""
	

	set oRsProductLine = Nothing
	'set oErr = GetMOLCategory(oRsProductLine, 29)
	set oRsProductLine = GetMOLCategory(29)	

	if oRsProductLine is Nothing then
		Response.Write("Recordset error: oRsProductLine")
		Response.End()
    end if
	
	sErr = GetFilterInfo(nCategoriesSelected, sStatusIDs, sBusSegIDs, strEOLDate,_
			 sModuleCategoryHTML, sStatusHTML, oRsAMOModules, _
			  nNumTotalModules, sBusSegHTML, bGpsyCom, bRasCom, bRAS ,_
			  bDisplayHideromMOL, bDisplayHideromSCM, bDisplayHideromSCL, strMRFromDate, strMRToDate, _
			  strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, _
			  strchkMRBlank, strchkMRBlank, strchkMRBlank, strRasObsoleteFromDate, strRasObsoleteToDate, sOwnersHTML, sOwnerIDs, bDisplayIncludeBlankGSD, strKeyWord, "NewBus")
		'4/18/2005 since strchkMRBlank value is used for all these three checkboxes, 
		' but teh 'complete' tab has it's own situation, so pass strchkMRBlank here to all three
   
	if sErr = "" then
		if oRsAMOModules.RecordCount > 0 then
			'user has to be in owner group in order to do AMO updates
			'removed check because Marketing Operations has to have rights to update too so only business segment will be checked
'			if not IsUserInGroup(oRsAMOModules("GroupID").value) then
'				bAMOCreate = False
'				bAMOUpdate = False
'				bAMODelete = False
'				bAMOView = False
'				bCostCreate = False
'				bCostView = False
'				bCostUpdate = False
'				bCostDelete = False
'			end if
			'user has to be in one of the target business segments in order to do RAS updates
			'if not UserInAMODivision(oRsAMOModules("DivisionIDs").value) then
			if not UserInAMODivision(sBusSegIDs) then
				bAMOCreate = False
				bAMOUpdate = False
				bAMODelete = False
				bAMOView = False
				bCostCreate = False
				bCostView = False
				bCostUpdate = False
				bCostDelete = False
				bRASCreate = False
				bRASUpdate = False
				bRASDelete = False
				bRASView = False
			end if
		end if
	end if
end if

function isIdBelong(byval dIds, byval sIds)
	dim i, arrIds, bFlag
	
	if Trim(dIds) <> "" then
		arrIds = Split(Trim(dIds), ",")
		bFlag = False
		for i = 0 to UBound(arrIds)
			if Trim(arrIds(i)) <> "" and instr(Trim(sIds), "," & Trim(arrIds(i)) & ",") > 0 then
				bFlag = True
				Exit For
			end if
		Next
	end if
	isIdBelong = bFlag	 
end function
'function SendEmailReject (byval strModuleIDs)
'
'	dim sTo, sBody, rsSEPM
'
'	'Get a list of all emails
'	set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
'	set oErr = oSvr.AMO_SendEmailReject(Application("REPOSITORY"), strModuleIDs, rsSEPM)
'Response.Write("after run the fcn in COM obj with module: "+strModuleIDs)
''Response.End
'		
'	if oErr is nothing then 
'		sTo = ""
'		if not rsSEPM.EOF and not rsSEPM.BOF then
'			rsSEPM.MoveFirst
'			do while (not rsSEPM.EOF)
'				if sTo <> "" then sTo = sTo & ";"
'				sTo = sTo & rsSEPM.Fields("Email").Value
'				rsSEPM.MoveNext
'			loop
'Response.Write("send IRS email list: "+sTo)
''Response.End
'		end if
'	end if
'	if sTo <> "" then
'		'sModuleLink = Server.URLEncode("/AMO/isapi/AMO_Properties.asp?Mode=2&ID=" & nID)
'		sBody = " The comment from RAS"
'		set oErr = SendIRSEmail_To("", sTo, "", "", "", "", "test SUBJ", "", sBody, 0, "" )
'Response.Write("send IRS email list: "+sTo)
'Response.End
'
'	end if
'end function
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
    <title><%=sHeader%> - All Options</title>
    <meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
    <link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
    <link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
    <link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
    <link rel="stylesheet" type="text/css" href="../style/amo.css" />
    <!--[if lte IE 8]>
       <link rel="stylesheet" type="text/css" href="../style/amo_ie8.css" />
    <![endif]-->
    <script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
    <script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
    <script type="text/javascript" src="../library/scripts/formChek.js"></script>
    <script type="text/javascript" src="../library/scripts/calendar.js"></script>
    <script type="text/javascript" src="../library/scripts/general.js"></script>
    <script type="text/javascript" src="../../Scripts/shared_functions.js"></script>
    <script type="text/javascript" src="../library/scripts/popitmenu.js"></script> 
    <script type="text/javascript" src="../scripts/amo.js"></script> 
    <SCRIPT LANGUAGE="JAVASCRIPT">
<!--
    function ClickEvent(evt) {
        var id;
        var ModuleID, CategoryID, PlatformID, RegionID;
        var Mtype;
        var repository = "<%=Application("REPOSITORY")%>"
        var fullname = "<%= session("FullName") %>";
        var sIRSLink = "<%= Application("IRSWebServer") %>";

        if (!evt) evt = window.event;
        objUnknown = evt.srcElement? evt.srcElement : evt.target;
        id = objUnknown.id;

        if ((objUnknown.tagName.toUpperCase() != "TD") && (objUnknown.tagName.toUpperCase() != "IMG"))
            return;

        if (objUnknown.tagName.toUpperCase() == "IMG") {
            // make object the parent TD
            objUnknown = objUnknown.parentNode; //parentElement
        }

        var objName = objUnknown.getAttribute("name"); //cross-browser

        // handle exception since we need to have moduleid on id
        if (objName == 'wwp' || objName == 'cpl' || objName == 'obd' || 
            objName == 'as' || objName == 'vem' || objName == 'vap' || 
            objName == 'vla' || objName == 'vna' || objName == 'reva' || 
            objName == 'rasdisc' || objName == 'gbsd')
            id = objName;

        //Extract ModuleID and CategoryID from 'm123c123'
        ModuleID = objUnknown.parentNode.getAttribute("mid");
        RegionID = objUnknown.parentNode.getAttribute("sid");
        CategoryID = objUnknown.parentNode.getAttribute("cid");
        Mtype = objUnknown.parentNode.getAttribute("mtype");

        switch (id){
            case 'prop': //Properties column
                ShowModuleProperties(ModuleID)
                break;
            case 'o': //Short Description column
                editText(objUnknown, ModuleID, 'shortdescription', 40, repository, fullname);
                break;
            case 'dsc': //long Description column
                enterComment(ModuleID, 'dsc')
                break;
            case 'rdsc': //rules Description column
                enterComment(ModuleID, 'rdsc')
                break;
            case 'reldes': //replacement av Description column
                enterComment(ModuleID, 'reldes')
                break;
            case 'ord': //order instruction column
                enterComment(ModuleID, 'ord')
                break;
            case 'asnew': //AMO status column for New modules
                cM_Status(evt, ModuleID, 'new')
                break;
            case 'asras': //AMO status column for RAS Review modules
                cM_Status(evt, ModuleID, 'ras')
                break;
            case 'as': //AMO status column
                cM_Status(evt, ModuleID, '')
                break;
            case 'asdis': //AMO status column
                cM_Status(evt, ModuleID, 'asdis')
                break;
            case 'ctr': //Comment to RAS column
                enterComment(ModuleID, 'ctr')
                break;
            case 'cfr': //Comment from RAS column
                enterComment(ModuleID, 'cfr')
                break;
            case 'bpn': //Blue PN column
                editText(objUnknown, ModuleID, 'bluepartno', 20, repository, fullname);
                break;
            case 'rpn': //Red PN column
                editText(objUnknown, ModuleID, 'redpartno', 20, repository, fullname);
                break;
            case 'molhide': //Hide from MOL
                if (objUnknown.innerHTML == "&nbsp;") {	
                    // add checkmark
                    cM_ChangeMOLHide(evt, ModuleID, 1, objUnknown, repository, fullname);
                } else {
                    // remove checkmark
                    if (Mtype=="Hardware")
                    {cM_ChangeMOLHide(evt, ModuleID, 0, objUnknown, repository, fullname);}
                }
                break;
            case 'scmhide': //Hide from SCM
                if (objUnknown.innerHTML == "&nbsp;") {	
                    // add checkmark
                    cM_ChangeSCMHide(evt, ModuleID, 1, objUnknown, repository, fullname);
                } else {
                    // remove checkmark
                    cM_ChangeSCMHide(evt, ModuleID, 0, objUnknown, repository, fullname);
                }
                break;
            case 'sclhide': //Hide from SCL
                if (objUnknown.innerHTML == "&nbsp;") {	
                    // add checkmark
                    cM_ChangeSCLHide(evt, ModuleID, 1, objUnknown, repository, fullname);
                } else {
                    // remove checkmark
                    cM_ChangeSCLHide(evt, ModuleID, 0, objUnknown, repository, fullname);
                }
                break;
            case 'ras': //RAS Complete column
                break;
            case 'gpsy': //GPSy Complete column
                break;
            case 'reva': //General Availability (GA)
                editDate(objUnknown, ModuleID, RegionID, 'bomrevadate', repository, fullname);	
                break;
            case 'rasdisc': //End of Manufacturing (EM) column
                editDate(objUnknown, ModuleID, RegionID, 'rasdiscontinuedate', repository, fullname);
                break;
            case 'cpl': //Select Availability (SA) column
                editDate(objUnknown, ModuleID, RegionID, 'cplblinddate', repository, fullname);
                break;<% if IsODM = 0 then %>
            case 'cst': //AMO Cost column
                editCurrency(objUnknown, ModuleID, 'amocost', 20, repository, fullname);
                break;
            case 'wwp': //AMO Price column
                //editCurrency(objUnknown, ModuleID, 'amowwprice', 20, repository, fullname);
                break;
            case 'acst': //Actual Cost column
                editCurrency(objUnknown, ModuleID, 'actualcost', 20, repository, fullname);
                break;<% end if %>
            case 'rep': //Replacement column
                    editText(objUnknown, ModuleID, 'replacement', 30, repository, fullname);
                break;
            case 'alt': //Alternative column
                editText(objUnknown, ModuleID, 'alternative', 30, repository, fullname);
                break;
            case 'nw': //Net Weight column
                editText(objUnknown, ModuleID, 'netweight', 9, repository, fullname);
                break;
            case 'ew': //Export Weight column
                editText(objUnknown, ModuleID, 'exportweight', 9, repository, fullname);
                break;
            case 'apw': //Air Packed Weight column
                editText(objUnknown, ModuleID, 'airpackedweight', 9, repository, fullname);
                break;
            case 'apc': //Air Packed Cubic column
                editText(objUnknown, ModuleID, 'airpackedcubic', 9, repository, fullname);
                break;
            case 'ec': //Export Cubic column
                editText(objUnknown, ModuleID, 'exportcubic', 9, repository, fullname);
                break;
            case 'vna': //Visibility_NA
                editOption(objUnknown, ModuleID, 'Visibility_NA', 9, repository, fullname);
                break;
            case 'vem': //Visibility_EM
                editOption(objUnknown, ModuleID, 'Visibility_EM', 9, repository, fullname);
                break;
            case 'vap': //Visibility_AP
                editOption(objUnknown, ModuleID, 'Visibility_AP', 9, repository, fullname);
                break;
            case 'vla': //Visibility_LA		
                editOption(objUnknown, ModuleID, 'Visibility_LA', 9, repository, fullname);
                break;
            case 'warc': //WarrantyCode
                editText(objUnknown, ModuleID, 'warrantycode', 5, repository, fullname);
                break;
            case 'obd': //ObsoleteDate
                editDate(objUnknown, ModuleID, RegionID, 'obsoletedate', repository, fullname);
                break;
            case 'gbsd': //Global Series Config EOL
                editDate(objUnknown, ModuleID, RegionID, 'globalseriesdate', repository, fullname);
                break;
            case 'maco': //ManufactureCountry
                editText(objUnknown, ModuleID, 'manufacturecountry', 2, repository, fullname);
                break;
            case 'prodl': //product line
                editProductOption(objUnknown, ModuleID, 'productline', 9, repository, fullname);
                break;
            case 'sf':
                editSerialFlag(objUnknown,ModuleID,'serialflag',2,repository,fullname);
                break;
        }
    }

    function enterComment(ModuleID, Field) {
        thisform.ID.value = ModuleID
        thisform.Field.value = Field
        thisform.action = "AMO_AddComment.asp";
        thisform.submit ();
    }

    function editProductOption(evtobj, ModuleID, Field, intMaxlength, repository, fullname) {
        var sHTML
        var objUnknown = evtobj;
        if (objUnknown.innerHTML.indexOf("optioncell" + ModuleID + Field) < 0) {
            objUnknown.style.textDecoration = "none"; // get rid of underline

            // have to replace quotation marks too and escape it and unescape it in the getOption function
            // so single and double quotes work		
            var sTempValue = objUnknown.innerHTML.replace(/&nbsp;/g, "").replace(/"/g, "&quot;")
		
            sHTML = "<select id='optionprodcell" + ModuleID + Field + "' name='optionprodcell" + ModuleID + Field + "' ";
            //sHTML += "onBlur='javascript:getProductOption(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")' ";
            sHTML += "onchange='javascript:getProductOption(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")'>";
	
            <% if oRsProductLine.RecordCount > 0 then %>
                sHTML += "<option value=''></option>"
            <%	do while (not oRsProductLine.EOF) %>
                    sHTML += "<option value=" + "'<%=oRsProductLine("Description").Value%>'"
            if (sTempValue == "<%=oRsProductLine("Description").Value%>")
            sHTML += " selected"
				
            sHTML += "><%=oRsProductLine("Description").Value%></option>"
    <%	oRsProductLine.MoveNext
            loop
            end if
            %>
            objUnknown.innerHTML = sHTML 
                document.getElementById("optionprodcell" + ModuleID + Field).focus()
        }
    }
    //sug 9881
    function editSerialFlag(evtobj, ModuleID, Field, intMaxlength, repository, fullname) {
        var sHTML;
        var objUnknown = evtobj
        if (objUnknown.innerHTML.indexOf("optioncell" + ModuleID + Field) < 0) {
            objUnknown.style.textDecoration = "none"; // get rid of underline

            // have to replace quotation marks too and escape it and unescape it in the getOption function
            // so single and double quotes work		
            var sTempValue = objUnknown.innerHTML.replace(/&nbsp;/g, "").replace(/"/g, "&quot;")            
            sHTML = "<select id='optionSerialflag" + ModuleID + Field + "' name='optionSerialflag" + ModuleID + Field + "' ";
            sHTML += "onBlur='javascript:getSerialFlag(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")' ";
            sHTML += "onchange = 'javascript:getSerialFlag(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")'>";
            if (sTempValue != 'Yes' && sTempValue != 'No')
                sHTML += "<option value=''></option>";
            sHTML += "<option value='Yes'";         
            if (sTempValue == 'Yes') 
                sHTML += " selected"  ;            
            sHTML += " >Yes</option><option value='No'";             
            if (sTempValue == 'No') 
                sHTML += "selected";            
            sHTML += ">No</option></select>";       
            objUnknown.innerHTML = sHTML;
            document.getElementById("optionSerialflag" + ModuleID + Field).focus()
        }
    }
    function cM_Status(evt, ModuleID, sStatus) {
        var menu1=new Array()
        var wide = 120;
	
        if (sStatus == 'ras') {
            <% if bRAS then %>
            menu1.push("<a href='../nj.asp' onClick=\"ChangeStatus(" + ModuleID + ", <%= Application("AMO_REJECT") %>); return false;\">Reject</a>");
            <% end if%>
            } else {
            <% if bAMOUpdate or bCostUpdate then %>
                if (sStatus == 'new')
            menu1.push("<a href='../nj.asp' onClick=\"ChangeStatus(" + ModuleID + ", <%= Application("AMO_RASREVIEW") %>); return false;\">RAS Review</a>");
        else if (sStatus == 'asdis') 
            menu1.push("<a href='../nj.asp' onClick=\"ChangeDisabledStatus(" + ModuleID + ", <%= Application("AMO_RASUPDATE") %>); return false;\">RAS Update</a>");
        else
                menu1.push("<a href='../nj.asp' onClick=\"ChangeStatus(" + ModuleID + ", <%= Application("AMO_RASUPDATE") %>); return false;\">RAS Update</a>");
            <% end if %>
            }

        showmenu(evt, menu1.join(""), wide+'px');
    }

    function ChangeStatus(ModuleID, StatusID) {
        var msg
        if (StatusID == <%= Application("AMO_RASREVIEW") %>) {
            msg = "Are you sure you want to set this Option to RAS Review?"
	}
    if (StatusID == <%= Application("AMO_RASUPDATE") %>) {
		msg = "Are you sure you want to set this Option to RAS Update?"
    }
        if (StatusID == <%= Application("AMO_REJECT") %>) {
            msg = "Are you sure you want to Reject this Option?"
            //msg = "Are you sure that you want to reject this Option? If you would like to add comments supporting the status change, please click Cancel and update the Comments from RAS field\n"	
        }
	
            if (confirm(msg)) {
                thisform.ID.value = ModuleID 
                thisform.StatusID.value = StatusID
                thisform.PreviousStatusID.value = "<%=Application("AMO_RASUPDATE")%>";
                thisform.action = "AMO_ModuleList.asp?nMode=1";
                thisform.target = "_self"
                thisform.submit ();
            }
    }

    function lbxGo_onhide() {
        thisform.action = "AMO_ViewCustomize.asp?nParent=AMO_ModuleList";
        thisform.target = "_self"
        thisform.submit ();
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
		
        if (!checkDate (thisform.txtDisFromDate, "Show options with End of Manufacturing (EM)", true))
            return false;
        if (!checkDate (thisform.txtDisToDate, "Show options with End of Manufacturing (EM)", true))
            return false;
        if (!checkDate (thisform.txtMRFromDate, "Show options with PHweb (General) Availability (GA)", true))
            return false;
        if (!checkDate (thisform.txtMRToDate, "Show options with PHweb (General) Availability (GA)", true))
            return false;
        if (!checkDate (thisform.txtCPLFromDate, "Show options with Select Availability (SA)", true))
            return false;
        if (!checkDate (thisform.txtCPLToDate, "Show options with Select Availability (SA)", true))
            return false;
		
        if (!checkDate (thisform.txtRasObsoleteFromDate, "Show options with End of Sales (ES)", true))
            return false;
        if (!checkDate (thisform.txtRasObsoleteToDate, "Show options with End of Sales (ES)", true))
            return false;
		
        SelectAll(thisform.lbxSelectedCategory);
        if (thisform.lbxSelectedCategory.value == "") {
            alert("Please select at least one AMO Category before proceeding");
            return false;
        } 	

        if (thisform.lbxSelectedCategory.selectedIndex == -1)
            thisform.emptycategory.value = "1"
		
        thisform.action = "AMO_ModuleList.asp?nMode=1";	
        thisform.target = "_self"
        thisform.submit();
    }

    //*****************************************************************
    //Function:     ChangeRAS_GPSy();
    //Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
    //*****************************************************************
    function ChangeRAS_GPSy(nChkIndex, ModuleID) {
        var RASItem = document.getElementById("chkRAS" + ModuleID);
        var GPSyItem = document.getElementById("chkGPSy" + ModuleID);
        var RASStatus, GPSyStatus;
        var bSaveEntry = 0;
        var bSetComplete = 0;
        var msg;
        var ajaxurl = "";
	
        if (RASItem.checked == true && GPSyItem.checked == true) {
            msg = "You are about to change the status of this Option to Complete. Do you wish to continue?"
            if (confirm(msg)) {
                bSaveEntry = 1;
                bSetComplete = 1;
            } else {
                // uncheck last option, no save
                if (nChkIndex == 1){	// RAS
                    RASItem.checked = false;
                } else { // GPSy
                    GPSyItem.checked = false;
                }
            }
        } else {
            bSaveEntry = 1;
        }

        if (bSaveEntry == 1) {
            if (RASItem.checked == true){
                RASStatus = 1;
            }else{
                RASStatus = 0;
            }
	
            if (GPSyItem.checked == true){
                GPSyStatus = 1;
            }else{
                GPSyStatus = 0;
            }

            //var objRS = RSGetASPObject("AMO_RS.asp");
            var fullname = "<%= session("FullName") %>";

            //PBI 17492/ Task 20251 - Change all remote scripting in IRS AMO List to jQuery AJAX; similar to AVL_RasReview.asp 
            // save value changed
            if (nChkIndex == 1){	// RAS
                ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=sys_ras&Value=" + RASStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setFieldValue("<%=Application("REPOSITORY")%>", ModuleID, "sys_ras", escape(RASStatus), "<%= session("FullName") %>");
            } else { // GPSy
                ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=sys_gpsy&Value=" + GPSyStatus + "&FullName=" + fullname  + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setFieldValue("<%=Application("REPOSITORY")%>", ModuleID, "sys_gpsy", escape(GPSyStatus), "<%= session("FullName") %>");
            }
            //if (objResult.return_value != "") {
            //erroutputArea.innerHTML = "<p><font color=red>" + objResult.return_value + "<\/font><\/p>";
            //}

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

            if (bSetComplete == 1) {
                thisform.ID.value = ModuleID ;
                thisform.StatusID.value = <%= Application("AMO_COMPLETE") %>
                thisform.action = "AMO_ModuleList.asp?nMode=1";
                thisform.target = "_self";
                thisform.submit ();
            }
        }
    }

    function btnChange_onClick(bSwitch) {
        if (bSwitch == 1) {
            thisform.fromAMO.value = '1'
        } else {
            thisform.fromAMO.value = ''
        }
        thisform.action = "AMO_ModuleList.asp?nMode=1";	
        thisform.target = "_self"
        thisform.submit();	
    }

    //*****************************************************************
    //Function:     BulkstatusChange();
    //Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
    //              Harris, Valerie (8/30/2016) - Bug 25663/ Task 25664 
    //*****************************************************************
    function BulkstatusChange() {
        var strModuleIDs = "";	
        var strAllModuleIDs = "";	
        var msg = "";
        var errormsg = "";
        var arrModId_Status;
        var strModId_Status;
        var ajaxurl = "";
        var iEmptyRASDiscontinue = 0;
        var sRASDiscontinue = "";

        var collStatus = document.getElementsByName("chkBlkStatus"); //thisform["chkBlkStatus"];
        var collActions = document.thisform["CboActions"];
	
        if (collStatus == null) {
            alert("There are no options to select.");
            return false;
        }

        for (var i=0;i< collStatus.length; i++) {
            if (collStatus[i].checked) {
                strModId_Status = collStatus[i].value;
                arrModId_Status = strModId_Status.split("|");
                strAllModuleIDs = strAllModuleIDs + "," + arrModId_Status[0];

                if(document.getElementById("rasdisc"+arrModId_Status[0]+""+arrModId_Status[3]+"")){
                    sRASDiscontinue = document.getElementById("rasdisc"+arrModId_Status[0]+""+arrModId_Status[3]+"").innerHTML;
                }

                if (collActions.value == "1") { //Complete
                    if (arrModId_Status[1] == "RAS Review" || arrModId_Status[1] == "RAS Update") {				
                        strModuleIDs = strModuleIDs + "," + arrModId_Status[0]; 
                    }
                    if(sRASDiscontinue == "&nbsp;" || sRASDiscontinue == " " || sRASDiscontinue == "-"){
                        iEmptyRASDiscontinue = iEmptyRASDiscontinue + 1;
                    }
                }else if (collActions.value == "2") { //RAS Review
                    if (arrModId_Status[1] == "New") {				
                        strModuleIDs = strModuleIDs + "," + arrModId_Status[0];
                    }
                } else if (collActions.value == "3") { //RAS Update
                    if (arrModId_Status[1] != "New" && arrModId_Status[1] != "RAS Update" && arrModId_Status[1] != "RAS Review"){ 				
                        strModuleIDs = strModuleIDs + "," + arrModId_Status[0];
                    }
                } else { //Reject
                    if (arrModId_Status[1] == "RAS Update" || arrModId_Status[1] == "RAS Review"){ 				
                        strModuleIDs = strModuleIDs + "," + arrModId_Status[0]; 
                    }
                }
            }
        }	
        
        if (strModuleIDs != ""){
            strModuleIDs = strModuleIDs.slice(1);
        }

        if (strAllModuleIDs == "") {
            alert("Please select at least one option for status change.");
            return false;
        }	

        if(collActions.value == "1" && iEmptyRASDiscontinue > 0){
            alert("Please select only options that have an End of Manufacturing date to change the status to Complete.");
            return false;
        }
	
        //var objRS = RSGetASPObject("AMO_RS.asp");		
	
        //PBI 17492/ Task 20251 - Change all remote scripting in IRS AMO List to jQuery AJAX; similar to AVL_RasReview.asp 
        if (collActions.value == "1") {	
            if (strAllModuleIDs != "" && strModuleIDs == "") {
                alert("Only options with RAS Review or RAS Update status can be set to Complete");
                return false;
            }	
		
            msg += "Only options with RAS Review or RAS Update status will be set to Complete. Are you sure you want to complete the appropriate selected options?\n"
				
            if (confirm(msg)) {
                var fullname = "<%= session("FullName") %>";
                ajaxurl = "AMO_SetBulkStatusValue.asp?RGS=3&ModuleID=" + strModuleIDs + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setBulkStatusValue("<%=Application("REPOSITORY")%>", <%=Application("AMO_COMPLETE")%>, strModuleIDs, "<%= session("FullName") %>");
				
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
                    thisform.action = "AMO_ModuleList.asp?nMode=1";
                    thisform.target = "_self"
                    thisform.submit ();
                }								
            }
        }  else if (collActions.value == "2") {	
            if (strAllModuleIDs != "" && strModuleIDs == "") {
                alert("Only options with New status can be set to RAS Review");
                return false;
            }	
			
            msg += "Only options with New status will be set to RAS Review. Are you sure you want to set to RAS Review the appropriate selected options?\n"
				
            if (confirm(msg)) {
                var fullname = "<%= session("FullName") %>";
                ajaxurl = "AMO_SetBulkStatusValue.asp?RGS=1&ModuleID=" + strModuleIDs + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setBulkStatusValue("<%=Application("REPOSITORY")%>", <%=Application("AMO_RASREVIEW")%>, strModuleIDs, "<%= session("FullName") %>");
			
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
                    thisform.action = "AMO_ModuleList.asp?nMode=1";
                    thisform.target = "_self"
                    thisform.submit ();
                }								
            }
        } else if (collActions.value == "3") {	
            if (strAllModuleIDs != "" && strModuleIDs == "") {
                alert("Only options with In Process, Complete, Enabled, and Reject status can be set to RAS Update");
                return false;
            }	
		
            msg += "Only options with In Process, Complete, Enabled, and Reject status can be set to RAS Update. Are you sure you want to set to RAS Update the appropriate selected options?\n"
			
            if (confirm(msg)) {
                var fullname = "<%= session("FullName") %>";
                ajaxurl = "AMO_SetBulkStatusValue.asp?RGS=2&ModuleID=" + strModuleIDs + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setBulkStatusValue("<%=Application("REPOSITORY")%>", <%=Application("AMO_RASUPDATE")%>, strModuleIDs, "<%= session("FullName") %>");
			
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
                    thisform.action = "AMO_ModuleList.asp?nMode=1";
                    thisform.target = "_self"
                    thisform.submit ();
                }								
            }
        } else {
            if (strAllModuleIDs != "" && strModuleIDs == "") {
                alert("Only options with RAS Review or RAS Update status can be set to Reject");
                return false;
            }
            msg += "Only options with RAS Review or RAS Update status will be set to Reject. Are you sure you want to reject the appropriate selected options?\n"
            if (confirm(msg)) {
                var fullname = "<%= session("FullName") %>";
                ajaxurl = "AMO_SetBulkStatusValue.asp?RGS=4&ModuleID=" + strModuleIDs + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //var objResult = objRS.setBulkStatusValue("<%=Application("REPOSITORY")%>", <%=Application("AMO_REJECT")%>, strModuleIDs, "<%= session("FullName") %>");

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
                    thisform.action = "AMO_ModuleList.asp?nMode=1";
                    thisform.target = "_self"
                    thisform.submit ();
                }			
            }
        }
    }

    function ChangeDisabledStatus(ModuleID, StatusID) {
        var msg = "Are you sure you want to set this Option to RAS Update?"
	
        if (confirm(msg)) {
            thisform.ID.value = ModuleID 
            thisform.StatusID.value = StatusID
            thisform.PreviousStatusID.value = "<%=Application("AMO_DISABLED")%>";
            thisform.action = "AMO_ModuleList.asp?nMode=1";
            thisform.target = "_self"
            thisform.submit ();
        }
    }

    function KeyWordSearch()
    {
        var keyword = $("#txtKeyWord").val();
	
        //var keyword = thisform.document.getElementById("txtKeyWord").value;
        var Start_index = keyword.indexOf("#"); //if keyword contain # char, then remove this char and everyhting after that (must likely user is trying to find a part number)
        if(Start_index != -1)
        {
            keyword = keyword.substring(0, Start_index); 
        }
        keyword = keyword.replace(/^\s+|\s+$/g, '') ;

        var strspc =  keyword.search("'");
        //alert(keyword);
        if (strspc == -1)
        {
            if (keyword !="")
            {
			
                thisform.action = "AMO_ModuleList.asp?nMode=12";	
                thisform.target = "_self"
                thisform.submit();
                return true;
            }
            else
            {
                alert("Please provide a key word to search for.");
                return false;
            }
        }
        else
        {
            alert("Single quote(') is not allowed to be included in the keyword.");
            return false;
        }
    }
        //-->
    </script>
</head>

<body bgcolor="#FFFFFF">
    <!-- #include file="../library/includes/popup.inc" -->
    <h1 class="page-title"><%=sHeader%></h1>
    <form name="thisform" id="thisform" method="post">
        <%
Response.Write ""
WriteTabs "All" & "&nbsp;" & "Options"

if sErr <> "" then
	Response.Write sErr 
else
%>
        <div id="erroutputArea"></div>
        There are two ways to find AMO Options and display them in a table. Either use the Search feature or use the Filters. If the Search is used, it will not use or affect the Filters so they will remain unchanged.<br>
        <br>
        Search for a key word
        <input type="text" name="txtKeyWord" id="txtKeyWord" maxlength="64" style="width: 350px;" value="<%=ckKeyWord %>">
        <input type="button" name="btnKeyWordSaerch" id="btnKeyWordSearch" value="Search" onclick="return KeyWordSearch()">
        <font style="font-style: italic; font-size: 11px;">Searches Marketing Description, Short Description, Long Description, and HP Part Number</font>
        <hr>
        <table border="0" cellpadding="1" cellspacing="1" width="100%">
            <% Call WriteCategoryFilter("All" & "&nbsp;" & "Options", sModuleCategoryHTML, sStatusHTML, sBusSegHTML, strEOLDate,_
		 bRAS, sGpsyCom, sRasCom, sDisplayHideromMOL, sDisplayHideromSCM, sDisplayHideromSCL, strMRFromDate, strMRToDate, _
		 strCPLFromDate, strCPLToDate, strDisFromDate, strDisToDate, strchkMRBlank, strchkCPLBlank, strchkDisBlank, _
		 strRasObsoleteFromDate, strRasObsoleteToDate, nCategoriesSelected, sOwnersHTML, sDisplayIncludeBlankGSD) %>
            <%if sHideColumns <> "" then%>
            <tr>
                <td colspan='3'><font color="red">Warning: some columns are hidden, click on "Select Column To Hide" to deselect hidden columns</font></td>
            </tr>
            <%end if %>

            <tr>
                <td colspan="2">
        <% '4/5/2004, removed bRAS checking
		if bRAS and bAMOCreate then %>
                    <input id='btnSelAll' name='btnSelAll' type='button' value='Select All' language='javascript' onclick="return btnSelectAll_Click()">&nbsp;
			<input id='btnDelSelAll' name='btnDeSelAll' type='button' value='Deselect All' language='javascript' onclick="return btnDeselectAll_Click()">&nbsp;
			&nbsp;&nbsp;Status Change:
                    <select name="CboActions" id="CboActions">
                        <option value="1">Complete</option>
                        <option value="2">RAS Review</option>
                        <option value="3">RAS Update</option>
                        <option value="4">Reject</option>
                    </select>
                    <input type="button" name="btnGo" value='Go' onclick="return BulkstatusChange()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%elseif bRas Then%>
                    <input id='btnSelAll' name='btnSelAll' type='button' value='Select All' language='javascript' onclick="return btnSelectAll_Click()">&nbsp;
			<input id='btnDelSelAll' name='btnDeSelAll' type='button' value='Deselect All' language='javascript' onclick="return btnDeselectAll_Click()">&nbsp;
			&nbsp;&nbsp;Status Change:
                    <select name="CboActions" id="CboActions">
                        <option value="1">Complete</option>
                        <option value="4">Reject</option>
                    </select>
                    <input type="button" name="btnGo" value='Go' onclick="return BulkstatusChange()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%elseif bAMOCreate Then%>
                    <input id='btnSelAll' name='btnSelAll' type='button' value='Select All' language='javascript' onclick="return btnSelectAll_Click()">&nbsp;
			<input id='btnDelSelAll' name='btnDeSelAll' type='button' value='Deselect All' language='javascript' onclick="return btnDeselectAll_Click()">&nbsp;		
			&nbsp;&nbsp;Status Change:
                    <select name="CboActions" id="CboActions">
                        <option value="2">RAS Review</option>
                        <option value="3">RAS Update</option>
                    </select>
                    <input type="button" name="btnGo" value='Go' onclick="return BulkstatusChange()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%end if %>


        <% if bAMOCreate then %>
            <input type='button' value='Create AMO Features' style="width: 150px" id="btnAdd" name="btnAdd" language="javascript" onclick="return btnAdd_onClick();">&nbsp;&nbsp;
			<input type='button' value='Bulk Date Change' style="width: 150px" id="btnAdd" name="btnAdd" language="javascript" onclick="return btnBulkDateChange_onClick();">&nbsp;&nbsp;			
		<%end if%>
                </td>
            </tr>
        </table>


        <table id="tblAMOList" border="0" cellpadding="1" cellspacing="1" width="100%">
            <colgroup></colgroup>
            <%
		if nNumTotalModules <= 0 then
			Response.Write "<tr><td>"
			Response.Write "No AMO options have been found for the above filter.<br><br><br>" 
			Response.Write "</td></tr>" & vbCrLf
			Response.Write "<tr><td>"
		else
            %>
            <tr>
                <td align="left">
                    <%
			if bRAS then
				WriteRASUpdateableGridHTML "All" & "&nbsp;" & "Options", oRsAMOModules, sDivisionIDs, sGroupIDs, Trim(sHideColumns)
			else
				WriteAMOUpdateableGridHTML "All" & "&nbsp;" & "Options", oRsAMOModules, sDivisionIDs, sGroupIDs, Trim(sHideColumns)
			end if
                    %>
                </td>
            </tr>
            <tr>
                <td>
                    <%
		end if 
                    %>
                    <div id="errorreport" name="errorreport"></div>

                </td>
            </tr>
        </table>

        <table>
            <tr>
                <td><%
			if oRsAMOModules.RecordCount > 0 then
				if oRsAMOModules.RecordCount = 1 then
					response.write "1 Option displayed"
				else
					response.write oRsAMOModules.RecordCount & " Options displayed"
				end if
			end if
                %></td>
            </tr>
            <tr>
                <td bgcolor='#FFFF99'>Highlighted Cell = Value changed. A Complete status resets.</td>
            </tr>
            <tr>
                <td>*=Required for RAS/GPSy.</td>
            </tr>
        </table>
        <%
	set oSvr = nothing
	oRsAMOModules.Close
	set oRsAMOModules = nothing
	set oRsProductLine = Nothing
end if 'no error
        %>
        <input type="hidden" name="ID" value="">
        <input type="hidden" name="Field" value="">
        <input type="hidden" name="StatusID" value="">
        <input type="hidden" name="PreviousStatusID" value="">
        <input type="hidden" name="emptycategory" value="">
        <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
        <input type="hidden" id="inpWebServer" value="<%=Application("IRSWebServer") %>"" />
        <!--<input type="hidden" name="txtModuleIDs" id = "txtModuleIDs" value="">-->
    </form>
</body>
</html>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->


