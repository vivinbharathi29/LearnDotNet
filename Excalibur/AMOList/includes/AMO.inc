<!-- #include file="../../includes/emailwrapper.asp" -->
<%
const ASCENDING = "ASC"
const DESCENDING = "DESC"
dim IsODM

if cint(session("IsODM")) = 0 then
	IsODM = 0	'HP user
else
	IsODM = 1	'ODM user
end if

'********************************************************************************
'* Put up or down arrow depending on which way the current sort is
'* 
'* Parameters:
'*  strCurrentField  = current field name
'*  sSortField       = field where the sort is
'*  sSortOrder       = Ascending or Descending sort order
'********************************************************************************
sub PutArrow(strCurrentField, sSortField, sSortOrder)
	if lcase(strCurrentField) = lcase(sSortField) then
		'we have the correct sort field, so determine what arrow to put
		select case sSortOrder
			case DESCENDING
				Response.Write "<img src=""../library/Images/gifs/downarrow.gif"" width=""10"" height=""9"" border=""0"" alt=""Descending"">"
			case ASCENDING
				Response.write "<img src=""../library/Images/gifs/uparrow.gif"" width=""10"" height=""9"" border=""0"" alt=""Ascending"">"
		end select
	end if
end sub

function WriteTabs(sActiveTab)
	dim sDefaultTab, sTabWidth
	dim sShorttabWidth
	if sActiveTab = "" then
		sDefaultTab = "All" & "&nbsp;" & "Options"
	else
		sDefaultTab = sActiveTab
	end if
	
	
	sTabWidth = "110px"
	sShorttabWidth = "110px"
	Response.write "<UL ID=""selectmenu""> "
	response.write WriteTab("All" & "&nbsp;" & "Options", sDefaultTab, "AMO_ModuleList.asp", sTabWidth)

	response.write WriteTab("WIP", sDefaultTab, "AMO_ModuleList_WIP.asp", sShorttabWidth)

	response.write WriteTab("RAS" & "&nbsp;" & "Review", sDefaultTab, "AMO_ModuleList_RAS.asp", sTabWidth)

	response.write WriteTab("Complete", sDefaultTab, "AMO_ModuleList_Complete.asp", sTabWidth)

	response.write WriteTab("Discontinue", sDefaultTab, "AMO_ModuleList_Discontinue.asp", sTabWidth)

	response.write WriteTab("Products", sDefaultTab, "AMO_Platforms.asp", sTabWidth)
	
	'7/12/16 - Per Efren, Remove Rules Tab:-------
	'response.write WriteSeparator
	'response.write WriteTab("Rules", sDefaultTab, "AMO_OptionCategoryRule.asp?nMol=1", sTabWidth)

	response.write WriteTab("Reports", sDefaultTab, "AMO_Reports.asp?nMol=1", sTabWidth)

	response.write WriteTab("Change History", sDefaultTab, "AMO_ChangeHistory.asp", sTabWidth)

	response.write "</UL> "
	response.write "<DIV ID=""menu-bar"" CLASS=""gray""></DIV> "

end function

function WriteTab(sTab, sDefaultTab, sLink, sTabWidth)
	dim sClass, sBGColor, sFontColor
	if sTab = sDefaultTab then
		sClass = "white current"
	else
		sClass = "white"
	end if
	response.write "<LI> " & vbCrLf
	response.write "<DIV ><A HREF=""" & sLink & """ CLASS="""&sClass&""">" & sTab & "</A><DIV>" & vbCrLf
	response.write "</LI> " & vbCrLf
end function

function WriteSeparator()
	response.write "    <TD BGCOLOR=""#FFFFFF"" WIDTH=""1""><FONT SIZE=""2"">&nbsp;</FONT></TD>" & vbCrLf
end function

'Format a number so it has 2 decimal places
function TwoPlaces(value)
	if IsNULL(value) or Trim(Value) = "" then
		TwoPlaces = ""
	else
		TwoPlaces = FormatNumber(value, 2)
	end if
end function

function ConvertNull(Byref Value)
	if IsNUll(Value) or Trim(Value)="" then 
		ConvertNull = "&nbsp;"
	else
		ConvertNull = Value
	end if
end function

'Get filter values from each tab and create Where clause for Module search
function GetFilterInfo( ByVal nCategoriesSelected, ByVal sStatusIDs, ByVal sBusSegIDs,_
						 ByVal strEOLDate, ByRef sModuleCategoryHTML, _
						 ByRef sStatusHTML, ByRef oRsReturn, ByRef nNumTotalModules,_
						  ByRef sBusSegHTML, ByVal bGpsyCom, ByVal bRasCom, ByVal bRAS,_
						  ByVal bDisplayHideromMOL,  ByVal bDisplayHideromSCM, ByVal bDisplayHideromSCL, ByVal strMRFromDate, ByVal strMRToDate, _
						  ByVal	strCPLFromDate, ByVal strCPLToDate, ByVal strDisFromDate, ByVal strDisToDate, _
						  ByVal	strchkMRBlank, ByVal strchkCPLBlank, ByVal strchkDisBlank, ByVal strRasObsoleteFromDate, _
						  ByVal strRasObsoleteToDate, ByRef sOwnersHTML, ByVal sOwnerIDs, ByVal bDisplayIncludeBlankGSD, ByVal strKeyWord, ByVal UseBusSeg)
						  
	dim sErr, sFilter, sCtrlStyle, bHideOwners
	dim oErr, oRsAllCategory, oRsStatus, oRsBusSeg, oSvr, sDateFilter, oRsGroupOwner
	'4/28/2005, need to re-construct dates filters 
	
	'Hide Owned By in Pulsar, not used
	bHideOwners = False

	'only get active modules
	sFilter =  " and (R.SCMID = 1 Or R.SCMID = null) and O.SCMID = 1 "
	
	'PBI 17498/ Task 21007 - Changes for AMO List tabs 
	If UseBusSeg <> "NewBus" Then
		if nCategoriesSelected <> "" then
			sFilter = sFilter & "  and C.CategoryID in (" & nCategoriesSelected & ")"
		else
			sFilter = sFilter & "  and C.CategoryID = -1"
		end if
	Else
		if nCategoriesSelected <> "" then
			sFilter = sFilter & "  and M.FeatureCategoryID in (" & nCategoriesSelected & ")"
		else
			sFilter = sFilter & "  and M.FeatureCategoryID = -1"
		end if
	End If

	
	if sErr = "" then
		'AMO Status
		'set oErr = GetMOLCategory(oRsStatus, 24)
		
		set oRsStatus = GetMOLCategory(24)	
		if (oRsStatus is Nothing) then
			Response.Write("Recordset error: oRsStatus")
			Response.End()
		else
			if not oRsStatus is nothing then
				if oRsStatus.RecordCount > 0 then
					oRsStatus.Filter = "Constant <> 'AMO_OBSOLETE'"
					oRsStatus.Sort = "Description ASC"
					oRsStatus.MoveFirst
					sStatusIDs = replace(sStatusIDs, " ", "")
					Do until oRsStatus.EOF
						if sStatusHTML <> "" then 
							sStatusHTML = sStatusHTML & "&nbsp;"
						end if
						if instr("," & sStatusIDs & ",", "," & cstr(oRsStatus("StatusID").Value) & ",") > 0 then 
							
							sStatusHTML = sStatusHTML & "<INPUT type='checkbox' id=chkStatus NAME=chkStatus value=" & oRsStatus("StatusID").Value & " checked " & sCtrlStyle & ">" & oRsStatus("Description").Value
						else
							
							sStatusHTML = sStatusHTML & "<INPUT type='checkbox' id=chkStatus NAME=chkStatus value=" & oRsStatus("StatusID").Value & ">" & oRsStatus("Description").Value
						end if
						oRsStatus.MoveNext
					Loop
					
				end if
				
			end if
			if sStatusIDs <> "" then 
				sFilter = sFilter & " and S.StatusID in (" & sStatusIDs & ")" 
			end if
		end if
	end if
	
	if sErr = "" then
		'Business segments
		'set oErr = GetMOLCategory(oRsBusSeg, 28)
		If UseBusSeg <> "NewBus" Then
			set oRsBusSeg = GetMOLCategory(28) 'Use Divisions
		Else
			set oRsBusSeg = GetMOLCategory(34) 'Use Pulsar Business Segments
		End If
			
		if (oRsBusSeg is Nothing) then
			Response.Write("Recordset error: oRsBusSeg")
			Response.End()
		else
			if not oRsBusSeg is nothing then
			
				if oRsBusSeg.RecordCount > 0 then
					'oRsBusSeg.Filter = "State=1"
					oRsBusSeg.Sort = "Description ASC"
					if sBusSegIDs = "" then
						'make every one checked
						oRsBusSeg.MoveFirst
						do until oRsBusSeg.EOF
							sBusSegIDs = sBusSegIDs & oRsBusSeg("BusinessSegmentID").Value & ","
							oRsBusSeg.MoveNext
						loop
						if right(sBusSegIDs, 1) = "," then
							sBusSegIDs = left(sBusSegIDs, len(sBusSegIDs)-1)
						end if

						'store the cookie
						If UseBusSeg <> "NewBus" Then
							Call SaveDBCookie( "AMO chkBusSeg", sBusSegIDs)	'Save BusSeg for Divisions
						Else
							Call SaveDBCookie( "AMO chkFeatureBusSeg", sBusSegIDs)	'Save BusSeg for Business Segments
						End If

					end if
					sBusSegIDs = replace(sBusSegIDs, " ", "")
					oRsBusSeg.MoveFirst
					Do until oRsBusSeg.EOF
						if sBusSegHTML <> "" then 
							sBusSegHTML = sBusSegHTML & "&nbsp;"
						end if
						if instr("," & sBusSegIDs & ",", "," & cstr(oRsBusSeg("BusinessSegmentID").Value) & ",") > 0 then 
							sBusSegHTML = sBusSegHTML & "<INPUT type='checkbox' id=chkBusSeg NAME=chkBusSeg value=" & oRsBusSeg("BusinessSegmentID").Value & " checked " & sCtrlStyle & ">" & oRsBusSeg("SegmentName").Value
						else
							sBusSegHTML = sBusSegHTML & "<INPUT type='checkbox' id=chkBusSeg NAME=chkBusSeg value=" & oRsBusSeg("BusinessSegmentID").Value & ">" & oRsBusSeg("SegmentName").Value
						end if
						oRsBusSeg.MoveNext
					Loop

				end if
			end if
			if sBusSegIDs <> "" then
				If UseBusSeg <> "NewBus" Then
					sFilter = sFilter & " and MD.DivisionID in (" & sBusSegIDs & ")"  'Search Business Segments by Division
				Else 
					sFilter = sFilter & " and AFD.DivisionID in (" & sBusSegIDs & ")" 'Search Business Segments by Feature Business Segments
				End If
			end if
		end if
	end if
	
	
	'if sErr = "" then

		'If bHideOwners = True Then
			'Group Owners
			'set oErr = GetCategory(oRsGroupOwner, 206, 0)
		
			'set oRsGroupOwner = GetCategory(206, 0)	
			'if (oRsGroupOwner is Nothing) then
				'Response.Write("Recordset error: oRsGroupOwner")
				'Response.End()
			'else
				'if not oRsGroupOwner is nothing then
					'if oRsGroupOwner.RecordCount > 0 then
						'oRsGroupOwner.Sort = "Name ASC"
						'if sOwnerIDs = "" then
							'make every one checked
							'oRsGroupOwner.MoveFirst
							'do until oRsGroupOwner.EOF
								'sOwnerIDs = sOwnerIDs & oRsGroupOwner("GroupID").Value & ","
								'oRsGroupOwner.MoveNext
							'loop
							'if right(sOwnerIDs, 1) = "," then
								'sOwnerIDs = left(sOwnerIDs, len(sOwnerIDs)-1)
							'end if
							'store the cookie
							'Call SaveDBCookie( "AMO chkGroupOwner", sOwnerIDs )
						'end if
						'sOwnerIDs = replace(sOwnerIDs, " ", "")
						'oRsGroupOwner.MoveFirst
						'Do until oRsGroupOwner.EOF
							'if sOwnersHTML <> "" then 
								'sOwnersHTML = sOwnersHTML & "&nbsp;"
							'end if
							'if instr("," & sOwnerIDs & ",", "," & cstr(oRsGroupOwner("GroupID").Value) & ",") > 0 then 
								'sOwnersHTML = sOwnersHTML & "<INPUT type='checkbox' id=chkGroupOwner NAME=chkGroupOwner value=" & oRsGroupOwner("GroupID").Value & " checked " & sCtrlStyle & ">" & oRsGroupOwner("Name").Value
							'else
								'sOwnersHTML = sOwnersHTML & "<INPUT type='checkbox' id=chkGroupOwner NAME=chkGroupOwner value=" & oRsGroupOwner("GroupID").Value & ">" & oRsGroupOwner("Name").Value
							'end if
							'oRsGroupOwner.MoveNext
						'Loop
					'end if
				'end if
				'if sOwnerIDs <> "" then 
					'sFilter = sFilter & " and M.GroupID in (" & sOwnerIDs & ")" 
				'end if
			'end if
		'end if
	'End If
			
	if sErr = "" then
		
		sDateFilter =""
		
		if strMRFromDate <> "" and strMRToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & "  and (BOMRevADate >= '" & strMRFromDate & "' and  BOMRevADate <= '" &  strMRToDate & "' or BOMRevADate is null) "
			else
				sDateFilter  = sDateFilter  & "  and  (BOMRevADate >= '" & strMRFromDate & "' and  BOMRevADate <= '" &  strMRToDate & "') "
			end if
		elseif strMRFromDate <> "" and strMRToDate = "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & "  and  (BOMRevADate >= '" & strMRFromDate & "' or BOMRevADate is null) "
			else
				sDateFilter  = sDateFilter  & "  and  (BOMRevADate >= '" & strMRFromDate & "') "
			end if
		
		elseif strMRFromDate = "" and strMRToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & "  and  (BOMRevADate <= '" &  strMRToDate & "' or BOMRevADate is null) "
			else
				sDateFilter  = sDateFilter  & "  and (BOMRevADate <= '" &  strMRToDate & "') "
			end if
		else
			if strchkMRBlank = "" then
				sDateFilter  = sDateFilter  & "  and  (BOMRevADate is not null) "
			end if
			
		end if
		
		if strCPLFromDate <> "" and strCPLToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate >= '" & strCPLFromDate & "' and  CPLBlindDate <= '" &  strCPLToDate & "' or CPLBlindDate is null) "
			else
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate >= '" & strCPLFromDate & "' and  CPLBlindDate <= '" &  strCPLToDate & "') "
			end if
		elseif strCPLFromDate <> "" and strCPLToDate = "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate >= '" & strCPLFromDate & "' or CPLBlindDate is null) "
			else
				sDateFilter  = sDateFilter  & "  and  (CPLBlindDate >= '" & strCPLFromDate & "') "
			end if
		
		elseif strCPLFromDate = "" and strCPLToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate <= '" &  strCPLToDate & "' or CPLBlindDate is null) "
			else
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate <= '" &  strCPLToDate & "') "
			end if
		else
			if strchkMRBlank = "" then
				sDateFilter  = sDateFilter  & " and  (CPLBlindDate is not null) "
			end if
		end if
		
		
		if strDisFromDate <> "" and strDisToDate <> "" then
			if strchkDisBlank = "checked" then
				sDateFilter  = sDateFilter  & " and (RASDiscontinueDate >= '" & strDisFromDate & "' and RASDiscontinueDate <= '" &  strDisToDate & "' or RASDiscontinueDate is null) "
			else
				sDateFilter  = sDateFilter  & " and (RASDiscontinueDate >= '" & strDisFromDate & "' and  RASDiscontinueDate <= '" &  strDisToDate & "') "
			end if
		elseif strDisFromDate <> "" and strDisToDate = "" then
			if strchkDisBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (RASDiscontinueDate >= '" & strDisFromDate & "' or RASDiscontinueDate is null) "
			else
				sDateFilter  = sDateFilter  & " and  (RASDiscontinueDate >= '" & strDisFromDate & "') "
			end if
		
		elseif strDisFromDate = "" and strDisToDate <> "" then
			if strchkDisBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (RASDiscontinueDate <= '" &  strDisToDate & "' or RASDiscontinueDate is null) "
			else
				sDateFilter  = sDateFilter  & " and (RASDiscontinueDate <= '" &  strDisToDate & "') "
			end if
		else
			if strchkDisBlank = "" then
				sDateFilter  = sDateFilter  & " and (RASDiscontinueDate is not null) "
			end if
		end if
			
		
		if strRasObsoleteFromDate <> "" and strRasObsoleteToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & "  and (ObsoleteDate >= '" & strRasObsoleteFromDate & "' and  ObsoleteDate <= '" &  strRasObsoleteToDate & "' or ObsoleteDate is null) "
			else
				sDateFilter  = sDateFilter  & "  and (ObsoleteDate >= '" & strRasObsoleteFromDate & "' and  ObsoleteDate <= '" &  strRasObsoleteToDate & "') "
			end if
		elseif strRasObsoleteFromDate <> "" and strRasObsoleteToDate = "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & " and  (ObsoleteDate >= '" & strRasObsoleteFromDate & "' or ObsoleteDate is null) "
			else
				sDateFilter  = sDateFilter  & "  and  (ObsoleteDate >= '" & strRasObsoleteFromDate & "') "
			end if
		
		elseif strRasObsoleteFromDate = "" and strRasObsoleteToDate <> "" then
			if strchkMRBlank = "checked" then
				sDateFilter  = sDateFilter  & " and (ObsoleteDate <= '" &  strRasObsoleteToDate & "' or ObsoleteDate is null) "
			else
				sDateFilter  = sDateFilter  & " and  (ObsoleteDate <= '" &  strRasObsoleteToDate & "') "
			end if
		else
			if strchkMRBlank = "" then
				sDateFilter  = sDateFilter  & " and  (ObsoleteDate is not null) "
			end if
		end if
		if sDateFilter  <> "" then
			sDateFilter  = " and  M.FeatureID in ( select distinct FeatureID from AMOFeature_Region  where  SCMID =1 " & sDateFilter  & ") "
		end if 
		sFilter = sFilter & sDateFilter
		
		'Display Hide from MOL/PRL
		if not bDisplayHideromMOL then
			sFilter = sFilter & " and O.MOLHide = 0  "
		end if 

		'Display Hide from SCM checked items'
		if not bDisplayHideromSCM then
			sFilter = sFilter & " and O.HideFromSCM = 0  "
		end if 

		'NOT USED - Removed Disaply Hide from SCL from PULSAR AMO List Filter
		'if not bDisplayHideromSCL then
			'sFilter = sFilter & " and O.HideFromSCL = 0  "
		'end if
		
		'Display Blank GSD also
		if not bDisplayIncludeBlankGSD then
			sFilter = sFilter & " and isnull(convert(varchar(10), R.GlobalSeriesDate, 101), '''') <> ''''" 
		end if	
		
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
		set oSvr = New ISAMO
		'Response.Write(sFilter)
		'Response.End()
		'sFilter = " and (R.SCMID = 1 Or R.SCMID = null) and O.SCMID = 1 and FeatureID = 5663 and R.RegionID = " & Application("AMO_GLOBAL_REGIONID") & ""
		set oRsReturn = oSvr.AMOModule_Search(Application("REPOSITORY"), sFilter, strKeyWord, null, null)
		if not (oRsReturn is nothing) then
			'oRsReturn.Sort = "FeatureCategory ASC, Description ASC"
			if not oRsReturn.EOF and not oRsReturn.BOF then 
			  if bRAS then			
				if bGpsyCom and bRasCom then
					'oRsReturn.Filter = "Ras = 0 and GPSy = 0"
					oRsReturn.Filter = "(AMOStatus = 'RAS Update' and Ras = 0 and GPSy <> 1) or (AMOStatus = 'RAS Review' and Ras = 0 and GPSy <> 1) or (AMOStatus = 'RAS Update' and GPSy = 0 and Ras <> 1) or (AMOStatus = 'RAS Review' and GPSy = 0 and Ras <> 1)"
				elseif bGpsyCom and Not bRasCom then
					'oRsReturn.Filter = "GPSy = 0"
					oRsReturn.Filter = "(GPSy = 0 and AMOStatus = 'RAS Update') or (GPSy = 0 and AMOStatus = 'RAS Review')"
				elseif Not bGpsyCom and bRasCom then
					'oRsReturn.Filter = "Ras = 0"
					oRsReturn.Filter = "(Ras = 0 and AMOStatus = 'RAS Update') or (Ras = 0 and AMOStatus = 'RAS Review')"
				end if	
			   end if			
			   nNumTotalModules = oRsReturn.RecordCount
			end if
		end if
	end if
	GetFilterInfo = sErr
end function


'*************************************************************************************
'* Sub/Function	: WriteAMOUpdateableGridHTML
'* Description	: Write the module data for the Options tab when NOT in RAS mode
'* Modified		: 07/13/16 - Harris, Valerie - Update sub to list AMO Features
'* Note			: In Pulsar, the Global Region ID is 334; In IRS, the Global Region ID is 43. 
'*				 Only AMO Features and Modules with Global Region IDs will editable fields for most columns.
'*				 Only the date fields are editable for AMO Features with Localized Region IDs assigned.
'*************************************************************************************
function WriteAMOUpdateableGridHTML (byval stab,byval rsModule, byVal sDivisionIDs, byVal sGroupIDs, byVal sHideColumns)
	dim strClass, sPD, sChange, strDisplay, strTitle, intHdrColSpan
	dim intChangeMask, intRegionChangeMask, intStatusID, bBaseModule
	dim nChassisID, nModuleID, numDL, nRegionId
	dim CurrentSection, CurrentGroup, HWSWType, bchkRight
	dim bLocked, bDisabled, bPutCheckmark, bCostPermission, bRasPermission, bMAMOUpdate, bStatusUpdate


	if IsODM = 0 then
		intHdrColSpan = 41
	else
		intHdrColSpan = 38	'hide 3 columns
	end if

	'--Hide Columns for AMO Features List: ---
	if (stab="All" & "&nbsp;" & "Options" or stab="WIP" or stab="RAS" & "&nbsp;" & "Review" or stab="Complete" or stab="Discontinue") then	
		'--PBI 17498/Task 21007 - Changes for AMO List tabs - Remove Ignore SCL Deployment Plan column: ---
		if instr("," & sHideColumns & ",", ",28,") = 0 Then
			sHideColumns = sHideColumns & ",28,"
		end if

		'--Per Sruthi, Owned By field isn't used in AMO Feature Properties, Remove Owned By column in AMO Features List
		if instr("," & sHideColumns & ",", ",16,") = 0 Then
			sHideColumns = sHideColumns & ",16,"
		end if

		'--PBI 17506/ Task 21612 - Changes for AMO List - Rules Tab - Per Efren, Rules not used for AMO Features, Remove Rules Description 
		if instr("," & sHideColumns & ",", ",20,") = 0 Then
			sHideColumns = sHideColumns & ",20,"
		end if

		'--Bug 25100 Pulsar Test: Resolve AMO List Errors and Changes - Task 25248 AMO List Changes 
		if instr("," & sHideColumns & ",", ",27,") = 0 Then
			sHideColumns = sHideColumns & ",27,"
		end if
	end if

	Response.Write  "<table id='oTable' name='oTable' align='left' border='1' CELLSPACING='0' CELLPADDING='3' " & _
			"onmouseover='row_onmouseover(event);' onmouseout='row_onmouseout(event);' onclick='javascript:ClickEvent(event);return true;' " & _
			" style='color:blue;'>" & vbCrLf	'text-decoration:underline
    
	'Row 1; Heading
	Response.Write  "<tr class='tblrow-pulsar'>" & vbCrLf
	'4/20/2005 add +/- sign for display/hide localization
	if stab<> "RAS" & "&nbsp;" & "Review" then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'></th>" & vbCrLf
	end if 
	Response.Write  "<th rowspan=2 id='misc' align='center' NAME='misc' class='rotate-text'><p>Select Option</p></th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' NAME='misc' class='rotate-text'><p>Properties</p></th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' width='200' nowrap NAME='misc'>Option (Short Description)*</th>" & vbCrLf
	if instr("," & sHideColumns & ",", ",19,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' width='200' NAME='misc'>Long Description</th>" & vbCrLf
	end if
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Status</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Comment to RAS*</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Comment from RAS</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>HP Part Number*</th>" & vbCrLf
	
	if instr("," & sHideColumns & ",", ",16,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Owned By</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",7,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from PRL</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",26,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from SCM</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",27,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from SCL</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",20,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Rules Description</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",1,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Select Availability (SA)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",4,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>PHweb (General) Availability (GA)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",5,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>End of Manufacturing (EM)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",3,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>End of Sales (ES)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",6,") = 0 Or sHideColumns = "6" Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Global Series Config EOL</th>" & vbCrLf
	end if
	if IsODM = 0 then
		if instr("," & sHideColumns & ",", ",8,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>AMO Cost*</th>" & vbCrLf
		end if
		if instr("," & sHideColumns & ",", ",2,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>AMO Price*</th>" & vbCrLf
		end if
		if instr("," & sHideColumns & ",", ",9,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Actual Cost</th>" & vbCrLf
		end if
	end if
	if instr("," & sHideColumns & ",", ",25,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Replacement*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",10,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Alternative</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",12,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Net Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",11,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Export Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",13,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Air Packed Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",14,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Air Packed Cubic</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",15,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Export Cubic</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",29,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Serial Flag*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",17,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Warranty Code*</th>" & vbCrLf
	end if	
	if instr("," & sHideColumns & ",", ",18,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Country of Manufacture*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",22,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Order Instructions</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",21,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Replacement AV Description</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",23,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Product Line*</th>" & vbCrLf
	end if
	
	'4-29-2008 Added by Dien
	if instr("," & sHideColumns & ",", ",28,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Ignore SCL Deployment Plan</th>" & vbCrLf
	end if
	'end
	
	if instr("," & sHideColumns & ",", ",24,") = 0 Then
		Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' colspan=4>SKU Visibility</th>" & vbCrLf
		Response.Write  "</tr>"
	
		Response.Write  "<tr style='color:#FFFFFF;text-decoration:none' bgcolor='#0096d6'>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>NA</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>EMEA</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>APJ</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>LA</th>" & vbCrLf
		Response.Write  "</tr>"
	else
		Response.Write  "</tr><tr></tr>"
	end if

	'Module Data
	CurrentGroup = 0
	rsModule.MoveFirst
	do while (not rsModule.EOF)
		'write Category Heading
		CurrentSection = rsModule.Fields("FeatureCategoryID").Value
		Response.Write  "<tr class='clsCatHeading'><td id='sd' colspan='" & intHdrColSpan & "' align='left' " & _
			"style='color:black;text-decoration:none'" & _
			"title=""" & rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("Category").Value & """>" & _
			rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("Category").Value & "</td>" & vbCrLf
		Response.Write  "</tr>" & vbCrLf
		do while ((not rsModule.EOF) and (rsModule.Fields("FeatureCategoryID").value = CurrentSection))
			'write each option
			'if collapse of columns desired, all columns that are to be collapsed must have same ID, thus make them all MISC.
			' then make an ID1 with the real code to check.
			nModuleID = rsModule.Fields("FeatureID").Value
			nRegionId = rsModule.Fields("RegionID").Value
			intChangeMask = rsModule.Fields("ChangeMask").value
			intRegionChangeMask = rsModule.Fields("RegionChangeMask").value
			intStatusID = rsModule.Fields("AMOStatusID").value
			if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") or intStatusID = Application("AMO_OBSOLETE")  then
				bLocked = True
			else
				bLocked = False
			end if
			
			if CLng(nRegionId) <> Application("AMO_GLOBAL_REGIONID") then
				bBaseModule = False
			else
				bBaseModule = True
			end if
			
			'Check if the User's Teams' BusinessSegments align with Business Segments selected in filter AND they have AMOList.Edit Permission: --
			bStatusUpdate = isIdBelong(Trim(rsModule.Fields("FeatureDivisionIDs").value), sDivisionIDs) and bAMOUpdate				
			
			if bAMOUpdate then
				bMAMOUpdate = True
				bchkRight = 1
			else
				bMAMOUpdate = False
				bchkRight = 0
			end if	
			
			if intStatusID = Application("AMO_DISABLED") then
				bDisabled = True
			else
				bDisabled = False
			end if
			
			'Check if the User's Team's BusinessSegments align with Business Segments selected in filter AND they have AMOCost.Edit Permission: --
			if isIdBelong(Trim(rsModule.Fields("DivisionIDs").value), sDivisionIDs) and bCostUpdate then 
				bCostPermission = True
			else
				bCostPermission = False
			end if
			
			'Check if the User's Team's BusinessSegments align with Business Segments selected in filter AND they have AMORAS.Edit Permission: --
			if isIdBelong(Trim(rsModule.Fields("DivisionIDs").value), sDivisionIDs) and bRASUpdate then
				bRasPermission = True
			else
				bRasPermission = False
			end if
			
			'Response.Write(bStatusUpdate &"--"& bRasPermission &"--"& Application("AMO_OBSOLETE") &"--"& Application("AMO_DISABLED") &"--"& Application("AMO_RASUPDATE") &"--"& Application("AMO_RASREVIEW")) 
			'Response.End()

			strDisplay = Trim(cstr(ConvertNull(rsModule.Fields("AMOStatus").Value)))
		
			strTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo:" & rsModule.Fields("BluePN").Value &  "#" & Trim(rsModule.Fields("DashCode").Value) & ")")

			'Response.Write  "<tr  mType='" & trim(rsModule.Fields("ModuleType").Value) & "' mid='" & nModuleID & "' sid='" & nRegionID & "' cid='" & CurrentSection & "' align='center'>" & vbCrLf
			
			Response.Write  "<tr "
			'4/20/2005 add +/- sign for display/hide localization
			if stab<> "RAS" & "&nbsp;" & "Review" then
				if rsModule.Fields("Localized").Value = 1  then
					if bBaseModule then
						response.write ""
					else
						response.write " id=AMO" & nModuleID & " name=AMO" & nModuleID & " style='display:none'"
					end if 
				else
					response.write ""
				end if 	
			end if 	
			Response.Write  " mType='" & trim(rsModule.Fields("FeatureType").Value) & "' mid='" & nModuleID & "' sid='" & nRegionID & "' cid='" & CurrentSection & "' align='center'>" & vbCrLf	
			
			'4/20/2005 add +/- sign for display/hide localization
			if stab<> "RAS" & "&nbsp;" & "Review" then
				if rsModule.Fields("Localized").Value = 1 then
					 if bBaseModule then
						response.write "<td style='text-decoration:none'><input id='btnExpand" & nModuleID & "' language='javascript' name='btnExpand" & nModuleID & "' type='button' style='HEIGHT:22px;WIDTH:22px;' value='+' onclick='return btnExpand_onclick(" & nModuleID & ")' ></td>"
					else
						response.write "<td></td>"
					end if 
				else
					response.write "<td style='color:black;text-decoration:none'>&nbsp;</td>"
				end if 	
			end if 	
			
			if bMAMOUpdate and (intStatusID <> Application("AMO_OBSOLETE")) And (intStatusID <> Application("AMO_DISABLED")) And (intStatusID <> Application("AMO_RASREVIEW")) And (intStatusID <> Application("AMO_RASUPDATE"))  then
				Response.Write "<td><input type='checkbox' NAME='chkBlkStatus' id='chkBlkStatus' value='" & nModuleID & "|" & strDisplay & "|" & bchkRight & "|" & nRegionId & "'></td>" & vbCrLf	
			elseif bStatusUpdate and (intStatusID <> Application("AMO_OBSOLETE"))  And (intStatusID <> Application("AMO_DISABLED")) And (intStatusID <> Application("AMO_RASREVIEW")) And (intStatusID <> Application("AMO_RASUPDATE")) then				
				Response.Write "<td><input type='checkbox' NAME='chkBlkStatus' id='chkBlkStatus' value='" & nModuleID & "|" & strDisplay & "|" & bchkRight & "|" & nRegionId & "'</td>" & vbCrLf		
			else
				Response.Write "<td id='sd' style='color:black;text-decoration:none' NAME='sd'>&nbsp;</td>" 
			end if
			
			Call DisplayField( true, "prop", "misc", "", server.htmlencode("<View>"), false, "", true, sHideColumns)
			
			Call DisplayField(bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "o", "misc", "align=left", ConvertNull(rsModule.Fields("ShortDescription").Value), SetHighlight(intChangeMask, "shortdescription"), strTitle, bBaseModule, sHideColumns)
			
			if trim(rsModule.Fields("LongDescription").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE"))  then
					Call DisplayField( true, "dsc", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "longDescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "dsc", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "longDescription"), strTitle, bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("LongDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "dsc", "misc", "", ConvertNull(left(trim(server.htmlEncode(rsModule.Fields("LongDescription").Value)),10) & "..."), SetHighlight(intChangeMask, "longDescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "dsc", "misc", "", ConvertNull(trim(server.htmlEncode(rsModule.Fields("LongDescription").Value))), SetHighlight(intChangeMask, "longDescription"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			if bLocked then
				'put a lock icon
				strDisplay = "<img src=""../library/Images/gifs/lock.gif"" alt=""Option Locked"" width=9 height=11 border=0> " & strDisplay
			end if
			'no need to display menu if RAS Review or RAS Update status already
			if intStatusID = Application("AMO_NEW") then
				Call DisplayField( bBaseModule and (bMAMOUpdate or bStatusUpdate or bCostPermission) and not bLocked and not bDisabled, "asnew", "misc", "", strDisplay, "", strTitle, bBaseModule, sHideColumns)
			elseif intStatusID = Application("AMO_DISABLED") then
				Call DisplayField( bBaseModule and bMAMOUpdate, "asdis", "asdis", "", strDisplay, "", strTitle, bBaseModule, sHideColumns)
			else
				'test: 7/9/07 try to remove AMO_IN_PROCESS in if.
				if (bMAMOUpdate or bStatusUpdate or bCostPermission)  and intStatusID <> Application("AMO_RASREVIEW") and intStatusID <> Application("AMO_RASUPDATE") and intStatusID <> Application("AMO_OBSOLETE") then
					Call DisplayField_Status(bBaseModule and (bMAMOUpdate or bCostPermission or bStatusUpdate) and not bLocked and not bDisabled, "as", "as", "", strDisplay, "", strTitle, nModuleID , bBaseModule)
				else
					Call DisplayField( false, "as", "as", "", strDisplay, "", strTitle, bBaseModule, sHideColumns)
				end if
		
			end if

			if trim(rsModule.Fields("InfoComment").Value) = "" then
				if bStatusUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "ctr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				elseif bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "ctr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "ctr", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				end if
			else
				'should always be able to click the link and let the Comment to RAS page handle security
				if len(trim(rsModule.Fields("InfoComment").Value)) > 10 then
					Call DisplayField( bBaseModule and true , "ctr", "misc", "", ConvertNull(left(trim(rsModule.Fields("InfoComment").Value),10) & "..."), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true , "ctr", "misc", "", ConvertNull(trim(rsModule.Fields("InfoComment").Value)), SetHighlight(intChangeMask, "infocomment"), strTitle , bBaseModule, sHideColumns)
				end if
			end if

			'should always be able to click the link and let the Comment from RAS page handle security
			if trim(rsModule.Fields("RASComment").Value) = "" then
				if bRASUpdate and not bLocked and not bDisabled then
					Call DisplayField( bBaseModule and true, "cfr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "rascomment"), strTitle , bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "cfr", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "rascomment"), strTitle , bBaseModule, sHideColumns)
				end if
			else
				'should always be able to click the link and let the Comment to RAS page handle security
				if len(trim(rsModule.Fields("RASComment").Value)) > 10 then
					Call DisplayField( bBaseModule and true , "cfr", "misc", "", ConvertNull(left(trim(rsModule.Fields("RASComment").Value),10) & "..."), SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true , "cfr", "misc", "", ConvertNull(trim(rsModule.Fields("RASComment").Value)), SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				end if
			end if

			if Trim(rsModule.Fields("DashCode").Value) <> "" then
				Call DisplayField_PartNo( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "bpn", "misc", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), strTitle , bBaseModule)
			else
				Call DisplayField_PartNo( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "bpn", "misc", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), strTitle, bBaseModule)
			end if

			'Call DisplayField( false, "own", "misc", "", rsModule.Fields("GroupName").value, false, strTitle, bBaseModule, sHideColumns)
			
			if rsModule.Fields("MOLHide").value = 1 then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if

			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "molhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)
			
			if rsModule.Fields("SCMHide").value = 1 then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if

			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "scmhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)
			
			'if rsModule.Fields("SCLHide").value = 1 then
				'bPutCheckmark = True
			'else
				bPutCheckmark = False
			'end if
			
			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "sclhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)
			
			if trim(rsModule.Fields("RuleDescription").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled then
					Call DisplayField( true, "rdsc", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, server.htmlEncode("ruleDescription")), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "rdsc", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "ruleDescription"), strTitle, bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("RuleDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "rdsc", "misc", "", ConvertNull(server.htmlencode(left(trim(rsModule.Fields("RuleDescription").Value),10)) & "..."), SetHighlight(intChangeMask, server.htmlEncode("RuleDescription")), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "rdsc", "misc", "", ConvertNull(server.htmlencode(trim(rsModule.Fields("RuleDescription").Value))), SetHighlight(intChangeMask, server.htmlEncode("RuleDescription")), strTitle, bBaseModule, sHideColumns)
				end if
			end if
			
			'highlight Blind Date"
			if bMAMOUpdate and not bLocked and not bDisabled then
				Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "cpl", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetLocalizationHighlight(intRegionChangeMask, "cplblinddate"), strTitle , true, sHideColumns)
   				'test Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "cpl", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetHighlight(intChangeMask, "cplblinddate"), strTitle , true)
			else
				Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "misc", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetLocalizationHighlight(intRegionChangeMask, "cplblinddate"), strTitle, true, sHideColumns)
				'test Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "misc", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetHighlight(intChangeMask, "cplblinddate"), strTitle, true)
			end if
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "reva" + cstr(nModuleID) + cstr(nRegionID), "reva", "", ConvertNull(rsModule.Fields("BOMRevADate").Value), SetLocalizationHighlight(intRegionChangeMask, "bomrevadate"), strTitle , true, sHideColumns)

			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "rasdisc" + cstr(nModuleID) + cstr(nRegionID), "rasdisc", "", ConvertNull(rsModule.Fields("RASDiscontinueDate").Value), SetLocalizationHighlight(intRegionChangeMask, "rasdiscontinuedate"), strTitle, true, sHideColumns)
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "obd" + cstr(nModuleID) + cstr(nRegionID), "obd", "", ConvertNull(rsModule.Fields("ObsoleteDate").Value), SetLocalizationHighlight(intRegionChangeMask, "ObsoleteDate"), strTitle, true, sHideColumns)
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "gbsd" + cstr(nModuleID) + cstr(nRegionID), "gbsd", "", ConvertNull(rsModule.Fields("GlobalSeriesDate").Value), SetLocalizationHighlight(intRegionChangeMask, "GlobalSeriesDate"), strTitle, bBaseModule, sHideColumns)
			
			if IsODM = 0 then
				Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "cst", "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOCost").Value)), SetHighlight(intChangeMask, "amocost"), strTitle, bBaseModule, sHideColumns )

				if (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled then
					Call DisplayField( bBaseModule and (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "wwp" + cstr(nModuleID), "wwp", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOWWPrice").Value)), SetHighlight(intChangeMask, "amowwprice"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "wwp" + cstr(nModuleID), "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOWWPrice").Value)), SetHighlight(intChangeMask, "amowwprice"), strTitle , bBaseModule, sHideColumns)
				end if

				Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "acst", "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("ActualCost").Value)), SetHighlight(intChangeMask, "actualcost"), strTitle , bBaseModule, sHideColumns)
			end if
			
			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "rep", "misc", "align=left", ConvertNull(rsModule.Fields("Replacement").Value), SetHighlight(intChangeMask, "replacement"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "alt", "misc", "align=left", ConvertNull(rsModule.Fields("Alternative").Value), SetHighlight(intChangeMask, "alternative"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "nw", "misc", "", ConvertNull(rsModule.Fields("NetWeight").Value), SetHighlight(intChangeMask, "netweight"), strTitle, bBaseModule, sHideColumns )

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "ew", "misc", "", ConvertNull(rsModule.Fields("ExportWeight").Value), SetHighlight(intChangeMask, "exportweight"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "apw", "misc", "", ConvertNull(rsModule.Fields("AirPackedWeight").Value), SetHighlight(intChangeMask, "airpackedweight"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "apc", "misc", "", ConvertNull(rsModule.Fields("AirPackedCubic").Value), SetHighlight(intChangeMask, "airpackedcubic"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "ec", "misc", "", ConvertNull(rsModule.Fields("ExportCubic").Value), SetHighlight(intChangeMask, "exportcubic"), strTitle , bBaseModule, sHideColumns)
			
			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "sf", "sf", "", ConvertNull(rsModule.Fields("SerialFlag").Value), SetHighlight(intChangeMask, "WarrantyCode"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "warc", "warc", "", ConvertNull(rsModule.Fields("WarrantyCode").Value), SetHighlight(intChangeMask, "WarrantyCode"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "maco", "maco", "", ConvertNull(rsModule.Fields("ManufactureCountry").Value), SetHighlight(intChangeMask, "ManufactureCountry"), strTitle , bBaseModule, sHideColumns)
			
			if trim(rsModule.Fields("OrderInstruction").Value) = "" then
				if bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "ord", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "ord", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule, sHideColumns )
				end if
			else
				if len(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value))) > 10 then
					Call DisplayField( bBaseModule and true, "ord", "misc", "", ConvertNull(left(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value)),10) & "..."), SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "ord", "misc", "", ConvertNull(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value))), SetHighlight(intChangeMask, "orderinstruction"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			if trim(rsModule.Fields("ReplacementAVDescription").Value) = "" then
				if bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "reldes", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "replacementavdescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "reldes", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "replacementavdescription"), strTitle , bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("ReplacementAVDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "reldes", "misc", "", ConvertNull(left(trim(rsModule.Fields("ReplacementAVDescription").Value),10) & "..."), SetHighlight(intChangeMask, "replacementavdescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "reldes", "misc", "", ConvertNull(trim(rsModule.Fields("ReplacementAVDescription").Value)), SetHighlight(intChangeMask, "replacementavdescription"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			Call DisplayField( bBaseModule and bMAMOUpdate and not bDisabled, "prodl", "prodl", "", ConvertNull(trim(rsModule.Fields("ProductLine").Value)), SetHighlight(intChangeMask, "productline"), strTitle , bBaseModule, sHideColumns)
			
			'4-29-2008 Added by Dien intStatusID
			if rsModule.Fields("NoSCLDeploy").value then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if

			Call DisplayCheckmark_IDP( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "idp", "" , intStatusID, bBaseModule, nModuleID, sHideColumns)
			'end
			
			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vna" & nModuleID, "vna", "", ConvertNull(rsModule.Fields("Visibility_NA").Value), SetHighlight(intChangeMask, "Visibility_NA"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vem" & nModuleID, "vem", "", ConvertNull(rsModule.Fields("Visibility_EM").Value), SetHighlight(intChangeMask, "Visibility_EM"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vap" & nModuleID, "vap", "", ConvertNull(rsModule.Fields("Visibility_AP").Value), SetHighlight(intChangeMask, "Visibility_AP"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vla" & nModuleID, "vla", "", ConvertNull(rsModule.Fields("Visibility_LA").Value), SetHighlight(intChangeMask, "Visibility_LA"), strTitle , bBaseModule, sHideColumns)

			Response.Write  "</tr>"  & vbCrLf
			rsModule.MoveNext
			if rsModule.EOF then exit do
		loop
		if rsModule.EOF then exit do
	loop
	Response.Write  "</table>" & vbCrLf
end function

'*************************************************************************************
'* Sub/Function	: WriteRASUpdateableGridHTML
'* Description	: Write the module data for the Options tab when in RAS mode
'* Modified		: 07/13/16 - Harris, Valerie - Update sub to list AMO Features
'* Note			: In Pulsar, the Global Region ID is 334; In IRS, the Global Region ID is 43. 
'*				 Only AMO Features and Modules with Global Region IDs will editable fields for most columns.
'*				 Only the date fields are editable for AMO Features with Localized Region
'*************************************************************************************
function WriteRASUpdateableGridHTML (byval stab, byval rsModule, byVal sDivisionIDs, ByVal sGroupIDs, ByVal sHideColumns)
	dim strClass, sPD, sChange, strDisplay, strTitle, strRASTitle
	dim intChangeMask, intRegionChangeMask, intStatusID, intHdrColSpan
	dim nChassisID, nModuleID, numDL, nRegionId, bBaseModule
	dim CurrentSection, CurrentGroup, HWSWType, bchkRight
	dim bLocked, bDisabled, bPutCheckmark, bCostPermission, bRasPermission, bMAMOUpdate, bStatusUpdate
	
	if IsODM = 0 then
		intHdrColSpan = 41
	else
		intHdrColSpan = 38	'hide 3 columns
	end if


	if (stab="All" & "&nbsp;" & "Options" or stab="WIP" or stab="RAS" & "&nbsp;" & "Review" or stab="Complete" or stab="Discontinue") then	
		'--PBI 17498/Task 21007 - Changes for AMO List tabs - Remove Ignore SCL Deployment Plan column: ---
		if instr("," & sHideColumns & ",", ",28,") = 0 Then
			sHideColumns = sHideColumns & "28,"
		end if

		'--Per Sruthi, Owned By field isn't used in AMO Feature Properties, Remove Owned By column in AMO Features List
		if instr("," & sHideColumns & ",", ",16,") = 0 Then
			sHideColumns = sHideColumns & "16,"
		end if

		'--PBI 17506/ Task 21612 - Changes for AMO List - Rules Tab - Per Efren, Rules not used for AMO Features, Remove Rules Description 
		if instr("," & sHideColumns & ",", ",20,") = 0 Then
			sHideColumns = sHideColumns & "20,"
		end if

		'--Bug 25100 Pulsar Test: Resolve AMO List Errors and Changes - Task 25248 AMO List Changes 
		if instr("," & sHideColumns & ",", ",27,") = 0 Then
			sHideColumns = sHideColumns & ",27,"
		end if
	end if
	
	Response.Write  "<table id='oTable' name='oTable' align='left' border='1' CELLSPACING='0' CELLPADDING='3' " & _
			"onmouseover='row_onmouseover(event);' onmouseout='row_onmouseout(event);' onclick='javascript:ClickEvent(event);return true;' " & _
			" style='color:blue;'>" & vbCrLf	'text-decoration:underline
    
	'Row 1; Heading
	Response.Write  "<tr class='tblrow-pulsar'>" & vbCrLf
	'4/20/2005 add +/- sign for display/hide localization
	if stab<> "RAS" & "&nbsp;" & "Review" then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'></th>" & vbCrLf
	end if 
	Response.Write  "<th rowspan=2 id='misc' align='center' NAME='misc' class='rotate-text'><p>Select Option</p></th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' NAME='misc' class='rotate-text'><p>Properties</p></th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' width='200' nowrap NAME='misc'>Option (Short Description)*</th>" & vbCrLf
	if instr("," & sHideColumns & ",", ",19,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Long Description</th>" & vbCrLf
	end if
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Status</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Comment to RAS*</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Comment from RAS</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>HP Part Number*</th>" & vbCrLf
	
	if instr("," & sHideColumns & ",", ",16,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Owned By</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",7,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from PRL</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",26,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from SCM</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",27,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Hide from SCL</th>" & vbCrLf
	end if
	
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>RAS Complete</th>" & vbCrLf
	Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>GPSy Complete</th>" & vbCrLf
	
	if instr("," & sHideColumns & ",", ",20,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Rules Description</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",1,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Select Availability (SA)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",4,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>PHweb (General) Availability (GA)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",5,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>End of Manufacturing (EM)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",3,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>End of Sales (ES)*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",6,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Global Series Config EOL</th>" & vbCrLf
	end if
	if IsODM = 0 then
		if instr("," & sHideColumns & ",", ",8,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>AMO Cost*</th>" & vbCrLf
		end if
		if instr("," & sHideColumns & ",", ",2,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>AMO Price*</th>" & vbCrLf
		end if
		if instr("," & sHideColumns & ",", ",9,") = 0 Then
			Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Actual Cost</th>" & vbCrLf
		end if
	end if
	if instr("," & sHideColumns & ",", ",25,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Replacement*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",10,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Alternative</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",12,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Net Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",11,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Export Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",13,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Air Packed Weight</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",14,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Air Packed Cubic</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",15,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Export Cubic</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",29,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Serial Flag*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",17,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Warranty Code*</th>" & vbCrLf
	end if	
	if instr("," & sHideColumns & ",", ",18,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Country of Manufacture*</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",22,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Order Instructions</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",21,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Replacement AV Description</th>" & vbCrLf
	end if
	if instr("," & sHideColumns & ",", ",23,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Product Line*</th>" & vbCrLf
	end if
	
	'4-29-2008 Added by Dien
	if instr("," & sHideColumns & ",", ",28,") = 0 Then
		Response.Write  "<th rowspan=2 id='misc' align='center' valign='bottom' NAME='misc'>Ignore SCL Deployment Plan</th>" & vbCrLf
	end if
	'end
	
	if instr("," & sHideColumns & ",", ",24,") = 0 Then
		Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' colspan=4>SKU Visibility</th>" & vbCrLf
		Response.Write  "</tr>"
	
		Response.Write  "<tr style='color:#FFFFFF;text-decoration:none' bgcolor='#0096d6'>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>NA</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>EMEA</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>APJ</th>" & vbCrLf
		Response.Write  "<th  id='misc' align='bottom'  valign='bottom' NAME='misc'>LA</th>" & vbCrLf
		Response.Write  "</tr>"
	else
		Response.Write  "</tr><tr></tr>"
	end if

	'Module Data
	CurrentGroup = 0
	rsModule.MoveFirst
	do while (not rsModule.EOF)
		'write Category Heading
		CurrentSection = rsModule.Fields("FeatureCategoryID").Value
		Response.Write  "<tr class='clsCatHeading'><td id='sd' colspan='" & intHdrColSpan & "' align='left' " & _
			"style='color:black;text-decoration:none'" & _
			"title=""" & rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("CategoryDesc").Value & """>" & _
			rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("CategoryDesc").Value & "</td>" & vbCrLf
		Response.Write  "</tr>" & vbCrLf
		do while ((not rsModule.EOF) and (rsModule.Fields("FeatureCategoryID").value = CurrentSection))
			'write each module
			nModuleID = rsModule.Fields("FeatureID").Value
			nRegionId = rsModule.Fields("RegionID").Value
			intChangeMask = rsModule.Fields("ChangeMask").value
			intRegionChangeMask = rsModule.Fields("RegionChangeMask").value
			intStatusID = rsModule.Fields("AMOStatusID").value
			
			if CLng(nRegionId) <> Application("AMO_GLOBAL_REGIONID") then
				bBaseModule = False
			else
				bBaseModule = True
			end if
			
			if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") or intStatusID = Application("AMO_OBSOLETE") then
				bLocked = True
			else
				bLocked = False
			end if
			
			'Check if the User's Teams' BusinessSegments align with Business Segments selected in filter AND they have AMOList.Edit Permission: --
			bStatusUpdate = isIdBelong(Trim(rsModule.Fields("FeatureDivisionIDs").value), sDivisionIDs) and bAMOUpdate	
			
			if bAMOUpdate then
				bMAMOUpdate = True
				bchkRight = 1
			else
				bMAMOUpdate = False
				bchkRight = 0
			end if	
			
			if intStatusID = Application("AMO_DISABLED") then
				bDisabled = True
			else
				bDisabled = False
			end if
			
			'Check if the User's Teams' BusinessSegments align with Business Segments selected in filter AND they have AMOCost.Edit Permission: --
			if isIdBelong(Trim(rsModule.Fields("DivisionIDs").value), sDivisionIDs) and bCostUpdate then 
				bCostPermission = True
			else
				bCostPermission = False
			end if
			
			'Check if the User's Teams' BusinessSegments align with Business Segments selected in filter AND they have AMORAS.Edit Permission: --
			if  isIdBelong(Trim(rsModule.Fields("DivisionIDs").value), sDivisionIDs) and bRASUpdate then
				bRasPermission = True
			else
				bRasPermission = False
			end if
			
			'Response.Write(bStatusUpdate &"--"& bRasPermission &"--"& Application("AMO_OBSOLETE") &"--"& Application("AMO_DISABLED") &"--"& Application("AMO_RASUPDATE") &"--"& Application("AMO_RASREVIEW")) 
			'Response.End()

			strDisplay = Trim(cstr(ConvertNull(rsModule.Fields("AMOStatus").Value)))
		
			strTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo:" & rsModule.Fields("BluePN").Value &  "#" & Trim(rsModule.Fields("DashCode").Value) & ")")
			strRASTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo:" & rsModule.Fields("BluePN").Value &  "#" & Trim(rsModule.Fields("DashCode").Value) & ")")

			Response.Write  "<tr "
			'4/20/2005 add +/- sign for display/hide localization
			if stab<> "RAS" & "&nbsp;" & "Review" then
				if rsModule.Fields("Localized").Value = 1  then
					if bBaseModule then
						response.write ""
					else
						response.write " id=AMO" & nModuleID & " name=AMO" & nModuleID & " style='display:none'"
					end if 
				else
					response.write ""
				end if 	
			end if 	
			Response.Write  " mType='" & trim(rsModule.Fields("FeatureType").Value) & "' mid='" & nModuleID & "' sid='" & nRegionID & "' cid='" & CurrentSection & "' align='center'>" & vbCrLf	
			
			'4/20/2005 add +/- sign for display/hide localization
			if stab<> "RAS" & "&nbsp;" & "Review" then
				if rsModule.Fields("Localized").Value = 1 then
					 if bBaseModule then
						response.write "<td style='text-decoration:none'><input id='btnExpand" & nModuleID & "' language='javascript' name='btnExpand" & nModuleID & "' type='button' style='HEIGHT:22px;WIDTH:22px;' value='+' onclick='return btnExpand_onclick(" & nModuleID & ")' ></td>"
					else
						response.write "<td></td>"
					end if 
				else
					response.write "<td style='color:black;text-decoration:none'>&nbsp;</td>"
				end if 	
			end if 	
		
			if bMAMOUpdate and bRasPermission and (intStatusID <> Application("AMO_OBSOLETE")) And (intStatusID <> Application("AMO_DISABLED")) then				
				Response.Write "<td><input type='checkbox' NAME='chkBlkStatus' id='chkBlkStatus' value='" & nModuleID & "|" & strDisplay &  "|" & bchkRight & "|" & nRegionId & "'</td>" & vbCrLf
			elseif bStatusUpdate and (intStatusID <> Application("AMO_OBSOLETE"))  And (intStatusID <> Application("AMO_DISABLED")) And (intStatusID <> Application("AMO_RASUPDATE")) And (intStatusID <> Application("AMO_RASREVIEW"))  then				
				Response.Write "<td><input type='checkbox' NAME='chkBlkStatus' id='chkBlkStatus' value='" & nModuleID & "|" & strDisplay &  "|" & bchkRight & "|" & nRegionId & "'</td>" & vbCrLf		
			elseif bRasPermission and (intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE")) then
				Response.Write "<td><input type='checkbox' NAME='chkBlkStatus' id='chkBlkStatus' value='" & nModuleID & "|" & strDisplay &  "|" & bchkRight & "|" & nRegionId & "'</td>" & vbCrLf	
			else
				Response.Write "<td id='sd' style='color:black;text-decoration:none' NAME='sd'>&nbsp;</td>" 
			end if
			
			Call DisplayField( true, "prop", "misc", "", server.htmlencode("<View>"), false, "", true, sHideColumns)
			
			Call DisplayField(bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "o", "misc", "align=left", ConvertNull(rsModule.Fields("ShortDescription").Value), SetHighlight(intChangeMask, "shortdescription"), strTitle, bBaseModule, sHideColumns)
			
			if trim(rsModule.Fields("LongDescription").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE"))  then
					Call DisplayField( true, "dsc", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "longDescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "dsc", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "longDescription"), strTitle , bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("LongDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "dsc", "misc", "", ConvertNull(left(trim(server.htmlEncode(rsModule.Fields("LongDescription").Value)),10) & "..."), SetHighlight(intChangeMask, "longDescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "dsc", "misc", "", ConvertNull(trim(server.htmlEncode(rsModule.Fields("LongDescription").Value))), SetHighlight(intChangeMask, "longDescription"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			if bLocked then
				'put a lock icon
				strDisplay = "<img src=""../library/Images/gifs/lock.gif"" alt=""Option Locked"" width=9 height=11 border=0> " & strDisplay
			end if
			
			if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") then			 
				Call DisplayField(bBaseModule and bRasPermission, "asras", "misc", "", strDisplay, "", strTitle , bBaseModule, sHideColumns)			  
			elseif intStatusID = Application("AMO_OBSOLETE") then
			    Call DisplayField_Status( (bBaseModule and bMAMOUpdate) and not bLocked and not bDisabled, "as", "as", "", strDisplay, "class='clsAMO_ChangedCell'", strTitle, nModuleID , bBaseModule, sHideColumns)
			elseif intStatusID = Application("AMO_DISABLED") then
				Call DisplayField( bBaseModule and bMAMOUpdate, "asdis", "asdis", "", strDisplay, "", strTitle, bBaseModule, sHideColumns)
			else
				if intStatusID = Application("AMO_NEW") then
					Call DisplayField( (bBaseModule and (bMAMOUpdate or bStatusUpdate or bCostPermission)) and not bLocked and not bDisabled, "asnew", "misc", "", strDisplay, "", strTitle , bBaseModule, sHideColumns)
				else
				 if bMAMOUpdate or bStatusUpdate or bCostPermission then
					Call DisplayField_Status( (bBaseModule and (bMAMOUpdate or bStatusUpdate or bCostPermission)) and not bLocked and not bDisabled, "as", "as", "", strDisplay, "", strTitle, nModuleID, bBaseModule)
				 else
					Call DisplayField( false, "as", "as", "", strDisplay, "", strTitle, bBaseModule)
				 end if
				end if
			end if

			if trim(rsModule.Fields("InfoComment").Value) = "" then
				if bBaseModule and bStatusUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "ctr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				elseif bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "ctr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "ctr", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				end if
			else
				'should always be able to click the link and let the Comment to RAS page handle security
				if len(trim(rsModule.Fields("InfoComment").Value)) > 10 then
					Call DisplayField( bBaseModule and true , "ctr", "misc", "", ConvertNull(left(trim(rsModule.Fields("InfoComment").Value),10) & "..."), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true , "ctr", "misc", "", ConvertNull(trim(rsModule.Fields("InfoComment").Value)), SetHighlight(intChangeMask, "infocomment"), strTitle, bBaseModule, sHideColumns)
				end if
			end if

			'allow adding a comment to a RAS Review or RAS Update option
			if trim(rsModule.Fields("RASComment").Value) = "" then
				if bRasPermission and not bDisabled  and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( bBaseModule and true, "cfr", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "cfr", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				end if
			else
				'should always be able to click the link and let the Comment to RAS page handle security
				if len(trim(rsModule.Fields("RASComment").Value)) > 10 then
					Call DisplayField( bBaseModule and true , "cfr", "misc", "", ConvertNull(left(trim(rsModule.Fields("RASComment").Value),10) & "..."), SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true , "cfr", "misc", "", ConvertNull(trim(rsModule.Fields("RASComment").Value)), SetHighlight(intChangeMask, "rascomment"), strTitle, bBaseModule, sHideColumns)
				end if
			end if

			if Trim(rsModule.Fields("DashCode").Value) <> "" then
				Call DisplayField_PartNo( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "bpn", "misc", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), strTitle, bBaseModule)
			else
				Call DisplayField_PartNo( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "bpn", "misc", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), strTitle, bBaseModule)
			end if

			'Call DisplayField( false, "own", "misc", "", rsModule.Fields("GroupName").value, false, strTitle , bBaseModule)
			
			if rsModule.Fields("MOLHide").value = 1 then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if

			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "molhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)
			
			if rsModule.Fields("SCMHide").value = 1 then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if
			
			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "scmhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)
			
			'if rsModule.Fields("SCLHide").value = 1 then
				'bPutCheckmark = True
			'else
				bPutCheckmark = False
			'end if
			
			Call DisplayCheckmark_MOLHide( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "sclhide", "" , rsModule.Fields("FeatureType").value, bBaseModule, nModuleID, sHideColumns)

			if bBaseModule then
			
			if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") then
				'Option is in RAS Review or RAS Update
				
					if bRasPermission  then
						if rsModule.Fields("RAS").Value = 1 then
							Response.Write "<td id='ras' title='RAS Complete : " & strRASTitle & "'><input type='checkbox' NAME='chkRAS" & nModuleID & "' ID='chkRAS" & nModuleID & "' value='1' checked "
							response.write "onClick=""javascript:ChangeRAS_GPSy(1," & nModuleID & "); return true;""></td>" & vbCrLf
						else
							Response.Write "<td id='ras' title='RAS Complete : " & strRASTitle & "'><input type='checkbox' NAME='chkRAS" & nModuleID & "' ID='chkRAS" & nModuleID & "' value='0' "
							response.write "onClick=""javascript:ChangeRAS_GPSy(1," & nModuleID & "); return true;""></td>" & vbCrLf
						end if
					else
						if rsModule.Fields("RAS").Value = 1 then
							Response.Write "<td id='ras' title='RAS Complete : " & strRASTitle & "'><input disabled type='checkbox' id='chkRAS' NAME='chkRAS' value='1' checked></td>" & vbCrLf
						else
							Response.Write "<td id='ras' title='RAS Complete : " & strRASTitle & "'><input disabled type='checkbox' id='chkRAS' NAME='chkRAS' value='0'></td>" & vbCrLf
						end if
					end if
	
					if bRasPermission  then
						if rsModule.Fields("GPSy").Value = 1 then
							Response.Write  "<td id='gpsy' title='RAS GPSY : " & strRASTitle & "'><input type='checkbox' NAME='chkGPSy" & nModuleID & "' ID='chkGPSy" & nModuleID & "' value='1' checked "
							response.write "onClick=""javascript:ChangeRAS_GPSy(2," & nModuleID & "); return true;""></td>" & vbCrLf
						else
							Response.Write  "<td id='gpsy' title='RAS GPSY : " & strRASTitle & "'><input type='checkbox' NAME='chkGPSy" & nModuleID & "' ID='chkGPSy" & nModuleID & "' value='0' "
							response.write "onClick=""javascript:ChangeRAS_GPSy(2," & nModuleID & "); return true;""></td>" & vbCrLf
						end if
					else
						if rsModule.Fields("GPSy").Value = 1 then
							Response.Write  "<td id='gpsy' title='RAS GPSY : " & strRASTitle & "'><input disabled type='checkbox' id='chkGPSy' NAME='chkGPSy' value='1' checked></td>" & vbCrLf
						else
							Response.Write  "<td id='gpsy' title='RAS GPSY : " & strRASTitle & "'><input disabled type='checkbox' id='chkGPSy' NAME='chkGPSy' value='0'></td>" & vbCrLf
						end if
					end if
				else 
					'Option is NOT in RAS Review or RAS Update thus don't put the checkboxes
					Response.Write "<td id='sd' style='text-decoration:none'>&nbsp;</td>" & vbCrLf
					Response.Write "<td id='sd' style='text-decoration:none'>&nbsp;</td>" & vbCrLf
				end if
			else
				Response.Write "<td bgcolor='lightgrey' id='sd' style='text-decoration:none'>&nbsp;</td>" & vbCrLf
				Response.Write "<td bgcolor='lightgrey' id='sd' style='text-decoration:none'>&nbsp;</td>" & vbCrLf
			end if 
			
			if trim(rsModule.Fields("RuleDescription").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled then
					Call DisplayField( true, "rdsc", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, server.htmlEncode("ruleDescription")), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "rdsc", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "ruleDescription"), strTitle , bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("RuleDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "rdsc", "misc", "", ConvertNull(server.htmlencode(left(trim(rsModule.Fields("RuleDescription").Value),10)) & "..."), SetHighlight(intChangeMask, server.htmlEncode("RuleDescription")), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "rdsc", "misc", "", ConvertNull(server.htmlencode(trim(rsModule.Fields("RuleDescription").Value))), SetHighlight(intChangeMask, server.htmlEncode("RuleDescription")), strTitle, bBaseModule, sHideColumns)
				end if
			end if
			
			if bMAMOUpdate and not bLocked and not bDisabled then
				Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "cpl", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetLocalizationHighlight(intRegionChangeMask, "cplblinddate"), strTitle, true, sHideColumns)
			else
				Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "cpl" + cstr(nModuleID) + cstr(nRegionID), "misc", "", ConvertNull(rsModule.Fields("CPLBlindDate").Value), SetLocalizationHighlight(intRegionChangeMask, "cplblinddate"), strTitle , true, sHideColumns)
			end if
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "reva" + cstr(nModuleID) + cstr(nRegionID), "reva", "", ConvertNull(rsModule.Fields("BOMRevADate").Value), SetLocalizationHighlight(intRegionChangeMask, "bomrevadate"), strTitle , true, sHideColumns)

			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "rasdisc" + cstr(nModuleID) + cstr(nRegionID), "rasdisc", "", ConvertNull(rsModule.Fields("RASDiscontinueDate").Value), SetLocalizationHighlight(intRegionChangeMask, "rasdiscontinuedate"), strTitle, true, sHideColumns)
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "obd" + cstr(nModuleID) + cstr(nRegionID), "obd", "", ConvertNull(rsModule.Fields("ObsoleteDate").Value), SetLocalizationHighlight(intRegionChangeMask, "ObsoleteDate"), strTitle , true, sHideColumns)
			
			Call DisplayField( bMAMOUpdate and not bLocked and not bDisabled, "gbsd" + cstr(nModuleID) + cstr(nRegionID), "gbsd", "", ConvertNull(rsModule.Fields("GlobalSeriesDate").Value), SetLocalizationHighlight(intRegionChangeMask, "GlobalSeriesDate"), strTitle, bBaseModule, sHideColumns)
			
			if IsODM = 0 then
				Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "cst", "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOCost").Value)), SetHighlight(intChangeMask, "amocost"), strTitle, bBaseModule, sHideColumns)

				if (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled then
					Call DisplayField( bBaseModule and (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "wwp" + cstr(nModuleID), "wwp", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOWWPrice").Value)), SetHighlight(intChangeMask, "amowwprice"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "wwp" + cstr(nModuleID), "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("AMOWWPrice").Value)), SetHighlight(intChangeMask, "amowwprice"), strTitle , bBaseModule, sHideColumns)
				end if

				Call DisplayField(bBaseModule and  (bMAMOUpdate or bCostPermission) and not bLocked and not bDisabled, "acst", "misc", "align=right", ConvertNull(TwoPlaces(rsModule.Fields("ActualCost").Value)), SetHighlight(intChangeMask, "actualcost"), strTitle , bBaseModule, sHideColumns)
			end if
			
			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "rep", "misc", "align=left", ConvertNull(rsModule.Fields("Replacement").Value), SetHighlight(intChangeMask, "replacement"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "alt", "misc", "align=left", ConvertNull(rsModule.Fields("Alternative").Value), SetHighlight(intChangeMask, "alternative"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "nw", "misc", "", ConvertNull(rsModule.Fields("NetWeight").Value), SetHighlight(intChangeMask, "netweight"), strTitle, bBaseModule, sHideColumns )

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "ew", "misc", "", ConvertNull(rsModule.Fields("ExportWeight").Value), SetHighlight(intChangeMask, "exportweight"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "apw", "misc", "", ConvertNull(rsModule.Fields("AirPackedWeight").Value), SetHighlight(intChangeMask, "airpackedweight"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "apc", "misc", "", ConvertNull(rsModule.Fields("AirPackedCubic").Value), SetHighlight(intChangeMask, "airpackedcubic"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "ec", "misc", "", ConvertNull(rsModule.Fields("ExportCubic").Value), SetHighlight(intChangeMask, "exportcubic"), strTitle , bBaseModule, sHideColumns)
			
			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "sf", "sf", "", ConvertNull(rsModule.Fields("SerialFlag").Value), SetHighlight(intChangeMask, "WarrantyCode"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "warc", "warc", "", ConvertNull(rsModule.Fields("WarrantyCode").Value), SetHighlight(intChangeMask, "WarrantyCode"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField(bBaseModule and  bMAMOUpdate and not bLocked and not bDisabled, "maco", "maco", "", ConvertNull(rsModule.Fields("ManufactureCountry").Value), SetHighlight(intChangeMask, "ManufactureCountry"), strTitle , bBaseModule, sHideColumns)
			
			if trim(rsModule.Fields("OrderInstruction").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( true, "ord", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "ord", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule , sHideColumns)
				end if
			else
				if len(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value))) > 10 then
					Call DisplayField( bBaseModule and true, "ord", "misc", "", ConvertNull(left(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value)),10) & "..."), SetHighlight(intChangeMask, "orderinstruction"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "ord", "misc", "", ConvertNull(trim(server.htmlEncode(rsModule.Fields("OrderInstruction").Value))), SetHighlight(intChangeMask, "orderinstruction"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			if trim(rsModule.Fields("ReplacementAVDescription").Value) = "" then
				if bBaseModule and bMAMOUpdate and not bLocked and not bDisabled and (intStatusID <> Application("AMO_OBSOLETE")) then
					Call DisplayField( true, "reldes", "misc", "", server.htmlencode("<Add>"), SetHighlight(intChangeMask, "replacementavdescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( false, "reldes", "misc", "", "&nbsp;", SetHighlight(intChangeMask, "replacementavdescription"), strTitle , bBaseModule, sHideColumns)
				end if
			else
				if len(trim(rsModule.Fields("ReplacementAVDescription").Value)) > 10 then
					Call DisplayField( bBaseModule and true, "reldes", "misc", "", ConvertNull(left(trim(rsModule.Fields("ReplacementAVDescription").Value),10) & "..."), SetHighlight(intChangeMask, "replacementavdescription"), strTitle, bBaseModule, sHideColumns)
				else
					Call DisplayField( bBaseModule and true, "reldes", "misc", "", ConvertNull(trim(rsModule.Fields("ReplacementAVDescription").Value)), SetHighlight(intChangeMask, "replacementavdescription"), strTitle , bBaseModule, sHideColumns)
				end if
			end if
			
			Call DisplayField( bBaseModule and bMAMOUpdate and not bDisabled, "prodl", "prodl", "", ConvertNull(trim(rsModule.Fields("ProductLine").Value)), SetHighlight(intChangeMask, "productline"), strTitle , bBaseModule, sHideColumns)
			
			'4-29-2008 Added by Dien intStatusID
			if rsModule.Fields("NoSCLDeploy").value then
				bPutCheckmark = True
			else
				bPutCheckmark = False
			end if

			Call DisplayCheckmark_IDP( bPutCheckmark, bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, strTitle, "", "idp", "" , intStatusID, bBaseModule, nModuleID, sHideColumns)
			'end
			
			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vna" & nModuleID, "vna", "", ConvertNull(rsModule.Fields("Visibility_NA").Value), SetHighlight(intChangeMask, "Visibility_NA"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vem" & nModuleID, "vem", "", ConvertNull(rsModule.Fields("Visibility_EM").Value), SetHighlight(intChangeMask, "Visibility_EM"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vap" & nModuleID, "vap", "", ConvertNull(rsModule.Fields("Visibility_AP").Value), SetHighlight(intChangeMask, "Visibility_AP"), strTitle , bBaseModule, sHideColumns)

			Call DisplayField( bBaseModule and bMAMOUpdate and not bLocked and not bDisabled, "vla" & nModuleID, "vla", "", ConvertNull(rsModule.Fields("Visibility_LA").Value), SetHighlight(intChangeMask, "Visibility_LA"), strTitle , bBaseModule, sHideColumns)

			Response.Write  "</tr>"  & vbCrLf
			rsModule.MoveNext
			if rsModule.EOF then exit do
		loop
		if rsModule.EOF then exit do
	loop
	Response.Write  "</table>" & vbCrLf
end function
 

'*************************************************************************************
'* Sub/Function	: WritePlatformUpdateableGridHTML
'* Description	: Write the module data for the Platforms tab
'* Modified		: 07/13/16 - Harris, Valerie - Update sub to list AMO Features and PBI 17505/ Task 21610
'* Note			: In Pulsar, the Global Region ID is 334; In IRS, the Global Region ID is 43. 
'*				 Only AMO Features and Modules with Global Region IDs will editable fields for most columns.
'*				 Only the date fields are editable for AMO Features with Localized Region
'*************************************************************************************
function WritePlatformUpdateableGridHTML (byval rsModule, byVal strBusIds, byVal sDivisionIDs, byVal sGroupIDs)
	dim strClass, sPD, sChange, sErr, strModuleIDs, strPlatform, strMarketingName, strTitle, strDisplay, strPlatformID, iFeatureID
	dim strExtra, strHighlight, strDivisionID
	dim intPlatformID, intChangeMask, intStatusID, strAvailableStatus, intPlatformLastRecord, intPlatform
	dim nChassisID, nModuleID, numDL
	dim oSvr, oErr, oRsPlatforms, oRsBusSeg, oRsCompatibility
	dim bPutCheckmark, bDisplayed, bLocked, bDisabled, bInDivision
	dim CurrentSection, CurrentGroup, HWSWType
	dim arrDivs, strAvailableDate, bPlatPermission, bPermission, intCount, intGroupCount, sCompatibilityId, strProductName, strPlatformFullName
	Dim strGroupName, strDisplayed
	
	strAvailableStatus = ""
	strAvailableDate = ""

	const NO_CHANGE = 0
	const CHANGE_MADE = 1
	const CHECKMARK_REMOVED = 2
	
	'set oErr = GetMOLCategory(oRsBusSeg, 28)
	set oRsBusSeg = GetMOLCategory(34)	
	if (oRsBusSeg is Nothing) then
		Response.Write("Recordset error: oRsBusSeg")
		Response.End()
	end if
	
	'combine all moduleids and pass it to the platforms function
	rsModule.Filter = "RegionId = "&Application("AMO_GLOBAL_REGIONID")&""
	if (rsModule.RecordCount = 0) then
		Response.Write("There are no AMO Features with an active Region. Try updating filter.")
		set rsModule = nothing
		exit function
	else
		strModuleIDs = CombineModuleIDs(rsModule)
	end if

	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	set oRsPlatforms = oSvr.AMOPlatforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & strBusIds, 1, 0)
	'Response.Write(strModuleIDs & "|" & strBusIds)
	'Response.End()

	if (oRsPlatforms is nothing) then
		Response.Write("Empty Recordset: oRsPlatforms")
		set oRsPlatforms = nothing
		exit function
	else
		intPlatformLastRecord = oRsPlatforms.RecordCount
	end if
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	set oRsCompatibility = oSvr.AMOCompatibility_Search(Application("REPOSITORY"), strModuleIDs, 1)
	if (oRsCompatibility is nothing) then
		Response.Write("Empty Recordset: oRsCompatibility")
		set oRsCompatibility = nothing
		exit function
	end if

	if oRsPlatforms.RecordCount = 0 then
		response.write "No platforms have been assigned to the displayed options." & vbCrLf
	end if

	Response.Write  "<table id='oTable' name='oTable' align='left' border='1' CELLSPACING='0' CELLPADDING='3' " & _
			"onmouseover='row_onmouseover(event);' onmouseout='row_onmouseout(event);' onclick='javascript:ClickEvent(event);return true;' " & _
			" style='color:blue;'>" & vbCrLf	'text-decoration:underline

	'Row 1; Heading
	Response.Write  "<tr class='tblrow-pulsar'>" & vbCrLf
	Response.Write  "<th id='misc' align='center' NAME='misc' rowspan='2' class='rotate-text'><p>Properties</p></th>" & vbCrLf
	Response.Write  "<th id='misc' align='center' valign='bottom' NAME='misc' style='white-space: nowrap !important;' rowspan='2'>Option (Short Description)</th>" & vbCrLf
	Response.Write  "<th id='misc' align='center' valign='bottom' NAME='as' style='white-space: nowrap !important;' rowspan='2'>Status</th>" & vbCrLf		
	Response.Write  "<th id='misc' align='center' valign='bottom' NAME='misc' style='white-space: nowrap !important;' rowspan='2'>HP Part Number</th>" & vbCrLf	
	Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' style='white-space: nowrap !important;' colspan='" & oRsBusSeg.recordcount & "'>Compatibility</th>" & vbCrLf

	'display product headings
	if oRsPlatforms.RecordCount > 0 then
		oRsPlatforms.MoveFirst
		strProductName = ""
		strPlatformFullName = ""
		intGroupCount = 0
		intCount = 0
		intPlatform = 0

		do while (not oRsPlatforms.EOF)
			if strProductName = "" then
				if oRsPlatforms("ProductName").value = "" Then
					strProductName = ""
				else
					strProductName = oRsPlatforms("ProductName").value
				end if
			end if

			if intCount = oRsPlatforms.RecordCount then
				Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' style='white-space: nowrap !important;' colspan='" & (intGroupCount) & "'><p>" & strProductName & "</p></th>" & vbCrLf 
			end if

			if oRsPlatforms("ProductName").value <> strProductName then
				If strProductName <> "" Then
					Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' style='white-space: nowrap !important;' colspan='" & (intGroupCount) & "'><p>" & strProductName & "</p></th>" & vbCrLf 
					strProductName = oRsPlatforms("ProductName").value
					strPlatformFullName = ""
					intGroupCount = 0
				Else
					Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' style='white-space: nowrap !important;' colspan='" & (intGroupCount) & "'><p>Not Associated to any Product</p></th>" & vbCrLf 
					strProductName = oRsPlatforms("ProductName").value
					strPlatformFullName = ""
					intGroupCount = 0
				End If
			else
				if oRsPlatforms("ProductPlatformName").value <> strPlatformFullName then
					strPlatformFullName = oRsPlatforms("ProductPlatformName").value 
					intGroupCount = intGroupCount + 1
					intPlatform = intPlatform + 1
				end if
			end if

			intCount = intCount + 1
			oRsPlatforms.MoveNext
		loop

		if intPlatformLastRecord = intCount then
			Response.Write  "<th id='misc' align='center' valign='center' NAME='misc' style='white-space: nowrap !important;' colspan='" & (intGroupCount) & "'><p>" & strProductName & "</p></th>" & vbCrLf 
			strProductName = ""
			strPlatformFullName = ""
			intCount = 0
		end if
	end if
	
	Response.Write  "</tr>"
	
	Response.Write "<tr class='tblrow-pulsar'>"
	
	'display compatibility		
	if oRsBusSeg.recordcount > 0 then
		oRsBusSeg.MoveFirst
	end if
	do while (not oRsBusSeg.EOF)
		Response.Write  "<th id='p' align='center' class='rotate-text'><p>" & Trim(oRsBusSeg("SegmentName").value) & "</p></th>" & vbCrLf 
		intCount = intCount + 1 
		oRsBusSeg.MoveNext
	loop

	'display platforms'
	if oRsPlatforms.RecordCount > 0 then
		oRsPlatforms.MoveFirst
		strPlatform = ""
		strMarketingName = ""
		bDisplayed = true
		intCount = 0
		do while (not oRsPlatforms.EOF)
			if oRsPlatforms("ProductPlatformName").value <> strPlatform and strPlatform <> "" then
				Response.Write  "<th id='p' align='center' class='rotate-text'><p>" & Trim(strMarketingName) & "</p></th>" & vbCrLf 
				bDisplayed = true
				intCount = intCount + 1
			end if
			strPlatform = oRsPlatforms("ProductPlatformName").value
			strMarketingName = oRsPlatforms("MarketingName").value
			bDisplayed = false
			oRsPlatforms.MoveNext
		loop
		if not bDisplayed then
			Response.Write  "<th id='p' align='center' class='rotate-text'><p>" &  Trim(strMarketingName) & "</p></th>" & vbCrLf
			intCount = intCount + 1
		end if
	end if
	
	Response.Write  "</tr>"

	'Module Data
	CurrentGroup = 0
	
	rsModule.MoveFirst
	do while (not rsModule.EOF)
		'write Category Heading
		CurrentSection = rsModule.Fields("FeatureCategoryID").Value
		Response.Write  "<tr class='clsCatHeading'><td id='sd' colspan='100' align='left' " & _
			"style='color:black;text-decoration:none'" & _
			"title=""" & rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("Category").Value & """>" & _
			rsModule.Fields("FeatureType").Value & " | " & rsModule.Fields("Category").Value & "</td>" & vbCrLf
		Response.Write  "</tr>" & vbCrLf
		do while ((not rsModule.EOF) and (rsModule.Fields("FeatureCategoryID").value = CurrentSection))
			'write each option
			'Use bgcolor='#FFFF99' for highlighted cells to indicate changes
			nModuleID = rsModule.Fields("FeatureID").Value
			intChangeMask = rsModule.Fields("ChangeMask").value
			intStatusID = rsModule.Fields("AMOStatusID").value
			strDisplayed = ""
		
			'7/9/07, remove 'in process'
			if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") or intStatusID = Application("AMO_OBSOLETE") then
			'TEST if intStatusID = Application("AMO_RASREVIEW") or intStatusID = Application("AMO_RASUPDATE") or intStatusID = Application("AMO_OBSOLETE") or intStatusID = Application("AMO_COMPLETE") or intStatusID = Application("AMO_IN_PROCESS") then
				bLocked = True
			else
				bLocked = False
			end if
			
			'add 2 more OR statement for IN_process and Complete
			'TEST if intStatusID = Application("AMO_DISABLED") OR intStatusID = Application("AMO_IN_PROCESS") OR intStatusID = Application("AMO_COMPLETE")then
			if intStatusID = Application("AMO_DISABLED") then
				bDisabled = True
			else
				bDisabled = False
			end if
			
			'Check if the User's Teams' BusinessSegments align with Business Segments selected in filter AND they have AMOCost.Edit Permission: --
			if isIdBelong(Trim(rsModule.Fields("FeatureDivisionIDs").value), sDivisionIDs) then 
				bPermission = True
			else
				bPermission = False
			end if

			'bPermission = isIdBelong(Trim(sDivisionIDs), Trim(rsModule.Fields("DivisionIDs").value)) 'In Pulsar, Owned By not used for Features: or (instr(sGroupIDs,"," & Trim(rsModule.Fields("GroupID").value) & ",") > 0)

			Response.Write  "<tr mType='" & trim(rsModule.Fields("FeatureType").Value) & "' mid='" & nModuleID & "' cid='" & CurrentSection & "' align='center'>" & vbCrLf

			'Just flip the ID and Name fields around for DisplayFuield so that it will work for the different Platforms tab
			Call DisplayField( true, "misc", "prop", "", server.htmlencode("<View>"), false, "", true, "" )
			'Call DisplayField( true, "prop", "misc", "", server.htmlencode("<View>"), false, "", true, "" )
			
			Call DisplayField( false, "", "o", "", ConvertNull(rsModule.Fields("ShortDescription").Value), SetHighlight(intChangeMask, "shortdescription"), "" , true, "")
			'Call DisplayField( false, "o", "", "", ConvertNull(rsModule.Fields("ShortDescription").Value), SetHighlight(intChangeMask, "shortdescription"), "" , true, "")

			strDisplay = Trim(cstr(ConvertNull(rsModule.Fields("AMOStatus").Value)))
			
			if bLocked then
				'put a lock icon
				strDisplay = "<img src=""../library/Images/gifs/lock.gif"" alt=""Option Locked"" width=9 height=11 border=0> " & strDisplay
			end if
			Call DisplayField( false, "as", "as", "", strDisplay, "", "", true, "")

			Call DisplayField( false, "", "bpn", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), "" , true, "")
			'Call DisplayField( false, "bpn", "", "", ConvertNull(rsModule.Fields("BluePN").Value), SetHighlight(intChangeMask, "bluepartno"), "" , true, "")

			'display compatibility
			if oRsBusSeg.recordcount> 0 then
				oRsBusSeg.MoveFirst
			end if
			do while (not oRsBusSeg.EOF)
				'*NEED CHANGE MASK STATUS IN COMPATIBILITY SEARCH RECORDSET: ----
				strHighlight = ""
				bPutCheckmark = false
				sCompatibilityId = oRsBusSeg.Fields("BusinessSegmentID").Value
				oRsCompatibility.filter = "FeatureID = " & nModuleID & " and DivisionId = " & sCompatibilityId
				strTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo :" & rsModule.Fields("BluePN").Value & ") / " + oRsBusSeg.Fields("SegmentName").Value)
				if oRsCompatibility.recordCount > 0 then
					select case oRsCompatibility("ChangeMask")
							case NO_CHANGE
								bPutCheckmark = True
								strHighlight = ""
							case CHANGE_MADE
								bPutCheckmark = True
								strHighlight = "class='clsAMO_ChangedCell'"
							case CHECKMARK_REMOVED
								bPutCheckmark = False
								strHighlight = "class='clsAMO_ChangedCell'"
					end select	
					strAvailableStatus = oRsCompatibility("AvailableStatus").value 
				end if
				'add bPermission
				'Call DisplayCheckmark_Comparability( bPutCheckmark,  not bLocked and not bDisabled and bAMOUpdate, strAvailableStatus, strTitle," pid='" &  sCompatibilityId & "' ", "comparability", strHighlight)
				Call DisplayCheckmark_Comparability( bPutCheckmark, bPermission and not bLocked and not bDisabled and bAMOUpdate, strAvailableStatus, strTitle, " pid='" & sCompatibilityId & "' ", "comparability", strHighlight, nModuleID, sCompatibilityId)
				oRsBusSeg.MoveNext
			loop

			'*NEED DIVISIONS ADDED TO PLATFORMS SEARCH: ----
			if oRsPlatforms.RecordCount > 0 then
				arrDivs = split(sDivisionIDs, ",")
				oRsPlatforms.MoveFirst
				strPlatform = ""
				strPlatformID = ""
				iFeatureID = ""
				strHighlight = ""
				bPutCheckmark = false
				bDisplayed = false
				strDisplayed = ""
				intCount = 0
												
				do while not oRsPlatforms.EOF	
					'determine if the user's assiged Team Divisions supports the feature and thus will allow editing
					'there can be multiple Roles that could be assigned to a Users so we will have to go through
					'one by one and compare the Team Divsion's against the platform's divisionids
					bPlatPermission = isIdBelong(Trim(sDivisionIDs), "," & cstr(oRsPlatforms.Fields("DivisionIDs").Value) & ",") 'Checks if User's Team Division IDs are linked to Platform; part of determining if User can edit Platform

					'determine if the platform supports the module and thus will allow editing
					'there can be multiple divisions for the module so we will have to go through
					'one by one and compare against the platform's divisionids
					bInDivision = isIdBelong(Trim(rsModule.Fields("DivisionIDs").value), "," & cstr(orsPlatforms("DivisionIDs")) & ",") 'Checks if Platform's DivisionIDs are linked to AMO Feature; part of determining if User can edit Platform

					iFeatureID = oRsPlatforms("FeatureID").value

					if strPlatform = "" then 'Get the first Product Platform
						strPlatform = oRsPlatforms("ProductPlatformName").value
					end if

					if oRsPlatforms("ProductPlatformName").value <> strPlatform and strPlatform <> "" then
						'If table cell not created for Feature, create cell before moving to next Platform column
						if not bDisplayed then
							if bInDivision then
								Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, bPermission and bInDivision and bPlatPermission and bAMOUpdate and not bLocked and not bDisabled, strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID)
							else
								Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, false, "Platform is not in Option's business segment." & vbCrLF & strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID)
							end if
						end if
						
						strPlatform = oRsPlatforms("ProductPlatformName").value
						
						bDisplayed = false
						bPutCheckmark = False
						strAvailableDate = ""
						strAvailableStatus = ""
						strHighlight = ""
					end if

					'If this is the Current Product Platform, check if FeatureID matches the parent, Module recordset's current FeatureID
					if (oRsPlatforms("ProductPlatformName").value = strPlatform) then 

						'If FeatureID matches parent Module recordset loop, create a table cell
						if (iFeatureID = nModuleID)  then
							strTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo :" & rsModule.Fields("BluePN").Value & ") / " + strPlatform)
							strPlatformID = oRsPlatforms("PlatformID").value
							strExtra = " pid='" & strPlatformID & "'"

							select case oRsPlatforms("ChangeMask")
								case NO_CHANGE
									bPutCheckmark = True
									strAvailableDate = oRsPlatforms("AvailableDate").value
									strHighlight = ""
								case CHANGE_MADE
									bPutCheckmark = True
									strAvailableDate = oRsPlatforms("AvailableDate").value
									strHighlight = "class='clsAMO_ChangedCell'"
								case CHECKMARK_REMOVED
									bPutCheckmark = False
									strHighlight = "class='clsAMO_ChangedCell'"
							end select
						
							if strAvailableDate = "" then
								strAvailableStatus = oRsPlatforms("AvailableStatus").value
							else
								strAvailableStatus = " (" & oRsPlatforms("AvailableStatus").value & ")"
							end if

							if bInDivision then
								Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, bPermission and bInDivision and bPlatPermission and bAMOUpdate and not bLocked and not bDisabled, strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID )
							else
								Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, false, "Platform is not in Option's business segment." & vbCrLF & strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID )
							end if

							bDisplayed = true
						else
							strTitle = server.htmlencode(rsModule.Fields("ShortDescription").Value & " (HP PartNo :" & rsModule.Fields("BluePN").Value & ") / " + strPlatform)
							strPlatformID = oRsPlatforms("PlatformID").value
							strExtra = " pid='" & strPlatformID & "'"
							
							'Else, make sure required cell values are blank						
							bPutCheckmark = False
							strAvailableDate = ""
							strAvailableStatus = ""
							strHighlight = ""
						end if

					end if
					
					intCount = intCount + 1
					oRsPlatforms.MoveNext
				loop

				if intPlatformLastRecord = intCount then
					if not bDisplayed then
						if bInDivision then
							Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, bPermission and bInDivision and bPlatPermission and bAMOUpdate and not bLocked and not bDisabled, strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID )
						else
							Call DisplayCheckmark_Platform(bPutCheckmark, strAvailableDate, strAvailableStatus, false, "Platform is not in Option's business segment." & vbCrLF & strTitle, strExtra, "plat", strHighlight, nModuleID, strPlatformID )
						end if
					end if
				end if

			end if
	
			Response.Write  "</tr>" & vbCrLf
			rsModule.MoveNext
			if rsModule.EOF then exit do
		loop
		if rsModule.EOF then exit do
	loop
	Response.Write  "</table>" & vbCrLf
	oRsPlatforms.Close
	set oRsPlatforms = nothing
end function


function WritePropertiesLocalizationGridHTML (ByVal strDivisionID, ByVal strModuleID, ByRef oRsCheckedRegion,  byval sCtrlStyle, byVal nMode, ByRef sHubCheckboxlist, ByRef sHideHubCheckboxlist)
	dim strGEO, strRegion, strGEOID, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay, strHighlight, strGEODate
	dim nModuleID
	dim intCount, intChangeMask, intStatusID, intColCount
	dim bDisplayed, bLocked, bPutCheckmark, bDisabled, strAllRegionIDs
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr, strChecked, strHideChecked
	dim strCPLBlindDate, strBOMRevADate, strRASDiscontinueDate, strObsoleteDate, strGlobalSeriesDate 

	const NO_CHANGE = 0
	const CHANGE_MADE = 1
	const CHECKMARK_REMOVED = 2
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	set oRsAllRegions = oSvr.AllRegions_Search(Application("REPOSITORY"), strDivisionID, 0)
	if (oRsAllRegions is nothing) then
		Response.Write("Empty Recordset: oRsAllRegions")
		set oRsAllRegions = nothing
		WritePropertiesLocalizationGridHTML = "Empty Recordset: oRsAllRegions"
		exit function
	end if
	oRsAllRegions.Sort = "SortOrder ASC, Region ASC"
	
	if oRsAllRegions.RecordCount = 0 then
		sErr = "No Regions have been assigned to the selected Business Segments. Please contact an AMO Admin to have them assign regions.; AMO_Localization.asp"
		response.write sErr
		set oRsAllRegions = nothing
		WritePropertiesLocalizationGridHTML = sErr
		exit function
	end if

	set oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleID, 1)
	if (oRsGEOs is nothing) then
		Response.Write("Empty Recordset: oRsGEOs")
		set oRsGEOs = nothing
		WritePropertiesLocalizationGridHTML = "Empty Recordset: oRsGEOs"
		exit function
	end if
	set oRsRegions = oSvr.AMORegions_Search(Application("REPOSITORY"), strModuleID, 1)
	if (oRsRegions is nothing) then
		Response.Write("Empty Recordset: oRsRegions")
		set oRsRegions = nothing
		WritePropertiesLocalizationGridHTML = "Empty Recordset: oRsRegions"
		exit function
	end if
	
	Set oRsCheckedRegion = CopyRs(oRsRegions, True)
	
	if Request.QueryString("nEditLocalization") <> "" then
		Response.Write  "<table id='oTable' name='oTable' align='left' border='1'  CELLSPACING='0' CELLPADDING='3' style='color:blue;'>" & vbCrLf
	else
		Response.Write  "<table id='oTable' name='oTable' align='left' border='1'  CELLSPACING='0' CELLPADDING='3' style='color:blue;display:none'>" & vbCrLf
	end if
    
	'Row 1; Heading
	Response.Write  "<tr class='tblrow-pulsar'>" 
	Response.Write  "<th align='center' valign='bottom' width='60' nowrap NAME='misc' >GEO</th>" 
	Response.Write  "<th align='center' valign='bottom' width='120' nowrap NAME='misc'>Country / Region</th>" 
	Response.Write  "<th align='center' valign='bottom' width='80' nowrap NAME='misc'>Localization Code</th>" 
	Response.Write  "<th align='center' valign='bottom' style='width: 115px;' NAME='misc'>Select Availability (SA)&nbsp;<font size=1 color='black'><i>(IRS calculated dates)</i></font></th>" 
	Response.Write  "<th bgcolor='gray' align='center' valign='bottom'style='width: 115px;' NAME='misc'>PHweb (General) Availability (GA)</th>" 
	Response.Write  "<th bgcolor='gray' align='center' valign='bottom' style='width: 115px;' NAME='misc'>End of Manufacturing (EM)</th>" 
	Response.Write  "<th align='center' valign='bottom' style='width: 115px;' NAME='misc'>End of Sales (ES)&nbsp;<font size=1 color='black'><i>(IRS calculated dates)</i></font></th>" 
	Response.Write  "<th align='center' valign='bottom' style='width: 115px;' NAME='misc'>Global Series Config EOL&nbsp;</th>" 
	Response.Write  "<th align='center' valign='bottom' width='30' nowrap NAME='misc'>Hub</th>" 
	Response.Write  "<th align='center' valign='bottom' width='30' nowrap NAME='misc'>Hide from Hub report</th>" 
	Response.Write  "</tr>"
	
	'display all the GEOs
	
	strAllRegionIDs = ""
	oRsAllRegions.MoveFirst
	
	bDisplayed = True
	do while (not oRsAllRegions.EOF)
	
		strChecked = "" 
		strHideChecked = ""
		strRegionID = oRsAllRegions("RegionID").value
	
		if instr(strAllRegionIDs, "," & strRegionID & ",") = 0 Then
	
			strAllRegionIDs = strAllRegionIDs & "," & strRegionID & ","
			oRsRegions.filter = "RegionID=" & strRegionID	
				
			if Trim(oRsAllRegions("Region").value) = "All Countries" then
				Response.Write  "<tr bgcolor='lightgrey' style='color:black'>"
			else
				Response.Write  "<tr style='color:black'>"
			end if
		
			if trim(ucase(oRsAllRegions("CountryGEO").value)) <> "GLOBAL" then
			   Response.Write  "<td id='misc'>" & oRsAllRegions("CountryGEO").value & "&nbsp;</td>"
			else
			   Response.Write  "<td id='misc'>&nbsp;</td>"
			end if
		
			if Trim(oRsAllRegions("Region").value) = "All Countries" then
				Response.Write  "<td id='misc'>" & oRsAllRegions("Region").value & "&nbsp;<font size=1 color='black'><i>(Base Module - data required)</i></font></td>"
			else
				Response.Write  "<td id='misc'>" & oRsAllRegions("Region").value & "&nbsp;</td>"
			end if
			Response.Write  "<td id='misc'>" & oRsAllRegions("HPDashCode").value & "&nbsp;</td>"
		
			if oRsRegions.RecordCount > 0 then
		
				if sHubCheckboxlist = "" then
					sHubCheckboxlist = oRsRegions("RegionID").Value
				else
					sHubCheckboxlist = sHubCheckboxlist & "," & oRsRegions("RegionID").Value 
				end if
				
				if oRsRegions("IsHubSku").Value Then
					strChecked = " checked "
				end if
		
				if oRsRegions("HideFromHub").Value Then
					strHideChecked = " checked "
				end if
				
				'if nMode <> 3 then
					strCPLBlindDate = oRsRegions("CPLBlindDate").Value
					strBOMRevADate = oRsRegions("BOMRevADate").Value
					strRASDiscontinueDate = oRsRegions("RASDiscontinueDate").Value
					strObsoleteDate = oRsRegions("ObsoleteDate").Value
					strGlobalSeriesDate = oRsRegions("GlobalSeriesDate").Value
					'if strCPLBlindDate <> "" then
					'	sNewCPLLocaleString = sNewCPLLocaleString & strRegionID & ",|"
					'end if 
					
				'else
				'	strCPLBlindDate = ""
				'	strBOMRevADate = ""
				'	strRASDiscontinueDate = ""					
				'	strObsoleteDate = ""
				'	strGlobalSeriesDate = ""
				'end if

				if sCtrlStyle = "" and Request.QueryString("nEditLocalization") <> ""  then
					Response.Write "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "cplblinddate" ) & "><input name='txtCPLBlindDate" & strRegionID & "' id='txtCPLBlindDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					strCPLBlindDate & "' onBlur='CPLBlindDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtCPLBlindDate" & strRegionID & "', document.thisform.txtCPLBlindDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "bomrevadate") & "><input name='txtBOMRevADate" & strRegionID & "' id='txtBOMRevADate" & strRegionID & "' maxlength=20 size=10 value='" & _
					strBOMRevADate & "' onBlur='BOMRevADate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtBOMRevADate" & strRegionID & "', document.thisform.txtBOMRevADate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "rasdiscontinuedate" ) & "><input name='txtRASDiscontinueDate" & strRegionID & "' id='txtRASDiscontinueDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					strRASDiscontinueDate & "' onBlur='RASDiscontinueDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtRASDiscontinueDate" & strRegionID & "', document.thisform.txtRASDiscontinueDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "obsoletedate" ) & "><input name='txtObsoleteDate" & strRegionID & "' id='txtObsoleteDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					strObsoleteDate & "' onBlur='ObsoleteDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtObsoleteDate" & strRegionID & "', document.thisform.txtObsoleteDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
					
					if oRsRegions("RegionID").Value = Application("AMO_GLOBAL_REGIONID") Then
			
						Response.Write "<td align='center'><input name='txtGlobalseriesDate' id='txtGlobalseriesDate' maxlength=20 size=10 value='" & _
						strGlobalSeriesDate & "' onBlur='GlobalseriesDate_Update(this, " & strRegionID & ");'>" & _
						"<a HREF=""javascript:popCalendar('document.thisform.txtGlobalseriesDate', document.thisform.txtGlobalseriesDate.value);"" onclick=event.cancelBubble=true;>" & _
						"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
					
					else
						Response.Write  "<td bgcolor='gray' align='center'>&nbsp;</td>" & vbCrLf
					end if
					
					if oRsRegions("RegionID").Value <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write "<td align='center'><input type='checkbox' name='chkHub' id='chkHub' value='" & strRegionID & "'" & strChecked & "></TD>"
					else
						Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					end if
					
					if oRsRegions("RegionID").Value <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write "<td align='center'><input type='checkbox' name='chkHideHub' id='chkHideHub' value='" & strRegionID & "'" & strHideChecked & "></TD>"
					else
						Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					end if

				else
					
					Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "cplblinddate" ) & ">" & strCPLBlindDate & "&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "bomrevadate" ) &  ">" & strBOMRevADate & "&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "rasdiscontinuedate" ) & ">" & strRASDiscontinueDate & "&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "obsoletedate" )  & ">" & strObsoleteDate & "&nbsp;</td>" & vbCrLf
					if oRsRegions("RegionID").Value = Application("AMO_GLOBAL_REGIONID") Then
						Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "globalseriesdate" )  & ">" & strGlobalSeriesDate & "&nbsp;</td>" & vbCrLf
					else
						Response.Write  "<td bgcolor='gray' align='center'>&nbsp;</td>" & vbCrLf
					end if
					
					if oRsRegions("IsHubSku").Value and oRsRegions("RegionID").Value <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "hub" )  & "><img src='../library/Images/gifs/bluechecktrans.gif' width=17 height=15 border=0></td>" & vbCrLf
					else
						Response.Write  "<td align='center' " & SetLocalizationHighlight(oRsRegions("ChangeMask"), "hub" )  & ">&nbsp;</td>" & vbCrLf
					end if
					
					if oRsRegions("HideFromHub").Value and oRsRegions("RegionID").Value <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write  "<td align='center'><img src='../library/Images/gifs/bluechecktrans.gif' width=17 height=15 border=0></td>" & vbCrLf
					else
						Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					end if
						
				end if
					 		
			else 
					
					
				if sCtrlStyle = "" and Request.QueryString("nEditLocalization") <> ""  then
					
					Response.Write "<td align='center'><input name='txtCPLBlindDate" & strRegionID & "' id='txtCPLBlindDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					"' onBlur='CPLBlindDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtCPLBlindDate" & strRegionID & "', document.thisform.txtCPLBlindDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center'><input name='txtBOMRevADate" & strRegionID & "' id='txtBOMRevADate" & strRegionID & "' maxlength=20 size=10 value='" & _
					"' onBlur='BOMRevADate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtBOMRevADate" & strRegionID & "', document.thisform.txtBOMRevADate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center'><input name='txtRASDiscontinueDate" & strRegionID & "' id='txtRASDiscontinueDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					"' onBlur='RASDiscontinueDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtRASDiscontinueDate" & strRegionID & "', document.thisform.txtRASDiscontinueDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					Response.Write "<td align='center'><input name='txtObsoleteDate" & strRegionID & "' id='txtObsoleteDate" & strRegionID & "' maxlength=20 size=10 value='" & _
					"' onBlur='ObsoleteDate_Update(this, " & strRegionID & ");'>" & _
					"<a HREF=""javascript:popCalendar('document.thisform.txtObsoleteDate" & strRegionID & "', document.thisform.txtObsoleteDate" &  strRegionID & ".value);"" onclick=event.cancelBubble=true;>" & _
					"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
					
					if strRegionID = Application("AMO_GLOBAL_REGIONID") Then

						Response.Write "<td align='center'><input name='txtGlobalseriesDate' id='txtGlobalseriesDate' maxlength=20 size=10 value='" & _
						"' onBlur='GlobalseriesDate_Update(this, " & strRegionID & ");'>" & _
						"<a HREF=""javascript:popCalendar('document.thisform.txtGlobalseriesDate', document.thisform.txtGlobalseriesDate.value);"" onclick=event.cancelBubble=true;>" & _
						"<img src=""../images/calendarButton.gif"" align=top ID=img1 name=img1 border=0 title=""Popup calendar"" WIDTH=26 HEIGHT=21></a></TD>"
						
					else
						Response.Write  "<td bgcolor='gray' align='center'>&nbsp;</td>" & vbCrLf
					end if
					
					if strRegionID <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write "<td align='center'><input type='checkbox' name='chkHub' id='chkHub' value='" & strRegionID & "'" & strChecked & "></TD>"
					else
						Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					end if
					
					if strRegionID <> Application("AMO_GLOBAL_REGIONID") Then
						Response.Write "<td align='center'><input type='checkbox' name='chkHideHub' id='chkHideHub' value='" & strRegionID & "'" & strHideChecked & "></TD>"
					else
						Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					end if
					
				else
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td bgcolor='gray' align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
					Response.Write  "<td align='center'>&nbsp;</td>" & vbCrLf
				end if
					
			end if
		end if
		
		Response.Write  "</tr>"
		oRsAllRegions.MoveNext
	loop

    Response.Write  "</table>" 
    oRsGEOs.Close
	set oRsGEOs = nothing
	oRsRegions.Close
	set oRsRegions = nothing
	oRsAllRegions.Close
	set oRsAllRegions = nothing

WritePropertiesLocalizationGridHTML = ""
end function
'********************************************************************************
'* Makes a local recordset that we can update
'* 
'* Parameters:
'*  oRs     = Recordset to use as the source of the copy
'*  bCopy   = true or false if existing data should be copied. If false, just the fields are created.
'*
'* Returns:
'*  Local Recordset - Excalibur data classes return disconnected records already, not needed
'********************************************************************************
function CopyRs(byval oRs, bCopy)
	dim oDupRs, oFld, i
    set oDupRs = Server.CreateObject ("ADODB.Recordset")
	set oDupRs = oRs

	'for i = 0 to oRs.Fields.Count - 1
		'set oFld =  oRs.Fields(i)
		'oDupRs.Fields.Append oFld.Name, oFld.Type, oFld.DefinedSize, oFld.Attributes
	'next
    'oDupRs.CursorLocation = 3	'Use client-side cursors
    'oDupRs.Open

	'if bCopy then
	    'If oRs.RecordCount > 0 Then
	        'oRs.MoveFirst
	    'End If
	
	    'While Not oRs.EOF
	        'oDupRs.AddNew
	        'For i = 0 To oRs.Fields.Count - 1
	            'oDupRs.Fields(oDupRs.Fields(i).Name).Value = oRs.Fields(i).Value
	        'Next
	        'oRs.MoveNext
	    'Wend
	
	    'If oRs.RecordCount > 0 Then
	        'oRs.MoveFirst
	    'End If
	'end if

	set CopyRs = oDupRs
end function

'Write the category filter for all tabs
sub WriteCategoryFilter(byval stab, ByVal sModuleCategoryHTML, ByVal sStatusHTML, byval sBusSegHTML, byval strEOLDate,_
			byVal bRAS, byVal bGpsyCom, byVal bRasCom,byval sDisplayHideromMOL, byval sDisplayHideromSCM, sDisplayHideromSCL, _
		  ByVal strMRFromDate, ByVal strMRToDate, ByVal strCPLFromDate, _
		  ByVal strCPLToDate, ByVal strDisFromDate, ByVal strDisToDate, _
		  ByVal strchkMRBlank, ByVal strchkCPLBlank, ByVal strchkDisBlank, _
		  strRasObsoleteFromDate, strRasObsoleteToDate, nCategoriesSelected, sOwnersHTML, sDisplayIncludeBlankGSD)
						  
	dim oRsAllCategorySelected, sErr, oRsAllCategory

	set oRsAllCategorySelected = Nothing
	sErr = ""
	'get the list of module/option category groups; Old code: set oErr = GetMOLCategory(oRsAllCategory, 33)	
	If (stab="All" & "&nbsp;" & "Options" or stab="WIP" or stab="Complete" or stab="Discontinue" or stab="Platforms" or stab="RAS" & "&nbsp;" & "Review") Then	
		set oRsAllCategory = GetMOLCategory(33)	
		if (oRsAllCategory is Nothing) then
			Response.Write("Recordset error: oRsAllCategory")
			Response.End()
		else
			if not oRsAllCategory is nothing then
				oRsAllCategory.Filter = "State = 'Active' and FeatureCategoryID <> 0"
				oRsAllCategory.Sort = "Name ASC"
			end if 
		end if
	Else
		set oRsAllCategory = GetMOLCategory(31)	
		if (oRsAllCategory is Nothing) then
			Response.Write("Recordset error: oRsAllCategory")
			Response.End()
		else
			if not oRsAllCategory is nothing then
				oRsAllCategory.Filter = "State = 1 and CategoryID <> 0"
				oRsAllCategory.Sort = "FullCatDescription ASC"
			end if 
		end if
	End If
	%>
	<tr>
		
		<td colspan='2'><% 
		If (stab="All" & "&nbsp;" & "Options" or stab="WIP" or stab="Complete" or stab="Discontinue" or stab="Platforms" or stab="RAS" & "&nbsp;" & "Review") Then	
			DualListboxRs_GetHTML6_Write oRsAllCategory, "Name", "FeatureCategoryID", oRsAllCategorySelected, _
			"Name", "FeatureCategoryID", true, true, nCategoriesSelected, "Available Categories", "Selected Categories", _
			"Category", true, 130, 360, false, true, false, 350, 10
		Else
			DualListboxRs_GetHTML6_Write oRsAllCategory, "FullCatDescription", "CategoryID", oRsAllCategorySelected, _
			"FullCatDescription", "CategoryID", true, true, nCategoriesSelected, "Available Categories", "Selected Categories", _
			"Category", true, 130, 360, false, true, false, 350, 10
		End If
		%></td>
	</tr>
	<%if not(stab="WIP" or stab="RAS" & "&nbsp;" & "Review" or stab="Complete" or stab="Discontinue") then	%>
	<tr>
		<td width=20%>AMO Status</td>
		<td><%=sStatusHTML%></td>
		</tr>
	<%end if %>	
	<tr>
		<td width=20%>Business Segment</td>
		<td><%=sBusSegHTML%></td>
	</tr>
	<!--//<tr>
		<td width=20%>Owned By</td>
		<td><%=sOwnersHTML%></td>
	</tr>//-->
	<% If bRAS Then %>
		<%if not(stab="WIP" or stab="Complete"  or stab="Discontinue") then	%>
		<tr>
			<td>Show Only RAS/GPSy With: </td>
			<td><INPUT id=chkGpsyCom name=chkGpsyCom type=checkbox <%=bGpsyCom%>>GPSy Incomplete Only <INPUT id=chkRasCom name=chkRasCom type=checkbox <%=bRasCom%>>RAS Incomplete Only</td>
		</tr>
		<%end if%>
	<%End If %>
	<%if not( stab="RAS" & "&nbsp;" & "Review" or stab="Complete"  or stab="Discontinue") then%>
		<tr>
		<td width=20%>Select Availability (SA)</td>
		<td>From&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtCPLFromDate NAME=txtCPLFromDate value="<%=strCPLFromDate %>" class="filter-dateselection">
		<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">To&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtCPLToDate NAME=txtCPLToDate value="<%=strCPLToDate%>" class="filter-dateselection">
		(MM/DD/YYYY)<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">
		</td>	
		</tr>
		<tr>
		<td width=20%>PHweb (General) Availability (GA)</td>
		<td>From&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtMRFromDate NAME=txtMRFromDate value="<%=strMRFromDate %>" class="filter-dateselection">
		<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">To&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtMRToDate NAME=txtMRToDate value="<%=strMRToDate%>" class="filter-dateselection">
		(MM/DD/YYYY)<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">	
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>Include All Blank Date</i>&nbsp;<INPUT id=chkMRBlankDate name=chkMRBlankDate type=checkbox <%=strchkMRBlank%>></td>
		</tr>	
		<tr>
		<tr>
		<td width=20%>End of Manufacturing (EM)</td>
		<td>From&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtDisFromDate NAME=txtDisFromDate value="<%=strDisFromDate %>" class="filter-dateselection">
		<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">To&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtDisToDate NAME=txtDisToDate value="<%=strDisToDate%>" class="filter-dateselection">
		(MM/DD/YYYY)<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">
		</td>	
		</tr>
		
		<tr>
		<td width=20%>End of Sales (ES)</td>
		<td>From&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtRasObsoleteFromDate NAME=txtRasObsoleteFromDate value="<%=strRasObsoleteFromDate %>" class="filter-dateselection">
		<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">To&nbsp;&nbsp;<input type=text maxlength=10 size=10 id=txtRasObsoleteToDate NAME=txtRasObsoleteToDate value="<%=strRasObsoleteToDate%>" class="filter-dateselection">
		(MM/DD/YYYY)<img src="../library/Images/gifs/empty.gif" alt="" width="5" height="1" border="0">
		</td>	
		</tr>
	<%end if %>	
	<%if not(stab="WIP" or stab="RAS" & "&nbsp;" & "Review" or stab="Complete" or stab="Discontinue") then	%>
		<tr>
			 <td colspan=2>
				Display options with "Hide from Product Requirements List (PRL)" checked:
				<INPUT id=chkIncludeHidefromMOL name=chkIncludeHidefromMOL type=checkbox <%=sDisplayHideromMOL%>>
			</td>
		</tr>
		<tr style=display:none;>
			 <td colspan=2>Display options with "Hide from SCL" checked:
			<INPUT id=chkIncludeHidefromSCL name=chkIncludeHidefromSCL type=checkbox ></td>
		</tr>
	<%end if %>	
	
	<%if not(stab="WIP" or stab="RAS" & "&nbsp;" & "Review" or stab="Complete" or stab="Discontinue") then	%>
		<tr>
			 <td colspan=2>Display options with "Hide from AMO SCM" checked: 
			<INPUT id=chkIncludeHidefromSCM name=chkIncludeHidefromSCM type=checkbox <%=sDisplayHideromSCM%>></td>
		</tr>
	<%end if %>	
			
	<tr>
		<td colspan=2>
	 	<%if not(stab="Change" & "&nbsp;" & "History" or stab="Report" or stab="Rules") then	%>
			Include blank Global Series Dates checked:
			<INPUT id=chkIncludeBlankGSD name=chkIncludeBlankGSD type=checkbox <%=sDisplayIncludeBlankGSD%>>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%end if%>		
			
		<INPUT type='button' value='Filter List' id=btnRefresh NAME=btnRefresh LANGUAGE=javascript onclick="return lbxGo_onchange();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% if stab <> "Platforms" then %>
		<INPUT type='button' value='Select Column To Hide' id=btnHide NAME=btnRefresh LANGUAGE=javascript onclick="return lbxGo_onhide();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<% end if %>
	 	<%if stab="RAS" & "&nbsp;" & "Review" then	%>
		<input type="button" name="btnCreateExportFile" id="btnCreateExportFile" value="Create Export Files" onclick="return btnCreateExportFile_onclick()">
		<% end if %>
		</td>
	</tr>
	
	<tr width=100%><td colspan=2><hr></td></tr>
	<%
end sub

'Determine if the user is in a division assigned to the Options in the AMO
function UserInAMODivision(strDivisionIDs)
	dim sErr, strItem
	dim oRs, oErr, oSvr
	dim bFound
	dim arrDivisions

	if instr(strDivisionIDs, ",") = 0 then
		strDivisionIDs = trim(strDivisionIDs) & ","
	end if

	'get user's groups to see if any division matches
	bFound = False
	
	set oSvr = New ISGeneral
	set oRs = oSvr.GetGroupsByUser( Application("Repository"), clng(Session("AMOUserID")))
	set Session("rsAMOGroups") = oRs

	if oRs.RecordCount > 0 then
		oRs.MoveFirst
		do until oRs.EOF
			arrDivisions = split(oRs("DivisionIDs").value, "|")
			for each strItem in arrDivisions
				if instr(strDivisionIDs, strItem & ",") > 0 then
					bFound = True
					exit do
				end if
			next
			oRs.MoveNext
		loop
	end if
	
	UserInAMODivision = bFound
end function

'Display a table cell that has text
sub DisplayField( bRights, strID, strName, strAlign, strValue, strHighlight, ByVal strTitle, byVal bBaseModule, ByVal sHideColumns)
	
	dim strHeader, strAlID', strUnderline
	
'	if bUnderline then
'		strUnderline = "text-decoration:underline"
'	else
'		strUnderline = "text-decoration:none"
'	end if
	
	strAlID = "xxx"
	if instr(strID, "cpl") > 0 then
	  strHeader = "Select Availability (SA)"
	  strAlID = "1"
	elseif instr(strID, "wwp") > 0 then
	  strHeader = "AMO Price"
	  strAlID = "2"
	elseif instr(strID, "obd") > 0 then
	 strHeader = "End of Sales (ES)"
	 strAlID = "3"
	elseif instr(strID, "reva") > 0 then
	 strHeader = "PHweb (General) Availability (GA)"
	 strAlID = "4"
	elseif instr(strID, "rasdisc") > 0 then
	 strHeader = "End of Manufacturing (EM)"
	 strAlID = "5"
	elseif instr(strID, "gbsd") > 0 then
	 strHeader = "Global Series Config EOL"
	 strAlID = "6"
	else
	
		Select Case strID 
			case "prop" 'Properties column
				strHeader = "Properties"
			case "o": 'Short Description column
				strHeader = "Short Description"
			case "asnew", "asras", "as", "asdis" 'AMO status column for New, RAS Review modules, and any status module
		    strHeader = "Status"
			case "ctr" 'Comment to RAS column
				strHeader = "Comment to RAS"
			case "cfr" 'Comment from RAS column
				strHeader = "Comment from RAS"
			case "bpn" 'Blue PN column
				strHeader = "HP Part Number"
			case "rpn" 'Red PN column
				strHeader = "Red Part Number"
			case "molhide" 'Hide from PRL
				strHeader = "Hide from PRL"
				strAlID = "7"
			case "scmhide" 'Hide from SCM
				strHeader = "Hide from SCM"
				strAlID = "26"
			case "sclhide" 'Hide from SCL
				strHeader = "Hide from SCL"
				strAlID = "27"
			case "ras" 'RAS Complete column
				strHeader = "RAS Complete"
			case "gpsy"  'GPSy Complete column
				strHeader = "GPSy Complete"
			case "cst" 'AMO Cost column
				strHeader = "AMO Cost"
				strAlID = "8"
			case "acst" 'Actual Cost column
				strHeader = "Actual Cost"
				strAlID = "9"
			case "rep" 'Replacement column
				strHeader = "Replacement"
				strAlID = "25"
			case "alt" 'Alternative column
				strHeader = "Alternative"
				strAlID = "10"
			case "nw" 'Net Weight column
				strHeader = "Net Weight"
				strAlID = "12"
			case "ew" 'Export Weight column
				strHeader = "Export Weight"
				strAlID = "11"
			case "apw" 'Air Packed Weight column
				strHeader = "Air Packed Weight"
				strAlID = "13"
			case "apc" 'Air Packed Cubic column
				strHeader = "Air Packed Cubic"
				strAlID = "14"
			case "ec" 'Export Cubic column
				strHeader = "Export Cubic"
				strAlID = "15"
			case "rc" 'Regional Comment
				strHeader = "Regional Comment"
			case "own" 'Owned By
				strHeader = "Owned By"
				strAlID = "16"
			case "warc" 'Warranty Code
				strHeader = "Warranty Code"	
				strAlID = "17"	
			case "sf" 'Serial Flag
				strHeader = "Serial Flag"	
				strAlID = "29"		
			case "maco" 'Manufacture Country
				strHeader = "Manufacture Country" 
				strAlID = "18"	
			case "dsc"
				strHeader = "Long Description"
				strAlID = "19"
			case "rdsc"
				strHeader = "Rules Description"
				strAlID = "20"
			case "reldes"
				strHeader = "Replacements Description"
				strAlID = "21"
			case "ord"
				strHeader = "Order Instruction"
				strAlID = "22"
			case "prodl"
				strHeader = "Product Line"
				strAlID = "23"
			case "idp" 'ignore deployment plan - 4-29-2008 Added by Dien
				strHeader = "Ignore SCL Deployment Plan"
				strAlID = "28"
		end select
		
		
		Select Case strName
			case "vna" 'Visibility_NA
				strHeader = "SKU Visibility - NA"
				strAlID = "24"
			case "vem" 'Visibility_EM
				strHeader = "SKU Visibility - EMEA"
				strAlID = "24"
			case "vap" 'Visibility_AP
				strHeader = "SKU Visibility - APJ"
				strAlID = "24"
			case "vla" 'Visibility_LA
				strHeader = "SKU Visibility - LA"	
				strAlID = "24"		
		end select
					
	end if
	
	if strTitle <> "" then	
		strTitle = " title=""" & strHeader & " : " & strTitle & " """
	end if
	
	if strID = "o" or strID = "cfr" or strID = "ctr" or strID = "bpn" or strID = "rpn" or strID = "rep" or strID = "alt" Then
		strValue = replace(strValue,"<","< ")  
		strValue = replace(strValue,">"," >") 
	end if
	
	if instr("," & sHideColumns & ",", "," & strAlID & ",") = 0 Then
	
		if bBaseModule then
	
			if bRights then
			 if strHeader = "AMO Price" then
				Response.Write  "<td " & strHighlight & " id='" & strID & "' NAME='" & strName & "' style='color:black;' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
			 else
			 	'added underline 7/9/07
				Response.Write  "<td " & strHighlight & " id='" & strID & "' NAME='" & strName & "' style='text-decoration:underline' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
			 end if
			else				
				Response.Write  "<td " & strHighlight & " id='sd' style='color:black;text-decoration:none' NAME='sd' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
			end if
		else
	
			Response.Write  "<td bgcolor='lightgrey' id='sd' style='color:black;text-decoration:none' NAME='sd'>&nbsp;</td>" & vbCrLf
	
		end if
		
	end if
	
end sub


'Display a table cell that has text
sub DisplayField_Status( bRights, strID, strName, strAlign, strValue, strHighlight, ByVal strTitle, ByVal nModuleId, byVal bBaseModule)
	
	if strTitle <> "" then
		strTitle = " title=""" & "Status : " & strTitle & " """
	end if
	
	if bBaseModule then
	
		if bRights then
			'added underline 7/9/07
			Response.Write  "<td " & strHighlight & " id='" & strID  & nModuleId & "' style='text-decoration:underline' NAME='" & strName  & "' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
		else
			Response.Write  "<td " & strHighlight & " id='" & strID  & nModuleId & "' style='color:black;text-decoration:none' NAME='" & strName & "' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
		end if
	else
		Response.Write  "<td bgcolor='lightgrey' id='sd' style='color:black;text-decoration:none' NAME='sd'>&nbsp;</td>" & vbCrLf
	end if
	
end sub

'Display a table cell that has text
sub DisplayField_PartNo( bRights, strID, strName, strAlign, strValue, strHighlight, ByVal strTitle, byVal bBaseModule)

	if strTitle <> "" then
		strTitle = " title=""" & "HP PartNo : " & strTitle & " """
	end if
		
	if bBaseModule then
	
		if bRights then
			'added underline 7/9/07
			Response.Write  "<td " & strHighlight & " id='" & strID & "' style='text-decoration:underline' NAME='" & strName & "' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
		else
			Response.Write  "<td " & strHighlight & " id='sd' style='color:black;text-decoration:none' NAME='sd' " & strAlign & strTitle & ">" & strValue & "</td>" & vbCrLf
		end if
	else
		Response.Write  "<td bgcolor='lightgrey' id='sd' style='color:black;text-decoration:none' NAME='sd'>" & strValue & "</td>" & vbCrLf
	end if
	
end sub

'4-29-2008 Added by Dien
'Display a table cell that has a checkmark or not
sub DisplayCheckmark_IDP( bPutCheckmark, bRights, byval strTitle, strExtra, strID, strHighlight, statusid, byVal bBaseModule, nModuleID, ByVal sHideColumns)
	dim strAlID
	strAlID = "28"
	strTitle = "Ignore SCL Deployment Plan" & " : " & strTitle
	if instr("," & sHideColumns & ",", "," & strAlID & ",") = 0 Then
		
		if bBaseModule then	
			
			if bPutCheckmark then
				
				if bRights and statusid = 170 then									
					Response.Write  "<td align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & _
					"<img id='" & strID & "' title=""" & strTitle & """" & strExtra & " " & _
					"onclick='javascript:ClickEvent(event);return true;' src=""../library/Images/gifs/bluechecktrans.gif"" alt="""" width=17 height=15 border=0>" & "</td>" & vbCrLf					
				else
					Response.Write  "<td align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & _
					"<img id='sd' " & "title=""" & strTitle & """ " & _
					"src=""../library/Images/gifs/bluechecktrans.gif"" alt="""" width=17 height=15 border=0>" & "</td>" & vbCrLf
				end if
			else	
				if bRights and statusid = 170 then
					Response.Write  "<td align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
				else
					Response.Write  "<td align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
				end if
			end if

		else
			Response.Write  "<td bgcolor='lightgrey' align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		end if
	end if
	
end sub

'Display a table cell that has a checkmark or not
sub DisplayCheckmark_MOLHide( bPutCheckmark, bRights, byval strTitle, strExtra, strID, strHighlight, strModuleType, byVal bBaseModule, nModuleID, ByVal sHideColumns)
	dim strHeader, strAlID
	
	strHeader = ""
	Select Case strID 
		case "molhide" 'Hide from PRL
			strHeader = "Hide from PRL"
			strAlID = "7"
		case "scmhide" 'Hide from SCM
			strHeader = "Hide from SCM"
			strAlID = "26"
		case "sclhide" 'Hide from SCL
			strHeader = "Hide from SCL"
			strAlID = "27"
	end select
   
	strTitle = strHeader & " : " & strTitle

	if instr("," & sHideColumns & ",", "," & strAlID & ",") = 0 Then
	
		if bBaseModule then
   
			if bPutCheckmark then
				'added the be_hardwarecondition 11/10/2004, sw is not allowed to be unchecked. This is only for MOL.
				if bRights then
					if trim(strModuleType) = "Hardware" or strID <> "molhide" then
						Response.Write  "<td align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & _
						"<img id='" & strID & "' title=""" & strTitle & """" & strExtra & " " & _
						"onclick='javascript:ClickEvent(event);return true;' src=""../library/Images/gifs/bluechecktrans.gif"" alt="""" width=17 height=15 border=0>" & "</td>" & vbCrLf
					else
						Response.Write  "<td align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & _
						"<img id='sd' " & "title=""" & strTitle & """ " & _
						"src=""../library/Images/gifs/bluechecktrans.gif"" alt="""" width=17 height=15 border=0>" & "</td>" & vbCrLf
					end if
				else
					Response.Write  "<td align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & _
					"<img id='sd' " & "title=""" & strTitle & """ " & _
					"src=""../library/Images/gifs/bluechecktrans.gif"" alt="""" width=17 height=15 border=0>" & "</td>" & vbCrLf
				end if
			else
				if bRights then
					Response.Write  "<td align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
				else
					Response.Write  "<td align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
				end if
			end if

		else
			Response.Write  "<td bgcolor='lightgrey' align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
					"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		end if
	end if
end sub

'Display a table cell that has a checkmark or not
sub DisplayCheckmark_Comparability( bPutCheckmark, bRights, strAvailableStatus, strTitle, strExtra, strID, strHighlight, nModuleID, strPID)
	if bPutCheckmark then
		if bRights then
			Response.Write  "<td align='center' " & strHighlight & " id='" & strID & cstr(nModuleID) & "p" & strPID & "' name='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableStatus & "</td>" & vbCrLf
		else
			Response.Write  "<td align='center' " & strHighlight & " name='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableStatus & "</td>" & vbCrLf
		end if
	else
		if bRights then
			Response.Write  "<td align='center' " & strHighlight & " id='" & strID & cstr(nModuleID) & "p" & strPID & "' name='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		else
			Response.Write  "<td align='center' " & strHighlight & " name='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		end if
	end if
end sub

'Display a table cell that has a checkmark or not in Platform tab
sub DisplayCheckmark_Platform( bPutCheckmark, strAvailableDate, strAvailableStatus, bRights, strTitle, strExtra, strID, strHighlight, nModuleID, strPlatformID)
	if bPutCheckmark then
		if bRights  then
			Response.Write  "<td  align='center' " & strHighlight & " id='" & strID & cstr(nModuleID) & "p" & strPlatformID & "' name='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableDate &  strAvailableStatus & "</td>" & vbCrLf
		else
			Response.Write  "<td  align='center' " & strHighlight & " name='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableDate &  strAvailableStatus & "</td>" & vbCrLf
		end if
	else
		if bRights then
			Response.Write  "<td  align='center' " & strHighlight & " id='" & strID & cstr(nModuleID) & "p" & strPlatformID & "' name='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		else
			Response.Write  "<td  align='center' " & strHighlight & " name='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		end if
	end if
end sub

'Display a table cell that has a checkmark or not
sub DisplayCheckmark( bPutCheckmark, strAvailableDate, strAvailableStatus, bRights, strTitle, strExtra, strID, strHighlight, nModuleID)
	if bPutCheckmark then
		if bRights  then
			Response.Write  "<td  align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableDate &  strAvailableStatus & "</td>" & vbCrLf
		else
			Response.Write  "<td  align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & ">" & strAvailableDate &  strAvailableStatus & "</td>" & vbCrLf
		end if
	else
		if bRights then
			Response.Write  "<td  align='center' " & strHighlight & " class='" & strID & cstr(nModuleID) & "' id='" & strID & "' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		else
			Response.Write  "<td  align='center' " & strHighlight & " id='sd' title=""" & server.htmlencode(strTitle) & """" & strExtra & " " & _
			"style='color:black;text-decoration:none'>&nbsp;" & "</td>" & vbCrLf
		end if
	end if
end sub


' append ModuleIDs together
function CombineModuleIDs(oRs)
	dim strModuleIDs
	
	strModuleIDs = ""
	oRs.MoveFirst
	do while (not oRs.EOF)
		strModuleIDs = strModuleIDs & oRs("FeatureID").value & ","
		oRs.MoveNext
	loop
	if right(strModuleIDs, 1) = "," then
		strModuleIDs = left(strModuleIDs, len(strModuleIDs)-1)
	end if

	CombineModuleIDs = strModuleIDs
end function

'Determines if the table cell should be highlighted
function SetHighlight(intChangeMask, strField)
	dim intMask
	intMask = 0
	SetHighlight = ""
	'msgBox "STEVE, exec setHighLight fcn "
	select case lcase(strField)
		case "bluepartno"
			intMask = 1
		case "ruledescription"
			intMask = 2
		case "bomrevadate"
			intMask = 4
		case "rasdiscontinuedate"
			intMask = 8
		case "cplblinddate"
			intMask = 16
		case "amocost"
			intMask = 32
		case "amowwprice"
			intMask = 64
		case "replacement"
			intMask = 128
		case "alternative"
			intMask = 256
		case "netweight"
			intMask = 512
		case "exportweight"
			intMask = 1024
		case "airpackedweight"
			intMask = 2048
		case "airpackedcubic"
			intMask = 4096
		case "exportcubic"
			intMask = 8192
		case "regioncomment"
			intMask = 16384
		case "description"
			intMask = 32768
		case "shortdescription"
			intMask = 65536
		case "actualcost"
			intMask = 131072
		case "infocomment"
			intMask = 262144
		case "rascomment"
			intMask = 524288
		case "visibility_na" 
			intMask = 1048576
		case "visibility_ap"
			intMask = 2097152
		case "visibility_em"
			intMask = 4194304
		case "visibility_la"
			intMask = 8388608
		case "warrantycode"
			intMask = 16777216
		case "obsoletedate"
			intMask = 33554432
		case "manufacturecountry"
			intMask = 67108864
		case "longdescription"
			intMask = 134217728
		case "orderinstruction"
			intMask = 268435456
		case "replacementavdescription"
			intMask = 536870912
		case "productline"
			intMask = 1073741824			
	end select	

	if intMask <> 0 then
		if (intChangeMask AND clng(intMask)) = intMask then
			SetHighlight = "class='clsAMO_ChangedCell'"
		end if
	end if
end function


'Determines if the table cell should be highlighted
function SetLocalizationHighlight(intChangeMask, strField)
	dim intMask
	intMask = 0
	SetLocalizationHighlight = ""

	select case lcase(strField)
		case "bomrevadate"
			intMask = 16
		case "rasdiscontinuedate"
			intMask = 8
		case "cplblinddate"
			intMask = 32
		case "obsoletedate"
			intMask = 4	
		case "globalseriesdate"
			intMask = 64		
	end select	

	if intMask <> 0 then
		if (intChangeMask AND intMask) = intMask then
			SetLocalizationHighlight = " class='clsAMO_ChangedCell' "
		end if
	end if
end function

function WriteOptionCategoryRulesGridHTML (byval sCtrlStyle, byval oRsOptionRules)

	dim arrColors, irowcount
	arrColors = Array("#CCCCCC", "white")
	irowcount = 1
	Response.Write  "<br><table border='1'  CELLSPACING='2' CELLPADDING='2' >" & vbCrLf
	Response.Write  "<tr bgcolor='lightgrey' style='color:black'>"
	Response.Write  "<th align='center' width='5%' nowrap NAME='misc'>&nbsp;</th>" 
	Response.Write  "<th align='center' width='10%' nowrap NAME='misc'>Category</th>" 
	Response.Write  "<th align='center' width='20%' nowrap NAME='misc'>Applies to SCMs Published by</th>" 
	Response.Write  "<th align='center' width='55%' nowrap NAME='misc'>Rules Description</th>" 
	Response.Write  "<th align='center' width='5%' nowrap NAME='misc'>Min</th>" 
	Response.Write  "<th align='center' width='5%' nowrap NAME='misc'>Max</th>"
	Response.Write  "</tr>"
	
	
	'add &MsgMode=1 in AMO_OptioncategoryRule.asp for NOT prompt msg in Edit mode
	'display all the GEOs
	oRsOptionRules.MoveFirst
	do while (not oRsOptionRules.EOF)
		Response.Write  "<tr bgcolor='" & arrColors(irowcount mod 2) & "' style='color:black'>"
		if sCtrlStyle = " disabled " Then
			Response.Write  "<td align='Center' width='5%'>&nbsp;</td>"
		else
			Response.Write  "<td align='Center' width='5%'><a href='AMO_OptionCategoryRule.asp?Mode=1&MsgMode=1&RuleID=" & oRsOptionRules("RuleID").Value & "&CategoryID=" & oRsOptionRules("CategoryID").Value & "'>edit</a>&nbsp;&nbsp;<a href='javascript:btnRemove_onclick(" & oRsOptionRules("CategoryID").Value & "," & oRsOptionRules("RuleID").Value & ");'>remove</a></td>"
		end if
		Response.Write  "<td align='Center' width='10%'>" & oRsOptionRules("Category_Desc").Value & "</td>"
		Response.Write  "<td align='Center' width='20%'>" & oRsOptionRules("Divisions").Value & "</td>"
		Response.Write  "<td align='Center' width='55%'>" & server.HTMLEncode(oRsOptionRules("Rule_Desc").Value) & "</td>"
		Response.Write  "<td align='Center' width='5%'>" & oRsOptionRules("Rule_Min").Value & "</td>"
		Response.Write  "<td align='Center' width='5%'>" & oRsOptionRules("Rule_Max").Value & "</td>"
		Response.Write  "</tr>"
		irowcount = irowcount + 1
		oRsOptionRules.MoveNext
	loop
	
	Response.Write  "</table>" 
	
	WriteOptionCategoryRulesGridHTML = ""
	
end function

function GetDelta(rsRegion, sModuleID)
	dim sValue
	sValue = ""
	rsRegion.Filter = "ModuleID = " &  sModuleID
	do while (not rsRegion.EOF)
		if rsRegion("ReportChangeMask").value > 0 Then
			sValue = "C"
		    Exit Do
		End If	
		rsRegion.MoveNext
	Loop
	GetDelta = sValue
end function

'Determines if the table cell should be highlighted
function SetRegionHubHighlight(intChangeMask)
			
	if  intChangeMask = True then
		SetRegionHubHighlight = " bgcolor='#00FF00' "
	else
		SetRegionHubHighlight = ""
	end if

end function

function WriteAMO_SCM_HUB_ReportGridHtml (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs, byVal sUserBusSeg)
	dim strGEO, strRegion, strDivisionIDs, strModuleIDs
	dim strGEOID, strGEOIDs, strRegionIDs, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay
	dim nModuleID, intCompatibility
	dim intCount, intIndex
	dim bDisplayed, bLocked
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr
	dim strClass, sPD, sChange, sErr, strPlatform, strMarketingName, strRevision
	dim intPlatformID
	dim nChassisID, numDL
	dim oRsPlatforms, rsModuleCopy
	dim bPlatformAssigned, sDelta
	dim PlatformCt, i, arrDate, oRsBusSeg, strAllRegionIDs
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup
	
	'set rsModuleCopy = CopyRS(rsModule, True)
	
	rsModule.filter = "RegionID=334"
	rsModule.movefirst	
	
	strDivisionIDs = rsModule("DivisionIDs").value
	strRevision = rsModule("Revision").value
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO
	set oRsAllRegions = oSvr.AllRegions_Search(Application("REPOSITORY"), strDivisionIDs, 1)
	if (oRsAllRegions is nothing) then
		response.write "<tr><td><strong>Error for AMO_HUB_Report</strong></td></tr>"
		set oRsAllRegions = nothing
		exit function
	end if
	
	
	'combine all moduleids and pass it to the GEO and Region functions
	strModuleIDs = ""
	rsModule.MoveFirst
	do while (not rsModule.EOF)
		strModuleIDs = strModuleIDs & rsModule("ModuleID").value & ","
		rsModule.MoveNext
	loop
	if right(strModuleIDs, 1) = "," then
		strModuleIDs = left(strModuleIDs, len(strModuleIDs)-1)
	end if
		set  oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsGEOs is nothing then
		response.write "<tr><td><strong>Error for AMO_HUB_Report</strong></td></tr>"
		exit function
	end if
	set oRsRegions = oSvr.AMO_SCM_Regions_Search(Application("REPOSITORY"), strModuleIDs,SCMID)
	if oRsRegions is nothing then
		response.write "<tr><td><strong>Error for AMO_HUB_Report</strong></td></tr>"
		exit function
	end if
	
			
		'Row 1; Heading
		Response.Write  "<tr class='tblrow-pulsar'>" & vbCrLf
		Response.Write  "<th align='center' valign='bottom' rowspan=3>PID/PN</th>" & vbCrLf
		Response.Write  "<th align='center' valign='bottom' rowspan=3>Market Description (64 characters)</th>" & vbCrLf
		Response.Write  "<th align='left' rowspan=3 valign=middle  rowspan=2>Info Delta</th>" & vbCrLf
		Response.Write  "<th align='left' rowspan=3 valign=middle  rowspan=2>Localization(Y/N)</th>" & vbCrLf
		Response.Write  "<th align='center' valign='bottom' rowspan=3>PHweb (General) Availability (GA)</th>" & vbCrLf
		Response.Write  "<th align='center' valign='bottom' rowspan=3>End of Manufacturing (EM)</th>" & vbCrLf
		

		'display all the regions
		'first is the GEO for the region
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strGEO = ""
		intCount = 0
		intIndex = 0
		strAllRegionIDs = ""
		dim ballcountryshown
		do while (not oRsAllRegions.EOF)
			oRsRegions.filter = "RegionID=" & oRsAllRegions("RegionID").value & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
			if oRsRegions.RecordCount > 0 then
				if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
					if trim(oRsAllRegions("CountryGEO").value) <> strGEO and strGEO <> "" then
						Response.Write "<th align='center' valign='bottom' colspan=" & cstr(intCount) & ">" & strGEO & "</th>" & vbCrLf 
						intCount = 0
						intIndex = intIndex + 1
					end if
					strGEO = trim(oRsAllRegions("CountryGEO").value)
					strRegion = trim(oRsAllRegions("Region").value)
					intCount = intCount + 1
					strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
				end if
			end if
			oRsAllRegions.MoveNext
		loop
		if intCount > 0  then
			Response.Write "<th align='center' valign='bottom' colspan=" & cstr(intCount) & ">" & strGEO & "</th>" & vbCrLf 
		end if
		
		
		Response.Write  "</tr><tr class='tblrow-pulsar'>" & vbCrLf  
	
	
		'second is the region name
		strAllRegionIDs = ""
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		do while (not oRsAllRegions.EOF)
			oRsRegions.filter = "RegionID=" & oRsAllRegions("RegionID").value & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
			if oRsRegions.RecordCount > 0 then
				if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
					Response.Write  "<th align=left width=10 >" & trim(oRsAllRegions("Region").value) & "</th>" & vbCrLf  
				end if
				strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
			end if
			oRsAllRegions.MoveNext
		loop
		
		Response.Write  "</tr><tr class='tblrow-pulsar'>" & vbCrLf  
		
		'third is the HP Dash Code for the region
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strRegionIDsBackup = ""
		intCount = 0
		strAllRegionIDs = ""
		do while (not oRsAllRegions.EOF)
			oRsRegions.filter = "RegionID=" & oRsAllRegions("RegionID").value & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
			if oRsRegions.RecordCount > 0 then
				if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
					Response.Write  "<th align=left width=10 >" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</th>" & vbCrLf  
					strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
					strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
					intCount = intCount + 1
				end if
			end if
			oRsAllRegions.MoveNext
		loop
						
		Response.write "</tr>" 
		
	
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value			
			Response.Write  "<tr bgcolor='#FFCC99' style='color:blue;text-decoration:none'>"
			Response.write "<td colspan='" & intCompatibility + intCount + 10 & "'><strong>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</strong></td>" & vbCrLf  
   			Response.Write  "</tr>" 

			do while ((not rsModule.EOF) and (rsModule.Fields("CategoryID").value = strCurrentSection))
				'write each option
				'Use bgcolor='#FFFF99' for highlighted cells to indicate changes
				nModuleID = rsModule.Fields("ModuleID").Value
				strGEOIDs = strGEOIDsBackup
				strRegionIDs = strRegionIDsBackup
				if rsModule.Fields("AMOStatusID").value = Application("AMO_RASREVIEW") or rsModule.Fields("AMOStatusID").value = Application("AMO_RASUPDATE") then
					bLocked = True
				else
					bLocked = False
				end if
				
				
				if clng(rsModule.Fields("ReportChangeMask").value) > 0 Then
					sDelta = "C"
				else
					sDelta = GetDelta(oRsRegions, nModuleID)
				end if

				Response.Write  "<tr style='color:black;text-decoration:none'>" & vbCrLf  
				Response.Write  "<td>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  & "</td>" & vbCrLf  
				Response.Write  "<td>" &  Server.HTMLEncode(rsModule.Fields("Description").Value) & "</td>" & vbCrLf  
				Response.Write  "<td>" & sDelta & "&nbsp;</td>"	 & vbCrLf  
				' Only check base unit only
				
				if rsModule.Fields("RegionID").Value = Application("AMO_GLOBAL_REGIONID") Then
				
					if rsModule.Fields("Localized").Value = 0 And cbool(rsModule.Fields("SCM_Localized").Value) = False Then
						response.write " <td>N</td>" & vbCrLf  
					else
						response.write "<td>Y</td>" & vbCrLf  	
					end if
				else
					Response.Write   "<td>&nbsp;</td>"
				end if	
				
			
				oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID = " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
				
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					'arrDate = Split(Trim(rsModule.Fields("BOMRevADate").Value),"/")
					Response.Write  "<td>" &  rsModule.Fields("BOMRevADate").Value & "&nbsp;</td>" & vbCrLf
				else
					Response.Write   "<td>&nbsp;</td>"

				end if 
				
				
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					'arrDate = Split(Trim(rsModule.Fields("RASDiscontinueDate").Value),"/")
					Response.Write  "<td>" &  rsModule.Fields("RASDiscontinueDate").Value  & "&nbsp;</td>" & vbCrLf
				else
					Response.Write   "<td>&nbsp;</td>"
				end if 
				
				
				oRsRegions.filter = ""		

				while instr(strRegionIDs, ",") > 0
					strRegionID = left(strRegionIDs, instr(strRegionIDs, "|")-1)
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID=" & strRegionID & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
					if oRsRegions.RecordCount > 0  then
						if oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write "<td" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & ">GA: " & oRsRegions("BOMRevADate").value & " - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</td>" & vbCrLf	
						elseif  oRsRegions("BOMRevADate").value = "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write "<td" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & ">GA: N/A - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</td>" & vbCrLf
						elseif  oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value = "" then
							response.write " <td" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & ">GA: " & oRsRegions("BOMRevADate").value & " - EM: N/A</td>" & vbCrLf	
						else
							response.write "<td " & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & ">GA: N/A - EM: N/A</td>" & vbCrLf
						end if
					else
						Response.Write  "<td>&nbsp;</td>" 
					end if

					strRegionIDs = mid(strRegionIDs, instr(strRegionIDs, ",")+1)
				wend
				
				
				Response.Write  "</tr>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop
		Response.Write  "</table>" & vbCrLf
		
		Response.Write  "<table>" & vbCrLf
		Response.Write  "<tr><td>&nbsp;</td></tr>" 
		Response.Write  "<tr><td>&nbsp;</td></tr>" 
		Response.Write  "<tr><td><strong>Legend:</strong></td></tr>" 
		Response.Write  "<tr><td><strong>Hub</strong></td></tr>" 
		Response.Write  "<tr bgcolor='#00FF00'><td>Indicates Hub SKU</td></tr>" 
		Response.Write  "<tr><td><strong>Info Delta</strong></td></tr>" 
		Response.Write  "<tr><td>C: Changed Since Last Published SCM</td></tr>" 
		
		
		oRsGEOs.Close
		set oRsGEOs = nothing
		oRsRegions.Close
		set oRsRegions = nothing
		oRsAllRegions.Close
		set oRsAllRegions = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function 

sub SendStatusChangeEmail (bBulkStatusChange, bStatusChanged, byval intStatus, byval strModuleID, byval strRegionID)
	
	dim sHTML
	dim oErr, oNotification, oRsReturn, oRsToUsers, oRsUC, oGeneral, oMessage
	dim sFilter, sErr, oSvr
	dim bMatch
	dim strSubject, strTo, strDivision, strModUser, x
	dim arrDivision, arrModUser
	
	'and M.StatusID = " & intStatus
	
	sFilter = " and (R.SCMID = 1 Or R.SCMID = null) and O.SCMID = 1 " 
	
	if bStatuschanged then 
		sFilter = sFilter & " and M.FeatureID = " & strModuleID & " and R.RegionID = " & strRegionID 
	end if
	
	if bBulkStatusChange then
		sFilter = sFilter & " and M.FeatureID IN (" & strModuleID & ") and R.RegionID IN (" & strRegionID & ")"
	end if
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
	set oSvr = New ISAMO	
	set oRsReturn = oSvr.AMOModule_Search(Application("REPOSITORY"), sFilter,"", null, null)
	if (oRsReturn is nothing) then
		Response.Write("Empty Recordset 1: oRsReturn")
		Response.End
		set oRsReturn = nothing
	else
			
		'Create Object for admin group email
		'set oNotification = Server.CreateObject("JF_S_Notification.ISNotification")
		'set oNotification = New ISNotification

		'Create Object for updater and creator
		'set oGeneral = Server.CreateObject("JF_S_General.ISGeneral")
		set oGeneral = New ISGeneral
					
		'oRsReturn.MoveFirst
		do while (not oRsReturn.EOF)
						
			strTo = ""			
	
			strSubject = "Pulsar AMO Feature - Part# " & oRsReturn.Fields("BluePN") & " has been rejected during RAS Review."
			
			'get Updater and Creator email
			strModUser = oRsReturn.Fields("CreatorID") & "," & oRsReturn.Fields("UpdaterID")
			arrModUser = Split(strModUser,",")
			For x = 0 To UBound(arrModUser)
				set oRsUC = oGeneral.GetOneUserInfo(Application("REPOSITORY"), CLng(arrModUser(x)))
				if (oRsUC is nothing) then
					Response.Write("Empty Recordset: oRsUC")
					set oRsUC = nothing
				else
					do while not oRsUC.EOF						
						if strTo <> "" then strTo = strTo & ";"
						if InStr(strTo, trim(oRsUC("Email"))) = 0 then
							strTo = strTo & trim(oRsUC("Email"))
						end if	
						oRsUC.MoveNext
					loop
				end if
			Next

			'get Ras Admin Email List - Pulsar doesn't use mailinglist or owned by so can't determine admin
			'strDivision = oRsReturn.Fields("DivisionIDs")
			'arrDivision = Split(strDivision,",")
			'For x = 0 To UBound(arrDivision)        
				'set oRsToUsers = oNotification.ViewUsersByEventRs(Application("REPOSITORY"), "AMO RAS REVIEW", arrDivision(x), -1 )
				'if (oRsToUsers is nothing) then
					'Response.Write("Empty Recordset: oRsToUsers")
					'set oRsToUsers = nothing
				'else
					'do while not oRsToUsers.EOF
						'if strTo <> "" then strTo = strTo & ";"
						'if InStr(strTo, trim(oRsToUsers("Email"))) = 0 then
							'strTo = strTo & trim(oRsToUsers("Email"))
						'end if							
						'oRsToUsers.MoveNext
					'loop
				'end if			
			'Next
				
			'Build body of email
			sHTML = "<p><font size =""2"" face=""Arial"" style=""font-weight:bolder;"">Pulsar - AMO Feature Rejected</font>&nbsp;</p>" & vbCrLf
			sHTML = sHTML & "<p><font size =""2"" face=""Arial"" style=""font-weight:bolder;"">Short Description:</font>&nbsp;" & vbCrLf
			sHTML = sHTML & "<font size =""2"" face=""Arial"">" & vbCrLf
			sHTML = sHTML & oRsReturn.Fields("Description") & "</font></p>" & vbCrLf
			sHTML = sHTML & "<p><font size =""2"" face=""Arial"" style=""font-weight:bolder;"">Part Number:</font>&nbsp;" & vbCrLf
			sHTML = sHTML & "<font size =""2"" face=""Arial"">" & vbCrLf
			sHTML = sHTML & oRsReturn.Fields("BluePN") & "</font></p>" & vbCrLf
			sHTML = sHTML & "<p><font size =""2"" face=""Arial"" style=""font-weight:bolder;"">Last Updated By:</font>&nbsp;" & vbCrLf
			sHTML = sHTML & "<font size =""2"" face=""Arial"">" & vbCrLf
			sHTML = sHTML & oRsReturn.Fields("Updater") & "</font></p>" & vbCrLf

			if oRsReturn.Fields("RASComment") = "" then
				sHTML = sHTML & "<p><font size =""2"" face=""Arial"">" & vbCrLf
				sHTML = sHTML & "There is no Comments from RAS.</font></p>" & vbCrLf
			else
				sHTML = sHTML & "<p><font size =""2"" face=""Arial""><strong>" & vbCrLf
				sHTML = sHTML & "Comments from RAS</strong></font></p>" & vbCrLf
			
				sHTML = sHTML & "<p><font size =""1"" face=""Arial"">" & vbCrLf
				sHTML = sHTML & oRsReturn.Fields("RASComment") & "</font></p>"	
			end if	
			
			On Error Resume Next
	
			'use pulsar email wrapper
			if sHTML <> "" and strTo <> "" then
				Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	'			Set oMessage.Configuration = Application("CDO_Config")	
				oMessage.From = "pulsar.support@hp.com"					
				oMessage.To = strTo
				oMessage.Subject = strSubject			
				oMessage.HTMLBody = sHTML
				oMessage.DSNOptions = cdoDSNFailure
				oMessage.Send 
				Set oMessage = Nothing 	
			end if
						
			oRsReturn.MoveNext
		loop
		'set oNotification = nothing
		set oGeneral = nothing
	end if
	
	set oSvr = nothing
end sub

function makeDateFormat(strFullDate)
	dim strDate
	strDate = trim(cstr(strFullDate))
	if strDate <> "" then
		makeDateFormat = PadZero(day(strDate),2) & "-" & MonthName(Month(strDate),true) & "-" & right(year(strDate), 2)
	else
		makeDateFormat = strFullDate
	end if
end function

function stripFancyChars(strString)
	'handle all these locations:
	'Broadcom 5751Gigabit ™ PCIe
	'Broadcom NIC ®
	'Audio®
	'Intel2d™
	'Intel HDA 2 channel ® audio ™

	'174 = registered symbol
	'153 = trademark symbol
	'169 = copyright symbol

	'Just to do it the easy way, replace those characters with empty strings. Then replace double spaces with one space.
	'The only flaw in this logic is that if someone intentionally wants two spaces in a description, they're not going to get it.
	'That should be an unlikely chance.
	
	stripFancyChars = replace(replace(replace(replace(replace(strString, chr(174), ""), chr(153), ""), chr(169), ""), "  ", " "), vbCrLf, " ")
end function

function tomorrow()
	dim thedate
	dim arrMonths
	arrMonths = array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
	
	thedate = dateAdd("d", 1, date)
	tomorrow = PadZero(datePart("d", thedate),2) & "-" & arrMonths(datePart("m", thedate)-1) & "-" & right(datePart("yyyy", thedate), 2)
end function
%>
