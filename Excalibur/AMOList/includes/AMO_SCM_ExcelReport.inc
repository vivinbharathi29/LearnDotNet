<%


function WriteAMO_SCM_ReportGridHTML (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs)
	dim strGEO, strRegion, strDivisionIDs, strModuleIDs
	dim strGEOID, strGEOIDs, strRegionIDs, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay
	dim nModuleID
	dim intCount
	dim bDisplayed, bLocked
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr, sErr
	dim strClass, sPD, sChange,strPlatform, strMarketingName
	dim intPlatformID
	dim nChassisID, numDL
	dim oRsPlatforms
	dim bPlatformAssigned
	dim PlatformCt, i
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup
	if rsModule.recordcount = 0 then
		response.write "No After Market Option found that matches filter"
		exit function
	else
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	
	set oSvr = New ISAMO
	set oRsAllRegions = oSvr.AllRegions_Search(Application("REPOSITORY"), strDivisionIDs, null)
	if oRsAllRegions is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
		exit function
	end if
	if oRsAllRegions.recordcount> 0 then
		oRsAllRegions.Sort = "SortOrder ASC, Region ASC"
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

	set oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsGEOs is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
		exit function
	end if
	set oRsRegions = oSvr.AMORegions_Search(Application("REPOSITORY"), strModuleIDs,SCMID)
	if oRsRegions is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
		exit function
	end if
	
	set oRsPlatforms = oSvr.AMOPlatforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs,SCMID, 0)
	if oRsPlatforms is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		Response.Write(sErr)
		Response.End()
		exit function
	end if
	
	Response.Write  "<pre><table id='oTable' name='oTable' align='left' border='1' "
		'Hardcoded info
		Response.Write  "<tr class='clsCatHeading'><td colspan=5 class='xlHeadingWhiteLarge'>AMO</td></tr>" 
		Response.Write  "<tr class='clsCatHeading'><td colspan=5 class='xlHeadingWhitemedium'>" & "Published: " & PublishDate & "</td></tr>" 
		
		Response.Write  "<tr class='clsCatHeading'><td colspan=5 class='xlHeadingWhitemedium'>" & "Today's Date: " & cstr(now()) &  " (All Dates mm/dd/yyyy)" & "</td></tr>" 
			
		'Row 1; Heading
		Response.Write  "<tr class='clsHeading' style='color:black;text-decoration:none'>"
		Response.Write  "<th rowspan=2 class='xlHeading' >Part Number</th>"
		Response.Write  "<th rowspan=2 class='xlHeading' >Option (Short Description)</th>"
		Response.Write  "<th  rowspan=2 class='xlHeadingRotate'>Replacement</th>"
		'platform info
		PlatformCt = 0
		strPlatformIDs =""
		strtempPlatformID = ""
		if oRsPlatforms.recordcount> 0 then
			oRsPlatforms.MoveFirst
		    strMarketingName = ""
			strPlatform = ""
			bDisplayed = true
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
				    if trim(strMarketingName) <> "" then
						Response.Write  "<th rowspan=2 class='xlHeadingRotatethin'>" & strMarketingName & "</th>" & vbCrLf
					else
						Response.Write  "<th rowspan=2 class='xlHeadingRotatethin'>N/A</th>" & vbCrLf
					end if
					PlatformCt = PlatformCt + 1
					bDisplayed = true
					strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
				end if
				strPlatform = oRsPlatforms("FullName").value
				strMarketingName = oRsPlatforms("MarketingName").value
				strtempPlatformID = oRsPlatforms("PlatformID").value
				bDisplayed = false
				oRsPlatforms.MoveNext
			loop
			if not bDisplayed then
				if trim(strMarketingName) <> "" then
					Response.Write  "<th rowspan=2 class='xlHeadingRotatethin'>" & strMarketingName & "</th>" & vbCrLf
				else
					Response.Write  "<th rowspan=2 class='xlHeadingRotatethin'>N/A</th>" & vbCrLf
				end if
				PlatformCt = PlatformCt + 1
				strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
			end if
			strPlatformIDs_backup = strPlatformIDs
		end if
		'end if platform headers
		Response.Write  "<th rowspan=2 class='xlHeading' >Select Availability (SA)</th>"
		Response.Write  "<th rowspan=2 class='xlHeading' >PHweb (General) Availability (GA)</th>"
		Response.Write  "<th rowspan=2 class='xlHeading' >End of Manufacturing (EM)</th>"
		Response.Write  "<th rowspan=2 class='xlHeading' >Net Weight</th>"
		Response.Write  "<th rowspan=2 class='xlHeadingRotateBottomborder' >Regions</th>"
		'display all the GEOs
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		strGEO = ""
		strGEOID = ""
		strGEOIDsBackup = ""
		bDisplayed = True
		do while (not oRsAllRegions.EOF)
			if oRsAllRegions("CountryGEO").value <> strGEO and strGEO <> "" then
				if ucase(strGEO) <> "GLOBAL" then	'don't display GLOBAL
					Response.Write  "<th rowspan=2 class='xlHeadingRotateBottomborder'>" & strGEO & "</th>" & vbCrLf
					strGEOIDsBackup = strGEOIDsBackup & strGEOID & ","
					bDisplayed = True
				end if
			end if
			strGEO = oRsAllRegions("CountryGEO").value
			strGEOID = oRsAllRegions("CountryGEOID").value
			bDisplayed = False
			oRsAllRegions.MoveNext
		loop
		if not bDisplayed then
			Response.Write  "<th rowspan=2 class='xlHeadingRotateBottomborder'>" & strGEO & "</th>" & vbCrLf
			strGEOIDsBackup = strGEOIDsBackup & strGEOID & ","
		end if
		'strGEOIDsBackup now has all the GEO IDs that were displayed

		'display all the regions
		'first is the GEO for the region
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strGEO = ""
		intCount = 0
		do while (not oRsAllRegions.EOF)
			if trim(oRsAllRegions("CountryGEO").value) <> strGEO and strGEO <> "" then
				if ucase(strGEO) = "GLOBAL" then
					Response.Write  "<th rowspan=2 id='p' align='left' class='xlHeadingRotate'>" & strRegion & "</th>" & vbCrLf
				else
					Response.Write  "<th colspan=" & cstr(intCount) & " class='xlHeadingLeftMedium' >" & strGEO & "</th>" & vbCrLf
				end if
				intCount = 0
			end if
			strGEO = trim(oRsAllRegions("CountryGEO").value)
			strRegion = trim(oRsAllRegions("Region").value)
			intCount = intCount + 1
			oRsAllRegions.MoveNext
		loop
		if intCount > 0 then
			Response.Write  "<th colspan=" & cstr(intCount) & " class='xlHeadingLeftMedium' >" & strGEO & "</th>" & vbCrLf
		end if

		'second is the region name
		Response.Write  "<tr class='clsHeading' style='color:black;text-decoration:none'>" & vbCrLf
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		do while (not oRsAllRegions.EOF)
			if trim(ucase(oRsAllRegions("CountryGEO").value)) <> "GLOBAL" then
				Response.Write  "<th id='p' align='left' class='xlHeadingRotate'>" & trim(oRsAllRegions("Region").value) & "</th>" & vbCrLf
			end if
			oRsAllRegions.MoveNext
		loop
		Response.Write  "</tr>" & vbCrLf
		'third is the HP Dash Code for the region
		'moved these codes under the 'BOM Eng' to make them same line
	'	Response.Write  "<tr class='clsHeading' style='color:black;text-decoration:none'>" & vbCrLf
	'	oRsAllRegions.MoveFirst
	'	strRegionIDsBackup = ""
	'	do while (not oRsAllRegions.EOF)
	'		Response.Write  "<th class='xlHeadingBottomBorder'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</th>" & vbCrLf
	'		strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
	'		oRsAllRegions.MoveNext
	'	loop
		'strRegionIDsBackup now has all the Region IDs & HP Dash Codes that were displayed
	'	Response.Write  "</tr>" & vbCrLf
	
		'BOM ENg info header row
		Response.Write  "<tr><th class='xlHeadingyellow' >BOM <Br>Eng</th><td class='xlHeadingyellow' >BOM <br>Eng</td><td class='xlHeadingyellow'>&nbsp;</td> "
		do while  i < PlatformCt
			Response.write "<td class='xlHeadingyellow' ><p>BOM</p>Eng</td>"
			i=i+1
		loop
		Response.write "<td class='xlHeadingyellowgreen' >Marketing <Br>Ops</td>"
		Response.write "<td class='xlHeadingaquamarine' >Supply <br>Chain</td>"
		Response.write "<td class='xlHeadingyellowgreen' >Marketing <br>Ops</td>"
		Response.write "<td class='xlHeadingaquamarine' >Supply <br>Chain</td>"
		'Response.write "<td class='xlHeadingWhite'>&nbsp;</td>"
		Response.write "<td class='xlHeadingWhite'>&nbsp;</td>"
		Response.write "<td class='xlHeadingWhite' >&nbsp;</td>"
		Response.write "<td class='xlHeadingWhite' >&nbsp;</td>"
		Response.write "<td class='xlHeadingWhite'>&nbsp;</td>"
		Response.write "<td class='xlHeadingWhite'>&nbsp;</td>"
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strRegionIDsBackup = ""
		do while (not oRsAllRegions.EOF)
			Response.Write  "<th class='xlHeadingBottomBorder'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</th>" & vbCrLf
			strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
			oRsAllRegions.MoveNext
		loop
				
		Response.write " </tr>" 
	
		'end of BOM ENg info header row
	
		'response.write strPlatformIDs_backup
		'	response.end
	
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value
			Response.Write  "<tr class='clsCatHeading'><td colspan=" & cstr(PlatformCt +7) & " class='xlCategory'>" & _
			rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value & "</td>" & vbCrLf
			Response.Write  "</tr>" & vbCrLf

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

				Response.Write  "<tr mid='" & nModuleID & "' cid='" & strCurrentSection & "' align='center'>"
				Response.Write  "<td class='xlcontentleft'>" & rsModule.Fields("BluePN").Value &"&nbsp;" & "</td>"
				Response.Write  "<td class='xlcontentleft'>" & rsModule.Fields("ShortDescription").Value & "&nbsp;" & "</td>"
				Response.Write  "<td class='xlcontentleft'>" & rsModule.Fields("Replacement").Value & "&nbsp;" & "</td>"	
			'	platform info
			'	if oRsPlatforms.recordcount> 0 then
			'		oRsPlatforms.MoveFirst
			'		strPlatform = ""
			'		bPlatformAssigned = false
			'		bDisplayed = true
			'		do while not oRsPlatforms.EOF
			'			if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
			'				if bPlatformAssigned then
			'					response.write "<td class='xl24' >ü</td>"
			'				else
			'					response.write "<td class='xlcontent'>" & "&nbsp;" & "</td>"
			'				end if
			'				bDisplayed = true
			'				bPlatformAssigned = False
			'			end if
			'			strPlatform = oRsPlatforms("FullName").value
			'			if oRsPlatforms("ModuleID") = nModuleID then
			'				bPlatformAssigned = True
			'			end if
			'			bDisplayed = false
			'			oRsPlatforms.MoveNext
			'		loop
			'		if not bDisplayed then
			'			if bPlatformAssigned then
			'					response.write "<td class='xl24' >ü</td>"
			'				else
			'					response.write "<td class='xlcontent'>" & "&nbsp;" & "</td>"
			'			end if
			'		end if
			'	end if	
			'	end of dispalying platform info
			'New way to display platform info-07/01/2003
			strPlatformIDs =strPlatformIDs_backup
			while instr(strPlatformIDs, ",") > 0
					strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
					oRsPlatforms.filter = "ModuleID = " & cstr(nModuleID) & " and platformID=" & strtempPlatformID
					if oRsPlatforms.RecordCount > 0 then
						response.write "<td class='xl24' >ü</td>"					
					else					
						Response.Write  "<td class='xlcontent'>&nbsp;</td>" & vbCrLf
					end if
					strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
			wend
			
		'	response.write strPlatformIDs_backup
		'	response.end
			
			
			
				if len(rsModule.Fields("CPLBlindDate").Value)>0 then
					Response.Write  "<td class='xlContent'>" & rsModule.Fields("CPLBlindDate").Value & "&nbsp;" & "</td>"
				else
					Response.Write  "<td class='xlcontent'>" & "&nbsp;" & "</td>"
				end if 
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					Response.Write  "<td class='xlContent'>" & rsModule.Fields("BOMRevADate").Value & "&nbsp;" & "</td>"
				else
					Response.Write  "<td class='xlcontent'>" & "&nbsp;" & "</td>"
				end if 
				
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					Response.Write  "<td class='xlContent'>" & rsModule.Fields("RASDiscontinueDate").Value & "&nbsp;" & "</td>"
				else
					Response.Write  "<td class='xlcontent'>" & "&nbsp;" & "</td>"
				end if 
				
				Response.Write  "<td class='xlcontent'>" & rsModule.Fields("NetWeight").Value & "&nbsp;" & "</td>"
			
				Response.Write  "<td class='xlcontent'>" & "&nbsp;" & "</td>"
			
			
				while instr(strGEOIDs, ",") > 0
					strGEOID = left(strGEOIDs, instr(strGEOIDs, ",")-1)
					oRsGEOs.filter = "ModuleID = " & cstr(nModuleID) & " and GEOID=" & strGEOID
					if oRsGEOs.RecordCount > 0 then
						'there is a matching date
						Response.Write  "<td class='xlcontent'>" & oRsGEOs("MRDate").value & "&nbsp;" & "</td>" & vbCrLf
					
					else
					
						Response.Write  "<td class='xlcontent'>&nbsp;</td>" & vbCrLf
					
					end if

					strGEOIDs = mid(strGEOIDs, instr(strGEOIDs, ",")+1)
				wend

				while instr(strRegionIDs, ",") > 0
					strRegionID = left(strRegionIDs, instr(strRegionIDs, "|")-1)
					strHPDASHCode = mid(strRegionIDs, instr(strRegionIDs, "|")+1, instr(strRegionIDs, ",")-instr(strRegionIDs, "|")-1)
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID=" & strRegionID
					if oRsRegions.RecordCount > 0 then
						'there is a matching entry, put a checkmark
						Response.Write  "<td class='xl24' >ü</td>" & vbCrLf
					else
						Response.Write  "<td class='xlcontent'>" & "&nbsp;" & "</td>" & vbCrLf
					end if

					strRegionIDs = mid(strRegionIDs, instr(strRegionIDs, ",")+1)
				wend

				Response.Write  "</tr>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop
		Response.Write  "</table></pre>" & vbCrLf
		oRsGEOs.Close
		set oRsGEOs = nothing
		oRsRegions.Close
		set oRsRegions = nothing
		oRsAllRegions.Close
		set oRsAllRegions = nothing
		oRsPlatforms.Close
		set oRsPlatforms = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function


function WriteAMO_SCM_CompareHTML (byval rsSCMCompare, byval RsTwoAMOPublishs, byval CompareCurrent)
	dim strtempModuleName
	dim iChangeTypeOrder
	strtempModuleName = ""
	iChangeTypeOrder = 0
	Response.Write  "<pre><table id='oTable' name='oTable' align='left' border='1' "
		if len(CompareCurrent) > 0 then
			Response.Write  "<tr class='clsCatHeading'><td colspan=4 class='xlHeadingWhitemedium'>Module changes between one publish and current data (Base Module Only)</td></tr>" 
		else
			Response.Write  "<tr class='clsCatHeading'><td colspan=4 class='xlHeadingWhitemedium'>Module changes between two publishes (Base Module Only)</td></tr>" 
		end if
		if RsTwoAMOPublishs.recordcount> 0 then
			Response.Write  "<tr>"
			Response.Write  "<td class='xlContent' >&nbsp;</td>"
			Response.Write  "<td class='xlContentleft' >Date Published</td>"
		
			do while (not RsTwoAMOPublishs.EOF)	
		
				Response.Write  "<td class='xlHeadingWhite' >" & "&nbsp;" &RsTwoAMOPublishs.Fields("TimeCreated").Value & "</td>"
						
				RsTwoAMOPublishs.MoveNext
					
			loop
			Response.Write  "</tr >"
			RsTwoAMOPublishs.movefirst
			Response.Write  "<tr >"
			Response.Write  "<td class='xlContent' >&nbsp;</td>"
			Response.Write  "<td class='xlContentleft' >Publisher</td>"
		
			do while (not RsTwoAMOPublishs.EOF)	
		
				Response.Write  "<td class='xlHeadingWhite' >" & RsTwoAMOPublishs("Creator") & " </td>"
						
				RsTwoAMOPublishs.MoveNext
					
			loop
			Response.Write  "</tr >"
		
		end if 
		Response.Write  "<tr class='clsHeading' style='color:black;text-decoration:none'>"
		Response.Write  "<th class='xlHeading' >Change Type</th>"
		Response.Write  "<th class='xlHeading' >Changed Item</th>"
		Response.Write  "<th class='xlHeading' >Changed From </th>"
		Response.Write  "<th class='xlHeading' >Changed To </th>"
		Response.Write  "</tr >"
		rsSCMCompare.MoveFirst
		
		do while (not rsSCMCompare.EOF)
		
			if iChangeTypeOrder <> rsSCMCompare("ChangeTypeOrder") then
				Response.Write  "<tr >"
				Response.Write  "<td class='xlChangeType' colspan = 4>" 
				if rsSCMCompare("ChangeTypeOrder") = 1 then
					Response.Write  "New Module Options"
				elseif  rsSCMCompare("ChangeTypeOrder") = 2 then
					Response.Write  "Changed Module Options"
				else
					Response.Write  "Deleted Module Options"
				end if 
				Response.Write  "</td>"
				Response.Write  "</tr >"
			end if 
			if strtempModuleName <> rsSCMCompare("ModuleName") then
				Response.Write  "<tr >"
				Response.Write  "<td class='xlCategoryGroup' colspan = 4>" & rsSCMCompare.Fields("ModuleName").Value &"&nbsp;" & "</td>"
				Response.Write  "</tr >"
			end if 
			
			Response.Write  "<tr >"
			Response.Write  "<td class='xlContent'>" & rsSCMCompare.Fields("ChangeType").Value &"&nbsp;" & "</td>"
			Response.Write  "<td class='xlcontentleft'>" & rsSCMCompare.Fields("Item").Value &"&nbsp;" & "</td>"
			Response.Write  "<td class='xlcontentleft'>" & rsSCMCompare.Fields("ChangeFrom").Value & "&nbsp;" & "</td>"
			Response.Write  "<td class='xlcontentleft'>" & rsSCMCompare.Fields("ChangeTo").Value & "&nbsp;" & "</td>"	
			Response.Write  "</tr >"
			strtempModuleName = rsSCMCompare("ModuleName")
			iChangeTypeOrder = rsSCMCompare("ChangeTypeOrder")
			rsSCMCompare.MoveNext
				
			
		loop
		Response.Write  "<tr><td class='xlcontentleft' colspan = 4>&nbsp;</td></tr >"
		
		Response.Write  "<tr><td class='xlcontentleft' colspan = 4><b>Legend:</b></font><br>"
		Response.Write "N:New<Br>C:Changed<Br>O:Obsolete<Br></td></tr >"
		
	
		
		Response.Write  "</table></pre>" & vbCrLf





end function




function Write_AVAMO_MAtrix (byval rsAVAMO, byval rsModules)
	dim PlatformCt, i
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup
	dim strCurrentSection
	dim nModuleID
	dim strTempPArtNo
		
	Response.Write  "<pre><table id='oTable' name='oTable' align='left' border='1' "
		
			Response.Write  "<tr><th class='xlHeading' >Light Description</th>"
	
		'platform info
		PlatformCt = 0
		strPlatformIDs =""
		strtempPlatformID = ""
		if rsAVAMO.recordcount> 0 then
			rsAVAMO.sort = "PlatformName, platformID"
			rsAVAMO.MoveFirst
		
			strPlatform = ""
			bDisplayed = true
			do while (not rsAVAMO.EOF)
				if rsAVAMO("PlatformName").value <> strPlatform then
					Response.Write  "<th class='xlHeading'>" & rsAVAMO("PlatformName").value & "</th>" & vbCrLf
					PlatformCt = PlatformCt + 1
					bDisplayed = true
					strPlatformIDs = strPlatformIDs & rsAVAMO("PlatformID").value & ","
				end if
				strPlatform = rsAVAMO("PlatformName").value
				strtempPlatformID = rsAVAMO("PlatformID").value
				bDisplayed = false
				rsAVAMO.MoveNext
			loop
			
			strPlatformIDs_backup = strPlatformIDs
		end if
		Response.Write  "</tr >"
		'end if platform headers
		
	
		'Module Data
		if rsAVAMO.recordcount> 0 then
			rsAVAMO.sort = "CategoryName asc, ModuleName asc"
			rsAVAMO.MoveFirst
			
		end if 	
		
		strCurrentSection = ""
		do while (not rsmodules.EOF)
		'write Category Heading
			nModuleID = rsModules.Fields("ModuleID").Value
			if strCurrentSection <> rsModules.Fields("CategoryID").Value then
			
				Response.Write  "<tr ><td class='xlCategoryGroup'>" & rsModuleS.Fields("Categoryname").Value & "</td>" & vbCrLf
				Response.Write  "</tr>" & vbCrLf
			end if 
			Response.Write  "<tr  align='center'>"
			Response.Write  "<td class='xlcontentleft'>" & rsModules.Fields("ModuleName").Value & "</td>"
			
			
			
			strPlatformIDs =strPlatformIDs_backup
			while instr(strPlatformIDs, ",") > 0
					strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
					rsAVAMO.filter = "ModuleID = " & cstr(nModuleID) & " and platformID=" & strtempPlatformID
					strTempPArtNo = ""
					if rsAVAMO.RecordCount > 0 then
						response.write "<td class='xlcontent' >"
						do while not rsAVAMO.eof
							strTempPArtNo = strTempPArtNo &  trim(rsAVAMO.Fields("PartNo").value) & "<br>"					
							rsAVAMO.movenext
						loop
						if len(strTempPArtNo) > 4 then
							strTempPArtNo = left(strTempPArtNo, len(strTempPArtNo)-4)
						end if 
						response.write strTempPArtNo
						response.write "</td>"
					else					
						Response.Write  "<td class='xlcontent'>&nbsp;</td>" & vbCrLf
					end if
					strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
			wend
					
				Response.Write  "</tr>" & vbCrLf
				strCurrentSection = rsModules.Fields("CategoryID").Value
					
			rsmodules.movenext
			
		loop
		Response.Write  "</table></pre>" & vbCrLf
		
		rsAVAMO.Close
		set rsAVAMO = nothing
		
		rsModules.Close
		set rsModules = nothing
	
end function




function WriteAMO_SCM_ReportGridXML (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs, byVal sUserBusSeg)
	dim strGEO, strRegion, strDivisionIDs, strModuleIDs
	dim strGEOID, strGEOIDs, strRegionIDs, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay
	dim nModuleID
	dim intCount
	dim bDisplayed, bLocked
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr
	dim strClass, sPD, sChange, sErr, strPlatform, strMarketingName, strRevision
	dim intPlatformID
	dim nChassisID, numDL
	dim oRsPlatforms
	dim bPlatformAssigned
	dim PlatformCt, i, arrDate
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup
	if rsModule.recordcount = 0 then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>No After Market Option found that matches filter</Data></Cell></Row>"
		exit function
	else	
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	strRevision = rsModule("Revision").value
	
	set oSvr = New ISAMO
	
	
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

	
	set oRsPlatforms = oSvr.AMO_SCM_Platforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs,SCMID)
	if oRsPlatforms is nothing then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	
		
		PlatformCt = 0
		strPlatformIDs =""
		strtempPlatformID = ""
		if oRsPlatforms.recordcount> 0 then
			oRsPlatforms.MoveFirst
		
			strPlatform = ""
			bDisplayed = true
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
					PlatformCt = PlatformCt + 1
					bDisplayed = true
					
				end if
				strPlatform = oRsPlatforms("FullName").value
				bDisplayed = false
				oRsPlatforms.MoveNext
			loop
			if not bDisplayed then
				PlatformCt = PlatformCt + 1
				
			end if
			
		end if

		if PlatformCt > 1 then
			Response.write  "<Column ss:AutoFitWidth='0' ss:Width='50.25' ss:Span='" & cstr(PlatformCt-1) & "'/>"
		elseif PlatformCt =1 then
			Response.write "<Column ss:Width='50.25'/>"
		end if 
		
		'add new columns 4/1/2004
		Response.write "<Column ss:Width='100.0'/>"
		Response.write "<Column ss:Width='50.75'/>"
		Response.write "<Column ss:Width='50.75'/>"
		Response.write "<Column ss:Width='50.75'/>"
		Response.write "<Column ss:Width='50.75'/>"
		Response.write "<Column ss:Width='200.0'/>"

		Response.write "<Column ss:Width='81.75'/>"
		Response.write " <Column ss:Width='84'/>"
			
		Response.write "<Column ss:Width='110.25'/>"

		Response.write "  <Column ss:Width='77.25'/>"

		Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='20' ss:StyleID='m28734884'><Data ss:Type='String'>" & sUserBusSeg & " AMO SCM</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" 
		'if trim(strRevision) <> "1" then
			Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Revision: " & strRevision &  "</Data><NamedCell ss:Name='Print_Titles'/></Cell></Row>"
			Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Published: " & PublishDate &  "</Data><NamedCell ss:Name='Print_Titles'/></Cell></Row>"
		'end if
		Response.Write  " <Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" 
			
		'Row 1; Heading
		Response.Write  "<Row ss:Height='15.75' ss:StyleID='s30'>"
		Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734914'><Data ss:Type='String'>Part Number</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734924'><Data ss:Type='String'>Option (Short Description)</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734934'><Data ss:Type='String'>Replacement</Data></Cell>"
		'platform info
		PlatformCt = 0
		strPlatformIDs =""
		strtempPlatformID = ""
		if oRsPlatforms.recordcount> 0 then
			oRsPlatforms.MoveFirst
			strMarketingName = ""
			strPlatform = ""
			bDisplayed = true
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
				  if trim(strMarketingName) <> "" then
					Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>"
				  else
					Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
				  end if
				  
					PlatformCt = PlatformCt + 1
					bDisplayed = true
					strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
				end if
				strPlatform = oRsPlatforms("FullName").value
				strMarketingName = oRsPlatforms("MarketingName").value & " " & oRsPlatforms("BusinessSegs").value
				strtempPlatformID = oRsPlatforms("AliasID").value
				bDisplayed = false
				oRsPlatforms.MoveNext
			loop
			if not bDisplayed then
				if trim(strMarketingName) <> "" then
				  Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>"
				else
				  Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
				end if
				PlatformCt = PlatformCt + 1
				strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
			end if
			strPlatformIDs_backup = strPlatformIDs
		end if
		'end if platform headers
		'new columns 4/1/2004
					
		'Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Rules Syntax</Data></Cell>"
		'Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>AV ID</Data></Cell>"
		'Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Group 1</Data></Cell>"
		'Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Group 2</Data></Cell>"
		'Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Group 3</Data></Cell>"
		Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Rules Description</Data></Cell>"
		Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>Localization</Data></Cell>"
				
		Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28734954'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>"
		Response.Write  " <Cell ss:MergeDown='1' ss:StyleID='m28735036'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28735046'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28735056'><Data ss:Type='String'>Net Weight</Data></Cell>"
		
	
		response.write " </Row>"

		'second is the region name
		dim firstregionrow
		firstregionrow =1
		Response.Write  "	<Row ss:Height='138' ss:StyleID='s30'>"
		
		Response.Write  "</Row>"		
		
		'third is the HP Dash Code for the region
		
		'BOM ENg info header row
		Response.Write  "<Row ss:Height='28.75'>"
		
		Response.write " <Cell  ss:StyleID='m28735502'><Data ss:Type='String'>BOM &#10;Eng</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.write " <Cell ss:StyleID='m28735502'><Data ss:Type='String'>BOM &#10;Eng</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.write " <Cell  ss:StyleID='m28735522'/>"
		
		do while  i < PlatformCt
			Response.write " <Cell ss:StyleID='s52'><Data ss:Type='String'>BOM &#10;Eng</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
			i=i+1
		loop
		'new columns 4/1/2004
		'Response.Write   "<Cell ss:StyleID='s101'/>"
		'Response.Write   "<Cell ss:StyleID='s101'/>"
		'Response.Write   "<Cell ss:StyleID='s101'/>"
		'Response.Write   "<Cell ss:StyleID='s101'/>"
		'Response.Write   "<Cell ss:StyleID='s101'/>"
		Response.Write   "<Cell ss:StyleID='s101'/>"
		Response.Write   "<Cell ss:StyleID='s101'/>"
		
		Response.write "<Cell  ss:StyleID='m28735532'><Data ss:Type='String'>Marketing &#10;Ops</Data></Cell>"
		Response.write "<Cell ss:StyleID='m28735542'><Data ss:Type='String'>Supply &#10;Chain</Data></Cell>"
		Response.write "<Cell ss:StyleID='m28735552'><Data ss:Type='String'>Marketing &#10;Ops</Data></Cell>"
		Response.write "<Cell ss:StyleID='m28735562'><Data ss:Type='String'>Supply &#10;Chain</Data></Cell>"			
		Response.write "</Row>" 
	
		'end of BOM ENg info header row
	
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value
			Response.Write  " <Row ss:AutoFitHeight='0' ss:Height='12.75'>"
			Response.write "<Cell ss:MergeAcross='" & cstr(PlatformCt +6) & "' ss:StyleID='m28737792'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>"
   			
   			Response.Write  "</Row>" 

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

				Response.Write  " <Row ss:StyleID='s70'>"
				if rsModule.Fields("RegionId").Value <> 43 then 
					Response.Write  "<Cell ss:StyleID='s93'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  &  "#" & Server.HTMLEncode (rsModule.Fields("DashCode").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				else
					Response.Write  "<Cell ss:StyleID='s93'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				end if
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode (rsModule.Fields("Replacement").Value) & "</Data></Cell>"	
			'	platform info
	
			'New way to display platform info-07/01/2003
			strPlatformIDs =strPlatformIDs_backup
			while instr(strPlatformIDs, ",") > 0
					strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
					oRsPlatforms.filter = "ModuleID = " & cstr(nModuleID) & " and AliasID=" & strtempPlatformID
					if oRsPlatforms.RecordCount > 0 then
						if oRsPlatforms.Fields("AvailableDate").Value <> "" Then
							arrDate = Split(Trim(oRsPlatforms.Fields("AvailableDate").Value ),"/")
							Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
						else
							response.write " <Cell ss:StyleID='s77'><Data ss:Type='String'>ü</Data></Cell>"	
						end if				
					else					
						Response.Write   "<Cell ss:StyleID='s74'/>"
					end if
					strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
			wend
			
		'	response.write strPlatformIDs_backup
		'	response.end
				'new columns 4/1/2004
				
				'Response.Write   "<Cell ss:StyleID='s101'/>"
				'Response.Write   "<Cell ss:StyleID='s101'/>"
				'Response.Write   "<Cell ss:StyleID='s101'/>"
				'Response.Write   "<Cell ss:StyleID='s101'/>"
				'Response.Write   "<Cell ss:StyleID='s101'/>"
				
				Response.Write  "<Cell ss:StyleID='s102'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("RuleDescription").Value) & "</Data></Cell>"
				
				' Only check base unit only
				
				if rsModule.Fields("RegionID").Value = 43 Then
				
					if rsModule.Fields("Localized").Value = 0 And cbool(rsModule.Fields("SCM_Localized").Value) = False Then
						response.write " <Cell ss:StyleID='s72'><Data ss:Type='String'></Data></Cell>"
					else
						response.write " <Cell ss:StyleID='s77'><Data ss:Type='String'>ü</Data></Cell>"	
					end if
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"
				end if				
				
				if len(rsModule.Fields("CPLBlindDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("CPLBlindDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" & arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("BOMRevADate").Value),"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("RASDiscontinueDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("NetWeight").Value & "</Data></Cell>"
			
				

				Response.Write  "</Row>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop

		oRsPlatforms.Close
		set oRsPlatforms = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function


function WriteAMO_SCM_History_ChangeGridXML (byval strModuleIDs)
	dim rsChangeLog,  oErr, oSvr, sSearchFilter
	
	sSearchFilter = " and AC.ModuleID in (" & strModuleIDs & ") and AC.HideFromSCM = 0|"
	
	set oSvr = New ISAMO
	set rsChangeLog = oSvr.AMO_ChangeHistory(Application("REPOSITORY"), sSearchFilter, rsChangeLog)
	if rsChangeLog is nothing then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM Change Log</Data></Cell></Row>"
		exit function
	end if
	
	if IsODM = 1 then 'hide cost items from ODM
		rsChangeLog.Filter = "FieldMask <> 32 and FieldMask <> 64 and FieldMask <> 131072"
	end if
	
	Response.write "<Column ss:Width='50.0'/>" & vbCrLf
	Response.write "<Column ss:Width='50.0'/>" & vbCrLf
	Response.write "<Column ss:Width='80.0'/>" & vbCrLf
	Response.write "<Column ss:Width='150.0'/>" & vbCrLf
	Response.write "<Column ss:Width='250.0'/>" & vbCrLf
	Response.write "<Column ss:Width='200.0'/>" & vbCrLf

	
	if rsChangeLog.recordcount> 0 then
		rsChangeLog.sort="ChangehistoryID desc"
		rsChangeLog.MoveFirst
	end if 
	
	Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='5' ss:StyleID='m28734894'><Data ss:Type='String'>SCM Change History</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" & vbCrLf
	Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='5' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" & vbCrLf
	
	
	Response.Write  "<Row ss:Height='50.25'>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>Change Type</Data></Cell>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>Category</Data></Cell>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>HP PartNo</Data></Cell>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>Short Description</Data></Cell>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>Change Description</Data></Cell>" & vbCrLf
	Response.Write  "<Cell ss:StyleID='m28734884'><Data ss:Type='String'>Change Reason</Data></Cell>" & vbCrLf
	Response.Write  "</Row >" & vbCrLf
	
	do while (not rsChangeLog.EOF)		
		Response.Write  "<Row>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & rsChangeLog.Fields("ChangeType").Value & "</Data></Cell>"  & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & rsChangeLog.Fields("AMOCategory").Value & "</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & server.HTMLencode(rsChangeLog.Fields("HPPartNo").Value) & "</Data></Cell>"   & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & server.HTMLencode(rsChangeLog.Fields("ShortDescription").Value) & "</Data></Cell>"   & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & rsChangeLog.Fields("ChangeDescription").Value & "</Data></Cell>"   & vbCrLf
		Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='String'>" & rsChangeLog.Fields("Reason").Value & "</Data></Cell>"  & vbCrLf
		Response.Write  "</Row>" & vbCrLf
		
		rsChangeLog.MoveNext
	loop
		
	set rsChangeLog = nothing
	set oErr = nothing
	set oSvr = nothing
		
end function


function WriteAMO_SCM_Alter_ReportGridXML (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs, byVal sUserBusSeg, ByVal intFormat, byref strModuleIDs)
	dim strGEO, strRegion, strDivisionIDs
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
	dim PlatformCt, i, arrDate, oRsBusSeg, strAllRegionIDs, oRsCompatibility
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup, strStype, intIsDiscontinue
	
	
	set oRsBusSeg = Nothing
	
	'set oErr = GetMOLCategory(oRsBusSeg, 34)
	Set oRsBusSeg = Server.CreateObject("ADODB.Recordset")
	set oRsBusSeg = GetMOLCategory(28)	

	if oRsBusSeg is Nothing then
		Response.Write("Recordset error: oRsBusSeg")
		Response.End()
	end if
	
	if rsModule.recordcount = 0 then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>No After Market Option found that matches filter</Data></Cell></Row>"
		exit function
	else
		rsModule.filter = "RegionID=43"
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	strRevision = rsModule("Revision").value
	
	set oSvr = New ISAMO
	set oRsAllRegions = oSvr.AllRegions_Search(Application("REPOSITORY"), strDivisionIDs, null)
	if oRsAllRegions is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	if oRsAllRegions.recordcount> 0 then
		oRsAllRegions.Sort = "SortOrder ASC, Region ASC"
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
	set oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsGEOs is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	set oRsRegions = oSvr.AMO_SCM_Regions_Search(Application("REPOSITORY"), strModuleIDs,SCMID)
	if oRsRegions is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	
	
	set oRsPlatforms = oSvr.AMO_SCM_Platforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs, SCMID)
	if oRsPlatforms is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	
	set oRsCompatibility = oSvr.AMOCompatibility_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsCompatibility is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
		
		'add new columns 4/1/2004
		Response.write "<Column ss:Width='50.75'/>" & vbCrLf
		Response.write "<Column ss:Width='150.75'/>" & vbCrLf
		Response.write "<Column ss:Width='120.75'/>" & vbCrLf
		Response.write "<Column ss:Width='30.75'/>" & vbCrLf
		Response.write "<Column ss:Width='30.75'/>" & vbCrLf
		
		
		if intFormat = 0 then
			intCompatibility = 0
			strPlatform = ""
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform then
					Response.write "<Column ss:Width='50.00'/>" & vbCrLf
					intCompatibility = intCompatibility + 1
				end if
				strPlatform = oRsPlatforms("FullName").value
				oRsPlatforms.MoveNext
			loop

			strCompatibilityPlatforms = "Platforms"
		else			
			intCompatibility = oRsBusSeg.RecordCount
			do while (not oRsBusSeg.EOF)
				Response.write "<Column ss:Width='30.00'/>" & vbCrLf
				oRsBusSeg.MoveNext
			loop
			
			strCompatibilityPlatforms = "Compatibility"
		end if
		
		strPlatformIDs_backup = strPlatformIDs
		
		Response.write "<Column ss:Width='200.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='20.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80.75'/>" & vbCrLf
		Response.write "<Column ss:Width='80' ss:Span='100'/>" & vbCrLf

		'if len(trim(PublishDate)) > 0 then
			Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='20' ss:StyleID='m28734894'><Data ss:Type='String'>Option SCM</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" & vbCrLf
		'else
			'Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='20' ss:StyleID='m28734894'><Data ss:Type='String'>" & sUserBusSeg & " AMO SCM</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" & vbCrLf
		'end if
		
		Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Published: " & PublishDate &  "</Data><NamedCell ss:Name='Print_Titles'/></Cell></Row>" & vbCrLf
		Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" & vbCrLf
			
		'Row 1; Heading
		Response.Write  "<Row ss:Height='15.75' ss:StyleID='s30'>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>PID/PN</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Market Description</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Short Description</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>Info Delta</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>Localization(Y/N)</Data></Cell>" & vbCrLf
		if intCompatibility <= 0 then
			Response.Write  "<Cell ss:MergeAcross='0' ss:StyleID='m28734884'><Data ss:Type='String'>" & strCompatibilityPlatforms & "</Data></Cell>" & vbCrLf
		else
			Response.Write  "<Cell ss:MergeAcross='" & intCompatibility-1 & "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strCompatibilityPlatforms & "</Data></Cell>" & vbCrLf
		end if
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Rules Description</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Global Series Config EOL</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Rplcmnt PN</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m21907241'><Data ss:Type='String'>Regions</Data></Cell>" & vbCrLf

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
			if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
				if trim(oRsAllRegions("CountryGEO").value) <> strGEO and strGEO <> "" then
					'if ucase(strGEO) <> "GLOBAL" then
						Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" & vbCrLf
					'end if
					intCount = 0
					intIndex = intIndex + 1
				end if
				strGEO = trim(oRsAllRegions("CountryGEO").value)
				strRegion = trim(oRsAllRegions("Region").value)
				intCount = intCount + 1
				strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
			end if
			oRsAllRegions.MoveNext
		loop
		if intCount > 0  then
			Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" & vbCrLf
		end if

		response.write " </Row>" & vbCrLf
		

		'display compatibility/platform
		
		
		if intFormat = 0 then
			strPlatformIDs =""
			strtempPlatformID = ""
			intCompatibility = 0
			if oRsPlatforms.recordcount> 0 then
				Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>" & vbCrLf
				oRsPlatforms.MoveFirst
				strMarketingName = ""
				strPlatform = ""
				bDisplayed = true
				do while (not oRsPlatforms.EOF)
					if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
					  if trim(strMarketingName) <> "" then
						Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>" & vbCrLf
					  else
						Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf
					  end if
					  
						intCompatibility = intCompatibility + 1
						bDisplayed = true
						strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
					end if
					strPlatform = oRsPlatforms("FullName").value
					strMarketingName = oRsPlatforms("MarketingName").value & " " & oRsPlatforms("BusinessSegs").value
					strtempPlatformID = oRsPlatforms("AliasID").value
					bDisplayed = false
					oRsPlatforms.MoveNext
				loop
				if not bDisplayed then
					if trim(strMarketingName) <> "" then
					  Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>" & vbCrLf
					else
					  Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf
					end if
					intCompatibility = intCompatibility + 1
					strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
				end if
				strPlatformIDs_backup = strPlatformIDs
			else	'no platforms	jec
				Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>" & vbCrLf	'jec
				Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf	'jec
			end if
		else
			if oRsBusSeg.recordcount > 0 then
				oRsBusSeg.MoveFirst
			end if
		
			Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>" & vbCrLf
			intCompatibility = 0
			do while (not oRsBusSeg.EOF)
				Response.Write  "<Cell ss:Index='" &cstr(intCompatibility + 6) & "' ss:MergeDown='1' ss:StyleID='m28734944'><Data ss:Type='String'>" & oRsBusSeg("Description").value & "</Data></Cell>" & vbCrLf
				oRsBusSeg.MoveNext
				intCompatibility = intCompatibility + 1
			loop
		
		end if
				

		'second is the region name
		dim firstregionrow
		firstregionrow =1
		strAllRegionIDs = ""
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		do while (not oRsAllRegions.EOF)
			'if trim(ucase(oRsAllRegions("CountryGEO").value)) <> "GLOBAL" then
				
				if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
					if firstregionrow = 1 then
						if intCompatibility = 0 then	'no platforms	jec
							Response.Write  "<Cell ss:Index='" &cstr(14 + intCompatibility) & "' ss:StyleID='m28734944'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>" & vbCrLf
						else
							Response.Write  "<Cell ss:Index='" &cstr(13 + intCompatibility) & "' ss:StyleID='m28734944'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>" & vbCrLf
						end if
					else
						Response.Write  "<Cell ss:StyleID='m28734944'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>" & vbCrLf
		
					end if
					firstregionrow = firstregionrow + 1
				end if
				strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
				
			'end if
			oRsAllRegions.MoveNext
		loop
		
		Response.Write  "</Row>" & vbCrLf		
		
		'third is the HP Dash Code for the region
		
		Response.Write  "<Row ss:Height='28.75'>" & vbCrLf
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strRegionIDsBackup = ""
		intCount = 0
		strAllRegionIDs = ""
		do while (not oRsAllRegions.EOF)
			if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
				'if trim(ucase(oRsAllRegions("CountryGEO").value)) <> "GLOBAL" then
					if intCompatibility = 0 then	'no platforms jec
						Response.Write  "  <Cell ss:Index='" & cstr(intCount + 14 + intCompatibility) & "' ss:StyleID='m28734884'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>" & vbCrLf
					else
						Response.Write  "  <Cell ss:Index='" & cstr(intCount + 13 + intCompatibility) & "' ss:StyleID='m28734884'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>" & vbCrLf
					end if
					strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
				'end if
				strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
				intCount = intCount + 1
			end if
			oRsAllRegions.MoveNext
		loop
						
		Response.write "</Row>" & vbCrLf
		
	
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value			
			Response.Write  " <Row>" & vbCrLf
			if intCompatibility = 0 then	'no platforms jec
				Response.write "<Cell ss:MergeAcross='" & intCompatibility + 5 & "' ss:StyleID='m21907240'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>" & vbCrLf
			else
				Response.write "<Cell ss:MergeAcross='" & intCompatibility + 4 & "' ss:StyleID='m21907240'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>" & vbCrLf
			end if
			Response.Write  "<Cell ss:StyleID='m21907240'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("CategoryRuleDesc").Value) & "</Data></Cell>" & vbCrLf
			if Trim(rsModule.Fields("RuleMin").Value) = ""	then
				Response.write "<Cell ss:StyleID='m21907240'><Data ss:Type='String'></Data></Cell>" & vbCrLf
			else
				Response.write "<Cell ss:StyleID='m21907240'><Data ss:Type='String'> Min = " &  rsModule.Fields("RuleMin").Value &  "</Data></Cell>" & vbCrLf
			end if
			if Trim(rsModule.Fields("RuleMax").Value) = ""	then
				Response.write "<Cell ss:StyleID='m21907240'><Data ss:Type='String'></Data></Cell>" & vbCrLf
			else
				Response.write "<Cell ss:StyleID='m21907240'><Data ss:Type='String'> Max = " &  rsModule.Fields("RuleMax").Value &  "</Data></Cell>" & vbCrLf
			end if
			Response.write "<Cell ss:MergeAcross='" & intCount+3 & "' ss:StyleID='m21907240'><Data ss:Type='String'></Data></Cell>" & vbCrLf
   			Response.Write  "</Row>" & vbCrLf

			do while ((not rsModule.EOF) and (rsModule.Fields("CategoryID").value = strCurrentSection))
				'write each option
				'Use bgcolor='#FFFF99' for highlighted cells to indicate changes
				nModuleID = rsModule.Fields("ModuleID").Value
				intIsDiscontinue = rsModule.Fields("IsDiscontinue").value
				strGEOIDs = strGEOIDsBackup
				strRegionIDs = strRegionIDsBackup
				if rsModule.Fields("AMOStatusID").value = Application("AMO_RASREVIEW") or rsModule.Fields("AMOStatusID").value = Application("AMO_RASUPDATE") then
					bLocked = True
				else
					bLocked = False
				end if
				
				if clng(rsModule.Fields("IsDiscontinue").value) = 1 Then 
					strStype = "s112"
				else
					strStype = "s102"
				end if
				
				if clng(rsModule.Fields("ReportChangeMask").value) > 0 Then
					sDelta = "C"
				else
					sDelta = GetDelta(oRsRegions, nModuleID)
				end if

				Response.Write  " <Row>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "bluepartno", intIsDiscontinue) & "'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "description", intIsDiscontinue) & "'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("Description").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "shortdescription", intIsDiscontinue) & "'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & strStype & "'><Data ss:Type='String'>" & sDelta & "</Data></Cell>"	
				' Only check base unit only
				
				if rsModule.Fields("RegionID").Value = 43 Then
				
					if rsModule.Fields("Localized").Value = 0 And cbool(rsModule.Fields("SCM_Localized").Value) = False Then
						response.write " <Cell ss:StyleID='" & strStype & "'><Data ss:Type='String'>N</Data></Cell>" & vbCrLf
					else
						response.write " <Cell ss:StyleID='" & strStype & "'><Data ss:Type='String'>Y</Data></Cell>" & vbCrLf
					end if
				else
					Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
				end if
				
					
				
				strPlatformIDs = strPlatformIDs_backup
				
				if intFormat = 0 then
					if instr(strPlatformIDs, ",") = 0 then
						Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
					else
						while instr(strPlatformIDs, ",") > 0
							strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
							oRsPlatforms.filter = "ModuleID = " & cstr(nModuleID) & " and AliasID=" & strtempPlatformID
							if oRsPlatforms.RecordCount > 0  then
								if oRsPlatforms.Fields("ReportChangeMask").Value <> 2 then
									response.write " <Cell ss:StyleID='" & SetPlatformComHighlight(oRsPlatforms.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'>" & oRsPlatforms.Fields("AvailableStatus").Value & "</Data></Cell>" & vbCrLf
								else
									response.write " <Cell ss:StyleID='" & SetPlatformComHighlight(oRsPlatforms.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'></Data></Cell>" & vbCrLf
								end if
							else					
								Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
							end if
							strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
						wend
					end if
					
					
				else
					if oRsBusSeg.RecordCount > 0 then
						oRsBusSeg.MoveFirst
					end if
					do while (not oRsBusSeg.EOF)
						oRsCompatibility.filter = "ModuleID = " & cstr(nModuleID) & " and DivisionID=" & oRsBusSeg.Fields("CategoryId").Value
						if oRsCompatibility.RecordCount > 0 then
							if oRsCompatibility.Fields("ReportChangeMask").Value <> 2 then
								Response.Write  "<Cell ss:StyleID='" & SetPlatformComHighlight(oRsCompatibility.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'>" & oRsCompatibility.Fields("AvailableStatus").Value & "</Data></Cell>" & vbCrLf
							else
								Response.Write  "<Cell ss:StyleID='" & SetPlatformComHighlight(oRsCompatibility.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'></Data></Cell>" & vbCrLf
							end if
						else
							Response.Write  "<Cell ss:StyleID='" & strStype & "'><Data ss:Type='String'></Data></Cell>" & vbCrLf
						end if
						oRsBusSeg.MoveNext
					loop
				
				end if
				
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "ruledescription", rsModule.Fields("IsDiscontinue").value) & "'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("RuleDescription").Value) & "</Data></Cell>" & vbCrLf
				
				
				oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID = " & Application("AMO_GLOBAL_REGIONID") & "  and SCMID = " & SCMID
				
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("BOMRevADate").Value),"/")
					Response.Write  "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "bomrevadate", intIsDiscontinue) & "'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>" & vbCrLf
				else
					Response.Write   "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "bomrevadate", intIsDiscontinue) & "'/>" & vbCrLf

				end if 
				
				if len(rsModule.Fields("CPLBlindDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("CPLBlindDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "cplblinddate", intIsDiscontinue) & "'><Data ss:Type='DateTime'>" & arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>" & vbCrLf
				else
					Response.Write   "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "cplblinddate", intIsDiscontinue) & "'/>" & vbCrLf

				end if 
				
				
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("RASDiscontinueDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "rasdiscontinuedate", intIsDiscontinue) & "'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>" & vbCrLf
				else
					Response.Write   "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "rasdiscontinuedate", intIsDiscontinue) & "'/>" & vbCrLf
				end if 
				
				if len(rsModule.Fields("GlobalSeriesDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("GlobalSeriesDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "globalseriesdate", intIsDiscontinue) & "'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>" & vbCrLf
				else
					Response.Write   "<Cell ss:StyleID='" & SetDateHighlight(oRsRegions.Fields("ReportChangeMask").Value, "globalseriesdate", intIsDiscontinue) & "'/>" & vbCrLf
				end if 
				
				
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "replacement", intIsDiscontinue) & "'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("Replacement").Value) & "</Data></Cell>" & vbCrLf
				
				
				Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
				
				oRsRegions.filter = ""		

				while instr(strRegionIDs, ",") > 0
					strRegionID = left(strRegionIDs, instr(strRegionIDs, "|")-1)
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID=" & strRegionID & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " "
					if oRsRegions.RecordCount > 0  then
						if oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write " <Cell ss:StyleID='" & SetRegionHighlight(oRsRegions.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'> GA: " & oRsRegions("BOMRevADate").value & " - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</Data></Cell>" & vbCrLf
						elseif  oRsRegions("BOMRevADate").value = "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write " <Cell ss:StyleID='" & SetRegionHighlight(oRsRegions.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'> GA: N/A - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</Data></Cell>" & vbCrLf
						elseif  oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value = "" then
							response.write " <Cell ss:StyleID='" & SetRegionHighlight(oRsRegions.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'> GA: " & oRsRegions("BOMRevADate").value & " - EM: N/A</Data></Cell>" & vbCrLf
						else
							response.write " <Cell ss:StyleID='" & SetRegionHighlight(oRsRegions.Fields("ReportChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'> GA: N/A - EM: N/A</Data></Cell>" & vbCrLf
						end if
					else
						response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
					end if

					strRegionIDs = mid(strRegionIDs, instr(strRegionIDs, ",")+1)
				wend
				Response.Write  "</Row>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop
		
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s109'><Data ss:Type='String'>Legend</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Compatibility:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>A: Available (currently supported)</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>N: Do Not Use</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>N/A: Not Available</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>O: Ongoing</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Info Delta:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>C = Change</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Color Coding:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s105'><Data ss:Type='String'>Cell contents changed since last publish</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s107'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s108'><Data ss:Type='String'>PN past End of Manufacturing (EM)</Data></Cell></Row>" & vbCrLf

		oRsGEOs.Close
		set oRsGEOs = nothing
		oRsRegions.Close
		set oRsRegions = nothing
		oRsAllRegions.Close
		set oRsAllRegions = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function


function WriteAMO_SCM_ReportGridXML_DateOnly (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs, byVal sUserBusSeg, ByVal intFormat, byref strModuleIDs)
	dim strGEO, strRegion, strDivisionIDs
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
	dim PlatformCt, i, arrDate, oRsBusSeg, strAllRegionIDs, oRsCompatibility
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup, strStype, intIsDiscontinue

	
	set oRsBusSeg = Nothing
	
	'set oErr = GetMOLCategory(oRsBusSeg, 28)
	Set oRsBusSeg = Server.CreateObject("ADODB.Recordset")
	set oRsBusSeg = GetMOLCategory(34)	

	if oRsBusSeg is Nothing then
		Response.Write("Recordset error: oRsBusSeg")
		Response.End()
	end if
	
	if rsModule.recordcount = 0 then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>No After Market Option found that matches filter</Data></Cell></Row>"
		exit function
	else
		rsModule.filter = "RegionID=43"
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	strRevision = rsModule("Revision").value
	
	set oSvr = New ISAMO
	
	
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
		
	set oRsPlatforms = oSvr.AMO_SCM_Platforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs, SCMID)
	if oRsPlatforms is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	

		
		'add new columns 4/1/2004
		Response.write "<Column ss:Width='50.75'/>" & vbCrLf
		Response.write "<Column ss:Width='150.75'/>" & vbCrLf
		Response.write "<Column ss:Width='120.75'/>" & vbCrLf
		
		
		if intFormat = 0 then
			intCompatibility = 0
			strPlatform = ""
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform then
					Response.write "<Column ss:Width='50.00'/>" & vbCrLf
					intCompatibility = intCompatibility + 1
				end if
				strPlatform = oRsPlatforms("FullName").value
				oRsPlatforms.MoveNext
			loop
			
			strCompatibilityPlatforms = "Platforms"
		else			
			intCompatibility = oRsBusSeg.RecordCount
			do while (not oRsBusSeg.EOF)
				Response.write "<Column ss:Width='30.00'/>" & vbCrLf
				oRsBusSeg.MoveNext
			loop
			
			strCompatibilityPlatforms = "Compatibility"
		end if
		
		strPlatformIDs_backup = strPlatformIDs
		
		Response.write "<Column ss:Width='80' ss:Span='100'/>" & vbCrLf

		Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='20' ss:StyleID='m28734894'><Data ss:Type='String'>Option SCM</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" & vbCrLf
		Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Published: " & PublishDate &  "</Data><NamedCell ss:Name='Print_Titles'/></Cell></Row>" & vbCrLf
		Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" & vbCrLf
			
		'Row 1; Heading
		Response.Write  "<Row ss:Height='15.75' ss:StyleID='s30'>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>PID/PN</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Market Description</Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Short Description</Data></Cell>" & vbCrLf
		if intCompatibility <= 0 then
			Response.Write  "<Cell ss:MergeAcross='0' ss:StyleID='m28734884'><Data ss:Type='String'>" & strCompatibilityPlatforms & "</Data></Cell>" & vbCrLf
		else
			Response.Write  "<Cell ss:MergeAcross='" & intCompatibility-1 & "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strCompatibilityPlatforms & "</Data></Cell>" & vbCrLf
		end if

		response.write " </Row>" & vbCrLf
		

		'display platform
		
		strPlatformIDs =""
		strtempPlatformID = ""
		intCompatibility = 0
		if oRsPlatforms.recordcount> 0 then
			Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>" & vbCrLf
			oRsPlatforms.MoveFirst
			strMarketingName = ""
			strPlatform = ""
			bDisplayed = true
			do while (not oRsPlatforms.EOF)
				if oRsPlatforms("FullName").value <> strPlatform and strPlatform <> "" then
				  if trim(strMarketingName) <> "" then
					Response.Write  "<Cell ss:MergeDown='1' ss:Index='" &cstr(intCompatibility + 4) & "'  ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>" & vbCrLf
				  else
					Response.Write  "<Cell ss:MergeDown='1' ss:Index='" &cstr(intCompatibility + 4) & "'  ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf
				  end if
					  
					intCompatibility = intCompatibility + 1
					bDisplayed = true
					strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
				end if
				strPlatform = oRsPlatforms("FullName").value
				strMarketingName = oRsPlatforms("MarketingName").value & " " & oRsPlatforms("BusinessSegs").value
				strtempPlatformID = oRsPlatforms("AliasID").value
				bDisplayed = false
				oRsPlatforms.MoveNext
			loop
			if not bDisplayed then
				if trim(strMarketingName) <> "" then
				  Response.Write  "<Cell ss:MergeDown='1' ss:Index='" &cstr(intCompatibility + 4) & "'  ss:StyleID='m28734944'><Data ss:Type='String'>" & strMarketingName & "</Data></Cell>" & vbCrLf
				else
				  Response.Write  "<Cell ss:MergeDown='1' ss:Index='" &cstr(intCompatibility + 4) & "'  ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf
				end if
				intCompatibility = intCompatibility + 1
				strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
			end if
			strPlatformIDs_backup = strPlatformIDs
		else	'no platforms	jec
			Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>" & vbCrLf	'jec		
			Response.Write  "<Cell ss:MergeDown='1' ss:Index='" &cstr(intCompatibility + 4) & "'  ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>" & vbCrLf 'jec
		end if
		
		Response.Write  "</Row><Row></Row>" & vbCrLf		
		
		
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value	
			intIsDiscontinue = rsModule.Fields("IsDiscontinue").value		
			Response.Write  " <Row>" & vbCrLf
			if intCompatibility = 0 then	'no platforms	jec
				Response.write "<Cell ss:MergeAcross='" & intCompatibility + 3 & "' ss:StyleID='m21907240'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>" & vbCrLf
			else
				Response.write "<Cell ss:MergeAcross='" & intCompatibility + 2 & "' ss:StyleID='m21907240'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>" & vbCrLf
			end if
   			Response.Write  "</Row>" & vbCrLf

			do while ((not rsModule.EOF) and (rsModule.Fields("CategoryID").value = strCurrentSection))
				'write each option
				'Use bgcolor='#FFFF99' for highlighted cells to indicate changes
				nModuleID = rsModule.Fields("ModuleID").Value
				
				if rsModule.Fields("AMOStatusID").value = Application("AMO_RASREVIEW") or rsModule.Fields("AMOStatusID").value = Application("AMO_RASUPDATE") then
					bLocked = True
				else
					bLocked = False
				end if
				
				if clng(rsModule.Fields("IsDiscontinue").value) = 1 Then 
					strStype = "s112"
				else
					strStype = "s102"
				end if
		
				

				Response.Write  " <Row>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "bluepartno", intIsDiscontinue) & "'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "description", intIsDiscontinue) & "'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("Description").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				Response.Write  "<Cell ss:StyleID='" & SetHighlight(rsModule.Fields("ReportChangeMask").Value, "shortdescription", intIsDiscontinue) & "'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>" & vbCrLf
				' Only check base unit only
				
				strPlatformIDs = strPlatformIDs_backup
				
				
				if instr(strPlatformIDs, ",") = 0 then
					Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
				else
					while instr(strPlatformIDs, ",") > 0
						strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
						oRsPlatforms.filter = "ModuleID = " & cstr(nModuleID) & " and AliasID=" & strtempPlatformID
						if oRsPlatforms.RecordCount > 0 then
							if oRsPlatforms.Fields("AvailableDate").Value <> "" Then
								arrDate = Split(Trim(oRsPlatforms.Fields("AvailableDate").Value ),"/")
								Response.Write  "<Cell ss:StyleID='" & SetPlatformComHighlight(oRsPlatforms.Fields("ChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>" & vbCrLf
							elseif oRsPlatforms.Fields("ChangeMask").Value <> 2 Then
								response.write " <Cell ss:StyleID='" & SetPlatformComHighlight(oRsPlatforms.Fields("ChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'>" & oRsPlatforms.Fields("AvailableStatus").Value & "</Data></Cell>" & vbCrLf
							else
								response.write " <Cell ss:StyleID='" & SetPlatformComHighlight(oRsPlatforms.Fields("ChangeMask").Value, intIsDiscontinue) & "'><Data ss:Type='String'></Data></Cell>" & vbCrLf
							end if				
						else					
							Response.Write   "<Cell ss:StyleID='" & strStype & "'/>" & vbCrLf
						end if
						strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
					wend
				end if
					
				Response.Write  "</Row>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop
						
		
		
		
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s109'><Data ss:Type='String'>Legend</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Compatibility:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>A: Available (currently supported)</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>N: Do Not Use</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>N/A: Not Available</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>O: Ongoing</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Info Delta:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s106'><Data ss:Type='String'>C = Change</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s110'><Data ss:Type='String'>Color Coding:</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s104'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s105'><Data ss:Type='String'>Cell contents changed since last publish</Data></Cell></Row>" & vbCrLf
		Response.Write  "<Row><Cell ss:StyleID='s107'><Data ss:Type='String'></Data></Cell>" & vbCrLf
		Response.Write  "<Cell ss:StyleID='s108'><Data ss:Type='String'>PN past End of Manufacturing (EM)</Data></Cell></Row>" & vbCrLf

		
		oRsPlatforms.Close
		set oRsPlatforms = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function


'Determines if the table cell should be highlighted
function SetHighlight(intChangeMask, strField, IsDiscontinue)
	dim intMask
	intMask = 0
	
	if clng(IsDiscontinue) = 1 then
		SetHighlight = "s112"
	else
		SetHighlight = "s72"
	end if

	select case lcase(strField)
		case "bluepartno"
			intMask = 1
		case "ruledescription"
			intMask = 2
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
		case "manufacturecountry"
			intMask = 67108864
		case "longdescription"
			intMask = 134217728
		case "orderinstruction"
			intMask = 268435456
		case "replacementavdescription"
			intMask = 536870912	
	end select	

	if intMask <> 0 then
		if (intChangeMask AND clng(intMask)) = intMask then
			SetHighlight = "s55"
		end if
	end if
end function


'Determines if the table cell should be highlighted
function SetDateHighlight(intChangeMask, strField, IsDiscontinue)
	dim intMask
	intMask = 0
	
	if clng(IsDiscontinue) = 1 then
		SetDateHighlight = "s113"
	else
		SetDateHighlight = "s101"
	end if


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
		if (intChangeMask AND clng(intMask)) = intMask then
			SetDateHighlight = "s103"
		end if
	end if
end function


'Determines if the table cell should be highlighted
function SetRegionHighlight(intChangeMask, IsDiscontinue)
	dim intBomrevadateMask, intRasdiscontinuedateMask
	
	
	if clng(IsDiscontinue) = 1 then
		SetRegionHighlight = "s114"
	else
		SetRegionHighlight = "s72"
	end if
	
	 
	intBomrevadateMask = 16	
	intRasdiscontinuedateMask = 8		
		
	if ((intChangeMask AND clng(intBomrevadateMask)) = intBomrevadateMask)  Or ((intChangeMask AND clng(intRasdiscontinuedateMask)) = intRasdiscontinuedateMask) then
		SetRegionHighlight = "s55"
	end if

end function


'Determines if the table cell should be highlighted
function SetRegionHubHighlight(intChangeMask)
			
	if  intChangeMask = True then
		SetRegionHubHighlight = "s56"
	else
		SetRegionHubHighlight = "s72"
	end if

end function

'Determines if the table cell should be highlighted
function SetPlatformComHighlight(intChangeMask, IsDiscontinue)
		
		
	if clng(IsDiscontinue) = 1 then
		SetPlatformComHighlight = "s113"
	else
		SetPlatformComHighlight = "s101"
	end if
		
	if  intChangeMask <> 0 then
		SetPlatformComHighlight = "s103"
	end if

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
    
   
function WriteAMO_SCM_HUB_ReportGridXML (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs, byVal sUserBusSeg)
	dim strGEO, strRegion, strDivisionIDs, strModuleIDs
	dim strGEOID, strGEOIDs, strRegionIDs, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay
	dim nModuleID
	dim intCount, intIndex
	dim bDisplayed, bLocked
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr
	dim strClass, sPD, sChange, sErr, strPlatform, strMarketingName, strRevision
	dim intPlatformID
	dim nChassisID, numDL
	dim oRsPlatforms
	dim bPlatformAssigned, sDelta
	dim PlatformCt, i, arrDate, strAllRegionIDs
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup

	
	if rsModule.recordcount = 0 then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>No After Market Option found that matches filter</Data></Cell></Row>"
		exit function
	else
		rsModule.filter = "RegionID=43"
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	strRevision = rsModule("Revision").value
	
	set oSvr = New ISAMO
	set oRsAllRegions = oSvr.AllRegions_Search(Application("REPOSITORY"), strDivisionIDs, 1)
	if oRsAllRegions is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
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
		set oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsGEOs is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	set oRsRegions = oSvr.AMO_SCM_Regions_Search(Application("REPOSITORY"), strModuleIDs,SCMID)
	if oRsRegions is nothing then
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		exit function
	end if
	
	
		
		'add new columns 4/1/2004
		Response.write "<Column ss:Width='50.75'/>"
		Response.write "<Column ss:Width='150.75'/>"
		'Response.write "<Column ss:Width='120.75'/>"
		Response.write "<Column ss:Width='30.75'/>"
		Response.write "<Column ss:Width='30.75'/>"
		Response.write "<Column ss:Width='80.75'/>"
		'Response.write "<Column ss:Width='80.75'/>"
		Response.write "<Column ss:Width='80.75'/>"
		'Response.write "<Column ss:Width='40.75'/>"
		Response.write "<Column ss:Width='80' ss:Span='100'/>"

		Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='20' ss:StyleID='m28734894'><Data ss:Type='String'>Factory/Hub SKU Localization Table</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" 
		Response.Write  "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" 
			
		'Row 1; Heading
		Response.Write  "<Row ss:Height='15.75' ss:StyleID='s30'>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>PID/PN</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Market Description (64 characters)</Data></Cell>"
		'Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Short Description (32 characters)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>Info Delta</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>Localization(Y/N)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>"
		'Response.Write  " <Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734884'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>"
		'Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>Regions</Data></Cell>"

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
						Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" 
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
			Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28734884'><Data ss:Type='String'>" & strGEO & "</Data></Cell>"  
		end if

		response.write " </Row>"
		

		
		
		Response.Write  "<Row ss:Height='138' ss:StyleID='s30'>"
		

		'second is the region name
		dim firstregionrow
		firstregionrow =1
		strAllRegionIDs = ""
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		do while (not oRsAllRegions.EOF)
			oRsRegions.filter = "RegionID=" & oRsAllRegions("RegionID").value & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
			if oRsRegions.RecordCount > 0 then
				
				if instr(strAllRegionIDs, "," & oRsAllRegions("RegionID").value & ",") = 0 Then
					if firstregionrow = 1 then
						Response.Write  "<Cell ss:Index='" &cstr(7) & "' ss:StyleID='m28734944'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>"
					else
						Response.Write  "<Cell ss:StyleID='m28734944'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>"
		
					end if
					firstregionrow = firstregionrow + 1
				end if
				strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
				
			end if
			oRsAllRegions.MoveNext
		loop
		
		Response.Write  "</Row>"		
		
		'third is the HP Dash Code for the region
		
		Response.Write  "<Row ss:Height='28.75'>"
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
					Response.Write  "  <Cell ss:Index='" & cstr(intCount + 7) & "' ss:StyleID='m28734884'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>" 
					strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
					strAllRegionIDs = strAllRegionIDs & "," & oRsAllRegions("RegionID").value & ","
					intCount = intCount + 1
				end if
			end if
			oRsAllRegions.MoveNext
		loop
						
		Response.write "</Row>"
	
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value			
			Response.Write  "<Row>"
			Response.write "<Cell ss:MergeAcross='" & intCount + 5 & "' ss:StyleID='m21907240'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & rsModule.Fields("Category").Value &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>"
   			Response.Write  "</Row>" 

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

				Response.Write  " <Row>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value)  & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("Description").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				'Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s102'><Data ss:Type='String'>" & sDelta & "</Data></Cell>"	
				' Only check base unit only
				
				if rsModule.Fields("RegionID").Value = Application("AMO_GLOBAL_REGIONID") Then
				
					if rsModule.Fields("Localized").Value = 0 And cbool(rsModule.Fields("SCM_Localized").Value) = False Then
						response.write " <Cell ss:StyleID='s102'><Data ss:Type='String'>N</Data></Cell>"
					else
						response.write " <Cell ss:StyleID='s102'><Data ss:Type='String'>Y</Data></Cell>"	
					end if
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"
				end if	
				

				
				
				oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID = " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
				
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("BOMRevADate").Value),"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
								
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					arrDate = Split(Trim(rsModule.Fields("RASDiscontinueDate").Value),"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"
				end if 
				
				
				oRsRegions.filter = ""		

				while instr(strRegionIDs, ",") > 0
					strRegionID = left(strRegionIDs, instr(strRegionIDs, "|")-1)
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID=" & strRegionID & " and RegionID <> " & Application("AMO_GLOBAL_REGIONID") & " and HideFromHub <> 1"
					if oRsRegions.RecordCount > 0  then
						if oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write " <Cell ss:StyleID='" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & "'><Data ss:Type='String'> GA: " & oRsRegions("BOMRevADate").value & " - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</Data></Cell>"	
						elseif  oRsRegions("BOMRevADate").value = "" and oRsRegions.Fields("RASDiscontinueDate").Value <> "" then
							response.write " <Cell ss:StyleID='" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & "'><Data ss:Type='String'> GA: N/A - EM: " &  oRsRegions.Fields("RASDiscontinueDate").Value & "</Data></Cell>"	
						elseif  oRsRegions("BOMRevADate").value <> "" and oRsRegions.Fields("RASDiscontinueDate").Value = "" then
							response.write " <Cell ss:StyleID='" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & "'><Data ss:Type='String'> GA: " & oRsRegions("BOMRevADate").value & " - EM: N/A</Data></Cell>"	
						else
							response.write " <Cell ss:StyleID='" & SetRegionHubHighlight(oRsRegions.Fields("IsHubSKU").Value) & "'><Data ss:Type='String'> GA: N/A - EM: N/A</Data></Cell>"	
						end if
					else
						response.Write   "<Cell ss:StyleID='s74'/>"
					end if

					strRegionIDs = mid(strRegionIDs, instr(strRegionIDs, ",")+1)
				wend
				Response.Write  "</Row>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop

		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s104'><Data ss:Type='String'>Legend:</Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s104'><Data ss:Type='String'>Hub</Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s105'><Data ss:Type='String'>Indicates Hub SKU</Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'></Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s104'><Data ss:Type='String'>Info Delta</Data></Cell></Row>" 
		Response.Write  "<Row><Cell ss:MergeAcross='1' ss:StyleID='s106'><Data ss:Type='String'>C: Changed Since Last Published SCM</Data></Cell></Row>" 
		
		oRsGEOs.Close
		set oRsGEOs = nothing
		oRsRegions.Close
		set oRsRegions = nothing
		oRsAllRegions.Close
		set oRsAllRegions = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function  
   
   
   
   
   


%>
