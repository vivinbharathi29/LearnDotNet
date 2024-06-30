<%


function ConvertNull(Byref Value)
	if IsNUll(Value) or Trim(Value)="" then 
		ConvertNull = "&nbsp;"
	else
		ConvertNull = Value
	end if
end function

function ConvertNull_Space(Byref Value)
	if IsNUll(Value) or Trim(Value)="" then 
		ConvertNull_Space = " "
	else
		ConvertNull_Space = Value
	end if
end function


'Xml format version

function WriteAMO_Discontinuance_ReportGridXML (byval rsModule, byval SCMID, byVal sBusSegIDs)
	dim strGEO, strRegion, strDivisionIDs, strModuleIDs
	dim strGEOID, strGEOIDs, strRegionIDs, strRegionID
	dim strGEOIDsBackup, strRegionIDsBackup, strCurrentSection
	dim strHPDASHCode, strDisplay
	dim nModuleID
	dim intCount
	dim bDisplayed, bLocked
	dim oRsGEOs, oRsRegions, oRsAllRegions, oErr, oSvr
	dim strClass, sPD, sChange, sErr, strPlatform, strMarketingName
	dim intPlatformID
	dim nChassisID, numDL
	dim oRsPlatforms
	dim bPlatformAssigned
	dim PlatformCt, i, arrDate
	dim strPlatformIDs, strtempPlatformID, strPlatformIDs_backup
	dim oRsAllGEOs
	if rsModule.recordcount = 0 then
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>No After Market Option found that matches filter</Data></Cell></Row>"
		exit function
	else
		rsModule.movefirst	
	end if 
	strDivisionIDs = rsModule("DivisionIDs").value
	
	'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
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

	set oRsAllRegions = oSvr.AMO_Report_AllRegions(Application("REPOSITORY"), strModuleIDs, 1)
	if oRsAllRegions is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
	  set oErr = nothing
		exit function
	end if
	if oRsAllRegions.recordcount> 0 then
		oRsAllRegions.Sort = "SortOrder ASC, Region ASC"
	end if
	
	set oRsAllGEOs = oSvr.AMO_Report_AllGEOs(Application("REPOSITORY"), strModuleIDs, 1)
	if oRsAllGEOs is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO Discontinuance Report</Data></Cell></Row>"
	  set oErr = nothing
		exit function
	end if
	if oRsAllGEOs.recordcount> 0 then
		oRsAllGEOs.Sort = "GEO"
	end if
	
	
	set oRsGEOs = oSvr.AMOGEOs_Search(Application("REPOSITORY"), strModuleIDs, SCMID)
	if oRsGEOs is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		set oErr = nothing
		exit function
	end if
	set oRsRegions = oSvr.AMO_SCM_Regions_Search(Application("REPOSITORY"), strModuleIDs,SCMID)
	if oRsRegions is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		set oErr = nothing
		exit function
	end if
	
	set oRsPlatforms = oSvr.AMO_SCM_Platforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs,SCMID)
	if oRsPlatforms is nothing then
		sErr = "Missing required parameters.  Unable to complete your request."
		'response.write sErr
		response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
		set oErr = nothing
		exit function
	end if
	
	'Hardcoded info
		'Response.write "<Column ss:Width=‘65.25’/>"
   		'Response.write "<Column ss:Width=‘151.5’/>"
   		'Response.write "<Column ss:Width=‘45’/>"
		'Response.write "<Column ss:Width=‘28.5’/>"
		'get platformcunt fo rcolumn info

		
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
		'i = 0
		'do while  i < PlatformCt
		'	Response.write "<Column ss:AutoFitWidth=‘0’ ss:Width=‘28.5’/>"
		'	i=i+1
		'loop
		if PlatformCt > 1 then
			Response.write  "<Column ss:AutoFitWidth='0' ss:Width='28.25' ss:Span='" & cstr(PlatformCt-1) & "'/>"
		elseif PlatformCt =1 then
			Response.write "<Column ss:Width='28.25'/>"
		end if 
		
		'add new columns 4/1/2004
		
		Response.write "<Column ss:Width='81.75'/>"
		Response.write " <Column ss:Width='84'/>"
			
		Response.write "<Column ss:Width='110.25'/>"

		Response.write "  <Column ss:Width='27.25'/>"
		'response.write "<Column ss:AutoFitWidth='0' ss:Width='27.75'/>"

		Response.Write  " <Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" 
		
		'Row 1; Heading
		Response.Write  "<Row ss:Height='15.75' ss:StyleID='s30'>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734914'><Data ss:Type='String'>Part Number</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734924'><Data ss:Type='String'>Option (Short Description)</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734924'><Data ss:Type='String'>Replacement</Data></Cell>"
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
					Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>" & server.htmlencode(strMarketingName) & "</Data></Cell>"
				  else
					Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
				  end if
				  
					PlatformCt = PlatformCt + 1
					bDisplayed = true
					strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
				end if
				strPlatform = oRsPlatforms("FullName").value
				strMarketingName = oRsPlatforms("MarketingName").value
				strtempPlatformID = oRsPlatforms("AliasID").value
				bDisplayed = false
				oRsPlatforms.MoveNext
			loop
			if not bDisplayed then
				if trim(strMarketingName) <> "" then
				  Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>" & server.htmlencode(strMarketingName) & "</Data></Cell>"
				else
				  Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
				end if
				PlatformCt = PlatformCt + 1
				strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
			end if
			strPlatformIDs_backup = strPlatformIDs
		end if
		'end if platform headers
		'new columns 4/1/2004			
				
		Response.Write  " <Cell ss:MergeDown='2' ss:StyleID='m28734954'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>"
		Response.Write  " <Cell ss:MergeDown='2' ss:StyleID='m28735036'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28735046'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>"
		Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28735066'><Data ss:Type='String'>Regions</Data></Cell>"
		
		'display all the GEOs
		'if oRsAllRegions.recordcount> 0 then
		'	oRsAllRegions.MoveFirst
		'end if
		dim GEOcnt 
		GEOcnt   =0
		strGEO = ""
		strGEOID = ""
		strGEOIDsBackup = ""
		bDisplayed = True
		do while (not oRsAllGEOs.EOF)
			if oRsAllGEOs("GEO").value <> strGEO and strGEO <> "" then
				if ucase(strGEO) <> "GLOBAL" then	'don't display GLOBAL
					Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28735188'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" 
					strGEOIDsBackup = strGEOIDsBackup & strGEOID & ","
					GEOcnt  = GEOcnt  + 1
					bDisplayed = True
				end if
			end if
			strGEO = oRsAllGEOs("GEO").value
			strGEOID = oRsAllGEOs("GEOID").value
			bDisplayed = False
			oRsAllGEOs.MoveNext
		loop
		if not bDisplayed then
			Response.Write  "<Cell ss:MergeDown='2' ss:StyleID='m28735188'><Data ss:Type='String'>" & strGEO & "</Data></Cell>"
			GEOcnt  = GEOcnt  + 1
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
		dim ballcountryshown
		do while (not oRsAllRegions.EOF)
			if trim(oRsAllRegions("CountryGEO").value) <> strGEO and strGEO <> "" then
				if ucase(strGEO) = "GLOBAL" then
					Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28735340'><Data ss:Type='String'>" & strRegion &"</Data></Cell>" 
					ballcountryshown =true
				else
					Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28735350'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" 
				end if
				intCount = 0
			end if
			strGEO = trim(oRsAllRegions("CountryGEO").value)
			strRegion = trim(oRsAllRegions("Region").value)
			intCount = intCount + 1
			oRsAllRegions.MoveNext
		loop
		if intCount > 0 then
			if ucase(strGEO) = "GLOBAL" then
					Response.Write  "<Cell ss:MergeDown='1' ss:StyleID='m28735340'><Data ss:Type='String'>" & strRegion &"</Data></Cell>" 
					ballcountryshown =true
			else
					Response.Write  "<Cell ss:MergeAcross='" & cstr(intCount-1) &  "' ss:StyleID='m28735350'><Data ss:Type='String'>" & strGEO & "</Data></Cell>" 
			end if
			
		end if

		response.write " </Row>"

		'second is the region name
		dim firstregionrow
		firstregionrow =1
		Response.Write  "	<Row ss:Height='138' ss:StyleID='s30'>"
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if
		do while (not oRsAllRegions.EOF)
			if trim(ucase(oRsAllRegions("CountryGEO").value)) <> "GLOBAL" then
				if firstregionrow = 1 then
					if ballcountryshown then
						Response.Write  "<Cell ss:Index='" &cstr(PlatformCt + GEOcnt + 9) & "' ss:StyleID='s37'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>"
					else
						Response.Write  "<Cell ss:Index='" &cstr(PlatformCt + GEOcnt + 8) & "' ss:StyleID='s37'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>"
					end if 
				else
			
					Response.Write  "<Cell ss:StyleID='s37'><Data ss:Type='String'>" & trim(oRsAllRegions("Region").value) & "</Data></Cell>"
		
				end if
				firstregionrow = firstregionrow + 1
			end if
			oRsAllRegions.MoveNext
		loop
		Response.Write  "</Row>"		
		
		'third is the HP Dash Code for the region
		
		'BOM ENg info header row
		Response.Write  "<Row ss:Height='28.75'>"
		
		
		
		dim bfirstHpcodecell
		bfirstHpcodecell =1
		if oRsAllRegions.recordcount> 0 then
			oRsAllRegions.MoveFirst
		end if	
		strRegionIDsBackup = ""
		do while (not oRsAllRegions.EOF)
			if bfirstHpcodecell =1 then
				if ballcountryshown then
						Response.Write  "  <Cell ss:Index='" &cstr(PlatformCt + GEOcnt + 8) & "' ss:StyleID='m28735796'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>" 
				else
						Response.Write  "<Cell ss:Index='" &cstr(PlatformCt + GEOcnt + 7) & "' ss:StyleID='m28735796'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>"
				end if 
				
			else
				Response.Write  "  <Cell ss:StyleID='m28735796'><Data ss:Type='String'>" & ConvertNull(trim(oRsAllRegions("HPDASHCode").value)) & "</Data></Cell>" 
			end if 
			bfirstHpcodecell = bfirstHpcodecell +1
			strRegionIDsBackup = strRegionIDsBackup & cstr(oRsAllRegions("RegionID").value) & "|" & oRsAllRegions("HPDASHCode").value & ","
			oRsAllRegions.MoveNext
		loop
				
		Response.write "</Row>" 
	
		'end of BOM ENg info header row
	
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value
			Response.Write  " <Row ss:AutoFitHeight='0' ss:Height='12.75'>"
			Response.write "<Cell ss:MergeAcross='" & cstr(PlatformCt +6) & "' ss:StyleID='m28737792'><Data ss:Type='String'>" & rsModule.Fields("ModuleType").Value & "-" & server.htmlencode(rsModule.Fields("Category").Value) &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>"
   			
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
				Response.Write  "<Cell ss:StyleID='s93'><Data ss:Type='String'>" & Server.HTMLEncode (rsModule.Fields("BluePN").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" &  Server.HTMLEncode (rsModule.Fields("Replacement").Value) & "</Data></Cell>"	
			'	platform info
	
			'New way to display platform info-07/01/2003
			strPlatformIDs =strPlatformIDs_backup
			while instr(strPlatformIDs, ",") > 0
					strtempPlatformID = left(strPlatformIDs, instr(strPlatformIDs, ",")-1)
					oRsPlatforms.filter = "ModuleID = " & cstr(nModuleID) & " and AliasID=" & strtempPlatformID
					if oRsPlatforms.RecordCount > 0 then
						response.write " <Cell ss:StyleID='s77'><Data ss:Type='String'>ü</Data></Cell>"					
					else					
						Response.Write   "<Cell ss:StyleID='s74'/>"
					end if
					strPlatformIDs = mid(strPlatformIDs, instr(strPlatformIDs, ",")+1)
			wend
			
		'	response.write strPlatformIDs_backup
		'	response.end
							
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
				Response.Write   "<Cell ss:StyleID='s74'/>"
	
			
				while instr(strGEOIDs, ",") > 0
					strGEOID = left(strGEOIDs, instr(strGEOIDs, ",")-1)
					oRsGEOs.filter = "ModuleID = " & cstr(nModuleID) & " and GEOID=" & strGEOID
					if oRsGEOs.RecordCount > 0 then
						'there is a matching date
						if Len(Trim(oRsGEOs("MRDate").value)) > 0 then
							arrDate = Split(oRsGEOs("MRDate").value,"/")
							Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
						else
							Response.Write   "<Cell ss:StyleID='s101'/>"
						end if
					else
						Response.Write   "<Cell ss:StyleID='s101'/>"
					end if

					strGEOIDs = mid(strGEOIDs, instr(strGEOIDs, ",")+1)
				wend

				while instr(strRegionIDs, ",") > 0
					strRegionID = left(strRegionIDs, instr(strRegionIDs, "|")-1)
					strHPDASHCode = mid(strRegionIDs, instr(strRegionIDs, "|")+1, instr(strRegionIDs, ",")-instr(strRegionIDs, "|")-1)
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) & " and RegionID=" & strRegionID
					if oRsRegions.RecordCount > 0 then
						'there is a matching entry, put a checkmark
						response.write " <Cell ss:StyleID='s77'><Data ss:Type='String'>" & Server.HTMLEncode("ü") & "</Data></Cell>"	
						'response.write " <Cell ss:StyleID='s74'><Data ss:Type='String'>X</Data></Cell>"
					else
						Response.Write   "<Cell ss:StyleID='s74'/>"

					end if

					strRegionIDs = mid(strRegionIDs, instr(strRegionIDs, ",")+1)
				wend

				Response.Write  "</Row>" & vbCrLf
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop

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


sub DLBoxHTML
	dim sStatus
	%>
	<TABLE WIDTH=50% BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
	<TH align=left>Available AMO Status</TH>
	<TH></TH>
	<TH align=left>Selected AMO status</TH>
	</TR>
	
	<TR>
	<TD align=left> <SELECT name='lbxAvailableStatus' id='lbxAvailableStatus' multiple size=7 style='WIDTH: 300px' onDblClick='return btnAddDLBCItem_onclick(AddStatus, lbxAvailableStatus,lbxSelectedStatus)' >
		<%


		do while not oRsAMOstatus.eof
			if sStatus<>oRsAMOstatus("Description")  and trim(oRsAMOstatus("Description")) <> "Obsolete"  then%>
				<OPTION value=<%=oRsAMOstatus("StatusID")%>> <%=oRsAMOstatus("Description")%> </OPTION>
				<%
			end if
			sStatus = oRsAMOstatus("Description")
			oRsAMOstatus.Movenext
		Loop
		%>
		</SELECT> </TD>
			    
	<TD><TABLE border=0 cellPadding=1 cellSpacing=1 width='10%'>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move all Available entries to Selected list' 
		name=AddAllStatus
		onclick='return btnAddAllDLBCItem_onclick(AddAllStatus, lbxAvailableStatus,lbxSelectedStatus)' type=button value='>>'></TD>
	 </TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move selected Available entries to Selected list' 
		name=AddStatus
		onclick='return btnAddDLBCItem_onclick(AddStatus, lbxAvailableStatus,lbxSelectedStatus)' type=button value='->'></TD>
	 </TR>
	 <TR><TD></TD></TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move selected entries to Available list' 
		name=RemoveStatus
		onclick='return btnRemoveDLBCItem_onclick(RemoveStatus,lbxAvailableStatus,lbxSelectedStatus)' type=button value='<-'></TD>
	 </TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move all Selected entries to Available list' 
		name=RemoveAllStatus
		onclick='return btnRemoveAllDLBCItem_onclick(RemoveAllStatus, lbxAvailableStatus,lbxSelectedStatus)' type=button value='<<'></TD>
	 </TR>
	 </TABLE></TD>
	<TD align=left> <SELECT name='lbxSelectedStatus' multiple id='lbxSelectedStatus' size=7 style='WIDTH: 300px' ondblclick='return btnRemoveDLBCItem_onclick(RemoveStatus,lbxAvailableStatus,lbxSelectedStatus)'  ></SELECT> </TD></TR>
	</table>
	<%
end sub






sub DLBoxHTML_BusSeg
	
	%>
	<TABLE WIDTH=50% BORDER=0 CELLSPACING=1 CELLPADDING=1>
	<TR>
	<TH align=left>Available Business Segment</TH>
	<TH></TH>
	<TH align=left>Selected Business Segment</TH>
	</TR>
	
	<TR>
	<TD align=left> <SELECT name='lbxAvailableBusSeg' id='lbxAvailableBusSeg' multiple size=7 style='WIDTH: 300px' onDblClick='return btnAddDLBCItem_onclick(AddBusSeg, lbxAvailableBusSeg,lbxSelectedBusSeg)' >
		<%

		if oRsBusSeg.RecordCount > 0 then
				oRsBusSeg.Filter = "State=1"
				oRsBusSeg.Sort = "Description ASC"
		end if 				
		

	do while not oRsBusSeg.eof
			%>
				<OPTION value=<%=oRsBusSeg("CategoryID")%>> <%=oRsBusSeg("Description")%> </OPTION>
				<%
			
			
			oRsBusSeg.Movenext
		Loop
		%>
		</SELECT> </TD>
			    
	<TD><TABLE border=0 cellPadding=1 cellSpacing=1 width='10%'>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move all Available entries to Selected list' 
		name=AddAllBusSeg
		onclick='return btnAddAllDLBCItem_onclick(AddAllBusSeg, lbxAvailableBusSeg,lbxSelectedBusSeg)' type=button value='>>'></TD>
	 </TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move selected Available entries to Selected list' 
		name=AddBusSeg
		onclick='return btnAddDLBCItem_onclick(AddBusSeg, lbxAvailableBusSeg,lbxSelectedBusSeg)' type=button value='->'></TD>
	 </TR>
	 <TR><TD></TD></TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move selected entries to Available list' 
		name=RemoveBusSeg
		onclick='return btnRemoveDLBCItem_onclick(RemoveBusSeg,lbxAvailableBusSeg,lbxSelectedBusSeg)' type=button value='<-'></TD>
	 </TR>
	 <TR>
	 <TD><INPUT language=javascript style = 'font-family:Courier New'
		title='Move all Selected entries to Available list' 
		name=RemoveAllBusSeg
		onclick='return btnRemoveAllDLBCItem_onclick(RemoveAllBusSeg, lbxAvailableBusSeg,lbxSelectedBusSeg)' type=button value='<<'></TD>
	 </TR>
	 </TABLE></TD>
	<TD align=left> <SELECT name='lbxSelectedBusSeg' multiple id='lbxSelectedBusSeg' size=7 style='WIDTH: 300px' ondblclick='return btnRemoveDLBCItem_onclick(RemoveBusSeg,lbxAvailableBusSeg,lbxSelectedBusSeg)'  ></SELECT> </TD></TR>
	</table>
	<%
end sub





%>
