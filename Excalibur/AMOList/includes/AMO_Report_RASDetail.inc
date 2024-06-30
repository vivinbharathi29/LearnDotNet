
<%
function WriteAMO_Report_rasdeTAILXML (byval rsModule, byval SCMID, byVal PublishDate, byVal sBusSegIDs)
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


set oRsRegions = oSvr.AMO_Report_RASDetail_Regions_Search(Application("REPOSITORY"), strModuleIDs,1)
if oRsRegions is nothing then
sErr = "Missing required parameters.  Unable to complete your request."
'response.write sErr
response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error for AMO_SCM</Data></Cell></Row>"
set oErr = nothing
exit function
end if

set oRsPlatforms = oSvr.AMO_SCM_Platforms_Search(Application("REPOSITORY"), strModuleIDs & "|" & sBusSegIDs,SCMID ,oRsPlatforms)
if oRsPlatforms is nothing then
sErr = "Missing required parameters.  Unable to complete your request."
'response.write sErr
response.write "<Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734894'><Data ss:Type='String'>Error searching AMO platforms</Data></Cell></Row>"
		set oErr = nothing
		exit function
	end if
	
	'Hardcoded info
		'Response.write "<Column ss:Width=‘65.25’/>"
   		'Response.write "<Column ss:Width=‘151.5’/>"
   		'Response.write "<Column ss:Width=‘45’/>"
		'Response.write "<Column ss:Width=‘28.5’/>"
		
		Response.write "<Column ss:Width='81.75'/>"
		Response.write " <Column ss:Width='84'/>"
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
			Response.write  "<Column  ss:AutoFitWidth='1'    ss:Span='" & cstr(PlatformCt-1) & "'/>"
		elseif PlatformCt =1 then
			Response.write "<Column ss:AutoFitWidth='1'/>"
		end if 
 


	
			
		Response.write "<Column ss:Width='110.25'/>"
		Response.write "<Column ss:Width='110.25'/>"
		Response.write "<Column ss:Width='77.25'/>"
		Response.write "<Column ss:Width='77.25'/>"
		Response.write "<Column ss:Width='77.25'/>"
		
		'response.write "<Column ss:AutoFitWidth='0' ss:Width='27.75'/>"

		Response.Write  "<Row ss:Height='20.25'><Cell ss:MergeAcross='4' ss:StyleID='m28734884'><Data ss:Type='String'>AMO RAS Detail</Data><NamedCell ss:Name='Print_Titles'/></Cell> </Row>" 
		Response.Write  " <Row ss:Height='15.75'><Cell ss:MergeAcross='4' ss:StyleID='m28734904'><Data ss:Type='String'>Today's Date: " & cstr(now()) & "(All Dates mm/dd/yyyy)</Data><NamedCell   ss:Name='Print_Titles'/></Cell></Row>" 
			
		'Row 1; Heading
		Response.Write  "<Row  ss:Height='75.75' ss:StyleID='s30'>"
		Response.Write  "<Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Blue Part Number</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:StyleID='m19282862'><Data ss:Type='String'>Option (Short Description)</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  "<Cell ss:StyleID='m19282862'><Data ss:Type='String'>Option (Long Description)</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Status</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Red Part Number</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Owned By</Data></Cell>"
		
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
					Response.Write  "<Cell  ss:StyleID='m28734944'><Data ss:Type='String'>" & server.htmlencode(strMarketingName) & "</Data></Cell>"
				  else
					Response.Write  "<Cell  ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
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
				  Response.Write  "<Cell  ss:StyleID='m28734944'><Data ss:Type='String'>" & server.htmlencode(strMarketingName) & "</Data></Cell>"
				else
				  Response.Write  "<Cell  ss:StyleID='m28734944'><Data ss:Type='String'>N/A</Data></Cell>"
				end if
				PlatformCt = PlatformCt + 1
				strPlatformIDs = strPlatformIDs & strtempPlatformID & ","
			end if
			strPlatformIDs_backup = strPlatformIDs
		end if
		'end if platform headers
		
		
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Comment to RAS</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Comment from RAS</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Hide from PRL</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>"
		Response.Write  "<Cell ss:StyleID='m19282862'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>"
		
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>Warranty Code</Data></Cell>"
		Response.Write  " <Cell  ss:StyleID='m19282862'><Data ss:Type='String'>End of Sales (ES)</Data></Cell>"
		Response.Write  "<Cell ss:StyleID='m19282862'><Data ss:Type='String'>Country of Manufacture</Data></Cell>"
		
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>AMO Cost</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>AMO Price</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Actual Cost</Data></Cell>"
		'Response.Write  "<Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Replacement</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Alternative</Data></Cell>"
		'Response.Write  "<Cell  ss:StyleID='m28735056'><Data ss:Type='String'>Net Weight</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Export Weight</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Air Packed Weight</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Air Packed Cubic</Data></Cell>"
		'Response.Write  " <Cell  ss:StyleID='m28734954'><Data ss:Type='String'>Export Cubic</Data></Cell>"
		
		Response.Write "</Row>"
	
		
		'Module Data
		rsModule.MoveFirst
		do while (not rsModule.EOF)
		
		'write Category Heading
			strCurrentSection = rsModule.Fields("CategoryID").Value
			Response.Write  " <Row ss:AutoFitHeight='0' ss:Height='12.75'>"
			Response.write "<Cell ss:MergeAcross='" & cstr(PlatformCt +6) & "' ss:StyleID='m19297342'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("ModuleType").Value) & "-" & server.htmlencode(rsModule.Fields("Category").Value) &  "</Data><NamedCell  ss:Name='Print_Titles'/></Cell>"
   			
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

				Response.Write  " <Row ss:AutoFitHeight='1' ss:StyleID='s70'>"
				Response.Write  "<Cell ss:StyleID='s93'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("BluePN").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("ShortDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("LongDescription").Value) & "</Data><NamedCell ss:Name='Print_Titles'/></Cell>"
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("AMOStatus").Value & "</Data></Cell>"	
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("RedPN").Value) & "</Data></Cell>"	
				Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("GroupName").Value & "</Data></Cell>"	
		
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
			 
			Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & Server.HTMLEncode(rsModule.Fields("InfoComment").Value) & "</Data></Cell>"	
			Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & Server.HTMLEncode(rsModule.Fields("RASComment").Value) & "</Data></Cell>"		
			
				if len(rsModule.Fields("CPLBlindDate").Value)>0 then
					arrDate = Split(rsModule.Fields("CPLBlindDate").Value,"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" & arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
				if len(rsModule.Fields("BOMRevADate").Value)>0 then
					arrDate = Split(rsModule.Fields("BOMRevADate").Value,"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
				if len(rsModule.Fields("RASDiscontinueDate").Value)>0 then
					arrDate = Split(rsModule.Fields("RASDiscontinueDate").Value,"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				
				Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='String'>" & server.htmlencode(rsModule.Fields("WarrantyCode").Value) & "</Data></Cell>"	
				if len(rsModule.Fields("ObsoleteDate").Value)>0 then
					arrDate = Split(rsModule.Fields("ObsoleteDate").Value,"/")
					Response.Write  "<Cell ss:StyleID='s101'><Data ss:Type='DateTime'>" &  arrDate(2) & "-" & arrDate(0) & "-" & arrDate(1) & "</Data></Cell>"
				else
					Response.Write   "<Cell ss:StyleID='s101'/>"

				end if 
				Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='String'>" & rsModule.Fields("ManufactureCountry").Value & "</Data></Cell>"	
				
				
				
		'	Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("AMOCost").Value & "</Data></Cell>"	
		'	Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("AMOWWPrice").Value & "</Data></Cell>"	
		'	Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("ActualCost").Value & "</Data></Cell>"	
		'	Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("Replacement").Value & "</Data></Cell>"	
		'	Response.Write  "<Cell ss:StyleID='s72'><Data ss:Type='String'>" & rsModule.Fields("Alternative").Value & "</Data></Cell>"	
		'	Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("NetWeight").Value & "</Data></Cell>"
		'	Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("ExportWeight").Value & "</Data></Cell>"
		'	Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("AirPackedWeight").Value & "</Data></Cell>"
		'	Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("AirPackedCubic").Value & "</Data></Cell>"
		'	Response.Write  "<Cell ss:StyleID='s74'><Data ss:Type='Number'>" & rsModule.Fields("ExportCubic").Value & "</Data></Cell>"
				
			
			Response.Write  "</Row>" & vbCrLf
				
				'Localization part
				
				
				if oRsRegions.recordcount > 0 then
					oRsRegions.filter = "ModuleID = " & cstr(nModuleID) 			
			
					if oRsRegions.RecordCount > 0 then
						Response.Write "<Row>"
						Response.Write   "<Cell ss:StyleID='s72'/>"
						Response.Write   "<Cell ss:StyleID='s72'/>"
						Response.Write   "<Cell ss:StyleID='s72'/>"
						Response.Write  "<Cell  ss:StyleID='m19297028'><Data ss:Type='String'>GEO</Data></Cell>"
						Response.Write  "<Cell  ss:StyleID='m19297028'><Data ss:Type='String'>Country/Region</Data></Cell>"			
						Response.Write  "<Cell  ss:StyleID='m19297028'><Data ss:Type='String'>Dash Code</Data></Cell>"
						Response.Write "</Row>"
						do while not oRsRegions.eof
							Response.Write "<Row>"
							Response.Write   "<Cell ss:StyleID='s72'/>"
							Response.Write   "<Cell ss:StyleID='s72'/>"
							Response.Write   "<Cell ss:StyleID='s72'/>"
							response.write " <Cell ss:StyleID='s72'><Data ss:Type='String'>" & oRsRegions("CountryGEO")  & "</Data></Cell>"
							response.write " <Cell ss:StyleID='s72'><Data ss:Type='String'>" & oRsRegions("Region")  & "</Data></Cell>"
				
							response.write " <Cell ss:StyleID='s72'><Data ss:Type='String'>" & oRsRegions("HPDASHCode")  & "</Data></Cell>"
							
							Response.Write "</Row>"
							oRsRegions.movenext
						loop	
					end if			
					oRsRegions.filter =""
				
				end if 
				
				'end of localization part
				
				
				
				
				rsModule.MoveNext
				if rsModule.EOF then exit do
			loop
			if rsModule.EOF then exit do
		loop

	
	
		oRsRegions.Close
		set oRsRegions = nothing
		
		oRsPlatforms.Close
		set oRsPlatforms = nothing
		set oErr = nothing
		set oSvr = nothing
	
end function


    
   
  
   
   
   
   
   


%>