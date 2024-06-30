<%
Function DrawImageLocalizationMatrix(ProductVersionID, PlatformName)
	Dim PreRow
	Dim MidRow
	Dim PostRow
	dim strRow
	dim strModelRow
	dim strOSRow
	dim strAppsRow
	dim strTypeRow
	dim strSKURow
	dim strModifiedRow
	dim strStatusRow
	dim strRTMDateRow
	dim strCommentsRow
	dim strHideRow
	dim blnFound
	dim MaxUpdated
	dim strImageIDList
	dim ImageArray
	dim OutArray
	dim strLastGeo
	dim strRegions
	dim ImageDefCount
	dim strLastRegion
	dim i
	dim blnFirst
	dim strRegionDef
	dim ImageCountArray
	dim ImageCountP1Array
	dim ImageCountP2Array
	dim ImageCountP3Array
	dim ImageCountP4Array
	dim ImageCountP5Array
	dim TotalImageCount
	dim strShading
	dim blnInclude
	dim strFilterColor
	dim strShowOption
	dim strTierheader
	dim strLangSub
	dim strTemp
	dim intNativeCount
	dim blnFoundImages
	dim strImagesDisplayed
	dim cm
	dim p
	dim strLangList
	dim strKwl
	dim rs2
	dim strTotalRows
	dim sOutput
		
	strImagesDisplayed = ""
	strModelRow = OpenXlRow("")
	strOSRow = OpenXlRow("")
	strAppsRow = OpenXlRow("")
	strTypeRow = OpenXlRow("")
	strSKURow = OpenXlRow("")
  	strModifiedRow= OpenXlRow("")
	strStatusRow = OpenXlRow("")
	strRTMDateRow = OpenXlRow("")
  	strHideRow= OpenXlRow("")
  	strCommentsRow = OpenXlRow("")
	blnFound = false

	strShowOption = ""	
	TotalImageCount = 0
	ImageDefCount = 0
	strImageIDList = ""
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListImageDefinitionsByProduct"
		
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = ProductVersionID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Report", 3, &H0001)
	p.Value = 1
	cm.Parameters.Append p

	if ProductVersionID = 268 then
		Set p = cm.CreateParameter("@IncludeProductID", 3, &H0001)
		p.Value = 267
		cm.Parameters.Append p
	end if	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	do while not rs.EOF
		blnInclude = true
		if request("Type") <> "" then
			if lcase(trim(rs("ImageType"))) <> lcase(trim(request("TYPE"))) then
				blnInclude = false
			end if
		end if
	
		if request("SW") <> "" then
			if lcase(trim(rs("SWType"))) <> lcase(trim(request("SW"))) then
				blnInclude = false
			end if
		end if

		if request("Brand") <> "" then
			if lcase(trim(rs("Brand"))) <> lcase(trim(request("Brand"))) then
				blnInclude = false
			end if
		end if
	
		if request("OpSys") <> "" then
			if lcase(trim(rs("OS"))) <> lcase(trim(request("OpSys"))) then
				blnInclude = false
			end if
		end if

		if request("Status") <> "" then
			if lcase(trim(rs("Status"))) <> lcase(trim(request("Status"))) then
				blnInclude = false
			end if
		end if

		if request("RTMDate") <> "" then
			if lcase(trim(rs("RTMDate") & "")) <> lcase(trim(request("RTMDate"))) then
				if not(lcase(trim(request("RTMDate"))) = "none" and lcase(trim(rs("RTMDate") & "")) = "") then
					blnInclude = false
				end if
			end if
		end if
	
	
		if blnInclude then
			strImageIDList = strImageIDList & "," & rs("ID")
			strModelRow = strModelRow & AddXlCell(rs("Brand"), "ILMCell")
			strOSRow = strOSRow & AddXlCell(rs("OS"), "ILMCell")
			strAppsRow = strAppsRow & AddXlCell(rs("SWType"), "ILMCell")
			strStatusRow = strStatusRow & AddXlCell(rs("Status"), "ILMCell")
			strRTMDateRow = strRTMDateRow & AddXlCell(rs("RTMDate"), "ILMCell")
			strTypeRow = strTypeRow & AddXlCell(rs("ImageType"), "ILMCell")

			strSKURow = strSKURow & AddXlCell(rs("SKUNumber"), "ILMCell")
			strCommentsRow = strCommentsRow & AddXlCell(rs("Comments"), "ILMCell")
			strModifiedRow = strModifiedRow & DrawXlCell(rs.Fields("Modified"), "ILMCell")
			strHideRow = strHideRow & AddXlCell("Hide", "ILMCell")
			if MaxUpdated = "" then
				MaxUpdated = formatdatetime(rs("Modified"),vbshortdate)
			elseif datediff("d",MaxUpdated,rs("Modified")) > 0 then
				MaxUpdated = formatdatetime(rs("Modified"),vbshortdate)
			end if
			blnFound= true
			'strTotalCells = strTotalCells & "<TD><font size=2 face=verdana></font></TD>"
			ImageDefCount = ImageDefCount + 1
		end if
		rs.movenext
	loop
	rs.Close
	if len(strImageIDList) > 0 then
		strImageIDList = mid(strImageIDList,2) 'Strip comma
	end if

	ImageArray = split (strImageIDList,",")
	Redim OutArray(Ubound(ImageArray))
	Redim PriorityArray(Ubound(ImageArray))
	Redim ImageCountArray(Ubound(ImageArray))
	Redim ImageCountP1Array(Ubound(ImageArray))
	Redim ImageCountP2Array(Ubound(ImageArray))
	Redim ImageCountP3Array(Ubound(ImageArray))
	Redim ImageCountP4Array(Ubound(ImageArray))
	Redim ImageCountP5Array(Ubound(ImageArray))

	
	
	if blnFound then
		strLastGeo = ""
		strLastRegion = ""
		for i = lbound(OutArray) to ubound(OutArray)
			OutArray(i) = ""
			PriorityArray(i) = ""
			ImageCountP1Array(i) = 0
			ImageCountP2Array(i) = 0
			ImageCountP3Array(i) = 0
			ImageCountP4Array(i) = 0
			ImageCountP5Array(i) = 0
			ImageCountArray(i) = 0
		next

		blnFirst = true
		
		dim strSKUBGColor
		
		strLangSub = ""
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListImagesForProduct"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = ProductVersionID
		cm.Parameters.Append p
	
		if ProductVersionID = 268 then
			Set p = cm.CreateParameter("@IncludeID", 3, &H0001)
			p.Value = 267
			cm.Parameters.Append p
		end if	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.open "spListImagesForProduct " & DisplayedID,cn,adOpenForwardOnly
		do while not rs.EOF	
			if trim(strlastRegion) <> trim(rs("ID")) then
				if blnFirst then
					blnFirst = false
					strLastRegion = rs("ID")
				else
					'Output row 
				'	strSKUBGColor = "cornsilk"
					blnFoundImages = false
					for i = lbound(PriorityArray) to ubound(PriorityArray)
						if trim(PriorityArray(i)) <> "" then
							blnFoundImages = true
							exit for
						end if
					next
					if blnFoundImages then
						strRegions = strRegions & OpenXlRow("")
						for i = lbound(PriorityArray) to ubound(PriorityArray)
							if trim(PriorityArray(i)) <> "" then
								Select case trim(PriorityArray(i))
								case "1"
									ImageCountP1Array(i) = ImageCountP1Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "2"
									ImageCountP2Array(i) = ImageCountP2Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "3"
									ImageCountP3Array(i) = ImageCountP3Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "4"
									ImageCountP4Array(i) = ImageCountP4Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								case "5"
									ImageCountP5Array(i) = ImageCountP5Array(i) + 1
									TotalImageCount = TotalImageCount + 1
									ImageCountArray(i) = ImageCountArray(i) + 1
									strSKUBGColor = "white"
								end select							
							else
								strSKUBGColor = "white"			
							end if
							if request("Display") <> "SKU" then					
								strRegions = strRegions & AddXlCell(OutArray(i), "ILMCell")
							else
								strRegions = strRegions & AddXlCell("SKU", "ILMCell")
							end if
						next
						strRegions = strRegions & strRegionDef & CloseXlRow()
					end if		
									
					'reset row array
					for i = lbound(OutArray) to ubound(OutArray)
						OutArray(i) = ""
						PriorityArray(i) = ""
					next
					
					'Remember Last Region
					strLastRegion = rs("ID")
					strTemp = ""
					intNativeCount = 0
					strRegionDef = ""
				end if
			end if

			if rs("Geo") & "" <> strLastGeo then
				strRegions = strRegions & OpenXlRow("") & AddXlCellSpan(rs("Geo") , "Region", ImageDefCount + 4) & CloseXlRow()
				strLastGeo = rs("Geo") & "" 
			end if
			
			for i = lbound(ImageArray) to ubound(ImageArray)
				if trim(ImageArray(i)) = trim(rs("ImageDefID")) then
					if trim(rs("Priority") & "" ) = "" then
						OutArray(i) = ""
						PriorityArray(i) = ""
					else
					strImagesDisplayed=strImagesDisplayed & "," & rs("ImageID")

						if isnumeric(rs("Priority")) then
							OutArray(i) = rs("Priority") & ""
							intNativeCount = intNativeCount + 1
						else
							OutArray(i) = replace(rs("Priority")," " ,"") & " Image "						
							if instr(strTemp,replace(rs("Priority")," " ,"") )=0 then
								set rs2 = server.CreateObject("ADODB.recordset")
							
								rs2.Open "spGetImageLanguage4Sub " & rs("ImageDefID") & ",'" & replace(rs("Priority")," " ,"")  & "'" ,cn,adOpenForwardOnly
								if rs2.EOF and rs2.BOF then
									strTemp = strTemp & "&#10;(" & replace(rs("Priority")," " ,"") & ")"
								else
									strLangList = rs2("OSLanguage")
									if rs2("OtherLanguage") & "" <> "" then
										strLangList = strLangList & "," & rs2("OtherLanguage")
									end if
									strTemp = strTemp & "&#10;" & replace(rs("Priority")," " ,"") & " (" & strLangList & ")"
								end if
								rs2.Close
								set rs2=nothing
							end if
						end if
						PriorityArray(i) = rs("Priority")
					end if
					exit for
				end if
			next
				strRegionDef = AddXlCell(rs("Name"), "ILMCell") & AddXlCell(rs("OptionConfig"), "ILMCell") & AddXlCell(rs("Dash"), "ILMCell")
				strLangList = rs("OSLanguage")
				if trim(rs("OtherLanguage") & "") <> "" then
					strLangList = strLangList & "," & trim(rs("OtherLanguage") & "")
				end if
				if trim(strTemp) = "" then
					strRegionDef = strRegionDef & AddXlCell(replace(rs("Dash")," ","") & " (" & strLangList & ")", "ILMCell")
				else
					if intNativeCount > 0 then
						strRegionDef = strRegionDef & AddXlHtmlCell("<Font html:Color=""#000000"">" & replace(rs("Dash")," ","") & " (" & strLangList & ")" & strTemp & "</Font>", "ILMCell", 1)
					else
						strRegionDef = strRegionDef & AddXlCell(mid(strTemp,5), "ILMCell")
					end if
				end if
				if rs("KWL") & "" <> "" then
					strKWL = rs("KWL") &  ""
				else
					strKWL = ""
				end if
				strLangSub = ""
			rs.MoveNext
		loop
		rs.Close

		strRegions = strRegions & OpenXlRow("0")
		for i = lbound(PriorityArray) to ubound(PriorityArray)
			if trim(PriorityArray(i)) <> "" then
				Select case trim(PriorityArray(i))
				case "1"
					ImageCountP1Array(i) = ImageCountP1Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "2"
					ImageCountP2Array(i) = ImageCountP2Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "3"
					ImageCountP3Array(i) = ImageCountP3Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "4"
					ImageCountP4Array(i) = ImageCountP4Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				case "5"
					ImageCountP5Array(i) = ImageCountP5Array(i) + 1
					TotalImageCount = TotalImageCount + 1
					ImageCountArray(i) = ImageCountArray(i) + 1
				end select							
			end if
			strRegions = strRegions & AddXlCell(OutArray(i), "ILMCell")
		next
		strRegions = strRegions & strRegionDef & CloseXlRow()

	strModelRow = strModelRow & AddXlCell("Model", "ILMCell") & AddXlCellSpanBoth("Total Image Count: " & TotalImageCount, "ILMCell", 3, 7) & CloseXlRow()
	strOSRow = strOSRow & AddXlCell("OS", "ILMCell") & CloseXlRow()
	strAppsRow = strAppsRow & AddXlCell("Apps Bundle", "ILMCell") & CloseXlRow()
	strTypeRow = strTypeRow & AddXlCell("BTO/CTO", "ILMCell") & CloseXlRow()
	strStatusRow = strStatusRow & AddXlCell("Status", "ILMCell") & CloseXlRow()
	strRTMDateRow = strRTMDateRow & AddXlCell("RTM Date", "ILMCell") & CloseXlRow()
	strSKURow = strSKURow & AddXlCell("Image #", "ILMCell") & CloseXlRow()
	strCommentsRow = strCommentsRow & AddXlCell("Comments", "ILMCell") & CloseXlRow()
	strModifiedRow = strModifiedRow & AddXlCell("Modified", "ILMCell") & CloseXlRow()

	strTotalRows = OpenXlRow("")
	strTierHeader = OpenXlRow("")
	for i = lbound(ImageCountArray) to ubound(ImageCountArray)
		strTotalRows = strTotalRows & AddXlCell(ImageCountArray(i), "ILMCell")
		strTierHeader = strTierHeader & AddXlCell("Tier", "Hdr")
	next
	strTierHeader = strTierHeader & AddXlCell("Localization", "Hdr") & AddXlCell("HP Code", "Hdr") & AddXlCell("Dash Code", "Hdr") & AddXlCell("Images", "Hdr") & CloseXlRow()
	
	for i = 1 to ImageDefCount + 4
		sOutput = sOutput & AddColumn(75)
	next
	
	sOutput = sOutput & OpenXlRow("7.5")
	sOutput = sOutput & AddXlCellSpan("", "TitleBorder",ImageDefCount + 4)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("24")
	sOutput = sOutput & AddXlCellSpan(PlatformName & " Image Localization Matrix", "TitleText",ImageDefCount + 4)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("7.5")
	sOutput = sOutput & AddXlCellSpan("", "TitleBorder",ImageDefCount + 4)
	sOutput = sOutput & CloseXlRow()
	sOutput = sOutput & OpenXlRow("7.5")
	sOutput = sOutput & CloseXlRow()

	sOutput = sOutput & strModelRow
	sOutput = sOutput & strOSRow
	sOutput = sOutput & strAppsRow
	sOutput = sOutput & strTypeRow
	sOutput = sOutput & strStatusRow
	sOutput = sOutput & strRTMDateRow
	sOutput = sOutput & strCommentsRow
	sOutput = sOutput & strSKURow
	sOutput = sOutput & strModifiedRow
	sOutput = sOutput & OpenXlRow("5") & AddXlCellSpan("", "TitleBorder", ImageDefCount + 4) & CloseXlRow()
	sOutput = sOutput & strTierHeader
	sOutput = sOutput & strRegions
	sOutput = sOutput & strTotalRows & AddXlCell("Tier x", "ILMCell") & CloseXlRow()

	sOutput = sOutput & OpenXlRow("")
	for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
		sOutput = sOutput &  AddXlCell(ImageCountP1Array(i), "ILMCell")
	next
	sOutput = sOutput &  AddXlCell("Tier 1", "ILMCell") & CloseXlRow()

	sOutput = sOutput & OpenXlRow("")
	for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
		sOutput = sOutput & AddXlCell(ImageCountP2Array(i), "ILMCell")
	next
	sOutput = sOutput & AddXlCell("Tier 2", "ILMCell") & CloseXlRow()

	sOutput = sOutput & OpenXlRow("")
	for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
		sOutput = sOutput & AddXlCell(ImageCountP3Array(i), "ILMCell")
	next
	sOutput = sOutput & AddXlCell("Tier 3", "ILMCell") & CloseXlRow()

	sOutput = sOutput & OpenXlRow("")
	for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
		sOutput = sOutput & AddXlCell(ImageCountP4Array(i), "ILMCell")
	next
	sOutput = sOutput & AddXlCell("Tier 4", "ILMCell") & CloseXlRow()

	sOutput = sOutput & OpenXlRow("")
	for i = lbound(ImageCountP1Array) to ubound(ImageCountP1Array)
		sOutput = sOutput & AddXlCell(ImageCountP5Array(i), "ILMCell")
	next
	sOutput = sOutput & AddXlCell("Tier 5", "ILMCell") & CloseXlRow()
	end if

	set rs = nothing

	DrawImageLocalizationMatrix = sOutput
End Function
%>

