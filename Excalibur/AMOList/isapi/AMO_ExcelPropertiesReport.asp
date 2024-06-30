<%@ language=vbscript %>
<% 

Response.Buffer = True
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "content-disposition", "inline; filename=AMO_Properties_Report.xls"

%> 

<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<%
Call ValidateSession

dim sDesc, sHeader, sModuleTypeHTML, sModuleDivisionHTML, sCtrlStyle, sHelpFile, strCheckDivisionIDs
dim sCreator, sUpdater, sCreatedDate, sUpdatedDate, sErr, sCategoryDesc, sAMOStatus
dim sRasDisconDate, sCPLBlindDate, sBOMRevADate, sBluePN, sRedPN
dim sNetWeight, sExportWeight, sAirPackedWeight, sAirPackedCubic, sExportCubic, sReplacement, sAlternative
dim sDivisionIDs, sShortDesc, sModuleType, sActualCost, sHWModuleCategoryHTML, sSWModuleCategoryHTML
dim sGroupName, sTmp, sDivisionID, strFrom, strOwnerHTML, strFilter, strDiv, strNotes, sCostCtrlStyle, sOptionCategory
dim sOriginalDivisionIDs, sOptionType, sDivisionTarget
dim nID, nMode, nTypeID, nModuleTypeID, nAMOStatusID
dim lngChangeMask, lngGroupID, lngMOLHide
dim oSvr, oErr
dim oRs, oRsPlatforms, oSelectedRs, oRsCheckedRegion, oRsGroups, oRsHWCategory, oRsSWCategory, oRsUserRights, oRsCreateGroups, oRsRegions
dim bUpdate, bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim bCostCreate, bCostView, bCostUpdate, bCostDelete
dim bFromAMO, bEdit
dim arrDivisions
dim sAMOPartNoRe, intTargetNA, intTargetLA, intTargetEMEA, intTargetAPJ, intTargetAllReg, intBurdenPer
dim intContraPer, intBurden, intContra, intNetRevenue, intDealerMargin, intTargetCostBurden, intGrossMarginPer
dim intGrossMargin, intLifetimeGrossMargin, sJustificationnotes
dim sAMOCost
dim sAMOWWPrice 
dim sManufactureCountry, sWarrantyCode, sObsoleteDate

%>

<?xml version="1.0" encoding="ISO-8859-1"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author></Author>
  <LastAuthor></LastAuthor>
  <Created>2003-08-05T18:30:31Z</Created>
  <Company>HP</Company>
  <Version>10.4219</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <DownloadComponents/>
  <LocationOfComponents HRef="file:///\\ccesam01\Corp\Office2002\"/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>9345</WindowHeight>
  <WindowWidth>11340</WindowWidth>
  <WindowTopX>360</WindowTopX>
  <WindowTopY>60</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s29" ss:Name="Hyperlink">
   <Font ss:Color="#0000FF" ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s25">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:Size="8" ss:Color="#000000" />
   <Interior ss:Color="#99CCFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s26">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s31" ss:Parent="s29">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s33">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:Size="7"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s32">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:Size="7"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s24">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:Size="8" ss:Color="#008000" ss:Bold="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s23">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:Size="12" ss:Color="#FF0000" ss:Bold="1" ss:Italic="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="s50">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:Size="12" ss:Color="#000000" ss:Bold="1"/>
  </Style>
 <Style ss:ID="s60">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:Size="7" ss:Color="#000000" ss:Bold="1" />
   <Interior ss:Color="#99CCFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s61">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Font ss:Size="7" ss:Color="#000000" ss:Bold="1" />
   <Interior ss:Color="#99CCFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s99">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" />
   <Font ss:Size="10" ss:Color="#000000" ss:Bold="1" />
  </Style>
 </Styles>
 
 <%
 
 Call GenerateReport()
 
 Function GenerateReport ()
 
		if Request.QueryString("ID") <> "" then
			nID = clng(Request.QueryString("ID"))
		else
			nID = 0
		end if
		
		if Request.QueryString("TypeID") <> "" then
			nTypeID = clng(Request.QueryString("TypeID"))
		else
			nTypeID = 0
		end if
		
		if Request.QueryString("DivisionTarget") <> "" then
			sDivisionTarget = Request.QueryString("DivisionTarget")
		else
			sDivisionTarget = ""
		end if
		
		'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
        set oSvr = New ISAMO
		'get option data
		set oRs = oSvr.AMOModule_Search(Application("REPOSITORY"), " and O.ModuleID=" & cstr(nID) & " and O.SCMID=1 ", "", null, null)
		if oRs is nothing then
			Response.Write("<Worksheet ss:Name='AMO Properties Report'>")
			Response.Write("<Table>")
			Response.Write("<Column ss:Width='75.0' ss:Span='100'/>")
			Response.Write("<Row>")
			Response.Write("<Cell  ss:MergeAcross='5' ss:StyleID='s23'><Data ss:Type='String'>There are no data found</Data></Cell>")
			Response.Write("</Row>")
			Response.Write("</Table>")
			Response.Write("</Worksheet>") 
			Response.Write("</Workbook>") 
			set oErr = nothing
		else
			if oRs.RecordCount = 0 then
				Response.Write("<Worksheet ss:Name='AMO Report'>")
				Response.Write("<Table>")
				Response.Write("<Column ss:Width='75.0' ss:Span='100'/>")
				Response.Write("<Row>")
				Response.Write("<Cell  ss:MergeAcross='5' ss:StyleID='s23'><Data ss:Type='String'>There are no data found</Data></Cell>")
				Response.Write("</Row>")
				Response.Write("</Table>")
				Response.Write("</Worksheet>") 
				Response.Write("</Workbook>") 
				set oErr = nothing
			else
				' get all property fields
				nTypeID = oRs.Fields("CategoryID").Value
				sDesc = oRs.Fields("Description").Value
				sShortDesc = oRs.Fields("ShortDescription").Value
				sCreator = oRs.Fields("Creator").Value
				sUpdater = oRs.Fields("Updater").Value
				sCreatedDate = oRs.Fields("TimeCreated").Value
				sUpdatedDate = oRs.Fields("TimeChanged").Value
				nModuleTypeID = oRs.Fields("ModuleTypeID").Value
				sModuleType = oRs.Fields("ModuleType").Value
				nAMOStatusID = oRs.Fields("AMOStatusID").Value
				sAMOStatus = oRs.Fields("AMOStatus").Value
				sGroupName = oRs.Fields("GroupName").Value
				sBluePN = oRs.Fields("BluePN").Value
				sRedPN = oRs.Fields("RedPN").Value
				sBOMRevADate = oRs.Fields("BOMRevADate").Value
				sRasDisconDate = oRs.Fields("RASDiscontinueDate").Value
				sCPLBlindDate = oRs.Fields("CPLBlindDate").Value
				sAMOCost = oRs.Fields("AMOCost").Value
				sAMOWWPrice = oRs.Fields("AMOWWPrice").Value
				sActualCost = oRs.Fields("ActualCost").Value
				sNetWeight = oRs.Fields("NetWeight").Value
				sAirPackedWeight = oRs.Fields("AirPackedWeight").Value
				sExportWeight = oRs.Fields("ExportWeight").Value
				sAirPackedCubic = oRs.Fields("AirPackedCubic").Value
				sExportCubic = oRs.Fields("ExportCubic").Value
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
				sManufactureCountry = oRs.Fields("ManufactureCountry").Value
				sWarrantyCode = oRs.Fields("WarrantyCode").Value
				sObsoleteDate = oRs.Fields("ObsoleteDate").Value
				
	
				sVisibility_NA = oRs.Fields("Visibility_NA").value
				sVisibility_EM = oRs.Fields("Visibility_EM").value
				sVisibility_AP = oRs.Fields("Visibility_AP").value
				sVisibility_LA = oRs.Fields("Visibility_LA").value
				
				
				
				if isNull(sNetWeight) then
					sNetWeight = ""
				end if 
				
				if isNull(sAirPackedWeight) then
					sAirPackedWeight = ""
				end if 
				
				if isNull(sExportWeight) then
					sExportWeight = ""
				end if 
				
				if isNull(sAirPackedCubic) then
					sAirPackedCubic = ""
				end if 
				
				if isNull(sExportCubic) then
					sExportCubic = ""
				end if 
				
			
							
			    ' calculate POR details
			    ' fix issue#5327 due to overflow when use clng()
				if isnumeric(intTargetNA) and isnumeric(intTargetLA) and isnumeric(intTargetEMEA) and isnumeric(intTargetAPJ) then
					intTargetAllReg	= cdbl(intTargetNA) + cdbl(intTargetLA) + cdbl(intTargetEMEA) + cdbl(intTargetAPJ)
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
							intLifetimeGrossMargin  = Round(intGrossMargin * intTargetAllReg,2)
						  else
						    intGrossMarginPer = ""
						  end if						  
					    end if			    
					end if					
							
							
				strNotes = oRs.Fields("Notes").value
				lngMOLHide = oRs.Fields("MOLHide").value
				
				
				' get option type
				
				set oRs = Nothing				
				'set oErr = GetMOLCategory(oRs, 10)	
	            set oRs = GetMOLCategory(10)	
	            if oRs is Nothing then
		            Response.Write("Recordset error: oRs")
		            Response.End()
	            end if		
     	
				oRs.Filter = "CategoryTypeID = " & nModuleTypeID						
				sOptionType = oRs("Description").Value
				
			    oRs.Close
			
				if sOptionType = "Hardware" then			
					'get module HW Categories
					'set oErr = GetMOLCategory(oRsHWCategory, 1)
	                set oRsHWCategory = GetMOLCategory(1)	
	                if oRsHWCategory is Nothing then
		                Response.Write("Recordset error: oRsHWCategory")
		                Response.End()
	                end if		

					if not oRsHWCategory.EOF and not oRsHWCategory.BOF then
						oRsHWCategory.Filter = "CategoryID = " & nTypeID
						sOptionCategory = oRsHWCategory.Fields("FullCatDescription").Value
					end if
					oRsHWCategory.Close
					set oRsHWCategory = Nothing	
					
				else
				   
				    'get module HW Categories
					'set oErr = GetMOLCategory(oRsSWCategory, 12)
	                set oRsSWCategory = GetMOLCategory(12)	
	                if oRsSWCategory is Nothing then
		                Response.Write("Recordset error: oRsSWCategory")
		                Response.End()
	                end if		               

					if not oRsSWCategory.EOF and not oRsSWCategory.BOF then
						oRsSWCategory.Filter = "CategoryID = " & nTypeID
						sOptionCategory = oRsSWCategory.Fields("FullCatDescription").Value
					end if
				    oRsSWCategory.close
					set oRsSWCategory = Nothing	
				end if
				
				
				set oSelectedRs = oSvr.AMOPlatforms_Search(Application("REPOSITORY"), nID & "|", 1, 0)			
				oSelectedRs.Filter = "ChangeMask <> 2"
				
				set oRsRegions = oSvr.AMORegions_Search(Application("REPOSITORY"), nID, 1)
				oRsRegions.Filter = "ChangeMask <> 2"
				
				
				
							
	
				Response.Write("<Worksheet ss:Name='AMO Report'>")
				Response.Write("<Table>")
				Response.Write("<Column ss:Index='1' ss:Width='70.0'/>")
				Response.Write("<Column ss:Index='2' ss:Width='120.0'/>")
				Response.Write("<Column ss:Index='3' ss:Width='70.0'/>")
				Response.Write("<Column ss:Index='4' ss:Width='120.0'/>")
				Response.Write("<Column ss:Index='5' ss:Width='70.0'/>")
				Response.Write("<Column ss:Index='6' ss:Width='120.0'/>")			
				Response.Write("<Column ss:Hidden='1' ss:Index='12' ss:Width='260.0'/>")
				Response.Write("<Column ss:Hidden='1' ss:Index='13' ss:Width='450.0'/>")
				Response.Write("<Column ss:Hidden='1' ss:Index='14' ss:Width='190.0'/>")
	
				Response.Write("<Row>")
				Response.Write("<Cell ss:MergeAcross='5' ss:StyleID='s50'><Data ss:Type='String'>After Market Option Report</Data></Cell>")
				Response.Write("</Row>")
				Response.Write("<Row ss:Height='20'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'></Data></Cell></Row>")
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Marketing Description</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sDesc) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Short Description</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sShortDesc) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Option Type</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sOptionType) & "</Data></Cell>")
				Response.Write("</Row>")
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Option Category</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sOptionCategory) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Business Segment</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sDivisionTarget & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>HP Part Number</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sBluePN) & "</Data></Cell>")
				Response.Write("</Row>")
				
				Response.Write("<Row>")				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>PHweb (General) Availability (GA)</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sBOMRevADate) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>End of Manufacturing (EM)</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sRasDisconDate) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Select Availability (SA)</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sCPLBlindDate) & "</Data></Cell>")
				
				Response.Write("</Row>")
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>End of Sales (ES)</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sObsoleteDate) & "</Data></Cell>")
				
				if IsODM = 0 then
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>AMO Cost</Data></Cell>")
					if sAMOCost <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & server.htmlencode(sAMOCost) & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>AMO Price</Data></Cell>")
					if sAMOWWPrice <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & server.htmlencode(sAMOWWPrice) & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if	
				else
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'></Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'></Data></Cell>")
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'></Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'></Data></Cell>")
				end if
				Response.Write("</Row>")
				
				Response.Write("<Row>")
				if IsODM = 0 then
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Actual Cost</Data></Cell>")
					if sActualCost <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sActualCost) & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
				else
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'></Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'></Data></Cell>")
				end if
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Replacement</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sReplacement) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Alternative</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sAlternative) & "</Data></Cell>")
				Response.Write("</Row>")
						
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Net Weight</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sNetWeight) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Export Weight</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sExportWeight) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Air Packed Weight</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sAirPackedWeight) & "</Data></Cell>")
				Response.Write("</Row>")
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Air Packed Cubic</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sAirPackedCubic) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Export Cubic</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sExportCubic) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Owner By</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sGroupName) & "</Data></Cell>")
				Response.Write("</Row>")
				
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Country of Manufacture</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sManufactureCountry) & "</Data></Cell>")
				
				Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Warranty Code</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sWarrantyCode) & "</Data></Cell>")
				
				
				Response.Write("</Row>")
				
				
				Response.Write("<Row ss:Height='5'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'></Data></Cell></Row>")
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:MergeAcross='2' ss:StyleID='s61'><Data ss:Type='String'>Notes</Data></Cell>")
				Response.Write("<Cell  ss:StyleID='s61'><Data ss:Type='String'>Created by</Data></Cell>")
				Response.Write("<Cell  ss:StyleID='s61'><Data ss:Type='String'>Status</Data></Cell>")
				Response.Write("<Cell  ss:StyleID='s61'><Data ss:Type='String'>Last Updated by</Data></Cell>")
				Response.Write("</Row><Row>")			
				Response.Write("<Cell ss:MergeAcross='2' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(strNotes) & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sCreator & " on " & sCreatedDate & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sAMOStatus) & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sUpdater & " on " & sUpdatedDate & "</Data></Cell>")
				Response.Write("<Cell ss:Index='12' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(strNotes) & "</Data></Cell>")				
				Response.Write("</Row>")
				
				Response.Write("<Row ss:Height='30'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'>POR Details</Data></Cell></Row>")
				
				if IsODM = 0 then
					Response.Write("<Row>")
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Lifetime Vol - NA</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intTargetNA) & "</Data></Cell>")
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Contra (%)</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intContraPer) & "</Data></Cell>")
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Net Revenue</Data></Cell>")
					if intNetRevenue <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intNetRevenue & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					Response.Write("</Row>")
					
					Response.Write("<Row>")
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Lifetime Vol - LA</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intTargetLA) & "</Data></Cell>")
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Burden (%)</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intBurdenPer) & "</Data></Cell>")	
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Gross Margin</Data></Cell>")
					if intGrossMargin <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intGrossMargin & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					Response.Write("</Row>")
				
					Response.Write("<Row>")	
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Lifetime Vol - APJ</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intTargetAPJ) & "</Data></Cell>")		
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Burden</Data></Cell>")
					if intBurden <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intBurden & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Gross Margin %</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & intGrossMarginPer & "</Data></Cell>")	
				    Response.Write("</Row>")
				
					Response.Write("<Row>")
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Lifetime Vol - EMEA</Data> </Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intTargetEMEA) & "</Data></Cell>")	
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Contra</Data></Cell>")	
					if intContra <> "" then		 
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intContra & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Cost + Burden</Data></Cell>")
					if intTargetCostBurden <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intTargetCostBurden & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					Response.Write("</Row>")
				
					Response.Write("<Row>")			
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Target Lifetime Vol - Total All Regions</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(intTargetAllReg) & "</Data></Cell>")
			
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Dealer Margin at 6%</Data></Cell>")
					Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & intDealerMargin & "</Data></Cell>")
					
					Response.Write("<Cell ss:StyleID='s60'><Data ss:Type='String'>Lifetime Gross Margin</Data></Cell>")
					if intLifetimeGrossMargin <> "" then
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$" & intLifetimeGrossMargin & "</Data></Cell>")
					else
						Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>$0.00</Data></Cell>")
					end if
					Response.Write("</Row>")
				end if

				Response.Write("<Row ss:Height='5'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'></Data></Cell></Row>")
				Response.Write("<Row>")
				if IsODM = 0 then
					Response.Write("<Cell ss:MergeAcross='4' ss:StyleID='s61'><Data ss:Type='String'>Provide Low Volume or Low Margin Justification</Data></Cell>")
				else
					Response.Write("<Cell ss:MergeAcross='4' ss:StyleID='s61'><Data ss:Type='String'></Data></Cell>")
				end if
				Response.Write("<Cell  ss:StyleID='s61'><Data ss:Type='String'>Replaces AMO PartNo</Data> </Cell>")
				Response.Write("</Row><Row>")
				if IsODM = 0 then
					Response.Write("<Cell ss:MergeAcross='4' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sJustificationnotes) & "</Data></Cell>")
				else
					Response.Write("<Cell ss:MergeAcross='4' ss:StyleID='s33'><Data ss:Type='String'></Data></Cell>")
				end if
				Response.Write("<Cell  ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sAMOPartNoRe) & "</Data></Cell>")
				Response.Write("<Cell  ss:Index='13' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(sJustificationnotes) & "</Data></Cell>")
				Response.Write("</Row>")
	
			
				Response.Write("<Row ss:Height='20'><Cell ss:MergeAcross='3' ss:StyleID='s99'><Data ss:Type='String'>SKU Visibility</Data></Cell></Row>")
				
				Response.Write("<Row>")
				Response.Write("<Cell ss:StyleID='s61'><Data ss:Type='String'>NA</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s61'><Data ss:Type='String'>EMEA</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s61'><Data ss:Type='String'>APJ</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s61'><Data ss:Type='String'>AL</Data></Cell>")
				Response.Write("</Row><Row>")			
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sVisibility_NA & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sVisibility_EM & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sVisibility_AP & "</Data></Cell>")
				Response.Write("<Cell ss:StyleID='s33'><Data ss:Type='String'>" & sVisibility_LA & "</Data></Cell>")
				Response.Write("</Row>")
				
				
				
				
				if oSelectedRs.recordCount > 0 then
					Response.Write("<Row ss:Height='10'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'></Data></Cell></Row>")
					Response.Write("<Row>")
					Response.Write("<Cell ss:MergeAcross='2' ss:StyleID='s61'><Data ss:Type='String'>Target Platforms</Data></Cell>")
					Response.Write("</Row>")
					Do until oSelectedRs.EOF		
						Response.Write("<Row>")
						Response.Write("<Cell ss:MergeAcross='2' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(oSelectedRs.Fields("FullName").value) & "</Data></Cell>")
						Response.Write("<Cell ss:Index='14' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(oSelectedRs.Fields("FullName").value) & "</Data></Cell>")
						Response.Write("</Row>") 
						oSelectedRs.MoveNext
					Loop
					
				end if
				
				if oRsRegions.recordCount > 0 then
					Response.Write("<Row ss:Height='10'><Cell ss:MergeAcross='5' ss:StyleID='s99'><Data ss:Type='String'></Data></Cell></Row>")
					Response.Write("<Row>")
					Response.Write("<Cell ss:Index='2' ss:StyleID='s61'><Data ss:Type='String'>Localization</Data></Cell>")
					Response.Write("</Row>")
					Do until oRsRegions.EOF		
						Response.Write("<Row>")
						Response.Write("<Cell  ss:Index='2' ss:StyleID='s33'><Data ss:Type='String'>" & server.htmlencode(oRsRegions.Fields("Region").value) & "</Data></Cell>")
						Response.Write("</Row>") 
						oRsRegions.MoveNext
					Loop
					
				end if
				
				Response.Write("</Table>")
				Response.Write("</Worksheet>") 
				Response.Write("</Workbook>") 			
		end if
	end if
	
	oSelectedRs.Close
	oRsRegions.Close
 
End Function
	
%>
								