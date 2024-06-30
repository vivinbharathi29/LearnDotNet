<%
'***************************************************************
'*  MAKE NOTE OF THE LAST PROCEDURE WHICH IS FOR ACTIVE ITEMS  *
'***************************************************************

' ========================================================================================
' Script for getting available categories
' ========================================================================================
' function GetCategory(byref rs, byval lngType, byval lngID, Optional byval lngBeanType)
' This function calls the server object to get the available list of categories.
'
' Note: Not all lngType values may be used in this function. Some may be used in the
' following procedure which is for Active items only
'
' rs (ADODB recordset): the recordset that contains the available items
' lngType: the category type
'			Possible values for lngType:
'			1 - Hardware category
'			2 - Software category
'			3 - Language
'			4 - OS
'			5 - Delivery Type/Method
'			6 - Hardware Feature
'			7 - Bus Type
'			8 - Test Package Category
'			9 - Test Group
'			10 - Software Feature
'			11 - Vendor
'			12 - Hardware Validations
'			13 - Hardware Regional Support
'			14 - Supported Regions for Product Drop
'			15 - Supported OS for Product Drop
'			16 - Test Dispersal Method
'			17 - Hardware Integrated
'			18 - Distribution Method
'			19 - Division
'			20 - Supported regions for Delivery
'			21 - Supported regions for BOM
'			22 - Supported OS for Delivery
'			23 - Supported OS for BOM
'			24 - PRISM Type
'			25 - Server Type
'			26 - Component Owners
'			27 - Software Status
'			28 - Software Status Validation
'			29 - Get Status
'			30 - Get Media type category
'			31 - Country
'			32 - Region
'			33 - test procedure category
'			34 - Milestone Types
'			35 - OTS Product Drop
'			36 - OTS Domain/Organization
'			37 - Component Developers
'			38 - Supported Language for Product Drop
'			39 - Supported Language for Delivery
'			40 - Supported Language for BOM
'			41,42 - Test Procedure Questions
'			43 - DIB Duplication Request User
'			44 - Get Business Segments for a User
'			45 - Get Business Segments Library for a User
'			46 - Get All CD Titles
'			47 - Get CD Titles for CD Project in Development 
'			48 - Get CD Titles for CD Project Complete
'			49 - Get duplication file shares
'			50 - Get all User Groups for a particular user
'			51 - Get System ROM Categories
'			52 - Get Vendors associated with specified CategoryType
'			53 - Vendors used via CategoryID
'			54 - Vendors used via CategoryType
'			55 - ODM R&D
'			56 - ODM R&D for a particular user
'			57 - PRISM servers in a R&D site
'			58 - OS by Business Segment
'			59 - Language by Business Segment
'			60 - Target Partition for LCM
'			61 - Platform for LCM - old way, this one gets generic name
'			62 - Product Type for LCM
'			63 - SystemID for LCM
'			64 - Release Type for LCM
'			65 - CD Category for LCM
'			66 - View vendor category
'			67 - OS with Release information
'			68 - Softpaq Category for LCM
'			69 - Distinct PCA for HW
'			70 - Get ProductDrop giving Category ID
'			71 - Get Platform without divisions
'			72 - Get ODM R&D without divisions
'			75 - HW PM Owners
' 		76 - Test Management Delivery Types
' 		77 - Test Plan Passes
' 		119- Business Segments
'			161 - Platform for LCM - new way, this one gets marketing name (product master)
' 		171 - Platform for LCM without divisions -  - new way, this one gets marketing name (product master)
'			204 - OS with no OS-Independent listed
'			205 - Risk Level Software Status 
'			206 - Get AMO Group Owner 
'			233 - Test Procedure Category
'			234 - Test Procedure Category For All Hardware Configs
' lngID: the category ID to get bus type or feature
' ========================================================================================
function GetCategory(byval lngType, byval lngID)
	dim oCategory

	' call the middle tier to get the selection
	'set oCategory = Server.CreateObject ("JF_S_WebCategory.ISWebCategory")
	set oCategory = New ISWebCategory
		
	'call the appropriate method to get the different categories
	select case lngType
	case 1	'HW Category		
		set GetCategory = oCategory.wHWCategory_All(Application("Repository"))
	case 2	'SW Category
		set GetCategory = oCategory.wSWCategory_All(Application("Repository"))
	case 3	'Language
		set GetCategory = oCategory.wLGCategory_All(Application("Repository"))
	case 4	'OS
		set GetCategory = oCategory.wOSCategory_All(Application("Repository"))
	case 5	'Delivery Type
		set GetCategory = oCategory.wDLMethod_All(Application("Repository"))
	case 6	'Hardware Feature
		set GetCategory = oCategory.wHWFeature_All2(Application("Repository"), lngID)
	case 7	'Bus Type
		set GetCategory = oCategory.wBusType_All2(Application("Repository"), lngID)
	case 8	'Test Package Category 
		set GetCategory = oCategory.wTPkgCategory_All(Application("Repository"))
	case 9	'Test Group 
		set GetCategory = oCategory.wTestGroup_All(Application("Repository"))
	case 10	'Software Feature 
		set GetCategory = oCategory.wSWFeature_All2(Application("Repository"), lngID)
	case 11	'Vendor
		set GetCategory = oCategory.wVendor_All(Application("Repository"))
	case 12	'Hardware Attributes
		set GetCategory = oCategory.wHWValidation_ViewAll(Application("Repository"), lngID)
	case 13	'Hardware Regional Support
		set GetCategory = oCategory.wHWRegionalSupport_All(Application("Repository"))
	case 14	'Supported regions for Product Drop
		set GetCategory = oCategory.wLGCategory_Supported(Application("Repository"), lngID)
	case 15	'Supported OS
		set GetCategory = oCategory.wOSCategory_Supported(Application("Repository"), lngID)
	case 16	'Test Dispersal Method		
		set GetCategory = oCategory.wDispersalMethod_All(Application("Repository"))
	case 17	'Hardware Integrated
		set GetCategory = oCategory.wHWIntegrated_All(Application("Repository"))
	case 18 'Distribution Methods
		set GetCategory = oCategory.wDistMethod_All(Application("Repository"))
	case 19 'Division
		set GetCategory = oCategory.wDivision_All(Application("Repository"))
	case 20 'Supported regions for Delivery
		set GetCategory = oCategory.wLGCategory_Supported_Delivery(Application("Repository"), lngID)	
	case 21 'Supported regions for BOM
		set GetCategory = oCategory.wLGCategory_Supported_BOM(Application("Repository"), lngID)	
	case 22	'Supported OS for Delivery
		set GetCategory = oCategory.wOSCategory_Supported_Delivery(Application("Repository"), lngID)
	case 23	'Supported OS for BOM
		set GetCategory = oCategory.wOSCategory_Supported_BOM(Application("Repository"), lngID)
	case 24 'PRISM Type
		set GetCategory = oCategory.wPRISMType_All(Application("Repository"))
	case 25 'Server for Deliverables
		set GetCategory = oCategory.wServer_All(Application("Repository"))
	case 26 'Component Owners
		set GetCategory = oCategory.wOwnerCategory_All(Application("Repository"), lngID)
	case 27 'Software Status
		set GetCategory = oCategory.wSWStatus_All(Application("Repository"))
	case 28 'Software Status Validation
		set GetCategory = oCategory.wSWStatusValidation_All(Application("Repository"))
	case 29 'Set Status 
		if (lngID <> 0) then
			set GetCategory = oCategory.wStatus_All(Application("Repository"), lngID)
		else
			set GetCategory = oCategory.wStatus_All(Application("Repository"))
		end if
	case 30 'Media Type
		set GetCategory = oCategory.wMTCategory_All(Application("Repository"))
	case 31 'Country
		set GetCategory = oCategory.wCountry_All(Application("Repository"))
	case 32 '130955 Region 
		set GetCategory = oCategory.w130955Region_All(Application("Repository"))
	case 33 'Test procedure catgory
	    set GetCategory = oCategory.wTprocCategory_All(Application("Repository"))
	case 34 'Milestone Types
	    set GetCategory = oCategory.wMilestoneTypes_All(Application("Repository"), lngID)
	case 35 'OTS Product Drop 
		set GetCategory = oCategory.wOTSDomain_All(Application("Repository"))
	case 36	'Get Category by type
		set GetCategory = oCategory.wCategory_All(Application("Repository"), lngID)
	case 37 'Component Developers
		set GetCategory = oCategory.wUserCategory_All(Application("Repository"), lngID, Application("COMP_DEVELOPER"))
	case 38	'Supported Language for Product Drop
		set GetCategory = oCategory.wLG2Category_Supported(Application("Repository"), lngID)
	case 39	'Supported Language for Delivery
		set GetCategory = oCategory.wLG2Category_Supported_Delivery(Application("Repository"), lngID)
	case 40	'Supported Language for BOM
		set GetCategory = oCategory.wLG2Category_Supported_BOM(Application("Repository"), lngID)
	case 41, 42 'Test Procedure Questions
		set GetCategory = oCategory.wQuestion_All(Application("Repository"))
	case 43 'DIB Duplication Request User
		set GetCategory = oCategory.wDIBRequestUser_All(Application("Repository"), lngID)
	case 44 'Get Business Segments for a User
		set GetCategory = oCategory.wUser_BusinessSegment(Application("Repository"), lngID)
	case 45 'Get User Business Segment Library
		set GetCategory = oCategory.wUser_DivisionLibrary(Application("Repository"), lngID)
	case 46 'Get All CD Titles
		set GetCategory = oCategory.wCDTitle_All(Application("Repository"))
	case 47 'Get All CD Titles for CD Project in Development
		set GetCategory = oCategory.wCDTitle_PrjInDevelop(Application("Repository"))
	case 48 'Get All CD Titles for CD Project Complete
		set GetCategory = oCategory.wCDTitle_PrjComplete(Application("Repository"))
	case 49 'Get duplication file shares
		set GetCategory = oCategory.wFileShare_All(Application("Repository"), lngID)
	case 50 'Get all User Groups for a particular user
		set GetCategory = oCategory.wUser_ViewUserGroups(Application("Repository"), lngID)
	case 51	'System ROM Category
		set GetCategory = oCategory.wROMCategory_All(Application("Repository"))
	case 53	'Vendors used via CategoryID 
		set GetCategory = oCategory.wVendor_CatID_Used(Application("Repository"), lngID)
	case 54	'Vendors used via CategoryType
		set GetCategory = oCategory.wVendor_CatType_Used(Application("Repository"), lngID)
	case 55	'ODMs used via CategoryType
		set GetCategory = oCategory.wODM_All(Application("Repository"))
	case 56 'ODM for a particular user		
		set GetCategory = oCategory.wUser_ODM(Application("Repository"), lngID)
	case 57 'PRISM servers for a R&D site
		set GetCategory = oCategory.wODM_PRISMServers(Application("Repository"), lngID)
	case 58 'OS by Business Segment		
		set GetCategory = oCategory.wOSCategory_All_ByDivision(Application("Repository"), lngID)
	case 59 'Language by Business Segment		
		set GetCategory = oCategory.wLGCategory_All_ByDivision(Application("Repository"), lngID)
	case 60	'Target Partition Category
		set GetCategory = oCategory.wTargetPartition_All(Application("Repository"))
	case 61	'Platform for LCM - old way, this one gets generic name
		set GetCategory = oCategory.wPlatform_All(Application("Repository"))
	case 62	'ProductType for LCM
		set GetCategory = oCategory.wProductType_All(Application("Repository"))
'	case 63	'SystemID for LCM
'		set GetCategory = oCategory.wSystemID_All(Application("Repository"))
	case 64	'Release Type for LCM
		set GetCategory = oCategory.wReleaseType_All(Application("Repository"))
	case 65	'CD Category for LCM
		set GetCategory = oCategory.wCDCategory_All(Application("Repository"))
	case 66	'View vendor - Category
		set GetCategory = oCategory.wVendor_Category(Application("Repository"), lngID)
	case 67	'OS with Release info
		set GetCategory = oCategory.wOSCategory_All_WithRelease(Application("Repository"))
	case 68	'SoftPaq Category for LCM
		set GetCategory = oCategory.wSoftpaqCategory_All(Application("Repository"))
	case 69	'Distinct PCA for HW
		set GetCategory = oCategory.wPlatform_PCA(Application("Repository"), lngID)
	case 70 ' Get product drop
	    set GetCategory = oCategory.wGetProductDrop(Application("Repository"), lngID)
	case 71	'Platform for LCM without divisions
		set GetCategory = oCategory.wPlatform_AllWithoutDivision(Application("Repository"))
	case 72	'ODM R&D without divisions
		set GetCategory = oCategory.wODM_AllWithoutDivision(Application("Repository"))
	case 75 'HW PM Owners
		set GetCategory = oCategory.wUserCategory_All(Application("Repository"), lngID, 2)
	case 76	'Test Management Delivery Types
		set GetCategory = oCategory.wTMDeliveryTypes_All(Application("Repository"))
	case 77	'Test Plan Passes
		set GetCategory = oCategory.wTestPlanPasses_All(Application("Repository"))
	case 119 'Business Segments
		set GetCategory = oCategory.wBusinessSegment_All(Application("Repository"))
	case 161	'Platform for LCM - new way, this one gets marketing name (product master)
		set GetCategory = oCategory.wPlatformMktNameMaster_All(Application("Repository"))
	case 171	'Platform for LCM without divisions -  - new way, this one gets marketing name (product master)
		set GetCategory = oCategory.wPlatformMktNameMaster_AllWithoutDivision(Application("Repository"))
	case 204	'OS with no OS-Independent
		set GetCategory = oCategory.wOSCategory_All_noNone(Application("Repository"))
	case 205 'Risk Level Software Status 
		set GetCategory = oCategory.wSWStatus_RiskLevel(Application("Repository"))
	case 206 'AMO Owner 
		set GetCategory = oCategory.AMOOwner(Application("Repository"))
	case else	'unrecognized type
		set rs = nothing
	end select

	set oCategory = nothing
end function

' lists only active category items
function GetCategoryActive(byval lngType, byval lngID)
	dim oCategory

	' call the middle tier to get the selection
	'set oCategory = Server.CreateObject ("JF_S_WebCategory.ISWebCategory")
	set oCategory = New ISWebCategory
		
	'call the appropriate method to get the different categories
	select case lngType
	case 1	'HW Category
		set GetCategoryActive = oCategory.Active_wHWCategory_All(Application("Repository"))
	case 2	'SW Category
		set GetCategoryActive = oCategory.Active_wSWCategory_All(Application("Repository"))
	case 3	'Language
		set GetCategoryActive = oCategory.Active_wLGCategory_All(Application("Repository"))
	case 4	'OS
		set GetCategoryActive = oCategory.Active_wOSCategory_All(Application("Repository"))
	case 5	'Delivery Type
		set GetCategoryActive = oCategory.wDLMethod_All(Application("Repository"))
	case 6	'Hardware Feature
		set GetCategoryActive = oCategory.Active_wHWFeature_All2(Application("Repository"), lngID)
	case 7	'Bus Type
		set GetCategoryActive = oCategory.Active_wBusType_All2(Application("Repository"), lngID)
	case 8	'Test Package Category 
		' didn't update because test package not used anymore
		set GetCategoryActive = oCategory.Active_wTPkgCategory_All(Application("Repository"))
	case 9	'Test Group 
		set GetCategoryActive = oCategory.wTestGroup_All(Application("Repository"))
	case 10	'Software Feature 
		set GetCategoryActive = oCategory.Active_wSWFeature_All2(Application("Repository"), lngID)
	case 11	'Vendor
		set GetCategoryActive = oCategory.wVendor_All(Application("Repository"))
	case 12	'Hardware Attributes
		set GetCategoryActive = oCategory.wHWValidation_ViewAll(Application("Repository"), lngID)
	case 13	'Hardware Regional Support
		set GetCategoryActive = oCategory.Active_wHWRegionalSupport_All(Application("Repository"))
	case 14	'Supported Language
		set GetCategoryActive = oCategory.wLGCategory_Supported(Application("Repository"), lngID)
	case 15	'Supported OS
		set GetCategoryActive = oCategory.wOSCategory_Supported(Application("Repository"), lngID)
	case 16	'Test Dispersal Method		
		set GetCategoryActive = oCategory.Active_wDispersalMethod_All(Application("Repository"))
	case 17	'Hardware Integrated
		set GetCategoryActive = oCategory.wHWIntegrated_All(Application("Repository"))
	case 18 'Distribution Methods
		set GetCategoryActive = oCategory.Active_wDistMethod_All(Application("Repository"))
	case 19 'Division
		set GetCategoryActive = oCategory.Active_wDivision_All(Application("Repository"))
	case 20 'Supported Language for Delivery
		set GetCategoryActive = oCategory.wLGCategory_Supported_Delivery(Application("Repository"), lngID)
	case 21 'Supported Language for BOM
		set GetCategoryActive = oCategory.wLGCategory_Supported_BOM(Application("Repository"), lngID)
	case 22	'Supported OS for Delivery
		set GetCategoryActive = oCategory.wOSCategory_Supported_Delivery(Application("Repository"), lngID)
	case 23	'Supported OS for BOM
		set GetCategoryActive = oCategory.wOSCategory_Supported_BOM(Application("Repository"), lngID)
	case 24 'PRISM Type
		set GetCategoryActive = oCategory.wPRISMType_All(Application("Repository"))
	case 25 'Server for Deliverables
		set GetCategoryActive = oCategory.wServer_All(Application("Repository"))
	case 26 'Component Owners
		set GetCategoryActive = oCategory.wOwnerCategory_All(Application("Repository"), lngID)
	case 27 'Software Status
		set GetCategoryActive = oCategory.wSWStatus_All(Application("Repository"))
	case 28 'Software Status Validation
		set GetCategoryActive = oCategory.wSWStatusValidation_All(Application("Repository"))
	case 29 'Set Status 
		if (lngID <> 0) then
			set GetCategoryActive = oCategory.wStatus_All(Application("Repository"), lngID)
		else
			set GetCategoryActive = oCategory.wStatus_All(Application("Repository"))
		end if
	case 30 'Software Status Validation
		set GetCategoryActive = oCategory.Active_wMTCategory_All(Application("Repository"))
	case 31 'Country
		set GetCategoryActive = oCategory.wCountry_All(Application("Repository"))
	case 32 '130955 Region 
		set GetCategoryActive = oCategory.w130955Region_All(Application("Repository"))
	case 33 'Test procedure catgory
		set GetCategoryActive = oCategory.Active_wTprocCategory_All(Application("Repository"))
	case 34 'Milestone Types
		set GetCategoryActive = oCategory.wMilestoneTypes_All(Application("Repository"), lngID)
	case 35 'OTS Product Drop 
		set GetCategoryActive = oCategory.wOTSDomain_All(Application("Repository"))
	case 36	'Get Category by type
		set GetCategoryActive = oCategory.Active_wCategory_All(Application("Repository"), lngID)
	case 37 'Component Developers
		set GetCategoryActive = oCategory.wUserCategory_All(Application("Repository"), lngID, Application("COMP_DEVELOPER"))
	case 38	'Supported Language
		set GetCategoryActive = oCategory.wLG2Category_Supported(Application("Repository"), lngID)
	case 39	'Supported Language for Delivery
		set GetCategoryActive = oCategory.wLG2Category_Supported_Delivery(Application("Repository"), lngID)
	case 40	'Supported Language for BOM
		set GetCategoryActive = oCategory.wLG2Category_Supported_BOM(Application("Repository"), lngID)
	case 41, 42 'Test Procedure Questions
		set GetCategoryActive = oCategory.Active_wQuestion_All(Application("Repository"))
	case 51	'System ROM Category
		set GetCategoryActive = oCategory.Active_wROMCategory_All(Application("Repository"))
	case 52	'Vendors associated with a CategoryType
		set GetCategoryActive = oCategory.Active_wVendor_CatType_All(Application("Repository"), lngID)
	case 58 'OS by Business Segment		
		set GetCategoryActive = oCategory.wOSCategory_All_ByDivision(Application("Repository"), lngID)
	case 59 'Language by Business Segment		
		set GetCategoryActive = oCategory.wLGCategory_All_ByDivision(Application("Repository"), lngID)
	case 60	'Target Partition Category
		set GetCategoryActive = oCategory.wTargetPartition_All(Application("Repository"), lngID)
	case 61	'Platform for LCM
		set GetCategoryActive = oCategory.Active_wPlatform_All(Application("Repository"))
	case 62	'ProductType for LCM
		set GetCategoryActive = oCategory.wProductType_All(Application("Repository"), lngID)
'	case 63	'SystemID for LCM
'		set GetCategoryActive = oCategory.wSystemID_All(Application("Repository"), lngID)
	case 64	'Release Type for LCM
		set GetCategoryActive = oCategory.wReleaseType_All(Application("Repository"), lngID)
	case 65	'CD Category for LCM
		set GetCategoryActive = oCategory.wCDCategory_All(Application("Repository"))
	case 67	'OS with Release info
		set GetCategoryActive = oCategory.wOSCategory_All_WithRelease(Application("Repository"))
	case 68	'SoftPaq Category for LCM
		set GetCategoryActive = oCategory.wSoftpaqCategory_All(Application("Repository"))
	case 76	'Test Management Delivery Types
		set GetCategoryActive = oCategory.Active_wTMDeliveryTypes_All(Application("Repository"))
	case 77	'Test Plan Passes
		set GetCategoryActive = oCategory.Active_wTestPlanPasses_All(Application("Repository"))
	case 119 'Active Business Segments
		set GetCategoryActive = oCategory.Active_wBusinessSegment_All(Application("Repository"))
	case 204	'OS with no OS-Independent
		set GetCategoryActive = oCategory.Active_wOSCategory_All_noNone(Application("Repository"))
	case 233 'Test procedure category
		set GetCategoryActive = oCategory.Active_wTprocCategory_All2(Application("Repository"), lngID)
	case 234 'Test procedure category
		set GetCategoryActive = oCategory.Active_wTprocCategory_AllConfig(Application("Repository"), lngID)
	case else	'unrecognized type
		set rs = nothing
	end select

	set oCategory = nothing
end function
%>