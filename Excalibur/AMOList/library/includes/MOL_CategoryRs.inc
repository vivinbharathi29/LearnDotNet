<%
' ========================================================================================
' Script for getting available categories
' ========================================================================================
' function GetMOLCategory(byval lngType)
' This function calls the server object to get the available list of categories/types.
'
' oRs (ADODB recordset): the recordset that contains the available items
' lngType: the category type
'			Possible values for lngType:
'			1 - Module Type
'			2 - Chassis Type
'			3 - Delivery Type
'			4 - 130955
'			5 - Countries that Compaq does business with
'			6 - Sign-off status
'			7 - User Group
'			8 - Request Status
'			9 - MOL division/business segment
'			10 - Module Type
'			11 - Product Year
'			12 - Module SW Category
'			13 - Module SW and HW Category
'			14 - HW Delivery Type
'			15 - SW Delivery Type
'			16 - Module HW Category
'			17 - Module SW and HW Category Group
'			18 - Module HW Category Group
'			19 - Module SW Category Group
'			20 - MOL Operating System
'			21 - MOL OS Supported Level
'			22 - MOL Commitment status list
'			23 - MOL HistoryType
'			24 - AMO Status
'			25 - PAL Category Groups
'			26 - MOL Eng. Responsibility Names
'			27 - MOL Resp. Types
'			28 - AMO Business Segment
'			29 - AMO ProductLine Category
'			30 - AMO Compatibility Business seg
'			31 - Module (HW and SW) Category WITHOUT the word All at the top
' ========================================================================================
function GetMOLCategory(byval lngType)
	dim oSvr

	' call the middle tier to get the selection
	'set oSvr = Server.CreateObject ("JF_S_MOLCategory.ISMOLCategory")
	set oSvr = New ISMOLCategory

	'call the appropriate method to get the different categories
	select case lngType
	case 1	'Module HW Category		
		set GetMOLCategory = oSvr.ModuleCategory_ViewAll(Application("REPOSITORY"), 0)
	case 2	'Chassis Type		
		set GetMOLCategory = oSvr.ChassisType_ViewAll(Application("REPOSITORY"))
	case 3	'All Delivery Type		
		set GetMOLCategory = oSvr.DeliveryType_ViewAll(Application("REPOSITORY"))
	case 4	'130955		
		set GetMOLCategory = oSvr.TB130955_ViewAll(Application("REPOSITORY"))
	case 5	'countries that Compaq does business with
		set GetMOLCategory = oSvr.CPQCountries_ViewAll(Application("REPOSITORY"))
	case 6	'sign-off status
		set GetMOLCategory = oSvr.SignoffStatus_ViewAll(Application("REPOSITORY"))
	case 7	'user group
		set GetMOLCategory = oSvr.UserGroup_ViewAll(Application("REPOSITORY"))
	case 8	'request status
		set GetMOLCategory = oSvr.ModuleStatus_ViewAll(Application("REPOSITORY"))
	case 9  'Module/MOL division
		set GetMOLCategory = oSvr.ModuleDivision_ViewAll(Application("REPOSITORY"))
	case 10  'Module type
		set GetMOLCategory = oSvr.ModuleType_ViewAll(Application("REPOSITORY"))
	case 11  'MOL Product Year
		set GetMOLCategory = oSvr.MOLProductYear_ViewAll(Application("REPOSITORY"))
	case 12	'Module SW Category		
		set GetMOLCategory = oSvr.ModuleCategory_ViewAll(Application("REPOSITORY"), 1)
	case 13	'Module All (HW and SW) Category		
		set GetMOLCategory = oSvr.ModuleCategory_ViewAll(Application("REPOSITORY"), 2)
	case 14	'HW Delivery Type		
		set GetMOLCategory = oSvr.HWDeliveryType_ViewAll(Application("REPOSITORY"))
	case 15	'SW Delivery Type		
		set GetMOLCategory = oSvr.SWDeliveryType_ViewAll(Application("REPOSITORY"))
	case 16	'Module HW Category		
		set GetMOLCategory = oSvr.ModuleCategory_ViewAll(Application("REPOSITORY"), 0)
	case 17	'Module all HW and SW category group
		set GetMOLCategory = oSvr.MOLCategoryGroup_ViewAll(Application("REPOSITORY"), 2)
	case 18	'Module HW category group
		set GetMOLCategory = oSvr.MOLCategoryGroup_ViewAll(Application("REPOSITORY"), 0)
	case 19	'Module SW category group
		set GetMOLCategory = oSvr.MOLCategoryGroup_ViewAll(Application("REPOSITORY"), 1)
	case 20  'MOL Operating System
		set GetMOLCategory = oSvr.MOLOS_ViewAll(Application("REPOSITORY"))
	case 21  'MOL OS Supported Level
		set GetMOLCategory = oSvr.MOLOSSupported_ViewAll(Application("REPOSITORY"))
	case 22 'Commitment status list
		set GetMOLCategory = oSvr.CommitmentStatus_ViewAll(Application("REPOSITORY"))
	case 23 'History Type
		set GetMOLCategory = oSvr.HistoryType_ViewAll(Application("REPOSITORY"))
	case 24 'AMO Status
		set GetMOLCategory = oSvr.AMOStatus_ViewAll(Application("REPOSITORY"))
	case 25	'PAL category group
		set GetMOLCategory = oSvr.PALCategoryGroup_ViewAll(Application("REPOSITORY"), 2)
	case 26	'MOL Eng. Responsibility
		set GetMOLCategory = oSvr.MOLResponsibility_ViewAll(Application("REPOSITORY"))
	case 27	'MOL Resp.Types
		set GetMOLCategory = oSvr.ModuleResponsibilityType_ViewAll(Application("REPOSITORY"))
	case 28  'AMO division
		set GetMOLCategory = oSvr.AMODivision_ViewAll(Application("REPOSITORY"))
	case 29  'AMO ProductLine 
		set GetMOLCategory = oSvr.AMOProductLine_ViewAll(Application("REPOSITORY"))
	case 31	'Module (HW and SW) Category WITHOUT the word All at the top
		set GetMOLCategory = oSvr.ModuleCategory_ViewAll(Application("REPOSITORY"), 3)
	case 32  'MOL Cycle
		set GetMOLCategory = oSvr.MOLCycle_ViewAll(Application("REPOSITORY"))
	case 33	
		set GetMOLCategory = oSvr.FeatureCategory_ViewAll(Application("REPOSITORY"))
	case 34	
		set GetMOLCategory = oSvr.AMOBusinessSegment_ViewAll(Application("REPOSITORY"))
	case else	'unrecognized type
		set oRs = nothing
	end select

	set oSvr = nothing
end function

function GetMOLCategoryByDivision(byval lngType, byval sDivisionIDs)
	dim oSvr

	' call the middle tier to get the selection
	'set oSvr = Server.CreateObject ("JF_S_MOLCategory.ISMOLCategory")
	set oSvr = New ISMOLCategory

	'call the appropriate method to get the different categories
	select case lngType
	case 26	'MOL Eng. Responsibility
		set GetMOLCategoryByDivision = oSvr.MOLResponsibility_ViewAll(Application("REPOSITORY"), sDivisionIDs)
	case 27	'MOL Resp.Types
		set GetMOLCategoryByDivision = oSvr.ModuleResponsibilityType_ViewAll(Application("REPOSITORY"), sDivisionIDs)
	case 30  'AMO Compatibility 
		set GetMOLCategoryByDivision = oSvr.AMOCompatibility_View(Application("REPOSITORY"), sDivisionIDs)
	case else	'unrecognized type
		set oRs = nothing
	end select

	set oSvr = nothing
end function

%>
