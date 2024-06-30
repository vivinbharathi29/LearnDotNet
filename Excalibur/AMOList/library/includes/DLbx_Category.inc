<%
' ========================================================================================
' Script for Base Category Dual Listbox
' Note: Please do not call this function in other page. It should be called only by
' functions in this page. Also, please do not make multiple copies of it.
' ========================================================================================
function BaseDLbxCat_GetHTML(byval strName, byref rsSel, byval sDescS, byval sIDS, _
			byval bExcluded, byval lngType, byval lngCatID, byval bDisableButtons, byval bActiveOnly, _
			byval bSelList, byval bNumeric, byval sSelList, byref oErr, byval bHorizontal, _
			bCentered, intPadSpaces, byVal sLeftHandler, byVal sRightHandler)
	dim sDescA, sIDA, sAHeading, sSHeading
	dim rsAvail
	dim lngHeight, lngWidth, lngFontHeight, arrlngCatID

	'initialize values
	sDescA = "Description"
	lngHeight = 7
	lngWidth = 300
	lngFontHeight = 14

	'response.end
	'change the ID field name and the headings of the list boxes
	'depending on the category type
	
	select case lngType
	case 1	'HW Category
		sIDA = "CategoryID"
		sAHeading = "Available Hardware Categories"
		sSHeading = "Selected Hardware Categories"
	case 101	'HW Category with horizontal scroll bar
		sIDA = "CategoryID"
		sAHeading = "Available Hardware Categories"
		sSHeading = "Selected Hardware Categories"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 1
	case 2	'SW Category
		sIDA = "CategoryID"
		sAHeading = "Available Software Categories"
		sSHeading = "Selected Software Categories"
	case 102	'SW Category with horizontal scroll bar
		sIDA = "CategoryID"
		sAHeading = "Available Software Categories"
		sSHeading = "Selected Software Categories"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 2
	case 3	'Language
		sIDA = "LanguageID"
		sAHeading = "Available Languages"
		sSHeading = "Selected Languages"
		sDescA = "LongDescription"
	case 103	'Language with horizontal scroll bar
		sIDA = "LanguageID"
		sAHeading = "Available Languages"
		sSHeading = "Selected Languages"
		sDescA = "LongDescription"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 3
	case 4	'OS
		sIDA = "OSID"
		sAHeading = "Available Operating Systems"
		sSHeading = "Selected Operating Systems"
	case 104	'OS with horizontal scroll bar
		sIDA = "OSID"
		sAHeading = "Available Operating Systems"
		sSHeading = "Selected Operating Systems"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 4
	case 204	'OS with horizontal scroll bar - no OS-Independent language
		sIDA = "OSID"
		sAHeading = "Available Operating Systems"
		sSHeading = "Selected Operating Systems"
		lngHeight = 115 'pixels
		lngWidth = 300
	case 5	'Delivery Type
		sIDA = "DeliveryTypeID"
		sAHeading = "Available Delivery Methods"
		sSHeading = "Selected Delivery Methods"
	case 6	'HW Feature
		sIDA = "FeatureID"
		sAHeading = "Available Hardware Features"
		sSHeading = "Selected Hardware Features"
	case 106	'HW Feature with horizontal scroll bar
		sIDA = "FeatureID"
		sAHeading = "Available Hardware Features"
		sSHeading = "Selected Hardware Features"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 6
	case 7	'Bus type
		sIDA = "HWValidationID"
		sAHeading = "Available Bus Types"
		sSHeading = "Selected Bus Types"
	case 107	'Bus type with horizontal scroll bar
		sIDA = "HWValidationID"
		sAHeading = "Available Bus Types"
		sSHeading = "Selected Bus Types"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 7
	case 9	'Test Group
		sIDA = "TestGroupID"
		sAHeading = "Available Test Groups"
		sSHeading = "Selected Test Groups"
	case 109	'Test Group with horizontal scroll bar
		sIDA = "TestGroupID"
		sAHeading = "Available Test Groups"
		sSHeading = "Selected Test Groups"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 9
	case 10	'Software Feature
		sIDA = "FeatureID"
		sAHeading = "Available Software Features"
		sSHeading = "Selected Software Features"
	case 110	'Software Feature with horizontal scroll bar
		sIDA = "FeatureID"
		sAHeading = "Available Software Features"
		sSHeading = "Selected Software Features"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 10
	case 11, 53, 54, 66	'Vendor
		sIDA = "VendorID"
		sAHeading = "Available Vendor"
		sSHeading = "Selected Vendor"
	case 12	'HW Validations
	case 13	'HW Regional Support
		sIDA = "CategoryID"
		sAHeading = "Available Hardware Regional Support"
		sSHeading = "Selected Hardware Regional Support"
	case 14, 20, 21 'Supported Languages for Product Drop, Delivery, BOM
		sIDA = "LanguageID"
		sAHeading = "Available DASH codes"
		sSHeading = "Selected DASH codes"
		sDescA = "LongDescription"
		lngHeight = 9
		lngWidth = 175
	case 17	'HW Integrated
		sIDA = "PassID"
		sAHeading = "Available Integrated Hardware"
		sSHeading = "Selected Integrated Hardware"
	case 15, 22, 23 'Supported OS for Product Drop, Delivery, BOM
		sIDA = "OSID"
		sAHeading = "Available Operating Systems"
		sSHeading = "Selected Operating Systems"
		lngWidth = 175
	case 26, 70	'Component Owners
		sDescA = "Owner"
		sAHeading = "Available Owners"
		sSHeading = "Selected Owners"
		if lngType = 26 then
			sIDA = "OwnerID"
		elseif lngType = 70 then
			sIDA = "ID"
			lngType = 26
		end if
	case 27	'Software Status
	case 28	'Software Status Validation
		sIDA = "ValidationID"
		sDescA = "DualName"
		sAHeading = "Available Software Status"
		sSHeading = "Selected Software Status"
	case 29	'CD Build Status
		sIDA = "StatusID"
		sAHeading = "Available Status"
		sSHeading = "Selected  Status"
	case 31	'Country
		sIDA = "CountryID"
		sAHeading = "Available Countries"
		sSHeading = "Selected Countries"
		sDescA = "LongDescription"
	case 32	'130955 Region
		sIDA = "RegionID"
		sAHeading = "Available Regions"
		sSHeading = "Selected Regions"
		sDescA = "LongDescription"
	case 33 ' Test procedure category
		sIDA = "CategoryID"
		sAHeading = "Available Test Procedure Categories"
		sSHeading = "Selected Test Procedure Categories"
	case 133 ' Test procedure category with horizontal scroll bar
		sIDA = "CategoryID"
		sAHeading = "Available Test Procedure Categories"
		sSHeading = "Selected Test Procedure Categories"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 33
	case 233 ' Test procedure category with horizontal scroll bar
		sIDA = "CategoryID"
		sAHeading = "Available Test Procedure Categories"
		sSHeading = "Selected Test Procedure Categories"
		lngHeight = 115 'pixels
		lngWidth = 300
		lngType = 233
	case 34 ' Milestone Types
		sIDA = "MilestoneID"
		sAHeading = "Available Milestone Types"
		sSHeading = "Selected Milestone Types"
	case 37	'Component Developers
		sIDA = "ID"
		sDescA = "Owner"
		sAHeading = "Available Developers"
		sSHeading = "Selected Developers"
	case 38,39,40 'Supported Languages for Product Drop, Delivery, BOM
		sIDA = "LanguageID"
		sAHeading = "Available Languages"
		sSHeading = "Selected Languages"
		sDescA = "LongDescription"
		'lngHeight = 9
		'lngWidth = 175
	case 41
		sIDA = "QuestionID"
		sAHeading = "Available Questions"
		sSHeading = "Selected Questions"
		sDescA = "Question"
		lngHeight = 100 'pixels
		lngWidth = 400
	case 42	'same as 41 but different width
		sIDA = "QuestionID"
		sAHeading = "Available Questions"
		sSHeading = "Selected Questions"
		sDescA = "Question"
		lngHeight = 100 'pixels
		lngWidth = 300
	case 46, 47, 48	'CD Titles for All, In Development, Complete
		sIDA = "BuildID"
		sAHeading = "Available CD Titles"
		sSHeading = "Selected CD Titles"
		sDescA = "CDTitle"
		
	case 75	'HW PM Owners
		sDescA = "Owner"
		sAHeading = "Available PM"
		sSHeading = "Selected PM"		
		sIDA = "ID"
		
	end select	

		
	'get the available categories

	arrlngCatID = Empty
	arrlngCatID = Split(lngCatID, ",")

	if bActiveOnly then
	' get category for multiple hardware configs
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
		if Not isEmpty(arrlngCatID) and UBound(arrlngCatID) > 0 Then
			'set oErr = GetCategoryActive(rsAvail, 234, lngCatID)
			set rsAvail = GetCategoryActive(234, lngCatID)	
		else
			'set oErr = GetCategoryActive(rsAvail, lngType, clng(lngCatID))
			set rsAvail = GetCategoryActive(lngType, clng(lngCatID))
		end if
	else
		'set oErr = GetCategory(rsAvail, lngType, lngCatID)
		set rsAvail = GetCategoryActive(lngType, clng(lngCatID))
	end if

	'check if any error happens
	if rsAvail is Nothing then
		'handle error
	    BaseDLbxCat_GetHTML = "<H1>Error<BR>"
	else
		'call the common dual list box script to get the HTML string
		BaseDLbxCat_GetHTML = DualListboxRs_GetHTML7(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, _
			bSelList, bNumeric, sSelList, sAHeading, sSHeading, strName, bExcluded, _
			lngHeight, lngWidth, bDisableButtons, bHorizontal, bCentered, intPadSpaces, _
			lngFontHeight, sLeftHandler, sRightHandler )
	end if
	
	'free objects
	set rsAvail = nothing
end function

' ========================================================================================
' The following functions return the HTML string that build the dual list box 
' for available and selected categories.
'
' strName (string): prefix for the list box name
' rsSel (ADODB recordset): the recordset to fill the selected items
' sDescS (string): the description field name
' sIDS (string): the ID field name
' bExcluded: the flag to specify whether to exclude the selected items
'			 from available box
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
'			14 - Supported regions for Product Drop
'			15 - Supported OS for Product Drop
'			17 - Hardware Integrated
'			20 - Supported regions for Delivery
'			21 - Supported regions for BOM
'			22 - Supported OS for Delivery
'			23 - Supported OS for BOM
'			24 - Component Owners
'			27 - Software Status
'			28 - Software Status Validation
'			31 - Country
'			32 - 130955 Region
'           33 - Test procedure Category
'           34 - Milestone Types
'			38 - Supported languages for Product Drop
'			39 - Supported languages for Delivery
'			40 - Supported languages for BOM
'			43 - CD Build status
'			46 - CD Title
' lngCatID (long): get the list for a specific category ID
' bDisabledButtons: the flag to initially disable buttons in the middle of dual list box
' bActiveOnly: the flag to get only active items
' ========================================================================================
function DLbxCat_GetHTML(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, byref oErr)
	DLbxCat_GetHTML = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, 0, False, False, False, False, "", oErr, false, false, 0, "", "")
end function

'with double click error handler on list box items
function DLbxCat_GetHTML2(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, byref oErr)
	DLbxCat_GetHTML2 = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, 0, False, False, False, False, "", oErr, false, false, 0, "", "")
end function

' lists only active category items
function DLbxCat_GetHTML2Active(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, byref oErr)
	DLbxCat_GetHTML2Active = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, 0, False, True, False, False, "", oErr, false, false, 0, "", "")
end function

' allows horizontal scroll bar to be passed
function DLbx_Category_GetHTML3Active(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, _
		byref oErr, byval bHorizontal, byval bCentered, byval intPadSpaces)
	DLbx_Category_GetHTML3Active = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, 0, False, True, False, False, "", oErr, bHorizontal, bCentered, intPadSpaces, "", "")
end function

' allows horizontal scroll bar to be passed
function DLbx_Category_GetHTML3Active2(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, _
		byref oErr, byval bHorizontal, byval bCentered, byval intPadSpaces, byval sTCfgID)
	DLbx_Category_GetHTML3Active2 = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, sTCfgID, False, True, False, False, "", oErr, bHorizontal, bCentered, intPadSpaces, "", "")
end function

' passes handlers for left and right side of the dual list box
function DLbx_Category_GetHTML4Active(byval strName, byref rsSel, byval sDescS, byval sIDS, byval bExcluded, byval lngType, _
		byref oErr, byval bHorizontal, byval bCentered, byval intPadSpaces, byVal sLeftHandler, byVal sRightHandler)
	DLbx_Category_GetHTML4Active = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, 0, False, True, False, False, "", oErr, bHorizontal, bCentered, intPadSpaces, sLeftHandler, sRightHandler)
end function


'should also get double click error handler
function DLbxCat_GetHTML_CID(byval strName, byref rsSel, byval sDescS, byval sIDS, _
			byval bExcluded, byval lngType, byval lngCatID, byref oErr)
    DLbxCat_GetHTML_CID = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, lngCatID, False, False, False, False, "", oErr, false, false, 0, "", "")
end function

'with double click error handler on list box items
function DLbxCat_GetHTML_CID2(byval strName, byref rsSel, byval sDescS, byval sIDS, _
			byval bExcluded, byval lngType, byval lngCatID, byref oErr)
		'call the common dual list box script to get the HTML string
	  	DLbxCat_GetHTML_CID2 = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, lngCatID, False, False, False, False, "", oErr, false, false, 0, "", "")
end function

' lists only active category items
function DLbxCat_GetHTML_CID2Active(byval strName, byref rsSel, byval sDescS, byval sIDS, _
			byval bExcluded, byval lngType, byval lngCatID, byref oErr)

  	DLbxCat_GetHTML_CID2Active = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, lngCatID, False, True, False, False, "", oErr, false, false, 0, "", "")
end function

' provide flags to get active items, to disable middle buttons and to exclude items in a list
function DLbxCat_GetHTML_CID3(byval strName, byref rsSel, byval sDescS, byval sIDS, _
			byval bExcluded, byval lngType, byval lngCatID, byval bDisabledButtons, byval bActiveOnly, _
			byval bSelList, byval bNumeric, byval sSelList, byref oErr)

  	DLbxCat_GetHTML_CID3 = BaseDLbxCat_GetHTML(strName, rsSel, sDescS, sIDS, _
			bExcluded, lngType, lngCatID, bDisabledButtons, bActiveOnly, bSelList, _
			bNumeric, sSelList, oErr, false, false, 0, "", "")
end function
%>
