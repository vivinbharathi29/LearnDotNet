<% 
' ========================================================================================
' Script for Category Listbox
' ========================================================================================
' LbxCat_GetHTML		: writes the HTML for a listbox
' ========================================================================================

' ========================================================================================
' function LbxCat_BaseGetHTML(sName, bMultiple, lngHeight, lngWidth, 
'							  lngType, rsSel, sIDS, nIDS, bIDSel, lngCatID, oErr)
' This function returns the HTML string that build the list box 
' for categories.
'
' sName (string): the list box name
' bMultiple (boolean): the flag to enable multiple selection
' lngHeight (long): the height of the list box in term of number of entries
'					(for combo box, size = 1)
' lngWidth (long): the size of the box in pixel
' lngType (long): the category type
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
'			14 - Supported Language for Product Drop
'			15 - Supported OS for Product Drop
'			16 - Test dispersal method
'			20 - Supported Language for Delivery
'			21 - Supported Language for BOM
'			22 - Supported OS for Delivery
'			23 - Supported OS for BOM
'
' rsSel (ADODB recordset): the recordset to fill the selected items
' sIDS (string): the selected ID field name
' nIDS (number): the selected ID
' bIDSel (boolean): the flag to indicate whether to use the selected ID or 
'				  the selected recordset
' bAddBlank (boolean): the flag to add a blank entry
' lngCatID (long): the category ID
' Return
' oErr (object): the return error object
' ========================================================================================
function LbxCat_BaseGetHTML(byval sName, byval bMultiple, _
							byval lngHeight, byval lngWidth, byval lngType, _
							byref rsSel, byval sIDS, byval nIDS, byval bIDSel, _
							byval bAddBlank, byval lngCatID, byval sHandler, byref oErr)
	dim sDescA
	dim sIDA
	dim rsAvail

	'initialize values
	sDescA = "Description"

	'change the ID field name and the headings of the list boxes
	'depending on the category type
	
	select case lngType
	case 3	'Language
		sIDA = "LanguageID"
		sDescA = "LongDescription"
	case 4	'OS
		sIDA = "OSID"
	case 5	'Delivery Type
		sIDA = "DeliveryTypeID"
	case 6, 10	'HW, SW feature
		sIDA = "FeatureID"
	case 7	'
		sIDA = "HWValidationID"
	case 9  'test group
		sIDA = "TestGroupID"
	case 11	'
		sIDA = "VendorID"
	case 14, 20, 21	'Supported Language
		sIDA = "LanguageID"	
		sDescA = "LongDescription"
	case 15, 22, 23	'Supported OS
		sIDA = "OSID"
	case else	
		'HW category, SW category, test package category, test dispersal method
		sIDA = "CategoryID"
	end select

	'get the available categories
	'set oErr = GetCategory(rsAvail, lngType, lngCatID)
	set rsAvail = GetCategory(lngType, lngCatID)	

	'check if any error happens
	if rsAvail is nothing then
		'handle error
	    LbxCat_BaseGetHTML = "<H1>Error!<BR>"
	else
		'call the common list box script to get the HTML string
		if bIDSel then
			LbxCat_BaseGetHTML = Lbx_GetHTML4(sName, bMultiple, lngHeight, lngWidth, _
					rsAvail, sDescA, sIDA, nIDS, bAddBlank, sHandler)
		else
	  		LbxCat_BaseGetHTML = Lbx_GetHTML(sName, bMultiple, lngHeight, lngWidth, _
					rsAvail, sDescA, sIDA, rsSel, sIDS)
		end if
	end if

	'free objects
	set rsAvail = nothing
end function
' lists only active category items
function LbxCat_BaseGetHTMLActive(byval sName, byval bMultiple, _
							byval lngHeight, byval lngWidth, byval lngType, _
							byref rsSel, byval sIDS, byval nIDS, byval bIDSel, _
							byval bAddBlank, byval lngCatID, byval sHandler, byref oErr)
	dim sDescA
	dim sIDA
	dim rsAvail

	'initialize values
	sDescA = "Description"

	'change the ID field name and the headings of the list boxes
	'depending on the category type
	
	select case lngType
	case 3	'Language
		sIDA = "LanguageID"
		sDescA = "LongDescription"
	case 4	'OS
		sIDA = "OSID"
	case 5	'Delivery Type
		sIDA = "DeliveryTypeID"
	case 6, 10	'HW, SW feature
		sIDA = "FeatureID"
	case 7	'
		sIDA = "HWValidationID"
	case 9  'test group
		sIDA = "TestGroupID"
	case 11	'
		sIDA = "VendorID"
	case 14, 20, 21	'Supported Language
		sIDA = "LanguageID"	
		sDescA = "LongDescription"
	case 15, 22, 23	'Supported OS
		sIDA = "OSID"
	case else	
		'HW category, SW category, test package category, test dispersal method
		sIDA = "CategoryID"
	end select

	'get the available categories
	'set oErr = GetCategoryActive(rsAvail, lngType, lngCatID)

	'get the available categories
	Set rsAvail = Server.CreateObject("ADODB.Recordset")
	set rsAvail = GetCategoryActive(lngType, lngCatID)	

	'check if any error happens
	if rsAvail is Nothing then
		'handle error
	    LbxCat_BaseGetHTMLActive = "<H1>Error!<BR>"
	else
		'call the common list box script to get the HTML string
		if bIDSel then
			LbxCat_BaseGetHTMLActive = Lbx_GetHTML4(sName, bMultiple, lngHeight, lngWidth, _
					rsAvail, sDescA, sIDA, nIDS, bAddBlank, sHandler)
		else
	  		LbxCat_BaseGetHTMLActive = Lbx_GetHTML(sName, bMultiple, lngHeight, lngWidth, _
					rsAvail, sDescA, sIDA, rsSel, sIDS)
		end if
	end if

	'free objects
	set rsAvail = nothing
end function

' ========================================================================================
' function LbxCat_GetHTML(sName, rsSel, sIDS, bMultiple, _
'						lngHeight, lngWidth, lngType,  oErr)
' This function returns the HTML string that build the list box 
' for categories.
'
' sName (string): the list box name
' rsSel (ADODB recordset): the recordset to fill the selected items
' sIDS (string): the ID field name
' bMultiple (boolean): the flag to enable multiple selection
' lngHeight (long): the height of the list box in term of number of entries
'					(for combo box, size = 1)
' lngWidth (long): the size of the box in pixel
' lngType: types allowed by LbxCat_BaseGetHTML
' Return
' oErr (object): the return error object
' ========================================================================================
function LbxCat_GetHTML(byval sName, byref rsSel, byval sIDS, byval bMultiple, _
						byval lngHeight, byval lngWidth, byval lngType, byref oErr)
	LbxCat_GetHTML = LbxCat_BaseGetHTML(sName, bMultiple, _
							lngHeight, lngWidth, lngType, _
							rsSel, sIDS, 0, false, false, 0, "", oErr)
end function

' ========================================================================================
' function LbxCat_GetHTML_CID(sName, rsSel, sIDS, bMultiple, _
'						lngHeight, lngWidth, lngType, lngCatID, oErr)
' This function returns the HTML string that build the list box 
' for categories.
'
' sName (string): the list box name
' rsSel (ADODB recordset): the recordset to fill the selected items
' sIDS (string): the ID field name
' bMultiple (boolean): the flag to enable multiple selection
' lngHeight (long): the height of the list box in term of number of entries
'					(for combo box, size = 1)
' lngWidth (long): the size of the box in pixel
' lngType: types allowed by LbxCat_BaseGetHTML
' lngCatID (long): the categoryID type
' Return
' oErr (object): the return error object
' ========================================================================================
function LbxCat_GetHTML_CID(byval sName, byref rsSel, byval sIDS, byval bMultiple, _
						byval lngHeight, byval lngWidth, byval lngType, byval lngCatID, byref oErr)
	LbxCat_GetHTML_CID = LbxCat_BaseGetHTML(sName, bMultiple, _
							lngHeight, lngWidth, lngType, _
							rsSel, sIDS, 0, false, false, lngCatID, "", oErr)
end function


' ========================================================================================
' function LbxCat_GetDropdownHTML(sName, nIDS, lngWidth, lngType, oErr)
' This function returns the HTML string that build the list box 
' for categories.
'
' sName (string): the list box name
' lngWidth (long): the size of the box in pixel
' lngType (long): types allowed by LbxCat_BaseGetHTML
' Return
' oErr (object): the return error object
' ========================================================================================
function LbxCat_GetDropdownHTML(byval sName, byval nIDS, byval lngWidth, byval lngType, byref oErr)
	LbxCat_GetDropdownHTML = LbxCat_BaseGetHTML(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, false, 0, "", oErr)
end function
' lists only active category items
function LbxCat_GetDropdownHTMLActive(byval sName, byval nIDS, byval lngWidth, byval lngType, byref oErr)
	LbxCat_GetDropdownHTMLActive = LbxCat_BaseGetHTMLActive(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, false, 0, "", oErr)
end function

' ========================================================================================
' function LbxCat_GetDropdownBHTML(sName, nIDS, lngWidth, lngType, oErr)
' This function returns the HTML string that build the list box 
' for categories. The first item of the list box can be blank.
'
' sName (string): the list box name
' lngWidth (long): the size of the box in pixel
' lngType (long): types allowed by LbxCat_BaseGetHTML
' bAddBlank (boolean): the flag to add a blank entry
' Return
' oErr (object): the return error object
' ========================================================================================
function LbxCat_GetDropdownBHTML(byval sName, byval nIDS, byval lngWidth, byval lngType, byval bAddBlank, byref oErr)
	LbxCat_GetDropdownBHTML = LbxCat_BaseGetHTML(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, bAddBlank, 0, "", oErr)
end function

function LbxCat_GetDropdownHTML_CID(byval sName, byval nIDS, byval lngWidth, byval lngType, lngCatID, byref oErr)
	LbxCat_GetDropdownHTML = LbxCat_BaseGetHTML(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, false, "", "", oErr)
end function


function LbxCat_GetDropdownBHTML_CID(byval sName, byval nIDS, byval lngWidth, byval lngType, byval bAddBlank, byval lngCatID, byref oErr)
	LbxCat_GetDropdownBHTML_CID = LbxCat_BaseGetHTML(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, bAddBlank, lngCatID, "", oErr)
end function

function LbxCat_GetDropdownBHHTML_CID(byval sName, byval nIDS, byval lngWidth, byval lngType, byval bAddBlank, byval lngCatID, sHandler, byref oErr)
	LbxCat_GetDropdownBHHTML_CID = LbxCat_BaseGetHTML(sName, false, 1, lngWidth, _
					lngType, nothing, "", nIDS, true, bAddBlank, lngCatID, sHandler, oErr)
end function
%>
