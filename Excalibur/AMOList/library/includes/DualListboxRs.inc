<%
' ========================================================================================
' Script for Dual Listbox
' ========================================================================================
' DLbx_GetHTML	: return the HTML string for the dual list box
' ========================================================================================

' ========================================================================================
' function DLbx_GetExcludeItems(byref rs, byval sIDFld)
' This function returns a filter string that excludes all items 
' contained in a recordset.
'
' rs (ADODB recordset): the recordset containing excluded items
' sIDFld (string): the ID field name
' ========================================================================================
function DLbx_GetExcludeItems(byref rs, byval sIDFld)
	' is the recordset defined?
	if (rs is nothing) then
		DLbx_GetExcludeItems = ""
		exit function
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		DLbx_GetExcludeItems = ""
		exit function
	end if

	dim	sFilter

	rs.MoveFirst

	while not rs.EOF
		if (sFilter <> "") then
			sFilter = sFilter & " AND "
		end if

		' add an item
		if isnumeric(rs.Fields(sIDFld).Value) then
			sFilter = sFilter & "(" & sIDFld & " <> " & rs.Fields(sIDFld).Value & ")"
		else
			sFilter = sFilter & "(" & sIDFld & " <> '" & rs.Fields(sIDFld).Value & "')"
		end if
		rs.MoveNext
	wend

	DLbx_GetExcludeItems = sFilter
end function
' ========================================================================================
' function DLbx_GetExcludeItems2(byval sList, byval sIDFld, byval bNumeric)
' This function returns a filter string that excludes all items 
' contained in a string delimited by comma.
'
' sList (string): the string containing excluded items delimited by comma
' sIDFld (string): the ID field name
' bNumeric (boolean): the flag to tell if the field is numeric or string
' ========================================================================================
function DLbx_GetExcludeItems2(byval sList, byval sIDFld, byval bNumeric)
	' is the list defined?
	if (trim(sList) = "") then
		DLbx_GetExcludeItems2 = ""
		exit function
	end if

	dim	sFilter, aVals, i
    aVals = split(sList, ",", -1, 1)

	for i = 0 to ubound(aVals)
		if (sFilter <> "") then
			sFilter = sFilter & " AND "
		end if

		' add an item
		if bNumeric then
			sFilter = sFilter & "(" & sIDFld & " <> " & aVals(i) & ")"
		else
			sFilter = sFilter & "(" & sIDFld & " <> '" & aVals(i) & "')"
		end if
	next

	DLbx_GetExcludeItems2 = sFilter
end function
' ========================================================================================
' function DLbx_DisableButtons
' This function returns the word disabled if the intdo value is True
'
' bDisabled: true or false to return DISABLED or not
' ========================================================================================
function DLbx_DisableButtons( byval bDisabled )
	if bDisabled then
		DLbx_DisableButtons = " DISABLED "
	else
		DLbx_DisableButtons = " "
	end if
end function

' ========================================================================================
' function BaseDLbx_GetHTML: returns the HTML string for the double list box.
'
' rsAvail (ADODB recordset): the recordset of available items
' sDescA (string):           the description field name of available items
' sIDA (string):             the ID field name of available items
' rsSel (ADODB recordset):   the recordset of selected items
' sDescS (string):           the description field name of selected items
' sIDS (string):             the ID field name of selected items
' sAHeading (string):        the heading for the list box of available items
' sSHeading (string):        the heading for the list box of available items
' sCtrlName (string):        the prefix for the list box names
' bExcluded (boolean):       the flag to exclude the selected items from the available box
' lngDefHeight (int):        Default Height of the list box
' lngDefWidth (int):         Default Width of the list box
' bDisableButtons:           the flag to disable the middle buttons or not
' bSelList (boolean):        the flag to exclude items in a list
' byval bNumeric:            the flag to tell if items in the exclude list is of numeric type
' byval sSelList:            the list of items to exclude delimited by comma
' byval bHorizonal:          True or False if a horizontal scroll bar should be added to the list boxes
' byval bCentered:           True or False if the dual list boxes should be centered. Only works with a horizontal scroll bar.
' byval intPadSpaces:        total length of description. This value minus description length will
'	                           determine how many spaces are padded at the end of the description. Only works with a 
'                            horizontal scroll bar.
' byval lngFontHeight:       height of the font used in list boxes
' byval sLeftHandler:        any javascript handler used for the available list box
' byval sRightHandler:       any javascript handler used for the selected list box
' byval lngFontHeight:       size of font height in pixels
' Please do not call this function from another page. This is the base function 
' for functions in this page to call only.
' ========================================================================================
function BaseDLbx_GetHTML(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, _
		byval sIDS, byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
		byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bSelList, byval bNumeric, _
		byval sSelList, byval bHorizontal, byval bCentered, byval intPadSpaces, byval lngFontHeight, _
		byval sLeftHandler, byval sRightHandler)
	dim sHTML
	dim sFilter
	dim lngDefButtonFontSize, lngSize
	dim sDblClickEvent

	'initialize values
	sFilter = ""
	lngDefButtonFontSize = 4
	sHTML = ""

	if bExcluded then
		' get the items excluded from the available box (because they are in the selected box already)
		if bSelList then
		    sFilter = DLbx_GetExcludeItems2(sSelList, sIDA, bNumeric)
		else
		    sFilter = DLbx_GetExcludeItems(rsSel, sIDS)
		end if

		' filter the recordset of available selections
		if (sFilter <> "") and (not rsAvail is nothing) then
			if cstr(rsAvail.Filter) <> "" and cstr(rsAvail.Filter) <> "0" then 
				rsAvail.Filter = rsAvail.Filter & " AND " & sFilter
			else
				rsAvail.Filter = sFilter
			end if				
		end if
	end if

	' write the HTML
	' table and heading for available and selected boxes
	if bCentered then
		sHTML = sHTML & "<div align=""center"">" & vbCrLf
	end if
	sHTML = sHTML & "<TABLE WIDTH=50% BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf
	sHTML = sHTML &  "<TR>" & vbCrLf
	sHTML = sHTML &  "<TH align=left>" & sAHeading & "</TH>" & vbCrLf 
	sHTML = sHTML &  "<TH></TH>" & vbCrLf 	
	sHTML = sHTML &  "<TH align=left>" & sSHeading & "</TH>" & vbCrLf 
	sHTML = sHTML &  "</TR>" & vbCrLf 	

	' fill options for the available box
	if bDisableButtons then
		sDblClickEvent = "" & sLeftHandler 'Disable the double click
	else
		sDblClickEvent = "onDblClick='return btnAddDLBCItem_onclick(btnAdd" & sCtrlName & ",lbxAvailable" & sCtrlName & ",lbxSelected" & sCtrlName & ")' " & sLeftHandler
	end if
	sHTML = sHTML &  "<TR>" & vbCrLf 
	sHTML = sHTML &  "<TD align=left> "
	if bHorizontal then
		if rsAvail is nothing then
			lngSize = 0
		else
			lngSize = rsAvail.recordcount
		end if
		if not rsSel is nothing then
			lngSize = lngSize + rsSel.recordcount
		end if
		sHTML = sHTML & ListboxRs_HorizontalGetHTML( bCentered, lngDefWidth, lngDefHeight, _
					"lbxAvailable" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
					rsAvail, sDescA, sIDA, nothing, "", 0, false, false, intPadSpaces, lngSize, _
					lngFontHeight)
	else
		sHTML = sHTML & Lbx_GetHTML6b("lbxAvailable" & sCtrlName, true, lngDefHeight, _
					lngDefWidth, rsAvail, sDescA, sIDA, nothing, "", _
					sDblClickEvent, bDisableButtons )
	end if
	sHTML = sHTML & "</TD>" & vbCrLf 	

	' button table
	sHTML = sHTML &  "<TD><TABLE border=0 cellPadding=1 cellSpacing=1 width='10%'>" & vbCrLf 

	sHTML = sHTML &  "<TR><TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	sHTML = sHTML &  "title='Move all Available entries to Selected list' " 
	sHTML = sHTML &  "name=" 
	sHTML = sHTML &  ("btnAddAll" & sCtrlName) 
	sHTML = sHTML &  " onclick='return btnAddAllDLBCItem_onclick(" 
	sHTML = sHTML &  ("btnAddAll" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	sHTML = sHTML &  ")' type=button value='>>'></TD>" & vbCrLf 
	sHTML = sHTML &  "	</TR>" & vbCrLf 	                    

	sHTML = sHTML &  "<TR><TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	sHTML = sHTML &  "title='Move selected Available entries to Selected list' " 
	sHTML = sHTML &  "name=" 
	sHTML = sHTML &  ("btnAdd" & sCtrlName) 
	sHTML = sHTML &  " onclick='return btnAddDLBCItem_onclick(" 
	sHTML = sHTML &  ("btnAdd" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	sHTML = sHTML &  ")' type=button value='->'></TD>" & vbCrLf 
	sHTML = sHTML &  "	</TR>" & vbCrLf 	                    

	sHTML = sHTML &  "	<TR><TD></TD></TR>" & vbCrLf 	                    

	sHTML = sHTML &  "	<TR>" & vbCrLf 	                    
	sHTML = sHTML &  "	<TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	sHTML = sHTML &  "title='Move selected Selected entries to Available list' " 
	sHTML = sHTML &  " name=" 
	sHTML = sHTML &  ("btnRemove" & sCtrlName) 
	sHTML = sHTML &  " onclick='return btnRemoveDLBCItem_onclick(" 
	sHTML = sHTML &  ("btnRemove" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	sHTML = sHTML &  ")' type=button value='<-'></TD>" & vbCrLf 
	sHTML = sHTML &  "	</TR>" & vbCrLf 	                    

	sHTML = sHTML &  "	<TR>" & vbCrLf 	                    
	sHTML = sHTML &  "	<TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	sHTML = sHTML &  "title='Move all Selected entries to Available list' " 
	sHTML = sHTML &  " name=" 
	sHTML = sHTML &  ("btnRemoveAll" & sCtrlName) 
	sHTML = sHTML &  " onclick='return btnRemoveAllDLBCItem_onclick(" 
	sHTML = sHTML &  ("btnRemoveAll" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	sHTML = sHTML &  ")' type=button value='<<'></TD>" & vbCrLf 
	sHTML = sHTML &  "	</TR>" & vbCrLf 	                    

	sHTML = sHTML &  "	</TABLE>" & vbCrLf
	if bCentered then
		sHTML = sHTML & "</div>" & vbCrLf
	end if
	sHTML = sHTML & "</TD>" & vbCrLf 

	' fill in the options for the selected box
	if bDisableButtons then
		sDblClickEvent = "" & sRightHandler 'Disable the double click
	else
		sDblClickEvent = "onDblClick='return btnRemoveDLBCItem_onclick(btnAdd" & sCtrlName & ",lbxAvailable" & sCtrlName & ",lbxSelected" & sCtrlName & ")' " & sRightHandler
	end if 
	sHTML = sHTML &  "<TD align=left> "
	if bExcluded and bSelList then
	    dim rs
		if (sFilter <> "") and (not rsAvail is nothing) then
		    sFilter = replace(sFilter, "<>", "=")
		    sFilter = replace(sFilter, "AND", "OR")
			rsAvail.Filter = sFilter
			set rs = rsAvail
	    else
	        set rs = nothing
		end if
		if bHorizontal then
			sHTML = sHTML & ListboxRs_HorizontalGetHTML( bCentered, lngDefWidth, lngDefHeight, _
						"lbxSelected" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
						rs, sDescA, sIDA, nothing, "", 0, false, false, intPadSpaces, lngSize, _
						lngFontHeight)
		else
		    sHTML = sHTML & Lbx_GetHTML6b("lbxSelected" & sCtrlName, true, lngDefHeight, _
		    			lngDefWidth, rs, sDescA, sIDA, nothing, "", _
		    			sDblClickEvent, bDisableButtons )
		end if
        set rs = nothing
	else
		if bHorizontal then
			sHTML = sHTML & ListboxRs_HorizontalGetHTML( bCentered, lngDefWidth, lngDefHeight, _
						"lbxSelected" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
						rsSel, sDescS, sIDS, nothing, "", 0, false, false, intPadSpaces, lngSize, _
						lngFontHeight)
		else
		    sHTML = sHTML & Lbx_GetHTML6b("lbxSelected" & sCtrlName, true, lngDefHeight, _
		    			lngDefWidth, rsSel, sDescS, sIDS, nothing, "", _
		    			sDblClickEvent, bDisableButtons )
		end if
    end if
	sHTML = sHTML & "</TD></TR>" & vbCrLf 	

	sHTML = sHTML &  "</TABLE>" & vbCrLf 

	BaseDLbx_GetHTML = sHTML
end function



'same as BaseDLbx_GetHTML but does writes instead of building HTML string
function BaseDLbx_GetHTML_Write(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, _
		byval sIDS, byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
		byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bSelList, byval bNumeric, _
		byval sSelList, byval bHorizontal, byval bCentered, byval intPadSpaces, byval lngFontHeight, _
		byval sLeftHandler, byval sRightHandler)
	dim sHTML
	dim sFilter
	dim lngDefButtonFontSize, lngSize
	dim sDblClickEvent

	'initialize values
	sFilter = ""
	lngDefButtonFontSize = 4
	sHTML = ""

	if bExcluded then
		' get the items excluded from the available box (because they are in the selected box already)
		if bSelList then
		    sFilter = DLbx_GetExcludeItems2(sSelList, sIDA, bNumeric)
		else
		    sFilter = DLbx_GetExcludeItems(rsSel, sIDS)
		end if

		' filter the recordset of available selections
		' JEC 8/26/10 - The below should really be here just like it is in the non-write version of the Base function.
		' However I don't know if it will cause problems elsewhere so we'll have to deal with passing in an Avail
		' recordset that's already filtered and by itself so the uncommented code below will work.
'		if (sFilter <> "") and (not rsAvail is nothing) then
'			if cstr(rsAvail.Filter) <> "" and cstr(rsAvail.Filter) <> "0" then 
'				rsAvail.Filter = rsAvail.Filter & " AND " & sFilter
'			else
'				rsAvail.Filter = sFilter
'			end if				
'		end if
		if (sFilter <> "") and (not rsAvail is nothing) then
			rsAvail.Filter = sFilter
		end if
	end if

	' write the HTML
	' table and heading for available and selected boxes
	if bCentered then
		response.write "<div align=""center"">" & vbCrLf
	end if
	response.write "<TABLE WIDTH=50% BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf
	response.write "<TR>" & vbCrLf
	response.write "<TH align=left>" & sAHeading & "</TH>" & vbCrLf 
	response.write "<TH></TH>" & vbCrLf 	
	response.write "<TH align=left>" & sSHeading & "</TH>" & vbCrLf 
	response.write "</TR>" & vbCrLf 	

	' fill options for the available box
	if bDisableButtons then
		sDblClickEvent = "" & sLeftHandler 'Disable the double click
	else
		sDblClickEvent = "onDblClick='return btnAddDLBCItem_onclick(btnAdd" & sCtrlName & ",lbxAvailable" & sCtrlName & ",lbxSelected" & sCtrlName & ")' " & sLeftHandler
	end if 
	response.write "<TR>" & vbCrLf 
	response.write "<TD align=left> "
	if bHorizontal then
		if rsAvail is nothing then
			lngSize = 0
		else
			if sFilter <> "" and rsSel is nothing then
				'set filter back so we can get the total size
				rsAvail.Filter = ""
				lngSize = rsAvail.recordcount
				rsAvail.Filter = sFilter
			else
				lngSize = rsAvail.recordcount
			end if
		end if
		if not rsSel is nothing then
			lngSize = lngSize + rsSel.recordcount
		end if
		Call ListboxRs_HorizontalGetHTML_Write( bCentered, lngDefWidth, lngDefHeight, _
					"lbxAvailable" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
					rsAvail, sDescA, sIDA, nothing, "", 0, false, false, intPadSpaces, lngSize, _
					lngFontHeight)
	else
		Call Lbx_GetHTML8b_Write("lbxAvailable" & sCtrlName, true, lngDefHeight, _
					lngDefWidth, rsAvail, sDescA, sIDA, nothing, "", _
					sDblClickEvent, bDisableButtons, lngFontHeight )
	end if
	response.write "</TD>" & vbCrLf 	

	' button table
	response.write "<TD><TABLE border=0 cellPadding=1 cellSpacing=1 width='10%'>" & vbCrLf 

	response.write "<TR><TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	response.write "title='Move all Available entries to Selected list' " 
	response.write "name=" 
	response.write ("btnAddAll" & sCtrlName) 
	response.write " onclick='return btnAddAllDLBCItem_onclick(" 
	response.write ("btnAddAll" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	response.write ")' type=button value='>>'></TD>" & vbCrLf 
	response.write "	</TR>" & vbCrLf 	                    

	response.write "<TR><TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	response.write "title='Move selected Available entries to Selected list' " 
	response.write "name=" 
	response.write ("btnAdd" & sCtrlName) 
	response.write " onclick='return btnAddDLBCItem_onclick(" 
	response.write ("btnAdd" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	response.write ")' type=button value='->'></TD>" & vbCrLf 
	response.write "	</TR>" & vbCrLf 	                    

	response.write "	<TR><TD></TD></TR>" & vbCrLf 	                    

	response.write "	<TR>" & vbCrLf 	                    
	response.write "	<TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	response.write "title='Move selected Selected entries to Available list' " 
	response.write " name=" 
	response.write ("btnRemove" & sCtrlName) 
	response.write " onclick='return btnRemoveDLBCItem_onclick(" 
	response.write ("btnRemove" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	response.write ")' type=button value='<-'></TD>" & vbCrLf 
	response.write "	</TR>" & vbCrLf 	                    

	response.write "	<TR>" & vbCrLf 	                    
	response.write "	<TD><INPUT " & DLbx_DisableButtons( bDisableButtons ) & "style = 'font-family:Courier New' language=javascript " 
	response.write "title='Move all Selected entries to Available list' " 
	response.write " name=" 
	response.write ("btnRemoveAll" & sCtrlName) 
	response.write " onclick='return btnRemoveAllDLBCItem_onclick(" 
	response.write ("btnRemoveAll" & sCtrlName) & "," & ("lbxAvailable" & sCtrlName) & "," & ("lbxSelected" & sCtrlName) 
	response.write ")' type=button value='<<'></TD>" & vbCrLf 
	response.write "	</TR>" & vbCrLf 	                    

	response.write "	</TABLE>" & vbCrLf
	if bCentered then
		response.write "</div>" & vbCrLf
	end if
	response.write "</TD>" & vbCrLf 

	' fill in the options for the selected box
	if bDisableButtons then
		sDblClickEvent = "" & sRightHandler 'Disable the double click
	else
		sDblClickEvent = "onDblClick='return btnRemoveDLBCItem_onclick(btnAdd" & sCtrlName & ",lbxAvailable" & sCtrlName & ",lbxSelected" & sCtrlName & ")' " & sRightHandler
	end if 
	response.write "<TD align=left> "
	if bExcluded and bSelList then
	    dim rs
		if (sFilter <> "") and (not rsAvail is nothing) then
		    sFilter = replace(sFilter, "<>", "=")
		    sFilter = replace(sFilter, "AND", "OR")
			rsAvail.Filter = sFilter
			set rs = rsAvail
	    else
	        set rs = nothing
		end if
		if bHorizontal then
			Call ListboxRs_HorizontalGetHTML_Write( bCentered, lngDefWidth, lngDefHeight, _
						"lbxSelected" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
						rs, sDescA, sIDA, nothing, "", 0, false, false, intPadSpaces, lngSize, _
						lngFontHeight)
		else
		    Call Lbx_GetHTML8_Write("lbxSelected" & sCtrlName, true, lngDefHeight, _
		    			lngDefWidth, rs, sDescA, sIDA, nothing, "", _
		    			sDblClickEvent, lngFontHeight )
		end if
        set rs = nothing
	else
		if bHorizontal then
			Call ListboxRs_HorizontalGetHTML_Write( bCentered, lngDefWidth, lngDefHeight, _
						"lbxSelected" & sCtrlName, true, bDisableButtons, sDblClickEvent, _
						rsSel, sDescS, sIDS, nothing, "", 0, false, false, intPadSpaces, lngSize, _
						lngFontHeight)
		else
		    Call Lbx_GetHTML8b_Write("lbxSelected" & sCtrlName, true, lngDefHeight, _
		    			lngDefWidth, rsSel, sDescS, sIDS, nothing, "", _
		    			sDblClickEvent, bDisableButtons, lngFontHeight )
		end if
    end if
	response.write "</TD></TR>" & vbCrLf 	

	response.write "</TABLE>" & vbCrLf 
end function

' ========================================================================================
' This is a legacy function (to provide backwards compatibility) and simply calls
'	the method of the same name below with default parameters for the height and
'	width of the listbox and to show buttons
' ========================================================================================
function DLbx_GetHTML(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, byval sAHeading, _
		 byval sSHeading, byval sCtrlName, byval bExcluded)
	DLbx_GetHTML = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, 7, 300, False, False, False, "", false, false, 0, 14, "", "")
end function

' ========================================================================================
' This is a legacy function (to provide backwards compatibility) and simply calls
'	the method of the same name below with default parameters for the height and
'	width of the listbox and to show buttons
' ========================================================================================
function DLbx_GetHTML2(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, byval sAHeading, _
		 byval sSHeading, byval sCtrlName, byval bExcluded, byval lngDefHeight, byval lngDefWidth)
	DLbx_GetHTML2 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, False, _
		 False, False, "", false, false, 0, 14, "", "")
end function

' ========================================================================================
' This function provides parameter to disable arrow buttons
' ========================================================================================
function DLbx_GetHTML3(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, byval sAHeading, _
		 byval sSHeading, byval sCtrlName, byval bExcluded, byval lngDefHeight, byval lngDefWidth, byval bDisableButtons)
	DLbx_GetHTML3 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, False, _
		 False, "", false, false, 0, 14, "", "")
end function

' ========================================================================================
' Exactly same as Dlbx_GetHTML3 
' ========================================================================================
function DLbx_GetHTML4(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, _
		byval sIDS, byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
		byval lngDefHeight, byval lngDefWidth, byval bDisableButtons)
	DLbx_GetHTML4 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, False, _
		 False, "", false, false, 0, 14, "", "")
end function

' ========================================================================================
' This function provides parameter to exclude items either by providing a recordset or  
' a list delimited by comma
' ========================================================================================
function DLbx_GetHTML5(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, _
          byval bSelList, byval bNumeric, byval sSelList, _
          byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
          byval lngDefHeight, byval lngDefWidth, byval bDisableButtons)
	DLbx_GetHTML5 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, bSelList, _
		 bNumeric, sSelList, false, false, 0, 14, "", "")
end function


' ========================================================================================
' This function allows a horizontal scroll bar and centered to be specified
' ========================================================================================
function DualListboxRs_GetHTML6(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, _
          byval bSelList, byval bNumeric, byval sSelList, _
          byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
          byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bHorizontal, _
		  byval bCentered, byval intPadSpaces, byval lngFontHeight )
	DualListboxRs_GetHTML6 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, bSelList, _
		 bNumeric, sSelList, bHorizontal, bCentered, intPadSpaces, lngFontHeight, "", "")
end function

function DualListboxRs_GetHTML7(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, _
          byval bSelList, byval bNumeric, byval sSelList, _
          byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
          byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bHorizontal, _
		  byval bCentered, byval intPadSpaces, byval lngFontHeight, byVal sLeftHandler, byVal sRightHandler )
	DualListboxRs_GetHTML7 = BaseDLbx_GetHTML(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, bSelList, _
		 bNumeric, sSelList, bHorizontal, bCentered, intPadSpaces, lngFontHeight, sLeftHandler, sRightHandler)
end function

'same as DualListboxRs_GetHTML6 but does writes instead of building the HTML string
function DualListboxRs_GetHTML6_Write(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, _
          byval bSelList, byval bNumeric, byval sSelList, _
          byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
          byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bHorizontal, _
				  byval bCentered, byval intPadSpaces, byval lngFontHeight )
	Call BaseDLbx_GetHTML_Write(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, bSelList, _
		 bNumeric, sSelList, bHorizontal, bCentered, intPadSpaces, lngFontHeight, "", "")
end function

function DualListboxRs_GetHTML7_Write(byref rsAvail, byval sDescA, byval sIDA, byval rsSel, byval sDescS, byval sIDS, _
          byval bSelList, byval bNumeric, byval sSelList, _
          byval sAHeading, byval sSHeading, byval sCtrlName, byval bExcluded, _
          byval lngDefHeight, byval lngDefWidth, byval bDisableButtons, byval bHorizontal, _
				  byval bCentered, byval intPadSpaces, byval lngFontHeight, byVal sLeftHandler, byVal sRightHandler )
	Call BaseDLbx_GetHTML_Write(rsAvail, sDescA, sIDA, rsSel, sDescS, sIDS, sAHeading, _
		 sSHeading, sCtrlName, bExcluded, lngDefHeight, lngDefWidth, bDisableButtons, bSelList, _
		 bNumeric, sSelList, bHorizontal, bCentered, intPadSpaces, lngFontHeight, sLeftHandler, sRightHandler)
end function
%>

<SCRIPT LANGUAGE=JAVASCRIPT>

// return comparison between two list options
function	strClientLessThan(a, b)
{
	var	aU = a.toString().toUpperCase();
	var	bU = b.toString().toUpperCase();
	return aU<bU?-1: a==b?0: a<b?-1: 1;
}


// transfers items from a source listbox to a destination listbox
function sortListbox(lbx, bAscending) 
{
	var oOption;
	var	i;

	// check if array empty
	if (lbx.options.length == 0)
		return;
/*		
    for (i=0; i < lbx.options.length; i++)
	{
		aTemp[i,0] = lbx.options[i].text;
		aTemp[i,1] = lbx.options[i].value;		
	}
    for (i=lbx.options.length-1; i >= 0 ; i--)
	{
		lbx.remove(i);
	}
	// sort the temporary storage

	// add the items back from the temporary storage back to listbox
	for (i=0; i < aTemp.length; i++)
	{	
		oOption = document.createElement("OPTION");
//		oOption.text=aTemp[i,0];
//		oOption.value=aTemp[i, 1];
		lbx.options.add(oOption);
	}
*/
}

// select all items in the list box
function SelectAll(lbx)
{
	var i;
	if (lbx.length > 0){
		for (i=0; i < lbx.options.length; i++){
			lbx.options[i].selected = true;
		}
	}
}

// transfers items from a source listbox to a destination listbox
function TransferSelectedItems(lbxSrc, lbxDest, bChangeAll) 
{
  // function to add and remove items to and from the list boxes
	var oOption;
	var i

	for (i=0; i < lbxSrc.options.length; i++)
		if (bChangeAll || lbxSrc.options[i].selected) {
			// add the selected software in the source list box to
			// the destination list box
			oOption = document.createElement("OPTION");
			oOption.text=lbxSrc.options[i].text;
			oOption.value=lbxSrc.options[i].value;
			lbxDest.options.add(oOption);
		}

	for (i=lbxSrc.options.length - 1; i >=0 ; i--)
		if (bChangeAll || lbxSrc.options[i].selected) {
			// remove the selected software in the source list box
			lbxSrc.remove(i);
		}

	// sort the listboxes
	sortListbox(lbxSrc, true);
	sortListbox(lbxDest, true);
}

// event handlers for DualListboxControl
function	btnAddAllDLBCItem_onclick(btnName, lbAvailable, lbSelected)
{
	TransferSelectedItems(lbAvailable, lbSelected, true);
	lbSelected.fireEvent("onchange");
	lbAvailable.fireEvent("onchange");
	bOnChange = true;
	return false;
}

function	btnAddDLBCItem_onclick(btnName, lbAvailable, lbSelected)
{
	TransferSelectedItems(lbAvailable, lbSelected, false);
	lbSelected.fireEvent("onchange");
	lbAvailable.fireEvent("onchange");
	bOnChange = true;
	return false;
}

function	btnRemoveDLBCItem_onclick(btnName, lbAvailable, lbSelected)
{
	TransferSelectedItems(lbSelected, lbAvailable, false);
	lbSelected.fireEvent("onchange");
	lbAvailable.fireEvent("onchange");
	bOnChange = true;
	return false;
}

function	btnRemoveAllDLBCItem_onclick(btnName, lbAvailable, lbSelected)
{
	TransferSelectedItems(lbSelected, lbAvailable, true);
	lbSelected.fireEvent("onchange");
	lbAvailable.fireEvent("onchange");
	bOnChange = true;
	return false;
}
</SCRIPT>