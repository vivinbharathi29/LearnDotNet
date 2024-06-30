<% 
' ========================================================================================
' Script for Listbox
' ========================================================================================
' Lbx_GetHTML	: return the HTML string for the list box (selected items in a recordset)
' Lbx_GetHTML2	: return the HTML string for the list box (selected item in an ID)
' ListboxRs_HorizontalGetHTML: return the HTML string for a list box with a horizontal scroll bar
' ========================================================================================
' ========================================================================================
' function Lbx_AddOptions(byref rs, byval sDescFld, byval sIDFld)
' This function returns the HTML string containing options from 
' records in a recordset.
'
' rs (ADODB recordset): items to construct the HTML string of options
' sDescFld (string): the description field name
' sIDFld (string): the ID field name
' nIDS (number): the selected ID
' bIDSel (boolean): the flag to set the selected items by ID or by recordset
' intPadSpaces: total length of description. This value minus description length will
'		determine how many spaces are padded at the end of the description.
' ========================================================================================
function Lbx_AddOptions(byref rs, byval sDescFld, byval sIDFld, _
						byref rsSel, byval sIDFldS, byval nIDS, byval bIDSel, _
						byval bAddBlank, byval intPadSpaces)
	dim sFilter, strTemp, sOptions, sInputRsFilter, sInputRsSelFilter
	dim	i
	dim bRecSel, bStrIDFld
	bRecSel = false
	sInputRsFilter = ""
	sInputRsSelFilter = ""
	
	' is the recordset defined?
	if (rs is nothing) then
		Lbx_AddOptions = ""
		exit function
	end if

	sOptions = ""
	select case rs.Fields(sIDFld).Type
		case 129, 200
			bStrIDFld = true	'string field
		case else
			bStrIDFld = false
	end select

	'add a blank entry if the addblank flag is set
	if (bAddBlank) then
		sOptions = sOptions & "<OPTION value=" 
		if bStrIDFld then
			sOptions = sOptions & ""
		else
			sOptions = sOptions & "0"
		end if
		sOptions = sOptions & "> </OPTION>"' & vbCrLf
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		Lbx_AddOptions = sOptions
		exit function
	end if

	if not bIDSel then
		if not rsSel is nothing then
			sInputRsSelFilter = rsSel.Filter
			if rsSel.RecordCount > 0 then
				bRecSel = true
			end if
		end if
	end if

	'preserve the filter of the input recordset
	sInputRsFilter = rs.Filter
	rs.MoveFirst

	while not rs.EOF
		' add an option
		sOptions = sOptions & "<OPTION value=" 
		if bStrIDFld then 
			sOptions = sOptions & """"
		end if
		sOptions = sOptions & rs.Fields(sIDFld).Value 
		if bStrIDFld then
			sOptions = sOptions & """"
		end if

		'set the selected entries
		if bIDSel then
			if bStrIDFld then
				if lcase(rs.Fields(sIDFld).Value) = lcase(nIDS) then
					sOptions = sOptions & " selected "
				end if
			else
				if rs.Fields(sIDFld).Value = nIDS then
					sOptions = sOptions & " selected " 
				end if
			end if
		else
		' add the selected flag for entries in the selected recordset
			if bRecSel then
				'4/15/02 - fixed code so it actually works
				if sInputRsSelFilter <> "" then
					sFilter = sInputRsSelFilter & " and "
				else
					sFilter = ""
				end if
				sFilter = sFilter & sIDFldS & " = " & rs.Fields(sIDFld).Value
				rsSel.Filter = sFilter
				if rsSel.RecordCount > 0 then
					sOptions = sOptions & " selected " 
				end if
				rsSel.Filter = sInputRsSelFilter
			end if
		end if
		sOptions = sOptions & ">"
		strTemp = TRIM(rs.Fields(sDescFld).Value)
		if intPadSpaces > 0 then
			strTemp = strTemp & SPACE(intPadSpaces - LEN(strTemp))
		end if
		sOptions = sOptions & replace(strTemp, " ", "&nbsp;") & "</OPTION>"' & vbCrLf
		'sOptions = sOptions & replace(trim(rs.Fields(sDescFld).Value), " ", "&nbsp;") & "</OPTION>"' & vbCrLf
		rs.MoveNext
	wend

	'reset the filter of the options recordset
	if bRecSel then
		rs.Filter = sInputRsFilter
	end if

	'return the HTML string with the options from the recordset
	Lbx_AddOptions = sOptions
end function

'same as Lbx_AddOptions but default to show all options
function Lbx_AddOptions_All(byref rs, byval sDescFld, byval sIDFld, _
						byref rsSel, byval sIDFldS, byval nIDS, byval bIDSel, _
						byval bAddBlank, byval intPadSpaces)
	dim sFilter, strTemp, sOptions, sInputRsFilter, sInputRsSelFilter
	dim	i
	dim bRecSel, bStrIDFld
	dim nDefaultValue
	bRecSel = false
	sInputRsFilter = ""
	sInputRsSelFilter = ""
	
	' is the recordset defined?
	if (rs is nothing) then
		Lbx_AddOptions_All = ""
		exit function
	end if

	sOptions = ""
	select case rs.Fields(sIDFld).Type
		case 129, 200
			bStrIDFld = true	'string field
		case else
			bStrIDFld = false
	end select

	'add a blank entry if the addblank flag is set
	
	if (bAddBlank) then
		sOptions = sOptions & "<OPTION value=" 
		if bStrIDFld then
			sOptions = sOptions & ""
		else
			sOptions = sOptions & "0"
		end if
		sOptions = sOptions & ">All</OPTION>"' & vbCrLf
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		Lbx_AddOptions_All = sOptions
		exit function
	end if

	if not bIDSel then
		if not rsSel is nothing then
			sInputRsSelFilter = rsSel.Filter
			if rsSel.RecordCount > 0 then
				bRecSel = true
			end if
		end if
	end if

	'preserve the filter of the input recordset
	sInputRsFilter = rs.Filter
	rs.MoveFirst

	while not rs.EOF
		' add an option
		sOptions = sOptions & "<OPTION value=" 
		if bStrIDFld then 
			sOptions = sOptions & """"
		end if
		sOptions = sOptions & rs.Fields(sIDFld).Value 
		if bStrIDFld then
			sOptions = sOptions & """"
		end if

		'set the selected entries
		if bIDSel then
			if bStrIDFld then
				if lcase(rs.Fields(sIDFld).Value) = lcase(nIDS) then
					sOptions = sOptions & " selected "
					'sOptions = sOptions & ""
				end if
			else
				if rs.Fields(sIDFld).Value = nIDS then
					sOptions = sOptions & " selected "
					'sOptions = sOptions & ""
				end if
			end if
		else
		' add the selected flag for entries in the selected recordset
			if bRecSel then
				'4/15/02 - fixed code so it actually works
				if sInputRsSelFilter <> "" then
					sFilter = sInputRsSelFilter & " and "
				else
					sFilter = ""
				end if
				sFilter = sFilter & sIDFldS & " = " & rs.Fields(sIDFld).Value
				rsSel.Filter = sFilter
				if rsSel.RecordCount > 0 then
					sOptions = sOptions & " selected " 
					'sOptions = sOptions & ""
				end if
				rsSel.Filter = sInputRsSelFilter
			end if
		end if
		sOptions = sOptions & ">"
		strTemp = TRIM(rs.Fields(sDescFld).Value)
		if intPadSpaces > 0 then
			strTemp = strTemp & SPACE(intPadSpaces - LEN(strTemp))
		end if
		sOptions = sOptions & replace(strTemp, " ", "&nbsp;") & "</OPTION>"' & vbCrLf
		sOptions = sOptions & replace(trim(rs.Fields(sDescFld).Value), " ", "&nbsp;") & "</OPTION>"' & vbCrLf
		rs.MoveNext
	wend

	'reset the filter of the options recordset
	if bRecSel then
		rs.Filter = sInputRsFilter
	end if

	'return the HTML string with the options from the recordset
	Lbx_AddOptions_All = sOptions
end function


'same as Lbx_AddOptions but does writes instead of building an HTML string
function Lbx_AddOptions_Write(byref rs, byval sDescFld, byval sIDFld, _
						byref rsSel, byval sIDFldS, byval nIDS, byval bIDSel, _
						byval bAddBlank, byval intPadSpaces)
	dim sFilter, strTemp, sOptions
	dim	i
	dim bRecSel, bStrIDFld
	bRecSel = false
	
	' is the recordset defined?
	if (rs is nothing) then
		exit function
	end if

	select case rs.Fields(sIDFld).Type
		case 129, 200
			bStrIDFld = true	'string field
		case else
			bStrIDFld = false
	end select

	'add a blank entry if the addblank flag is set
	if (bAddBlank) then
		response.write "<OPTION value=" 
		if bStrIDFld then
			response.write ""
		else
			response.write "0"
		end if
		response.write "> </OPTION>" & vbCrLf
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		exit function
	end if

	if not bIDSel then
		if not rsSel is nothing then
			if rsSel.RecordCount > 0 then
				bRecSel = true
			end if
		end if
	end if

	rs.MoveFirst

	while not rs.EOF
		' add an option
		response.write "<OPTION value=" 
		if bStrIDFld then 
			response.write """"
		end if
		response.write rs.Fields(sIDFld).Value 
		if bStrIDFld then
			response.write """"
		end if

		'set the selected entries
		if bIDSel then
			if bStrIDFld then
				if lcase(rs.Fields(sIDFld).Value) = lcase(nIDS) then
					response.write " selected "
				end if
			else
				if rs.Fields(sIDFld).Value = nIDS then
					response.write " selected " 
				end if
			end if
		else
		' add the selected flag for entries in the selected recordset
			if bRecSel then
				'4/15/02 - fixed code so it actually works
				sFilter = sIDFldS & " = " & rs.Fields(sIDFld).Value
				rsSel.Filter = sFilter
				if rsSel.RecordCount > 0 then
					response.write " selected " 
				end if
				rsSel.Filter = ""
			end if
		end if
		response.write ">"
		strTemp = TRIM(rs.Fields(sDescFld).Value)
		if intPadSpaces > 0 then
			strTemp = strTemp & SPACE(intPadSpaces - LEN(strTemp))
		end if
		response.write replace(strTemp, " ", "&nbsp;") & "</OPTION>" & vbCrLf
		'sOptions = sOptions & replace(trim(rs.Fields(sDescFld).Value), " ", "&nbsp;") & "</OPTION>" & vbCrLf
		rs.MoveNext
	wend

	'reset the filter of the selected recordset
	if bRecSel then
		rs.Filter = ""
	end if
end function

'same as Lbx_AddOptions_All but defaults to show all options and does writes instead of building an HTML string
function Lbx_AddOptions_All_Write(byref rs, byval sDescFld, byval sIDFld, _
						byref rsSel, byval sIDFldS, byval nIDS, byval bIDSel, _
						byval bAddBlank, byval intPadSpaces)
	dim sFilter, strTemp, sOptions
	dim	i
	dim bRecSel, bStrIDFld
	bRecSel = false
	
	' is the recordset defined?
	if (rs is nothing) then
		exit function
	end if

	select case rs.Fields(sIDFld).Type
		case 129, 200
			bStrIDFld = true	'string field
		case else
			bStrIDFld = false
	end select

	'add a blank entry if the addblank flag is set
	if (bAddBlank) then
		response.write "<OPTION value=" 
		if bStrIDFld then
			response.write ""
		else
			response.write "0"
		end if
		response.write ">All</OPTION>" & vbCrLf
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		exit function
	end if

	if not bIDSel then
		if not rsSel is nothing then
			if rsSel.RecordCount > 0 then
				bRecSel = true
			end if
		end if
	end if

	rs.MoveFirst

	while not rs.EOF
		' add an option
		response.write "<OPTION value=" 
		if bStrIDFld then 
			response.write """"
		end if
		response.write rs.Fields(sIDFld).Value 
		if bStrIDFld then
			response.write """"
		end if

		'set the selected entries
		if bIDSel then
			if bStrIDFld then
				if lcase(rs.Fields(sIDFld).Value) = lcase(nIDS) then
					response.write " selected "
				end if
			else
				if rs.Fields(sIDFld).Value = nIDS then
					response.write " selected " 
				end if
			end if
		else
		' add the selected flag for entries in the selected recordset
			if bRecSel then
				'4/15/02 - fixed code so it actually works
				sFilter = sIDFldS & " = " & rs.Fields(sIDFld).Value
				rsSel.Filter = sFilter
				if rsSel.RecordCount > 0 then
					response.write " selected " 
				end if
				rsSel.Filter = ""
			end if
		end if
		response.write ">"
		strTemp = TRIM(rs.Fields(sDescFld).Value)
		if intPadSpaces > 0 then
			strTemp = strTemp & SPACE(intPadSpaces - LEN(strTemp))
		end if
		response.write replace(strTemp, " ", "&nbsp;") & "</OPTION>" & vbCrLf
		'sOptions = sOptions & replace(trim(rs.Fields(sDescFld).Value), " ", "&nbsp;") & "</OPTION>" & vbCrLf
		rs.MoveNext
	wend

	'reset the filter of the selected recordset
	if bRecSel then
		rs.Filter = ""
	end if
end function

sub Lbx_AddOptionswrite(byref rs, byval sDescFld, byval sIDFld, _
						byref rsSel, byval sIDFldS, byval nIDS, byval bIDSel, _
						byval bAddBlank)
	dim sFilter
	dim	sOptions, i, bRecSel, bStrIDFld
	bRecSel = false
	
	' is the recordset defined?
	if (rs is nothing) then
		'Lbx_AddOptions = ""
		exit sub
	end if

	sOptions = ""
	select case rs.Fields(sIDFld).Type
		case 129, 200
			bStrIDFld = true	'string field
		case else
			bStrIDFld = false
	end select

	'add a blank entry if the addblank flag is set
	if (bAddBlank) then
		response.write  "<OPTION value=" 
		if bStrIDFld then
			response.write  ""
		else
			response.write  "0"
		end if
		response.write  "> </OPTION>" & vbCrLf
	end if

	' is there any record?
	if (rs.RecordCount = 0) then
		'Lbx_AddOptionswrite = sOptions
		exit sub
	end if

	if not bIDSel then
		if not rsSel is nothing then
			if rsSel.RecordCount > 0 then
				bRecSel = true
			end if
		end if
	end if

	rs.MoveFirst

	while not rs.EOF
		' add an option
		response.write  "<OPTION value=" 
		if bStrIDFld then 
			response.write  """"
		end if
		response.write   rs.Fields(sIDFld).Value 
		if bStrIDFld then
			response.write   """"
		end if

		'set the selected entries
		if bIDSel then
			if bStrIDFld then
				if lcase(rs.Fields(sIDFld).Value) = lcase(nIDS) then
					response.write   " selected "
				end if
			else
				if rs.Fields(sIDFld).Value = nIDS then
					response.write   " selected " 
				end if
			end if
		else
		' add the selected flag for entries in the selected recordset
			if bRecSel then
				sFilter = sIDFld & " = " & rs.Fields(sIDFld).Value
				rs.Filter = sFilter
				if rs.RecordCount > 0 then
					response.write   " selected " 
				end if
			end if
		end if
		response.write   ">"
		response.write   trim(rs.Fields(sDescFld).Value) & "</OPTION>" & vbCrLf
		rs.MoveNext
	wend

	'reset the filter of the selected recordset
	if bRecSel then
		rs.Filter = ""
	end if

	'return the HTML string with the options from the recordset

end sub


'duplication of Lbx_AddOptionswrite, but uses CSV list of ID's for selected values 
sub ListBox_AddOptionsWrite(byref resultSetOptions, byval stringDescriptionField, byval stringIDField, byval string_CSV_SelectedIDs, byval booleanAddBlank)
	
	dim	stringSelectedIDsArray
	dim filteredArray
	dim i
	
	stringSelectedIDsArray = Split(string_CSV_SelectedIDs, ",")	

	' is the recordset defined?
	if (resultSetOptions is nothing) then
		exit sub
	else
		if (resultSetOptions.RecordCount = 0) then
			exit sub
		end if
	end if

	'add a blank entry if the addblank flag is set
	if (booleanAddBlank) then
		response.write  "<OPTION value=""0""> </OPTION>" & vbCrLf
	end if

	resultSetOptions.MoveFirst

	while not resultSetOptions.EOF

		' add an option
		response.write  "<OPTION value=""" & resultSetOptions.Fields(stringIDField).Value & """"

		'reduce selected values to those similar to the current value
		filteredArray=Filter(stringSelectedIDsArray, resultSetOptions.Fields(stringIDField).Value)
		'if there are still any left search through them
		For i = 0 To Ubound(filteredArray, 1)
			if trim(filteredArray(i)) = trim(resultSetOptions.Fields(stringIDField).Value) then
				'exact match has been found, set selected flag and exit loop
				response.write " selected"
				'exit for
			end if
		next

		response.write   " >" & trim(resultSetOptions.Fields(stringDescriptionField).Value) & "</OPTION>" & vbCrLf
		resultSetOptions.MoveNext
	wend

end sub


function Lbx_BaseGetHTML(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS, _
					byval nIDS, byval bIDSel, byval bAddBlank, byval sHandler, byval bDisabled)
	dim sHTML
	dim sName

	sHTML = "<SELECT name=" 
	sHTML = sHTML & "'" & (sCtrlName) & "'" 
	if bMultiple then
		sHTML = sHTML & " multiple "
	end if
	if bDisabled then
		sHTML = sHTML & " disabled "
	end if
	
	sHTML = sHTML &  " id='" & sCtrlName & "' size=" & lngSize 

	if lngWidth > 0 then
		if bDisabled then
			'gray out the text
			sHTML = sHTML & " style='WIDTH: " & lngWidth & "px;color:gray' "
		else
			sHTML = sHTML & " style='WIDTH: " & lngWidth & "px' "
		end if
	else
		if bDisabled then
			sHTML = sHTML & " style='color:gray' "
		end if
	end if

	if sHandler <> "" then
		sHTML = sHTML & sHandler
	end if
	sHTML = sHTML & " >" '& vbCrLf
	sHTML = sHTML &  Lbx_AddOptions(rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, nIDS, bIDSel, bAddBlank, 0)
	sHTML = sHTML &  "</SELECT>" '& vbCrLf
	Lbx_BaseGetHTML = sHTML
end function 

'to show all options
function Lbx_BaseGetHTML_All(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS, _
					byval nIDS, byval bIDSel, byval bAddBlank, byval sHandler, byval bDisabled)
	dim sHTML
	dim sName

	sHTML = "<SELECT name=" 
	sHTML = sHTML & "'" & (sCtrlName) & "'" 
	if bMultiple then
		sHTML = sHTML & " multiple "
	end if
	if bDisabled then
		sHTML = sHTML & " disabled "
	end if
	
	if lngWidth > 0 then
		sHTML = sHTML &  " id='" & sCtrlName & "' size=" & lngSize &" style='WIDTH: " & lngWidth & "px' "
	else
		sHTML = sHTML &  " id='" & sCtrlName & "' size=" & lngSize & " "
	end if
	
	if sHandler <> "" then
		sHTML = sHTML & sHandler
	end if
	sHTML = sHTML & " >" '& vbCrLf
	sHTML = sHTML &  Lbx_AddOptions_All(rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, nIDS, bIDSel, bAddBlank, 0)
	sHTML = sHTML &  "</SELECT>" '& vbCrLf
	Lbx_BaseGetHTML_All = sHTML
end function 

sub Lbx_BaseGetHTMLwrite(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS, _
					byval nIDS, byval bIDSel, byval bAddBlank, byval sHandler, byval bDisabled, byval lngFontHeight)
	dim sHTML
	dim sName

	response.write  "<SELECT name=" 
	response.write   "'" & (sCtrlName) & "'" 
	if bMultiple then
		response.write   " multiple "
	end if
	if bDisabled then
		response.write  " disabled "
	end if

	if lngWidth > 0 then
		response.write " id='" & sCtrlName & "' size=" & lngSize &" style='WIDTH: " & lngWidth & "px' "
	else
		response.write " id='" & sCtrlName & "' size=" & lngSize & " "
	end if

	if lngFontHeight <> "" then
		response.write " style=""font-size:" & cstr(lngFontHeight) & "px """
	end if

	if sHandler <> "" then
		response.write  sHandler
	end if

	response.write   " >" & vbCrLf
	call  Lbx_AddOptionswrite(rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, nIDS, bIDSel, bAddBlank) 
	response.write    "</SELECT>" & vbCrLf
end sub

'to show all
sub Lbx_BaseGetHTML_All_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS, _
					byval nIDS, byval bIDSel, byval bAddBlank, byval sHandler, byval bDisabled, byval lngFontHeight)
	dim sHTML
	dim sName

	response.write  "<SELECT name=" 
	response.write   "'" & (sCtrlName) & "'" 
	if bMultiple then
		response.write   " multiple "
	end if
	if bDisabled then
		response.write  " disabled "
	end if

	if lngWidth > 0 then
		response.write " id='" & sCtrlName & "' size=" & lngSize &" style='WIDTH: " & lngWidth & "px' "
	else
		response.write " id='" & sCtrlName & "' size=" & lngSize & " "
	end if

	if lngFontHeight <> "" then
		response.write " style=""font-size:" & cstr(lngFontHeight) & "px """
	end if

	if sHandler <> "" then
		response.write  sHandler
	end if

	response.write   " >" & vbCrLf
	call  Lbx_AddOptions_All_Write(rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, nIDS, bIDSel, bAddBlank, 0) 
	response.write    "</SELECT>" & vbCrLf
end sub

'duplication of Lbx_BaseGetHTMLWrite but calls new option subroutine 
sub ListBox_BaseGetHTMLWrite(byval stringControlName, byval booleanAllowMultipleSelectionsFlag, byval intSize, byval intWidth, _ 
					byref resultSetOptions, byval stringDescriptionField, byval stringIDField, byval string_CSV_SelectedIDs, byval booleanDisabledFlag, byval booleanAddBlankFlag)

	response.write  "<SELECT name=""" & (stringControlName) & """" 
	if booleanAllowMultipleSelectionsFlag then
		response.write   " multiple"
	end if
	if booleanDisabledFlag then
		response.write  " disabled"
	end if

	if intWidth > 0 then
		response.write " id=""" & stringControlName & """ size=""" & intSize &""" style=""WIDTH:" & intWidth & "px;"""
	else
		response.write " id=""" & stringControlName & """ size=""" & intSize & """"
	end if

	response.write   " >" & vbCrLf
	call  ListBox_AddOptionsWrite(resultSetOptions, stringDescriptionField, stringIDField, string_CSV_SelectedIDs, booleanAddBlankFlag) 
	response.write    "</SELECT>" & vbCrLf
end sub

'the default selections are in a recordset
function Lbx_GetHTML(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS)
	Lbx_GetHTML = Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, "", false)
end function

sub Lbx_GetHTMLwrite(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, byval sIDFldS)
	call Lbx_BaseGetHTMLwrite (sCtrlName, bMultiple, lngSize, lngWidth, rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, "", false, "")
end sub


'the default selection is an ID and no blank entry
function Lbx_GetHTML2(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID)
	Lbx_GetHTML2 = 	Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, false, "", false)
end function

'the default selection is an ID and can have a blank entry
function Lbx_GetHTML3(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank)
	Lbx_GetHTML3 = 	Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, "", false)
end function

'the default selection is an ID, can have a blank entry and a handler
function Lbx_GetHTML4(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, sHandler)
	Lbx_GetHTML4 = 	Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, false)
end function
sub Lbx_GetHTML4write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, sHandler)
	call Lbx_BaseGetHTMLwrite (sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, false, "")
end sub

'the default selection is an ID, can have a blank entry, a handler and can be disabled
function Lbx_GetHTML5(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, byval sHandler, byval bDisabled)
	Lbx_GetHTML5 = 	Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, bDisabled)
end function
sub Lbx_GetHTML5write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, byval sHandler, byval bDisabled)
	call Lbx_BaseGetHTMLwrite (sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, bDisabled, "")
end sub

'the default selections are in a recordset
function Lbx_GetHTML6(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, _
					byval sIDFldS, byval sHandler)
	'same as Lbx_GetHMTL except allows a handler to be specified
	Lbx_GetHTML6 = Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, sHandler, false)
end function
'same as Lbx_GetHTML6 but does writes instead of building an HTML string
function Lbx_GetHTML6_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, _
					byval sIDFldS, byval sHandler)
	'same as Lbx_GetHMTL except allows a handler to be specified
	Call Lbx_BaseGetHTMLWrite(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, sHandler, false, "")
end function
'same as Lbx_GetHTML6 but allows disabled
function Lbx_GetHTML6b(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, _
					byval sIDFldS, byval sHandler, byval bDisabled)
	Lbx_GetHTML6b = Lbx_BaseGetHTML(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, sHandler, bDisabled)
end function

'the default selection is to show all, can have a blank entry, a handler and can be disabled
function Lbx_GetHTML7(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, byval sHandler, byval bDisabled)
	Lbx_GetHTML7 = 	Lbx_BaseGetHTML_All(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, bDisabled)
end function
function Lbx_GetHTML7_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, byval sHandler, byval bDisabled)
	Call Lbx_BaseGetHTMLWrite(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, bDisabled, "")
end function
'does the default to show all just like Lbx_GetHTML7
function Lbx_GetHTML7_All_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byval nSelID, byval bAddBlank, byval sHandler, byval bDisabled)
	Call Lbx_BaseGetHTML_All_Write(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, nothing, "", nSelID, true, bAddBlank, sHandler, bDisabled, "")
end function



'same as Lbx_GetHTML6_Write but can specify the font size
function Lbx_GetHTML8_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, _
					byval sIDFldS, byval sHandler, byval lngFontHeight)
	'same as Lbx_GetHMTL except allows a handler to be specified
	Call Lbx_BaseGetHTMLWrite(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, sHandler, false, lngFontHeight)
end function
'add in bDisabled
function Lbx_GetHTML8b_Write(byval sCtrlName, byval bMultiple, byval lngSize, byval lngWidth, _
					byref rsOptions, byval sDescFld, byval sIDFld, byref rsSel, _
					byval sIDFldS, byval sHandler, byval bDisabled, byval lngFontHeight)
	'same as Lbx_GetHMTL except allows a handler to be specified
	Call Lbx_BaseGetHTMLWrite(sCtrlName, bMultiple, lngSize, lngWidth, _
		rsOptions, sDescFld, sIDFld, rsSel, sIDFldS, 0, false, false, sHandler, bDisabled, lngFontHeight)
end function

'list box that allows for multiple items to be selected based on a csv list
function ListBox_GetHTML_Write(byval stringControlName, byval booleanAllowMultipleSelectionsFlag, byval intSize, byval intWidth, _
					byref resultSetOptions, byval stringDescriptionField, byval stringIDField, byref string_CSV_SelectedIDs, byval booleanDisabledFlag, byval booleanAddBlankFlag)
	Call ListBox_BaseGetHTMLWrite(stringControlName, booleanAllowMultipleSelectionsFlag, intSize, intWidth, _
					resultSetOptions, stringDescriptionField, stringIDField, string_CSV_SelectedIDs, booleanDisabledFlag, booleanAddBlankFlag)
end function

' ========================================================================================
' function Lbx_BaseHorizontalGetHTML
' This function returns the HTML string containing options from 
' records in a recordset. It builds a list box with a horizontal scroll bar.
'
' bCentered:     True or False if list box should be centered
' lngWidth:      width of list box in pixels
' lngHeight:     height of list box in pixels
' strCtrlName:   name of listbox
' bMultiple:     True or False if multiple items can be selected
' bDisabled:     True or False if list box should be disabled
' strHandler:    any string for Javascript functions
' rsOptions:     recordset containing items to include in list box
' strDescFld:    field in rsOptions that is the description field to display to user
' strIDFld:      field in rsOptions that is the ID field to send to next page
' rsSel:         recordset containing currently selected items in list box
' strIDFldS:     field in rsSel that is the ID field that is currently selected
' lngIDS:        the selected ID in rsSel if appropriate. 0 if more than one can be selected
' bIDSel:        True if selected items by ID or False irf selected items by recordset
' bAddBlank:     True or False if a blank entry should be added at top of list box
' intPadSpaces:  total length of description. This value minus description length will
'				  determine how many spaces are padded at the end of the description.
' lngSize:       Total number of items in both dual list boxes to give a total number of lines in a list box
' lngFontHeight: Font height in pixels
' ========================================================================================
function Lbx_BaseHorizontalGetHTML(byval bCentered, byval lngWidth, byval lngHeight, _
		byval strCtrlName, byval bMultiple, byval bDisabled, byval strHandler, _
		byref rsOptions, byval strDescFld, byval strIDFld, byref rsSel, byval strIDFldS, _
		byval lngIDS, byval bIDSel, byval bAddBlank, byval intPadSpaces, byval lngSize, _
		byval lngFontHeight )
	dim sHTML, sName
	dim lngTotalRecords
	
	if rsOptions is nothing then
		lngTotalRecords = 0
	else
		lngTotalRecords = rsOptions.RecordCount
	end if
	if bAddBlank then
		'add one since the top line will be blank
		lngSize = lngSize + 1
	end if

	sHTML = ""
	if bCentered then
		sHTML = sHTML & "<div align=""center"">" & vbCrLf
	end if
	if lngWidth > 0 then
		sHTML = sHTML & "<table border=1 cellpadding=0 cellspacing=0 width=""" & cstr(lngWidth) & "px""" & vbCrLf
	else
		sHTML = sHTML & "<table border=1 cellpadding=0 cellspacing=0 " & vbCrLf
	end if
	sHTML = sHTML & "	style=""border-bottom: none; border-left: none; border-right: none; border-top: thin;"">" & vbCrLf
	sHTML = sHTML & "<tr><td>" & vbCrLf
	sHTML = sHTML & "<DIV style=""OVERFLOW: auto; "
	if lngWidth > 0 then
		sHTML = sHTML & "WIDTH: " & cstr(lngWidth) & "px; "
	end if
	if lngHeight > 0 then
		sHTML = sHTML & "HEIGHT: " & cstr(lngHeight) & "px"
	end if
	sHTML = sHTML & """>" & vbCrLf
	if lngSize < (cint(lngHeight / lngFontHeight)) then
		'since there are less lines than what is needed for the list box,
		'build up list box so the internal SELECT tag fills the list box.
		lngSize = cint(lngHeight / lngFontHeight)
	end if
	sHTML = sHTML & "	<select name=""" & strCtrlName & """ id=""" & strCtrlName & """ size=""" & cstr(lngSize) & """ "
	sHTML = sHTML & "style=""font-size:" & cstr(lngFontHeight) & "px"
	'if lngTotalRecords = 0 then
		'no records so make the list box wider so it looks okay. Just picked 10 as a large number.
		sHTML = sHTML & ";width:" & cstr(lngWidth*10) & "px"
	'end if
	if bDisabled then	'gray out the text
		sHTML = sHTML & ";color:gray"
	end if
	sHTML = sHTML & """ "	'close style
	if bMultiple then
		sHTML = sHTML & "multiple "
	end if
	if bDisabled then
		sHTML = sHTML & " disabled "
	end if
	if strHandler <> "" then
		sHTML = sHTML & strHandler
	end if
	sHTML = sHTML & " >" & vbCrLf

	sHTML = sHTML &  Lbx_AddOptions(rsOptions, strDescFld, strIDFld, rsSel, strIDFldS, lngIDS, bIDSel, bAddBlank, intPadSpaces)

	sHTML = sHTML &  "</SELECT>" & vbCrLf
	sHTML = sHTML &  "</DIV>" & vbCrLf
	sHTML = sHTML &  "</TD></TR>" & vbCrLf
	sHTML = sHTML &  "</TABLE>" & vbCrLf
	if bCentered then
		sHTML = sHTML &  "</DIV>" & vbCrLf
	end if

	Lbx_BaseHorizontalGetHTML = sHTML
end function 


'same as Lbx_BaseHorizontalGetHTML but does writes instead of building an HTML string
function Lbx_BaseHorizontalGetHTML_Write(byval bCentered, byval lngWidth, byval lngHeight, _
		byval strCtrlName, byval bMultiple, byval bDisabled, byval strHandler, _
		byref rsOptions, byval strDescFld, byval strIDFld, byref rsSel, byval strIDFldS, _
		byval lngIDS, byval bIDSel, byval bAddBlank, byval intPadSpaces, byval lngSize, _
		byval lngFontHeight )
	dim sHTML, sName
	dim lngTotalRecords
	
	if rsOptions is nothing then
		lngTotalRecords = 0
	else
		lngTotalRecords = rsOptions.RecordCount
	end if
	if bAddBlank then
		'add one since the top line will be blank
		lngSize = lngSize + 1
	end if

	if bCentered then
		response.write "<div align=""center"">" & vbCrLf
	end if
	if lngWidth > 0 then
		response.write "<table border=1 cellpadding=0 cellspacing=0 width=""" & cstr(lngWidth) & "px""" & vbCrLf
	else
		response.write "<table border=1 cellpadding=0 cellspacing=0 " & vbCrLf
	end if
	response.write "	style=""border-bottom: none; border-left: none; border-right: none; border-top: thin;"">" & vbCrLf
	response.write "<tr><td>" & vbCrLf
	response.write "<DIV style=""OVERFLOW: auto; "
	if lngWidth > 0 then
		response.write "WIDTH: " & cstr(lngWidth) & "px; "
	end if
	if lngHeight > 0 then
		response.write "HEIGHT: " & cstr(lngHeight) & "px"
	end if
	response.write """>" & vbCrLf
	if lngSize < (cint(lngHeight / lngFontHeight)) then
		'since there are less lines than what is needed for the list box,
		'build up list box so the internal SELECT tag fills the list box.
		lngSize = cint(lngHeight / lngFontHeight)
	end if
	response.write "	<select name=""" & strCtrlName & """ id=""" & strCtrlName & """ size=""" & cstr(lngSize) & """ "
	response.write "style=""font-size:" & cstr(lngFontHeight) & "px"
	'if lngTotalRecords = 0 then
		'no records so make the list box wider so it looks okay. Just picked 10 as a large number.
		response.write ";width:" & cstr(lngWidth*10) & "px"
	'end if
	if bDisabled then	'gray out the text
		response.write ";color:gray"
	end if
	response.write """ "	'close style
	if bMultiple then
		response.write "multiple "
	end if
	if bDisabled then
		response.write " disabled "
	end if
	if strHandler <> "" then
		response.write strHandler
	end if
	response.write " >" & vbCrLf

	response.write Lbx_AddOptions_Write(rsOptions, strDescFld, strIDFld, rsSel, strIDFldS, lngIDS, bIDSel, bAddBlank, intPadSpaces)

	response.write "</SELECT>" & vbCrLf
	response.write "</DIV>" & vbCrLf
	response.write "</TD></TR>" & vbCrLf
	response.write "</TABLE>" & vbCrLf
	if bCentered then
		response.write "</DIV>" & vbCrLf
	end if
end function 

function ListboxRs_HorizontalGetHTML(byval bCentered, byval lngWidth, byval lngHeight, _
		byval strCtrlName, byval bMultiple, byval bDisabled, byval strHandler, _
		byref rsOptions, byval strDescFld, byval strIDFld, byref rsSel, byval strIDFldS, _
		byval lngIDS, byval bIDSel, byval bAddBlank, byval intPadSpaces, byval lngSize, _
		byval lngFontHeight )
	ListboxRs_HorizontalGetHTML = Lbx_BaseHorizontalGetHTML( bCentered, lngWidth, lngHeight, _
		strCtrlName, bMultiple, bDisabled, strHandler, _
		rsOptions, strDescFld, strIDFld, rsSel, strIDFldS, _
		lngIDS, bIDSel, bAddBlank, intPadSpaces, lngSize, lngFontHeight )
end function

function ListboxRs_HorizontalGetHTML2(byval bCentered, byval lngWidth, byval lngHeight, _
		byval strCtrlName, byval bMultiple, byval bDisabled, byval strHandler, _
		byref rsOptions, byval strDescFld, byval strIDFld, byref rsSel, byval strIDFldS, _
		byval lngIDS, byval bIDSel, byval bAddBlank, byval intPadSpaces, byval lngSize, _
		byval lngFontHeight )
	ListboxRs_HorizontalGetHTML2 = Lbx_BaseHorizontalGetHTML2( bCentered, lngWidth, lngHeight, _
		strCtrlName, bMultiple, bDisabled, strHandler, _
		rsOptions, strDescFld, strIDFld, rsSel, strIDFldS, _
		lngIDS, bIDSel, bAddBlank, intPadSpaces, lngSize, lngFontHeight )
end function

'same as ListboxRs_HorizontalGetHTML but does writes instead of building HTML string
function ListboxRs_HorizontalGetHTML_Write(byval bCentered, byval lngWidth, byval lngHeight, _
		byval strCtrlName, byval bMultiple, byval bDisabled, byval strHandler, _
		byref rsOptions, byval strDescFld, byval strIDFld, byref rsSel, byval strIDFldS, _
		byval lngIDS, byval bIDSel, byval bAddBlank, byval intPadSpaces, byval lngSize, _
		byval lngFontHeight )
	Call Lbx_BaseHorizontalGetHTML_Write( bCentered, lngWidth, lngHeight, _
		strCtrlName, bMultiple, bDisabled, strHandler, _
		rsOptions, strDescFld, strIDFld, rsSel, strIDFldS, _
		lngIDS, bIDSel, bAddBlank, intPadSpaces, lngSize, lngFontHeight )
end function
%>