<%
function GetFullGridHTML(byref oRs, byref aCfg)
	GetFullGridHTML = GetPageGridHTML(oRs.RecordCount, 1, false, oRs, aCfg)
end function

function GetPageGridHTML(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg)
	dim aX, aY
	GetPageGridHTML = BasePageGridHTML(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, "", "", aX, aY, 0, 0, "")
end function

function GetPageGridHTML2(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink)
	dim aX, aY
	GetPageGridHTML2 = BasePageGridHTML(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, sEmptyDesc, sLink, aX, aX, 0, 0, "")
end function

function GetPageGridHTML3(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle)
	GetPageGridHTML3 = BasePageGridHTML(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, sEmptyDesc, sLink, aUCLink, aCLink, nUCLink, nCLink, sStyle)
end function

function GetPageGridHTML4(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, byval sTitleField)
	dim aX, aY
	GetPageGridHTML4 = BasePageGridHTML(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, sEmptyDesc, sLink, aX, aX, 0, 0, sTitleField)
end function

'following function copied from GetPageGridHTML4, handling checked checkbox issue

function GetPageGridHTML5(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, byval sTitleField, _
						 ByRef scheckboxstr)
	dim aX, aY
	GetPageGridHTML5 = BasePageGridHTML2(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, sEmptyDesc, sLink, aX, aX, 0, 0, sTitleField,scheckboxstr )
end function

function ReplaceTags( byval strString)
	strString = replace(replace(strString, "<", "&lt;"), ">", "&gt;")
	ReplaceTags = strString
end function

function ReplaceSpaces( ByVal strString)
	strString = replace(rtrim(strString), "  ", " &nbsp;")
	'try to get the first and last space to show up
	if left(strString,1) = " " and len(strString) > 1 then
		strString = "&nbsp;" + mid(strString, 2)
	end if
	if right(strString,1) = " " and len(strString) > 1 then
		strString = left(strString, len(strString)-1) + "&nbsp;"
	end if
	ReplaceSpaces = strString
end function

function BasePageGridHTML(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle)
	dim sHTML, sNewLine, sField, sIDField, nHidden, nType, sHandler, bDisplay
	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
	dim sCmpOp, sOptions, j, bStop, sID
	dim sCtrlName
	dim arrData, sTField
	
	sNewLine = chr(13) & chr(10)
    dim arrColors
	arrColors = Array("#CCCCCC", "white")

	sHTML = ""
	nHidden = 3

	if oRs.RecordCount = 0 then
		BasePageGridHTML = sEmptyDesc
		exit function
	end if

	nNumPage = (oRs.RecordCount\nNumRow)
	if (oRs.RecordCount Mod nNumRow > 0) then
		nNumPage = nNumPage + 1
	end if

	if nSelPage > nNumPage then
		nRequestedPage = 1
	else
		nRequestedPage = nSelPage
	end if

	nNumCol = ubound(aCfg)
	
	' build the table
	sHTML = "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>"

	'format column group
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
			sHTML = sHTML & "<colgroup " & aCfg(i, 6) & "></colgroup>"
		end if
	next
	' print column headers
	sHTML = sHTML & "<tr bgcolor= lightsteelblue>"
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
		   if ubound(aCfg, 2) = 8 then
		      if len(acfg(i, 7)) > 0 then
		         sHTML = sHTML & "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>"
		      else
		         sHTML = sHTML & "<th>" & aCfg(i, 1) & "</th>"
		      end if
		   else
			   sHTML = sHTML & "<th>" & aCfg(i, 1) & "</th>"
		   end if
		end if
	next
	sHTML = sHTML & "</tr>"

	' print table content
	oRs.MoveFirst

	nStartItem = nNumRow * (nRequestedPage - 1) + 1

	' Position record set to a start item
	if (nStartItem > 1) then
		arrData = oRs.GetRows(CInt(nStartItem-1))
	end if

	nCount = 1
	while (nCount <= nNumRow) and (not oRs.EOF)
		sHTML = sHTML & "<tr   bgcolor=" & arrColors(nCount mod 2) & ">"
		for i = 0 to nNumCol
			sField = aCfg(i, 0)
			nType = aCfg(i, 2)
			if nType <> nHidden then
				sHTML = sHTML & "<td>"
				select case nType
				case 1		'print a link
					sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					sHTML = sHTML & "'>" & ReplaceSpaces(ReplaceTags(oRs.Fields(sField).Value)) & "</A>"
						
				case 2		'print a check box
					sHTML = sHTML & "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						sHTML = sHTML & "<input name='" & sCtrlName & nCount 
						sHTML = sHTML & "' id='" & sCtrlName & nCount & "' type=checkbox"
						if sHandler <> "" then
							sHTML = sHTML & " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						end if
						sHTML = sHTML & ">"
					end if
				case 4		'print an operation link with string input
					sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					sHTML = sHTML & "'>" & sField & "</A>"
				case 5		'print a button
					sCtrlName = "btn" & sField
					sHTML = sHTML & "<input name='" & sCtrlName & nCount 
					sHTML = sHTML & "' id='" & sCtrlName & nCount & "' type=button "
					sHTML = sHTML & aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
					sHTML = sHTML & ">"
				case 6		'print drop down for operations
					sOptions = ""
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) & ">" & aUCLink(j) & "</OPTION>"
					next
					for j = 1 to nCLink
						sCmpOp = aCLink(j, 2)
						bStop = aCLink(j, 4)
						select case sCmpOp
							case "="
								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case "<>"
								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case else
								sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
						end select
					next
					if sOptions <> "" then
						sCtrlName = aCfg(i, 3)
						sHandler = aCfg(i, 5)
						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
						sHTML = sHTML & "<SELECT id=" & sCtrlName & sID & " name=" & _
							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
						sHTML = sHTML & "<input id=btnGo" & nCount & " name=btnGo" & _
						    nCount & " type=submit value='Go' LANGUAGE=javascript onclick='return " & sHandler & "(" & sID & ")'>"
					else
						sHTML = sHTML & "&nbsp;"
					end if
				case 11 'print a link with constant str
				    sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					sHTML = sHTML & "'>" & sField & "</A>"
				case 12		'print a link with title
					sIDField = aCfg(i, 4)
					sTField = trim(oRs.Fields(sStyle).Value)
					if trim(sTField) <> "" then
						sTField = replace(sTField,"""","''")
					end if
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					sHTML = sHTML & "' Title = " &chr(34)& sTField &chr(34)
					sHTML = sHTML & ">" & ReplaceSpaces(oRs.Fields(sField).Value) & "</A>"
				case else	'print a value
					sHTML = sHTML & oRs.Fields(sField).Value & "&nbsp;"
					'if not isnull(oRs.Fields(sField)) then 
					'	sHTML = sHTML & Server.HTMLEncode(CStr(oRs.Fields(sField).Value)) & "&nbsp;"
					'else
					'	sHTML = sHTML & "&nbsp;"
					'end if
				end select
				sHTML = sHTML & "</td>"
			else
				sIDField = aCfg(i, 4)
				sCtrlName = aCfg(i, 3)
				sHTML = sHTML & "<input type=hidden name='" & sCtrlName & nCount 
				sHTML = sHTML & "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >"
			end if
		next
		sHTML = sHTML & "</tr>"
		nCount = nCount + 1
		oRs.MoveNext
	wend
	sHTML = sHTML & "</table>"

	if sLink <> "" then	'print a link  at the bottom right of the table
		sHTML = sHTML & "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
		sHTML = sHTML & "<tr><td align='right'>" & sLink
		sHTML = sHTML & "</td></tr></table>"
	end if

	if bPrintPage then
		nEndItem = nStartItem + nNumRow - 1
		if nEndItem > oRs.RecordCount then
			nEndItem = oRs.RecordCount
		end if
		sHTML = sHTML & "<BR>"
		sHTML = sHTML & "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
		sHTML = sHTML & "<br>Pages: "
		
		if (nRequestedPage > 1) then
			'display the Prev link
				sHTML = sHTML & " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
				sHTML = sHTML & " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
		end if
		
		if (nRequestedPage < nNumPage) then
			'display the Next link
				sHTML = sHTML & " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
				sHTML = sHTML & " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
		end if

		for i = 1 to nNumPage
			if (i = nRequestedPage) then
				sHTML = sHTML & " " & i
			else
				sHTML = sHTML & " <a onclick='SelectPage(" & i & ")'" 
				sHTML = sHTML & " class='BlueHand'>" & i & "</a>"
			end if
		next

	end if
	BasePageGridHTML = sHTML
end function
    
'the following function is copied from   BasePageGridHTML to handle checkbox stuff   
function BasePageGridHTML2(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
	dim sHTML, sNewLine, sField, sIDField, nHidden, nType, sHandler, bDisplay
	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
	dim sCmpOp, sOptions, j, bStop, sID
	dim sCtrlName
	dim arrData, sTField
	
	sNewLine = chr(13) & chr(10)
    dim arrColors
		arrColors = Array("#CCCCCC", "white")

   
	sHTML = ""
	nHidden = 3

	if oRs.RecordCount = 0 then
		BasePageGridHTML2 = sEmptyDesc
		exit function
	end if

	nNumPage = (oRs.RecordCount\nNumRow)
	if (oRs.RecordCount Mod nNumRow > 0) then
		nNumPage = nNumPage + 1
	end if

	if nSelPage > nNumPage then
		nRequestedPage = 1
	else
		nRequestedPage = nSelPage
	end if

	nNumCol = ubound(aCfg)
	
	' build the table
	sHTML = "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>"

	'format column group
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
			sHTML = sHTML & "<colgroup " & aCfg(i, 6) & "></colgroup>"
		end if
	next
	' print column headers
	sHTML = sHTML & "<tr bgcolor= lightsteelblue>"
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
		   if ubound(aCfg, 2) = 8 then
		      if len(acfg(i, 7)) > 0 then
		         sHTML = sHTML & "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>"
		      else
		         sHTML = sHTML & "<th>" & aCfg(i, 1) & "</th>"
		      end if
		   else
			   sHTML = sHTML & "<th>" & aCfg(i, 1) & "</th>"
		   end if
		end if
	next
	sHTML = sHTML & "</tr>"

	' print table content
	oRs.MoveFirst

	nStartItem = nNumRow * (nRequestedPage - 1) + 1

	' Position record set to a start item
	if (nStartItem > 1) then
		arrData = oRs.GetRows(CInt(nStartItem-1))
	end if

	nCount = 1
	while (nCount <= nNumRow) and (not oRs.EOF)
		sHTML = sHTML & "<tr   bgcolor=" & arrColors(nCount mod 2) & ">"
		for i = 0 to nNumCol
			sField = aCfg(i, 0)
			nType = aCfg(i, 2)
			if nType <> nHidden then
				sHTML = sHTML & "<td>"
				select case nType
				case 1		'print a link
					sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					sHTML = sHTML & "'>" & ReplaceSpaces(ReplaceTags(oRs.Fields(sField).Value)) & "</A>"

				case 2		'print a check box
					sHTML = sHTML & "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						sHTML = sHTML & "<input name='" & sCtrlName
						sHTML = sHTML & "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=checkbox"
						if instr(1, scheckboxstr, "," & cstr(oRs.Fields(sField).Value) & "," ,1) then
							sHTML = sHTML & " checked"
						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
						end if
						
						if sHandler <> "" then
							sHTML = sHTML & " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						end if
						sHTML = sHTML & ">"
					end if
				case 4		'print an operation link with string input
					sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					sHTML = sHTML & "'>" & sField & "</A>"
				case 5		'print a button
					sCtrlName = "btn" & sField
					sHTML = sHTML & "<input name='" & sCtrlName & nCount 
					sHTML = sHTML & "' id='" & sCtrlName & nCount & "' type=button "
					sHTML = sHTML & aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
					sHTML = sHTML & ">"
				case 6		'print drop down for operations
					sOptions = ""
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) & ">" & aUCLink(j) & "</OPTION>"
					next
					for j = 1 to nCLink
						sCmpOp = aCLink(j, 2)
						bStop = aCLink(j, 4)
						select case sCmpOp
							case "="
								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case "<>"
								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case else
								sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
						end select
					next
					if sOptions <> "" then
						sCtrlName = aCfg(i, 3)
						sHandler = aCfg(i, 5)
						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
						sHTML = sHTML & "<SELECT id=" & sCtrlName & sID & " name=" & _
							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
						sHTML = sHTML & "<input id=btnGo" & nCount & " name=btnGo" & _
						    nCount & " type=submit value='Go' LANGUAGE=javascript onclick='return " & sHandler & "(" & sID & ")'>"
					else
						sHTML = sHTML & "&nbsp;"
					end if
				case 11 'print a link with constant str
				    sIDField = aCfg(i, 4)
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					sHTML = sHTML & "'>" & sField & "</A>"
				case 12		'print a link with title
					sIDField = aCfg(i, 4)
					sTField = trim(oRs.Fields(sStyle).Value)
					if trim(sTField) <> "" then
						sTField = replace(sTField,"""","''")
					end if
					sHTML = sHTML & "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					sHTML = sHTML & "' Title = " &chr(34)& sTField &chr(34)
					sHTML = sHTML & ">" & oRs.Fields(sField).Value & "</A>"
						
				case else	'print a value
					sHTML = sHTML & oRs.Fields(sField).Value & "&nbsp;"
					'if not isnull(oRs.Fields(sField)) then 
					'	sHTML = sHTML & Server.HTMLEncode(oRs.Fields(sField).Value) & "&nbsp;"
					'else
					'	sHTML = sHTML & "&nbsp;"
					'end if
				end select
				sHTML = sHTML & "</td>"
			else
				sIDField = aCfg(i, 4)
				sCtrlName = aCfg(i, 3)
				sHTML = sHTML & "<input type=hidden name='" & sCtrlName & nCount 
				sHTML = sHTML & "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >"
			end if
		next
		sHTML = sHTML & "</tr>"
		nCount = nCount + 1
		oRs.MoveNext
	wend
	sHTML = sHTML & "</table>"

	if sLink <> "" then	'print a link  at the bottom right of the table
		sHTML = sHTML & "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
		sHTML = sHTML & "<tr><td align='right'>" & sLink
		sHTML = sHTML & "</td></tr></table>"
	end if

	if bPrintPage then
		nEndItem = nStartItem + nNumRow - 1
		if nEndItem > oRs.RecordCount then
			nEndItem = oRs.RecordCount
		end if
		sHTML = sHTML & "<BR>"
		sHTML = sHTML & "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
		sHTML = sHTML & "<br>Pages: "
		
		if (nRequestedPage > 1) then
			'display the Prev link
				sHTML = sHTML & " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
				sHTML = sHTML & " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
		end if
		
		if (nRequestedPage < nNumPage) then
			'display the Next link
				sHTML = sHTML & " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
				sHTML = sHTML & " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
		end if

		for i = 1 to nNumPage
			if (i = nRequestedPage) then
				sHTML = sHTML & " " & i
			else
				sHTML = sHTML & " <a onclick='SelectPage(" & i & ")'" 
				sHTML = sHTML & " class='BlueHand'>" & i & "</a>"
			end if
		next

	end if
	BasePageGridHTML2 = sHTML
end function    
    
sub GetPageGridHTMLWrite(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg)
	' calls other Base function to do response.write's instead of building up the HTML code
	dim aX, aY
	Call BasePageGridHTMLWrite(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, "", "", aX, aY, 0, 0, "", "")
end sub

sub GetPageGridHTMLWrite6(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, byval scheckboxstr)
	' like GetPageGridHTMLWrite but check box is data aware and the scheckboxstr is not modified
	dim aX, aY
	Call BasePageGridHTMLWrite(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, "", "", aX, aY, 0, 0, "", scheckboxstr)
end sub

sub GetPageGridHTMLWrite2(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle)
	' calls other Base function to do response.write's instead of building up the HTML code
	Call BasePageGridHTMLWrite(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, sEmptyDesc, sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, "")
end sub

'copied from GetPageGridHTMLWrite, address the checkbox issue
sub GetPageGridHTMLWrite3(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, byref scheckboxstr)
	' calls other Base function to do response.write's instead of building up the HTML code
	dim aX, aY
	Call BasePageGridHTMLWrite2(nNumRow, nSelPage, bPrintPage, _
						oRs, aCfg, "", "", aX, aY, 0, 0, "",scheckboxstr)
end sub

'like GetPageGridHTMLWrite3 but allow the arrays for Operation drop downs
sub GetPageGridHTMLWrite4(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
	Call BasePageGridHTMLWrite2(nNumRow, nSelPage, bPrintPage, _
				oRs, aCfg, "", "", aUCLink, aCLink, nUCLink, nCLink, sStyle, scheckboxstr)
end sub

sub BasePageGridHTMLWrite(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, scheckboxstr)
	dim sHTML, sNewLine, sField, sIDField, nHidden, nHidden2, nType, sHandler, bDisplay
	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
	dim sCmpOp, sOptions, j, bStop, sID
	dim sCtrlName
	dim arrData, sPrefixField 
	
	sNewLine = chr(13) & chr(10)
    dim arrColors
		arrColors = Array("#CCCCCC", "white")

	sHTML = ""
	nHidden = 3
	nHidden2 = 18

	nNumPage = (oRs.RecordCount\nNumRow)
	if (oRs.RecordCount Mod nNumRow > 0) then
		nNumPage = nNumPage + 1
	end if

	if nSelPage > nNumPage then
		nRequestedPage = 1
	else
		nRequestedPage = nSelPage
	end if

	nNumCol = ubound(aCfg)
	
	' build the table
	response.write "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>"

	'format column group
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
			response.write "<colgroup " & aCfg(i, 6) & "></colgroup>"
		end if
	next
	' print column headers
	response.write "<tr bgcolor= lightsteelblue>"
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden and aCfg(i, 2) <> nHidden2 then
		   
		   if ubound(aCfg, 2) = 8 then
		      if len(acfg(i, 7)) > 0 then
		         response.write "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>"
		      else
		         response.write "<th>" & aCfg(i, 1) & "</th>"
		      end if
		   else
			   response.write "<th>" & aCfg(i, 1) & "</th>"
		   end if
		end if
	next
	response.write "</tr>"

	' print table content
	oRs.MoveFirst

	nStartItem = nNumRow * (nRequestedPage - 1) + 1

	' Position record set to a start item
	if (nStartItem > 1) then
		arrData = oRs.GetRows(CInt(nStartItem-1))
	end if

	nCount = 1
	while (nCount <= nNumRow) and (not oRs.EOF)
		response.write "<tr   bgcolor=" & arrColors(nCount mod 2) & ">"
		for i = 0 to nNumCol
			sField = aCfg(i, 0)
			nType = aCfg(i, 2)
			if nType <> nHidden and nType <> nHidden2 then
				response.write "<td>"
				select case nType
				case 1		'print a link
					sIDField = aCfg(i, 4)
					if not isnull(oRs.Fields(sField)) and trim(oRs.Fields(sField)) <> "" then 
						response.write "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
						response.write "'>" & ReplaceSpaces(ReplaceTags(oRs.Fields(sField).Value)) & "</A>"
					else
						response.write "&nbsp;"
					end if

				case 2		'print a check box
					response.write "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName & nCount 
						response.write "' id='" & sCtrlName & nCount & "' type=checkbox " & aCfg(i, 4) 
						if sHandler <> "" then
							response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						end if
						response.write ">"
					end if
				case 20		'print a data aware check box
					response.write "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName
						response.write "' id='" & sCtrlName & nCount & "' type=checkbox value=" & cstr(oRs.Fields(aCfg(i, 4)).Value) 
						if instr(1, scheckboxstr, cstr(oRs.Fields(sField).Value), 1) then
						   response.write " checked"
						end if
						if sHandler <> "" then
							response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						end if
						response.write ">"
					end if
				case 4		'print an operation link with string input
					sIDField = aCfg(i, 4)
					response.write "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>"
				case 5		'print a button
					sCtrlName = "btn" & sField
					response.write "<input name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' type=button "
					response.write aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
					response.write ">"
				case 6		'print drop down for operations
					sOptions = ""
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) & ">" & aUCLink(j) & "</OPTION>"
					next
					for j = 1 to nCLink
						sCmpOp = aCLink(j, 2)
						bStop = aCLink(j, 4)
						select case sCmpOp
							case "="
								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case "<>"
								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case else
								sOptions = sOptions & "<OPTION value=" & aCLink(j, 0) & ">" & aCLink(j, 0) & "</OPTION>"
						end select
					next
					if sOptions <> "" then
						sCtrlName = aCfg(i, 3)
						sHandler = aCfg(i, 5)
						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
						response.write "<SELECT id=" & sCtrlName & sID & " name=" & _
							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
						response.write "<input id=btnGo" & nCount & " name=btnGo" & _
						    nCount & " type=submit value='Go' LANGUAGE=javascript onclick='return " & sHandler & "(" & sID & ")'>"
					else
						response.write "&nbsp;"
					end if
				case 7	'print a drop down and highlight a value
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) 
						if aUCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>"
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>"

				case 8	'print a drop down with the display and the value different
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aCLink(j) 
						if aCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>"
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>"

				case 11 'print a link with constant str
				    sIDField = aCfg(i, 4)
					response.write "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>"
					
				case 12		'print a link with field value
				    sIDField = aCfg(i, 4)
				    sField = oRs.Fields(aCfg(i, 3))
					response.write "<A href='javascript:" & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>"
				
				case 14 'print a input box
					sField = aCfg(i, 0)
					sCtrlName = "txt" & aCfg(i, 3)
					response.write "<input name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' type=text value='" & oRs.Fields(sField).Value & "'"
					response.write aCfg(i, 6)
					response.write ">"

				case 15 'print a textarea input
					sField = aCfg(i, 1)
					sCtrlName = "txt" & aCfg(i, 0)
					response.write "<textarea name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' " 
					response.write aCfg(i, 6)
					response.write ">" & oRs.Fields(sField).Value & "</textarea>"
				
				case 16		'print a link with special chars
					sIDField = aCfg(i, 4)
					if not isnull(oRs.Fields(sField)) and trim(oRs.Fields(sField)) <> "" then 
						response.write "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					else
						response.write "&nbsp;"
					end if
					if not isnull(oRs.Fields(sField)) then 
					response.write "'>" & Server.HTMLEncode(oRs.Fields(sField).Value) & "</A>"
					else
						response.write "'>" & "&nbsp;" & "</A>"
					end if

				case 17		'print a link with db fields
					sIDField = aCfg(i, 4)
					sPrefixField = aCfg(i, 3)
					if not isnull(oRs.Fields(sPrefixField)) and trim(oRs.Fields(sPrefixField)) <> "" then 
						response.write oRs.Fields(sPrefixField).Value 
						if not isnull(oRs.Fields(sIDField)) and trim(oRs.Fields(sIDField)) <> "" then 
							response.write Server.HTMLEncode(oRs.Fields(sIDField).Value) & aCfg(i, 5)
						end if
					else
						response.write "&nbsp;"
					end if
				case 19 'print a textarea input with htmlencode, added 11/02/2004
					sField = aCfg(i, 1)
					sCtrlName = "txt" & aCfg(i, 0)
					response.write "<textarea name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' " 
					response.write aCfg(i, 6)
					response.write ">" & server.htmlencode(oRs.Fields(sField).Value) & "</textarea>"
			
				case else	'print a value
					select case ucase(oRs.Fields(sField).Value) 
						case "TRUE"
							response.write "Yes"
						case "FALSE"
							response.write "No"
						case else 
							response.write oRs.Fields(sField).Value & "&nbsp;"
					end select
				end select
				response.write "</td>"
			else
				select case nType
				case 18		'print a hidden text value between quotes with HTMLEncode()
					sIDField = aCfg(i, 4)
					sCtrlName = aCfg(i, 3)
					response.write "<input type=hidden name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' value='" & Server.HTMLEncode(oRs.Fields(sIDField).Value) & "' >"
				case else
					sIDField = aCfg(i, 4)
					sCtrlName = aCfg(i, 3)
					response.write "<input type=hidden name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >"
				end select
			end if
		next
		response.write "</tr>"
		nCount = nCount + 1
		oRs.MoveNext
	wend
	response.write "</table>"

	if sLink <> "" then	'print a link  at the bottom right of the table
		response.write "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
		response.write "<tr><td align='right'>" & sLink
		response.write "</td></tr></table>"
	end if

	if bPrintPage then
		nEndItem = nStartItem + nNumRow - 1
		if nEndItem > oRs.RecordCount then
			nEndItem = oRs.RecordCount
		end if
		response.write "<BR>"
		response.write "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
		response.write "<br>Pages: "
		
		if (nRequestedPage > 1) then
			'display the Prev link
				response.write " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
				response.write " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
		end if
		
		if (nRequestedPage < nNumPage) then
			'display the Next link
				response.write " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
				response.write " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
		end if

		for i = 1 to nNumPage
			if (i = nRequestedPage) then
				response.write " " & i
			else
				response.write " <a onclick='SelectPage(" & i & ")'" 
				response.write " class='BlueHand'>" & i & "</a>"
			end if
		next

	end if
end sub

'copied from BasePageGridHTMLWrite, addressing the checkbox issue
sub BasePageGridHTMLWrite2(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
	dim sHTML, sNewLine, sField, sIDField, nHidden, nType, sHandler, bDisplay
	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
	dim sCmpOp, sOptions, j, bStop, sID
	dim sCtrlName
	dim arrData

	sNewLine = chr(13) & chr(10)
    dim arrColors
	arrColors = Array("#CCCCCC", "white")

	sHTML = ""
	nHidden = 3

	nNumPage = (oRs.RecordCount\nNumRow)
	if (oRs.RecordCount Mod nNumRow > 0) then
		nNumPage = nNumPage + 1
	end if

	if nSelPage > nNumPage then
		nRequestedPage = 1
	else
		nRequestedPage = nSelPage
	end if

	nNumCol = ubound(aCfg)
	
	' build the table
	response.write "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>"

	'format column group
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
			response.write "<colgroup " & aCfg(i, 6) & "></colgroup>"
		end if
	next
	' print column headers
	response.write "<tr bgcolor= lightsteelblue>"
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
		   
		   if ubound(aCfg, 2) = 8 then
		      if len(acfg(i, 7)) > 0 then
		         response.write "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>"
		      else
		         response.write "<th>" & aCfg(i, 1) & "</th>"
		      end if
		   
		   else
			   response.write "<th>" & aCfg(i, 1) & "</th>"
		   end if
		end if
	next
	response.write "</tr>"

	' print table content
	oRs.MoveFirst

	nStartItem = nNumRow * (nRequestedPage - 1) + 1

	' Position record set to a start item
	if (nStartItem > 1) then
		arrData = oRs.GetRows(CInt(nStartItem-1))
	end if

	nCount = 1
	while (nCount <= nNumRow) and (not oRs.EOF)
		response.write "<tr   bgcolor=" & arrColors(nCount mod 2) & ">"
		for i = 0 to nNumCol
			sField = aCfg(i, 0)
			nType = aCfg(i, 2)
			if nType <> nHidden then
				response.write "<td>"
				select case nType
				case 1		'print a link
					sIDField = aCfg(i, 4)
					response.write "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					response.write "'>" & ReplaceSpaces(ReplaceTags(oRs.Fields(sField).Value)) & "</A>"
					
				case 2		'print a check box
				'sIDField = aCfg(i, 4)
				 '  response.write sIDField
					'sIDField = "TestProcedureID"
					response.write "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName 
						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=checkbox"
						'if sHandler <> "" then
						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						'end if
						
						if instr(1, scheckboxstr, "," & cstr(oRs.Fields(sField).Value) & "," ,1) then
						   response.write " checked"
						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
						end if

						response.write ">"
					end if
					
				case 13		'print a radio
				'sIDField = aCfg(i, 4)
				 '  response.write sIDField
					'sIDField = "TestProcedureID"
					response.write "&nbsp;"
					bDisplay = true
'					if sField <> "" then
'					  if oRs.Fields(sField).Value = clng(aCfg(i, 4)) then
'						bDisplay = false
'					  end if
'					end if
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName 
						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=radio"
						'if sHandler <> "" then
						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						'end if
						
						if instr(1, scheckboxstr,  cstr(oRs.Fields(sField).Value)  ,1) then
						   response.write " checked"
						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
						end if
						
						response.write ">"
					end if
				case 4		'print an operation link with string input
					sIDField = aCfg(i, 4)
					response.write "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>"
				case 5		'print a button
					sCtrlName = "btn" & sField
					response.write "<input name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' type=button "
					response.write aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
					response.write ">"
					case 6		'print drop down for operations
					sOptions = ""
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=""" & aUCLink(j) & """>" & aUCLink(j) & "</OPTION>"
					next
					for j = 1 to nCLink
						sCmpOp = aCLink(j, 2)
						bStop = aCLink(j, 4)
						select case sCmpOp
							case "="
								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case "<>"
								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
									end if
								end if
							case else
								sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
						end select
					next
					if sOptions <> "" then
						sCtrlName = aCfg(i, 3)
						sHandler = aCfg(i, 5)
						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
						response.write "<SELECT id=" & sCtrlName & sID & " name=" & _
							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
						response.write "<input id=btnGo" & nCount & " name=btnGo" & _
						    nCount & " type=submit value='Go' LANGUAGE=javascript onclick='return " & sHandler & "(" & sID & ")'>"
					else
						response.write "&nbsp;"
					end if

				case 7	'print a drop down and highlight a value
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) 
						if aUCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>"
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>"

				case 8	'print a drop down with the display and the value different
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aCLink(j) 
						if aUCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>"
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>"

				case 11 'print a link with constant str
				    sIDField = aCfg(i, 4)
					response.write "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>"
					
				case else	'print a value
					response.write oRs.Fields(sField).Value & "&nbsp;"
					'if not isnull(oRs.Fields(sField)) then 
					'	response.write Server.HTMLEncode(CStr(oRs.Fields(sField).Value)) & "&nbsp;"
					'else
					'	response.write "&nbsp;"
					'end if

				end select
				response.write "</td>"
			else
				sIDField = aCfg(i, 4)
				sCtrlName = aCfg(i, 3)
				response.write "<input type=hidden name='" & sCtrlName & nCount 
				response.write "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >"
			end if
		next
		response.write "</tr>"
		nCount = nCount + 1
		oRs.MoveNext
	wend
	response.write "</table>"

	if sLink <> "" then	'print a link  at the bottom right of the table
		response.write "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
		response.write "<tr><td align='right'>" & sLink
		response.write "</td></tr></table>"
	end if

	if bPrintPage then
		nEndItem = nStartItem + nNumRow - 1
		if nEndItem > oRs.RecordCount then
			nEndItem = oRs.RecordCount
		end if
		response.write "<BR>"
		response.write "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
		response.write "<br>Pages: "
		
		if (nRequestedPage > 1) then
			'display the Prev link
				response.write " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
				response.write " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
		end if
		
		if (nRequestedPage < nNumPage) then
			'display the Next link
				response.write " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
				response.write " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
		end if

		for i = 1 to nNumPage
			if (i = nRequestedPage) then
				response.write " " & i
			else
				response.write " <a onclick='SelectPage(" & i & ")'" 
				response.write " class='BlueHand'>" & i & "</a>"
			end if
		next

	end if
end sub


function GetGridSelectedItems(byval sChkPrefix, byval sIDPrefix, byval sDelimiter, _
							byval nPage, byval nNumRow)
	'only get the selected items on the submit 
	dim sSelItems, i
    for i = 1 to nNumRow
      if Request.Form(sChkPrefix & i) = "on" then
        sSelItems = sSelItems & Request.Form(sIDPrefix & i) & sDelimiter
      end if
    next
    'strip off the last comma separator
    if sSelItems <> "" then sSelItems = left(sSelItems, len(sSelItems) - len(sDelimiter))
	GetGridSelectedItems = sSelItems
end function

function MarkGridSelectedItems(byval sChkPrefix, byval nPage, byval nNumRow, _
					byval sMarkField, byref oRs)
	if oRs.RecordCount = 0 then
		exit function
	end if
	
	dim i, nStartItem, nEndItem
	GetInterval nPage, nNumRow, oRs.RecordCount, nStartItem, nEndItem
	oRs.Move nStartItem - 1, 1
	for i = 1 to (nEndItem - nStartItem + 1)
      if Request.Form(sChkPrefix & i) = "on" then
		oRs.Fields(sMarkField).Value = true
	  else
		oRs.Fields(sMarkField).Value = false
	  end if
	  oRs.MoveNext
    next
end function

function GetInterval (byval nPage, byval nNumRowPPage, byval nTotalRow, _
					byref nStartItem, byref nEndItem)
	nStartItem = nNumRowPPage * (nPage - 1) + 1
	nEndItem = nStartItem + nNumRowPPage - 1
	if nEndItem > nTotalRow then
		nEndItem = nTotalRow
	end if
end function

function GetMarkedItems(byval sMarkField, byval sIDField, byval sDelimiter, byref oRs)
	dim sSelItems
	if oRs.RecordCount > 0 then
		oRs.MoveFirst
		while not oRs.EOF
			if oRs.Fields(sMarkField).Value then
				sSelItems = sSelItems & oRs.Fields(sIDField).Value & sDelimiter
			end if
			oRs.MoveNext
		wend
	end if
    'strip off the last comma separator
    if sSelItems <> "" then sSelItems = left(sSelItems, len(sSelItems) - len(sDelimiter))
	GetMarkedItems = sSelItems
end function

function ConfigColumn(byref aCfg, byval nCol, byval sField, byval sColName, _
					byval nColType, byval sLinkPrefix, byval sLinkField, byval sLinkSuffix, byval sColFormat)
	aCfg(nCol, 0) = sField
	aCfg(nCol, 1) = sColName
	aCfg(nCol, 2) = nColType
	aCfg(nCol, 3) = sLinkPrefix
	aCfg(nCol, 4) = sLinkField
	aCfg(nCol, 5) = sLinkSuffix
	aCfg(nCol, 6) = sColFormat
end function

'following function is added for the column sorting 
function ConfigColumn2(byref aCfg, byval nCol, byval sField, byval sColName, _
					byval nColType, byval sLinkPrefix, byval sLinkField, byval sLinkSuffix, byval sColFormat, byval sSortHTML, byval sSortFld, byval sArrowSrc )
	aCfg(nCol, 0) = sField
	aCfg(nCol, 1) = sColName
	aCfg(nCol, 2) = nColType
	aCfg(nCol, 3) = sLinkPrefix
	aCfg(nCol, 4) = sLinkField
	aCfg(nCol, 5) = sLinkSuffix
	aCfg(nCol, 6) = sColFormat
	aCfg(nCol, 7) = sSortHTML
	if sSortFld = sField then
	   aCfg(nCol, 8) = sArrowSrc
	else   
	    aCfg(nCol, 8) = ""
	end if 
end function

'8/26/02 - THIS IS A DUPLICATE FUNCTION. COMMENTING THIS ONE OUT FOR NOW JUST IN CASE IT WAS THE ONE DESIRED TO BE USED.
'sub BasePageGridHTMLWrite3(ByVal nNumRow, ByVal nSelPage, _
'						 byval bPrintPage, byref oRs, byref aCfg, _
'						 byval sEmptyDesc, byval sLink, _
'						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
'	dim sHTML, sNewLine, sField, sIDField, nHidden, nType, sHandler, bDisplay
'	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
'	dim sCmpOp, sOptions, j, bStop, sID
'	dim sCtrlName
'	dim arrData
'
'	sNewLine = chr(13) & chr(10)
'    dim arrColors
'	arrColors = Array("#CCCCCC", "white")
'
'	sHTML = ""
'	nHidden = 3
'
'	nNumPage = (oRs.RecordCount\nNumRow)
'	if (oRs.RecordCount Mod nNumRow > 0) then
'		nNumPage = nNumPage + 1
'	end if
'
'	if nSelPage > nNumPage then
'		nRequestedPage = 1
'	else
'		nRequestedPage = nSelPage
'	end if
'
'	nNumCol = ubound(aCfg)
'	
'	' build the table
'	response.write "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>"
'
'	'format column group
'	for i = 0 to nNumCol
'		if aCfg(i, 2) <> nHidden then
'			response.write "<colgroup " & aCfg(i, 6) & "></colgroup>"
'		end if
'	next
'	' print column headers
'	response.write "<tr bgcolor= lightsteelblue>"
'	for i = 0 to nNumCol
'		if aCfg(i, 2) <> nHidden then
'		   
'		   if ubound(aCfg, 2) = 8 then
'		      if len(acfg(i, 7)) > 0 then
'		         response.write "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>"
'		      else
'		         response.write "<th>" & aCfg(i, 1) & "</th>"
'		      end if
'		   
'		   else
'			   response.write "<th>" & aCfg(i, 1) & "</th>"
'		   end if
'		end if
'	next
'	response.write "</tr>"
'
'	' print table content
'	oRs.MoveFirst
'
'	nStartItem = nNumRow * (nRequestedPage - 1) + 1
'
'	' Position record set to a start item
'	if (nStartItem > 1) then
'		arrData = oRs.GetRows(CInt(nStartItem-1))
'	end if
'
'	nCount = 1
'	while (nCount <= nNumRow) and (not oRs.EOF)
'		response.write "<tr   bgcolor=" & arrColors(nCount mod 2) & ">"
'		for i = 0 to nNumCol
'			sField = aCfg(i, 0)
'			nType = aCfg(i, 2)
'			if nType <> nHidden then
'				response.write "<td>"
'				select case nType
'				case 1		'print a link
'					sIDField = aCfg(i, 4)
'					response.write "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
'					response.write "'>" & replace(trim(oRs.Fields(sField).Value), "  ", " &nbsp;") & "</A>"
'					'8/16/02 - the following is what we want but it is too close to production release and it
'					'would require testing the whole system again to make sure nothing broke.
'					'sHTML = sHTML & "'>" & ReplaceSpaces(oRs.Fields(sField).Value) & "</A>"
'				case 2		'print a check box
'				'sIDField = aCfg(i, 4)
'				 '  response.write sIDField
'					'sIDField = "TestProcedureID"
'					response.write "&nbsp;"
'					bDisplay = true
''					if sField <> "" then
''					  if oRs.Fields(sField).Value = clng(aCfg(i, 4)) then
''						bDisplay = false
''					  end if
''					end if
'					sCtrlName = aCfg(i, 3)
'					sHandler = aCfg(i, 5)
'					if bDisplay then
'						response.write "<input name='" & sCtrlName 
'						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=checkbox"
'						'if sHandler <> "" then
'						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
'						'end if
'						
'						if instr(1, scheckboxstr, "," & cstr(oRs.Fields(sField).Value) & "," ,1) then
'						   response.write " checked"
'						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
'						end if
'						
'						
'						response.write ">"
'					end if
'					
'				case 13		'print a radio
'				'sIDField = aCfg(i, 4)
'				 '  response.write sIDField
'					'sIDField = "TestProcedureID"
'					response.write "&nbsp;"
'					bDisplay = true
''					if sField <> "" then
''					  if oRs.Fields(sField).Value = clng(aCfg(i, 4)) then
''						bDisplay = false
''					  end if
''					end if
'					sCtrlName = aCfg(i, 3)
'					sHandler = aCfg(i, 5)
'					if bDisplay then
'						response.write "<input name='" & sCtrlName 
'						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=radio"
'						'if sHandler <> "" then
'						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
'						'end if
'						
'						if instr(1, scheckboxstr,  cstr(oRs.Fields(sField).Value)  ,1) then
'						   response.write " checked"
'						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
'						end if
'						
'						
'						response.write ">"
'					end if
'				case 4		'print an operation link with string input
'					sIDField = aCfg(i, 4)
'					response.write "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
'					response.write "'>" & sField & "</A>"
'				case 5		'print a button
'					sCtrlName = "btn" & sField
'					response.write "<input name='" & sCtrlName & nCount 
'					response.write "' id='" & sCtrlName & nCount & "' type=button "
'					response.write aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
'					response.write ">"
'				case 6		'print drop down for operations
'					sOptions = ""
'					for j = 1 to nUCLink
'						if aUCLink(j) = "State" then
'							if oRs.Fields(aUCLink(j)).Value = "Enabled" then
'								sOptions = sOptions & "<OPTION value='Disabled'" & ">" & "Set Disabled" & "</OPTION>"
'							else
'								sOptions = sOptions & "<OPTION value='Enabled' " & ">" & "Set Enabled" & "</OPTION>"
'							end if
'						else
'							sOptions = sOptions & "<OPTION value=""" & aUCLink(j) & """>" & aUCLink(j) & "</OPTION>"
'						end if
'					next
'					for j = 1 to nCLink
'						sCmpOp = aCLink(j, 2)
'						bStop = aCLink(j, 4)
'						select case sCmpOp
'							case "="
'								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
'									if bStop then
'										exit for
'									else
'										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
'									end if
'								end if
'							case "<>"
'								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
'									if bStop then
'										exit for
'									else
'										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
'									end if
'								end if
'							case else
'								sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>"
'						end select
'					next
'					if sOptions <> "" then
'						sCtrlName = aCfg(i, 3)
'						sHandler = aCfg(i, 5)
'						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
'						response.write "<SELECT id=" & sCtrlName & sID & " name=" & _
'							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
'						response.write "<input id=btnGo" & nCount & " name=btnGo" & _
'						    nCount & " type=submit value='Go' LANGUAGE=javascript onclick='return " & sHandler & "(" & sID & ")'>"
'					else
'						response.write "&nbsp;"
'					end if
'
'				case 7	'print a drop down and highlight a value
'					sOptions = ""
'					sCtrlName = aCfg(i, 3)
'					for j = 1 to nUCLink
'						sOptions = sOptions & "<OPTION value=" & aUCLink(j) 
'						if aUCLink(j) = oRs.Fields(sField).Value then
'							sOptions = sOptions & " selected"
'						end if
'						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>"
'					next
'					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
'						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>"
'
'				case 11 'print a link with constant str
'				    sIDField = aCfg(i, 4)
'					response.write "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
'					response.write "'>" & sField & "</A>"
'					
'				case else	'print a value
'					response.write oRs.Fields(sField).Value & "&nbsp;"
'				end select
'				response.write "</td>"
'			else
'				sIDField = aCfg(i, 4)
'				sCtrlName = aCfg(i, 3)
'				response.write "<input type=hidden name='" & sCtrlName & nCount 
'				response.write "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >"
'			end if
'		next
'		response.write "</tr>"
'		nCount = nCount + 1
'		oRs.MoveNext
'	wend
'	response.write "</table>"
'
'	if sLink <> "" then	'print a link  at the bottom right of the table
'		response.write "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
'		response.write "<tr><td align='right'>" & sLink
'		response.write "</td></tr></table>"
'	end if
'
'	if bPrintPage then
'		nEndItem = nStartItem + nNumRow - 1
'		if nEndItem > oRs.RecordCount then
'			nEndItem = oRs.RecordCount
'		end if
'		response.write "<BR>"
'		response.write "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
'		response.write "<br>Pages: "
'		
'		if (nRequestedPage > 1) then
'			'display the Prev link
'				response.write " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
'				response.write " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
'		end if
'		
'		if (nRequestedPage < nNumPage) then
'			'display the Next link
'				response.write " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
'				response.write " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
'		end if
'
'		for i = 1 to nNumPage
'			if (i = nRequestedPage) then
'				response.write " " & i
'			else
'				response.write " <a onclick='SelectPage(" & i & ")'" 
'				response.write " class='BlueHand'>" & i & "</a>"
'			end if
'		next
'
'	end if
'end sub

sub GetPageGridHTMLWrite5(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
	Call BasePageGridHTMLWrite3(nNumRow, nSelPage, bPrintPage, _
				oRs, aCfg, "", "", aUCLink, aCLink, nUCLink, nCLink, sStyle, scheckboxstr)
end sub

'copied from BasePageGridHTMLWrite, addressing the checkbox issue
sub BasePageGridHTMLWrite3(ByVal nNumRow, ByVal nSelPage, _
						 byval bPrintPage, byref oRs, byref aCfg, _
						 byval sEmptyDesc, byval sLink, _
						 aUCLink, aCLink, nUCLink, nCLink, sStyle, byref scheckboxstr)
	dim sHTML, sNewLine, sField, sIDField, nHidden, nType, sHandler, bDisplay
	dim i, nCount, nNumCol, nNumPage, nRequestedPage, nStartItem, nEndItem
	dim sCmpOp, sOptions, j, bStop, sID
	dim sCtrlName
	dim arrData
	dim arrColors

	sNewLine = chr(13) & chr(10)
	arrColors = Array("#CCCCCC", "white")

	sHTML = ""
	nHidden = 3

	nNumPage = (oRs.RecordCount\nNumRow)
	if (oRs.RecordCount Mod nNumRow > 0) then
		nNumPage = nNumPage + 1
	end if

	if nSelPage > nNumPage then
		nRequestedPage = 1
	else
		nRequestedPage = nSelPage
	end if

	nNumCol = ubound(aCfg)
	
	' build the table
	response.write "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>" & vbCRLF

	'format column group
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
			response.write "<colgroup " & aCfg(i, 6) & "></colgroup>" & vbCRLF
		end if
	next
	' print column headers
	response.write "<tr bgcolor= lightsteelblue>"
	for i = 0 to nNumCol
		if aCfg(i, 2) <> nHidden then
		   
		   if ubound(aCfg, 2) = 8 then
		      if len(acfg(i, 7)) > 0 then
		         response.write "<th>" & acfg(i, 7) & " " & acfg(i, 8) & "</th>" & vbCRLF
		      else
		         response.write "<th>" & aCfg(i, 1) & "</th>" & vbCRLF
		      end if
		   
		   else
			   response.write "<th>" & aCfg(i, 1) & "</th>" & vbCRLF
		   end if
		end if
	next
	response.write "</tr>"

	' print table content
	oRs.MoveFirst

	nStartItem = nNumRow * (nRequestedPage - 1) + 1

	' Position record set to a start item
	if (nStartItem > 1) then
		arrData = oRs.GetRows(CInt(nStartItem-1))
	end if

	nCount = 1
	while (nCount <= nNumRow) and (not oRs.EOF)
		response.write "<tr   bgcolor=" & arrColors(nCount mod 2) & ">" & vbCRLF
		for i = 0 to nNumCol
			sField = aCfg(i, 0)
			nType = aCfg(i, 2)
			if nType <> nHidden then
				response.write "<td>"
				select case nType
				case 1		'print a link
					sIDField = aCfg(i, 4)
					response.write "<A href='" & aCfg(i, 3) & oRs.Fields(sIDField).Value & aCfg(i, 5)
					response.write "'>" & ReplaceSpaces(ReplaceTags(oRs.Fields(sField).Value)) & "</A>"
					
				case 2		'print a check box
				'sIDField = aCfg(i, 4)
				 '  response.write sIDField
					'sIDField = "TestProcedureID"
					response.write "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName 
						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=checkbox"
						'if sHandler <> "" then
						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						'end if
						
						if instr(1, scheckboxstr, "," & cstr(oRs.Fields(sField).Value) & "," ,1) then
						   response.write " checked"
						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
						end if
						
						response.write ">" & vbCRLF
					end if
					
				case 13		'print a radio
				'sIDField = aCfg(i, 4)
				 '  response.write sIDField
					'sIDField = "TestProcedureID"
					response.write "&nbsp;"
					bDisplay = true
					sCtrlName = aCfg(i, 3)
					sHandler = aCfg(i, 5)
					if bDisplay then
						response.write "<input name='" & sCtrlName 
						response.write "' id='" & sCtrlName & "' value= " & oRs.Fields(sField).Value & " type=radio"
						'if sHandler <> "" then
						'	response.write " onclick='return " & sHandler & "(""" & sCtrlName & nCount & """)'"
						'end if
						
						if instr(1, scheckboxstr,  cstr(oRs.Fields(sField).Value)  ,1) then
						   response.write " checked"
						   scheckboxstr = replace(scheckboxstr,"," & cstr(oRs.Fields(sField).Value) & ",","")
						end if
						
						response.write ">" & vbCRLF
					end if
				case 4		'print an operation link with string input
					sIDField = aCfg(i, 4)
					response.write "<A href='javascript:" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>" & vbCRLF
				case 5		'print a button
					sCtrlName = "btn" & sField
					response.write "<input name='" & sCtrlName & nCount 
					response.write "' id='" & sCtrlName & nCount & "' type=button "
					response.write aCfg(i, 3) & aCfg(i, 4) & aCfg(i, 5)
					response.write ">" & vbCRLF
				case 6		'print drop down for operations
					sOptions = ""
					for j = 1 to nUCLink					
						if aUCLink(j) = "State" then
						    if cint(session("GroupID")) = oRs("TestGroupID") then
								if oRs.Fields(aUCLink(j)).Value = "Enabled" then
									sOptions = sOptions & "<OPTION value='Disabled'" & ">" & "Set Disabled" & "</OPTION>" & vbCRLF
								else
									sOptions = sOptions & "<OPTION value='Enabled' " & ">" & "Set Enabled" & "</OPTION>" & vbCRLF
								end if
							end if
						else
							sOptions = sOptions & "<OPTION value=""" & aUCLink(j) & """>" & aUCLink(j) & "</OPTION>" & vbCRLF
						end if
					next
					for j = 1 to nCLink
						sCmpOp = aCLink(j, 2)
						bStop = aCLink(j, 4)
						select case sCmpOp
							case "="
								if oRs.Fields(aCLink(j, 1)).Value = aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>" & vbCRLF
									end if
								end if
							case "<>"
								if oRs.Fields(aCLink(j, 1)).Value <> aCLink(j, 3) then
									if bStop then
										exit for
									else
										sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>" & vbCRLF
									end if
								end if
							case else
								sOptions = sOptions & "<OPTION value=""" & aCLink(j, 0) & """>" & aCLink(j, 0) & "</OPTION>" & vbCRLF
						end select
					next
					if sOptions <> "" then
						sCtrlName = aCfg(i, 3)
						sHandler = aCfg(i, 5)
						sID = cstr(oRs.Fields(aCfg(i, 4)).Value)
						response.write "<SELECT id=" & sCtrlName & sID & " name=" & _
							sCtrlName & sID & " " & sStyle & ">" & sOptions & "</SELECT>&nbsp;"
						response.write "<input id=btnGo" & nCount & " name=btnGo" & _
						     nCount & " type=submit value='Go' LANGUAGE=javascript onclick=""return " & sHandler & _
								 "(" & sID & ",'" & Replace(ReplaceQuotesGrid(oRs.Fields("Name").Value), "\", "\\") & "')"">" & vbCRLF
					else
						response.write "&nbsp;"
					end if

				case 7	'print a drop down and highlight a value
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aUCLink(j) 
						if aUCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>" & vbCRLF
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>" & vbCRLF

				case 8	'print a drop down with the display and the value different
					sOptions = ""
					sCtrlName = aCfg(i, 3)
					for j = 1 to nUCLink
						sOptions = sOptions & "<OPTION value=" & aCLink(j) 
						if aCLink(j) = oRs.Fields(sField).Value then
							sOptions = sOptions & " selected"
						end if
						sOptions = sOptions & ">" & aUCLink(j) & "</OPTION>" & vbCRLF
					next
					response.write "<SELECT id=" & sCtrlName & nCount & " name=" & _
						sCtrlName & nCount & " " & sStyle & ">" & sOptions & "</SELECT>" & vbCRLF

				case 11 'print a link with constant str
				    sIDField = aCfg(i, 4)
					response.write "<A href='" & aCfg(i, 3) & replace(oRs.Fields(sIDField).Value, "\", "\\") & aCfg(i, 5)
					response.write "'>" & sField & "</A>" & vbCRLF
					
				case else	'print a value
					response.write oRs.Fields(sField).Value & "&nbsp;"
					'if not isnull(oRs.Fields(sField)) then 
					'	response.write Server.HTMLEncode(CStr(oRs.Fields(sField).Value)) & "&nbsp;"
					'else
					'	response.write "&nbsp;"
					'end if

				end select
				response.write "</td>"
			else
				sIDField = aCfg(i, 4)
				sCtrlName = aCfg(i, 3)
				response.write "<input type=hidden name='" & sCtrlName & nCount 
				response.write "' id='" & sCtrlName & nCount & "' value=" & oRs.Fields(sIDField).Value & " >" & vbCRLF
			end if
		next
		response.write "</tr>" & vbCRLF
		nCount = nCount + 1
		oRs.MoveNext
	wend
	response.write "</table>" & vbCRLF

	if sLink <> "" then	'print a link  at the bottom right of the table
		response.write "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
		response.write "<tr><td align='right'>" & sLink
		response.write "</td></tr></table>" & vbCRLF
	end if

	if bPrintPage then
		nEndItem = nStartItem + nNumRow - 1
		if nEndItem > oRs.RecordCount then
			nEndItem = oRs.RecordCount
		end if
		response.write "<BR>"
		response.write "<b>" & nStartItem & " to " & nEndItem & " of " & oRs.RecordCount & " displayed</b>"
		response.write "<br>Pages: "
		
		if (nRequestedPage > 1) then
			'display the Prev link
				response.write " <a onclick='SelectPage(" & (nRequestedPage - 1) & ")'" 
				response.write " class='BlueHand'>Prev</a>&nbsp;&nbsp;&nbsp;"		
		end if
		
		if (nRequestedPage < nNumPage) then
			'display the Next link
				response.write " <a onclick='SelectPage(" & (nRequestedPage + 1) & ")'" 
				response.write " class='BlueHand'>Next</a>&nbsp;&nbsp;&nbsp;"		
		end if

		for i = 1 to nNumPage
			if (i = nRequestedPage) then
				response.write " " & i
			else
				response.write " <a onclick='SelectPage(" & i & ")'" 
				response.write " class='BlueHand'>" & i & "</a>"
			end if
		next

	end if
end sub

'********************************************************************************
'* Replaces single and double quotes with the corresponding hex character
'* so Javascript can accept the values. Make sure the Javascript then uses
'* UNESCAPE to get it back to normal format.
'* 
'* Parameters:
'*  strString = string to replace the quotes
'*
'* Returns:
'*  string with characters replaced
'********************************************************************************
function ReplaceQuotesGrid( strString )
	if isnull(strString) or strString = "" then
		'don't do anything
	else
		'replace single quotes
		strString = replace(strString, "'", "%27")
		'replace double quotes
		strString = replace(strString, chr(34), "%22")
	end if
	
	ReplaceQuotesGrid = strString
end function
%>