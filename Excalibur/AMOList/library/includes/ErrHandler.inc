<%
' ========================================================================================
' Generating Error Messages
' ========================================================================================
'function ErrPacific_GetHTML(Byref objErr, ByVal sErrPacific, Byval sErrMessage, ByVal sLinks)
'	dim i, bOutLoop
'	dim tblStr
'	
'	i = 1
'	bOutLoop = false
'	while (i <= objErr.Count) and (not bOutLoop)
'		if (InStr(objErr.Item(i).strDescription, sErrPacific) > 0) then
'			bOutLoop = true
'		end if
'		i = i + 1
'	wend
'	if (bOutLoop) then
'		tblStr = "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
'		tblStr = tblStr & "<tr><td><font>ERROR!</font></td></tr>"
'		tblStr = tblStr & "<tr><td><strong>" & sErrMessage & "</strong></td></tr>"
'		if (sLinks <> "") then
'			tblStr = tblStr & "<tr><td><br>" & sLinks & "</tr></td>"
'		end if
'		tblStr = tblStr & "</table>"
'	else
'		tblStr = Errors_GetHTML(objErr)
'	end if
'
'	ErrPacific_GetHTML = tblStr
'end function

' ========================================================================================
' Generating Error Messages
' ========================================================================================
function Errors_GetHTML(Byref objErr)
	dim tblStr, strDesc
	dim i

	'tblStr = "<img src=""" & Session("ServerPath") & "/library/Images/jpegs/yellowtriangle.jpg""></img>" & vbCrLf
	tblStr = "<table WIDTH=""100%"" BORDER=""0"" CELLSPACING=""1"" CELLPADDING=""1"">" & vbCRLF
	for i = 0 to objErr.Count - 1
		tblStr = tblStr & "<tr>" & vbCRLF
		tblStr = tblStr & "<td>" & vbCRLF
		strDesc = objErr.Item(i+1).strDescription
		while (InStr(strDesc, "]") > 0)
			strDesc = Mid(strDesc, InStr(strDesc, "]") + 1)
		wend
		tblStr = tblStr & "<ul><li>" & strDesc & "</li></ul>"
		tblStr = tblStr & "</td>" & vbCRLF
		tblStr = tblStr & "</tr>" & vbCRLF
	Next
	tblStr = tblStr & "</table>" & vbCRLF
	
	Errors_GetHTML = tblStr
end function

' ========================================================================================
' Generating Error Text Messages 
' ========================================================================================
function Errors_GetMsg(Byref objErr)
	dim tblStr, strDesc
	dim i

	for i = 0 to objErr.Count - 1
		
		strDesc = objErr.Item(i+1).strDescription
		while (InStr(strDesc, "]") > 0)
			strDesc = Mid(strDesc, InStr(strDesc, "]") + 1)
		wend
		
	Next
	
	
	Errors_GetMsg= strDesc 
end function

function HandleError(byref oErr)
  'Display Error
  Response.Write Errors_GetHTML(oErr)
  'free the error object
  set oErr = nothing
  Response.End 
end function

'********************************************************************************
'* Checks for specific error checks and then returns the designed error message
'* Primarily used for returning good looking messages from stored procedures
'* 10/8/2003 - Don't display duplicate error messages if they are in the array
'* 
'* Parameters:
'*  objErr  = object from call to stored procedure
'*	aErrors = array with 2 dimensions. (x,0) = string to look for to identify error message
'*                                     (x,1) = string of error message to output
'*						If the message is a dynamic one coming from SQL, then end the (x,0)
'*							string with the characters |||
'*  sLinks  = any HTML code to list after error messages which are links
'*
'* Returns:
'*  HTML code to output whether it's the found error messages or the standard
'*  error message
'********************************************************************************
function ErrSpecific_GetHTML(Byref objErr, ByVal arrErrors, ByVal strLinks)
	dim strErrMessage
	dim intCount1, intCount2
	dim tblStr
	dim arrDisplayed
	redim arrDisplayed(ubound(arrErrors))
	for intCount2 = 0 to ubound(arrErrors)
		arrDisplayed(intCount2) = 0
	next

	intCount1 = 1
	strErrMessage = ""
	while (intCount1 <= objErr.Count)
		for intCount2 = 0 to ubound(arrErrors)
			if InStr(objErr.Item(intCount1).strDescription, arrErrors(intCount2,0)) > 0 then
				if arrDisplayed(intCount2) <> 1 then
					arrDisplayed(intCount2) = 1
					'check for a dynamic message from SQL
					if right(arrErrors(intCount2,0), 3) = "|||" then
						'it is so use the right part of the description as the message
						strErrMessage = strErrMessage & "<li>" & mid(objErr.Item(intCount1).strDescription, instr(objErr.Item(intCount1).strDescription, "|||")+3)
					else
						strErrMessage = strErrMessage & "<li>" & arrErrors(intCount2,1)
					end if
					strErrMessage = strErrMessage & "</li>"
				end if
			end if
		next
		intCount1 = intCount1 + 1
	wend
	if strErrMessage <> "" then
		tblStr = tblStr & "<img src='" & Session("ServerPath") & "/library/Images/jpegs/yellowtriangle.jpg'>" & vbCrLf
		tblStr = tblStr & "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCRLF
		tblStr = tblStr & "<tr><td><ul>" & strErrMessage & "</ul></td></tr>" & vbCrLf
		if (strLinks <> "") then
			tblStr = tblStr & "<tr><td><br>" & strLinks & "</td></tr>"
		end if
		tblStr = tblStr & "</table>" & vbCRLF
	else
		tblStr = Errors_GetHTML(objErr)
	end if

	ErrSpecific_GetHTML = tblStr
end function


'function HandleError2(Byref objErr, ByVal sErrPacific, Byval sErrMessage, ByVal sLinks)
'  'Display Error
'  Response.Write ErrPacific_GetHTML( objErr, sErrPacific, sErrMessage, sLinks )
'  'free the error object
'  set objErr = nothing
'  Response.End 
'end function

'function PageError(byref sError, byval sModule)
'  dim oErr
'  set oErr = Server.CreateObject("JF_H_Error.CErrors")
'  oErr.Add 1, cstr(sError), cstr(sModule)
'  HandleError oErr
'end function

'function GenerateError(byref sError, byval sModule)
'  dim oErr
'  set oErr = Server.CreateObject("JF_H_Error.CErrors")
'  oErr.Add 1, cstr(sError), cstr(sModule)
'  set GenerateError = oErr
'  set oErr = nothing
'end function

'function Err_GetHTML(ByVal sErr, ByVal sLinks)
'	dim tblStr
'	tblStr = "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>"
'	tblStr = tblStr & "<tr><td><strong>" & sErr & "</strong></td></tr>"
'	if (sLinks <> "") then
'		tblStr = tblStr & "<tr><td><br>" & sLinks & "</tr></td>"
'	end if
'	tblStr = tblStr & "</table>"
'	Err_GetHTML = tblStr
'end function

'function MutipleSpecificErr (byref objErr, byref aErr2Chk, byval nCount)
'	dim i, nLastErr, sErr
'	nLastErr = objErr.Count
'	if (nLastErr > 0) then
'		for i = 0 to nCount
'			'check if a specific error occurs
'			if instr(1, objErr.Item(nLastErr).strDescription, aErr2Chk(i, 0)) then
'				select case aErr2Chk(i, 1)
'					case 0	'display the original message only
'						sErr = objErr.Item(nLastErr).strDescription
'					case 1	'display the addition message only
'						sErr = aErr2Chk(i, 2)
'					case 2	'display the original message and the additional message
'						sErr = objErr.Item(nLastErr).strDescription & aErr2Chk(i, 2)
'				end select
'				MutipleSpecificErr = Err_GetHTML(sErr, aErr2Chk(i, 3))
'				exit function
'			end if
'		next
'	end if
'	MultipleSpecificErr = Errors_GetHTML(objErr)
'end function

'function AddErrToCheck(byref aErr2Chk, nCol, sErr, nType, sMsg, sLink)
'	aErr2Chk(nCol, 0) = sErr
'	aErr2Chk(nCol, 1) = nType
'	aErr2Chk(nCol, 2) = sMsg
'	aErr2Chk(nCol, 3) = sLink
'end function
%>