<%
' ========================================================================================
' Include file for General functions. Note that this is more than just email functions.
' Assumes these include files are in the main ASP file calling this:
'	"/library/includes/GlobalFooter.inc"
'	"/library/includes/Overview.inc"
' ========================================================================================

'Puts all users in the BCC field
function SendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, strEventName, _
			strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink )
	set SendIRSEmail = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, -1, -1, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		"", "", True, True, "" )
end function

' This function adds the ability to pass DivisionID and CategoryID when searching for an event
' in addition to specifying the sender's Name and Email Address
'Puts all users in the BCC field
function SendIRSEmail2(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, _
			strErrorReturnLink, strFromName, strFromAddress )
	set SendIRSEmail2 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, _
		strErrorReturnLink, strFromName, strFromAddress, True, True, "" )
end function

' ========================================================================================
' This function sends emails to the specific To field
' ========================================================================================
function SendIRSEmail_To(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, strEventName, _
			strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink )
	set SendIRSEmail_To = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, -1, -1, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		"", "", True, False, "" )
end function

' ========================================================================================
' This function adds the ability to pass DivisionID and CategoryID when searching for an event
' in addition to specifying the sender's Name and Email Address
' ========================================================================================
function SendIRSEmail_To2(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, _
			strErrorReturnLink, strFromName, strFromAddress )
	set SendIRSEmail_To2 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, _
		strErrorReturnLink, strFromName, strFromAddress, True, False, "" )
end function

' ========================================================================================
' Option to not display Server name and automated message at end
' Do not care about DivisionID and CategoryID when searching for an event
' ========================================================================================
function SendIRSEmail_To3(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
			strFromName, strFromAddress, intDisplayServerAutomated )
	set SendIRSEmail_To3 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, -1, -1, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		strFromName, strFromAddress, intDisplayServerAutomated, False, "" )
end function

' ========================================================================================
' All options specified without attachments
' ========================================================================================
function SendIRSEmail4(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
			strFromName, strFromAddress, intDisplayServerAutomated, bMakeAllBCC )
	set SendIRSEmail4 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		strFromName, strFromAddress, intDisplayServerAutomated, bMakeAllBCC, "" )
end function

' ========================================================================================
' All options specified with attachments
' ========================================================================================
function SendIRSEmail5(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
			strFromName, strFromAddress, intDisplayServerAutomated, bMakeAllBCC, strAttachments )
	set SendIRSEmail5 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		strFromName, strFromAddress, intDisplayServerAutomated, bMakeAllBCC, strAttachments )
end function

' ========================================================================================
' This function sends emails to the specific To field in addition to specifying the sender's Name and Email Address
' ========================================================================================
function SendIRSEmail_To6(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, strEventName, _
			strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, strFromName, strFromAddress )
	set SendIRSEmail_To6 = BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
		strEventName, -1, -1, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, strErrorReturnLink, _
		strFromName, strFromAddress, True, False, "" )
end function

' ========================================================================================

' ========================================================================================
' BaseSendIRSEmail
' This function sends an email formatted correctly for IRS
' 
' Input: 
' strMailingListName        = name of mailing list
' strAllToEmails            = list of email addresses separated by semicolons (;) with no trailing ;
'                               This is usually blank because the Event name should pull the email
'                               addresses but there may be times when email addresses should be passed
' strAddToEmails            = this is one or more email addresses separated by semicolons (;) with no trailing ;
'                               that needs to be added to the To list whether given by strAllToEmails
'                               or from getting the email addresses via the Event name.
' strCCEmails               = list of email addresses for CC separated by semicolons (;) with no trailing ; (can be empty)
' strBCCEmails              = list of email addresses for BCC separated by semicolons (;) with no trailing ; (can be empty)
' strEventName              = mailing list event name
' lngDivisionID             = divisionid to search for with event name, if don't care use -1
' lngCatID                  = categoryid to search for with event name, if don't care use -1
' strSubject                = subject of email
' strHeadHTML               = extra HTML to put in HEAD
' strBodyHTML               = body of email in HTML
' intDisplayUnsubscribe     = type of Subscribe/Unsubscribe info that should be displayed
'						                    0=none, -1 or 1=mailing list, 2=category subscription
' strErrorReturnLink        = full link for user to return to if error (optional), e.g. <a>Link</a>
' strFromName               = Name of person sending email, empty gets "ISIS_IRS"
' strFromAddress            = Address of person sending email, empty gets "ISIS_IRS@hp.com"
' intDisplayServerAutomated = true or false if Server name and automated message should be displayed at the bottom
' bMakeAllBCC               = true or false if all recipients should be in the BCC field or not. If True and there is a BCC value
'                               passed, it is appended to the To Values
' strAttachments			= paths of the files to be attached to the email
' Returns:
' oErr (object): the returned error object
' ========================================================================================
function BaseSendIRSEmail(strMailingListName, strAllToEmails, strAddToEmails, strCCEmails, strBCCEmails, _
			strEventName, lngDivisionID, lngCatID, strSubject, strHeadHTML, strBodyHTML, intDisplayUnsubscribe, _
			strErrorReturnLink, strFromName, strFromAddress, intDisplayServerAutomated, bMakeAllBCC, strAttachments )
	Dim strServername, strToUserIDs, strToUserEmails', strValue
	dim strHTML, strName, strTemp
	dim intError, intCount
	dim oNotification, oErr, oSendMail
	dim arrSplit, arrSplitTo, arrSplitCC, arrSplitBCC
	dim fs
	
	'if instr(1,  Request.ServerVariables("Server_Name"), application("ServerFullAddr"), 1) then
	'	strServername =  Request.ServerVariables("Server_Name")
	'else
	'	strServername =  Request.ServerVariables("Server_Name") & application("ServerFullAddr")
	'end if
	strServername = Session("PrefixFullPath")

	Set oNotification = Server.CreateObject("JF_S_Notification.ISNotification")
	if trim(strAllToEmails) = "" then
		set oErr = oNotification.ViewUsersByEvent(Application("Repository"), strEventName, strToUserIDs, strToUserEmails, clng(lngDivisionID), clng(lngCatID) )
		if not oErr is nothing then 
			set BaseSendIRSEmail = oErr
			exit function
		end if
	else
		strToUserEmails = trim(strAllToEmails)
		'trim last ; just in case user gave one
		strToUserEmails = RemoveTrailingChar( strToUserEmails, ";")
	end if

	if trim(strAddToEmails) <> "" then
		'trim last ; just in case user gave one
		strAddToEmails = RemoveTrailingChar( strAddToEmails, ";")
		strToUserEmails = strToUserEmails & ";" & trim(strAddToEmails)
	end if
	
	if trim(strToUserEmails) = "" then
		'it is possible that there is no mailing list found with email addresses or
		'that everybody unsubscribed from that mailing list so it shouldn't be an error.
		'We'll leave this Call here just in case we decide there is a way to verify for sure
		'there is an error in the future.
		'Call NoEmailAddressesFound( strMailingListName, strErrorReturnLink )
		set BaseSendIRSEmail = nothing
		exit function
	end if

'	set oErr = oNotification.GetDateTimeMailHeader(strValue) 
'	if not oErr is nothing then
'		set BaseSendIRSEmail = oErr
'		exit function
'	end if

	strHTML = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">" & vbCrLf
	strHTML = strHTML & "<html>" & vbCrLf
	strHTML = strHTML & "<head>" & vbCrLf
	strHTML = strHTML & "<STYLE>BODY" & vbCrLf
	strHTML = strHTML & "{" & vbCrLf
	strHTML = strHTML & "    COLOR: black;" & vbCrLf
	strHTML = strHTML & "    FONT-FAMILY: Arial;" & vbCrLf
	strHTML = strHTML & "    FONT-SIZE: 9pt;" & vbCrLf
	strHTML = strHTML & "    MARGIN-LEFT: 10px;" & vbCrLf
	strHTML = strHTML & "    MARGIN-RIGHT: 10px" & vbCrLf
	strHTML = strHTML & "}" & vbCrLf
	strHTML = strHTML & "TABLE" & vbCrLf
	strHTML = strHTML & "{" & vbCrLf
	strHTML = strHTML & "    FONT-FAMILY: Arial;" & vbCrLf
	strHTML = strHTML & "    FONT-SIZE: 9pt;" & vbCrLf
	strHTML = strHTML & "    MARGIN: 0px;" & vbCrLf
	strHTML = strHTML & "    PADDING-BOTTOM: 0px;" & vbCrLf
	strHTML = strHTML & "    PADDING-LEFT: 0px;" & vbCrLf
	strHTML = strHTML & "    PADDING-RIGHT: 0px;" & vbCrLf
	strHTML = strHTML & "    PADDING-TOP: 0px" & vbCrLf
	strHTML = strHTML & "}</STYLE>" & vbCrLf
	strHTML = strHTML & "<meta http-equiv=""Content-Type""" & vbCrLf
	strHTML = strHTML & "content=""text/html; charset=iso-8859-1"">" & vbCrLf
	strHTML = strHTML & "<meta name=""GENERATOR""" & vbCrLf
	strHTML = strHTML & " content=" & Server.HTMLEncode(strSubject) & ">" & vbCrLf
	strHTML = strHTML & "<title>" & strMailingListName & "</title>" & vbCrLf
	if strHeadHTML <> "" then
		strHTML = strHTML & strHeadHTML & vbCrLf
	end if
	strHTML = strHTML & "</head>" & vbCrLf

	strHTML = strHTML & "<body bgcolor=""FFFFFF"">" & vbCrLf
	strHTML = strHTML & strBodyHTML
	
	if intDisplayServerAutomated then
		strHTML = strHTML & "<p><span style='font-size:8.0pt;'><b>Server: </b>"
		strHTML = strHTML & Session("ServerURL") & "</p>" & vbCrLf
	end if
	if strMailingListName <> "" then
		strHTML = strHTML & "<p><span style='font-size:8.0pt;'><b>Mailing List: </b>"
		strHTML = strHTML & strMailingListName & "&nbsp;<Br>" & vbCrLf
	end if
	select case intDisplayUnsubscribe
		case 0 		'don't display any message
		case 1, -1	'mailing list
			strHTML = strHTML & "If you would like to Subscribe/Unsubscribe to this, or any other Mailing Lists, please go to the User Preferences section within IRS,"
			strHTML = strHTML & " or click <A target = '_blank' HREF='" & strServername & "irsplus/default.aspx?Link=admin/UserPreferences.aspx'> here</A>."
		case 2		'category subscription
			strHTML = strHTML & "If you would like to Subscribe/Unsubscribe to this, or any other Category Subscription, please go to the User Preferences section within IRS,"
			strHTML = strHTML & " or click <A target = '_blank' HREF='" & strServername & "irsplus/default.aspx?Link=admin/UserPreferences.aspx'> here</A>."
	end select
	if strMailingListName <> "" then
		strHTML = strHTML & "</span></p>" & vbCrLf
	end if

	if intDisplayServerAutomated then
		strHTML = strHTML & "<p align = center> <i style='mso-bidi-font-style:normal'><span "
		strHTML = strHTML & "style='font-size:8.0pt;'>This is an automated email from IRS. Please do not respond.</span></i></p>" & vbCrLf
	end if
	strHTML = strHTML & "</body>" & vbCrLf
	strHTML = strHTML & "</html>" & vbCrLf

	Set oSendMail = Server.CreateObject("EasyMail.SMTP")
	if strFromName = "" then
		oSendMail.From = "ISIS_IRS"
	else
		'strip off leading and trailing " if there. ASPSmartMail email component needed it, EasyMail doesn't.
		if left(strFromName, 1) = chr(34) then strFromName = mid(strFromName, 2)
		if right(strFromname, 1) = chr(34) then strFromName = left(strFromName, len(strFromName)-1)
		oSendMail.From = strFromName
	end if
	if strFromAddress = "" then
		oSendMail.FromAddr = "ISIS_IRS@hp.com"
	else
		oSendMail.FromAddr = strFromAddress
	end if

'	oSendMail.DateTime = strValue
	oSendMail.Subject = strSubject
	oSendMail.BodyText = strHTML
	oSendMail.BodyFormat = 1
	oSendMail.AutoWrap = 0
	oSendMail.LicenseKey = "Hewlett-Packard  (2 Developer)/0990630B10A08B004A77"
	
	if strAttachments <> "" then
		'could have multiple attached files seperated by semicolons
		arrSplit = split(strAttachments, ";")
		Set fs = Server.CreateObject("Scripting.FileSystemObject")
		for each strName in arrSplit
			if fs.FileExists(strName) then
				intError = oSendMail.AddAttachment (trim(strName), 0)
				if intError <> 0 then
					set oErr = Server.CreateObject("JF_H_Error.CErrors")
					select case intError
						case 1
							oErr.Add 1, "Error with opening attachment. Error: " & cstr(intError), "Email Attachment Error"
						case 2
							oErr.Add 1, "Error with reading attachment. Error: " & cstr(intError), "Email Attachment Error"
						case else
							oErr.Add 1, "Error with attachment. Error: " & cstr(intError), "Email Attachment Error"
					end select
					Set BaseSendIRSEmail = oErr
					exit function
				end if
			end if
		next
		set fs = nothing
	end if
	
	if bMakeAllBCC then
		'just one name to show in the To field
		'oSendMail.AddRecipient "IRS User", "ISIS_IRS@hp.com", 1
		'don't put anything and "undisclosed-recipients" will show in the To field
	else
		strToUserEmails = RemoveTrailingChar( strToUserEmails, ";")
	end if
	
	if trim(strCCEmails) <> "" then
		strCCEmails = RemoveTrailingChar( trim(strCCEmails), ";")
	end if

	if bMakeAllBCC then
		'Really use To emails as the BCC ones
		if trim(strBCCEmails) = "" then
			strBCCEmails = strToUserEmails
		else
			strBCCEmails = strToUserEmails & ";" & RemoveTrailingChar( trim(strBCCEmails), ";")
		end if
		strToUserEmails = ""	'since these names are being sent in the BCC
	else
		if trim(strBCCEmails) <> "" then
			strBCCEmails = RemoveTrailingChar( trim(strBCCEmails), ";")
		end if
	end if

	On Error Resume Next
	'try to split up the emails with limited number of users per email
	while strToUserEmails <> "" or strCCEmails <> "" or strBCCEmails <> ""
		' Can't do this because this is just an invalid email address that keeps bouncing back
		'if len(strToUserEmails) > 0 then
		'	call AddName( oSendMail, strToUserEmails, 1)
		'else
		'	oSendMail.AddRecipient "undisclosed-recipients", "name@company.com", 1
		'end if

		call AddName( oSendMail, strToUserEmails, 1)
		call AddName( oSendMail, strCCEmails, 2)
		call AddName( oSendMail, strBCCEmails, 3)
		
		intError = oSendMail.SubmitToExpress(left(Request.ServerVariables("APPL_PHYSICAL_PATH"), 1) & ":\EasyMail SMTP Express\")
		
		If intError <> 0 Then
			set oErr = Server.CreateObject("JF_H_Error.CErrors")
			select case intError
				case 1
					oErr.Add 1, "Error with sending email. An exception has occurred. Error: " & cstr(intError), "Email Error"
				case 3
					oErr.Add 1, "Error with sending email. Out of memory. Error: " & cstr(intError), "Email Error"
				case 5
					oErr.Add 1, "Error with sending email. Access denied. Error: " & cstr(intError), "Email Error"
				case 10
					oErr.Add 1, "Error with sending email. Error opening or creating file. Among other things, this could indicate access problems with attachment files or an invalid path to the EasyMail SMTP Express installation, or invalid access rights in that path. Error: " & cstr(intError), "Email Error"
				case 11
					oErr.Add 1, "Error with sending email. Error reading from file. Error: " & cstr(intError), "Email Error"
				case 12
					oErr.Add 1, "Error with sending email. Error writing to file. Error: " & cstr(intError), "Email Error"
				case 105
					oErr.Add 1, "Error with sending email. Invalid license key. Error: " & cstr(intError), "Email Error"
				case else
					oErr.Add 1, "Error with sending email. Error: " & cstr(intError), "Email Error"
			end select
			Set BaseSendIRSEmail = oErr
			exit function
		End If
		
		oSendMail.Clear (1)	'clear recipients
	wend

	set oSendMail = nothing
	Set oNotification = Nothing

	set BaseSendIRSEmail = nothing
end function

' ========================================================================================
' AddName
' This subroutine adds the names from a string to the mail object. It limits the email to 100
' users. If there are more than 100 users, it returns the extra names in the same string.
' 
' Input: 
' oSendMail     = name of mail object
' strUserEmails = string of email addresses to send to.
' intField      = Which field to put the email addresses into: 1=To, 2=CC, 3=BCC
'
' Returns:
' strUserEmails = If there are more than 100, this gets returned with the remaining email addresses to send
' ========================================================================================
sub AddName( byref oSendMail, byref strUserEmails, byval intField)
	dim intCount
	dim strTemp, strName
	dim arrArray
	
	const PeoplePerEmail = 100	'total number of email addresses per email to be sent
	
	arrArray = split(strUserEmails, ";")
	if ubound(arrArray)+1 > 0 then
		intCount = 0
		strTemp = ""
		for each strName in arrArray
			intCount = intCount + 1
			if cint(intCount) > cint(PeoplePerEmail) then
				'build up string again
				strTemp = strTemp & strName & ";"
			else
				oSendMail.AddRecipient "", strName, intField
			end if
		next
		strUserEmails = RemoveTrailingChar( trim(strTemp), ";")
	end if
end sub

' ========================================================================================
' NoEmailAddressesFound
' This subroutine prints out an error message if no email addresses are found
' 
' Input: 
' strMailingListName = name of mailing list
' strErrorReturnLink = full link for user to return to if error (optional), e.g. <a>Link</a>
'
' Returns:
' none
' ========================================================================================
sub NoEmailAddressesFound(strMailingListName, strErrorReturnLink)
	response.write "<html>" & vbCrLf
	response.write "<head>" & vbCrLf
	response.write "<title>No Email Addresses Found</title>" & vbCrLf
	response.write "<link rel=""stylesheet"" type=""text/css"" href=""" & Session("ServerPath") & "/library/stylesheets/simple.css"">" & vbCrLf
	response.write "</head>" & vbCrLf
	response.write "<body bgcolor=""#ffffff"">" & vbCrLf
	
	response.write BuildHelp(strMailingListName, "")
	response.write "<h2>No Email Addresses Found</h2>" & vbCrLf
	response.write "<p>No email addresses were found to send a notification to. Please alert an IRS Administrator.</p>" & vbCrLf
	if strErrorReturnLink <> "" then
		response.write "<p>" & strErrorReturnLink & "</p>" & vbCrLf
	end if
	response.write "<p>&nbsp;</p>" & vbCrLf
	
	response.write "</body>" & vbCrLf
	response.write "</html>" & vbCrLf
	response.end
end sub

' ========================================================================================
' RemoveTrailingChar
' This function strips off a trailing character
' 
' Input: 
' strString = string to check
' strChar   = character to strip
'
' Returns:
' strString with trailing character stripped if present
' ========================================================================================
function RemoveTrailingChar( strString, strChar )
	if right(strString, 1) = strChar then
		RemoveTrailingChar = left(strString, len(strString)-1)
	else
		RemoveTrailingChar = strString
	end if
end function

' ========================================================================================
' replaceXMLChars
' This function replaces characters that XML doesn't like with values it does like.
' 
' Input: 
' strString = string to check
'
' Returns:
' acceptable XML string
' ========================================================================================
function replaceXMLChars(strString)
	strString = replace(strString, "&", "&amp;")
	strString = replace(strString, "<", "&lt;")
	strString = replace(strString, ">", "&gt;")
	strString = replace(strString, chr(34), "&quot;")
	strString = replace(strString, "'", "&apos;")
	replaceXMLChars = strString
end function

' ========================================================================================
' DelimitedStringToArray
' This subroutine is to convert a delimited-string into an array.
' 
' Input: 
' sInput = string to check
' aOutput = output array 
' sDelimiter = the delimiter character
'
' ========================================================================================
Sub DelimitedStringToArray(ByVal sInput, aOutput, sDelimiter)

    Dim iPos1, iPos2 
    Dim iCount
    
    If Right(sInput, 1) <> sDelimiter Then
        sInput = sInput + sDelimiter
    End If
    iPos1 = InStr(1, sInput, sDelimiter)
    iCount = 0
    While iPos1 > 0
        iCount = iCount + 1
        iPos1 = InStr(iPos1 + 1, sInput, sDelimiter)
    Wend
    
    If iCount > 0 Then
    
        ReDim aOutput(iCount - 1)
        Dim iIndex
        iIndex = 0
        iPos1 = 1
        iPos2 = InStr(1, sInput, sDelimiter)
        
        While iPos2 > 0
            aOutput(iIndex) = Mid(sInput, iPos1, iPos2 - iPos1)
            iPos1 = iPos2 + 1
            iPos2 = InStr(iPos1, sInput, sDelimiter)
            iIndex = iIndex + 1
        Wend
    End If
End Sub

' ========================================================================================
' DelimitedStringToRecordset
' This subroutine is to convert a delimited-string into a ADO recordset.
' 
' Input: 
' sInput = string to check
' sDelimiter = the delimiter character
' iSize = the length of the record
' sColName = the name of the column
' nColType = the ADO data type of the record (integer:3, varchar:200)
' rsOutput = output recordset
'
' ========================================================================================
Sub DelimitedStringToRecordset(sInput, sDelimiter, iSize, sColName, nColType, rsOutput)
    Dim iPos1, iPos2, iCount, oRs 

    Set rsOutput = Server.CreateObject("ADODB.Recordset")
    With rsOutput
        .Fields.Append sColName, nColType, iSize
        .CursorType = 2   'adUseClient
        .Open
    End With
    
    sInput = UCase(Trim(sInput))

		if sInput <> "" then
	    if Right(sInput, 1) <> sDelimiter then
	        sInput = sInput + sDelimiter
	    end if
	    iPos1 = InStr(1, sInput, sDelimiter)
	    iCount = 0
	    while iPos1 > 0
	        iCount = iCount + 1
	        iPos1 = InStr(iPos1 + 1, sInput, sDelimiter)
	    wend
	    if iCount > 0 then
	        iPos1 = 1
	        iPos2 = InStr(1, sInput, sDelimiter)
	
	        while iPos2 > 0
	            rsOutput.AddNew
	            if nColType = 3 then
					rsOutput(sColName) = clng(Mid(sInput, iPos1, iPos2 - iPos1))
				else
					rsOutput(sColName) = Mid(sInput, iPos1, iPos2 - iPos1)
				end if
	            iPos1 = iPos2 + 1
	            iPos2 = InStr(iPos1, sInput, sDelimiter)
	        wend
	    end if
		end if
end sub

' ========================================================================================
' RecordsetToDelimitedString
' This subroutine is to convert an ADO recordset to a delimited-string.
' 
' Input: 
' rsInput    = input recordset
' sDelimiter = the delimiter character
' sColName   = the name of the column to use
' ========================================================================================
function RecordsetToDelimitedString(rsInput, sDelimiter, sColName)
	dim strTemp, bFirst

	strTemp = ""
	if rsInput.RecordCount > 0 then
		rsInput.MoveFirst
		bFirst = True
		do while not rsInput.EOF
			if bFirst = False then
				strTemp = strTemp & sDelimiter
			end if
			strTemp = strTemp & rsInput(sColName).Value
			bFirst = False
			rsInput.MoveNext
		loop
	end if
	
	RecordsetToDelimitedString = strTemp
end function

'********************************************************************************
'* Makes a local recordset that we can update
'* 
'* Parameters:
'*  oRs     = Recordset to use as the source of the copy
'*  intFill = true or false if all entries in the recordset should be copied to the new one
'*
'* Returns:
'*  Local Recordset
'********************************************************************************
function CopyRecordset(oRs, intFill)
	dim oDupRs, oFld, i
	set oDupRs = Server.CreateObject ("ADODB.Recordset")

	for i = 0 to oRs.Fields.Count - 1
		set oFld =  oRs.Fields(i)
		oDupRs.Fields.Append oFld.Name, oFld.Type, oFld.DefinedSize, oFld.Attributes
	next

	oDupRs.CursorLocation = 3	'Use client-side cursors
	oDupRs.Open
	
	If oRs.RecordCount > 0 Then
	    oRs.MoveFirst
	End If

	if intFill then
		While Not oRs.EOF
			oDupRs.AddNew
			For i = 0 To oRs.Fields.Count - 1
				oDupRs.Fields(oDupRs.Fields(i).Name).Value = oRs.Fields(i).Value
			Next
			oRs.MoveNext
		Wend
	end if

	If oRs.RecordCount > 0 Then
		oRs.MoveFirst
	End If

	set CopyRecordset = oDupRs
end function

' ========================================================================================
' RecordReportUsage
' This subroutine is to record the usage of reports 
' 
' Input: ReportID, UserID, 
' ========================================================================================
Sub RecordReportUsage(ByVal ReportID, byval UserID)
	Dim oGeneral, oErr
	
	set oGeneral = Server.CreateObject("JF_S_General.ISGeneral")
	set oErr = oGeneral.RecordReportUsage (Application("Repository"), ReportID, UserID)
	
	set oGeneral =nothing
	set oErr =nothing
End Sub

' ========================================================================================
' RecordReportUsageByCode
' This subroutine is to record the usage of reports 
' 
' Input: ReportCode, UserID
' ========================================================================================
Sub RecordReportUsageByCode(ByVal ReportCode, byval UserID)
	Dim oGeneral, oErr
	
	set oGeneral = Server.CreateObject("JF_S_General.ISGeneral")
	set oErr = oGeneral.RecordReportUsageByCode (Application("Repository"), ReportCode, UserID)
	
	set oGeneral =nothing
	set oErr =nothing
End Sub

' ========================================================================================
' Writeborder_firefox
' This function is to make the border show up when ruuning on firefox for many reports 
' ========================================================================================
function Writeborder_firefox()

	dim user_agent 
	dim strborder
	user_agent = request.servervariables("HTTP_USER_AGENT")
	if (instr(1,user_agent, "firefox",1) > 0) then
		strborder = " Border=1 "
	else
		strborder = ""
	end if
	
	Writeborder_firefox = strborder
end function

'********************************************************************************
'* This function pads the necessary number of zeros to the beginning of a number
'* 
'* Parameters:
'*  strNumber   = Number in string format
'*  totalDigits = total number of digits the number should be
'*
'* Returns:
'*  number as a string with the specified number of digits and leading zeros if necessary
'********************************************************************************
Function PadZero(strNumber, totalDigits) 
    if totalDigits > len(strNumber) then 
        PadZero = String(totalDigits-len(strNumber),"0") & strNumber
    else 
        PadZero = strNumber
    end if 
End Function 

'********************************************************************************
'* This subroutine verifies the given string to determine if it is an integer only comma delimited string
'* 
'* Parameters:
'*  strString		= comma delimited string
'*  strErrorMessage = error message to display if it is an error
'*
'* Returns:
'*  Nothing. If it doesn't, an error message is produced.
'********************************************************************************
Sub IntegerCommaDelimited(strString, strErrorMessage)
	dim RegularExpressionObject, strError, oPageErr
	strError = ""
	
	Set RegularExpressionObject = New RegExp
	With RegularExpressionObject
		'.Pattern = "^([1-9][0-9]{0,7},)*[1-9][0-9]{0,7}$"	'original that doesn't allow leading 0's or any 0's alone
		.Pattern = "^([0-9][0-9]{0,7},)*[0-9][0-9]{0,7}$"
		.IgnoreCase = False
		.Global = True
	End With

	if not RegularExpressionObject.Test(strString) then
		'error string
		 Response.Write("Missing required parameters.  Unable to complete RegularExpressionObject.")
		 Response.End()
	end if

	Set RegularExpressionObject = nothing

	if strError <> "" then
		response.write "<html><body>" & vbCrLf
		response.write "<div align=""center"">" & vbCrLf
		response.write "<img src=""" & Session("ServerPath") & "/library/Images/gifs/irscomet-darkblue.gif"" width=""329"" height=""65"" border=""0"">" & vbCrLf
		response.write "</div>" & vbCrLf
		response.Write strError
		response.write "</body></html>" & vbCrLf
		response.end
	end if
End Sub 

'********************************************************************************
'* This subroutine verifies the given string to determine if it matches certain criteria
'* 
'* Note: Must also include the file "/library/includes/ErrHandler.inc" in the calling page
'* 
'* Parameters:
'*  intTest         = Test to check:
'*						1 = string can only have integers, a comma, or dash
'*  strString		= comma delimited string
'*  strErrorMessage = error message to display if it is an error
'*
'* Returns:
'*  Nothing. If it doesn't pass, an error message is produced.
'********************************************************************************
Sub ValidateStringData(intTest, strString, strErrorMessage)
	dim RegularExpressionObject, strError, oPageErr
	strError = ""
	
	Set RegularExpressionObject = New RegExp
	
	select case intTest
		case 1
			'string can only have integers, a comma, or dash
			With RegularExpressionObject
				.Pattern = "^([0-9\-\,]*)$"
				.IgnoreCase = False
				.Global = True
			End With
	end select

	if not RegularExpressionObject.Test(strString) then
		 Response.Write("Missing required parameters.  Unable to complete RegularExpressionObject.")
		 Response.End()
	end if

	Set RegularExpressionObject = nothing

	if strError <> "" then
		response.write "<html><body>" & vbCrLf
		response.write "<div align=""center"">" & vbCrLf
		response.write "<img src=""" & Session("ServerPath") & "/library/Images/gifs/irscomet-darkblue.gif"" width=""329"" height=""65"" border=""0"">" & vbCrLf
		response.write "</div>" & vbCrLf
		response.Write strError
		response.write "</body></html>" & vbCrLf
		response.end
	end if
End Sub 
%>