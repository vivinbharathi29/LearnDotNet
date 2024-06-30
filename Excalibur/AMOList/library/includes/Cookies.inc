<% 
' ========================================================================================
' Script for Cookies
' ========================================================================================

' ========================================================================================
' function GetDBCookie
' This function returns the database cookie
'
' Input: 
' strName = name of the cookie
' The user requesting the cookie is already know via a session variable
'
' Output:
' string of the cookie value
' ========================================================================================
function GetDBCookie( byval strName )
	dim oRecords, oError, oRsCookie
	dim strCookie

	strCookie = ""
	if Session("AMOUserID") <> "" and Session("AMOUserID") <> 0 then
		'set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
		set oRecords = New ISGeneral
		set oRsCookie = oRecords.GetDBCookie (Application("Repository"), clng(Session("AMOUserID")), strName)
		if (oRsCookie is nothing) then
			Response.Write("Empty Recordset: oRsCookie")
			Response.End()
		end if

		if oRsCookie.RecordCount > 0 then
			strCookie = oRsCookie("Value")
		end if
		oRsCookie.Close
		set oRsCookie = nothing
		set oRecords = nothing
	end if

	GetDBCookie = strCookie
end function

' ========================================================================================
' function SaveDBCookie
' This function saves the database cookie
'
' Input: 
' strName  = name of the cookie
' strValue = value to save
' The user saving the cookie is already known via a session variable
'
' Output:
' None
' ========================================================================================
function SaveDBCookie( byval strName, byval strValue )
	dim oRecords, oError, sError

	if Session("AMOUserID") <> "" and Session("AMOUserID") <> 0 then
		'set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
		set oRecords = New ISGeneral
		sError = oRecords.SaveDBCookie (Application("Repository"), clng(Session("AMOUserID")), strName, strValue)
		if sError <> "True" then
			Response.Write("Missing required parameters.  Unable to complete your SaveDBCookie request. Tried to Save cookie: " & strName & " with value: " & strValue)
			Response.End()
		end if
		set oRecords = nothing
	end if
end function

' ========================================================================================
' function GetDBCookieSet
' This function returns the database cookie
'
' Input: 
' strName = name of the cookie
' The user requesting the cookie is already know via a session variable
'
' Output:
' string of the cookie value
' ========================================================================================
function GetDBCookieSet( byval strName )
	dim oRecords, oError
	dim oRsCookie

	set oRsCookie = nothing
	if Session("AMOUserID") <> "" and Session("AMOUserID") <> 0 then
		'set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
		set oRecords = New ISGeneral
		set oRsCookie = oRecords.GetDBCookieSet (Application("Repository"), clng(Session("AMOUserID")), strName)
		if (oRsCookie is nothing) then
			Response.Write("Empty Recordset: oRsCookie")
			Response.End()
		end if
		
		set oRecords = nothing
	end if

	set GetDBCookieSet = oRsCookie
end function

' ========================================================================================
' function SaveDBCookieSet
' This function saves the strings as database cookies. The recordset can be passed as an actual recordset
' or as a comma delimited string to turn into a recordset.
'
' Input: 
' strName   = name of the cookie
' strString = comma delimited string to use as the recordset
' oRsValues = recordset to store as a cookie. Recordset must have only 1 field and it doesn't matter what it is called.
' The user saving the cookie is already known via a session variable
'
' Output:
' None
' ========================================================================================
function SaveDBCookieSet( byval strName, byval strString, byval oRsValues )
	dim oRecords, oError, sError

	if Session("AMOUserID") <> "" and Session("AMOUserID") <> 0 then
		'set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
		set oRecords = New ISGeneral
		sError = oRecords.SaveDBCookieSet (Application("Repository"), clng(Session("AMOUserID")), strName, strString, oRsValues )
		if sError <> True then
			Response.Write("Cookie not saved. Try again.")
			Response.End()
		end if
		set oRecords = nothing
	end if
end function

%>