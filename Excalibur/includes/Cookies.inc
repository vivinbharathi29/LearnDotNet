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
function GetDBCookie( byval strName, byval intUserID )	
	Dim dw, cn, cmd, rs, strCookie
	strCookie = ""
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_COOKIE_GetDBCookie")
	dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 4, intUserID
	dw.CreateParameter cmd, "@p_chrName", adVarChar, adParamInput, 128, strName
	Set rs = dw.ExecuteCommandReturnRS(cmd)	

	If Not (rs.EOF And rs.BOF) Then
		strCookie = rs("Value")
	End If
	rs.Close

	Set rs = nothing
	Set cmd = nothing
	Set cn = nothing
	Set dw = nothing

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
function SaveDBCookie( byval strName, byval strValue, byval intUserID )
	Dim dw, cn, cmd, rs, strCookie
	ON ERROR RESUME NEXT	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_COOKIE_SaveDBCookie")
	dw.CreateParameter cmd, "@p_intUserID", adInteger, adParamInput, 4, intUserID
	dw.CreateParameter cmd, "@p_chrName", adVarChar, adParamInput, 128, strName
	dw.CreateParameter cmd, "@p_chrValue", adVarChar, adParamInput, 255, strValue
	dw. ExecuteNonQuery(cmd)		
	if Err.number <> 0 then
		Response.Write("Unable to complete your SaveDBCookie request. Tried to Save cookie: " & strName & " with value: " & strValue)
		Response.End()
	end if
end function

%>