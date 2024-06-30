<%
 Response.Write Request.ServerVariables("HTTP_USERID")
 Response.Write "<br>"
 Response.Write Request.ServerVariables("HTTP_SM_UNIVERSALID")
 Response.Write "<br>"
 Response.Write Request.ServerVariables("LOGON_USER")
 Response.Write "<br>"
 Response.Write "<br>"
 Response.Write "<br>"

	Dim AuthUser
	AuthUser = Request.ServerVariables("LOGON_USER")
	Response.Write AuthUser
	Dim UserVariables
	UserVariables = Split(AuthUser, ",")
	Response.Write UBound(UserVariables)
	
	 Response.Write "<br>"
 Response.Write "<br>"
 Response.Write "<br>"
 
	For i = 0 To UBound(UserVariables)
		If LEFT(LCASE(UserVariables(i)), 4) = "uid=" Then
			Response.Write REPLACE(LCASE(UserVariables(i)), "uid=", "")
		End If
	Next
	
	 Response.Write "<br>"
 Response.Write "<br>"
 Response.Write "<br>"
 
 for each x in Request.ServerVariables
   response.write "<p>" & x & ": " & Request.ServerVariables(x) & "</p>"
 next
%>