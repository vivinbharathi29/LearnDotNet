<%
Class UserInfo

	Private m_CurrentUser
	Private m_CurrentUserFullName
	Private m_CurrentUserID
	Private m_CurrentUserDomain
	Private m_CurrentUserEmail
	
	Public Property Get CurrentUser()
		CurrentUser = m_CurrentUser
	End Property
	
	Public Property Get CurrentUserFullName()
		CurrentUserFullName = m_CurrentUserFullName
	End Property
	
	Public Property Get CurrentUserID()
		CurrentUserID = m_CurrentUserID
	End Property
	
	Public Property Let CurrentUserID( value )
		m_CurrentUserID = value
	End Property
	
	Public Property Get CurrentUserDomain()
		CurrentUserDomain = m_CurrentUserDomain
	End Property
	
	Public Property Get CurrentUserEmail()
		CurrentUserEmail = m_CurrentUserEmail
	End Property
	
	Private Sub Class_Initialize()		
		m_CurrentUser = lcase(Session("LoggedInUser"))

		if InStr(m_CurrentUser,"\") > 0 then
			m_CurrentUserDomain = Left(m_CurrentUser, instr(m_CurrentUser,"\") - 1)
			m_CurrentUser = mid(m_CurrentUser,instr(m_CurrentUser,"\") + 1)
		end if

		Dim cn, rs, cm, p, cnString
		cnString =Session("PDPIMS_ConnectionString")
		
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = cnString
		cn.Open

		set rs = server.CreateObject("ADODB.recordset")
		rs.ActiveConnection = cn
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetUserInfo"	
		
		Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		p.Value = m_CurrentUser
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		p.Value = m_CurrentUserDomain
		cm.Parameters.Append p
		
		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set	rs = cm.Execute 
		
		set cm=nothing	
			
		if not (rs.EOF and rs.BOF) then
			m_CurrentUserID = rs("ID") & ""
			m_CurrentUserEmail = rs("email") & ""
			m_CurrentUserFullName = rs("name") & ""
			
		end if
		rs.Close

		
		Set rs = nothing
		Set cm = nothing
		Set cn = nothing
	End Sub
		

End Class
%>