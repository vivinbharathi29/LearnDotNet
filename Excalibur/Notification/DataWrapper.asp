<%

Class DataWrapper

	Private m_ActiveConnection
	Private m_cmd
	
	Public Function CreateConnection(ConnectionString)
		Dim cn
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Application(ConnectionString) 
		cn.Open
		set m_ActiveConnection = cn	
		set CreateConnection = cn
		
	End Function

	Public Function CreateCommandSP(ConnectionObject, ProcedureName)
		Dim cm
		set cm = server.CreateObject("ADODB.Command")
		cm.ActiveConnection = ConnectionObject
		cm.CommandType = 4
		cm.CommandText = ProcedureName
		
		set CreateCommandSP = cm
		
	End Function

	Public Function CreateCommandSQL(ConnectionObject, SQLText)
	End Function

	Public Function CreateParameter(ByRef CommandObject, ParamName, ParamType, ParamDirection, ParamSize, ParamValue)

		Dim p
		set p = CommandObject.CreateParameter(ParamName, ParamType, ParamDirection, ParamSize)
		p.Value = IsDbNull(ParamValue)
		CommandObject.Parameters.Append p
		
	End Function

	Public Function ExecuteCommandReturnRS(ByRef CommandObject)
		Dim rs
		'Set rs = server.CreateObject("ADODB.recordset")
		'rs.CursorType = adOpenForwardOnly
		'rs.LockType = AdLockReadOnly

		Set rs = CommandObject.Execute 
		
		set ExecuteCommandReturnRS = rs
		
	End Function
	
	Public Function ExecuteNonQuery(ByRef CommandObject)
		Dim iRecordCount
		
		CommandObject.Execute iRecordCount
		
		ExecuteNonQuery = iRecordCount
		
	End Function

	Public Function UserIsAdmin(ConnectionObject, UserID)
		dim cmd
		set cmd = CreateCommandSP(ConnectionObject, "usp_SelectEmployees")
		CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 4, UserID 
		CreateParameter cmd, "@p_IsAdmin", adBoolean, adParamInput, 1, 1 
		CreateParameter cmd, "@p_NTName", adVarChar, adParamInput, 30, ""
		CreateParameter cmd, "@p_Domain", adVarChar, adParamInput, 30, ""
		CreateParameter cmd, "@p_PartnerID", adInteger, adParamInput, 4, ""
		
		Dim rs
		Set rs = ExecuteCommandReturnRS(cmd)
		
		If rs.EOF and rs.BOF then
			UserIsAdmin = false
		else
			UserIsAdmin = true
		end if
		
		rs.Close
		set rs = nothing
		set cmd = nothing
		
	End Function
	
	Public Function IsDbNull(InputValue)
		
		Dim ReturnValue
		
		ReturnValue = InputValue
		
		If Len(Trim(InputValue)) = 0 Then
			ReturnValue = NULL
		End If
		
		IsDbNull = ReturnValue
		
	End Function

End Class 'DataWapper
%>