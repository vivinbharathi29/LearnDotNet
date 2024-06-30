<%

Class DataWrapper

	Private m_ActiveConnection
	Private m_cmd
	
	Public Function CreateConnection(ConnectionString)
		Dim cn
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session(ConnectionString) 
		cn.Open
		'set m_ActiveConnection = cn	
		set CreateConnection = cn
		set cn = nothing
	End Function

	Public Function CreateCommandSP(ConnectionObject, ProcedureName)
		Dim cm
		set cm = server.CreateObject("ADODB.Command")
		cm.ActiveConnection = ConnectionObject
		cm.CommandType = 4
		cm.CommandText = ProcedureName
		cm.CommandTimeout = 120
		'cm.NamedParameters = True
		
		set CreateCommandSP = cm
		
	End Function
	
		Public Function CreateCommandSPwTimeout(ConnectionObject, ProcedureName, TimeOut)
		Dim cm
		set cm = server.CreateObject("ADODB.Command")
		cm.ActiveConnection = ConnectionObject
		cm.CommandType = 4
		cm.CommandText = ProcedureName
		cm.CommandTimeout = TimeOut
		'cm.NamedParameters = True
		
		set CreateCommandSPwTimeout = cm
		
	End Function

	Public Function CreateCommandSQL(ConnectionObject, SQLText)
	    Dim cm
	    Set cm = Server.CreateObject("ADODB.Command")
	    cm.ActiveConnection = ConnectionObject
	    cm.CommandType = 1
	    cm.CommandText = SQLText
	    cm.CommandTimeout = 120
	    
	    Set CreateCommandSQL = cm
	End Function

	Public Function CreateParameter(ByRef CommandObject, ParamName, ParamType, ParamDirection, ParamSize, ParamValue)
    ON ERROR RESUME NEXT
		Dim p
		set p = CommandObject.CreateParameter(ParamName, ParamType, ParamDirection, CLng(ParamSize))
		Select Case ParamType
			Case adBoolean
				If IsNull(ParamValue) Then
				    p.Value = NULL
				Else
				    Select Case LCase(CStr(ParamValue))
					    Case "1", "on", "checked", True
						    p.Value = True
					    Case Else
						    p.Value = False
				    End Select
				End If
			Case Else
				p.Value = IsDbNull(ParamValue)
		End Select
		
		
		CommandObject.Parameters.Append p

		If Err.number <> 0 Then
            Dim errDescription
		    Response.Write "Error Description: " & err.Description & "<br>"
		    Response.Write "Error Source:" & err.Source & "<br>"
		    Response.Write "Error Number: " & err.number & "<br>"
   		    errDescription = err.Description
            ON ERROR GOTO 0
		    Err.Raise 500,"DataWrapper.CreateParameter",ParamName & " - " & errDescription
		End If
    ON ERROR GOTO 0		
	End Function

	Public Function ExecuteCommandReturnRS(ByRef CommandObject)
        ON ERROR RESUME NEXT
        
		Dim rs
		Set rs = server.CreateObject("ADODB.recordset")
		rs.CursorType = adOpenStatic
		rs.CursorLocation = adUseClient

		rs.Open(CommandObject)

		If Err.number <> 0 Then
		    Dim errDescription
		    Response.Write "Error Description: " & err.Description & "<br>"
		    Response.Write "Error Source:" & err.Source & "<br>"
		    Response.Write "Error Number: " & err.number & "<br>"
		    errDescription = err.Description
            ON ERROR GOTO 0
		    Err.Raise -500,"DataWrapper.ExecuteCommandReturnRS", CommandObject.CommandText & " Error.Description:" & errDescription
		End If

        ON ERROR GOTO 0
		
		Set ExecuteCommandReturnRS = rs
		
	End Function
	
    Public Function ExecuteCommandNonQuery(ByRef CommandObject)
        ExecuteCommandNonQuery = ExecuteNonQuery(CommandObject)
    End Function

	Public Function ExecuteNonQuery(ByRef CommandObject)
		Dim iRecordCount
		iRecordCount = 0
        ON ERROR RESUME NEXT
		
		CommandObject.Execute iRecordCount
		
		If Err.number <> 0 Then
		    Dim errDescription
		    Response.Write "Error Description: " & err.Description & "<br>"
		    Response.Write "Error Source:" & err.Source & "<br>"
		    Response.Write "Error Number: " & err.number & "<br>"
		    iRecordCount = err.number
		    errDescription = err.Description
            ON ERROR GOTO 0
		    Err.Raise -500,"DataWrapper.ExecuteNonQuery", CommandObject.CommandText & " Error.Description:" & errDescription
		End If

        ON ERROR GOTO 0

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