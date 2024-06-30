<%
'*************************************************************************************
'* FileName		: oDataGeneral.asp
'* Description	: Class for Cookie Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************

Class ISGeneral

    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 


    '-----------------------------------------------------------------
    'Procedure: GetDBCookie
    '@Purpose:  Get a database cookie
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_intUserID            | id of user
    '           @parm | String  | p_strName              | cookie name desired
    'Outputs:
    '           @parm | Variant | p_oRs                  | recordset with value and times
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function GetDBCookie(p_strRepository, _
                            p_intUserID, _
                            p_strName)

        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
        oCommand.CommandText = "usp_COOKIE_GetDBCookie" 
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrName", adVarChar, adParamInput, 128)
		oParameter.Value = p_strName
		oCommand.Parameters.Append oParameter
            
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)

            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetDBCookie = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetDBCookie = oRSelection
            Exit Function

        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveDBCookie
    '@Purpose:  Save a database cookie
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_intUserID            | id of user
    '           @parm | String  | p_strName              | cookie name to save
    '           @parm | String  | p_strValue             | cookie value to save
    'Outputs:
    '           None
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveDBCookie(p_strRepository, _
                            p_intUserID, _
                            p_strName, _
                            p_strValue)

        'set optional default value
        If IsNull(p_strValue) = True Or IsEmpty(p_strValue) = True Or p_strValue = "" Then
            p_strValue = ""
        End If

            
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
	    oCommand.CommandText = "usp_COOKIE_SaveDBCookie"

        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrName", adVarChar, adParamInput, 128)
		oParameter.Value = p_strName
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrValue", adVarChar, adParamInput, 255)
		oParameter.Value = p_strValue
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			SaveDBCookie = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
				
			Exit Function
        Else 
            SaveDBCookie = True
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
                
				
			Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: GetDBCookieSet
    '@Purpose:  Get a database cookie recordset
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_intUserID            | id of user
    '           @parm | String  | p_strName              | cookie name desired
    'Outputs:
    '           @parm | Variant | p_oRs                  | recordset with value and times
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function GetDBCookieSet(p_strRepository, _
                            p_intUserID, _
                            p_strName)

        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
        oCommand.CommandText = "usp_COOKIE_GetDBCookieSet" 
                
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrName", adVarChar, adParamInput, 128)
		oParameter.Value = p_strName
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetDBCookieSet = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing

            'return recordset
            Set GetDBCookieSet = oRSelection
            Exit Function

        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveDBCookieSet
    '@Purpose:  Save a database cookie recordset
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_intUserID            | id of user
    '           @parm | String  | p_strName              | cookie name to save
    '           @parm | String  | p_strString            | comma delimited string of values
    '           @parm | String  | p_oRsValues            | recordset of cookie values to save (this or p_strString)
    'Outputs:
    '           None
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveDBCookieSet(p_strRepository , _
                            p_intUserID, _
                            p_strName, _
                            p_strString, _
                            p_oRsValues)
            
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
	    oCommand.CommandText = "usp_COOKIE_SaveDBCookieSet" 
	       

        Dim strDSN, strXML
        Dim arrString
        Dim i 
    
        strXML = "<?xml version='1.0'  encoding='iso-8859-1' ?>"
        strXML = strXML & "<FormInfo>"
    
        'use the recordset first if present
        If p_strString = "" Then
            While Not p_oRsValues.EOF
                'just take the first field, whatever it is
                strXML = strXML & "<Cookie Value=""" & RTrim(CStr(p_oRsValues(0).Value)) & """/>"
                p_oRsValues.MoveNext
            Wend
        Else
            arrString = Split(p_strString, ",")
            For i = 0 To UBound(arrString)
                strXML = strXML & "<Cookie Value=""" & RTrim(arrString(i)) & """/>"
            Next
        End If
    
        strXML = strXML & "</FormInfo>"
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrName", adVarChar, adParamInput, 128)
		oParameter.Value = p_strName
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_txtXML", adLongVarChar, adParamInput, Len(strXML) + 1)
		oParameter.Value = strXML
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			SaveDBCookieSet = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
				
			Exit Function
        Else 
            SaveDBCookieSet = True
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
                
				
			Exit Function
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: GetOneUserInfo
    '@Purpose:  Gets information for given User IDs
    '
    'Inputs:    @parm | String        | p_strRepository    | database connection string
    '           @parm | Integer       | p_lngUserID        | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs              | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function GetOneUserInfo(p_strRepository, p_lngUserID) 
        
    
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
       ' oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
        oCommand.CommandText = "usp_USR_GetOneUserInfo" 
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetOneUserInfo = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetOneUserInfo = oRSelection
            Exit Function
        End If
    
    End Function


    ' -----------------------------------------------------------------------------
    ' Function: GetGroupsByUser
    '
    ' @Purpose: This procedure gets all the user groups that an user belongs the system.
    ' Inputs:
    '   @parm string | p_strRepository | database repository string
    '   @parm string | p_lngID | UserID
    '
    ' Outputs:
    '   @parm variant | p_rs | the recordset containing the groups
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetGroupsByUser(p_strRepository, p_lngID)
      Set GetGroupsByUser = ViewList2(p_strRepository, p_lngID, "@p_intID", adInteger, "usp_USR_ViewGroupsByUser")
    End Function


    '-------------------------------------------------------------------------------------
    '* Purpose		: ViewList
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function ViewList2(p_strRepository, _
                                p_lngID, _
                                p_param, _
                                p_datatype, _
                                p_strSP)

        'Set optional default value
        If IsNumeric(p_lngID) = True Then
			p_lngID = CLng(p_lngID)
		End If
			
        Set oConnect = Session("oConnect") 
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = p_strSP

        ' set stored procedure parameter(s)
        If IsNumeric(p_lngID) = True Then
            Set oParameter = oCommand.CreateParameter(p_param, p_datatype, adParamInput, Len(p_lngID))
			oParameter.Value = p_lngID
			oCommand.Parameters.Append oParameter
        End If
                
         ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set ViewList2 = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewList2 = oRSelection

            Exit Function
        End If
            
    End Function

    '-----------------------------------------------------------------
    'Procedure: ViewUsers
    '@Purpose:  Display the users for user groups
    'Inputs:    @parm                 | String | p_strRepository  | database connection string
    '           @parm                 | String | p_intGroupID     | user's group ID to filter
    '           @parm                 | String | p_intUserID      | user's ID to filter
    '           @parm                 | String | p_intNTGroupID   | NT Group's ID to filter
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs  | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function ViewUsers(p_strRepository, _
                            p_intGroupID, _
                            p_intUserID, _
                            p_intNTGroupID) 
        
        'Set optional default value
        If IsNull(p_intGroupID) = True Or IsEmpty(p_intGroupID) = True Or p_intGroupID = ""  Then
			p_intGroupID = 0
		End If

        If IsNull(p_intUserID) = True Or IsEmpty(p_intUserID) = True Or p_intUserID = ""  Then
			p_intUserID = 0
		End If

        If IsNull(p_intNTGroupID) = True Or IsEmpty(p_intNTGroupID) = True Or p_intNTGroupID = ""  Then
			p_intNTGroupID = 0
		End If

      
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc	
        oCommand.CommandText = "usp_USR_ViewUsers" 
    
    
        ' set the parameters
         ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intGroupID", adInteger, adParamInput, 4)
		oParameter.Value = p_intGroupID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intNTGroupID", adInteger, adParamInput, 4)
		oParameter.Value = p_intNTGroupID
		oCommand.Parameters.Append oParameter

   
       ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)

            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set ViewUsers = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewUsers = oRSelection
            Exit Function
        End If
    
    End Function


    '-------------------------------------------------------------------------------------
    '* Procedure	: IIF    
    '* Purpose		: Custom IIF function for single line If..Else..Then statements in VBScript
    '* Inputs		: Condition, TrueValue, FalseValue
    '-------------------------------------------------------------------------------------
    Private Function IIf(Condition, TrueValue, FalseValue)
    Dim bCondition
        bCondition = False 
        On Error Resume Next
        bCondition = CBool(Condition)
        On Error Goto 0
        If bCondition Then 
            If IsObject(TrueValue) Then 
                Set IIf = TrueValue
            Else 
                IIf = TrueValue
            End If 
        Else
            If IsObject(FalseValue) Then 
                Set IIf = FalseValue
            Else 
                IIf = FalseValue
            End If 
        End If 
    End Function 
		
	 '*************************************************************************************
	'* Purpose		: Build database connection string.
	'* Inputs		: None
	'* Returns		: PULSARDB - database connection string.
	'*************************************************************************************
	Private Function PULSARDB()
        Dim oSvr
        '---Close DB Connection: ---
        Set oSvr = New DBConnection 
        PULSARDB = oSvr.PulsarConnectionString()
	End Function


End Class
%>