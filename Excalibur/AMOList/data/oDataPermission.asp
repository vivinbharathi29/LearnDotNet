<%
'*************************************************************************************
'* FileName		: oDataPermission.asp
'* Description	: Class for Permission Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************

Class ISRole

    '--DECLARE LOCAL VARIABLES------------------------------------------------------------
    Dim oErrors		                        'ERROR OBJECT
    Dim sErrorMessage                       'ERROR MESSAGE
    Dim oConnect			                'DB OBJECT
    Dim oCommand                            'COMMAND OBJECT
    Dim oRSelection, oRSelection2           'RECORDSET OBJECT
    Dim oParameter                          'COMMAND PARAMETER
    Dim sSQL                                'SQL STRING
    Dim bProcessComplete                    'STATUS 

    ' -----------------------------------------------------------------------------
    ' Function: GetAllRolesByUserGroup
    '
    ' @Purpose: This procedure gets all the permissions of the
    '           current user FROM THE DATABASE.
    ' Inputs:
    '   @parm string | p_strRepository   | database repository string
    '   @parm Long   | p_lngUserID       | UserID of person to search
    '   @parm string | p_strUserGroupIDs | Comma delimited string of UserGroupIDs to search
    '
    ' Outputs:
    '   @parm variant | p_rs | the recordset that contains all the permissions
    '                          of the current user
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetAllRolesByUserGroup(p_strRepository, _
                                p_lngUserID, _
                                p_strUserGroupIDs) 
            
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
        oCommand.CommandText = "usp_USR_GetAllRolesByUserGroup" 
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUserGroups", adVarChar, adParamInput, Len(p_strUserGroupIDs))
		oParameter.Value = p_strUserGroupIDs
		oCommand.Parameters.Append oParameter
            
         ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetAllRolesByUserGroup = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetAllRolesByUserGroup = oRSelection

            'Set Session("oConnect") = oConnect
            Exit Function

        End If

    End Function

    ' -----------------------------------------------------------------------------
    ' Function: GetAllRolesByDivision
    '
    ' @Purpose: This procedure gets all the permissions of the
    '           current user FROM THE DATABASE.
    ' Inputs:
    '   @parm string | p_strRepository  | database repository string
    '   @parm Long   | p_lngUserID      | UserID of person to search
    '   @parm string | p_strDivisionIDs | Comma delimited string of DivisionIDs to search
    '
    ' Outputs:
    '   @parm variant | p_rs | the recordset that contains all the permissions
    '                          of the current user
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetAllRolesByDivision(p_strRepository, _
                                p_lngUserID, _
                                p_strDivisionIDs) 

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
            oCommand.CommandText = "usp_USR_GetAllRolesByDivision" 
    
            ' set database stored procedure command
            Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, 4)
			oParameter.Value = p_lngUserID
			oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_chrDivisionIDs", adVarChar, adParamInput, Len(p_strDivisionIDs))
			oParameter.Value = p_strDivisionIDs
			oCommand.Parameters.Append oParameter
            
             ' execute the stored procedure
            oRSelection.CursorType = adOpenStatic
		    oRSelection.CursorLocation = adUseClient
            oRSelection.Open(oCommand)
            
           
            ' if, no error return recordset object
            If Err.Number > 0 Then		    
                'Close database connection, oConnect.
                'oRSelection.Close
                'Set oRSelection = Nothing
                
                'oConnect.close
			    'Set oConnect = Nothing
                
                'return empty recordset object and exit function
                Set GetAllRolesByDivision = Nothing
                Exit Function
            Else
                'disconnect the recordset
                Set oCommand.ActiveConnection = Nothing
                Set oCommand = Nothing
                
                'return recordset
                Set GetAllRolesByDivision = oRSelection

                'Set Session("oConnect") = oConnect
                Exit Function

            End If

    End Function

    ' -----------------------------------------------------------------------------
    ' Function: GetAllRoles
    '
    ' @Purpose: This procedure gets all the permissions of the
    '           current user FROM THE DATABASE.
    ' Inputs:
    '   @parm string | p_strRepository | database repository string
    '
    ' Outputs:
    '   @parm variant | p_rs | the recordset that contains all the permissions
    '                          of the current user
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetAllRoles(p_strRepository, p_User, p_Permission) 

        Dim strUser
        strUser = p_User

        Dim varP1, varP2, lngPos 
        If InStr(strUser, "\") > 0 Then   'domain\username
            lngPos = InStr(1, strUser, "\")
            If lngPos > 0 Then
                varP1 = Mid(strUser, lngPos + 1, Len(strUser))
                varP2 = Left(strUser, lngPos - 1)
            Else
                varP1 = strUser
                varP2 = ""
            End If
        Else  'email address
            lngPos = InStr(1, strUser, "@")
            If lngPos > 0 Then
                varP2 = Mid(strUser, lngPos + 1, Len(strUser))
                varP1 = Left(strUser, lngPos - 1)
            Else
                varP1 = strUser
                varP2 = ""
            End If
        End If

        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        'Response.Write(varP1 & "--" & varP2 & "--" & p_Permission)
        'Response.End()

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
        oCommand.CommandText = "usp_USR_ViewPermission"
        oCommand.CommandTimeout = 120
          
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrUser", adVarChar, adParamInput, Len(varP1))
		oParameter.Value = varP1
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDomain", adVarChar, adParamInput, Len(varP2))
		oParameter.Value = varP2
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_PName", adVarChar, adParamInput, Len(p_Permission))
		oParameter.Value = p_Permission
		oCommand.Parameters.Append oParameter
            
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)

        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetAllRoles = Nothing
            Exit Function
        Else
          
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetAllRoles = oRSelection

            'Set Session("oConnect") = oConnect
            Exit Function

        End If
    
    End Function

    ' -----------------------------------------------------------------------------
    ' Function: GetUserID
    '
    ' @Purpose: This procedure gets the ID of the a user in the database.
    ' Inputs:
    '   @parm string | p_strRepository | database repository string
    '   @parm string | p_varUser | optional NT user name to get ID
    '   (if this parameter is not specified, current user is assumed.)
    '
    ' Outputs:
    '   @parm variant | p_lngID | the user's ID
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetUserID(p_strRepository, p_User)
                                  
        Dim lng_UserID

        Set oConnect = Session("oConnect")
		'oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
        ' Handle unexpected errors
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect 
	    oCommand.CommandText = "usp_USR_ViewID" 
        oCommand.CommandType = adCmdStoredProc 
        oCommand.CommandTimeout = 120 
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrDomainUserName", adVarChar, adParamInput, Len(p_User))
        oParameter.Value = p_User
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intID", adInteger, adParamOutput, cg_lngLEN_INT)
        oCommand.Parameters.Append oParameter
            
       ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			GetUserID = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
				
			Exit Function
        Else 
            lng_UserID = oCommand.Parameters("@p_intID").Value
            GetUserID = lng_UserID
				
			'Open database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                
			Exit Function
        End If
    End Function

    ' -----------------------------------------------------------------------------
    ' Function: GetRoleGroups
    '
    ' @Purpose: This procedure gets the list of user groups belong to one or more roles.
    ' Inputs:
    '   @parm variant | p_strRepository | the database repository string
    '   @parm variant | p_strRoles | the comma-delimited list of roles (using
    '   role numbers) to get members
    '
    ' Outputs:
    '   @parm variant | p_rs | the recordset that contains all the user groups of a role
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetRoleGroups(p_strRepository, _
                                  p_strRoles, _
                                  p_strPermission, _
                                  p_blnCreate, _
                                  p_blnUpdate, _
                                  p_blnView, _
                                  p_blnDelete, p_intUserID) 
        
        If IsNull(p_intUserID) = True Or IsEmpty(p_intUserID) = True Or p_intUserID = "" Then
           p_intUserID = 0
        End If   

        Set GetRoleGroups = GetRoleLowerLevel(p_strRepository, p_strRoles, p_strPermission, p_blnCreate, p_blnUpdate, p_blnView, p_blnDelete, 3, p_intUserID)
    End Function

    ' -----------------------------------------------------------------------------
    ' Function: GetRoleLowerLevel
    '
    ' @Purpose: This procedure gets the list of users/NT groups/groups based on the list of role
    ' numbers.
    ' Inputs:
    '   @parm variant | p_strRepository | the database repository string
    '   @parm variant | p_strRoles | the comma-delimited list of roles (using
    '   role numbers)
    '
    ' Outputs:
    '   @parm variant | p_strRoleNames | the list of role names
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Private Function GetRoleLowerLevel(p_strRepository, _
                                  p_strRoles, _
                                  p_strPermission, _
                                  p_blnCreate, _
                                  p_blnUpdate, _
                                  p_blnView, _
                                  p_blnDelete, _
                                  p_lngType, _
                                  p_intUserID) 

        Dim strRoleNames 
        Dim strSP 

        If IsNull(p_intUserID) = True Or IsEmpty(p_intUserID) = True Or p_intUserID = "" Then
            p_intUserID = 0
        End If   
        
        'translate the list of role codes to role names
        strRoleNames = GetRoleNames(p_strRepository, p_strRoles, p_strPermission, p_blnCreate, p_blnUpdate, p_blnView, p_blnDelete)
        
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
         ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        If IsNumeric(p_lngType) Then
            Select Case p_lngType
                Case 1 'get role members
                    strSP = "usp_ROLE_ViewNTGroups"
                Case 2 'get NT groups
                    strSP = "usp_ROLE_ViewNTGroups"
                Case 3 'get groups
                    strSP = "usp_ROLE_ViewGroups"
                Case Else
            End Select
        Else
            Set GetRoleLowerLevel = Nothing
            Exit Function
        End If

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
        oCommand.CommandText = strSP
        oCommand.CommandTimeout = 120 

        
        'get the list of users in the input roles
        Set oParameter = oCommand.CreateParameter("@p_chrRoleNames", adVarChar, adParamInput, IIf(Len(strRoleNames) > 0, Len(strRoleNames), 1))
        oParameter.Value = strRoleNames
        oCommand.Parameters.Append oParameter

        If (p_lngType = 2) Or (p_lngType = 3) Then 'change to @p_Userid int
            Set oParameter = oCommand.CreateParameter("@p_Userid", adInteger, adParamInput, cg_lngLEN_INT)
            oParameter.Value = p_intUserID
            oCommand.Parameters.Append oParameter
        End If

        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetRoleLowerLevel = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetRoleLowerLevel = oRSelection

            'Set Session("oConnect") = oConnect
            Exit Function

        End If

    End Function


    ' -----------------------------------------------------------------------------
    ' Function: GetRoleNames
    '
    ' @Purpose: This procedure gets the list of role names based on the list of role
    ' codes.
    ' Inputs:
    '   @parm variant | p_strRepository | the database repository string
    '   @parm variant | p_strRoles | the comma-delimited list of roles (using
    '   role numbers)
    '
    ' Outputs:
    '   @parm variant | p_strRoleNames | the list of role names
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Private Function GetRoleNames(p_strRepository, _
                                    p_strRoles, _
                                    p_strPermission, _
                                    p_blnCreate, _
                                    p_blnUpdate, _
                                    p_blnView, _
                                    p_blnDelete) 
        Dim strRoles, p_strRoleNames

        

        strRoles = Trim(p_strRoles)

        If strRoles = "" Then
            p_strRoleNames = ""
        Else
            'initialize the list of role names
            p_strRoleNames = ""
        
            'strip off the last comma in the list of role codes
            If Right(strRoles, 1) = "," Then
                strRoles = Left(strRoles, Len(strRoles) - 1)
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
	        oCommand.CommandText = "usp_ROLE_ViewbyRoleCode"
            oCommand.CommandTimeout = 120  
           
            Set oParameter = oCommand.CreateParameter("@p_chrList", adVarChar, adParamInput, IIf(Len(strRoles) > 0, Len(strRoles), 1))
            oParameter.Value = strRoles
            oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intCreate", adInteger, adParamInput, cg_lngLEN_INT)
            oParameter.Value = IIf(p_blnCreate, 1, 0)
            oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intUpdate", adInteger, adParamInput, cg_lngLEN_INT)
            oParameter.Value = IIf(p_blnUpdate, 1, 0)
            oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intView", adInteger, adParamInput, cg_lngLEN_INT)
            oParameter.Value = IIf(p_blnView, 1, 0)
            oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intDelete", adInteger, adParamInput, cg_lngLEN_INT)
            oParameter.Value = IIf(p_blnDelete, 1, 0)
            oCommand.Parameters.Append oParameter
    
            'execute the stored procedure
            oRSelection.CursorType = adOpenStatic
		    oRSelection.CursorLocation = adUseClient
            oRSelection.Open(oCommand)

            ' if, no error return recordset object
            If Err.Number > 0 Then		    
			    'Close database connection, oConnect.
                'oRSelection.Close
                'Set oRSelection = Nothing
                
                'oConnect.close
			    'Set oConnect = Nothing
                
                'return empty recordset object and exit function
                GetRoleNames = ""
                Exit Function
            Else
                If p_strPermission = "" Then
                    If Not oRSelection.EOF Then
                        Do While Not oRSelection.EOF
                            p_strRoleNames = p_strRoleNames & oRSelection.Fields("RoleName").Value & ","
                        oRSelection.MoveNext()
                        Loop
                    End If
                Else    'only return roles assigned a specific Permission
                    If Not oRSelection.EOF Then
                        Do While Not oRSelection.EOF
                            If InStr(oRSelection.Fields("Name").Value, p_strPermission) > 0 Then    '--Check permission name for specific permission 
                                p_strRoleNames = p_strRoleNames & oRSelection.Fields("RoleName").Value & ","
                            End If
                        oRSelection.MoveNext()
                        Loop
                    End If
                End If

               'strip off the last comma in the list of role codes
                If Right(p_strRoleNames, 1) = "," Then
                    p_strRoleNames = Left(p_strRoleNames, Len(p_strRoleNames) - 1)
                End If

                'return recordset
                GetRoleNames = p_strRoleNames

                'Set Session("oConnect") = oConnect
                Exit Function

            End If

        End If
    End Function


        ' -----------------------------------------------------------------------------
    ' Function: GetRoleNames
    '
    ' @Purpose: This procedure gets the list of role names based on the list of role
    ' codes.
    ' Inputs:
    '   @parm variant | p_strRepository | the database repository string
    '   @parm variant | p_strRoles | the comma-delimited list of roles (using
    '   role numbers)
    '
    ' Outputs:
    '   @parm variant | p_strRoleNames | the list of role names
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function GetUserRoleNames(p_strRepository, p_intUserID) 
            
        Dim strRoles, p_strRoleNames

        'initialize the list of role names
        p_strRoleNames = ""
        
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
	    oCommand.CommandText = "usp_USR_GetRoles"
        oCommand.CommandTimeout = 120  
           

        Set oParameter = oCommand.CreateParameter("@p_intUserId", adInteger, adParamInput, cg_lngLEN_INT)
        oParameter.Value = p_intUserID
        oCommand.Parameters.Append oParameter

    
        'execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)

        ' if, no error return recordset object
        If Err.Number > 0 Then		    
			'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            GetUserRoleNames = ""
            Exit Function
        Else
            If Not oRSelection.EOF Then
                Do While Not oRSelection.EOF
                    p_strRoleNames = p_strRoleNames & oRSelection.Fields("RoleName").Value & ", "
                oRSelection.MoveNext()
                Loop
            End If

           'strip off the last comma in the list of role codes
            If Right(p_strRoleNames, 1) = "," Then
                p_strRoleNames = Left(p_strRoleNames, Len(p_strRoleNames) - 1)
            End If
                
            'return recordset
            GetUserRoleNames = p_strRoleNames

            'Set Session("oConnect") = oConnect
            Exit Function

        End If

    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: GetActualUserInfo
    '* Inputs		: NTName
    '-------------------------------------------------------------------------------------
    Public Function GetActualUserInfo(p_strRepository, p_LoggedInUser)
        Dim p_CurrentDomain, p_CurrentUser

        'Get Actual Current User's Netowrk Domain and Name:
        if instr(p_LoggedInUser,"\") > 0 then
		    p_CurrentDomain = left(p_LoggedInUser, instr(p_LoggedInUser,"\") - 1)
		    p_CurrentUser = mid(p_LoggedInUser,instr(p_LoggedInUser,"\") + 1)
	    end if
			
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
	    oCommand.CommandText = "spGetUserInfo"

        Set oParameter = oCommand.CreateParameter("@UserName", adVarChar, adParamInput, 80)
        oParameter.Value = p_CurrentUser
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@Domain", adVarChar, adParamInput, 30)
        oParameter.Value = p_CurrentDomain
        oCommand.Parameters.Append oParameter
                
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
           'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetActualUserInfo = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetActualUserInfo = oRSelection
            
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
            
    End Function

        '-------------------------------------------------------------------------------------
    '* Purpose		: GetImpersonateUserID
    '* Inputs		: NTName
    '-------------------------------------------------------------------------------------
    Public Function GetImpersonateUser(p_strRepository, p_LoggedInUser)
        Dim p_CurrentDomain, p_CurrentUser

        'Get Actual Current User's Netowrk Domain and Name:
        if instr(p_LoggedInUser,"\") > 0 then
		    p_CurrentDomain = left(p_LoggedInUser, instr(p_LoggedInUser,"\") - 1)
		    p_CurrentUser = mid(p_LoggedInUser,instr(p_LoggedInUser,"\") + 1)
	    end if
			
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
	    oCommand.CommandText = "spGetEmployeeImpersonateID"

        Set oParameter = oCommand.CreateParameter("@NTName", adVarChar, adParamInput, 50)
        oParameter.Value = p_CurrentUser
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@Domain", adVarChar, adParamInput, 50)
        oParameter.Value = p_CurrentDomain
        oCommand.Parameters.Append oParameter
                
        ' execute the stored procedure
        oRSelection.CursorType = adOpenStatic
		oRSelection.CursorLocation = adUseClient
        oRSelection.Open(oCommand)
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
           'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set GetImpersonateUser = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetImpersonateUser = oRSelection
            
            'Set Session("oConnect") = oConnect
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