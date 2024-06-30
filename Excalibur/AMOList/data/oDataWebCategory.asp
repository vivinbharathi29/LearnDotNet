<%
'*************************************************************************************
'* FileName		: oDataWebCategory.asp
'* Description	: Class for Web Category Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO 
'*************************************************************************************

Class ISWebCategory

    '--DECLARE LOCAL VARIABLES------------------------------------------------------------
    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 
            

    '-----------------------------------------------------------------
    'Function: Dependent functions that return Recordsets
    '
    '-----------------------------------------------------------------
    
    Public Function wUser_BusinessSegment(p_strRepository, p_lngUserID) 
        Set wUser_BusinessSegment = GetCategory(p_strRepository, "usp_USR_GetBusinessSegment", p_lngUserID, null)
    End Function
    
    
    'NOT NEEDED FOR AMO FEATURES - "Owned By" not used to manage or create AMO Features (see Efren): ---
    'Public Function AMOOwner(p_strRepository)
        'Set AMOOwner = GetCategory(p_strRepository, "IRS_usp_AMO_OptionGroupOwner", null, null)
    'End Function

    'NOT NEEDED FOR AMO FEATURES - Used for AMO_Properties.asp which is replaced by IPulsar's AMO Feature Properties: ---      
    'Public Function wPlatform_All(p_strRepository) 
        'Set wPlatform_All = GetCategory(p_strRepository, "IRS_usp_PLF_ViewAll", null, null)
    'End Function

    'NOT NEEDED FOR AMO FEATURES - Used for AMO_Properties and AMO_Save pages which are replaced by IPulsar's AMO Feature Properties: ---
    'Public Function Active_wTprocCategory_AllConfig(p_strRepository, p_strTCFgID)
        'Set Active_wTprocCategory_AllConfig = GetCategoryAllConfig(p_strRepository, "IRS_usp_CAT_TProcCat_ViewAllStateActive", p_strTCFgID)
    'End Function

    'NOT NEEDED FOR AMO FEATURES - Used for User.inc permission to create sessions that aren't needed for AMO List OR AMO Feature Properties
    'Public Function ViewOneUserODMs2(p_strRepository, p_lngUserID) 
        'Set ViewOneUserODMs2 = GetCategory(p_strRepository, "usp_USR_ViewOneUserODMs2", p_lngUserID, null)
    'End Function
            

    '-------------------------------------------------------------------------------------
    '* Purpose		: GetCategory
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function GetCategory(p_strRepository, _
                                p_strSP, _
                                p_lngCategoryID, _
                                p_lngTypeID)

        'Set optional default value
        If IsNumeric(p_lngCategoryID) = True Then
			p_lngCategoryID = CLng(p_lngCategoryID)
		End If
			
		If IsNumeric(p_lngTypeID) = True Then
			p_lngTypeID = CLng(p_lngTypeID)
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
        If IsNumeric(p_lngCategoryID) = True Then
            Set oParameter = oCommand.CreateParameter("VersionID", adInteger, adParamInput, Len(p_lngCategoryID))
			oParameter.Value = p_lngCategoryID
			oCommand.Parameters.Append oParameter
        End If

        If IsNumeric(p_lngTypeID) = True Then
            Set oParameter = oCommand.CreateParameter("@p_intTypeID", adInteger, adParamInput, Len(p_lngTypeID))
			oParameter.Value = p_lngTypeID
			oCommand.Parameters.Append oParameter
        End If
            
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
            Set GetCategory = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetCategory = oRSelection
            Exit Function

        End If
            
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: GetCategoryAllConfig
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function GetCategoryAllConfig(p_strRepository, _
                                p_strSP, _
                                p_strTCFgIDs)

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
        Set oParameter = oCommand.CreateParameter("@p_strTCFgIDs", adVarChar, adParamInput, Len(p_strTCFgIDs) + 1)
		oParameter.Value = p_strTCFgIDs
		oCommand.Parameters.Append oParameter
                           
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
            Set GetCategoryAllConfig = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set GetCategoryAllConfig = oRSelection
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
