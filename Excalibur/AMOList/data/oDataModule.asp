<%
'*************************************************************************************
'* FileName		: oDataModule.asp
'* Description	: Class for Module Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'* FILE NOT NEEDED - FUNCTIONS WAS USED FOR AMO_SAVE.asp BUT NOT NEEDED TO CREATE PULSAR AMO FEATURES
'*************************************************************************************

Class ISMODULE

    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 

    '-----------------------------------------------------------------
    'Procedure: Module_MOLWhereUsed
    '
    '-----------------------------------------------------------------
    Public Function Module_MOLWhereUsed(p_strRepository, p_lngID) 
        Set Module_MOLWhereUsed = ViewList(p_strRepository, p_lngID, "IRS_usp_RPT_ModuleWhereUsed")
    End Function

    '-----------------------------------------------------------------
    'Procedure: Module_GetPORSEPM
    '
    '-----------------------------------------------------------------
    Public Function Module_GetPORSEPM(p_strRepository, p_lngModuleID)
        Set Module_GetPORSEPM = ViewList(p_strRepository, p_lngModuleID, "IRS_usp_MD_GetPORSEPM")
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: ViewList
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function ViewList(p_strRepository, _
                                p_lngID, _
                                p_strSP)

        'Set optional default value
        If IsNumeric(p_lngID) = True Then
			p_lngID = CLng(p_lngID)
		End If
			
        Set oConnect  = Session("oConnect")
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
            Set oParameter = oCommand.CreateParameter("@p_intID", adInteger, adParamInput, Len(p_lngID))
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
            Set ViewList = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewList = oRSelection

            Exit Function
        End If
            
    End Function


    '-----------------------------------------------------------------
    'Procedure: Module_Remove
    '
    '-----------------------------------------------------------------
    Public Function Module_Remove( _
                    p_strRepository, _
                    p_strListIDs, _
                    p_rsRemains)

        Set oConnect  = Session("oConnect")
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
	    oCommand.CommandText = "IRS_usp_MD_Remove"    

        ' set stored procedure parameter(s)
        Set oParameter = oCommand.CreateParameter("@p_chrIDs", adVarChar, adParamInput, Len(p_strListIDs))
	    oParameter.Value = p_strListIDs
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
            Set Module_Remove = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set Module_Remove = oRSelection
            Exit Function
        End If
    End Function


    Public Function Module_Search(p_strRepository, p_strFilter) 
              
        Set oConnect  = Session("oConnect")
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
        oCommand.CommandText = "IRS_usp_MD_Search" 
    
    
        ' set the parameters
        Set oParameter = oCommand.CreateParameter("@p_strFilter", adVarChar, adParamInput, 2048)
	    oParameter.Value = p_strFilter
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
            Set Module_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set Module_Search = oRSelection
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
