<%
'*************************************************************************************
'* FileName		: oDataNotifciation.asp
'* Description	: Class for Notification Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'* Note         : Owners not Used for AMO Features, not notification is needed for AMO
'*************************************************************************************

Class ISNotifciation

    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 


    ' -----------------------------------------------------------------------------
    ' Function: ViewUsersByEventRs
    '
    ' @Purpose: Get Users for a mailing list by giving the event name but return recordset
    ' Inputs:   @parm                 | String | p_strRepository    | database connection string
    '           @parm                 | String | p_strEvent         | event for mailing list
    '           @parm  OPTIONAL       | Long   | p_lngDivisionID    | division id to filter users
    '           @parm  OPTIONAL       | Long   | p_lngObjectID      | object id to filter users
    ' Outputs:
    '           @parm                 | String | p_strToUserIDs     | semi-colon delimited list of user ids
    '           @parm                 | String | p_strToUserEmails  | semi-colon delimited list of user email addresses
    '
    ' @Returns: nothing on success, an error object otherwise.
    ' -----------------------------------------------------------------------------
    Public Function ViewUsersByEventRs(p_strRepository, _
                            p_strEvent, _
                            p_lngDivisionID, _
                            p_lngObjectID) 

        'Set optional default value
        If IsNull(p_lngDivisionID) = True Or IsEmpty(p_lngDivisionID) = True Or p_lngDivisionID = "" Then
			p_lngDivisionID = -1
		End If
			
		If IsNull(p_lngObjectID) = True Or IsEmpty(p_lngObjectID) = True Or p_lngObjectID = "" Then
			p_lngObjectID = -1
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
	    oCommand.CommandText = "IRS_usp_MAIL_ViewUsersByEvent"

        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrEvent", adVarChar, adParamInput, 64)
		oParameter.Value = p_strEvent
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intDivisionID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngDivisionID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intObjectID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngObjectID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            'oRSelection.Close
            'Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set ViewUsersByEventRs = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewUsersByEventRs = oRSelection
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
