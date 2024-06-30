<%
'*************************************************************************************
'* FileName		: oDataAVL.asp
'* Description	: Class for PAL AVL Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************

Class ISAVL

    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 


    '-----------------------------------------------------------------
    'Procedure: AVL_ProductLine
    '
    '-----------------------------------------------------------------
    Public Function AVL_ProductLine(p_strRepository)

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
        oCommand.CommandText = "usp_AMOFeature_AVL_ProductLine"    
            
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AVL_ProductLine = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AVL_ProductLine = oRSelection
            Exit Function
        End If
    End Function

    Public Function AVL_SupplierCodeGBU(p_strRepository)

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
        oCommand.CommandText = "usp_AMOFeature_AVL_SupplierCodeGBU"    
            
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
           'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AVL_SupplierCodeGBU = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AVL_SupplierCodeGBU = oRSelection
            Exit Function
        End If
    End Function

    Public Function AVL_InitialSupplierCode(p_strRepository)

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
        oCommand.CommandText = "usp_AMOFeature_AVL_InitialSupplierCode"    
            
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' if, no error return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AVL_InitialSupplierCode = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AVL_InitialSupplierCode = oRSelection
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
