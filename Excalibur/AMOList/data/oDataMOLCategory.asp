<%
'*************************************************************************************
'* FileName		: oDataMOLCategory.asp
'* Description	: Class for MOL Category Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/29/2016 - PBI 17487/ Task 21005 - Only includes functions used in AMO
'*************************************************************************************

Class ISMOLCategory

    '--DECLARE LOCAL VARIABLES------------------------------------------------------------
    Dim oErrors		        'ERROR OBJECT
    Dim sErrorMessage       'ERROR MESSAGE
    Dim oConnect			'DB OBJECT
    Dim oCommand            'COMMAND OBJECT
    Dim oRSelection         'RECORDSET OBJECT
    Dim oParameter          'COMMAND PARAMETER
    Dim sSQL                'SQL STRING
    Dim bProcessComplete    'STATUS 
    Dim p_intCategoryType   'CATEGORY TYPE


    '-----------------------------------------------------------------
    'Function: Dependent functions that return Recordsets
    '
    '-----------------------------------------------------------------
    Public Function AMOProductLine_ViewAll(p_strRepository)
        Set AMOProductLine_ViewAll = Category_View(p_strRepository, tcAMOProductLine)
    End Function
        
    Public Function AMODivision_ViewAll(p_strRepository)
        Set AMODivision_ViewAll = Category_View(p_strRepository, tcAMOModuleDivision)
    End Function

    Public Function AMOBusinessSegment_ViewAll(p_strRepository)
        Set AMOBusinessSegment_ViewAll = List_View(p_strRepository, "usp_AMOFeature_GetBusinessSegments", 0, "@p_intFeatureID", adInteger, 4)
    End Function

    Public Function ModuleType_ViewAll(p_strRepository)
        Set ModuleType_ViewAll = Category_View(p_strRepository, tcModuleType)
    End Function

    'No longer Used - Used for AMO_Properties and AMO_Save pages which aren't needed for AMO Features
    'Public Function ModuleDivision_ViewAll(p_strRepository)
        'Set ModuleDivision_ViewAll = Category_View(p_strRepository, tcMOLModuleDivision)
    'End Function


    Public Function ModuleCategory_ViewAll(p_strRepository, p_intHWSW) 
        'Set optional default value
        If IsNumeric(p_intHWSW) = True Then
			p_intHWSW = CLng(p_intHWSW)
        Else
            p_intHWSW = 2
		End If

        If p_intHWSW = 0 Then
            Set ModuleCategory_ViewAll = Category_View(p_strRepository, tcHWModuleCategory)
        ElseIf p_intHWSW = 1 Then
            Set ModuleCategory_ViewAll = Category_View(p_strRepository, tcSWModuleCategory)
        ElseIf p_intHWSW = 2 Then
            Set ModuleCategory_ViewAll = Category_View(p_strRepository, tcHWSWModuleCategory)
        ElseIf p_intHWSW = 3 Then
            Set ModuleCategory_ViewAll = Category_View(p_strRepository, tcHWSWModuleCategory_NoAll)
        End If
    End Function

    Public Function AMOStatus_ViewAll(p_strRepository)
        'Replaces sp_Status_View 
        Set AMOStatus_ViewAll = List_View(p_strRepository, "usp_AMOFeature_FeatureStatus", tsAMOStatus, "@p_intStatusType", adInteger, cg_lngLEN_INT) 
    End Function

     Public Function FeatureCategory_ViewAll(p_strRepository) 
        Set FeatureCategory_ViewAll = List_View(p_strRepository, "usp_ADMIN_GetAllFeatureCategory", null, null, null, null)
    End Function


    '-------------------------------------------------------------------------------------
    '* Purpose		: List_View
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function List_View(p_strRepository, p_strSP, p_varP1, p_strP1Desc, p_lngP1Type, p_lngP1Len)

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
        oCommand.CommandTimeout = 120  
	    

        ' set stored procedure parameter(s)
        If IsNull(p_varP1) <> True And IsEmpty(p_varP1) <> True And p_varP1 <> "" Then
            Set oParameter = oCommand.CreateParameter(p_strP1Desc, p_lngP1Type, adParamInput, p_lngP1Len)
            oParameter.Value = p_varP1
            oCommand.Parameters.Append oParameter
        End If
            
        'execute the stored procedure
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
            Set List_View = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set List_View = oRSelection
            Exit Function
        End If
    End Function   
    
    'Set AMODivision_ViewAll = Category_View(p_strRepository, tcAMOModuleDivision, p_rs)
    '-------------------------------------------------------------------------------------
    '* Purpose		: Category_View
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function Category_View(p_strRepository, p_lngType)

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
	    oCommand.CommandText = "usp_AMOFeature_CT_ViewAll" 
        oCommand.CommandTimeout = 120  

        ' set stored procedure parameter(s)
        Set oParameter = oCommand.CreateParameter("@p_intCategoryType", adInteger, adParamInput, 4)
        oParameter.Value = p_lngType
        oCommand.Parameters.Append oParameter
            
        'execute the stored procedure
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
            Set Category_View = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set Category_View = oRSelection
            Exit Function
        End If
    End Function     


    '-------------------------------------------------------------------------------------
    '* Purpose		: Category_View2
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function Category_View2(p_strRepository, p_lngType, p_strDivisionIDs)

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
	    oCommand.CommandText = "usp_AMOFeature_CT_ViewAll" 
        oCommand.CommandTimeout = 120  

        ' set stored procedure parameter(s)
        Set oParameter = oCommand.CreateParameter("@p_intCategoryType", adInteger, adParamInput, 4)
        oParameter.Value = p_lngType
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIDs", adVarChar, adParamInput, IIf(Len(p_strDivisionIDs) = 0, 1, Len(p_strDivisionIDs)))
        oParameter.Value = p_strDivisionIDs
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
            Set Category_View2 = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set Category_View2 = oRSelection
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
