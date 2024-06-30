<%
'*************************************************************************************
'* FileName		: oDataAMO.asp
'* Description	: Class for AMO Functions and Sub Routines
'* Creator		: Harris, Valerie
'* Created		: 05/27/2016 - PBI 17487/ Task 21005
'*************************************************************************************

Class ISAMO      
    
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
	'Procedure: AMOModule_Search
	'@Purpose:  Search for filtered AMO Options
	'Inputs:    @parm | String  | p_strRepository        | database connection string
	'           @parm | String  | p_strFilter            | WHERE clause for stored procedure
	'Outputs:
	'           @parm ADODB.Recordset | p_oRs            | record data
	'
	'@Returns:  A null object if no errors; otherwise, a collection of errors encountered
	'-----------------------------------------------------------------
    Public Function AMOModule_Search( _
                    p_strRepository, _
                    p_strFilter, _
                    p_strKeyWord, _
                    p_strDivisionIds, _
					p_intSCMId)
		
        'Set optional default value
        If IsNull( p_strKeyWord) = True Or IsEmpty( p_strKeyWord) = True Or p_strKeyWord = "" Then
			p_strKeyWord = ""
		End If

        If IsNull(p_strDivisionIds) = True Or IsEmpty(p_strDivisionIds) = True Or p_strDivisionIds = "" Then
			p_strDivisionIds = ""
		End If
			
		If IsNull(p_intSCMId) = True Or IsEmpty(p_intSCMId) = True Or p_intSCMId = "" Then
			p_intSCMId = 1
		End If


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

        if p_strKeyWord = "" then
            oCommand.CommandText = "usp_AMOFeature_Search"
        else
            oCommand.CommandText = "usp_AMOFeature_SearchByKeyWord"
        end if
            
        oCommand.CommandTimeout = 6000 
    
        'Response.Write("Filter:" & p_strFilter & "<br/>Division IDs:" & p_strDivisionIds & "<br/>SCM ID:" & p_intSCMId & "<br/>Keyword:" & p_strKeyWord & "")
        'Response.End()

		' set stored procedure parameter(s)
        if p_strKeyWord = "" then
			Set oParameter = oCommand.CreateParameter("@p_strFilter", adVarChar, adParamInput, Len(p_strFilter))
			oParameter.Value = p_strFilter
			oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_strDivisionIds", adVarChar, adParamInput, 512)
			oParameter.Value = p_strDivisionIds
			oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intSCMId", adInteger, adParamInput, 2)
			oParameter.Value = p_intSCMId
			oCommand.Parameters.Append oParameter
        end if
			
		Set oParameter = oCommand.CreateParameter("@p_chrKeyWord", adVarChar, adParamInput, 128)
		oParameter.Value = p_strKeyWord
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
            Set AMOModule_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMOModule_Search = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function   


    '-----------------------------------------------------------------
    'Procedure: UpdateStatus
    '@Purpose:  Update an AMO Option's status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngStatusID          | Status ID
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function UpdateStatus(p_strRepository, _
                            p_lngModuleID, _
                            p_lngStatusID, _
                            p_strPerson, _
                            IsDisabled, _
                            p_UserID)
                        
        'set optional default value
        If IsNull(IsDisabled) = True Or IsEmpty(IsDisabled) = True Or IsDisabled = "" Then
            IsDisabled = 0
        End If

        If IsDisabled = True Then
            IsDisabled = 1
        Else
            IsDisabled = 0
        End If
    
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_UpdateStatus"  
		oCommand.CommandType = adCmdStoredProc    
             
        If IsDisabled = True Then
		    ' set stored procedure parameter(s)
            Set oParameter = oCommand.CreateParameter("@p_intFeatureID", adInteger, adParamInput, 4)
			oParameter.Value = p_lngModuleID
			oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intStatusID", adInteger, adParamInput, 4)
			oParameter.Value = p_lngStatusID
			oCommand.Parameters.Append oParameter

			Set oParameter = oCommand.CreateParameter("@p_UpdatedBy", adVarChar, adParamInput, 64)
			oParameter.Value = p_strPerson
			oCommand.Parameters.Append oParameter

            'Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
			'oParameter.Value = p_UserID
			'oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_bitIsDisabled", adInteger, adParamInput, 4)
			oParameter.Value = IsDisabled
			oCommand.Parameters.Append oParameter

        Else
            ' set stored procedure parameter(s)
            Set oParameter = oCommand.CreateParameter("@p_intFeatureID", adInteger, adParamInput, 4)
			oParameter.Value = p_lngModuleID
			oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_intStatusID", adInteger, adParamInput, 4)
			oParameter.Value = p_lngStatusID
			oCommand.Parameters.Append oParameter

			Set oParameter = oCommand.CreateParameter("@p_UpdatedBy", adVarChar, adParamInput, 64)
			oParameter.Value = p_strPerson
			oCommand.Parameters.Append oParameter

            'Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
			'oParameter.Value = p_UserID
			'oCommand.Parameters.Append oParameter
        End If
            
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			UpdateStatus =  Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else 
            UpdateStatus = True
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing

			Exit Function
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: Update Category Rules
    '@Purpose:  Update Category Rules
    'Inputs:    @parm | String  | p_strRepository           | database connection string
    '           @parm | Long    | p_intCategoryId           | Option Category ID
    '           @parm | String  | p_chrRuleDescription      | Rules Description
    '           @parm | Long    | p_intMin                  | Min
    '           @parm | Long    | p_intMax                  | Max
    '           @parm | String  | p_chrDivisionIds          | Division
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function UpdateCategoryRules(p_strRepository, _
                            p_intCategoryID, _
                            p_intRuleID, _
                            p_chrRuleDescription, _
                            p_intMin, _
                            p_intMax, _
                            p_chrDivisionIds)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
        ' Handle unexpected errors
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_UpdateCategoryRules"  
		oCommand.CommandType = adCmdStoredProc    
                         
        Dim p_intValue

		' set stored procedure parameter(s)
        Set oParameter = oCommand.CreateParameter("@p_intCategoryId", adInteger, adParamInput, 4)
		oParameter.Value = p_intCategoryID
		oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_intRuleId", adInteger, adParamInput, 4)
		oParameter.Value = p_intRuleID
		oCommand.Parameters.Append oParameter

		Set oParameter = oCommand.CreateParameter("@p_chrRuleDescription", adVarChar, adParamInput, Len(p_chrRuleDescription))
		oParameter.Value = p_chrRuleDescription
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intMin", adInteger, adParamInput, 4)
		oParameter.Value = p_intMin
		oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_intMax", adInteger, adParamInput, 4)
		oParameter.Value = p_intMax
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIds", adVarChar, adParamInput, 300)
		oParameter.Value = p_chrDivisionIds
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intValue", adInteger, adParamOutput, 1)
		oCommand.Parameters.Append oParameter
       
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			UpdateCategoryRules = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else 
            p_intValue = oCommand.Parameters("@p_intValue").Value
            UpdateCategoryRules = p_intValue
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
    	
			Exit Function
        End If

    End Function

    '-----------------------------------------------------------------
    'Procedure: remove Category Rules
    '@Purpose:  remove Category Rules
    'Inputs:    @parm | String  | p_strRepository           | database connection string
    '           @parm | Long    | p_intRuleId           | Option Rule ID
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function RemoveCategoryRules(p_strRepository, _
                            p_intRuleID)
            
        Set oConnect = Session("oConnect")
		'oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
	    oCommand.CommandText = "usp_AMOFeature_RemoveCategoryRules"
        oCommand.CommandType = adCmdStoredProc   
    
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_intRuleId", adInteger, adParamInput, 4)
		oParameter.Value = p_intRuleID
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
			RemoveCategoryRules = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
				
			Exit Function
        Else 
            RemoveCategoryRules = True
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing

			Exit Function
        End If

    End Function

    '-----------------------------------------------------------------
    'Procedure: UpdateBulkStatus
    '@Purpose:  Update multiple AMO Option's status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_strModuleIDs          | Module ID
    '           @parm | Long    | p_intStatusID          | Status ID
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function UpdateBulkStatus(p_strRepository, _
                            p_intStatusID, _
                            p_strModuleIDs, _
                            p_strPerson, _
                            p_UserID)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_UpdateBulkStatus"  
		oCommand.CommandType = adCmdStoredProc    
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_strModuleIDs",adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intStatusID", adInteger, adParamInput, 4)
		oParameter.Value = p_intStatusID
		oCommand.Parameters.Append oParameter

		Set oParameter = oCommand.CreateParameter("@p_chrPerson", adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then
            UpdateBulkStatus = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
				
            'Close database connection, oConnect.
            oConnect.close
            Set oConnect = Nothing
				
            Exit Function
        Else 
            UpdateBulkStatus = True
				
            'Close database connection, oConnect.
            oConnect.close
            Set oConnect = Nothing
                
            Exit Function
        End If
    
    End Function


    '-----------------------------------------------------------------
    'Procedure: UpdateBulkDate
    '@Purpose:  Update multiple AMO Option's status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_strModuleIDs          | Module ID
    '           @parm | Long    | p_intStatusID          | Status ID
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function UpdateBulkDate(p_strRepository, _
                            chrModuleId_RegionIds, _
                            chrLocalCPLBlindDate, _
                            chrLocalBOMRevADate, _
                            chrLocalRASDisconDate, _
                            charLocObsoleteDate, _
                            charGlobalSeriesDate, _
                            p_strPerson, _
                            p_UserID)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next

        'Response.Write(chrModuleId_RegionIds & "--" & chrLocalCPLBlindDate & "--" & chrLocalBOMRevADate & "--" & chrLocalRASDisconDate & "--" & charLocObsoleteDate & "--" & charGlobalSeriesDate & "--" &  p_strPerson & "--" & p_UserID)
        'Response.End()
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_UpdateBulkDate"   
		oCommand.CommandType = adCmdStoredProc            
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrModuleId_RegionIds",adVarChar, adParamInput, Len(chrModuleId_RegionIds))
		oParameter.Value = chrModuleId_RegionIds
		oCommand.Parameters.Append oParameter
        
        Set oParameter = oCommand.CreateParameter("@p_chrLocalCPLBlindDate",adVarChar, adParamInput, 10)
		oParameter.Value = chrLocalCPLBlindDate
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLocalBOMRevADate",adVarChar, adParamInput, 10)
		oParameter.Value = chrLocalBOMRevADate
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLocalRASDisconDate",adVarChar, adParamInput, 10)
		oParameter.Value = chrLocalRASDisconDate
		oCommand.Parameters.Append oParameter
        
        Set oParameter = oCommand.CreateParameter("@p_charLocObsoleteDate",adVarChar, adParamInput, 10)
		oParameter.Value = charLocObsoleteDate
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_charGlobalSeriesDate",adVarChar, adParamInput, 10)
		oParameter.Value = charGlobalSeriesDate
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson",adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
    
    
        ' execute the stored procedure
        oCommand.Execute 
            
            
        ' return recordset object
        If Err.number <> 0 Then
            UpdateBulkDate = Err.Number & " -- " & 	Err.Description& " -- " &  Err.Source 
				
            'Close database connection, oConnect.
            oConnect.close
            Set oConnect = Nothing
				
            Exit Function
        Else 
            UpdateBulkDate = True
				
            'Close database connection, oConnect.
            oConnect.close
            Set oConnect = Nothing
                
            Exit Function
        End If

    End Function
        
    'This stored procedure isn't there for some reason so don't call this function
    'Public Function AMO_View( _
                    'p_strRepository, _
                    'p_lngID, _
                    'p_rs _
                    ')
        'Set AMO_View = ViewList(p_strRepository, p_lngID, "usp_AMO_View", p_rs)
    'End Function



    '-----------------------------------------------------------------
    'Procedure: AMO_Update
    '@Purpose:  Update an AMO Option's properties
    'Inputs:    @parm | String  | p_strRepository      | database connection string
    '           @parm | Long    | p_lngID              | Module ID
    '           @parm | String  | p_strDesc            | Marketing description
    '           @parm | String  | p_strShortDesc       | Short description
    '           @parm | Long    | p_lngOptionType      | OptionType ID
    '           @parm | Long    | p_lngOptionCategory  | Option Category ID
    '           @parm | String  | p_strBusSegIDs       | Business Segment IDs
    '           @parm | String  | p_strBluePN          | blue part number
    '           @parm | String  | p_strRedPN           | red part number
    '           @parm | String  | p_strBOMRevADate     | BOM Rev A. Date
    '           @parm | String  | p_strRasDisconDate   | RAS Discontinue date
    '           @parm | String  | p_strCPLBlindDate    | CPL Blind date
    '           @parm | String  | p_strAMOCost         | AMO Cost
    '           @parm | String  | p_strAMOWWPrice      | AMO Price
    '           @parm | String  | p_strActualCost      | Actual Cost
    '           @parm | String  | p_strReplacement     | Replacement
    '           @parm | String  | p_strAlternative     | Alternative
    '           @parm | String  | p_strNetWeight       | Net Weight
    '           @parm | String  | p_strExportWeight    | Export Weight
    '           @parm | String  | p_strAirPackedWeight | Air Packed Weight
    '           @parm | String  | p_strAirPackedCubic  | Air Packed Cubic
    '           @parm | String  | p_strExportCubic     | Export Cubic
    '           @parm | String  | p_strPlatformIDs     | Platform IDs assigned to options
    '           @parm | Long    | p_lngGroupID         | Owner Group ID
    '           @parm | Long    | p_lngChangeMask      | ChangeMask field
    '           @parm | String  | p_strUpdater         | Person changing status
    '           @parm | String  | p_strNotes           | Notes field
    '           @parm | Long    | p_lngMOLHide         | Hide option from MOL
    '
    '           @parm | String  | p_sAMOPartNoRe       | This product replaces AMO Part Number
    '           @parm | String  | p_sTargetNA          | Target Lifetime Volume - North America
    '           @parm | String  | p_sTargetLA          | Target Lifetime Volume - LA
    '           @parm | String  | p_sTargetEMEA        | Target Lifetime Volume - EMEA
    '           @parm | String  | p_sTargetAPJ         | Target Lifetime Volume - APJ
    '           @parm | String  | p_sBurdenPer         | Burden Percentage
    '           @parm | String  | p_sContraPer         | Contra percentage
    '           @parm | String  | p_sJustifications    | Low Volume or Low Margin Justification
    '           @parm | int     | sVisibility_NA       | Visibility NA
    '           @parm | int     | sVisibility_EM       | Visibility EMEA
    '           @parm | int     | sVisibility_AP       | Visibility APJ
    '           @parm | int     | sVisibility_LA       | Visibility LA
    'Outputs:
    '           @parm | Variant | p_lngModuleID        | ID after stored in database
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_Update( _
        p_strRepository, p_lngID, _
        p_strDesc, p_strShortDesc, _
        p_lngOptionType, p_lngOptionCategory, _
        p_strBusSegIDs, p_strBluePN, _
        p_strGlobalSeriesDate, p_strBOMRevADate, _
        p_strRasDisconDate, p_strCPLBlindDate, _
        p_strAMOCost, p_strAMOWWPrice, _
        p_strActualCost, p_strReplacement, _
        p_strAlternative, p_strNetWeight, _
        p_strExportWeight, p_strAirPackedWeight, _
        p_strAirPackedCubic, p_strExportCubic, _
        p_strPlatformIDs, p_lngGroupID, _
        p_lngChangeMask, p_strUpdater, _
        p_strNotes, p_strObsoleteDate, p_lngMOLHide, p_lngSCLHide, _
        p_sAMOPartNoRe, p_sTargetNA, _
        p_sTargetLA, p_sTargetEMEA, p_sTargetAPJ, p_sBurdenPer, _
        p_sContraPer, p_sJustifications, p_lngModuleID, _
        sVisibility_NA, sVisibility_EM, sVisibility_AP, sVisibility_LA, _
        p_strRuleID, p_strPath, _
        p_ManufactureCountry, p_WarrantyCode, _
        p_lngSCMHide, p_strLongDesc, _
        p_strRepDesc, p_strOrderIstr, p_chrRuleDescription, _
        p_intClone, p_intLocalized, p_chrComDivIDs, _
        p_chrRegionIDs, p_chrHideRegionIDs, p_intProductLineID, p_intLocalization, p_UserID)
  
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_Update"    
		oCommand.CommandType = adCmdStoredProc

        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intID", adInteger, adParamInput, 4)
        oParameter.Value = p_lngID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDesc", adVarChar, adParamInput, 200)
        oParameter.Value = p_strDesc
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrShortDesc", adVarChar, adParamInput, 40)
        oParameter.Value = p_strShortDesc
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLongDesc", adVarChar, adParamInput, 160) 
        oParameter.Value = p_strLongDesc
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intOptionType", adInteger, adParamInput, 4) 
            oParameter.Value = p_lngOptionType
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intOptionCategory", adInteger, adParamInput, 4) 
            oParameter.Value = p_lngOptionCategory
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, 200)
            oParameter.Value = p_strBusSegIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrBluePN", adVarChar, adParamInput, 20)
            oParameter.Value = p_strBluePN
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRedPN", adVarChar, adParamInput, 20)
        oParameter.Value = ""
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrAMOCost", adVarChar, adParamInput, 20) 
        oParameter.Value = p_strAMOCost
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrAMOWWPrice", adVarChar, adParamInput, 20)
        oParameter.Value = p_strAMOWWPrice
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrActualCost", adVarChar, adParamInput, 20)
        oParameter.Value = p_strActualCost
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrReplacement", adVarChar, adParamInput, 30)
        oParameter.Value = p_strReplacement
        oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_chrAlternative", adVarChar, adParamInput, 30)
        oParameter.Value = p_strAlternative
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrNetWeight", adVarChar, adParamInput, 10)
        oParameter.Value = p_strNetWeight
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrExportWeight", adVarChar, adParamInput, 10)
        oParameter.Value = p_strExportWeight
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrAirPackedWeight", adVarChar, adParamInput, 10)
        oParameter.Value = p_strAirPackedWeight
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrAirPackedCubic", adVarChar, adParamInput, 10)
        oParameter.Value = p_strAirPackedCubic
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrExportCubic", adVarChar, adParamInput, 10)
        oParameter.Value = p_strExportCubic
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPlatformIDs", adVarChar, adParamInput, 4000)
        oParameter.Value = p_strPlatformIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intGroupID", adInteger, adParamInput, 4)
        oParameter.Value = p_lngGroupID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intChangeMask", adInteger, adParamInput, 4)
        oParameter.Value = p_lngChangeMask
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUpdater", adVarChar, adParamInput, 64)
        oParameter.Value = p_strUpdater
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrNotes", adVarChar, adParamInput, 1024)
        oParameter.Value = p_strNotes
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intMOLHide", adInteger, adParamInput, 4)
        oParameter.Value = p_lngMOLHide
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCLHide", adInteger, adParamInput, 4)
        oParameter.Value = p_lngSCLHide
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sAMOPartNoRe", adVarChar, adParamInput, 30)
        oParameter.Value = p_sAMOPartNoRe
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sTargetNA", adVarChar, adParamInput, 10)
        oParameter.Value = p_sTargetNA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sTargetLA", adVarChar, adParamInput, 10)
        oParameter.Value = p_sTargetLA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sTargetEMEA", adVarChar, adParamInput, 10)
        oParameter.Value = p_sTargetEMEA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sTargetAPJ", adVarChar, adParamInput, 10)
        oParameter.Value = p_sTargetAPJ
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sBurdenPer", adVarChar, adParamInput, 10)
        oParameter.Value = p_sBurdenPer
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sContraPer", adVarChar, adParamInput, 10)
        oParameter.Value = p_sContraPer
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_sJustifications", adVarChar, adParamInput, 300)
        oParameter.Value = p_sJustifications
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_Visibility_NA", adInteger, adParamInput, 4)
        oParameter.Value = sVisibility_NA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_Visibility_EM", adInteger, adParamInput, 4)
        oParameter.Value = sVisibility_EM
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_Visibility_AP", adInteger, adParamInput, 4)
        oParameter.Value = sVisibility_AP
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_Visibility_LA", adInteger, adParamInput, 4)
        oParameter.Value = sVisibility_LA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRuleID", adVarChar, adParamInput, IIf(Len(p_strRuleID) > 0, Len(p_strRuleID), 1))
        oParameter.Value = p_strRuleID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_ManufactureCountry", adVarChar, adParamInput, 2)
        oParameter.Value = p_ManufactureCountry
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_WarrantyCode", adVarChar, adParamInput, IIf(Len(p_WarrantyCode) > 0, Len(p_WarrantyCode), 1))
        oParameter.Value = p_WarrantyCode
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMHide", adInteger, adParamInput, 4)
        oParameter.Value = p_lngSCMHide
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRepDesc", adVarChar, adParamInput, 80)
        oParameter.Value = p_strRepDesc
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOrderIstr", adVarChar, adParamInput, 600)
        oParameter.Value = p_strOrderIstr
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRuleDescription", adVarChar, adParamInput, IIf(Len(p_chrRuleDescription) > 0, Len(p_chrRuleDescription), 1))
        oParameter.Value = p_chrRuleDescription
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intLocalized", adInteger, adParamInput, 1)
        oParameter.Value = p_intLocalized
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrComDivIDs", adVarChar, adParamInput, 512)
        oParameter.Value = p_chrComDivIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intProductLineID", adInteger, adParamInput, 2)
        oParameter.Value = p_intProductLineID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamOutput, 4)
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPath", adVarChar, adParamOutput, 255)
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUpdaterID", adInteger, adParamInput, 4)
        oParameter.Value = p_UserID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intIDP", adInteger, adParamInput, 4)
        oParameter.Value = intIgnoreDeploy
        oCommand.Parameters.Append oParameter
            
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If  Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_Update = Err.Number & " -- " &  Err.Description & " -- " &  Err.Source  
            Exit Function
        Else
            bProcessComplete = True
        End If
        
        If bProcessComplete = True Then
    
                p_lngModuleID = oCommand.Parameters("@p_intModuleID").Value
                p_strPath = oCommand.Parameters("@p_chrPath").Value

                
            If p_intLocalization = 1 Then
                    
                'reset command and connection objects
                If IsObject(oCommand) = True Then
                    'disconnect the command object
                    If Not (oCommand Is Nothing) Then
                        Set oCommand.ActiveConnection = Nothing
                        Set oCommand = Nothing
                    End If
                End If

                If IsObject(oConnect) = True Then
                    'Close database connection, oConnect.
                    If Not (oConnect Is Nothing) Then
                        oConnect.close
                        Set oConnect = Nothing
                    End If
                End If
                       
                
                Set oConnect = Server.CreateObject("ADODB.Connection")
			    oConnect.Open(PULSARDB())
			    Set oCommand = Server.CreateObject("ADODB.Command")
       
			    On Error Resume Next
			
            	' set database stored procedure command
			    Set oCommand.ActiveConnection = oConnect
			    oCommand.CommandText = "IRS_usp_AMO_UpdateLocalizationDates" 
			    oCommand.CommandType = adCmdStoredProc    
    
                              
                Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
                oParameter.Value = p_lngModuleID
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrLocalCPLBlindDates", adVarChar, adParamInput, IIf(Len(p_strCPLBlindDate) > 0, Len(p_strCPLBlindDate), 1))
                oParameter.Value = p_strCPLBlindDate
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrLocalBOMRevADates", adVarChar, adParamInput, IIf(Len(p_strBOMRevADate) > 0, Len(p_strBOMRevADate), 1))
                oParameter.Value = p_strBOMRevADate
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrLocalRASDisconDates", adVarChar, adParamInput, IIf(Len(p_strRasDisconDate) > 0, Len(p_strRasDisconDate), 1))
                oParameter.Value = p_strRasDisconDate
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_charLocObsoleteDates", adVarChar, adParamInput, IIf(Len(p_strObsoleteDate) > 0, Len(p_strObsoleteDate), 1)) 
                oParameter.Value = p_strObsoleteDate
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_charGlobalSeriesDate", adVarChar, adParamInput, IIf(Len(p_strGlobalSeriesDate) > 0, Len(p_strGlobalSeriesDate), 1)) 
                oParameter.Value = p_strGlobalSeriesDate
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrRegionIDs", adVarChar, adParamInput, 1024)
                oParameter.Value = p_chrRegionIDs
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrHideRegionIDs", adVarChar, adParamInput, 1024)
                oParameter.Value = p_chrHideRegionIDs
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_chrUpdater", adVarChar, adParamInput, 64)
                oParameter.Value = p_strUpdater
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_intClone", adInteger, adParamInput, 64)
                oParameter.Value = p_intClone
                oCommand.Parameters.Append oParameter

                Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
                oParameter.Value = p_UserID
                oCommand.Parameters.Append oParameter
                    
                ' execute the stored procedure
                oCommand.Execute 
            
                ' return recordset object
                If  Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
                    AMO_Update = Err.Number & " -- " &  Err.Description & " -- " &  Err.Source 

                    'Close database connection, oConnect.
				    oConnect.close
				    Set oConnect = Nothing
                        
                    Exit Function
                Else
                    AMO_Update = True
                    Exit Function
                End If
            Else 'if p_Localization not 1 and first process is complete
                AMO_Update = True
                Exit Function
            End If
      
        Else  'if process isn't complete, return error and close connection      
            AMO_Update = Err.Number & " -- " &  Err.Description & " -- " &  Err.Source 

            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing

            Exit Function
        End If
  
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMORegions_Search
    '@Purpose:  Get an AMO Option's regions
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMORegions_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)
            
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
        oCommand.CommandText = "usp_AMOFeature_RegionsSearch"

        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMORegions_Search = Nothing
            Exit Function
        Else
                'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMORegions_Search = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMOGEOs_Search
    '@Purpose:  Get an AMO Option's GEOs
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMOGEOs_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)
            
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
        oCommand.CommandText = "usp_AMOFeature_GEOsSearch"

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
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
            Set AMOGEOs_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMOGEOs_Search = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
	End Function
    '-----------------------------------------------------------------
    'Procedure: AMOPlatforms_Search
    '@Purpose:  Get an AMO Option's platforms
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    '           @parm | Integeer   | p_intReturnRegions | 1=return Regions with the Platforms, 0=no regions returned
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMOPlatforms_Search( _
					p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID, _
                    p_intReturnRegions)

        If IsNull(p_intReturnRegions) = True Or IsEmpty(p_intReturnRegions) = True Or p_intReturnRegions = ""  Then
            p_intReturnRegions = 1
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
        oCommand.CommandText = "usp_AMOFeature_PlatformsSearch"

        'Response.Write("'" & p_strModuleIDs & "'," & p_lngSCMID & "," & p_intReturnRegions)
        'Response.End()

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intReturnRegions", adInteger, adParamInput, 4)
		oParameter.Value = p_intReturnRegions
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
            Set AMOPlatforms_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMOPlatforms_Search = oRSelection

            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: AMO_AllPlatforms_Search
    '@Purpose:  Get an All Option's platforms
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_AllPlatforms_Search(p_strRepository, _
                        p_strDivisionIds)
         
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
        oCommand.CommandText = "usp_AMO_PLF_ViewAll"

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strDivisionIds))
		oParameter.Value = p_strDivisionIds
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
            Set AMO_AllPlatforms_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_AllPlatforms_Search = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AllRegions_Search
    '@Purpose:  Get all active regions in the specified divisions
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strDivisionIDs   | comma delimited string of division ids
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AllRegions_Search( _
                    p_strRepository, _
                    p_strDivisionIds, _
                    p_IsHubReport)

        If IsNull(p_IsHubReport) = True Or IsEmpty(p_IsHubReport) Or p_IsHubReport = "" Then
            p_IsHubReport = 0
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
        oCommand.CommandText = "usp_AMOFeature_AllRegionsSearch"

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIDs", adVarChar, adParamInput, Len(p_strDivisionIds))
		oParameter.Value = p_strDivisionIds
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_bitHubReport", adInteger, adParamInput, 1)
		oParameter.Value = p_IsHubReport
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
            Set AllRegions_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AllRegions_Search = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function

        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AllRegions_Search_SCL
    '@Purpose:  Get all active regions in the specified divisions, for SCL
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strDivisionIDs   | comma delimited string of division ids
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AllRegions_Search_SCL( _
                    p_strRepository, _
                    p_strDivisionIds)
            
        Set oConnect = Session("oConnect")
		'oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
        oCommand.CommandText = "usp_AMO_AllRegionsSearch_SCL"

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIDs", adVarChar, adParamInput, Len(p_strDivisionIds))
		oParameter.Value = p_strDivisionIds
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AllRegions_Search_SCL = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AllRegions_Search_SCL = oRSelection
            
            Exit Function
        End If
    End Function
    '-----------------------------------------------------------------
    'Procedure: AMOCompatibility_Search
    '@Purpose:  Get Compatibility divisions
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs   | comma delimited string of Module ids
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMOCompatibility_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_intSCMId)
            
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
        oCommand.CommandText = "usp_AMOFeature_CompatibilitySearch"

        ' get database connection and command
        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, IIf(Len(p_strModuleIDs) > 0, Len(p_strModuleIDs), 1))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_intSCMId
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
            Set AMOCompatibility_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMOCompatibility_Search = oRSelection

            ''Set Session("oConnect") = oConnect
            Exit Function
        End If

    End Function

    '-----------------------------------------------------------------
    'Procedure: SavePlatformStatus
    '@Purpose:  Updates an AMO Option's Platform status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngPlatformID        | Platform ID
    '           @parm | Long    | p_lngValue             | 1=set, 0=clear
    '           @parm | Long    | p_chrNewValue              | user set date
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SavePlatformStatus(p_strRepository, _
                            p_lngModuleID, _
                            p_lngPlatformID, _
                            p_lngValue, _
                            p_chrNewValue, _
                            p_strPerson, _
                            p_UserID)
         
        'Response.Write(p_lngModuleID & "--" & p_lngPlatformID & "--" & p_lngValue & "--" & p_chrNewValue & "--" & p_strPerson & "--" & p_UserID)
        'Response.End()

        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SavePlatformStatus"
		oCommand.CommandType = adCmdStoredProc
                        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPlatformID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngPlatformID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intValue", adInteger, adParamInput, 4)
		oParameter.Value = p_lngValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrNewValue",adVarChar, adParamInput, Len(p_chrNewValue))
		oParameter.Value = p_chrNewValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson",adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            SavePlatformStatus = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            SavePlatformStatus = True
                
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveComparabilityStatus
    '@Purpose:  Updates an AMO Option's Comparability status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngDivisionID        | Division ID
    '           @parm | Long    | p_lngValue             | 1=set, 0=clear
    '           @parm | Long    | p_chrNewValue              | user set date
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveComparabilityStatus(p_strRepository, _
                            p_lngModuleID, _
                            p_lngDivisionID, _
                            p_lngValue, _
                            p_strPerson, _
                            p_UserID)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveComparabilityStatus" 
		oCommand.CommandType = adCmdStoredProc        
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intDivisionID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngDivisionID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intValue", adInteger, adParamInput, 4)
		oParameter.Value = p_lngValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson",adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            SaveComparabilityStatus = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            SaveComparabilityStatus = True
                
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveRegionStatus
    '@Purpose:  Updates an AMO Option's Region status
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngRegionID          | Region ID
    '           @parm | Long    | p_lngValue             | 1=set, 0=clear
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveRegionStatus(p_strRepository, _
                            p_lngModuleID, _
                            p_lngRegionID, _
                            p_lngValue, _
                            p_strPerson, _ 
                            p_UserID)

        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveRegionStatus"  
		oCommand.CommandType = adCmdStoredProc        
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intRegionID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngRegionID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intValue", adInteger, adParamInput, 4)
		oParameter.Value = p_lngValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson",adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            SaveRegionStatus = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            SaveRegionStatus = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If

    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveFieldValue
    '@Purpose:  Updates an AMO Option's field
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | String  | p_strField             | lowercase field name to change
    '           @parm | String  | p_strNewValue          | value to store
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveFieldValue(p_strRepository, _
                            p_lngModuleID, _
                            p_strField, _
                            p_strNewValue, _
                            p_strPerson, _
							p_UserID)
    
		Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveFieldValue" 
		oCommand.CommandType = adCmdStoredProc
			
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 6)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrField", adVarChar, adParamInput, 30)
		oParameter.Value = p_strField
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrNewValue", adVarChar, adParamInput, 2000)
		oParameter.Value = p_strNewValue
		oCommand.Parameters.Append oParameter

		Set oParameter = oCommand.CreateParameter("@p_chrPerson", adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 6)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
        				
        ' execute the stored procedure
        oCommand.Execute 
            			
        ' return recordset object
        If Err.number <> 0 Then
			SaveFieldValue = SaveFieldValue = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else 
            SaveFieldValue = True
				
			'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                	
			Exit Function
        End If
			
    End Function

    '-----------------------------------------------------------------
    'Procedure: SaveDateFieldValue
    '@Purpose:  Updates an AMO Option's field
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngRgionID           | Region ID
    '           @parm | String  | p_strField             | lowercase field name to change
    '           @parm | String  | p_strNewValue          | value to store
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveDateFieldValue(p_strRepository, _
                            p_lngModuleID, _
                            p_lngRegionID, _
                            p_strField, _
                            p_strNewValue, _
                            p_strPerson,_
                            p_UserID)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
	    oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveDateFieldValue"  
		oCommand.CommandType = adCmdStoredProc        
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter    
    
        Set oParameter = oCommand.CreateParameter("@p_intRegionID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngRegionID
		oCommand.Parameters.Append oParameter    

        Set oParameter = oCommand.CreateParameter("@p_chrField",adVarChar, adParamInput, 30)
		oParameter.Value = p_strField
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrNewValue",adVarChar, adParamInput, 1000)
		oParameter.Value = p_strNewValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson", adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            SaveDateFieldValue = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
            
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            SaveDateFieldValue = True
            
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If

    End Function
    '-----------------------------------------------------------------
    'Procedure: SaveGEODate
    '@Purpose:  Updates an AMO Option's GEO Date
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | Long    | p_lngModuleID          | Module ID
    '           @parm | Long    | p_lngGEOID             | GEO ID
    '           @parm | String  | p_strValue             | new date
    '           @parm | String  | p_strPerson            | Person changing status
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function SaveGEODate(p_strRepository, _
                            p_lngModuleID, _
                            p_lngGEOID, _
                            p_strValue, _
                            p_strPerson, _ 
                            p_UserID)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveGEODate"  
		oCommand.CommandType = adCmdStoredProc
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngModuleID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intGEOID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngGEOID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrValue", adInteger, adParamInput, 20)
		oParameter.Value = p_strValue
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPerson",adVarChar, adParamInput, 64)
		oParameter.Value = p_strPerson
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intPersonID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            SaveGEODate = Err.Number & " -- " & Err.Description & " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            SaveGEODate = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If
    
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_Publish
    '@Purpose:  publish/snapshot AMO_SCM data
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | string    | p_strCreator          | person created the snapshot
    '           @parm | String  | p_strModuleIDs           | Modules that will be taken a snapshot
    'Outputs:
    '           none
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_Publish(p_strRepository, _
                            p_strCreator, _
                            p_strModuleIDs, _
                            p_lngSCMID, _
                            p_strDivisionIds, _
                            p_intFormatType)

        If IsNull(p_strDivisionIds) = True Or IsEmpty(p_strDivisionIds) = True Or p_strDivisionIds = "" Then
            p_strDivisionIds = ""
        End If    
    
        If IsNull(p_intFormatType) = True Or IsEmpty(p_intFormatType) = True Or p_intFormatType = "" Then
            p_intFormatType = 0
        End If     

                        
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SCM_Publish"  
		oCommand.CommandType = adCmdStoredProc             
        
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrName",adVarChar, adParamInput, 64)
		oParameter.Value = p_strCreator
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs",adVarChar, adParamInput, IIf(Len(p_strModuleIDs) > 0, Len(p_strModuleIDs), 1))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIds",adVarChar, adParamInput, 512)
		oParameter.Value = p_strDivisionIds
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intFormatType", adInteger, adParamInput, 4)
		oParameter.Value = p_intFormatType
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamOutput, 4)
		oCommand.Parameters.Append oParameter
    
        ' execute the stored procedure
        oCommand.Execute 

        ' return recordset object
        If  Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_SCM_Publish = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
            
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            p_lngSCMID = oCommand.Parameters("@p_intSCMID").Value
            AMO_SCM_Publish = p_lngSCMID
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: ViewAll_AMO_SCM_PublishList
    '@Purpose:  Get the list of all published AMO_SCM/filtered by bus segment
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function ViewAll_AMO_SCM_PublishList( _
                    p_strRepository, _
                    p_strBusSegIDs)
            
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        'Response.Write(p_strBusSegIDs)
        'Response.End()

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = "usp_AMOFeature_ViewAllPublishes"

        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, 500)
		oParameter.Value = p_strBusSegIDs
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
            Set ViewAll_AMO_SCM_PublishList = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewAll_AMO_SCM_PublishList = oRSelection
            ''Set Session("oConnect") = oConnect
        End If
    End Function
    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_Comparison
    '@Purpose:  compare common modules in two snapshots and return the diffreence
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | String  | p_strSCMIDs        | Two snapshot IDs
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_Comparison( _
                    p_strRepository, _
                    p_strSCMIDs)
            
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
	    oCommand.CommandText = "usp_AMOFeature_SCM_Compare"

        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrSCMIDs", adVarChar, adParamInput, 600)
		oParameter.Value = p_strSCMIDs
		oCommand.Parameters.Append oParameter

        '  ' execute the stored procedure
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
            Set AMO_SCM_Comparison = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_SCM_Comparison = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function

        End If
    End Function
    '-----------------------------------------------------------------
    'Procedure: ViewTwo_AMO_SCM_Publish
    '@Purpose:  Get info for the two published AMO_SCM comparison spreadsheet
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | String  | p_strSCMIDs        | Two snapshot IDs
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function ViewTwo_AMO_SCM_Publish( _
                    p_strRepository, _
                    p_strSCMIDs)

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
	    oCommand.CommandText = "usp_AMOFeature_ViewTwoPublishes"

        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrSCMIDs", adVarChar, adParamInput, 600)
		oParameter.Value = p_strSCMIDs
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
            Set ViewTwo_AMO_SCM_Publish = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set ViewTwo_AMO_SCM_Publish = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_CompareCurrent
    '@Purpose:  compare modules in ONE snapshot WITH CURRENT DATA and return the diffreence
    'Inputs:    @parm | String  | p_strRepository       | database connection string
    '           @parm | String  | p_strSCMIDs           | Two snapshot IDs
    '           @parm | String  | p_strFilter           | TSQL Query String
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_CompareCurrent( _
                    p_strRepository, _
                    p_strSCMIDs, _
                    p_strFilter)

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
	    oCommand.CommandText = "usp_AMOFeature_SCM_CompareCurrent"

        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrSCMIDs", adVarChar, adParamInput, 600)
		oParameter.Value = p_strSCMIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrFilter", adVarChar, adParamInput, 3000)
		oParameter.Value = p_strFilter
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
            Set AMO_SCM_CompareCurrent = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_SCM_CompareCurrent = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_AllRegions
    '@Purpose:  Get all AMO module option regions
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_AllRegions( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)

        Set oConnect = Session("oConnect")
		'oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = "usp_AMOFeature_SCM_AllRegions"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_SCM_AllRegions = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
            'Set oRSelection.ActiveConnection = Nothing
                
            'return recordset
            Set AMO_SCM_AllRegions = oRSelection

            'Close database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing

            'oRSelection.Close
            'Set oRSelection = Nothing
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: AVAMO_Modules
    '@Purpose:  get the moduleIDs for the AVAMO_Matrix report
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | String  | p_strFilter            | WHERE clause for stored procedure
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AVAMO_Modules( _
                    p_strRepository, _
                    p_strFilter)

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
	    oCommand.CommandText = "IRS_RPT_AVAMO_Modules"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrFilter", adVarChar, adParamInput, 2000)
		oParameter.Value = p_strFilter
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
            Set AVAMO_Modules = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
            'Set oRSelection.ActiveConnection = Nothing
                
            'return recordset
            Set AVAMO_Modules = oRSelection
            ''Set Session("oConnect") = oConnect    
            Exit Function
        End If
    End Function



    '-----------------------------------------------------------------
    'Procedure: AVAMO_Matrix
    '@Purpose:  get the moduleIDs for the AVAMO_Matrix report
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | String  | p_strFilter            | WHERE clause for stored procedure
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AVAMO_Matrix( _
                    p_strRepository, _
                    p_strFilter)

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
	    oCommand.CommandText = "IRS_RPT_AVAMO_Matrix"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrFilter", adVarChar, adParamInput, 2000)
		oParameter.Value = p_strFilter
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
                
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AVAMO_Matrix = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AVAMO_Matrix = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: AVAMO_Allproductfamily
    '@Purpose:  get all the platform product families for the AVAMO_Matrix report
    'Inputs:    @parm | String  | p_strRepository        | database connection string
    '           @parm | String  | p_strFilter            | WHERE clause for stored procedure
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AVAMO_Allproductfamily( _
                    p_strRepository)
 
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
	    oCommand.CommandText = "IRS_RPT_AVAMO_AllProductFamily"
            			
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
            Set AVAMO_Allproductfamily = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AVAMO_Allproductfamily = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_Platforms_Search
    '@Purpose:  Get an AMO Option's platforms for the SCM report
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_Platforms_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)

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
	    oCommand.CommandText = "usp_AMOFeature_SCM_PlatformsSearch"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
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
            Set AMO_SCM_Platforms_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_SCM_Platforms_Search = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SCM_Regions_Search
    '@Purpose:  Get an AMO Option's regions for the SCM report
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SCM_Regions_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)

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
	    oCommand.CommandText = "usp_AMOFeature_SCM_RegionsSearch"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSCMID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngSCMID
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
            Set AMO_SCM_Regions_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_SCM_Regions_Search = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function
    '-----------------------------------------------------------------
    'Procedure: AMO_Report_RASDEtail_Regions_Search
    '@Purpose:  Get an AMO Option's regions for the SCM report
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_Report_RASDEtail_Regions_Search( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)

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
	    oCommand.CommandText = "IRS_usp_AMO_Report_RASDEtail_ModuleRegions"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
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
            Set AMO_Report_RASDEtail_Regions_Search = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_Report_RASDEtail_Regions_Search = oRSelection

            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_ClonefromCTO
    '@Purpose:  Update an AMO Option's properties
    'Inputs:    @parm | String  | p_strRepository      | database connection string
    '           @parm | Long    | p_lngID              | Module ID
    '           @parm | String  | p_strUpdater         | Person changing status

    '           @parm | int     | sVisibility_LA       | Visibility LA
    'Outputs:
    '           @parm | Variant | p_lngModuleID        | ID after stored in database
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_ClonefromCTO( _
        p_strRepository, _ 
        p_lngID, _
        p_strUpdater, _
        p_UserID)
  
        Set oConnect = Session("oUpdateConnect")
		'oConnect.Open(IRSDB(p_strRepository))
        Set oCommand = Server.CreateObject("ADODB.Command")

		'oConnect.Open(PULSARDB())
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "IRS_usp_AMO_ClonefromCTO"
		oCommand.CommandType = adCmdStoredProc        
                
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_intModuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_lngID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUpdater", adVarChar, adParamInput, 64)
		oParameter.Value = p_strUpdater
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intModuleID_output", adInteger, adParamOutput, 4)
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUpdaterID", adInteger, adParamInput, 4)
		oParameter.Value = p_UserID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_ClonefromCTO = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
                
			Exit Function
        Else
            p_lngModuleID = oCommand.Parameters("@p_intModuleID_output").Value
            AMO_ClonefromCTO = p_lngModuleID
                
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
                
			Exit Function
        End If
         
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_Report_AllRegions
    '@Purpose:  Get All regions for the AMO reports/discontinuance report
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_Report_AllRegions( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)
          
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
	    oCommand.CommandText = "IRS_usp_AMO_Report_AllRegions"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
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
            Set AMO_Report_AllRegions = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_Report_AllRegions = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If

    End Function


    '-----------------------------------------------------------------
    'Procedure: AMO_Report_AllGEOs
    '@Purpose:  Get All regions for the AMO reports/discontinuance report
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | String     | p_strModuleIDs     | comma delimited string of module ids
    '           @parm | Long       | p_lngSCMID         | SCMID for joining with AMO_SCM table
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_Report_AllGEOs( _
                    p_strRepository, _
                    p_strModuleIDs, _
                    p_lngSCMID)

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
	    oCommand.CommandText = "usp_AMO_Report_AllGEOs"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrModuleIDs", adVarChar, adParamInput, Len(p_strModuleIDs))
		oParameter.Value = p_strModuleIDs
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
            Set AMO_Report_AllGEOs = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_Report_AllGEOs = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function


    '-----------------------------------------------------------------
    'Procedure: AMO_RulesSearch
    '@Purpose:  Get All Option category Rules
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_intCategoryID    | CategoryID
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_RulesSearch( _
                    p_strRepository, _
                    p_intCategoryID, _
                    p_intRuleID)
          
        'Response.Write(p_intCategoryID & "--" & p_intRuleID)
        'Response.End()

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
	    oCommand.CommandText = "IRS_usp_AMO_RulesSearch"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_intCategoryID", adInteger, adParamInput, 4)
		oParameter.Value = p_intCategoryID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intRuleID", adInteger, adParamInput, 4)
		oParameter.Value = p_intRuleID
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
            Set AMO_RulesSearch = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            Set AMO_RulesSearch = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If

    End Function

        
    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ValidateData
    '* Inputs		: 
    '-------------------------------------------------------------------------------------

    Public Function AMO_ValidateData( _
                    p_strRepository, _
                    p_strDivisionIds, _
                    p_strRASDiscontinueDate)
                
          
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
		oCommand.CommandText = "usp_AMOFeature_ValidateData"  
	         
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_RAS_DiscontinueDate", adVarChar, adParamInput, 10)
		oParameter.Value = Trim(p_strRASDiscontinueDate)
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDivisionIDs", adVarChar, adParamInput, 512)
		oParameter.Value = p_strDivisionIds
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
            Set AMO_ValidateData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_ValidateData = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If

    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ChangeHistory_AMOCategories
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_ChangeHistory_AMOCategories( _
                    p_strRepository)

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
	    oCommand.CommandText = "usp_AMOFeature_ChangeHistory_AMOCategories"
            
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
            Set AMO_ChangeHistory_AMOCategories = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_ChangeHistory_AMOCategories = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

        '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ChangeHistory_AllUpdaters
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_ChangeHistory_AllUpdaters( _
                    p_strRepository)
                
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
	    oCommand.CommandText = "usp_AMOFeature_ChangeHistory_AllUpdaters"
            
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
            Set AMO_ChangeHistory_AllUpdaters = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_ChangeHistory_AllUpdaters = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ChangeHistory
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_ChangeHistory( _
                    p_strRepository, _
                    p_chrSearchFilter)

        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        'On Error Resume Next

        'Response.Write(p_chrSearchFilter)
        'Response.End()

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = "usp_AMOFeature_ChangeHistory"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrSearchFilter", adVarChar, adParamInput, 4000)
		oParameter.Value = p_chrSearchFilter
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
            Set AMO_ChangeHistory = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_ChangeHistory = oRSelection

            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ChangeHistory_Update
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_ChangeHistory_Update(p_strRepository, _
                            p_chrReasonIDs, _
                            p_chrReasons, _
                            p_chrHidefromSCMIDs)
    
        Set oConnect = Session("oConnect")
		'oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_ChangeHistory_UpdateReason" 
		oCommand.CommandType = adCmdStoredProc 
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrReasonIDs", adVarChar, adParamInput, IIf(Len(p_chrReasonIDs) > 0, Len(p_chrReasonIDs), 1))
		oParameter.Value = p_chrReasonIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrReasons", adVarChar, adParamInput, IIf(Len(p_chrReasons) > 0, Len(p_chrReasons), 1))
		oParameter.Value = p_chrReasons
		oCommand.Parameters.Append oParameter

		Set oParameter = oCommand.CreateParameter("@p_chrHidefromSCMIDs", adLongVarChar, adParamInput, IIf(Len(p_chrHidefromSCMIDs) > 0, Len(p_chrHidefromSCMIDs), 1))
		oParameter.Value = p_chrHidefromSCMIDs
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_ChangeHistory_Update = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
				
			'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
				
			Exit Function
        Else
            AMO_ChangeHistory_Update = True
                
            'Open database connection, oConnect.
			'oConnect.close
			'Set oConnect = Nothing
                
			Exit Function
        End If
    End Function


    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_ChangeHistory_AllChangeTypes
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_ChangeHistory_AllChangeTypes( _
                    p_strRepository)

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
	    oCommand.CommandText = "usp_AMOFeature_ChangeHistory_AllChangeTypes"
            
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
            Set AMO_ChangeHistory_AllChangeTypes = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing

            Set AMO_ChangeHistory_AllChangeTypes = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SaveDFT
    '@Purpose:  Save data DFT for AMO before Generate 2 txt files
    'Inputs:    @parm      | String        | p_strRepository    | database connection string
    '           @parm long | p_lngUserID   | UserID
    'Outputs:
    '          
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SaveDFT(p_strRepository, _
                    p_strBusSegIDs, p_strOwnerIDs, p_intUserID, _
                    p_strUserName, p_intProductLineID, _
                    p_intSupplier_Code_DivisionID, p_strSupplier_Code, _
                    p_strHW_PROD_FAMILY, p_strSW_PROD_FAMILY, _
                    p_strHW_TAX_CLASS_CD, p_strSW_TAX_CLASS_CD, _
                    p_strOverride_COM_Code, p_strA, _
                    p_strBUS_DEF_FIELD4, p_strCTRY_CD, _
                    p_strCURR_CD, p_strDIFF_CD, _
                    p_strENTRY_SOURCE_CD, p_strM, _
                    p_strMKT, p_strMKT_CD, _
                    p_strPA_DISC_FLG, p_strPRC_DISP_CD, _
                    p_strPRC_TERM_CD, p_strPROD, _
                    p_strPROD_DISP_EXCL_CD, p_strQBL_SEQ_NBR, _
                    p_strSERIAL_FLG, p_strSERV_CD, _
                    p_strSRT_CD, p_strSUI, _
                    p_strUOM_CD)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveDFT"
		oCommand.CommandType = adCmdStoredProc        

        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1) 
        oParameter.Value = p_strBusSegIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1) 
        oParameter.Value = p_strOwnerIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
        oParameter.Value = p_intUserID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUserName", adVarChar, adParamInput, 64)
        oParameter.Value = p_strUserName
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intProductLineID", adInteger, adParamInput, cg_lngLEN_INT)
        oParameter.Value = p_intProductLineID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intSupplier_Code_DivisionID", adInteger, adParamInput, cg_lngLEN_INT)
        oParameter.Value = p_intSupplier_Code_DivisionID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSupplier_Code", adVarChar, adParamInput, Len(p_strSupplier_Code))
        oParameter.Value = p_strSupplier_Code
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrHW_PROD_FAMILY", adVarChar, adParamInput, 4)
        oParameter.Value = p_strHW_PROD_FAMILY
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSW_PROD_FAMILY", adVarChar, adParamInput, 4)
        oParameter.Value = p_strSW_PROD_FAMILY
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrHW_TAX_CLASS_CD", adVarChar, adParamInput, 4)
        oParameter.Value = p_strHW_TAX_CLASS_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSW_TAX_CLASS_CD", adVarChar, adParamInput, 4)
        oParameter.Value = p_strSW_TAX_CLASS_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOverride_COM_Code", adVarChar, adParamInput, 2)
        oParameter.Value = p_strOverride_COM_Code
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrA", adVarChar, adParamInput, 16)
        oParameter.Value = p_strA
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrBUS_DEF_FIELD4", adVarChar, adParamInput, 16)
        oParameter.Value = p_strBUS_DEF_FIELD4
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrCTRY_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strCTRY_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrCURR_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strCURR_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDIFF_CD", adVarChar, adParamInput, 16)
        oParameter.Value =  p_strDIFF_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrENTRY_SOURCE_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strENTRY_SOURCE_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrM", adVarChar, adParamInput, 16)
        oParameter.Value = p_strM
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrMKT", adVarChar, adParamInput, 16)
        oParameter.Value = p_strMKT
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrMKT_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strMKT_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPA_DISC_FLG", adVarChar, adParamInput, 16)
        oParameter.Value = p_strPA_DISC_FLG
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPRC_DISP_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strPRC_DISP_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPRC_TERM_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strPRC_TERM_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPROD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strPROD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPROD_DISP_EXCL_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strPROD_DISP_EXCL_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrQBL_SEQ_NBR", adVarChar, adParamInput, 16)
        oParameter.Value = p_strQBL_SEQ_NBR
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSERIAL_FLG", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSERIAL_FLG
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSERV_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSERV_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSRT_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSRT_CD
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSUI", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSUI
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUOM_CD", adVarChar, adParamInput, 16)
        oParameter.Value = p_strUOM_CD
        oCommand.Parameters.Append oParameter
   
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.number <> 0 Then	    'CUSTOM ERROR HANDLING.   
            AMO_SaveDFT = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        Else
            AMO_SaveDFT = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
				
			Exit Function
        End If
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_GenerateDFT
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_GenerateDFT( _
                    p_strRepository, _
                    p_lngUserID)

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
	    oCommand.CommandText = "usp_AMOFeature_GenerateDFT"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GenerateDFT = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GenerateDFT = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function

        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_AnyDFTFileData
    '@Purpose:  Return information to determine if a Base and Localized DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_AnyDFTFileData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs)
 
        'Response.Write(p_strBusSegIDs & "--" & p_strOwnerIDs)
        'Response.End()

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
	    oCommand.CommandText = "usp_AMOFeature_AnyDFTFileData"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter
        
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_AnyDFTFileData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_AnyDFTFileData = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
               
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_GetDFTFileData_Localized
    '@Purpose:  Return information so a Localized DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetDFTFileData_Localized( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID)
 
        'Response.Write(p_strBusSegIDs & "---" & p_strOwnerIDs & "--" & p_lngUserID & "--1")
        'Response.End()

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
	    oCommand.CommandText = "usp_AMOFeature_GetDFTFileData"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intLocalized", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = 1
		oCommand.Parameters.Append oParameter

           
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
        If Err.Number > 0 Then		    
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetDFTFileData_Localized = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GetDFTFileData_Localized = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function

        End If
    End Function 


    '-----------------------------------------------------------------
    'Procedure: AMO_GetDFTFileData_Base
    '@Purpose:  Return information so a Base Part Number DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetDFTFileData_Base( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID)
 
        'Response.Write(p_strBusSegIDs & "---" & p_strOwnerIDs & "--" & p_lngUserID & "--0")
        'Response.End()

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
	    oCommand.CommandText = "usp_AMOFeature_GetDFTFileData"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

		Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intLocalized", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = 0
		oCommand.Parameters.Append oParameter
                  
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
        If Err.Number > 0 Then			    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetDFTFileData_Base = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                                
            'return recordset
            Set AMO_GetDFTFileData_Base = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function

        End If
               
    End Function
            
            '-----------------------------------------------------------------
    'Procedure: AMO_GetDFTFileSuppliers
    '@Purpose:  Return information so a Base Part Number DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rsSupplier    | Supplier record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetDFTFileSuppliers_Base( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID)
 
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
	    oCommand.CommandText = "usp_AMOFeature_GetDFTFileSuppliers"  
    
        ' set database stored procedure command
     ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetDFTFileSuppliers_Base = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GetDFTFileSuppliers_Base = oRSelection
            ''Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SaveAutoload
    '@Purpose:  Save data DFT for AMO before Generate 2 txt files
    'Inputs:    @parm      | String        | p_strRepository    | database connection string
    '           @parm long | p_lngUserID   | UserID
    'Outputs:
    '           
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SaveAutoload(p_strRepository, _
                    p_strBusSegIDs, p_strOwnerIDs, _
                    p_intUserID, _
                    p_strParentPin, p_strRowLevel1, _
                    p_strRowLevel2, p_strODMCode, _
                    p_strFormatCode, p_strBrandName, _
                    p_strCountrification, p_strLCStatus, _
                    p_strMLORW, p_strSDFFlag, _
                    p_strLevel, _
                    p_strLCStatusLocal, p_strSDFFlagLocal, _
                    p_strLCStatusBaseWithLocal, p_strSDFFlagBaseWithLocal)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		'On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveAutoload"
		oCommand.CommandType = adCmdStoredProc
        
    
        Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
        oParameter.Value = p_strBusSegIDs
        oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
        oParameter.Value = p_strOwnerIDs
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
        oParameter.Value =  p_intUserID
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrParentPin", adVarChar, adParamInput, 16)
        oParameter.Value = p_strParentPin
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRowLevel1", adVarChar, adParamInput, 16)
        oParameter.Value = p_strRowLevel1
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrRowLevel2", adVarChar, adParamInput, 16)
        oParameter.Value = p_strRowLevel2
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrODMCode", adVarChar, adParamInput, 16)
        oParameter.Value = p_strODMCode
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrFormatCode", adVarChar, adParamInput, 16)
        oParameter.Value = p_strFormatCode
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrBrandName", adVarChar, adParamInput, 16)
        oParameter.Value = p_strBrandName
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrCountrification", adVarChar, adParamInput, 16)
        oParameter.Value = p_strCountrification
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLCStatus", adVarChar, adParamInput, 16)
        oParameter.Value = p_strLCStatus
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrMLORW", adVarChar, adParamInput, 16)
        oParameter.Value = p_strMLORW
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSDFFlag", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSDFFlag
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLevel", adVarChar, adParamInput, 16)
        oParameter.Value = p_strLevel
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLCStatusLocal", adVarChar, adParamInput, 16)
        oParameter.Value = p_strLCStatusLocal
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSDFFlagLocal", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSDFFlagLocal
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrLCStatusBaseWithLocal", adVarChar, adParamInput, 16)
        oParameter.Value = p_strLCStatusBaseWithLocal
        oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrSDFFlagBaseWithLocal", adVarChar, adParamInput, 16)
        oParameter.Value = p_strSDFFlagBaseWithLocal
        oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_SaveAutoload = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                	
			Exit Function
        Else
            AMO_SaveAutoload = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                	
			Exit Function
        End If
    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_GenerateAutoload
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_GenerateAutoload( _
                    p_strRepository, _
                    p_lngUserID)

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
	    oCommand.CommandText = "usp_AMOFeature_GenerateAutoload"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

            ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing

            'return empty recordset object and exit function
            Set AMO_GenerateAutoload = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GenerateAutoload = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
                
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_AnyAutoloadFileData
    '@Purpose:  Return information to determine if a Base and Localized DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_AnyAutoloadFileData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs)

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
	    oCommand.CommandText = "usp_AMOFeature_AnyAutoloadFileData"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter
  
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_AnyAutoloadFileData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_AnyAutoloadFileData = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_GetAutoloadFileData
    '@Purpose:  Return information so a Localized DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    '           @parm | Long       | p_lngLocalized     | get Localized or not, 0=no, 1=yes
    'Outputs:
    '           @parm ADODB.Recordset | p_rs            | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetAutoloadFileData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID, _
                    p_lngLocalized)

        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        'set database connection
        'oConnect.ConnectionString = PULSARDB() 
        'oConnect.Open
            
        ' Handle unexpected errors
        On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
        oCommand.CommandText = "usp_AMOFeature_GetAutoloadFileData"
            
		' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intLocalized", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngLocalized
		oCommand.Parameters.Append oParameter
           
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetAutoloadFileData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GetAutoloadFileData = oRSelection
            'Set Session("oConnect") = oConnect
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SaveDFTDescription
    '@Purpose:  Save data DFT for AMO before Generate 2 txt files
    'Inputs:    @parm      | String        | p_strRepository    | database connection string
    '           @parm long | p_lngUserID   | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_oRs           | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SaveDFTDescription(p_strRepository, _
                    p_strBusSegIDs, p_strOwnerIDs, _
                    p_intUserID, _
                    p_strUserName, _
                    p_strDESC_CD_Common, _
                    p_strDESC_CD_Quote, _
                    p_strM)
            
        Set oConnect = Server.CreateObject("ADODB.Connection")
		oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
       
		On Error Resume Next
			
        ' set database stored procedure command
		Set oCommand.ActiveConnection = oConnect
		oCommand.CommandText = "usp_AMOFeature_SaveDFTDescription"
		oCommand.CommandType = adCmdStoredProc


        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUserName", adVarChar, adParamInput, 64)
		oParameter.Value = p_strUserName
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDESC_CD_Common", adVarChar, adParamInput, 16)
		oParameter.Value = p_strDESC_CD_Common
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrDESC_CD_Quote", adVarChar, adParamInput, 16)
		oParameter.Value = p_strDESC_CD_Quote
		oCommand.Parameters.Append oParameter

            Set oParameter = oCommand.CreateParameter("@p_chrM", adVarChar, adParamInput, 16)
		oParameter.Value = p_strM
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_SaveDFTDescription = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                
			Exit Function
        Else
            AMO_SaveDFTDescription = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                
			Exit Function
        End If

    End Function


    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_GenerateDFTDescription
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_GenerateDFTDescription( _
                    p_strRepository, _
                    p_lngUserID)

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
	    oCommand.CommandText = "usp_AMOFeature_GenerateDFTDescription"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GenerateDFTDescription = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GenerateDFTDescription = oRSelection
            'Set Session("oConnect") = oConnect
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_AnyDFTDescriptionFileData
    '@Purpose:  Return information to determine if a DFT Description file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_AnyDFTDescriptionFileData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs)
 
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
	    oCommand.CommandText = "usp_AMOFeature_AnyDFTDescriptionFileData"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter
  
            ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_AnyDFTDescriptionFileData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_AnyDFTDescriptionFileData = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
                
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_GetDFTDescriptionData
    '@Purpose:  Return information so a Base Part Number DFT file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetDFTDescriptionData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID)
 
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
	    oCommand.CommandText = "usp_AMOFeature_GetDFTDescriptionData"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter
  
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetDFTDescriptionData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GetDFTDescriptionData = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
            
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_SaveBlindDate
    '@Purpose:  Save Blind Date data for AMO before Generate text file
    'Inputs:    @parm      | String        | p_strRepository    | database connection string
    '           @parm long | p_lngUserID   | UserID
    '
    'Outputs:   @parm ADODB.Recordset      | p_oRs              | record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_SaveBlindDate(p_strRepository, _
                    p_strBusSegIDs, p_strOwnerIDs, _
                    p_intUserID, _
                    p_strUserName, _
                    p_strOld_Eff_DT, _
                    p_strM)

        Set oConnect = Server.CreateObject("ADODB.Connection")
        oConnect.Open(PULSARDB())
		Set oCommand = Server.CreateObject("ADODB.Command")
            
            
        ' Handle unexpected errors
        On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = "usp_AMOFeature_SaveBlindDate"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_intUserID
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrUserName", adVarChar, adParamInput, 64)
		oParameter.Value = p_strUserName
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOld_Eff_DT", adVarChar, adParamInput, 16)
		oParameter.Value = p_strOld_Eff_DT
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrM", adVarChar, adParamInput, 16)
		oParameter.Value = p_strM
		oCommand.Parameters.Append oParameter
  
        ' execute the stored procedure
        oCommand.Execute 
            
        ' return recordset object
        If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            AMO_SaveBlindDate = Err.Number & " -- " & Err.Description& " -- " &  Err.Source 
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing
                
			Exit Function
        Else
            AMO_SaveBlindDate = True
                
            'Close database connection, oConnect.
			oConnect.close
			Set oConnect = Nothing

			Exit Function
        End If

    End Function

    '-------------------------------------------------------------------------------------
    '* Purpose		: AMO_GenerateBlindDate
    '* Inputs		: 
    '-------------------------------------------------------------------------------------
    Public Function AMO_GenerateBlindDate( _
                    p_strRepository, _
                    p_lngUserID)
     
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
	    oCommand.CommandText = "usp_AMOFeature_GenerateBlindDate"  
    
        ' set database stored procedure command
        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
		oCommand.Parameters.Append oParameter

             
        ' execute the stored procedure
        oRSelection.CursorType = adOpenForwardOnly
	    oRSelection.LockType = AdLockReadOnly
	    Set oRSelection = oCommand.Execute 
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GenerateBlindDate = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GenerateBlindDate = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function

        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_AnyBlindDateFileData
    '@Purpose:  Return information to determine if a Blind Date file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | UserID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_AnyBlindDateFileData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs)

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
	    oCommand.CommandText = "usp_AMOFeature_AnyBlindDateFileData"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
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
            Set AMO_AnyBlindDateFileData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_AnyBlindDateFileData = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_GetBlindDateData
    '@Purpose:  Return information so a Base Part Number Blind Date file can be created
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetBlindDateData( _
                    p_strRepository, _
                    p_strBusSegIDs, _
                    p_strOwnerIDs, _
                    p_lngUserID)
 
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
	    oCommand.CommandText = "usp_AMOFeature_GetBlindDateData"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrBusSegIDs", adVarChar, adParamInput, Len(p_strBusSegIDs) + 1)
		oParameter.Value = p_strBusSegIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrOwnerIDs", adVarChar, adParamInput, Len(p_strOwnerIDs) + 1)
		oParameter.Value = p_strOwnerIDs
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_intUserID", adInteger, adParamInput, cg_lngLEN_INT)
		oParameter.Value = p_lngUserID
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
            Set AMO_GetBlindDateData = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                
            'return recordset
            Set AMO_GetBlindDateData = oRSelection
            'Set Session("oConnect") = oConnect
            Exit Function
        End If
    End Function

    '-----------------------------------------------------------------
    'Procedure: AMO_GetPHWebCategory
    '@Purpose:  Return PHWeb Categories
    'Inputs:    @parm | String     | p_strRepository    | database connection string
    '           @parm | Long       | p_lngUserID        | User ID
    'Outputs:
    '           @parm ADODB.Recordset | p_rs          | AV record data
    '
    '@Returns:  A null object if no errors; otherwise, a collection of errors encountered
    '-----------------------------------------------------------------
    Public Function AMO_GetPHWebCategory( _
                    p_strRepository)
 
        Set oConnect = Session("oConnect")
		Set oCommand = Server.CreateObject("ADODB.Command")
        Set oRSelection = Server.CreateObject("ADODB.Recordset")
            
        ' set database connection
        'oConnect.Open(IRSDB(p_strRepository))
            
        ' Handle unexpected errors
        On Error Resume Next

        ' set database stored procedure command
        Set oCommand.ActiveConnection = oConnect
        oCommand.CommandType = adCmdStoredProc
	    oCommand.CommandText = "usp_ADMIN_AMOPHWebCategories"
 
        ' set stored procedure parameter(s)
		Set oParameter = oCommand.CreateParameter("@p_chrTable", adVarChar, adParamInput, 50)
		oParameter.Value = "amophweb_category_viewactive"
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrID", adVarChar, adParamInput, 1)
		oParameter.Value = ""
		oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_chrFieldValues", adVarChar, adParamInput, 1)
		oParameter.Value = ""
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrPersonName", adVarChar, adParamInput, 1)
		oParameter.Value = ""
		oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_chrPersonFullName", adVarChar, adParamInput, 1)
		oParameter.Value =""
		oCommand.Parameters.Append oParameter

        Set oParameter = oCommand.CreateParameter("@p_chrChildLinkID", adVarChar, adParamInput, 1)
		oParameter.Value = ""
		oCommand.Parameters.Append oParameter
            
        Set oParameter = oCommand.CreateParameter("@p_chrReturnID", adVarChar, adParamOutput, 1000)
		oCommand.Parameters.Append oParameter

        ' execute the stored procedure
        oRSelection.CursorLocation = adUseClient
        oRSelection.Open oCommand, , adOpenForwardOnly, adLockReadOnly
            
        ' return recordset object
            If Err.Number <> 0 Then		    'CUSTOM ERROR HANDLING.   
            'Close database connection, oConnect.
            oRSelection.Close
            Set oRSelection = Nothing
			    
            'oConnect.close
			'Set oConnect = Nothing
                
            'return empty recordset object and exit function
            Set AMO_GetPHWebCategory = Nothing
            Exit Function
        Else
            'disconnect the recordset
            Set oCommand.ActiveConnection = Nothing
            Set oCommand = Nothing
                 
            'return recordset
            Set AMO_GetPHWebCategor = oRSelection

            'Set Session("oConnect") = oConnect

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
