<%
Class ExcaliburSecurity

	Private m_IsSysAdmin
	Private m_CurrentUser
	Private m_CurrentUserFullName
	Private m_CurrentUserID
	Private m_CurrentUserDomain
	Private m_CurrentUserEmail
	Private m_CurrentPartnerID
	Private m_CommodityPM
	Private m_SCFactoryEngineer
	Private m_AccessoryPM
	Private m_EngCoordinator
	Private m_PulsarSystemAdmin
    Private m_IsProgramCoordinatorPermissions
    Private m_IsConfigurationManagerPermissions
    Private m_IsSEPMProductsPermissions
	
	Public Property Get IsPulsarSystemAdmin()
		if m_PulsarSystemAdmin = 1 then
			IsPulsarSystemAdmin = true
		else
			IsPulsarSystemAdmin = false
		end if
	End Property

	Public Property Get IsSysAdmin()
		IsSysAdmin = m_IsSysAdmin
	End Property
	
	Public Property Get IsProgramCoordinatorPermissions()
		if m_IsProgramCoordinatorPermissions > 0 then
			IsProgramCoordinatorPermissions = true
		else
			IsProgramCoordinatorPermissions = false
		end if
	End Property

	Public Property Get IsConfigurationManagerPermissions()
		if m_IsConfigurationManagerPermissions > 0 then
			IsConfigurationManagerPermissions = true
		else
			IsConfigurationManagerPermissions = false
		end if
	End Property

    Public Property Get IsSEPMProductsPermissions()
		if m_IsSEPMProductsPermissions > 0 then
			IsSEPMProductsPermissions = true
		else
			IsSEPMProductsPermissions = false
		end if
	End Property

	Public Property Get CurrentUser()
		CurrentUser = m_CurrentUser
	End Property
	
	Public Property Get CurrentUserFullName()
		CurrentUserFullName = m_CurrentUserFullName
	End Property
	
	Public Property Get CurrentUserID()
		CurrentUserID = m_CurrentUserID
	End Property
	
	Public Property Let CurrentUserID( value )
		m_CurrentUserID = value
	End Property
	
	Public Property Get CurrentUserDomain()
		CurrentUserDomain = m_CurrentUserDomain
	End Property
	
	Public Property Get CurrentUserEmail()
		CurrentUserEmail = m_CurrentUserEmail
	End Property
	
	Public Property Get CurrentPartnerID()
		CurrentPartnerID = m_CurrentPartnerID
	End Property

    Public Property Get IsCommodityPM()
        IsCommodityPM = m_CommodityPM
    End Property
    
    Public Property Get IsSCFactoryEngineer()
        IsSCFactoryEngineer = m_SCFactoryEngineer
    End Property
    
    Public Property Get IsAccessoryPM()
        IsAccessoryPM = m_AccessoryPM
    End Property
    
    Public Property Get IsEngineeringCoordinator()
        IsEngineeringCoordinator = m_EngCoordinator
    End Property

	Private Sub Class_Initialize()
		m_CurrentUser = lcase(Session("LoggedInUser"))

		if InStr(m_CurrentUser,"\") > 0 then
			m_CurrentUserDomain = Left(m_CurrentUser, instr(m_CurrentUser,"\") - 1)
			m_CurrentUser = mid(m_CurrentUser,instr(m_CurrentUser,"\") + 1)
		end if
		
		Dim dw, cn, cmd, rs
		
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spGetUserInfo")
		dw.CreateParameter cmd, "@UserName", adVarChar, adParamInput, 80, m_CurrentUser
		dw.CreateParameter cmd, "@Domain", adVarChar, adParamInput, 30, m_CurrentUserDomain
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If Not (rs.EOF And rs.BOF) Then
			m_CurrentUserID = rs("ID") & ""
			m_CurrentUserEmail = rs("email") & ""
			m_CurrentPartnerID = rs("partnerid") & ""
			m_CurrentUserFullName = rs("name") & ""
			m_CommodityPM = rs("CommodityPM")
			m_SCFactoryEngineer = rs("SCFactoryEngineer")
			m_AccessoryPM = rs("AccessoryPM")
			m_EngCoordinator = rs("EngCoordinator")
			m_PulsarSystemAdmin = rs("PulsarSystemAdmin")
            m_IsProgramCoordinatorPermissions = rs("PCProductCount")
            m_IsConfigurationManagerPermissions = rs("CMProductCount")
            m_IsSEPMProductsPermissions  = rs("SEPMProducts")
		Else
			m_CurrentUserID = 0
		End If
		rs.Close

		m_IsSysAdmin = IsSystemAdmin()
		
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing
	End Sub
	
	Private Sub Class_Terminate()
	End Sub
	
	Public Function IsHardwarePm(ProductVersionId)
		Dim dw, cn, cmd, rs
		Dim bPlatFormDevelopmentPm, bProcessorPm, bCommPm, bGraphicsControllerPm, bVideoMemoryPM, bSuperUser

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spGetHardwareTeamAccessList")
		dw.CreateParameter cmd, "@EmployeeID", adInteger, adParamInput, 30, m_CurrentUserID
		dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 30, CLng(Trim(ProductVersionId))
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
	
		do while not rs.EOF
			if rs("HWTeam") = "ProgramCoordinator" or m_EngCoordinator > 0 then
				bPlatFormDevelopmentPm = true
			elseif rs("HWTeam") = "PlatformDevelopment" and rs("Products") > 0 then
				bPlatFormDevelopmentPm = true
			elseif rs("HWTeam") = "Processor" and rs("Products") > 0 then
				bProcessorPm = true
			elseif rs("HWTeam") = "Comm" and rs("Products") > 0 then
				bCommPm = true
			elseif rs("HWTeam") = "GraphicsController" and rs("Products") > 0 then
				bGraphicsControllerPm = true
			elseif rs("HWTeam") = "VideoMemory" and rs("Products") > 0 then
				bVideoMemoryPM = true
			elseif rs("HWTeam") = "SuperUser" and rs("Products") > 0 then
				bSuperUser = true
			'elseif rs("HWTeam") = "Commodity" and rs("Products") > 0 then
			'	blnCommodityPM = true
			end if
			rs.MoveNext			
		loop
		rs.Close

		if m_CommodityPm or  bCommPM or bProcessorPM or bVideoMemoryPM or bGraphicsControllerPM then
			IsHardwarePm = true
		else
			IsHardwarePm = false
		end if
	End Function
	
	Public Function IsDeliverableOwner(DeliverableVersionID, DeliverableRootID)
		If Len(Trim(DeliverableVersionID)) = 0 And Len(Trim(DeliverableRootID)) = 0 And (Not IsNumeric(DeliverableRootID)) Then
			
			IsDeliverableOwner = False
			Exit Function
		End If

		Dim dw, cn, cmd, rs

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectDeliverableVersion")
		dw.CreateParameter cmd, "@p_DeliverableVersionID", adInteger, adParamInput, 30, DeliverableVersionID
		dw.CreateParameter cmd, "@p_DeliverableRootID", adInteger, adParamInput, 30, CLng(Trim(DeliverableRootID))
		dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 30, CLng(Trim(m_CurrentUserID))
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof And rs.bof Then
			IsDeliverableOwner = False
		Else
			IsDeliverableOwner = True
		End If
		
		rs.Close

		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsSystemTeamLead(ProductVersionID)
	
		If Len(Trim(ProductVersionID)) = 0 Then
			IsSystemTeamLead = False
			Exit Function
		End If

		Dim dw, cn, cmd, rs
	
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductVersion")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 30, CLng(Trim(ProductVersionID))
		dw.CreateParameter cmd, "@p_PMID", adInteger, adParamInput, 30, NULL
		dw.CreateParameter cmd, "@p_SEPMID", adInteger, adParamInput, 30, NULL
		dw.CreateParameter cmd, "@p_STLID", adInteger, adParamInput, 30, CLng(m_CurrentUserID)
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof And rs.bof Then
			IsSystemTeamLead = False
		Else
			IsSystemTeamLead = True
		End If
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsTestLead()
	
		Dim dw, cn, cmd, rs
		Dim blnOdmTestLead, blnWwanTestLead, blnSeTestLead
	
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spGetTestLeadsAll")
		dw.CreateParameter cmd, "@SE", adBoolean, adParamInput, 30, 1
		dw.CreateParameter cmd, "@WWAN", adBoolean, adParamInput, 30, 1
		dw.CreateParameter cmd, "@ODM", adBoolean, adParamInput, 30, 1
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        
	    Do While Not rs.EOF
		    If trim(m_CurrentUserID) = Trim(rs("ID")) Then
			    If rs("role") = "ODM Test Lead"  Then
				    blnODMTestLead = 1
			    ElseIf rs("role") = "WWAN Test Lead"  Then
				    blnWWANTestLead = 1
			    ElseIf rs("role") = "SE Test Lead" Then
				    blnSETestLead = 1
				    If rs("PartnerID")  = 1 Then
					    blnODMTestLead = 1
				    End If
			    End If
		    End If		
		    rs.MoveNext
	    Loop

	    rs.Close

		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	    If blnODMTestLead =1 Or blnWWANTestLead =1 Then 'blnSETestLead or 
		    IsTestLead = true
		Else
		    IsTestLead = false
	    End If

	End Function
	
	Public Function IsToolsPm()
	
		Dim dw, cn, cmd, rs
		Dim blnToolsPM
	
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spListToolsPMs")
        Set rs = dw.ExecuteCommandReturnRs(cmd)
	
	    blnToolsPM = False
		do while not rs.EOF
			if trim(m_CurrentUserID) = trim(rs("ID")) then
				blnToolsPM = true
				exit do
			end if
			rs.MoveNext
		loop
		rs.Close	

        IsToolsPm = blnToolsPM

    End Function

	Public Function IsProgramManager(ProductVersionID)
	
		If Len(Trim(ProductVersionID)) = 0 Then
			IsProgramManager = False
			Exit Function
		End If

		Dim dw, cn, cmd, rs
	
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductVersion")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 30, CLng(Trim(ProductVersionID))
		dw.CreateParameter cmd, "@p_PMID", adInteger, adParamInput, 30, CLng(m_CurrentUserID)
		dw.CreateParameter cmd, "@p_SEPMID", adInteger, adParamInput, 30, NULL
		dw.CreateParameter cmd, "@p_STLID", adInteger, adParamInput, 30, NULL
		dw.CreateParameter cmd, "@p_PCID", adInteger, adParamInput, 30, NULL
		dw.CreateParameter cmd, "@p_PlatformDevID", adInteger, adParamInput, 30, NULL
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof And rs.bof Then
			IsProgramManager = False
		Else
			IsProgramManager = True
		End If
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsSysEngProgramManager(ProductVersionID)
	
		If Len(Trim(ProductVersionID)) = 0 Then
			IsSysEngProgramManager = False
			Exit Function
		End If

        Dim isSEPM
		Dim dw, cn, cmd, rs
	
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		'Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductVersion")
		'dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 30, CLng(Trim(ProductVersionID))
		'dw.CreateParameter cmd, "@p_PMID", adInteger, adParamInput, 8, null
		'dw.CreateParameter cmd, "@p_SEPMID", adInteger, adParamInput, 8, CLng(Trim(m_CurrentUserID))
		'dw.CreateParameter cmd, "@p_STLID", adInteger, adParamInput, 30, NULL
		'Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		Set cmd = dw.CreateCommandSP(cn, "spListPMsActive")
		dw.CreateParameter cmd, "@TypeID", adInteger, adParamInput, 0, 1
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		isSEPM = false

        Do Until rs.EOF
            If rs("ID") & "" = m_CurrentUserID Then isSEPM = true
            rs.MoveNext
        Loop

		IsSysEngProgramManager = isSEPM
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function

	Private Function IsSystemAdmin()

		Dim dw, cn, cmd, rs

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectEmployees")
		dw.CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 4, CLng(Trim(m_CurrentUserID))
		dw.CreateParameter cmd, "@p_IsAdmin", adBoolean, adParamInput, 1, 1 
		dw.CreateParameter cmd, "@p_NTName", adVarChar, adParamInput, 30, ""
		dw.CreateParameter cmd, "@p_Domain", adVarChar, adParamInput, 30, ""
		dw.CreateParameter cmd, "@p_PartnerID", adInteger, adParamInput, 4, ""
		
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		If rs.EOF and rs.BOF then
			IsSystemAdmin = false
		'ElseIf rs("ImpersonateID")&"" <> "" then
        '    IsSystemAdmin = false
        Else
			IsSystemAdmin = true
		end if
		
		rs.Close

		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing
		
	End Function

	Public Function IsProgramCoordinator(ProductVersionID)
	
		If Len(Trim(ProductVersionID)) = 0 Then
			IsProgramCoordinator = False
			Exit Function
		End If

		Dim dw, cn, cmd, rs

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
        Set cmd = dw.CreateCommandSP(cn, "spGetProductVersion")
        dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, CLng(ProductVersionID)
        Set rs = dw.ExecuteCommandReturnRS(cmd)

        If rs.eof And rs.bof Then
			IsProgramCoordinator = False
			Exit Function
		End If

        Dim BusinessID
        Dim FusionRequirement : FusionRequirement = 0
        BusinessID = rs("BusinessID") & ""
        if (rs("FusionRequirements")) then
            FusionRequirement = 1
        else
            FusionRequirement = 0  
        end if  
      
        rs.close

        if (FusionRequirement = 1) then
        	Set cmd = dw.CreateCommandSP(cn, "usp_GetProgramCoordinatorStatus")
		    dw.CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 8, CLng(m_CurrentUserID)
            dw.CreateParameter cmd, "@p_BusinessID", adInteger, adParamInput, 8, null
            Set rs = dw.ExecuteCommandReturnRS(cmd)	
       else
            Set cmd = dw.CreateCommandSP(cn, "usp_GetProgramCoordinatorStatus")
		    dw.CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 8, CLng(m_CurrentUserID)
		    dw.CreateParameter cmd, "@p_BusinessID", adInteger, adParamInput, 8, CLng(BusinessID)
            Set rs = dw.ExecuteCommandReturnRS(cmd)	
        end if
		   
		If rs.eof And rs.bof Then
			IsProgramCoordinator = False
		Else
			IsProgramCoordinator = True
		End If
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing
		
	End Function    
	
	Public Function IsPlatformDevMgr(ProductVersionID)

		If Len(Trim(ProductVersionID)) = 0 Then
			IsPlatformDevMgr = False
			Exit Function
		End If

		Dim dw, cn, cmd, rs

		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_SelectProductVersion")
		dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, CLng(Trim(ProductVersionID))
		dw.CreateParameter cmd, "@p_PMID", adInteger, adParamInput, 8, NULL
		dw.CreateParameter cmd, "@p_SEPMID", adInteger, adParamInput, 8, NULL
		dw.CreateParameter cmd, "@p_STLID", adInteger, adParamInput, 8, NULL
		dw.CreateParameter cmd, "@p_PCID", adInteger, adParamInput, 8, NULL
        dw.CreateParameter cmd, "@p_PlatformDevID", adInteger, adParamInput, 8, CLNG(Trim(m_CurrentUserID))
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof And rs.bof Then
			IsPlatformDevMgr = False
		Else
			IsPlatformDevMgr = True
		End If
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsSpdm()
	    Dim dw, cn, cmd, rs
	    
	    Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_GetSpdmAccess")
		dw.CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 8, CLng(m_CurrentUserID)
		Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		If rs.eof And rs.bof Then
			IsSpdm = False
		Else
			IsSpdm = True
		End If
		
		rs.Close
	
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsServiceBomAnalyst()
	    Dim dw, cn, cmd, rs
	    
	    Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_GetServiceBomAnalystAccess")
		dw.CreateParameter cmd, "@p_EmployeeID", adInteger, adParamInput, 8, CLng(m_CurrentUserID)
		Set rs = dw.ExecuteCommandReturnRS(cmd)
		
		If rs.Eof And rs.Bof Then
		    IsBomAnalyst = False
		Else
		    IsBomAnalyst = True
		End If
		
		rs.Close
		
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
	
	Public Function IsActivePm()
	    Dim dw, cn, cmd, rs
	    Dim blnEditProductProperties
	    Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "spListPMsActive")
		dw.CreateParameter cmd, "@TypeId", adInteger, adParamInput, 8, 3
		Set rs = dw.ExecuteCommandReturnRS(cmd)	

        blnEditProductProperties = false
        do while not rs.EOF
		    if trim(currentuserid) =  trim(rs("ID")) then
           	    blnEditProductProperties = true
                exit do
            end if
            rs.MoveNext				
        loop
        rs.Close	
        
        IsActivePm = blnEditProductProperties
        
	End Function
	
	Public Function UserInRole(ProductVersionID, Role)
        UserInRole = False

		Dim dw, cn, cmd, rs
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		
        Set cmd = dw.CreateCommandSP(cn, "usp_GetUserInRole")
	    dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 0, m_CurrentUserID
	    dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, ProductVersionId
	    dw.CreateParameter cmd, "@p_RoleCd", adVarChar, adParamInput, 20, Role
        Set rs = dw.ExecuteCommandReturnRS(cmd)
        If Not (rs.EOF And rs.BOF) Then
            UserInRole = (CLng(rs("UserInRole")) > 0)
        End If
        
	

        If Not UserInRole Then
        
		    Set cmd = dw.CreateCommandSP(cn, "usp_ListUserInRoles")
		    dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 0, CLng(m_CurrentUserID)
		    Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
		    If rs.eof And rs.bof Then
		    Else
			    Select Case UCase(Role)
				    Case "STL"
					    If CLng(rs("SMID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "POPM"
					    If CLng(rs("PMID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SEPM"
					    If CLng(rs("SEPMID")) = CLng(m_CurrentUserID) Then UserInRole = True
			        Case "CM"
			            If CLng(rs("PMID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "COMMERCIALMARKETING"
					    If CLng(rs("ComMarketingID")&"") = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "CONSUMERMARKETING"
					    If CLng(rs("ConsMarketingID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SMBMARKETING"
					    If CLng(rs("SmbMarketingID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "PLATFORMDEVELOPMENT"
					    If CLng(rs("PlatformDevelopmentID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SUPPLYCHAIN"
					    If CLng(rs("SupplyChainID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SERVICE"
					    If CLng(rs("ServiceID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "FINANCE"
					    If CLng(rs("FinanceID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "COMMODITYPM"
					    If CLng(rs("PDEID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "ACCESSORYPM"
					    If CLng(rs("AccessoryPMID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "MARKETINGOPS"
					    If CLng(rs("MarketingOpsID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "PC"
					    If CLng(rs("PCID")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SEPE"
					    If CLng(rs("SEPE")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "PINPM"
					    If CLng(rs("PINPM")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "SETL"
					    If CLng(rs("SETestLead")) = CLng(m_CurrentUserID) Then UserInRole = True
				    Case "GPLM"
				        If rs("GPLM")&"" <> "" Then
    				        If CLng(rs("GPLM")) = CLng(m_CurrentUserId) Then UserInRole = True
    				    End If
				    Case "SBA"
				        If rs("SvcBomAnalyst")&"" <> "" Then
    				        If CLng(rs("SvcBomAnalyst")) = CLng(m_CurrentUserId) Then UserInRole = True
    				    End If
    				    If Not UserInRole AND rs("SPDM")&"" <> "" Then
    				        If CLng(rs("SPDM")) = CLng(m_CurrentUserId) Then UserInRole = True
    				    End If
			    End Select
		    End If
    		
		    rs.Close
	    End If
	    
		Set rs = nothing
		Set cmd = nothing
		Set cn = nothing
		Set dw = nothing

	End Function
End Class
%>