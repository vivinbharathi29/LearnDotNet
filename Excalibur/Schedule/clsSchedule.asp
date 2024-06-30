<%
Class Schedule

    Public Function UpdateMilestone(Connection, PorEnd, Item, FullUserName, UserEmail, Comments, ItemNotes, ProjectedStartOld, ProjectedEndOld, ProjectedStart, ProjectedEnd, ActualStart, ActualEnd, ShowOnReports, ItemPhase, ItemDefinition, ItemOwner)
	    
	    Dim dw
	    Dim cn
	    Dim cmd
	    Dim iRowsChanged
	    Dim bFoundErrors
        Dim bNotifyOfPddLock
        Dim rs1, rs2
        Dim sSubject
        Dim sBody
        Dim oMessage
        Dim IsMilestone

	    Set dw = New DataWrapper
        Set cn = Connection
        
	    bFoundErrors = False

		Set cmd = dw.CreateCommandSP(cn, "usp_UpdateScheduleData")
        dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, Item
		dw.CreateParameter cmd, "@p_LastUpdUser", adVarChar, adParamInput, 50, Trim(FullUserName)
		dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, NULL
		dw.CreateParameter cmd, "@p_ItemNotes", adVarChar, adParamInput, 5000, Comments
		dw.CreateParameter cmd, "@p_ChangeNotes", adVarChar, adParamInput, 5000, ItemNotes
		dw.CreateParameter cmd, "@p_PorStartDt", adDate, adParamInput, 0, NULL
		dw.CreateParameter cmd, "@p_PorEndDt", adDate, adParamInput, 0, NULL
		dw.CreateParameter cmd, "@p_ProjectedStartDt", adDate, adParamInput, 0, ProjectedStart
		dw.CreateParameter cmd, "@p_ProjectedEndDt", adDate, adParamInput, 0, ProjectedEnd
		dw.CreateParameter cmd, "@p_ActualStartDt", adDate, adParamInput, 0, ActualStart
		dw.CreateParameter cmd, "@p_ActualEndDt", adDate, adParamInput, 0, ActualEnd
		dw.CreateParameter cmd, "@p_ShowOnReports_YN", adChar, adParamInput, 1, ShowOnReports
		dw.CreateParameter cmd, "@p_ItemPhase", adInteger, adParamInput, 8, ItemPhase
		dw.CreateParameter cmd, "@p_ItemOwner", adInteger, adParamInput, 8, ItemOwner
		Set rs1 = dw.ExecuteCommandReturnRS(cmd)
        'Set rs1 = server.CreateObject("ADODB.RecordSet")
        
        'Set cmd = server.CreateObject("ADODB.Command")
        
        response.Write ItemDefinition & "<br>"
        response.Write PorEnd & "<br>"
        response.Write ActualEnd & "<br>"

        Response.Write ItemDefinition = "7" And Trim(PorEnd) = "" And Trim(ActualEnd) <> "" & "<BR>"
        '
	    ' If PDD Locked Milestone Actual Date Set then Set the PDDLocked bit in the Product_Release Table
	    '
        If ItemDefinition = "7" And Trim(PorEnd) = "" And Trim(ActualEnd) <> "" Then
			
			response.Write "Set PDD Locked<br>"
				
			Set cmd = dw.CreateCommandSP(cn, "usp_SetPddLockedBit")
			dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, Item
			iRowsChanged = dw.ExecuteNonQuery(cmd)
				
		    Set cmd = dw.CreateCommandSP(cn, "usp_SelectScheduleData")
		    dw.CreateParameter cmd, "@p_ScheduleDataID", adInteger, adParamInput, 8, NULL
		    dw.CreateParameter cmd, "@p_ScheduleID", adInteger, adParamInput, 8, Request.Form("hidScheduleID")
		    dw.CreateParameter cmd, "@p_ReleaseID", adInteger, adParamInput, 8, NULL
		    dw.CreateParameter cmd, "@p_PhaseID", adInteger, adParamInput, 8, NULL
		    dw.CreateParameter cmd, "@p_Active_YN", adChar, adParamInput, 1, NULL
		    dw.CreateParameter cmd, "@p_ScheduleDefinitionDataID", adInteger, adParamInput, 8, 60
		    Set rs2 = dw.ExecuteCommandReturnRS(cmd)
    		
		    If Not rs2.EOF Then
			    '
			    ' Send email to the MobileExcalNotification-FCS@hp.com list.      
			    '
			    
			    If rs2("projected_end_dt")&"" = "" Then
			        Response.Clear
			        Response.Write "<h2>The MV Production Release and Mass Production First Customer Ship Dates must be set before the POR is complete.</h2>"
			        cn.RollbackTrans
			        Response.End
                End If			        
			    
			    sSubject = Request("hidProgram") & " has reached PRL/PDD Lock"
			    sBody = "The " & Request("hidProgram") & " schedule has locked." & vbcrlf & _
					    vbcrlf & _
					    "MV Production Release POR Date: " & FormatDateTime(rs2("projected_end_dt")&"", vbShortDate)
    		
			    'Set oMessage = Server.CreateObject("CDO.Message")
			    Set oMessage = New EmailWrapper
			    'Set oMessage.Configuration = Application("CDO_Config")
    		
			    oMessage.To = "MobileExcalNotification-FCS@hp.com"
			    oMessage.From = UserEmail
			    oMessage.Subject = sSubject
			    oMessage.TextBody = sBody
			    oMessage.DSNOptions = 2 'cdoDSNFailure
			    
			    if Request("PVID") = "100" Then
			        oMessage.To = UserEmail
			        oMessage.TextBody = "Email Is Inactive this message would have gone to MobileExcalNotification-FCS@hp.com" & vbCrLf & vbCrLf & sBody
			    End If
			    
			    oMessage.Send
			    Set oMessage = Nothing 		
    					
		    End If
		    rs2.close
		    set rs2 = nothing
		End If
		
		'Schedule Change Alert
		If (Not rs1.EOF) And (ProjectedStartOld <> ProjectedStart Or ProjectedEndOld <> ProjectedEnd) Then

            sSubject = "Schedule Change Alert: " & rs1("name") & " " & rs1("version") & " - " & rs1("item_description")
		    
		    sBody = "<P class=tabletitle>The " & rs1("item_description") & " Date on the " & rs1("name") & " " & rs1("version") & " schedule has changed.</P>"
		    
		    If Ucase(rs1("Milestone_YN")) = "Y" Then
		        IsMilestone = True
		    Else 
		        IsMilestone = False
		    End If
		    
		    If IsMilestone Then
		        	sBody = sBody & "<TABLE cellSpacing=1 cellPadding=3 border=1 borderColor=tan bgColor=ivory>"
		        	sBody = sBody & "<THEAD><TR><TH width='50%' bgColor=cornsilk>Old Plan</TH><TH width='50%' bgColor=cornsilk>New Plan</TH></TR></THEAD>"
		        	sBody = sBody & "<TBODY><TR><TD>" & ProjectedStartOld & "</TD><TD>" & ProjectedStart & "</TD></TR></TBODY></TABLE>"
            Else
            	sBody = sBody & "<TABLE cellSpacing=1 cellPadding=3 border=1 borderColor=tan bgColor=ivory>"
              sBody = sBody & "<THEAD><TR><TH colspan=2 width='50%' bgColor=cornsilk>Old Plan</TH><TH colspan=2 width='50%' bgColor=cornsilk>New Plan</TH></TR>"
              sBody = sBody & "<TR><TH width='25%' bgColor=cornsilk>Start</TH><TH width='25%' bgColor=cornsilk>End</TH><TH width='25%' bgColor=cornsilk>Start</TH><TH width='25%' bgColor=cornsilk>End</TH></TR></THEAD>"
              sBody = sBody & "<TBODY><TR><TD>" & ProjectedStartOld & "</TD><TD>" & ProjectedEndOld & "</TD><TD>" & ProjectedStart & "</TD><TD>" & ProjectedEnd & "</TD></TR></TBODY></TABLE>"
            End If
		    
'		    "The " & Request("hidMilestoneName") & " date on the " & Request("hidProgram") & " schedule has changed." & vbcrlf & _
'				vbcrlf & _
'				"POR Date: " & FormatDateTime(Request("hidPorEndDt"), vbShortDate) & vbcrlf & _
'				"Old Date: " & FormatDateTime(Request("hidProjectedEndDt"), vbShortDate) & vbcrlf & _
'				"New Date: " & FormatDateTime(Request("txtProjectedEndDt"), vbShortDate)
		
		    'Set oMessage = Server.CreateObject("CDO.Message")
		    Set oMessage = New EmailWrapper
		    'Set oMessage.Configuration = Application("CDO_Config")
		
		    oMessage.To = rs1("NotificationRecipients")
		    'oMessage.To = "kenneth.berntsen@hp.com" 'rs1("NotificationRecipients")
		    oMessage.From = UserEmail
		    oMessage.Subject = sSubject
		    oMessage.HTMLBody = sBody
		    oMessage.DSNOptions = 2 'cdoDSNFailure
			    if Request("PVID") = "100" Then
			        oMessage.HTMLBody = "Email Is Inactive this message would have gone to " & rs1("NotificationRecipients") & "<br><br>" & sBody
			        oMessage.To = UserEmail
			    End If

		    oMessage.Send
		    Set oMessage = Nothing 

		End If

    End Function
End Class
%>
