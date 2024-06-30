<%@ Language=VBScript %>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/emailwrapper.asp" -->
<%
	dim strVersionID
	dim strProductID
	dim strType
	dim strRejected
	dim cn
	dim rs
    dim strArray
	dim IDArray
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
    dim CurrentUserName
    dim CurrentUserEmail
	dim cm
	dim TargetArray
    dim strTo
    dim strDeliverableName
    dim strProductName
    dim strSubject
    dim strBody
    dim strCC
    dim strEnvironment

	if request("txtMultiID") = "" then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0
	else       

  
		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open


		'Get User
		CurrentUser = lcase(Session("LoggedInUser"))

		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetUserInfo"
	
		Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		p.Value = Currentuser
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		p.Value = CurrentDomain
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
	
		set cm=nothing
	
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID") 
            CurrentUserName = rs("Name") & ""
            CurrentUserEmail = rs("Email") & ""
		end if
		rs.Close
	
		Response.Write "<BR>UserID:" & CurrentUserID & "<BR>"

		cn.BeginTrans
		
		ProcessArray = split(request("txtMultiID"),",")
		
		Response.Write "<BR>Processing:" & "<BR>"
		
       'Get Server Environment
        rs.open "usp_GetServerEnvironment"
        if not (rs.EOF and rs.BOF) then
            strEnvironment = rs("Environment") & ""
		end if
        rs.Close

		for i = 0 to ubound(ProcessArray)
            strTO = ""
            strCC = ""
            strSubject = ""
            strBody = ""
			
            strID = trim(ProcessArray(i))		
			Response.Write strID & " - " & clng(request("NewValue")) & "<BR>"
			strArray = split(strID,"_")
            IDArray = split(strArray(1),":")

            'get Email Data first
            if clng(IDArray(0)) = 1 and clng(request("NewValue")) = 2 then 'Rejecting Root
                rs.open "spGetDeliverableRootProductPM " & clng(IDArray(0)),cn
                if rs.eof and rs.bof then
                    strTo = ""
                    strCC = ""
                    strDeliverableName = ""
                    strProductName = ""
                else
                    strTo = trim(rs("Email") & "")
                    if trim(rs("TypeID")) = "1" then
                        strCC = trim(rs("SystemTeamEmail") & "")
                    else
                        strCC = ""
                    end if
                    strDeliverableName = trim(rs("Deliverablename") & "")
                    strProductName = trim(rs("Productname") & "")
                    strSubject = "Deliverable Root Support Request Rejected"
                    strBody = "The request to support " & strDeliverableName & " on " & strProductName & " was rejected.<BR><BR>"
                    if request("txtComments") <> "" then
                        strBody = strBody & "Reason Rejected: " & request("txtComments") & "<BR><BR>"
                    end if

                end if
               rs.Close
            elseif clng(request("NewValue")) = 2 then 'Rejecting Version
                rs.open "spGetDeliverableVersionProductPM " & clng(IDArray(0)),cn
                if rs.eof and rs.bof then
                    strTo = ""
                    strCC = ""
                    strDeliverableName = ""
                    strProductName = ""
                else
                    strTo = trim(rs("Email") & "")
                    strCC = trim(rs("SystemTeamEmail") & "")
                    strDeliverableName = trim(rs("Deliverablename") & "")
                    strProductName = trim(rs("Productname") & "")
                    strSubject = "Deliverable Version Support Request Rejected"
                    strBody = "The request to support this version of " & strDeliverableName & " on " & strProductName & " was rejected." & "<BR><BR>"
                    if request("txtComments") <> "" then
                       strBody = strBody & "Reason Rejected: " & request("txtComments") & "<BR><BR>"
                    end if
                    strBody = strBody & "Vendor: " & rs("Vendor") & "<BR>"
                    strBody = strBody & "Hardware Version: " & rs("Version") & "<BR>"
                    strBody = strBody & "Firmware Version: " & rs("Revision") & "<BR>"
                    strBody = strBody & "Revision: " & rs("Pass") & "<BR>"
                    strBody = strBody & "Part Number: " & rs("PartNumber") & "<BR>"
                    strBody = strBody & "Model Number: " & rs("ModelNumber") & "<BR>"
                end if
               rs.Close
            end if

            'Process update

			set cm = server.CreateObject("ADODB.Command")		
			cm.ActiveConnection = cn

            if clng(IDArray(1)) = 0 then 
			    cm.CommandText = "spUpdateDeveloperNotificationStatus"
            	cm.CommandType = &H0004
						
			    Set p = cm.CreateParameter("@PDID",3, &H0001)
			    p.Value = clng(IDArray(0))
			    cm.Parameters.Append p
            else 
                cm.CommandText = "spUpdateDeveloperNotificationStatusPulsar"
            	cm.CommandType = &H0004
						
                Set p = cm.CreateParameter("@PDID",3, &H0001)
			    p.Value = clng(IDArray(0))
			    cm.Parameters.Append p

                Set p = cm.CreateParameter("@PDRID",3, &H0001)
			    p.Value = clng(IDArray(1))
			    cm.Parameters.Append p
            end if
			
			Set p = cm.CreateParameter("@Status",16, &H0001)
			p.Value = clng(request("NewValue"))
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@DeveloperTestNotes", 200, &H0001, 256)
			p.Value = request("txtComments")
			cm.Parameters.Append p

			Set p = cm.CreateParameter("@Type",16, &H0001)
			p.Value = clng(strArray(0))
			cm.Parameters.Append p
			
			Set p = cm.CreateParameter("@UserID",3, &H0001)
			p.Value = clng(CurrentUserID)
			cm.Parameters.Append p
					
			cm.Execute
			Set cm = Nothing
	
			if cn.Errors.count > 0 then
				Response.Write "<BR>Failed.<BR>"
				strSuccess = "0"
				cn.RollbackTrans
				Response.Write "<BR>Records were not saved correctly.<BR>"
				exit for
			else
				strSuccess = "1"

                if strEnvironment <> "1" Then
                  strCC = CurrentUserEmail  
                  strTO = CurrentUserEmail 
                else
                  strCC = CurrentUserEmail & ";" & strCC 
                end if


                'Notify PM
                if (trim(strTo) <> "" or trim(strCC) <> "") and clng(request("NewValue")) =2 then
                    response.write "<BR>Preparing Email"
                    if strDeliverableName = "" or strproductName = "" then
                        strBody = "ProdDelRootID: " &  clng(IDArray(0)) & "<BR>"
                        strBody = strBody & "To: " &  strTo & "<BR>"
                        strBody = strBody & "Product: " & strProductName & "<BR>"
                        strBody = strBody & "Root: " &  strDeliverableName & "<BR>"

                    else
                        if strTo="" then
                            strTO = CurrentUserEmail
                        end if
                        if lcase(trim(strProductname)) = "test product 1.0" then
                            strBody = "This update occured on Test Product 1.0.  The following notifications would have been sent to <BR>TO: " & strTo & "<BR>CC: " & strCC & ".<BR><BR>"  & strBody

                        end if
                    end if

                    response.write "<BR>Email Ready: " & strTO & "_" & strCC & "_" & strBody
                    if (strTo <> "" or strCC <> "" )  then
                        response.write "<BR>Sending Email"

                        Set oMessage = New EmailWrapper 
	                    oMessage.From = CurrentUserEmail
		
	                    if strTo="" then
		                    oMessage.To = CurrentUserEmail
               		    else
		                    oMessage.To = strTo 
	                    end if
                        
                        oMessage.CC = strCC
            
	                    oMessage.Subject = strSubject
	                    oMessage.HTMLBody = "<font size=2 face=verdana color=black>" & strBody & "</font>"
	                    oMessage.DSNOptions = cdoDSNFailure
	                    oMessage.Send 
	                    Set oMessage = Nothing 			
                    end if

                end if

			end if
		next
		if strSuccess = "1" then
			Response.Write "<BR>Committing<BR>"
			cn.CommitTrans
			Response.Write "<BR>Records saved successfully.<BR>"
		end if
					            
		set rs = nothing
		set cn = nothing

	end if


%>
