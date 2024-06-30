<%@  language="VBScript" %>
<%Response.Buffer = True %>
<!-- #include file="../../includes/EmailQueue.asp" -->
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">

    <script id="clientEventHandlersJS" language="javascript">
<!--
function window_onload() {
		window.close();
}

//-->
    </script>

</head>
<style>
A:link
{
    COLOR: blue;
}
A:visited
{
    COLOR: blue;
}
A:hover
{
    COLOR: red;
}
</style>
<body bgcolor="Ivory" language="javascript" onload="return window_onload()">
<%
Response.Write "<label id=SavePrompt><b><font size=3 face=verdana>Saving.  Please Wait...</b></font><br></label>"


'
' To do
'
' Loop through DCR IDs
' Get ActionAppoverID
' Find out if first dissapproval
' Set Approver Status
' Send Email

    Dim saDcrId

    
    saDcrId = split(Request.Form("txtMultiID"), ",")

'
'   Loop through Items
'
    For Each Dcr In saDcrId
        Call ProcessDcr(Dcr)
    Next

Sub ProcessDcr(txtID)

    Dim cn :    Set cn = Server.CreateObject("ADODB.Connection")
    Dim rs : 	Set rs = Server.CreateObject("ADODB.RecordSet")
    Dim rsDcr : Set rsDcr = Server.CreateObject("ADODB.RecordSet")
    Dim cmd
    Dim CurrentUserId
    Dim CurrentUserName
    Dim CurrentUserEmail
    Dim CurrentDomain
    Dim CurrentUserPartner
    Dim blnSendDisapprovedEmail

'
' Configure Connection
'
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 60
	cn.Open


'
' Get User Info
'

	CurrentUserID = 0
	CurrentUser = lcase(Session("LoggedInUser"))
	
	if instr(currentuser,"\") > 0 Then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	End If

	Set cmd = server.CreateObject("ADODB.Command")
	Set cmd.ActiveConnection = cn
	cmd.CommandType = 4
	cmd.CommandText = "spGetUserInfo"
	
	Set p = cmd.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cmd.Parameters.Append p

	Set p = cmd.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cmd.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cmd.Execute 

	Set cmd=nothing	

	if not (rs.EOF and rs.BOF) Then
	    If rs("PhWebImpersonate") > 0 Then
	        CurrentUserID = rs("PhWebImpersonate")
	    Else
		    CurrentUserID = rs("ID")
		End If
		    CurrentUserName = rs("Name")
		    CurrentUserEmail = rs("Email")
	End If
	rs.Close
'
'   Pull DCR record from the database.
'        
        Set cmd = Server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection = cn
		cmd.CommandType = 4
		cmd.CommandText = "spGetActionProperties"

		Set p = cmd.CreateParameter("@ID", 3, &H0001)
		p.Value = CLng(txtID)
		cmd.Parameters.Append p
	
		'rsDcr.CursorType = adOpenStatic
		Set rsDcr = cmd.Execute 
		
		If rsDcr.EOF and rsDcr.BOF Then
		    Errors = True
		End If
		
		Set cmd=nothing
'
'   Get ApprovalID from the database.
'		
        ApprovalID = 0
        If Not Errors Then
            Set cmd = Server.CreateObject("ADODB.Command")
            Set cmd.ActiveConnection = cn
            cmd.CommandType = 4
            cmd.CommandText = "spListApprovals"
    
            Set p = cmd.CreateParameter("@ActionID", 3, &H0001)
            p.Value = txtID
    	    cmd.Parameters.Append p

            Set p = cmd.CreateParameter("@ApproverID", 3, &H0001)
            p.Value = CurrentUserID
    	    cmd.Parameters.Append p
    	
    	    rs.CursorType = adOpenStatic
    	    Set rs = cmd.Execute
    	    Set cmd = Nothing
    	
    	    If Not (rs.EOF And rs.BOF) Then
        	    ApprovalID = rs("ID")
        	Else
    	        Errors = True
    	    End If
    	    rs.Close
    	    
        End If
'
'   Check to see if we are the first Disapproval or the last Approval.
'
        NewStatus = 0
        blnSendDisapprovedEmail = False
        If Not Errors Then
		    Set cmd = server.CreateObject("ADODB.Command")
		    Set cmd.ActiveConnection = cn
		    cmd.CommandType = 4
		    cmd.CommandText = "spVerifyAutoApprove"
    	
		    Set p = cmd.CreateParameter("@ID", 3, &H0001)
		    p.Value = txtID
		    cmd.Parameters.Append p

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cmd.Execute 
		    Set cmd=nothing

		    if not(rs.EOF and rs.BOF) Then
			    if Request.QueryString("ApproverStatus") = "2" Then 'Approved
   				    if rs("Verified") = 1 Then	
					    if CLng(rsDcr("Type")) = 3 Or CLng(rsDcr("Type")) = 7 Then				
						    NewStatus = 4
					    Else
						    NewStatus = 2
					    End If
				    End If
		        ElseIf Request.QueryString("ApproverStatus") = "3" Then 'Disapproved
			        If rs("DisapprovedCount") = 0 Then	
				        blnSendDisapprovedEmail = True
			        End If
		    End If
		    rs.Close
        End If
'
'   Update Approval
'        
        If Not Errors Then
            Set cmd = Server.CreateObject("ADODB.Command")
		    cmd.ActiveConnection = cn
		    cmd.CommandText = "spUpdateApproval"
		    cmd.CommandType =  &H0004
    	
		    Set p = cmd.CreateParameter("@ApprovalID", 3,  &H0001)
		    p.value = CLng( ApprovalID )
		    cmd.Parameters.Append p

		    Set p = cmd.CreateParameter("@Status", 16,  &H0001)
	        p.value = CLng(Request.QueryString("ApproverStatus"))
   		    cmd.Parameters.Append p

		    Set p = cmd.CreateParameter("@Comments", 200, &H0001, 300)
		    p.Value = Left(Request.Form("txtComments"),300)
		    cmd.Parameters.Append p

		    cmd.execute Rowseffected
    	    If rowseffected <> 1 Then
			    Errors = True
		    Else
			    blnUpdateApprovals = True											
		    End If
        End If
        
        if (Not Errors) and blnUpdateApprovals then
			Set cmd = server.CreateObject("ADODB.Command")
			Set cmd.ActiveConnection = cn
			cmd.CommandType = 4
			cmd.CommandText = "spSetApprovalList"
		
			Set p = cmd.CreateParameter("@ID", 3, &H0001)
			p.Value = clng( txtID )
			cmd.Parameters.Append p
	
			cmd.Execute 
			Set cmd = Nothing
		End If
					
		If Not Errors Then
			'cn.CommitTrans
		End If
'
'   Update Status
'		
        If (Not Errors) And NewStatus <> 0 And NewStatus <> rsDcr("Status") Then
            Set cmd = Server.CreateObject("ADODB.Command")
            Set cmd.ActiveConnection = cn
            cmd.CommandType = 4
            cmd.CommandText = "spUpdateDeliverableActionStatus"
            
            Set p = cmd.CreateParameter("@ID", 3, &H0001)
            p.Value = CLng(txtID)
            cmd.Parameters.Append p
            
            Set p = cmd.CreateParameter("@Status", 3, &H0001)
            p.Value = CLng(NewStatus)
            cmd.Parameters.Append p
            
            Set p = cmd.CreateParameter("@LastUpdUser", 200, &H0001, 200)
            p.Value = CurrentUserName
            cmd.Parameters.Append p
            
            cmd.Execute
            Set cmd = Nothing
        End If

'
'   Send Emails if Needed
'
        If Not Errors Then
		    Dim oMessage 

		    dim strNewOwnerEmail
		    dim strSubmitterEmail
		    dim strPM
		    dim strDivision
    		
		    set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "spGetEmployeeByID"
    		

		    Set p = cm.CreateParameter("@ID", 3, &H0001)
		    p.Value = rsDcr("OwnerID")
		    cm.Parameters.Append p

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cm.Execute 
		    Set cm=nothing

		    strNewOwnerEmail = rs("Email") & ""		
		    rs.Close

		    set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "spGetEmployeeByID"
    		

		    Set p = cm.CreateParameter("@ID", 3, &H0001)
		    p.Value = rsDcr("SubmitterID")
		    cm.Parameters.Append p

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cm.Execute 
		    Set cm=nothing

		    if rs.EOF and rs.BOF then
			    strSubmitterEmail = ""
		    else
			    strSubmitterEmail = rs("Email") & ""		
		    end if
		    rs.Close

		    set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "spGetProductVersion"

		    Set p = cm.CreateParameter("@ID", 3, &H0001)
		    p.Value = rsDcr("ProductVersionID")
		    cm.Parameters.Append p

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cm.Execute 
		    Set cm=nothing

		    strProductEmail = rs("Distribution") & ""		
		    strPM = rs("PMName") & ""
		    strDivision = rs("Division") & ""
		    strPMEmail = rs("PMEmail") & ""
		    rs.Close
    		
		    DisplayedID = txtID


            ' Get the current list of approvers 
            '
            Dim strCurrentApproverEmails : strCurrentApproverEmails = ""
        
            rs.open "spListApprovals " & DisplayedId, cn, adOpenStatic
            Do While Not rs.EOF
                    if InStr(strCurrentApproverEmails, rs("Email")) = 0 then
                        strCurrentApproverEmails = strCurrentApproverEmails & rs("Email") & ";"
                    end if
                rs.MoveNext
            Loop
            rs.close
    	    


		    'Notifications
    		
		    dim notifycount
		    notifycount = 0
		    dim strTO
		    dim strSubject
		    dim strBody
		    dim strProgramName
		    dim strID
		    dim strCoreTeam
		    dim strOwnerName
		    dim strBusiness
		    Dim strSummary : strSummary = rsDcr("Summary")&""

            dim strBodyJustification
	        dim strHPBody        
            dim strODMBody
            dim strToHPEmail
            dim strToODMEmail
            dim strExcaliburHPLink       
            dim strExcaliburODMLink      

            strBodyJustification = ""
	        strHPBody = ""
            strODMBody = ""
            strToHPEmail = ""
            strToODMEmail = ""
            strExcaliburHPLink = ""
            strExcaliburODMLink = ""
       	
		    strTO = ""

		    if rsDcr("DocChange") then
			    if txtID <> "" then 'Only notify them when editing
				    if strTo = "" then
					    strTO = "houdcrdocs@hp.com;"
				    else
					    strTO = strTO & "houdcrdocs@hp.com;"
				    end if
			    end if
		    end if
    	
		    '---------
    		
		    strGroupList=""
		    rs.open "spListGroups4Action " & Displayedid,cn,adOpenForwardOnly
		    do while not rs.eof	
			    if trim(rs("ID") & "" )  <> "" then
				    strGroupList=strGroupList & "," & rs("GroupName")
			    end if
			    rs.movenext	
		    loop
		    rs.close
		    if strGroupList <> "" then
			    strGroupList = mid(strGroupList,2)
		    end if
    		
		    TypeID = rsDcr("Type")
		    Select Case TypeID
		    Case "1"
			    strType = "Issue"
		    Case "2"
			    strType = "Action Item"
		    Case "3"
			    strType = "Change Request"
		    Case "4"
			    strType = "Status Note"
		    Case "5"
			    strType = "Improvement Opportunity"
		    Case "6"
			    strType = "Test Request"
			Case "7"
			    strType = "Service ECR"
		    End Select


		    set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "spGetAction4Mail"

		    Set p = cm.CreateParameter("@ID", 3, &H0001)
		    p.Value = DisplayedID
		    cm.Parameters.Append p

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cm.Execute 
		    Set cm=nothing

		    if rs.EOF and rs.BOF then
			    strBody = "<a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strType & "</a><br><br>"
		    else
    		
		    If strType = "4" And rs("Status") = "4" Then
                strTo = strTo & ";houprtdcrnotif@hp.com;"
		    End If
    						
		    strProgramName = rs("Program")
		    strProgramMail=  rs("EmailActive")
    			
		    strBusiness = ""
		    if (not isnull(rs("Consumer"))) and (not isnull(rs("Commercial"))) and (not isnull(rs("SMB"))) then
			    if rs("Consumer") then
				    strBusiness = strBusiness & ",Consumer"  
			    end if
			    if rs("Commercial") then
				    strBusiness = strBusiness & ",Commercial"  
			    end if
			    if rs("SMB") then
				    strBusiness = strBusiness & ",SMB"  
			    end if
			    if strBusiness <> "" then
				    strBusiness = mid(strBusiness,2)
			    end if
		    end if
    			
		    strBody = "<font face=Arial size=2><a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strType & "</a><br>"
		    strBody = strBody & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/Excalibur.asp"">Open Pulsar Today Page</a><br><br></font>"
		    strBody = strbody & "<font face=Arial size=2>"
		    strBody = strBody & "<b>NUMBER:</B> " & DisplayedID & "<BR>"
		    strBody = strBody & "<b>SUBMITTER:</b> " & rs("Submitter") & "<BR>"
		    If IsNull(rs("Created")) Then
			    strBody = strBody & "<b>SUBMITTED:</b> N/A" & "<BR>"
		    Else
			    strBody = strBody & "<b>SUBMITTED:</b> " & formatdatetime(rs("Created"), vbshortdate) & "<BR>"
		    End If
		    strProductName = rs("Program") & ""
		    strBody = strBody & "<b>PROGRAM:</b> " & rs("Program") & "<BR>"
		    strBody = strBody & "<b>TYPE:</b> " & strType & "<BR>"
		    if trim(TypeID) = "5" then
			    strBody = strBody & "<b>ISSUE/ACCOMPLISHMENT:</B> " & replace(strSummary,"""","&QUOT;") & "<BR>"
		    else
			    strBody = strBody & "<b>SUMMARY:</B> " & replace(strSummary,"""","&QUOT;") & "<BR>"
		    end if
		    Select Case rs("Status") & ""
		    Case "0"
                strBody = strBody & "<b>STATUS:</b> N/A" & "<BR>"
                strStatus = "N/A"
		    Case "1"
			    strBody = strBody & "<b>STATUS:</b> Open" & "<BR>"
                strStatus = "Open"
		    Case "2"
                strBody = strBody & "<b>STATUS:</b> Closed" & "<BR>"
                strStatus = "Closed"
		    Case "3"
                strBody = strBody & "<b>STATUS:</b> Need More Information" & "<BR>"
                strStatus = "Need More Information"
		    Case "4"
                strBody = strBody & "<b>STATUS:</b> Approved" & "<BR>"
                strStatus = "Approved"
		    Case "5"
                strBody = strBody & "<b>STATUS:</b> Disapproved" & "<BR>"
                strStatus = "Disapproved"
		    Case "6"
                strBody = strBody & "<b>STATUS:</b> Investigating" & "<BR>"
                strStatus = "Investigating"
		    End Select
    		
		    strBody = strBody & "<b>OWNER:</b> " & rs("Owner") & "<BR>"
		    strNewOwner = rsDcr("OwnerID")
		    strOwnerName = rs("Owner") & ""
    		
		    if strBusiness <> "" then
			    strBody = strBody & "<b>BUSINESS:</b> " & strBusiness & "<BR>"
		    end if
    		
		    if rs("Description") & "" <> "" then
			    strDescription = replace(rs("Description")& "",vbcrlf,"<BR>")
			    if trim(TypeID) = "5" then
				    StringArray = split(strDescription,chr(1))
				    if ubound(StringArray) > -1 then
					    if trim(StringArray(0)) <> "" then
						    strDescription = "<b>POSITIVE IMPACT:</b><br>" & StringArray(0)
					    else
						    strDescription = ""				
					    end if
				    end if
				    if ubound(StringArray) > 0 then
					    if trim(StringArray(0)) <> "" and trim(StringArray(1)) <> ""  then
						    strDescription = strDescription & "<BR><BR>"
					    end if
					    if trim(StringArray(1)) <> "" then
						    strDescription = strDescription & "<b>NEGATIVE IMPACT:</b><br>" & StringArray(1)
					    end if
				    end if
			    end if
			    if trim(TypeID) = "5" then
				    strBody = strBody & strDescription & "<BR>" & "<BR>"
			    else
				    strBody = strBody & "<b>DESCRIPTION:</b> " & "<BR>" & strDescription & "<BR>" & "<BR>"
			    end if
		    end if
    		
		    If trim(TypeID) = "3" Then
                dim Details
                if IsNull(rsDcr("Details")) then
                    Details = ""
                else
                    Details = rsDcr("Details")
                end if
		        strBody = strBody & "<b>DETAILS:</b> "  & "<BR>" & Replace(Details, VbCrLf, "<BR>") & "<BR>"
		    End If

		    if rs("Justification") & "" <> "" then
			    if trim(TypeID) = "5" then
				    strBodyJustification = "<b>ROOT CAUSE:</b><BR>" & replace(rs("Justification"),vbcrlf,"<BR>") & "<BR><BR>"
			    else	
				    strBodyJustification = "<b>JUSTIFICATION:</b><BR>" & replace(rs("Justification"),vbcrlf,"<BR>") & "<BR><BR>"
			    end if
		    end if

            strBody = strBody & strBodyJustification & ""

		    If strType <> "Status Note" Then
    		
			    strActions = replace(rs("Actions") & "",vbcrlf,"<BR>")
			    if trim(TypeID) = "5" then
				    StringArray = split(strActions,chr(1))
				    if ubound(StringArray) > -1 then
					    if trim(StringArray(0)) <> "" then
						    strActions = "<b>CORRECTIVE ACTIONS:</b><br>" & StringArray(0)
					    else
						    strActions = ""				
					    end if
				    end if
				    if ubound(StringArray) > 0 then
					    if trim(StringArray(0)) <> "" and trim(StringArray(1)) <> ""  then
						    strActions = strActions & "<BR><BR>"
					    end if
					    if trim(StringArray(1)) <> "" then
						    strActions = strActions & "<b>PREVENTIVE ACTIONS:</b><br>" & StringArray(1)
					    end if
				    end if
			    end if
			    if trim(TypeID) = "5" then
				    strBody = strBody & "<font color=red>" & strActions & "</font>" & "<BR>" & "<BR>"
			    else
				    strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & strActions & "</font>" & "<BR>" & "<BR>"
			    end if
			    strBody = strBody & "<b>RESOLUTION:</b> " & "<BR>" & replace(rs("Resolution") & "",vbcrlf,"<BR>") & "<BR>"
		    End If

		    if trim(rsDcr("AvailableForTest")) <> "" then
			    strBody = strBody & "<font color=red><b>SAMPLES AVAILABLE:</b> " & rsDcr("AvailableForTest") & "</font><BR><BR>"
		    end if
    			
		    if trim(TypeID) = "5" then
			    select case trim(rs("Priority") & "")
			    case "1"
				    strPriority="High"
			    case "2"
				    strPriority="Medium"
			    case "3"
				    strPriority="Low"
			    case else
				    strPriority=""
			    end select			
			    strBody = strBody & "<b>IMPACT:</b> " & strPriority & "<BR>"
		    end if
    			
		    if trim(TypeID) = "5" then
			    if rs("AffectsCustomers")=1  then
				    strCustomers = "Positive"
			    elseif rs("AffectsCustomers")=0 then
				    strCustomers = "&nbsp;"
			    else
				    strCustomers = "Negative"
			    end if
			    strBody = strBody & "<b>NET AFFECT:</b> " & strCustomers & "<BR>"
		    end if


            If rsDcr("ZsrpRequired") Then
                strBody = strBody & "<b>ZSRP READY TARGET: </b> " & rsDcr("ZsrpReadyTargetDt") & "<br>"
				strBody = strBody & "<b>ZSRP READY ACTUAL: </b> " & rsDcr("ZsrpReadyActualDt") & "<br><br><br>"
            End If

		    if trim(TypeID) = "5"  and trim(rs("AvailableNotes") & "") <> "" then
			    strBody = strBody & "<b>METRIC IMPACTED:</b> " & rs("AvailableNotes") & "<BR>"
		    end if

		    strCoreTeam = rs("CoreTeamRep")

		    if (not isnull(rs("BTODate"))) and (rs("Distribution") = "BTO" or rs("Distribution") = "BOTH") then
	            strBody = strBody & "<b>BTO-IMPLEMENT BY:</b> " & rs("BTODate") & "<BR>"
		    end if
		    if (not isnull(rs("CTODate"))) and (rs("Distribution") = "CTO" or rs("Distribution") = "BOTH") then
	            strBody = strBody & "<b>CTO-IMPLEMENT BY:</b> " & rs("CTODate") & "<BR>"
		    end if

		    If strType <> "Status Note" Then
			    strBody = strBody & "<b>NOTIFY ON CLOSURE:</b> " & rs("Notify") & "<BR>"
		    End If
		    strNotify = rs("Notify") & ""
		    if rs("Approvals") & "" <> "" then
			    strBody = strBody & "<b>APPROVALS:</b><font color=teal><BR>" & replace(rs("Approvals"),vbcrlf,"<BR>") & "</font><BR><BR>"
		    end if
		    if strGroupList <> "" then
			    strBody = strBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
		    end if

		    strBody = strBody & "<br><font size=1 color=red face=verdana>HP Restricted</font>"

		    end if
		    rs.Close
  		
		    if blnSendDisapprovedEmail then 'First Disapprover Found
			    strTo = strTO & strNewOwnerEmail & ";"
			    strTo = replace(StrTo,CurrentUserEmail & ";","")
			    if strCoreTeam = "Sustaining System Team" then
				    if strTO = "" then
					    strTO = CCTMailList & ";"
				    else
					    strTO = strTo & CCTMailList & ";"
				    end if
			    end if			
			    Response.Write "<BR>First time disapproved by Approver: " & strTO & "<BR>"		
			    notifycount = notifycount + 1

                'separate the hp and odm email addresses
                strToHPEmail = ""
                strToODMEmail = ""
                EmailArray = split(strTO,";")
				for each emailaddress in EmailArray
                    if Len(Trim(emailaddress)) > 0 then                    
					    if instr(UCase(emailaddress), "@HP.COM") = 0 then
                            strToODMEmail = strToODMEmail & ";" & emailaddress
                        else
                            strToHPEmail = strToHPEmail & ";" & emailaddress
                        end if		
                    end if
				next

'''''''''''' send to HP User '''start
			    if strToHPEmail <> "" then
				    Set oMessage = New EmailQueue

                    strHPBody = strBody & vbcrlf  & "<!-- HP Notify --><!-- Multiple Approve From TodayPage -->" & vbcrlf 
		  
				    oMessage.From = CurrentUserEmail
				    if strProgramMail = "1" then
					    oMessage.To= strToHPEmail
					    oMessage.Subject = strType & " " & DisplayedID  & " (Disapproved by Approver) : " & strProductName &  " : " & replace(strSummary,"""","'")
				    else
					    oMessage.To= CurrentUserEmail 
					    oMessage.Subject = "TEST MAIL: " & strType & " " & DisplayedID  & " (Disapproved by Approver) : " & strProductName &  " : " & replace(strSummary,"""","'")
				    end if
				    strHPBody = "<font face=Arial size=2 color=red><b>This item has been disapproved by an approver. Further disapprovals will not generate and email notification.</font></b><BR><BR>" & strHPBody
    				
				    if strProgramMail = "1" then
					    oMessage.HTMLBody = strHPBody
				    else
					    oMessage.HTMLBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strToHPEmail & "</font><BR><BR>" & strHPBody
				    end if
    				
				    oMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				    Set oMessage = Nothing 			
			    end if
'''''''''''' send to HP User '''end

'''''''''''' send to ODM User '''start
			    if strToODMEmail <> "" then
				    Set oMessage = New EmailQueue

                    strODMBody = strBody & vbcrlf  & "<!-- ODM Notify --><!-- Multiple Approve From TodayPage -->" & vbcrlf 

                    ''' prp url
                    strODMBody = replace(strODMBody, "href=""https://" & Application("Excalibur_ODM_ServerName"), "href=""https://pulsarweb-pro.prp.ext.hp.com")

                    if strBodyJustification <> "" then
					    strODMBody = replace(strODMBody, strBodyJustification, "")
                    end if

				    oMessage.From = CurrentUserEmail
				    if strProgramMail = "1" then
					    oMessage.To= strToODMEmail
					    oMessage.Subject = strType & " " & DisplayedID  & " (Disapproved by Approver) : " & strProductName &  " : " & replace(strSummary,"""","'")
				    else
					    oMessage.To= CurrentUserEmail 
					    oMessage.Subject = "TEST MAIL: " & strType & " " & DisplayedID  & " (Disapproved by Approver) : " & strProductName &  " : " & replace(strSummary,"""","'")
				    end if
				    strODMBody = "<font face=Arial size=2 color=red><b>This item has been disapproved by an approver. Further disapprovals will not generate and email notification.</font></b><BR><BR>" & strODMBody
    				
				    if strProgramMail = "1" then
                        oMessage.HTMLBody = strODMBody
				    else
					    oMessage.HTMLBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strToODMEmail & "</font><BR><BR>" & strODMBody
				    end if
    				
				    oMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				    Set oMessage = Nothing 			
			    end if
'''''''''''' send to ODM User '''end

		    end if

		    if (NewStatus = 2 or NewStatus = 4 or NewStatus = 5) and (rsDcr("Status") <> trim(cstr(newStatus)) ) then
			    Response.Write strType & ":" & strProductEmail & "<BR>"
			    if TypeID = "3" then
				    strTO = strTO & strNewOwnerEmail  & ";" & strProductEmail & ";" & strSubmitterEmail & ";" & strCurrentApproverEmails & ";"
				    if rsDcr("notify") <> "" then
					    if strTO = "" then
						    strTO = rsDcr("notify") & ";" 
					    else
						    strTO = strTO & rsDcr("notify") & ";" 
					    end if
				    end if

				    if trim(strDivision) = "1" then
					    if strTO = "" then
						    strTO = "NotebookDCRNotification@hp.com;"
					    else
						    strTO = strTO & "NotebookDCRNotification@hp.com;"
					    end if
				    end if

				    'strTo = replace(lcase(StrTo),lcase(trim(CurrentUserEmail)) & ";","")
				    Response.Write "<BR>Item Closed: " & strTO & "<BR>"
			    else
				    'if blnReleaseNotification = 1 then
				    '	strTO = strNewOwnerEmail
				    'else
					    strTO = strTO & strNewOwnerEmail  & ";" & strSubmitterEmail  & ";"
				    'end if
				    if rsDcr("notify") <> "" then
					    if strTO = "" then
						    strTO = rsDcr("notify") & ";"
					    else
						    strTO = strTO & rsDcr("notify") & ";"
					    end if
				    end if
				    strTo = replace(StrTo,CurrentUserEmail & ";","")
				    Response.Write "<BR>Item Closed: " & strTO &  "<BR>"
			    end if

			    if strCoreTeam = "Sustaining System Team" then
				    if strTO = "" then
					    strTO = CCTMailList & ";"
				    else
					    strTO = strTo & CCTMailList & ";"
				    end if
			    end if

                'separate the hp and odm email addresses
                strToHPEmail = ""
                strToODMEmail = ""
                EmailArray = split(strTO,";")
				for each emailaddress in EmailArray
                    if Len(Trim(emailaddress)) > 0 then                    
					    if instr(UCase(emailaddress), "@HP.COM") = 0 then
                            strToODMEmail = strToODMEmail & ";" & emailaddress
                        else
                            strToHPEmail = strToHPEmail & ";" & emailaddress
                        end if		
                    end if
				next



'''''''''''' send to HP User   '''start  			
			    if strToHPEmail <> "" then
				    notifycount = notifycount + 1

				    Set oMessage = New EmailQueue

                    strHPBody = strBody & vbcrlf & "<!-- HP Notify --><!-- Multiple Approve From TodayPage -->" & vbcrlf 

				    oMessage.From = CurrentUserEmail
				    if strProgramMail = "1" then
					    oMessage.To=strToHPEmail
					    oMessage.Subject = strType & " " & DisplayedID  & " (" & strStatus & ") : " & strProductName &  " : " & replace(strSummary,"""","'")
				    else
					    oMessage.To=CurrentUserEmail
					    oMessage.Subject = "TEST MAIL: " & strType & " " & DisplayedID  &  " (" & strStatus & ") : " & strProductName &  " : " & replace(strSummary,"""","'")
				    end if	
    				
				    if strProgramMail = "1" then
					    oMessage.HTMLbody = "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strHPBody 
				    else
					    oMessage.HTMLbody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strToHPEmail & "</font><BR><BR>" & "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strHPBody 
				    end if
    	
				    oMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				    Set oMessage = Nothing 
			    end if
'''''''''''' send to HP User '''end

'''''''''''' send to ODM User '''start
			    if strToODMEmail <> "" then
				    notifycount = notifycount + 1

				    Set oMessage = New EmailQueue

                    strODMBody = strBody & vbcrlf & "<!-- ODM Notify --><!-- Multiple Approve From TodayPage -->" & vbcrlf 

                    ''' prp url
                    strODMBody = replace(strODMBody, "href=""https://" & Application("Excalibur_ODM_ServerName"), "href=""https://pulsarweb-pro.prp.ext.hp.com")

                    if strBodyJustification <> "" then
					    strODMBody = replace(strODMBody, strBodyJustification, "")
                    end if

				    oMessage.From = CurrentUserEmail
				    if strProgramMail = "1" then
					    oMessage.To=strToODMEmail
					    oMessage.Subject = strType & " " & DisplayedID  & " (" & strStatus & ") : " & strProductName &  " : " & replace(strSummary,"""","'")
				    else
					    oMessage.To=CurrentUserEmail
					    oMessage.Subject = "TEST MAIL: " & strType & " " & DisplayedID  &  " (" & strStatus & ") : " & strProductName &  " : " & replace(strSummary,"""","'")
				    end if	
    				
				    if strProgramMail = "1" then
					    oMessage.HTMLbody = "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strODMBody
				    else
					    oMessage.HTMLbody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strToODMEmail & "</font><BR><BR>" & "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strODMBody 
				    end if
    	
				    oMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				    Set oMessage = Nothing 
			    end if

'''''''''''' send to ODM User '''end


    			
		    end if

            rsDcr.Close

        End If

        Set rs = Nothing
        Set rsDcr = Nothing
        cn.Close
        Set cn = nothing

    End If
End Sub



response.Redirect "today.asp"

Function StripHTMLTag(ByVal sText)
   StripHTMLTag = ""
   fFound = False
   Response.Write sText & "<BR>" & vbcrlf
   Do While InStr(sText, "<")
      fFound = True
      StripHTMLTag = StripHTMLTag & " " & Left(sText, InStr(sText, "<")-1)
      strTag = lcase(trim(mid(sText,InStr(sText, "<"),InStr(sText, ">") - InStr(sText, "<")+1)))

'	  if strTag = "<b>" or strTag = "</b>" or strTag = "<i>" or strTag = "</i>" or strTag = "<u>" or strTag = "</u>" Then
		if left(replace(ucase(strTag)," ",""),5) <> "<" & trim("FONT") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("FONT") and left(replace(ucase(strTag)," ",""),5) <> "<" & trim("SPAN") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("SPAN") and left(replace(ucase(strTag)," ",""),4) <> "<" & trim("DIV") and left(replace(ucase(strTag)," ",""),5) <> "</" & trim("DIV") and left(replace(ucase(strTag)," ",""),2) <> "<" & trim("P") and left(replace(ucase(strTag)," ",""),3) <> "</" & trim("P") Then
			StripHTMLTag = StripHTMLTag & strTag
      End If

	  
      sText = MID(sText, InStr(sText, ">") + 1)

      
   Loop
   StripHTMLTag = StripHTMLTag & sText
   If Not fFound Then StripHTMLTag = sText
End Function

%>
</body>
</html>