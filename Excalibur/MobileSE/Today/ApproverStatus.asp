<%@ Language=VBScript %>
<!-- #include file="../../includes/EmailQueue.asp" -->
<!-- #include file="../../includes/no-cache.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	window.returnValue = txtResult.value;
	window.close();
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<b><font face=verdana size=2>

<BR><BR>&nbsp;&nbsp;Updating Status.  Please wait...</font></b>

<%
	dim cm
	dim cn
	dim strResult
	dim RowsEffected
	dim rs 

	if request("ID") <> "" and request("Status") <> "" then
		strConnect = Session("PDPIMS_ConnectionString")
		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = strConnect
		cn.CommandTimeout = 60
		cn.Open
		
		cn.BeginTrans
		
		set cm = server.CreateObject("ADODB.command")
			
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateApprovalStatus"
		cm.CommandType =  &H0004
	
		Set p = cm.CreateParameter("@ApprovalID", 3,  &H0001)
		p.value = clng( request("ID"))
		cm.Parameters.Append p
	
		Set p = cm.CreateParameter("@Status", 16,  &H0001)
		p.value = clng(request("Status"))
		cm.Parameters.Append p

		cm.execute Rowseffected
		
		set cm=nothing
		
		'response.Write RowsEffected & "<BR>"
		
		if rowseffected <> 1 then
			cn.RollbackTrans
			strResult = "0"
		else

			cn.CommitTrans

			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spSetApprovalList"
	
			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = clng( request("ActionID"))
			cm.Parameters.Append p
	
			cm.Execute 
			Set cm=nothing

			'cn.Execute "spSetApprovalList " & clng( request("ActionID"))

			strResult = "1"
			
			set rs = server.CreateObject("ADODB.recordset")
			
			'Get CurrentUser Email address
			dim CurrentDomain
			dim CurrentUserPartner
			CurrentUser = lcase(Session("LoggedInUser"))

			if instr(currentuser,"\") > 0 then
				CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
				Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
			end if

			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			set rs = server.CreateObject("ADODB.recordset")

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

			CurrentUserID = 0
			if not (rs.EOF and rs.BOF) then
				CurrentUserID = rs("ID")
				CurrentUserName = rs("Name")
				CurrentUserEmail = rs("Email")
			end if
			rs.Close

			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetApproverEmail"
			
			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("ID")
			cm.Parameters.Append p
	
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			'rs.Open "spGetApproverEmail " & request("ID") ,cn,adOpenForwardOnly
			
			if not (rs.EOF and rs.BOF) then


				strTo = rs("email") & ""
				
				select case rs("Type")
				case "1"
					strType = "Issue"
				case "2"
					strType = "Action Item"
				case "3"
					strType = "Change Request"
				case "4"
					strType = "Status Note"
				case else
					strType = "Item"
				end select
				
				strBody = "<font face=Arial size=2><a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & rs("Type") & "&id=" & rs("ActionID") & """>Open this " & strtype & "</a><br>"
				strBody = strBody & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/Excalibur.asp"">Open Pulsar Today Page</a><br><br></font>"
				strBody = strbody & "<font face=Arial size=2>"
				strBody = strBody & "<b>NUMBER:</B> " & rs("ActionID") & "<BR>"
				strBody = strBody & "<b>SUBJECT:</B> " & rs("Summary") & "<BR>"
				strBody = strBody & "<b>PROGRAM:</b> " & rs("Product") & "<BR>"
				strBody = strBody & "<b>TYPE:</b> " & strtype & "<BR>"
    
				strBody = strBody & "<b>OWNER:</b> " & rs("Owner") & "<BR>"
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
				
				strBody = strBody & "<b>CORE TEAM REP:</b> " & rs("CoreTeamRep") & "<BR>"
				strCoreTeam = rs("CoreTeamRep")
				strBody = strBody & "<b>SUBMITTER:</b> " & rs("Submitter") & "<BR>"
				If IsNull(rs("Created")) Then
					strBody = strBody & "<b>SUBMITTED:</b> N/A" & "<BR>"
				Else
					strBody = strBody & "<b>SUBMITTED:</b> " & formatdatetime(rs("Created"), vbshortdate) & "<BR>"
				End If
				If strtype <> "Status Note" Then
			        strBody = strBody & "<b>TARGET CLOSURE:</b> " & rs("TargetDate") & "<BR>"
					strBody = strBody & "<b>NOTIFY ON CLOSURE:</b> " & rs("Notify") & "<BR>"
				End If
				strNotify = rs("Notify") & ""
				if rs("Approvals") & "" <> "" then
					strBody = strBody & "<b>APPROVALS:</b><font color=teal><BR>" & replace(rs("Approvals"),vbcrlf,"<BR>") & "</font><BR><BR>"
				end if
				strBody = strBody & "<b>DESCRIPTION:</b> " & "<BR>" & replace(rs("Description")& "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
				If strtype <> "Status Note" Then
					strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & replace(rs("Actions") & "",vbcrlf,"<BR>") & "</font>" & "<BR>" & "<BR>"
					strBody = strBody & "<b>RESOLUTION:</b> " & "<BR>" & replace(rs("Resolution") & "",vbcrlf,"<BR>") & "<BR>"
				End If
				strBody = strBody & "<br><br><br><br><br><br><br><br><br><font size=1 color=red face=verdana>HP Restricted</font>"

				
				'Send Mail
				
				Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
				'Set oMessage.Configuration = Application("CDO_Config")				
				
				oMessage.From = CurrentUserEmail
				if rs("EmailActive") = 1 then
					oMessage.To= strTo
					if request("Status") = "4" then
						oMessage.Subject =  strtype & " " & rs("ActionID") &  " (Approval no Longer Required): " & rs("Summary")
					else
						oMessage.Subject =  strtype & " " & rs("ActionID") &  " (Approval Requested): "  & rs("Summary")
					end if
				else
					oMessage.To= CurrentUserEmail
					if request("Status") = "4" then
						oMessage.Subject =  "TEST MAIL: " & strtype & " " & rs("ActionID") &  " (Approval Status RESET or CANCELLED): " & rs("Summary")
					else
						oMessage.Subject =  "TEST MAIL: " & strtype & " " & rs("ActionID") &  " (Approval Requested): "  & rs("Summary")
					end if
				end if
		
				if request("Status") = "4" then
					strBody = "<font face=Arial size=2 color=red><b>Your request for approval has been RESET or CANCELLED.  Please use Excalibur to review the " & strtype & " again for details.</font></b><BR><BR>" & strBody 
				else
					strBody = "<font face=Arial size=2 color=red><b>Your request for approval has been RESET or CANCELLED.  Please use Excalibur to review the " & strtype & " again for details.</font></b><BR><BR>" & strBody 
				end if
						
				if rs("EmailActive") = 1 then
					oMessage.HTMLBody = strBody
				else
					oMessage.HTMLBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strTo & "</font><BR><BR>" & strBody
				end if
						
				oMessage.Send 
				Set oMessage = Nothing 			
				rs.Close
			end if		
			set rs = nothing		
		end if
		
	else
		strResult = "0"
	end if

	set cm = nothing
	set cn = nothing
%>
<INPUT type="hidden" id=txtResult name=txtResult value="<%=strResult%>">
</BODY>
</HTML>
