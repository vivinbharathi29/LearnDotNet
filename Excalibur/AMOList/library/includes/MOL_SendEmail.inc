<%
Function SendMOLEmail(ByVal sRepository, ByVal nMOLID, ByVal nType, Byval nModuleID, ByVal sMsg, ByVal nModuleType)

	dim oSvr, oRs, oRsUser, sTo, sFrom, sFromName, sSubject, sBody, sLink, sServer, sAutomatedMesg, strServername
	dim oErr, oRecords, oError
	dim strError', sFromName, sFrom, sColor
	dim intShowGroupID
	dim sToFilter

	if nType = 0 or nType = "" then
		nType = 1
	end if
	
	'Subject	
    select case nType
      case 1
		sSubject = "**URGENT** PLEASE REVIEW FOR APPROVAL - MOL Marketing Sign-off"
      case 2
		sSubject = "MOL Request Sign-off"
      case 3
        sSubject = "MOL Engineering Commitment"
      case 4
        sSubject = "MOL Engineering(SEPM) Review"
	  case 5
        sSubject = "MOL Platform PM Review"
	  case 6
        sSubject = "MOL Scope Change"
    end select

	
	if instr(1,  Request.ServerVariables("Server_Name"), application("ServerFullAddr"), 1) then
		strServername =  Request.ServerVariables("Server_Name")
	else
		strServername =  Request.ServerVariables("Server_Name") & application("ServerFullAddr")
	end if
	
	'From List
	'dim strError, sFromName, sFrom
	'dim intShowGroupID

	strError = ""

	set oRecords = Server.CreateObject("JF_S_General.ISGeneral")
	set oError = oRecords.ViewUsers (Application("Repository"), oRs, 0, CLng(Session("AMOUserID")))
	'check if any error happens
	if not oError is nothing then
		strError = Errors_GetHTML(oError)
	else
		if oRs.RecordCount = 0 then
			strError = strError & "<img src='" & Session("ServerPath") & "/library/Images/jpegs/yellowtriangle.jpg'>" & vbCrLf
			strError = strError & "<table WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCRLF
			strError = strError & "<tr><td><ul><li>Your Email address was not found.</li></ul></td></tr>" & vbCrLf
			strError = strError & "</table>" & vbCRLF
		else
			sFrom = oRs("Email")
			sFromName = oRs("FirstName") & " " & oRs("LastName")
		end if
		oRs.Close
		set oRs = nothing
	end if
	set oError = nothing
	set oRecords = nothing
	
	'To list
	set oSvr = server.CreateObject("JF_S_MOL.ISMOL")
	if not oSvr is nothing then
		set oErr = oSvr.MOL_View(sRepository, nMOLID, oRs)
		if oErr is nothing then
			if nType = 1 or nType = 2 or nType = 3 or nType = 6  then
				dim nSignoffListType
				if nType = 6 then
				    nSignoffListType = 2
				else
				    nSignoffListType = nType
				end if
				if nType=3 then
					sToFilter =sMsg
					set oErr = oSvr.MOL_ViewSignoff(sRepository, nMOLID, nSignoffListType, oRsUser, sToFilter, 4)
				else
					set oErr = oSvr.MOL_ViewSignoff(sRepository, nMOLID, nSignoffListType, oRsUser)
				end if 
				
			elseif nType = 4 then			    
				set oErr = oSvr.MOL_ViewSEPM(sRepository, nMOLID, oRsUser, 2)
			elseif nType = 5 then			    
				set oErr = oSvr.MOL_ViewPPM(sRepository, nMOLID, oRsUser)
            end if
    		set oSvr = nothing

			'Compile Addressee string
			sTo = ""
			if oErr is nothing then
				if oRsUser.RecordCount <> 0 then
					While not oRsUser.EOF and not oRsUser.BOF
						if not IsNull(oRsUser("Email")) AND oRsUser("Email") <> "" Then
							sTo = sTo & oRsUser("Email") & ";"
						end if
						oRsuser.MoveNext
					Wend
					set oRsUser = nothing
				end if
			end if
		end if
	end if		
	
	'Build body of email
	if oErr is nothing then			
		if (nType = 1 or nType = 2) then
			sLink = Server.URLEncode("MOL/ModuleList/isapi/Signoff_Properties.asp?Type=" & nType & "&MOLID=" & nMOLID)
			sBody = "You are receiving this email because you are one of the signatories for the MOL, " _ 
				& "'" & oRs("Name") & "' . Please go to the following link to respond. <br>" _			
				& "<A href='" & Session("PrefixFullPath")  & "default.asp?link=" & sLink & "' style=color:blue target=_blank>MOL Marketing Sign-off page</A>" 

		elseif (nType = 3 ) then
			sLink = Server.URLEncode("MOL/ModuleList/isapi/MOL_Commitment.asp?MOLID=" & nMOLID)
		    sBody = "You are receiving this email because you are one of the commitment signatories for the MOL, " _
		    	  &	"'" & oRs("Name") & "'.  Please go to the following link to review the MOL.<br>" _
		    	  & "<A href='" & Session("PrefixFullPath")  & "default.asp?link=" & sLink & "' style=color:blue target=_blank>MOL Commitment page</A> "
		
		elseif (nType = 4) then
			sLink = Server.URLEncode("MOL/ModuleList/isapi/MOL_ModuleList.asp?MOLID=" & nMOLID)
		    sBody = "You are receiving this email because you are assigned as the SEPM for the MOL, " _
		    	  &	"'" & oRs("Name") & "'.  Please go to the following link to review the MOL.<br>" _
		    	  & "<A href='" & Session("PrefixFullPath")  & "default.asp?link=" & sLink & "' style=color:blue target=_blank>MOL Engineering(SEPM) Review page</A><br> " 
		    	  
		    	
		
		elseif (nType = 5) then
			sLink = Server.URLEncode("MOL/ModuleList/isapi/MOL_ModuleList.asp?MOLID=" & nMOLID)
		    sBody = "You are receiving this email because you are assigned as the Platform PM for the MOL, " _
		    	  &	"'" & oRs("Name") & "'.  Please go to the following link to review the MOL.<br>" _
		    	  & "<A href='" & Session("PrefixFullPath")  & "default.asp?link=" & sLink & "' style=color:blue target=_blank>MOL Module List page</A> "
		
		elseif (nType = 6) then
			sLink = Server.URLEncode("MOL/ModuleList/isapi/Request_Properties.asp?Mode=2&MOLID=" & nMOLID & "&ModuleID=" & nModuleID & "&ModuleType=" & nModuleTypeID)
		    sBody = "Scope Change has been entered/updated for the MOL, " _
		    	  &	"'" & oRs("Name") & "'. New/Updated scope change ID(s) are: " & sMsg _
		    	  & ". Please go to the following link for more details about the update.<br>" _
		    	  & "<A href='" & Session("PrefixFullPath")  & "default.asp?link=" & sLink & "' style=color:blue target=_blank>MOL Request Properties</A> "
		end if
		
	 	set oErr = SendIRSEmail_To2( _
		 "", _
		 sTo, _
		 "", _
		 "", _
		 "", _
		 "", _
		 -1, _
		 -1, _
		 sSubject, _
	     "", _
	     sBody, _
	     0, _
	     "", _
	     sFromName, _
	     sFrom _
	     )
	end if     
	set SendMOLEmail = oErr	
end function


function SendMOL_ModuleChangeEmail(sMode, nMOLID, sModuleIDs, nDivisionID, sMOLName)
	dim sHTML, sPriority, sServername, sIUO, strSubject, strType, sBusinessCritical
	dim intCount
	dim oErr
	dim arrColors
	dim oSvr, oRs
	
	arrColors = Array("#CCCCCC", "white")
	
	set oSvr = server.CreateObject("JF_S_Module.ISModule")
	if not oSvr is nothing then
		set oErr = oSvr.Module_Search(sRepository, " m.moduleID in (" & sModuleIDs & ") and md.DivisionID=" & nDivisionID, oRs)
	end if 	
		
	strSubject = "MOL Module Change: "  
	
	if instr(1,  Request.ServerVariables("Server_Name"), application("ServerFullAddr"), 1) then
		sServername =  Request.ServerVariables("Server_Name")
	else
		if instr(sServername, "irsweb") > 0 or instr(sServername, "irssandbox") > 0 or instr(sServername, "irstest") > 0 then
			sServername =  Request.ServerVariables("Server_Name") & application("ServerFullAddr")
		else
			'must be a developer's server
			sServername =  Request.ServerVariables("Server_Name")
		end if
	end if
	
	if instr(sServername, "irsweb") = 0 then
		strSubject = strSubject & " (" & sServerName & ")"
	end if

	'Build body of email
	sHTML = "<p><font size ='3' face='Arial'><strong>" & vbCrLf
	
	sHTML = sHTML & " Module and Option List: " & sMOLName & "<br>"

	sHTML = sHTML & " The following modules have been " & sMode & "."
	
	
	sHTML = sHTML & "<table border=1 cellPadding=2 cellSpacing=0 WIDTH=""100%"">" & vbCrLf
	
	intCount = 1
	sHTML = sHTML & "	<tr bgcolor=""" & arrColors(intCount mod 2) & """>" & vbCrLf
	sHTML = sHTML & "		<td><b>Module Name </b></td>" & vbCrLf
'	sHTML = sHTML & "		<td><b>Time Changed</b></td>" & vbCrLf
	sHTML = sHTML & "	</tr>" & vbCrLf

		
	if oRs.RecordCount > 0 then
		While not oRs.EOF and not oRs.BOF
			intCount = intCount +1
			sHTML = sHTML & "	<tr bgcolor=""" & arrColors(intCount mod 2) & """>" & vbCrLf
			sHTML = sHTML & "		<td>" & oRs("Description") & "</td>" & vbCrLf
		'	sHTML = sHTML & "		<td></td>" & vbCrLf
			sHTML = sHTML & "	</tr>" & vbCrLf

			oRs.MoveNext
		Wend
		set oRs = nothing
	end if



	sHTML = sHTML & "</table>" & vbCrLf
	
	
	
	'sHTML ="dsdsfswdfw"
	set oErr = SendIRSEmail2( "MOL Module change notification",_
		"", _
		"", _
		"", _
		"", _
		Application("MAIL_MOL_MODULE_CHANGE"), _
		nDivisionID, _
		-1, _
		strSubject, _
		"", _
		sHTML, _
		False, _
		"", _
		"", _
		"" _
		)
		
	set SendMOL_ModuleChangeEmail = oErr	
end function

function SendMOL_ModuleChangeEmail_update(sMode, nMOLID, sModuleIDs, nDivisionID, ByVal oRs, sDesc)
	dim sHTML, sPriority, sServername, sIUO, strSubject, strType, sBusinessCritical
	dim intCount
	dim oErr
	dim arrColors
	'dim oSvr, oRs
	
	arrColors = Array("#CCCCCC", "white")
	
	'set oSvr = server.CreateObject("JF_S_Module.ISModule")
	'if not oSvr is nothing then
	'	set oErr = oSvr.Module_Search(sRepository, " m.moduleID in (" & sModuleIDs & ")", oRs)
	'end if 	
		
	strSubject = "MOL Module Change: "  
	
	if instr(1,  Request.ServerVariables("Server_Name"), application("ServerFullAddr"), 1) then
		sServername =  Request.ServerVariables("Server_Name")
	else
		if instr(sServername, "irsweb") > 0 or instr(sServername, "irssandbox") > 0 or instr(sServername, "irstest") > 0 then
			sServername =  Request.ServerVariables("Server_Name") & application("ServerFullAddr")
		else
			'must be a developer's server
			sServername =  Request.ServerVariables("Server_Name")
		end if
	end if
	
	if instr(sServername, "irsweb") = 0 then
		strSubject = strSubject & " (" & sServerName & ")"
	end if

	'Build body of email
	sHTML = "<p><font size ='3' face='Arial'><strong>" & vbCrLf
	
	sHTML = sHTML & " Module name: " & sDesc & "<br>"

	sHTML = sHTML & " The module has been " & sMode & " for the following MOLs:"
	
	
	sHTML = sHTML & "<table border=1 cellPadding=2 cellSpacing=0 WIDTH=""100%"">" & vbCrLf
	
	intCount = 1
	sHTML = sHTML & "	<tr bgcolor=""" & arrColors(intCount mod 2) & """>" & vbCrLf
	sHTML = sHTML & "		<td><b>MOL Name </b></td>" & vbCrLf
'	sHTML = sHTML & "		<td><b>Time Changed</b></td>" & vbCrLf
	sHTML = sHTML & "	</tr>" & vbCrLf

		
	if oRs.RecordCount > 0 then
		oRs.filter="divisionID=" & nDivisionID
		if oRs.RecordCount > 0 then
			oRs.movefirst
			While not oRs.EOF and not oRs.BOF
				if oRs("StatusID") >= Application("MOL_EG_COMMITMENT") then
				'if oRs("StatusID") >= 43 then 'for testing
					intCount = intCount +1
					sHTML = sHTML & "	<tr bgcolor=""" & arrColors(intCount mod 2) & """>" & vbCrLf
					sHTML = sHTML & "		<td>" & oRs("Path") & "</td>" & vbCrLf
				'	sHTML = sHTML & "		<td></td>" & vbCrLf
					sHTML = sHTML & "	</tr>" & vbCrLf
				end if 
				oRs.MoveNext
			Wend
		end if	
		oRs.filter=""
		'set oRs = nothing
	end if



	sHTML = sHTML & "</table>" & vbCrLf
	
	
	
	'sHTML ="dsdsfswdfw"
	if intCount > 1 then
		set oErr = SendIRSEmail2( "MOL Module change notification",_
		"", _
		"", _
		"", _
		"", _
		Application("MAIL_MOL_MODULE_CHANGE"), _
		nDivisionID, _
		-1, _
		strSubject, _
		"", _
		sHTML, _
		False, _
		"", _
		"", _
		"" _
		)
		set SendMOL_ModuleChangeEmail_update = oErr	
	else
		set SendMOL_ModuleChangeEmail_update = nothing	
	end if	
	
end function

'Procedure: Delay_5Seconds

'@Purpose:  When loop through to send different emails, somehow if it's too 'fast' there

'           was a problem that some emails are missing eventhough the log file say it was sent,

'        dely 5 seconds betwwen them have been proved be working fine

'Inputs:

'@Returns:  none.

'-----------------------------------------------------------------

Sub Delay_5Seconds()

    Dim CurrentSeconds
    Dim currentminute
    CurrentSeconds = Second(Now())
	currentminute = Minute(Now())
	If currentminute = 60 Then
		currentminute = 0
	End If
    If CurrentSeconds < 55 Then
        Do While Second(Now()) - CurrentSeconds < 5
        Loop
    Else
		Do While (Second(Now()) < 5 Or (Second(Now()) < 61 And Second(Now()) > 54))
		Loop
    End If  
End Sub

function CopyRs(oRs, bCopy)
	dim oDupRs, oFld, i
    set oDupRs = Server.CreateObject ("ADODB.Recordset")

	for i = 0 to oRs.Fields.Count - 1
		set oFld =  oRs.Fields(i)
		oDupRs.Fields.Append oFld.Name, oFld.Type, oFld.DefinedSize, oFld.Attributes
	next
    oDupRs.CursorLocation = 3	'Use client-side cursors
    oDupRs.Open

	if bCopy then
	    If oRs.RecordCount > 0 Then
	        oRs.MoveFirst
	    End If
	
	    While Not oRs.EOF
	        oDupRs.AddNew
	        For i = 0 To oRs.Fields.Count - 1
	            oDupRs.Fields(oDupRs.Fields(i).Name).Value = oRs.Fields(i).Value
	        Next
	        oRs.MoveNext
	    Wend
	
	    If oRs.RecordCount > 0 Then
	        oRs.MoveFirst
	    End If
	end if

	set CopyRs = oDupRs
end function

%>
