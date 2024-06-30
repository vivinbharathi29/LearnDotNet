<%@ Language=VBScript %>
<!-- #include file="includes/emailwrapper.asp" -->
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value=="1")
			{
			window.parent.returnValue = 1;
			window.parent.close();
			}
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

	<font size=2 face=verdana>Saving.  Please wait...</font>

<%

 function ProperCase(sString)
	if sString = "PM" then
		Propercase = sstring
	else
	
		Dim lTemp
		Dim sTemp, sTemp2
		Dim x
		sString = LCase(sString)
		if Len(sString) Then
		    sTemp = Split(sString, " ")
			lTemp = UBound(sTemp)
			For x = 0 To lTemp
				sTemp2 = sTemp2 & UCase(Left(sTemp(x), 1)) & Mid(sTemp(x), 2) & " " 
			Next
			ProperCase = trim(sTemp2)
		Else
			ProperCase = sString
	    End if
	end if
 End function


	dim WorkflowComplete
	dim strSQL
	dim cn
	dim FoundErrors
	dim strLangList
	dim Lang
	dim LangArray
	dim blnFailed
	dim blnAllReleased
	dim blnAllFailed
	dim blnAnyReleased
	dim strNextMilestone
	dim blnAnyFailed
	dim strFullNotify
	dim NotifyArray
	dim NotifyAddress


	if request("NextMilestoneID") = "" or request("NextMilestoneID") = "0" then
		WorkflowComplete = 1
		strNextMilestone = request("SelectedMilestone")
	else
		WorkflowComplete = 0
		strNextMilestone = request("NextMilestoneID")
	end if

	dim strSubject
	dim strBody
	dim strTo
	dim rs 
	dim strComments	
	dim strTemp
	dim strReleaseLanguages
	dim CurrentUser
	dim currentuserid
	dim CurrentUserEmail
	dim strSupported
	dim strItem
	dim strDeliverableReleaseString
	Response.Write "."
	'Create Database Connection
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.ConnectionTimeout = 180
	cn.Open


	'Get User
	dim CurrentDomain
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
		CurrentUserEmail = rs("Email")
		currentuserid = rs("ID")
	end if
	rs.Close

	Response.Write "."


	strSubject = ""
	strBody = ""
	strTo = ""
	strSupported = ""

	if 1 then 'request("txtNotify") <> "" then
	
	    on error resume next
		rs.Open "spGetOTSByDelVersion " & clng(request("DeliverableID")),cn,adOpenForwardOnly
        on error goto 0
        if cn.errors.count = 0 then
		strItem = ""
        strItemSummary = ""
		if not (rs.EOF and rs.BOF) then
			do while not rs.EOF
				strItem = strItem & ", " & rs("OTSNumber") 
				strItemSummary = strItemSummary & rs("OTSNumber") & " - " & rs("shortdescription") & "<BR>"
                rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3)
			end if
			strSupported = strSupported & "OBSERVATIONS FIXED:<BR>" & strItem & "<BR>-------------------------------------------------------------------------------------<BR>"
		    strSupported = strSupported & strItemSummary & "-------------------------------------------------------------------------------------<BR><BR>"
        end if
		
		rs.Close
        end if
		
		dim strDesktopProducts
		strDesktopProducts = "" 
		if instr(request("txtOldComments"),"DESKTOP PRODUCTS:") > 0 then
			strDesktopProducts = mid(request("txtOldComments"),instr(request("txtOldComments"),"DESKTOP PRODUCTS:")+18)
			if instr(strDesktopProducts,vbcrlf) then
				strDesktopProducts = left(strDesktopProducts,instr(strDesktopProducts,vbcrlf)-1)
			end if

			dim strDesktopPartNumber
			strDesktopPartNumber = "" 
			if instr(request("txtOldComments"),"PART NUMBER:") > 0 then
				strDesktopPartNumber = mid(request("txtOldComments"),instr(request("txtOldComments"),"PART NUMBER:"))
				if instr(strDesktopPartNumber,vbcrlf) then
					strDesktopPartNumber = left(strDesktopPartNumber,instr(strDesktopPartNumber,vbcrlf)-1)
				end if
			end if
		end if


		
		Response.Write "."
		rs.Open "spGetProductsForVersion " & clng(request("DeliverableID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF and  strDesktopProducts = "" then
			strSupported = strSupported & "PRODUCTS: Product Independent<BR>"
		else
			strItem = ""
			do while not rs.EOF
				strItem = strItem & ", " & rs("Family") & " " & rs("Version")
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3) 
				if strDesktopProducts <> "" then
					strItem = strItem & ", " & strDesktopProducts
				end if
			else
				strItem = strDesktopProducts
			end if
			strSupported = strSupported & "PRODUCTS: " & strItem & "<BR>"
		end if
		
		rs.Close
	

		Response.Write "."
		rs.Open "spGetSelectedOS " & clng(request("DeliverableID")),cn,adOpenForwardOnly
		strItem = ""
		if rs("ID")=16 and request("txtType")="1" then
			strSupported = strSupported
		elseif rs.EOF and rs.BOF then
			strSupported = strSupported & "OPERATING SYSTEMS: OS Independent<BR>"
		else
			do while not rs.EOF
				strItem = strItem & ", " & rs("Name") 
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3)
			end if
			if trim(strItem) <> "" then
				strSupported = strSupported & "OPERATING SYSTEMS: " & strItem & "<BR>"
			else
				strSupported = strSupported & "OPERATING SYSTEMS: OS Independent<BR>"
			end if
		end if
		
		rs.Close


		Response.Write "."
		rs.Open "spGetSelectedLanguages " & clng(request("DeliverableID")),cn,adOpenForwardOnly
		strItem = ""
		if rs.EOF and rs.BOF then
			strSupported = strSupported & "LANGUAGES: Language Independent<BR>"
		else
			do while not rs.EOF
				strItem = strItem & ", " & rs("Abbreviation") 
				rs.MoveNext
			loop
			if len(strItem)>0 then
				strItem = mid(strItem,3)
			end if
			if trim(request("txtHFCN")) = "true" then
				strSupported = strSupported & "COUNTRIES: " & strItem & "<BR>"
			else
				strSupported = strSupported & "LANGUAGES: " & strItem & "<BR>"
			end if
		end if
		
		rs.Close


	
		Response.Write "."
		rs.open "spGetVersionProperties4Web " & clng(request("DeliverableID")),cn,adOpenForwardOnly
	
		if rs.EOF and rs.BOF then
			strDeliverableReleaseString = ""
			strBody = "Deliverable " & request("DeliverableID") & " has been released, but the mail could not be sent."
			strTo = "max.yu@hp.com"
		else
		    strComments = ""
		    
		    If trim(request("txtTestRecommendations")) <> "" Then
		        strComments = strComments & "TEST RECOMMENDATIONS:" & vbcrlf & request("txtTestRecommendations") & vbcrlf & vbcrlf
		    End If
		    
		    If trim(request("txtSampleNotes")) <> "" Then
		        strComments = strComments & "SAMPLE NOTES:" & vbcrlf & request("txtSampleNotes") & vbcrlf & vbcrlf
		    End If
		    
			if trim(request("txtComments"))<> "" then
				strComments = strComments &  propercase(request("SelectedMilestoneName")) & " Comments:" & vbcrlf & request("txtComments") & vbcrlf & vbcrlf &  request("txtOldComments") 
			else
				strComments = strComments & request("txtOldComments")
			end if
			
			strDeliverableReleaseString = request("SelectedMilestoneName") & ": " & rs("Name") & " " & rs("Version")
			
			if rs("Revision") & "" <> "" then
				strDeliverableReleaseString = strDeliverableReleaseString & "," & rs("Revision")
			end if
			if rs("Pass") & "" <> "" then
				strDeliverableReleaseString = strDeliverableReleaseString & "," & rs("Pass")
			end if
			
			strBody = ""
			strBody = strBody & "ID: " & request("DeliverableID") & "<BR>"
			strBody = strBody & "NAME: " & request("DeliverableName") & "<BR>"
			if trim(request("txtType")) = "1" then
				if rs("PartNumber") & "" <> "" then
					strBody = strBody & "PART NUMBER: " & rs("PartNumber") & "<BR>"
				end if

			    if rs("VENDOR") & "" <> "" and lCase(rs("VENDOR")) & "" <> "hp" and lCase(rs("VENDOR")) & "" <> "compaq" then
				    strBody = strBody & "VENDOR: " & rs("VersionVendor") & "<BR>"
			    end if
			
				if rs("ModelNumber") & "" <> "" then
					strBody = strBody & "MODEL NUMBER: " & rs("ModelNumber") & "<BR>"
				end if
			
				strBody = strBody & "HW VERSION: " & rs("Version") & "<BR>"
			else
				strBody = strBody & "VERSION: " & rs("Version") & "<BR>"
			end if

			if rs("Revision") & "" <> "" then
				if trim(request("txtType")) = "1" then
					strBody = strBody & "FW VERSION: " & rs("Revision") & "<BR>"
				else
					strBody = strBody & "REVISION: " & rs("Revision") & "<BR>"				
				end if
			end if
			
			if rs("Pass") & "" <> "" then
				strBody = strBody & "PASS: " & rs("Pass") & "<BR>"
			end if

			if rs("vendorVersion") & "" <> "" then
				strBody = strBody & "VENDOR VERSION: " & rs("Vendorversion") & "<BR>"
			end if
			
			if rs("MD5") & "" <> "" then
				strBody = strBody & "MD5: " & rs("MD5") & "<BR>"
			end if

            strBody = strBody & "<BR>"

			if rs("Changes") & "" <> "" then
				strBody = strBody & "MODIFICATIONS/ENHANCEMENTS: " & rs("changes") & "<BR><BR>"
			end if

			if request("txtType")="1" then
				if request("txtHWLocation") <> "" then
					strBody = strBody & "DELIVERABLE LOCATION: " & request("txtHWLocation") & "<BR>"
				elseif rs("ImagePath") <> "" then
					strBody = strBody & "DELIVERABLE LOCATION: " & rs("ImagePath") & "<BR>"
				else
					strBody = strBody & "DELIVERABLE LOCATION: " & "N/A" & "<BR>"
				end if
			else
				dim TransferText
				if rs("AR")=1 then
					strBody = strBody & "TRANSFER PATH: This deliverable is not available online.  The deliverable can be obtained from the developer, Release Team, and/or directly from the replicater.<BR>"
				elseif trim(request("txtISOFilename")) <> "" then
				    TransferText = request("txtISOFilename")
					strBody = strBody & "TRANSFER PATH: " & TransferText & "<BR>"
				elseif request("txtTransfer") <> "" then
					TransferText = request("cboTransferServer")
					if right(TransferText,1) = "\" then
						TransferText = left(TransferText,len(TransferText)-1)
					end if
					if ucase(left(request("txtTransfer"),len(transferText))) = ucase(TransferText) then
						TransferText = request("txtTransfer")
					elseif left(request("txtTransfer"),1) = "\" then
						TransferText = TransferText & request("txtTransfer")
					else
						TransferText = TransferText & "\" & request("txtTransfer")
					end if
					strBody = strBody & "TRANSFER PATH: " & TransferText & "<BR>"
				end if
			end if			

			if rs("CDPartNumber") & "" <> "" then
				strBody = strBody & "CD PART NUMBER: " & rs("CDPartNumber") & "<BR>"
			end if
			
			if strDesktopPartNumber & "" <> "" then
				strBody = strBody & strDesktopPartNumber & "<BR>"
			end if

			if trim(request("txtType")) = "1" then
				if rs("SampleDate") & "" <> "" then
					strBody = strBody & "SAMPLES AVAILABLE: " & rs("SampleDate") & "<BR><BR>"
				end if
			end if 
			
			if rs("DevManager") & "" <> "" then
				strBody = strBody & "DEVELOPMENT MANAGER: " & rs("devmanager") & "<BR>"
			end if

            if trim(request("txtDeveloperID")) <> "" and trim(request("txtDeveloperID")) <> "0" and (not blnFailed) and request("txtType")="1" then
				strBody = strBody & "DEVELOPER: " & request("txtDeveloperName") & "<BR><BR>"
			elseif rs("developer") & "" <> "" then
				strBody = strBody & "DEVELOPER: " & rs("developer") & "<BR><BR>"
			end if

			if rs("InstallableUpdate") & ""  = "1" then
				strBody = strBody & "SPECIAL NOTES: Installable Update" & "<BR><BR>"
			end if
		
		
			strBody = strBody & strSupported & "<BR><BR>"
			
			if rs("TypeID") = 3 then 'FW
				strTemp = ""
				if rs("CAB") then
					strTemp = strTemp & ",CAB"
				end if			
				if rs("Rompaq") then
					strTemp = strTemp & ",Rompaq"
				end if			
				if rs("PreinstallROM") then
					strTemp = strTemp & ",Preinstall"
				end if			
				if rs("Binary") then
					strTemp = strTemp & ",Binary"
				end if		
				if strTemp <> "" then
					strTemp=mid(strTemp,2)
				end if
				strBody = strBody & "ROM COMPONENTS: " & strTemp & "<BR><BR>"
			else
				strTemp = ""
				if rs("Preinstall") then
					strTemp = strTemp & ",Preinstall"
				end if			
				if rs("AR") then
					strTemp = strTemp & ",Replicator Only"
				end if			
				if rs("floppydisk") then
					strTemp = strTemp & ",Diskette"
				end if			
				if rs("CDIMage") then
					strTemp = strTemp & ",CD Image"
				end if			
				if rs("Scriptpaq") then
					strTemp = strTemp & ",Scriptpaq"
				end if		
				if strTemp <> "" then
					strTemp=mid(strTemp,2)
					strBody = strBody & "DISTRIBUTION METHODS: " & strTemp & "<BR><BR>"
				end if
			end if
			
			
			if trim(strComments) <> "" then
				strBody = strBody & "COMMENTS: <BR>" & strComments & "<BR><BR>"
			else 
				strBody = strBody & "COMMENTS: N/A <BR>" & "<BR><BR>"
			end if
			

			if rs("Description") <> "" then
				strBody = strBody & "DESCRIPTION: <BR>" & rs("Description") & "<BR><BR>"
			end if
			

			if trim(rs("Notes") & "") <> "" then
				strBody = strBody & "NOTES: <BR>" & rs("Notes") & "<BR><BR>"
			end if
			
			dim strReleasePriorityText
			select case trim(request("cboReleasePriority"))
			case "1"
				strReleasePriorityText = "High"
			case "2"
				strReleasePriorityText = "Normal"
			case "3"
				strReleasePriorityText = "After-hours Support"
			case else
				strReleasePriorityText = trim(request("cboReleasePriority"))
			end select
			if lcase(request("NextMilestoneNameset cn")) = "release team" then
				strBody = strBody & "RELEASE PRIORITY: " & strReleasePriorityText & "<BR>"
				if request("txtReleasePriorityJust") <> "" then
					strBody = strBody & "RELEASE JUSTIFICATION: " & request("txtReleasePriorityJust") & "<BR><BR>"
				end if
			end if

		end if
		rs.close
		set rs=nothing
	end if
	Response.Write "."

'    Dim cn2
'    set cn2 = Server.CreateObject("ADODB.Connection")	
'	cn2.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
'	cn2.ConnectionTimeout = 180
'	cn2.Open
 '   cn2.BeginTrans
    
	FoundErrors = false	
	cn.BeginTrans

	Response.write "<BR>>" & request("LanguageList") & "<<BR>"
	LangArray = split(request("LanguageList"),",")
	blnFailed = false
	blnAllReleased	= true
	blnAllFailed = true
	blnAnyReleased = false
	blnAnyFailed = false
	strReleaseLanguages = ""
	Response.Write ubound(Langarray) & "<BR>"
	for i = 0 to ubound(Langarray)
		Lang = langArray(i)
		Response.Write Lang & ":"
		Response.Write request("cbo" & lang & "Release")  & "<BR>"
		if trim(Lang) <> "" then
			Select case request("cbo" & lang & "Release") 
				case "0"
					cn.Execute "spupdateLanguageLocation " & clng(request("DeliverableID")) & "," & clng(langArray(i+1)) & "," & clng(request("SelectedMilestone")) & "," & 0 & ",0",RowsEffected
					blnAllReleased = false
					blnAllFailed = false
				case "1"
					cn.Execute "spupdateLanguageLocation " & clng(request("DeliverableID")) & "," & clng(langArray(i+1)) & "," & clng(strNextMilestone) & "," & clng(WorkflowComplete) & ",0",RowsEffected
					blnAllFailed = false
					blnAnyReleased = true
					strReleaseLanguages = strReleaseLanguages & ", " & lang
				case "2"
					cn.Execute "spupdateLanguageLocation " & clng(request("DeliverableID")) & "," & clng(langArray(i+1)) & "," & clng(request("SelectedMilestone")) & "," & 0 & ",1",RowsEffected
					blnAllReleased = false
					blnAnyFailed = true
			end select
			if RowsEffected <> 1 then
				blnfailed = true
				exit for
			end if
		end if
	next 

	Response.Write "."

	'Add Server
	if request("txtType")<>"1" then

   		set cm = server.CreateObject("ADODB.Command")
					            
	    cm.ActiveConnection = cn
	    cm.CommandText = "spAddServer"
	    cm.CommandType = &H0004
		                
	    dim ServerName
	    ServerName = left(request("cboTransferServer"),255)
	        
	    Set p = cm.CreateParameter("@Name",200, &H0001,255)
		if right(ServerName,1) = "\" then
			p.Value = left(ServerName,len(ServerName)-1)
		else
			p.Value = ServerName
	    end if
	    cm.Parameters.Append p
	
	    cm.Execute recordseffected
		Set cm = Nothing

		Response.Write "."
	
	end if
	
	if strReleaseLanguages	 <> "" and trim(request("txtType")) <> "1" then
		strBody = strBody & "LANGUAGES RELEASED: " & mid(strReleaseLanguages,3)
	end if

	if (not blnFailed) and (not blnAllFailed) then
		cn.Execute "spresumeafterfailure " & clng(request("DeliverableID")) & "," & clng(request("SelectedMilestone")) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnfailed = true
		end if
	end if


	Response.Write "."
	if (not blnFailed) and blnAllReleased then
		cn.Execute "spupdatemilestonerelease " & clng(request("DeliverableID")) & "," & clng(request("SelectedMilestone")) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnfailed = true
		end if
	elseif (not blnFailed) and blnAnyReleased then
		if request("NextMilestoneID") <> "" and request("NextMilestoneID") <> "0" then
			cn.Execute "spupdatemilestoneStart " & clng(request("DeliverableID")) & "," & clng(request("NextMilestoneID")) ,RowsEffected		
			if cn.Errors.count > 0 then
				blnfailed = true
			end if
		end if
	end if
	


	Dim p
	dim cm
	Response.Write "."
	if (not blnFailed) then
		set cm = server.CreateObject("ADODB.Command")
					            
		cm.ActiveConnection = cn
		cm.CommandText = "spUpdateDeliverable4Release"
		cm.CommandType = &H0004
		                
		Set p = cm.CreateParameter("@DelID", 3, &H0001)
		p.Value = clng(request("DeliverableID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@NextMilestoneID", 3, &H0001)
		if request("NextMilestoneID") <> "" then
			p.Value = clng(request("NextMilestoneID"))
		else
			p.value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Filename", 200, &H0001,60)
		p.Value = left(request("txtFilename"),60)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Transfer", 200, &H0001,255)
		if request("txtType")="1" then
			p.Value = left(request("txtHWLocation"),255)
		else
			p.Value = left(TransferText,255)
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@MD5", 200, &H0001,50)
		p.Value = left(request("txtMD5"),50)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@BaseFilePath", 200, &H0001,1024)
		p.Value = ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@CVASubPath", 200, &H0001,255)
		p.Value = ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ReleasePriority", 16, &H0001)
		p.Value = clng(request("cboReleasePriority"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ReleasePriorityJustification", 200, &H0001,80)
		p.Value = left(request("txtReleasePriorityJust"),80)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Comments", 201, &H0001, 2147483647)
		if trim(request("txtComments")) <> "" then
			p.value =  propercase(request("SelectedMilestoneName")) & " Comments:" & vbcrlf & request("txtComments") & vbcrlf & vbcrlf & request("txtOldComments")
		else
			p.value = request("txtOldComments")
		end if
		cm.Parameters.Append p
	            
		cm.Execute RowsEffected
	
		If cn.Errors.count > 0 Then
			blnFailed = true
		End If
		set cm = nothing
		set p = nothing
	end if
	
	Response.Write "."
	if (not blnFailed) and blnAnyFailed then
		cn.Execute "spupdatemilestonefailure " & clng(request("DeliverableID")) & "," & clng(request("SelectedMilestone")) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnfailed = true
		end if
	end if

	if request("SelectedMilestoneName") = "RELEASE TEAM" and (not blnAllFailed) then
		'Response.Write "."
		'if (not blnFailed) then
		'	cn.Execute "spUpdateProdVerDistribution " & clng(request("DeliverableID")) & ",1" ,RowsEffected		
		'	if cn.Errors.count > 0 then
		'		blnfailed = true
		'	end if
		'end if

		if (not blnFailed) then
			cn.Execute "spPreReleaseVersion " & clng(request("DeliverableID")) ,RowsEffected		
			if cn.Errors.count > 0 then
				blnfailed = true
			end if
		end if
	end if

    if (not blnFailed) and trim(WorkflowComplete) = "1" then
		cn.Execute "spLogReleaseToProducts " & clng(request("DeliverableID")) & "," & clng(currentuserid) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnFailed = true
		end if
    end if

	if (not blnFailed) then
		cn.Execute "spUpdateDeliverableLocation " & clng(request("DeliverableID")) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnFailed = true
		end if
	end if


	if (not blnFailed) and lcase(request("NextMilestoneName")) = "release team" and blnAnyReleased then
		cn.Execute "spUpdateReleaseStatuses " & clng(request("DeliverableID")) ,RowsEffected		
		if cn.Errors.count > 0 then
			blnFailed = true
		end if
	end if


	if blnFailed then
		cn.RollbackTrans
		'cn2.RollbackTrans
		Response.Write "<BR>Unable to release this version."
		Response.Write "<INPUT type=""hidden"" id=txtSuccess name=txtSuccess value=""0"">"
	else
		cn.CommitTrans
		'cn2.CommitTrans
		Response.Write "<INPUT type=""hidden"" id=txtSuccess name=txtSuccess value=""1"">"

		if blnAllReleased then
			'if trim(request("txtHFCN")) = "true" then
			'	strSubject = "NEW HFCN: Deliverable Released from " & strDeliverableReleaseString
			'else
				strSubject = "Deliverable Released from " & strDeliverableReleaseString
			'end if
		elseif blnAllFailed then
			strSubject = "Deliverable Failed or Cancelled by " & strDeliverableReleaseString
		elseif blnAnyFailed and blnAnyReleased then
			strSubject = "Deliverable Released (Some Languages Failed) from " & strDeliverableReleaseString
		elseif blnAnyFailed  then
			strSubject = "Deliverable Failed by " & strDeliverableReleaseString
		elseif blnAnyReleased  then
			strSubject = "Deliverable Released from " & strDeliverableReleaseString
		end if


	    if trim(request("txtDeveloperID")) <> "" and trim(request("txtDeveloperID")) <> "0" and (not blnFailed) and request("txtType")="1" then
    	    cn.execute "spUpdateDeliverableDeveloper " & clng(request("DeliverableID")) & "," & clng(request("txtDeveloperID")),RowsEffected
            if RowsEffected <> 1 then
                blnFailed = true
            end if
        end if

		if blnAnyReleased or blnAnyFailed then
			if strTo = "" then 'Error message could be sent to admin.  If so, don't override with regular notification list
   				strFullNotify = ""
   				NotifyArray = split(request("txtNotify"),";")
   				for each NotifyAddress in NotifyArray
		   			if trim(NotifyAddress) <> "" then
   						strFullNotify = strFullNotify & ";" & NotifyAddress
   					end if
   				next
				if strFullNotify = "" then
					strFullNotify = "max.yu@hp.com" 
				else
					strFullNotify = mid(strFullNotify,2)
				end if
   				strTo = strFullNotify 
			end if	
		end if

	'	if strTO <> "" then
		
			if trim(CurrentUserEmail) = "" then
				CurrentUserEmail = "max.yu@hp.com"
			end if
			
   			dim strMailHeader	
  		
			strMailHeader = "<font face=Arial size=2>"
			if blnAnyFailed and trim(request("txtComments")) <> "" then
				strMailHeader = strMailHeader & trim(server.HTMLEncode(request("txtComments"))) & "<BR><BR>"
			end if	
			strMailHeader = strMailHeader & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur.asp"">Open Pulsar</a></font><br><br>"
			strMailHeader = strMailHeader & "<font face=Arial size=2>"
				
			strBody = strMailHeader & replace(replace(strBody,vbcrlf,"<BR>"),"""","&QUOT;")
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")		

			oMessage.From = CurrentUserEmail

			if trim(strTo) = "" then
				oMessage.To= "max.yu@hp.com"
			else
				oMessage.To= replace(strTo,";;",";")
			end if
			'oMessage.CC = "max.yu@hp.com"
			oMessage.Subject = strSubject
				
			oMessage.HTMLBody = strBody
			oMessage.DSNOptions = cdoDSNFailure
			'oMessage.BCC = "max.yu@hp.com"
			oMessage.Send 
			Set oMessage = Nothing 
   	'	end if

        if left(lcase(trim(strDeliverableReleaseString)),11) = "development" then
             'AddToMsgQueue "9," & clng(request("DeliverableID")) & ",0,'Send Deliverable Version To IRS - Release'"
             cn.execute "usp_SSSB_SendSync_Message 'MSMQLegacy', 9, " & clng(request("DeliverableID")) & ",0,0,'Send Deliverable Version To IRS - Release(ServiceBroker)'"
        end if

	end if

	cn.Close
	set cn=nothing
	'cn2.Close
	'set cn2 = nothing

	Response.Write "."

    Sub AddToMsgQueue (strMessage)
        Dim objQInfo
        Dim objQSend
        Dim objMessage
     
        'open the queue
        Set objQInfo = Server.CreateObject("MSMQ.MSMQQueueInfo")
        objQInfo.PathName = ".\private$\SIExcalSync" 
        Set objQSend = objQInfo.Open(2, 0)
  
        'build/send the message
        Set objMessage = Server.CreateObject("MSMQ.MSMQMessage")
        objMessage.Body = "<?xml version=""1.0""?><string>" & strMessage & "</string>"
        objMessage.Send objQSend
        objQSend.Close

        'clean up
        Set objQInfo = Nothing
        Set objQSend = Nothing
        Set objMessage = Nothing
    end sub
%>


</BODY>
</HTML>
