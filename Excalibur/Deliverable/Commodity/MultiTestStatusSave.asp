<%@ Language=VBScript %>
<!-- #include file="../../includes/EmailQueue.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<%if request("Type") = 1 then%>
<TITLE>Target Version</TITLE>
<%else%>
<TITLE>Reject Version</TITLE>
<%end if%>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined")
	{
	    if (txtSuccess.value == "1") {
	        //close window
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            parent.window.parent.closeExternalPopup();
	        }
	        else if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.parent.PulsarPlusCallback(txtSuccess.value);	            
	            ClosePulsarPlusPopup()
	        }
	        else {
	            window.parent.closewindow(false);
	        }        
	       
	    }
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">
<BR>
<table width=100%><TR><TD align=center>
<!--<font face=verdana size =2>Processing Request.  Please wait...</font>-->
</td></tr></table>
<%


	
	
	dim cn
	dim rs
	dim rs2
	dim IDArray
	dim CurrentDomain
	dim CurrentUserID
	dim CurrentUser
	dim UserSaveName
	dim cm
	dim rowsupdated
	dim strFrom
	dim strBody
	dim strTo
	dim strSubject
	dim blnIsTDCCNB
    dim blnIsTDCBNB
	dim strPMs
	dim strDevelopers
	dim VersionIDs
	dim VersionArray
	dim ID
	dim strTargetNotes
	dim blnPilotEngineer
	dim blnHardwarePM
	dim blnAccessoryPM
    dim blnCommodityPM
	dim blnPilotStatusChanged
	dim blnAccessoryStatusChanged
	dim blnQualStatusChanged
	dim strStatusText
	dim strQualTo
	dim strQualSubject
	dim strQualBody
	dim strPilotTo
	dim strPilotSubject
	dim strPilotBody
	dim strAccessoryTo
	dim strAccessorySubject
	dim strAccessoryBody
	dim strQCompleteListHP
	dim strQualBodyInventec
	dim strQualBodyQuanta
	dim strQualBodyCompal
	dim strQualBodyWistron
	dim strQualBodyDevelopers
	dim strPilotBodyInventec
	dim strPilotBodyQuanta
	dim strPilotBodyCompal
	dim strPilotBodyWistron
	dim strWWANCell
	dim strRow
	dim strFailedListODM
	dim TestStatusArray
	
  	TestStatusArray = split("TBD,Passed,Failed,Blocked,Watch,N/A",",")

	strQCompleteListHP = "TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;claire.lin@hp.com"
	
	
	strQualTo = ""
	strQualSubject = ""
	strQualBody = ""
	strPilotTo = ""
	strPilotSubject = ""
	strPilotBody = ""
	strAccessoryTo = ""
	strAccessorySubject = ""
	strAccessoryBody = ""
	strQualBodyInventec = ""
	strQualBodyQuanta = ""
	strQualBodyCompal = ""
	strQualBodyWistron = ""
	strQualBodyDevelopers = ""
	strPilotBodyInventec = ""
	strPilotBodyQuanta = ""
	strPilotBodyCompal = ""
	strPilotBodyWistron = ""
	
	strFailedListODM= ""
	
	blnCommodityPM = false

	blnIsTDCCNB = false
    blnIsTDCBNB = false
	blnPilotStatusChanged = false
	blnAccessoryStatusChanged = false
	blnQualStatusChanged = false

  


	if  request("txtMultiID") = "" or (request("NewValue") = "" and request("NewPilotValue") = "" and request("NewAccessoryValue") = "") then
		Response.Write "<BR><font face=verdana size=2>Not enough information supplied</font><BR>" 
		strSuccess = 0

	else

		set cn = server.CreateObject("ADODB.Connection")
		set rs = server.CreateObject("ADODB.recordset")
		set rs2 = server.CreateObject("ADODB.recordset")
	
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open


		'Get User
		CurrentUser = lcase(Session("LoggedInUser"))

		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

		UserSaveName = CurrentDomain & "_" & CurrentUser

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
	
		if rs.EOF and rs.BOF then
			CurrentUserID = 0 
			strFrom = "max.yu@hp.com"
		else
			CurrentUserID = rs("ID") 
			strFrom = rs("Email")
			if rs("CommodityPM") or rs("engcoordinator") or rs("ServicePM") then
				blnCommodityPM = true ' Check to see if the person is a Commodity PM (with or without an assigned product)
			end if
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
			if rs("SCFactoryEngineer") or trim(request("txtPilotEngineer")) = "True" then
				blnPilotEngineer = true
			else
				blnPilotEngineer = false
			end if
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
			if rs("AccessoryPM") or trim(request("txtAccessoryPM")) = "True" then
				blnAccessoryPM = true
			else
				blnAccessoryPM = false
			end if
			
		end if
		rs.Close
	
		blnHardwarePM = false
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if blnCommodityPM then
			blnHardwarePM = true
		else
			rs.Open "spGetHardwareTeamAccessList " & CurrentUserID,cn,adOpenStatic
			do while not rs.EOF
				if rs("Products") > 0 then
					blnHardwarePM = true
					exit do
				end if
				rs.MoveNext
			loop
			rs.Close
		end if		
	

		if CurrentUserID = 0 then
			strSuccess = "0"
		else
		
			if blnAccessoryPM then
				strBody=""

				strSQL = "Select pd.id, v.id  as ProductID, d.id as VersionID , v.DevCenter, pd.accessorystatusid,pd.pilotstatusid, pd.teststatusid, v.dotsname as Product, pd.targetnotes, r.name as Deliverable, d.version, d.revision, d.pass, d.modelnumber, d.partnumber, vd.name as vendor, e.email as Developer, e2.email as DevManager " & _
						  "from product_Deliverable pd with (NOLOCK), productversion v with (NOLOCK), deliverableversion d with (NOLOCK), deliverableroot r with (NOLOCK), vendor vd with (NOLOCK), employee e with (NOLOCK), employee e2 with (NOLOCK) " & _
						  "where pd.productversionid = v.id " & _
						  "and pd.deliverableversionid = d.id " & _
						  "and r.id = d.deliverablerootid " & _
						  "and vd.id = d.vendorid " & _
						  "and d.developerid = e.id " & _
						  "and r.devmanagerid = e2.id " & _
						  "and pd.id in (" & request("txtMultiID") & ") " 

				VersionIDs = ""
				strDevelopers = ""
				strPMs = ""
				
				rs.open strSQL,cn,adOpenForwardOnly
				do while not rs.EOF
					if trim(rs("AccessoryStatusID") & "") <> request("NewAccessoryValue")  and trim(request("NewAccessoryValue")) <> ""then
						if trim(rs("TargetNotes") & "") = "" then
							strTargetNotes = "No Comments Entered"
						else
							strTargetNotes = trim(rs("TargetNotes") & "")
						end if
					
						blnAccessoryStatusChanged = trim(request("NewAccessoryValue")) = "3"  or trim(request("NewAccessoryValue")) = "4" or trim(request("NewAccessoryValue")) = "6" or trim(request("NewAccessoryValue")) = "5"
							
						if blnAccessoryStatusChanged then 
							VersionIDs = VersionIDs & "," & rs("VersionID")
							if instr(strDevelopers,rs("Developer") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("Developer")
							end if
							if instr(strDevelopers,rs("DevManager") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("DevManager")
							end if
							strVersion = rs("Version") & ""
							if rs("Revision")&"" <> "" then
								strVersion = strVersion & "," & rs("Revision")
							end if
							if rs("Pass")&"" <> "" then
								strVersion = strVersion & "," & rs("Pass")
							end if
							strBody = strBody & "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
							strBody = strBody & "<TD nowrap>" & rs("Product") & "</TD>"
							strBody = strBody & "<TD>" & rs("Vendor") & "</TD>"
							strBody = strBody & "<TD>" & rs("Deliverable") & "</TD>"
							strBody = strBody & "<TD nowrap>" & strVersion & "&nbsp;</TD>"
							strBody = strBody & "<TD>" & rs("Modelnumber") & "&nbsp;</TD>"
							strBody = strBody & "<TD nowrap>" & rs("Partnumber") & "&nbsp;</TD>"
							strBody = strBody & "<TD>" & strTargetNotes & "&nbsp;</TD>"
							strBody = strBody & "</TR>"
						end if
							
					end if
					rs.MoveNext
				loop
				rs.Close

				if trim(strDevelopers)<> "" then
					strDevelopers = mid(strDevelopers,2)
				end if
				if trim(VersionIDs)<> "" then
					VersionIDs = mid(VersionIDs,2)
				end if

				VersionArray = split(VersionIDs,",")

				strPMs = ""
				for each ID in VersionArray
					rs.Open "spListAccessoryPMs4Version " & ID,cn,adOpenStatic
					do while not rs.EOF
						if instr(strPMs,rs("Email")&"")=0 then
							strPMs = strPMs	& ";" & rs("Email")
						end if
						rs.MoveNext
					loop
					rs.Close
				next
					
				if trim(strPMs) = "" then
					strPMs = "max.yu@hp.com"
				else
					strPMs = mid(strPMs,2)
				end if


				strStatusText = ""
				if strBody <> "" then ' and (trim(request("NewValue")) = "5" or trim(request("NewValue")) = "10") then 'Send Emails
					'Lookup Status
					rs.Open "spGetAccessoryStatus2 " & clng(request("NewAccessoryValue")),cn,adOpenStatic
					if rs.EOF and rs.BOF then
						strStatusText = " a new status "
					else
						strStatusText = rs("name") & ""
					end if
					rs.Close
					'Set To
					strTo = "max.yu@hp.com"
					'if strPMs <> "" then
					'	strTo = strTo & ";" & strPMs
					'end if
					'if strDevelopers <> "" then
					'	strTo = strTo & ";" & strDevelopers
					'end if
					
					'Set Subject
					Response.Write ">" & request("NewPilotValue") & "<"
					'Response.Flush
					if clng(request("NewAccessoryValue")) = 3 then
						strSubject = "Accessory Hold Notification"
					elseif clng(request("NewAccessoryValue")) = 4 then
						strSubject = "Accessory Cancellation Notification"
					elseif clng(request("NewAccessoryValue")) = 5 then
						strSubject = "Accessory Failure Notification"
					elseif clng(request("NewAccessoryValue")) = 6 then
						strSubject = "Accessory Complete Notification"
					else
						strSubject = "Accessory Status Updated"
					end if

					strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strBody & "</table>"
					strBody = "<font size=2 face=verdana color=black><b>The Accessory status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR></font>" & strBody 

					strAccessoryTo = strTo
					strAccessorySubject = strSubject
					strAccessoryBody = strBody
				end if
				
			end if
			
			 '****************Accessory above, Pilot below***************
			if blnPilotEngineer then
				strBody=""

				strSQL = "Select v.partnerid, pd.id, v.id  as ProductID, d.id as VersionID , v.DevCenter, pd.pilotstatusid, pd.teststatusid, v.dotsname as Product, pd.pilotnotes, pd.targetnotes, r.name as Deliverable, d.version, d.revision, d.pass, d.modelnumber, d.partnumber, vd.name as vendor, e.email as Developer, e2.email as DevManager " & _
						  "from product_Deliverable pd with (NOLOCK), productversion v with (NOLOCK), deliverableversion d with (NOLOCK), deliverableroot r with (NOLOCK), vendor vd with (NOLOCK), employee e with (NOLOCK), employee e2 with (NOLOCK) " & _
						  "where pd.productversionid = v.id " & _
						  "and pd.deliverableversionid = d.id " & _
						  "and r.id = d.deliverablerootid " & _
						  "and vd.id = d.vendorid " & _
						  "and d.developerid = e.id " & _
						  "and r.devmanagerid = e2.id " & _
						  "and pd.id in (" & request("txtMultiID") & ") " 

				VersionIDs = ""
				strDevelopers = ""
				strPMs = ""
				
				rs.open strSQL,cn,adOpenForwardOnly
				do while not rs.EOF
					if trim(rs("PilotStatusID") & "") <> request("NewPilotValue")  and trim(request("NewPilotValue")) <> ""then
						if trim(rs("TargetNotes") & "") = "" then
							strTargetNotes = "No Comments Entered"
						else
							strTargetNotes = trim(rs("TargetNotes") & "")
						end if
					
						blnPilotStatusChanged = trim(request("NewPilotValue")) = "3"  or trim(request("NewPilotValue")) = "6" or trim(request("NewPilotValue")) = "4" or trim(request("NewPilotValue")) = "7" or trim(request("NewPilotValue")) = "5"
							
						if blnPilotStatusChanged then 
							VersionIDs = VersionIDs & "," & rs("VersionID")
							if instr(strDevelopers,rs("Developer") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("Developer")
							end if
							if instr(strDevelopers,rs("DevManager") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("DevManager")
							end if
							strVersion = rs("Version") & ""
							if rs("Revision")&"" <> "" then
								strVersion = strVersion & "," & rs("Revision")
							end if
							if rs("Pass")&"" <> "" then
								strVersion = strVersion & "," & rs("Pass")
							end if
							strRow = "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
							strRow = strRow & "<TD nowrap>" & rs("Product") & "</TD>"
							strRow = strRow & "<TD>" & rs("Vendor") & "</TD>"
							strRow = strRow & "<TD>" & rs("Deliverable") & "</TD>"
							strRow = strRow & "<TD nowrap>" & strVersion & "&nbsp;</TD>"
							strRow = strRow & "<TD>" & rs("Modelnumber") & "&nbsp;</TD>"
							strRow = strRow & "<TD nowrap>" & rs("Partnumber") & "&nbsp;</TD>"
							strRow = strRow & "<TD>" & strTargetNotes & "&nbsp;</TD>"
							strRow = strRow & "</TR>"
							strBody = strBody & strRow
							
							'Add the row to the correct ODM email body
							Select case trim(rs("PartnerID"))
								case "2"
									strPilotBodyInventec = strPilotBodyInventec & strRow
								case "4"
									strPilotBodyQuanta = strPilotBodyQuanta & strRow
								case "3"
									strPilotBodyCompal = strPilotBodyCompal & strRow
								case "7"
									strPilotBodyWistron =  strPilotBodyWistron & strRow
							end select

						end if
							
					end if
					rs.MoveNext
				loop
				rs.Close

				if trim(strDevelopers)<> "" then
					strDevelopers = mid(strDevelopers,2)
				end if
				if trim(VersionIDs)<> "" then
					VersionIDs = mid(VersionIDs,2)
				end if

				VersionArray = split(VersionIDs,",")

				strPMs = ""
				for each ID in VersionArray
					rs.Open "spListHardwarePMs4Version " & ID,cn,adOpenStatic
					do while not rs.EOF
						if instr(strPMs,rs("Email")&"")=0 then
							strPMs = strPMs	& ";" & rs("Email")
						end if
						rs.MoveNext
					loop
					rs.Close
				next
					
				if trim(strPMs) = "" then
					strPMs = "max.yu@hp.com"
				else
					strPMs = mid(strPMs,2)
				end if


				strStatusText = ""
				if strBody <> "" then ' and (trim(request("NewValue")) = "5" or trim(request("NewValue")) = "10") then 'Send Emails
					'Lookup Status
					rs.Open "spGetPilotStatus2 " & clng(request("NewPilotValue")),cn,adOpenStatic
					if rs.EOF and rs.BOF then
						strStatusText = " a new status "
					else
						strStatusText = rs("name") & ""
					end if
					rs.Close
					'Set To
					strTo = "TWNPDCNBCommodityTechnology@hp.com;kidwell.proceng@hp.com;GPS.Taiwan.NB.Buy-Sell@hp.com;NotebookCommodityPlanningTeam@hp.com;twnpdccnbcommoditypm@hp.com"
					if strPMs <> "" then
						strTo = strTo & ";" & strPMs
					end if
					if strDevelopers <> "" then
						strTo = strTo & ";" & strDevelopers
					end if
					
					'Set Subject
					if clng(request("NewPilotValue")) = 3 then
						strSubject = "Pilot Hold Notification"
					elseif clng(request("NewPilotValue")) = 4 then
						strSubject = "Pilot Cancellation Notification"
					elseif clng(request("NewPilotValue")) = 5 then
						strSubject = "Pilot Failure Notification"
					elseif clng(request("NewPilotValue")) = 6 then
						strSubject = "Pilot Complete Notification"
					elseif clng(request("NewPilotValue")) = 7 then
						strSubject = "Factory Hold Notification"
						strTo = strTo & ";notebook.npi.mm@hp.com"
					else
						strSubject = "Pilot Status Updated"
					end if

					if trim(request("txtPilotComments")) = "" then
						strInsertComments = ""
					else
						strInsertComments = "COMMENTS: " & request("txtPilotComments") & "<BR><BR>"
					end if

					strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strBody & "</table>"
					strBody = "<font size=2 face=verdana color=black><b>The Pilot status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR>" & strInsertComments & "</font>" & strBody 
					
					'Finish building each ODM notification to be sent

					if trim(strPilotBodyInventec) <> "" then
						strPilotBodyInventec = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strPilotBodyInventec & "</table>"
						strPilotBodyInventec = "<font size=2 face=verdana color=black><b>The Pilot status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR>" & strInsertComments & "</font>" & strPilotBodyInventec 
					end if
					if trim(strPilotBodyQuanta) <> "" then
						strPilotBodyQuanta = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strPilotBodyQuanta & "</table>"
						strPilotBodyQuanta = "<font size=2 face=verdana color=black><b>The Pilot status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR>" & strInsertComments & "</font>" & strPilotBodyQuanta 
					end if
					if trim(strPilotBodyCompal) <> "" then
						strPilotBodyCompal = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strPilotBodyCompal & "</table>"
						strPilotBodyCompal = "<font size=2 face=verdana color=black><b>The Pilot status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR>" & strInsertComments & "</font>" & strPilotBodyCompal 
					end if
					if trim(strPilotBodyWistron) <> "" then
						strPilotBodyWistron = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>Comments</b></TD>" & strPilotBodyWistron & "</table>"
						strPilotBodyWistron = "<font size=2 face=verdana color=black><b>The Pilot status of the following deliverables has been set to " & strStatusText & " on the listed products:</b><BR><BR>" & strInsertComments & "</font>" & strPilotBodyWistron 
					end if
					
					
					strPilotTo = strTo
					strPilotSubject = strSubject
					strPilotBody = strBody
				end if
				
			end if
			
			 '****************Pilot above, qual below***************
			
			if blnHardwarePM then
				strBody=""
				strTo = ""
				strSubject = ""

				strSQL = "Select d.tts, pd.wwanteststatus, pd.riskrelease, pd.id, v.id  as ProductID, d.id as VersionID , v.partnerid, v.DevCenter, pd.pilotstatusid, pd.teststatusid, v.dotsname as Product, pd.targetnotes, r.name as Deliverable, d.version, d.revision, d.pass, d.modelnumber, d.partnumber, vd.name as vendor, e.email as Developer, e2.email as DevManager, c.requiresWWANtestfinalapproval, v.wwanproduct " & _
						  "from product_Deliverable pd with (NOLOCK), productversion v with (NOLOCK), deliverableversion d with (NOLOCK), deliverableroot r with (NOLOCK), vendor vd with (NOLOCK), employee e with (NOLOCK), employee e2 with (NOLOCK), deliverablecategory c with (NOLOCK) " & _
						  "where pd.productversionid = v.id " & _
						  "and pd.deliverableversionid = d.id " & _
						  "and c.id = r.categoryid " & _
						  "and r.id = d.deliverablerootid " & _
						  "and vd.id = d.vendorid " & _
						  "and d.developerid = e.id " & _
						  "and r.devmanagerid = e2.id " & _
						  "and pd.id in (" & request("txtMultiID") & ") " 

				VersionIDs = ""
				strDevelopers = ""
				strPMs = ""
				
				rs.open strSQL,cn,adOpenForwardOnly
				do while not rs.EOF
					strWWANCell = ""
					if (trim(rs("TestStatusID") & "") <> request("NewValue") or (replace(trim(rs("riskrelease") & ""),"0","") <> replace(request("chkRiskRelease"),"on","1")  )  ) and trim(request("NewValue")) <> "" then
						if trim(rs("DevCenter") & "") = "2" then
							blnIsTDCCNB = true
						end if

                        if trim(rs("DevCenter") & "") = "3" then
							blnIsTDCBNB = true
						end if

                        if trim(request("txtTestComments")) <> "" then
							strTargetNotes = trim(request("txtTestComments"))
						elseif trim(rs("TargetNotes") & "") = "" then
							strTargetNotes = "No Comments Entered"
						else
							strTargetNotes = trim(rs("TargetNotes") & "")
						end if
					
						blnQualStatusChanged = (trim(request("NewValue")) = "6"  or trim(request("NewValue")) = "7" or trim(request("NewValue")) = "10" or trim(request("NewValue")) = "5" or trim(request("NewValue")) = "18" )
							
						if blnQualStatusChanged then 
							if not (rs("wwanproduct") and rs("requiresWWANtestfinalapproval")) then
								strWWANCell = "N/A"
							else
								strWWANCell = TestStatusArray(clng(rs("WWANTestStatus")))
							end if
				
							VersionIDs = VersionIDs & "," & rs("VersionID")
							if instr(strDevelopers,rs("Developer") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("Developer")
							end if
							if instr(strDevelopers,rs("DevManager") & "") = 0 then
								strDevelopers = strDevelopers & ";" & rs("DevManager")
							end if
							strVersion = rs("Version") & ""
							if rs("Revision")&"" <> "" then
								strVersion = strVersion & "," & rs("Revision")
							end if
							if rs("Pass")&"" <> "" then
								strVersion = strVersion & "," & rs("Pass")
							end if
							strRow = "<TR><TD><a href=""http://16.81.19.70/query/DeliverableVersionDetails.asp?ID=" & rs("VersionID") & """>" & rs("VersionID") & "</a></TD>"
							strRow = strRow & "<TD nowrap>" & rs("Product") & "</TD>"
							strRow = strRow & "<TD>" & rs("Vendor") & "</TD>"
							strRow = strRow & "<TD>" & rs("Deliverable") & "</TD>"
							strRow = strRow & "<TD nowrap>" & strVersion & "&nbsp;</TD>"
							strRow = strRow & "<TD>" & rs("Modelnumber") & "&nbsp;</TD>"
							strRow = strRow & "<TD nowrap>" & rs("Partnumber") & "&nbsp;</TD>"
							strRow = strRow & "<TD>" & strWWANCell & "&nbsp;</TD>"
							strRow = strRow & "<TD>" & strTargetNotes & "&nbsp;</TD>"
							strRow = strRow & "</TR>"
							
							strQualBodyDevelopers = strQualBodyDevelopers & strRow
							
							strBody = strBody & strRow

						if (blnIsTDCCNB and trim(request("NewValue")) = "5") or ((not blnIsTDCCNB) and (trim(request("NewValue")) = "6" or trim(request("NewValue")) = "7" or trim(request("NewValue")) = "10" or trim(request("NewValue")) = "18" or trim(request("NewValue")) = "5")) then
								Select case trim(rs("PartnerID"))
									case "2"
										strQualBodyInventec = strQualBodyInventec & strRow
									case "4"
										strQualBodyQuanta = strQualBodyQuanta & strRow
									case "3"
										strQualBodyCompal = strQualBodyCompal & strRow
									case "7"
										strQualBodyWistron =  strQualBodyWistron & strRow
								end select
							end if							
						end if
							
					end if
					rs.MoveNext
				loop
				rs.Close

				if trim(strDevelopers)<> "" then
					strDevelopers = mid(strDevelopers,2)
				end if
				if trim(VersionIDs)<> "" then
					VersionIDs = mid(VersionIDs,2)
				end if

				VersionArray = split(VersionIDs,",")

				strPMs = ""
				for each ID in VersionArray
					rs.Open "spListHardwarePMs4Version " & ID,cn,adOpenStatic
					do while not rs.EOF
						if instr(strPMs,rs("Email")&"")=0 then
							strPMs = strPMs	& ";" & rs("Email")
						end if
						rs.MoveNext
					loop
					rs.Close
				next
					
				if trim(strPMs) = "" then
					strPMs = "max.yu@hp.com"
				else
					strPMs = mid(strPMs,2)
				end if

				strStatusText = ""
				if strBody <> "" then ' and (trim(request("NewValue")) = "5" or trim(request("NewValue")) = "10") then 'Send Emails
					'Lookup Status
					rs.Open "spGetQualificationStatus " & clng(request("NewValue")),cn,adOpenStatic
					if rs.EOF and rs.BOF then
						strStatusText = " a new status "
					elseif request("chkRiskRelease") = "on" and rs("status") & "" = "QComplete" then
						strStatusText = "Risk Release"
					else
						strStatusText = rs("status") & ""
					end if
					rs.Close
					
					if trim(request("NewValue")) = "5" then 'Qcomplete
						if blnIsTDCCNB then 'DevCenter = 2
							strTo = strDevelopers & ";twnpdccnbcommoditypm@hp.com;tdcesmail@hp.com;" & strQCompleteListHP '& ";kenneth.berntsen@hp.com"
                        elseif blnIsTDCBNB then 'DevCenter = 3
                            strTo = strDevelopers & ";tdcesmail@hp.com;twinkle.k.s@hp.com;sridevi.s@hp.com;rajendran.m@hp.com" & strQCompleteListHP
						else
							strTo = strDevelopers & ";" & strQCompleteListHP
						end if

						if request("chkRiskRelease") = "on"  then
							strSubject = "Hardware Risk Release Notification"
						else
							strSubject = "Hardware QComplete Notification"
						end if
                    else 'Failed (or other negative status)
						'Set TO
						if blnIsTDCCNB then 'DevCenter = 2
							strTo = ";TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com;twnpdccnbcommoditypm@hp.com;tdcesmail@hp.com" 
						elseif blnIsTDCBNB then 'DevCenter = 3
                            strTo = ";TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com;tdcesmail@hp.com;twinkle.k.s@hp.com;sridevi.s@hp.com;rajendran.m@hp.com"
                        else
							strTo = ";TWNPDCNBCommodityTechnology@hp.com;APJ-RCTO.SC@hp.com;claire.lin@hp.com;kidwell.proceng@hp.com"
						end if
						
						if strPMs <> "" then
							strTo = strTo & ";" & strPMs
						end if
						if strDevelopers <> "" then
							strTo = strTo & ";" & strDevelopers
						end if
						if strTo <> "" then
							strTo= mid(strTo,2) 
						else
							strTo = "max.yu@hp.com"
						end if
						                   
						'Set Subject
						if clng(request("NewValue")) = 6 then
							strSubject = "Hardware Drop Notification"
                            strTo = strTo + ";tdcdtoedmfunction@hp.com"
						elseif clng(request("NewValue")) = 7 then
							strSubject = "Hardware Hold Notification"
						elseif clng(request("NewValue")) = 10 then
							strSubject = "Hardware Failure Notification"
						elseif clng(request("NewValue")) = 18 then
							strSubject = "Hardware Status Changed to 'Service Only'"
						else
							strSubject = "Hardware Status Updated"
						end if
					end if
					strBody = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strBody & "</table>"
					if clng(request("NewValue")) = 18 then
					    strBody = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strBody 
					else
					    strBody = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strBody 
					end if
					if trim(strQualBodyInventec) <> "" then
						strQualBodyInventec = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strQualBodyInventec & "</table>"
                        if clng(request("NewValue")) = 18 then
						    strQualBodyInventec = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strQualBodyInventec 
					    else
						    strQualBodyInventec = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strQualBodyInventec 
					    end if
					end if
					if trim(strQualBodyCompal) <> "" then
						strQualBodyCompal = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strQualBodyCompal & "</table>"
                        if clng(request("NewValue")) = 18 then
    						strQualBodyCompal = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strQualBodyCompal 
	                    else
    						strQualBodyCompal = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strQualBodyCompal 
	                    end if
					end if
					if trim(strQualBodyQuanta) <> "" then
						strQualBodyQuanta = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strQualBodyQuanta & "</table>"
                        if clng(request("NewValue")) = 18 then
    						strQualBodyQuanta = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strQualBodyQuanta 
	                    else
    						strQualBodyQuanta = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strQualBodyQuanta 
	                    end if
					end if
					if trim(strQualBodyWistron) <> "" then
						strQualBodyWistron = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strQualBodyWistron & "</table>"
                        if clng(request("NewValue")) = 18 then
    						strQualBodyWistron = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strQualBodyWistron 
	                    else
    						strQualBodyWistron = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strQualBodyWistron 
	                    end if
					end if
					if trim(strQualBodyDevelopers) <> "" then
						strQualBodyDevelopers = "<STYLE>TD{Font-Family:verdana;Font-Size:xx-small;}</STYLE><TABLE cellpadding=2 cellspacing=1 bgcolor=ivory border=1><TR bgcolor=beige><TD><b>ID</b></TD><TD><b>Product</b></TD><TD><b>Vendor</b></TD><TD><b>Deliverable</b></TD><TD><b>HW,FW,Rev</b></TD><TD><b>Model</b></TD><TD><b>Part</b></TD><TD><b>WWAN</b></TD><TD><b>Comments</b></TD>" & strQualBodyDevelopers & "</table>"
                        if clng(request("NewValue")) = 18 then
    						strQualBodyDevelopers = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR>Note: This component has been 'Dropped' for Manufacturing but is still being used by Service.<BR><BR></font>" & strQualBodyDevelopers 
	                    else
    						strQualBodyDevelopers = "<font size=2 face=verdana color=black><b>The Qualification Status on the following deliverables has been set to " & strStatusText& " on the listed products:</b><BR><BR></font>" & strQualBodyDevelopers 
	                    end if
					end if
					
					strQualTo = strTo
					strQualSubject = strSubject
					strQualBody = strBody
				end if
			
			end if
			
			'Save Updates	
			cn.BeginTrans
			
			ProcessArray = split(request("txtMultiID"),",")
    

			
			Response.Write "<BR><font face=verdana size=2>Processing: </font>"
			strSuccess = "1"
			for i = 0 to ubound(ProcessArray)
				strID = trim(ProcessArray(i))
	
				if blnHardwarePM and trim(request("NewValue")) <> "" then
					set cm = server.CreateObject("ADODB.Command")
			
					cm.ActiveConnection = cn
					cm.CommandText = "spUpdateCommodityStatus"
					cm.CommandType = &H0004
								
					Set p = cm.CreateParameter("@ID",3, &H0001)
					p.Value = clng(strID)
					cm.Parameters.Append p
				
					Set p = cm.CreateParameter("@Status",3, &H0001)
					p.Value = clng(request("NewValue"))
					cm.Parameters.Append p
						
					Set p = cm.CreateParameter("@RiskRelease", 16,  &H0001)
					if request("chkRiskRelease") = "on" then
						p.Value = 1
					else
						p.Value = 0
					end if
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserID",3, &H0001)
					p.Value = clng(CurrentUserID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserName",200, &H0001,80)
					p.Value = left(UserSaveName,80)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@DCRID",3, &H0001)
					p.Value = 0
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@SupplyChainRestriction", 16,  &H0001)
					p.Value = null
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@ConfigurationRestriction", 16,  &H0001)
					p.Value = null
					cm.Parameters.Append p
			
					Set p = cm.CreateParameter("@Comments",200, &H0001,255)
					p.Value = left(request("txtTestComments"),255)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@TestDate", 135,  &H0001)
					if isdate(request("txtTestDate")) then
						p.Value = cdate(request("txtTestDate"))
					else
						p.value = null
					end if
					cm.Parameters.Append p
		
					Set p = cm.CreateParameter("@TestConfidence",3, &H0001)
					if request("cboConfidence") = "" then
						p.Value = 1
					else
						p.Value = clng(request("cboConfidence"))
					end if
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@BatchMode",3, &H0001)
					p.Value = 1
					cm.Parameters.Append p
										
					cm.Execute rowsupdated
					Set cm = Nothing
					
					if cn.Errors.count <> 0 then
						Response.Write "<BR>Failed.<BR>"
						strSuccess = "0"
						cn.RollbackTrans
						Response.Write "<BR>Records were not saved correctly.<BR>"
						exit for
					else
						strSuccess = "1"
        
					end if
   
					'End of Commodity Status Change Section
				end if
				if blnPilotEngineer and trim(request("NewPilotValue")) <> "" then
					'Start of Pilot Change Section

					set cm = server.CreateObject("ADODB.Command")
			
					cm.ActiveConnection = cn
					cm.CommandText = "spUpdatePilotStatus"
					cm.CommandType = &H0004
								
					Set p = cm.CreateParameter("@ID",3, &H0001)
					p.Value = clng(strID)
					cm.Parameters.Append p
				
					Set p = cm.CreateParameter("@Status",3, &H0001)
					p.Value = clng(request("NewPilotValue"))
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserID",3, &H0001)
					p.Value = clng(CurrentUserID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserName",200, &H0001,80)
					p.Value = left(UserSaveName,80)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@Comments",200, &H0001,255)
					p.Value = left(request("txtPilotComments"),255)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@PilotDate", 135,  &H0001)
					if isdate(request("txtPilotDate")) then
						p.Value = cdate(request("txtPilotDate"))
					else
						p.value = null
					end if
					cm.Parameters.Append p
		
					Set p = cm.CreateParameter("@BatchMode",3, &H0001)
					p.Value = 1
					cm.Parameters.Append p
										
					cm.Execute rowsupdated
					Set cm = Nothing
					
					if cn.Errors.count <> 0 then
						Response.Write "<BR>Failed.<BR>"
						strSuccess = "0"
						cn.RollbackTrans
						Response.Write "<BR>Records were not saved correctly.<BR>"
						exit for
					else
						strSuccess = "1"
					end if

				
					'End of Pilot Change Section
				end if
				
				if blnAccessoryPM and trim(request("NewAccessoryValue")) <> "" then
					'Start of Accessory Change Section

					set cm = server.CreateObject("ADODB.Command")
			
					cm.ActiveConnection = cn
					cm.CommandText = "spUpdateAccessoryStatus"
					cm.CommandType = &H0004
								
					Set p = cm.CreateParameter("@ID",3, &H0001)
					p.Value = clng(strID)
					cm.Parameters.Append p
				
					Set p = cm.CreateParameter("@Status",3, &H0001)
					p.Value = clng(request("NewAccessoryValue"))
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserID",3, &H0001)
					p.Value = clng(CurrentUserID)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@UserName",200, &H0001,80)
					p.Value = left(UserSaveName,80)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@Comments",200, &H0001,255)
					p.Value = ""
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@KitNumber",200, &H0001,20)
					p.Value = ""
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@KitDescription",200, &H0001,120)
					p.Value = ""
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@AccessoryDate", 135,  &H0001)
					if isdate(request("txtAccessoryDate")) then
						p.Value = cdate(request("txtAccessoryDate"))
					else
						p.value = null
					end if
					cm.Parameters.Append p
		
					Set p = cm.CreateParameter("@BatchMode",3, &H0001)
					p.Value = 1
					cm.Parameters.Append p
										
					cm.Execute rowsupdated
					Set cm = Nothing
					
					if cn.Errors.count <> 0 then
						Response.Write "<BR>Failed.<BR>"
						strSuccess = "0"
						cn.RollbackTrans
						Response.Write "<BR>Records were not saved correctly.<BR>"
						exit for
					else
						strSuccess = "1"
					end if

				
					'End of Accessory Change Section
				end if
				
			next
		end if			

		if strSuccess = "1" then
			cn.CommitTrans
			Response.Write "<font face=verdana size=2>Records saved successfully.</font><BR>"
		end if

'		Response.Flush

		set rs2=nothing	            
		set rs = nothing
		cn.Close
		set cn = nothing

        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if strQualBody <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = strQualTo 
            oMessage.CC = "isc.sj@hp.com; "	
			oMessage.Subject = strQualSubject
			
'			if trim(request("NewValue")) = "10" then
'				oMessage.Importance = cdoHigh
'			end if
			
			if trim(request("NewValue")) = "5" then
				strQualBody = strQualBody & "<font size=1 face=verdana><BR><BR><BR><BR><b>Note:</b> The appropriate ODM contacts were notified of this change in separate emails.</font>"
			end if
			
				oMessage.HTMLBody = strQualBody
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if

		'Developers only - Qual Complete on first product
'		if trim(strQualBodyDevelopers) <> ""  and strSuccess = "1" then
'			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
'			'Set oMessage.Configuration = Application("CDO_Config")	
'			oMessage.From = strFrom
'			if currentuserid = 31 or currentuserid = 8 then
'				oMessage.To = "max.yu@hp.com"
'			else
'				oMessage.To = strDevelopers
'			end if
'			oMessage.Subject = strQualSubject
'			
''			if trim(request("NewValue")) = "10" then
''				oMessage.Importance = cdoHigh
''			end if
'			
'			strQualBodyDevelopers = strQualBodyDevelopers & "<font size=1 face=verdana><BR><BR><BR><BR><b>Note:</b> The appropriate HP and ODM supply chain contacts were notified of this change in separate emails.</font>"
'			
'			if currentuserid = 31 or currentuserid = 8 then
'				oMessage.HTMLBody = "TO: " & strDevelopers & "<BR>" & strQualBodyDevelopers
'			else
'				oMessage.HTMLBody = strQualBodyDevelopers
''			end if
'			oMessage.DSNOptions = cdoDSNFailure
'			oMessage.Send 
'			Set oMessage = Nothing 	

'		end if



		'Inventec 
		if strQualBodyInventec <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = "IPC-ED1@inventec.com;IPCHP-Excalibur@inventec.com"
			oMessage.Subject = strQualSubject
				oMessage.HTMLBody = strQualBodyInventec 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if

		'Compal
		if strQualBodyCompal <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = "A32KeyCommodity@compal.com"
			oMessage.Subject = strQualSubject
				oMessage.HTMLBody = strQualBodyCompal 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if
		
		'Quanta
		if strQualBodyQuanta <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = "TOPCommodityTeam@quantacn.com"
			oMessage.Subject = strQualSubject
				oMessage.HTMLBody = strQualBodyQuanta 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if
		
		'Wistron 
		if strQualBodyWistron <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = "max.yu@hp.com;key_commodity@wistron.com" 
			oMessage.Subject = strQualSubject
				oMessage.HTMLBody = strQualBodyWistron 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if

		if strPilotBody <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = strPilotTo 
			oMessage.Subject = strPilotSubject
			
'			if trim(request("NewValue")) = "10" then
'				oMessage.Importance = cdoHigh
'			end if
			strPilotBody  = strPilotBody  & "<font size=1 face=verdana><BR><BR><BR><BR><b>Note:</b> The appropriate ODM contacts were notified of this change in separate emails.</font>"
			
			oMessage.HTMLBody = strPilotBody 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

			'Inventec - Pilot
			if strPilotBodyInventec <> "" and strSuccess = "1" then
				Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
				'Set oMessage.Configuration = Application("CDO_Config")	
				oMessage.From = strFrom
					oMessage.To = "IPC-ED1@inventec.com;IPCHP-Excalibur@inventec.com"
				oMessage.Subject = strPilotSubject
					oMessage.HTMLBody = strPilotBodyInventec 
				oMessage.DSNOptions = cdoDSNFailure
				oMessage.Send 
				Set oMessage = Nothing 	

			end if

			'Compal - Pilot
			if strPilotBodyCompal <> "" and strSuccess = "1" then
				Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
				'Set oMessage.Configuration = Application("CDO_Config")	
				oMessage.From = strFrom
					oMessage.To = "A32KeyCommodity@compal.com" 
				oMessage.Subject = strPilotSubject
					oMessage.HTMLBody = strPilotBodyCompal 
				oMessage.DSNOptions = cdoDSNFailure
				oMessage.Send 
				Set oMessage = Nothing 	

			end if

			'Quanta - Pilot
			if strPilotBodyQuanta <> "" and strSuccess = "1" then
				Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
				'Set oMessage.Configuration = Application("CDO_Config")	
				oMessage.From = strFrom
					oMessage.To = "TOPCommodityTeam@quantacn.com" 
				oMessage.Subject = strPilotSubject
					oMessage.HTMLBody = strPilotBodyQuanta 
				oMessage.DSNOptions = cdoDSNFailure
				oMessage.Send 
				Set oMessage = Nothing 	

			end if
					
			'Wistron - Pilot
			if strPilotBodyWistron <> "" and strSuccess = "1" then
				Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
				'Set oMessage.Configuration = Application("CDO_Config")	
				oMessage.From = strFrom
					oMessage.To = "max.yu@hp.com;key_commodity@wistron.com" 
				oMessage.Subject = strPilotSubject
					oMessage.HTMLBody = strPilotBodyWistron 
				oMessage.DSNOptions = cdoDSNFailure
				oMessage.Send 
				Set oMessage = Nothing 	

			end if			


		end if
		
		
		
		if strAccessoryBody <> "" and strSuccess = "1" then
			Set oMessage = New EmailQueue 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")	
			oMessage.From = strFrom
				oMessage.To = strAccessoryTo 
			oMessage.Subject = strAccessorySubject
			
'			if trim(request("NewValue")) = "10" then
'				oMessage.Importance = cdoHigh
'			end if
			
			oMessage.HTMLBody = strAccessoryBody 
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 	

		end if		
		
	end if

	if request("Remaining") < 1 then
		strReload = "1"
	else
		strReload = "0"
	end if
%>

<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
<INPUT type="hidden" id=txtReload name=txtReload value="<%=strReload%>">
<INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=request("app")%>">
</BODY>
</HTML>




