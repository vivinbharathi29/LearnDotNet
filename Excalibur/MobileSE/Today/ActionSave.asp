<%@  language="VBScript" %>
<%Response.Buffer = True %>
<!-- #include file="../../includes/EmailQueue.asp" -->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
  <script src="../../Scripts/PulsarPlus.js"></script>
  <script id="clientEventHandlersJS" language="javascript">
<!--
function window_onload() {
    if (document.getElementById('layout').value == 'pulsar2') {
        alert('Change Request Updated Successfully');
        history.go(-1);
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
        window.parent.parent.parent.popupCallBack(1);
    }
    else {
        var iframeName = parent.window.name;
        if (iframeName != '') {
            parent.window.parent.ClosePropertiesDialog(1);
        } else if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel(true);
        } else if (DisplayedID.value == "") {
            SavePrompt.style.display = "none";
            Results.style.display = "";
            window.returnValue = 1;
        } else {
            SavePrompt.style.display = "none";
            Results.style.display = "";
            window.returnValue = 1;

            window.parent.opener = 'X';
            window.parent.open('', '_parent', '')
            window.parent.close();
        }
    }
}

  //-->
  </script>
</head>
<style type="text/css">
  A:link {
    color: blue;
  }

  A:visited {
    color: blue;
  }

  A:hover {
    color: red;
  }
</style>
<body bgcolor="Ivory" language="javascript" onload="return window_onload()">
  <%
Response.Write "<label id=SavePrompt><b><font size=3 face=verdana>Saving.  Please Wait...</b></font><br></label>"

'response.Flush

Function StripHTMLTag(ByVal sText)
   StripHTMLTag = ""
   fFound = False
   Response.Write sText & "<BR>" & vbcrlf
   Do While InStr(sText, "<")
      fFound = True
      StripHTMLTag = StripHTMLTag & " " & Left(sText, InStr(sText, "<")-1)
      strTag = lcase(trim(mid(sText,InStr(sText, "<"),InStr(sText, ">") - InStr(sText, "<")+1)))

'	  if strTag = "<b>" or strTag = "</b>" or strTag = "<i>" or strTag = "</i>" or strTag = "<u>" or strTag = "</u>" then
		if left(replace(ucase(strTag)," ",""),5) <> "<" & trim("FONT") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("FONT") and left(replace(ucase(strTag)," ",""),5) <> "<" & trim("SPAN") and left(replace(ucase(strTag)," ",""),6) <> "</" & trim("SPAN") and left(replace(ucase(strTag)," ",""),4) <> "<" & trim("DIV") and left(replace(ucase(strTag)," ",""),5) <> "</" & trim("DIV") and left(replace(ucase(strTag)," ",""),2) <> "<" & trim("P") and left(replace(ucase(strTag)," ",""),3) <> "</" & trim("P") then
			StripHTMLTag = StripHTMLTag & strTag
      end if

	  
      sText = MID(sText, InStr(sText, ">") + 1)

      
   Loop
   StripHTMLTag = StripHTMLTag & sText
   If Not fFound Then StripHTMLTag = sText
End Function

Function getUniqueItems(arrItems)
        Dim objDict, strItem

        Set objDict = Server.CreateObject("Scripting.Dictionary")

        For Each strItem in arrItems
            objDict.Item(strItem) = 1
        Next

        getUniqueItems = objDict.Keys
    End Function

	dim cn
	dim cm
	dim p
	dim DisplayedID
	dim strConnect
	dim strProducts
	dim strProductID
	dim strProductName
	dim CurrentUser 
	dim CurrentUserID
	dim strType
	dim CurrentUserName
	dim strTypeName 
	dim strChangeType
	dim strGEOS
	dim strStatus
	dim strJustification
	dim CurrentUserEmail
	dim oldSummary
    dim oldProduct
 	dim oldstatus
	dim oldOwner
	dim oldZsrpTargetDt
	dim oldZsrpActualDt
	dim oldZsrpRequired
	dim oldCoreTeam
	dim oldTargetDate
	dim oldNotify
	dim oldDescription
	dim oldAction
	dim oldResolution
	dim strNewOwner
    dim strPM
    dim strPMID
	dim strPMEmail
	dim strPCEmail
	dim strProgramMail
	dim oldOnStatus
	dim strApproverList
	dim blnUpdateApprovals
	dim strRestoreType
	dim strOSList
	dim strLanguageList
	dim CCTMailList
	dim blnSendDisapprovedEmail
	dim strChangeTypeCat
	dim strDescription
	dim strApproverIDs
	dim strApproverEmails
	dim strApprovers    
	dim CommodityPMEmail
	dim strProductStatusID
	dim strAddDCRNotificationList
	dim bZsrpReadyRequired : bZsrpReadyRequired = (request("chkZsrpRequired") = "on")
    dim bID : bID = (request("chkIDChange") = "on")
	dim bBios : bBios = (request("chkBiosChange") = "on")
	dim bSCR : bSCR = (Request("chkSwChange") = "on")
	dim bRequirement : bRequirement = (request("chkReqChange") = "on")
	dim bSku : bSku = (request("chkSKUChange") = "on")
	dim bSoftware : bSoftware = (request("chkImageChange") = "on")
    dim bCategoryBiosChange : bCategoryBiosChange = (request("chkCategoryBiosChange") = "on")
	dim bCommodity : bCommodity = (request("chkCommodityChange") = "on")
	dim bDocs : bDocs = (request("chkDocChange") = "on")
	dim bOther : bOther = (request("chkOtherChange") = "on")
	dim strRegions
	dim strCustomerImpact
	dim strDeliverableRootID
	dim strDeliverableRootName
	dim strDeliverableManagerID
	dim strODMBody
    dim strHPBody
    dim strHPEmailAddress
    dim strODMEmailAddress   
    dim strExcaliburHPLink       
    dim strExcaliburODMLink      
    Dim iChangeRequestID
    Dim sProductRelease
    Dim iIssueID
    Dim sReleaseField
    Dim OwnerID
    Dim OwnerName
    Dim strTxtNotify : strTxtNotify = request("txtnotify") 
    Dim strTxtNotifyFin 'Notify list for each product
    Dim sOperation
    DIM bAVRequired : bAVRequired = (request("chkAVRequired") = "on")
    DIM bQualificationRequired : bQualificationRequired = (request("chkQualificationRequired") = "on")
	dim strTO
	dim strSubject
	dim strBody
	dim strProgramName
	dim strID
	dim strCoreTeam
	dim strOwnerName
	dim strBusiness
    dim strODMCC
    dim strHPCC
	dim strCC
    DIM bImportant : bImportant = (request("chkImportant") = "on")
    dim oldConsumer
    dim oldSMB
    dim oldCommercial
    dim oldNA
    dim oldLA
    dim oldEMEA
    dim oldAPJ
    dim oldJustification
	strCC = ""

    strExcaliburHPLink = ""
    strExcaliburODMLink = ""

	dim strBusinessID
	strBusinessID = ""
	
	strChangeTypeCat = ""
	strCustomerImpact = "None"
	blnSendDisapprovedEmail = false
	
	CCTMailList = "HOUPRTSUSTAININGSYSTEMTEAM@hp.com"
	response.flush  
	strRestoreType = ""
	strOSList = ""
	strLanguageList = ""
	strDescription = ""	
	oldSummary = ""
    oldProduct = ""
 	oldstatus = ""
	oldOwner = ""
	oldZsrpTargetDt = ""
	oldZsrpActualDt = ""
	oldZsrpRequired = ""
	oldCoreTeam = ""
	oldTargetDate = ""
	oldNotify = ""
	oldDescription = ""
	oldAction = ""
	oldResolution = ""
	oldOnStatus = ""
	strProductStatusID = ""
	strDeliverableRootID = ""
	strDeliverableRootName = ""
	strDeliverableManagerID = ""
	strAddDCRNotificationList = ""
    oldConsumer=""
    oldSMB=""
    oldCommercial=""
    oldNA=""
    oldLA=""
    oldEMEA=""
    oldJustification=""
    Displayedid = 0

	strType=request("txtType")
	
	select case strType
	case "1"
		strTypeName = "Issue"
	case "2"
		strTypeName = "Action Item"
	case "3"
		strTypeName = "Change Request"
	case "4"
		strTypeName = "Status Note"
	case "5"
		strTypeName = "Improvement Opportunity"
	case "6"
		strTypeName = "Test Request"
	end select
	
	strConnect = Session("PDPIMS_ConnectionString")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = strConnect
	cn.CommandTimeout = 60
	cn.Open


	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


	CurrentUserID = 0
	'Get User
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

	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserName = rs("Name")
		CurrentUserEmail = rs("Email")
        CurrentUserPartner = rs("PartnerID") & ""
	end if
	rs.Close


	strDescription = request("txtDescription")
	if trim(request("txtType")) = "4" then
		strDescription = StripHTMLTag(strDescription)
	end if
	'if lcase(left(strDescription,26)) = "<font face=verdana size=2>" then
	'	strDescription = mid(strDescription,27)
	'end if
	'if lcase(right(strDescription,7)) = "</font>" then
	'	strDescription = left(strDescription,len(strDescription)-7)
	'end if


	dim strGroupsAdded
	dim strGroupsRemoved
	dim GroupAddArray
	dim GroupRemoveArray
	dim GroupSelectedArray
	dim GroupLoadedArray
	dim strDcrAutoOpen
	
	GroupSelectedArray = split(request("lstFunctionalGroup"),",")
	GroupLoadedArray = split(request("txtGroupsLoaded"),",")
	
	strGroupsAdded=""
	strGroupsRemoved = ""
	for i = lbound(GroupSelectedArray) to ubound(GroupSelectedArray)	
		if instr("," & request("txtGroupsLoaded") & ",", "," & trim(GroupSelectedArray(i)) & ",") = 0 then
			strGroupsAdded = strGroupsAdded & "," & trim(GroupSelectedArray(i))
		end if
	next

	for i = lbound(GroupLoadedArray) to ubound(GroupLoadedArray)	
		if instr(", " & request("lstFunctionalGroup") & ",", ", " & trim(GroupLoadedArray(i)) & ",") = 0 then
			strGroupsRemoved = strGroupsRemoved & "," & trim(GroupLoadedArray(i))
		end if
	next
	
	if strGroupsAdded <> "" then
		GroupAddArray = split(mid(strGroupsAdded,2),",")
	end if

	if strGroupsRemoved <> "" then
		GroupRemoveArray = split(mid(strGroupsRemoved,2),",")
	end if

    'Harris, Valerie -  02/9/2016 - PBI 15660/ Task 16234 - Set iIssueID variable
    If Request("txtID") <> "" Then
        iIssueID = Request("txtID")
    Else
        iIssueID = ""
    End If

	if bBios then
	    strProducts = "347,"
    elseif bID then
        strProducts = "1107,"
	elseif bScr then
	    strProducts = "344,"
        strDeliverableRootID = Request.Form("hidDeliverableRootId")
        strDeliverableRootName = Request.Form("txtDeliverableRootName")
    elseif request("chkPreinstallDeliverable") = "on" then
		strProducts = "170,"	
	else
        'Harris, Valerie -  02/4/2016 - PBI 15660/ Task 16234 - If Type 3 and Add mode, Get Products from checkbox table
        If strType = "3" And iIssueID = "" Then
            If Request("chkProducts") <> "" Then
                strProducts = Request("chkProducts") & ","
            Else
                strProducts = ","
            End If
        Else
            '---:Get Products from lstProducts drop-down: ---
		strProducts = request("lstProducts") & ","
        End If
	end if

    'Harris, Valerie -  02/4/2016 - PBI 15660/ Task 16234 - If Type 3 and Add mode, Get Change Request ID
    If strType = "3" And iIssueID = "" Then
        If Request("inpChangeRequestID") <> "" Then
            iChangeRequestID = CLng(Request("inpChangeRequestID"))
        Else
            iChangeRequestID = 0
        End If
    Else
        iChangeRequestID = 0
    End If

	dim NewStatus
	NewStatus = clng(request("cboStatus"))
	
    strDeliverableRootID = Request.Form("hidDeliverableRootId")
    strDeliverableRootName = Request.Form("txtDeliverableRootName")

	if request("txtID") = "" then 'Adding
		if strproducts <> "," then 'Just to make sure we don't get into an infinite loop
			
			if request("cboOwner") > 0 then
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetEmployeeByID"
		

				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = request("cboOwner")
				cm.Parameters.Append p
	

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

			
				if rs.EOF and rs.BOF then
					OwnerID = request("cboOwner")
					OwnerName =""
				else
					OwnerID = request("cboOwner")
					OwnerName = rs("Name")				
				end if
		
				rs.Close
			end if
			strRestoreType = ""
			
			Dim sBiosNotes
			sBiosNotes = ""
			
			If bBios Then
			    sBiosNotes = "BIOS CHANGE NOTES:" & vbcrlf & "This is a " & request("rbBiosNewChange") & vbcrlf
			    sBiosNotes = sBiosNotes & "Current Implementation: " & request("txtBiosCurrentImp") & vbcrlf
			    sBiosNotes = sBiosNotes & "Future Implementation: " & request ("txtBiosFutureImp") & vbcrlf
			    sBiosNotes = sBiosNotes & vbcrlf & "Customer Impact: " & request("txtCustomerImpact") & vbcrlf
			End If
			
			If bSCR Then
			    sBiosNotes = "SOFTWARE CHANGE NOTES:" & vbcrlf & "This is a " & request("rbBiosNewChange") & vbcrlf
			    sBiosNotes = sBiosNotes & "Current Implementation: " & request("txtBiosCurrentImp") & vbcrlf
			    sBiosNotes = sBiosNotes & "Future Implementation: " & request ("txtBiosFutureImp") & vbcrlf
			    sBiosNotes = sBiosNotes & vbcrlf & "Customer Impact: " & request("txtCustomerImpact") & vbcrlf
			End If
			
			if strType = "3" then		
			
				'Restore
				if request("chkFull") = "on" then
					strRestoreType = strRestoreType & ", Full"
				end if
				if request("chkSelect")= "on" then
					strRestoreType = strRestoreType & ", Select"
				end if
				if request("chkDIB")= "on" then
					strRestoreType = strRestoreType & ", DIB"
				end if
				if len(strRestoreType) > 0 then
					strRestoreType =  vbcrlf & "RESTORE SOLUTION: " & mid(strRestoreType,3)
				end if


				'OS
				if request("lstOS") <> "" then
					strOSList =  "OPERATING SYSTEMS: " & request("lstOS")
				end if

				'Language
				strLanguageList = request("lstLanguages")
				if len(strLanguageList) > 0 then
					strLanguageList = "LANGUAGES: " & strLanguageList
				end if
			end if

			'cn.BeginTrans
			'loopcount = 0
			do while instr(strproducts,",")' and loopcount < request("txtProductCount") + 1 'Just to make sure we don't get into an infinite loop
				'loopcount = loopcount  + 1
				strChangeTypeCat=""
				strBusiness = ""
				strRegions = ""
    			
				strProductID = left(strproducts,instr(strproducts,",")-1)
				strproducts = mid(strproducts,instr(strproducts,",")+1)

                sReleaseField = "chkRelease_"&Trim(strProductID)&""

                If CInt(strType) = 3 Then
                    If Request(""&sReleaseField&"") <> "" Then
                        sProductRelease = Request(""&sReleaseField&"") 
                    Else
                        sProductRelease = ""
                    End If
                End If

				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetProductVersionName"
		

				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = strproductID
				cm.Parameters.Append p
	

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				'rs.Open "spGetProductVersionName " & strproductID,cn,adOpenForwardOnly	
				strProductStatusID = rs("ProductStatusID") & ""
				strProductName = rs("Name") & ""
				strProgramMail =  rs("Email") & ""
				strDcrAutoOpen = rs("DCRAutoOpen") 
				strAddDCRNotificationList = rs("AddDCRNotificationList")
				strBusinessID = rs("BusinessID")
				sOperation = rs("Operation") & ""
				'TODO:Check with Walter
				' Add DCR Notification
				'
				if trim(rs("Division")& "") = "1" and strDcrAutoOpen >= 2 and cint(strType) = 3 AND NOT (bBios OR bSCR OR bID) and LCase(strAddDCRNotificationList) = "true" then
					if trim(rs("Distribution") & "") <> "" then
					    if trim(sOperation) = "0" then
					        strProductEmail = rs("Distribution") & ";NotebookDCRNotification@hp.com"
                        else
                            strProductEmail = rs("Distribution") & ""
                        end if
					else
						strProductEmail = "NotebookDCRNotification@hp.com"
					end if
				elseif strDcrAutoOpen >= 2 and cint(strType) = 3 then
					strProductEmail= rs("Distribution") & ""
				else
					strProductEmail= ""
				end if				
				
				
				rs.Close		
				if request("cboOwner") > 0 then
					OwnerID = request("cboOwner")
				else
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetDefaultDCROwner"

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = strproductID
					cm.Parameters.Append p
	
					rs.CursorType = adOpenForwardOnly
					rs.LockType=AdLockReadOnly
					Set rs = cm.Execute 
					Set cm=nothing

					if rs.EOF and rs.BOF then
						OwnerID = ""
						OwnerName = ""
					else
						OwnerID = rs("ID")
						OwnerName = rs("Name")				
					end if
		
					rs.Close
				end if

                if trim(OwnerID) = "" then

					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetPM"

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = strproductID
					cm.Parameters.Append p
	
					rs.CursorType = adOpenForwardOnly
					rs.LockType=AdLockReadOnly
					Set rs = cm.Execute 
					Set cm=nothing

					if rs.EOF and rs.BOF then
						OwnerID = ""
						OwnerName = ""
					else
						OwnerID = rs("ID")
						OwnerName = rs("Name")				
					end if
		
					rs.Close

                end if

                '''get Program Coordinator's Email, PCEmail 
        		set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 1
				cm.CommandText = "SELECT  ISNULL(ui.Email, '') AS PCEmail FROM ProductVersion AS v INNER JOIN BusinessSegment bs ON v.BusinessSegmentID = bs.BusinessSegmentID LEFT OUTER JOIN UserInfo ui ON v.PCID = ui.UserID WHERE v.TypeID IN (1, 3) AND v.ID = " + cstr(cint(strProductID)) + " AND bs.BusinessID=1 and bs.Operation =1 "

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				if rs.EOF and rs.BOF then
					strPCEmail = ""
				else
					strPCEmail = rs("PCEmail")	
				end if
		
				rs.Close
                '''get Program Coordinator's Email, PCEmail 

                '''Add PCEmail into Notify email list, strTxtNotify
                if len(strPCEmail)>5 and instr(strTxtNotify,strPCEmail)<1 then
                    strTxtNotifyFin = strPCEmail + ";" + strTxtNotify
                else
                    strTxtNotifyFin = strTxtNotify
                end if
			
				set cm = server.CreateObject("ADODB.command")
		
				cm.ActiveConnection = cn
				cm.CommandText = "spAddDeliverableActionWeb2"
				cm.CommandType =  &H0004
				cm.NamedParameters = True

				set p =  cm.CreateParameter("@ProductID", 3, &H0001)
				p.value = strproductid
				cm.Parameters.Append p
	            
				set p =  cm.CreateParameter("@DeliverableRootId", 3, &H0001)
				If Trim(strDeliverableRootId) = "" Then
				    p.value = 0
				Else
				    p.value = CLng(strDeliverableRootId)
				End If
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Type", 16, &H0001)
				p.Value = clng(strType)
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Status", 16, &H0001)
				if strDcrAutoOpen >= 2 and cint(strType) = 3 then
					p.Value = 6
				else
					p.Value = 1
				end if
				cm.Parameters.Append p
		
				Set p = cm.CreateParameter("@Submitter", 200, &H0001, 50)
				p.Value = left(currentusername,50)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SubmitterID", 3, &H0001)
				p.Value = clng(currentuserid)
				cm.Parameters.Append p
	
				set p =  cm.CreateParameter("@CategoryID", 3, &H0001)
				p.value = 1'request("cboCategory")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OwnerID", 3, &H0001)
				p.Value = OwnerID
				cm.Parameters.Append p
            	strNewOwner = OwnerID

				Set p = cm.CreateParameter("@PreinstallOwner", 3, &H0001)
				p.Value = clng(request("cboPreinstallApprover"))
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@AffectsCustomers", 16, &H0001)
		        if strType = "5" then
					p.value=request("cboNetAffect")
		        else
					if request("chkCustomers") = "on" then
						p.Value = 1
						strCustomerImpact = "Affects Images And/Or BIOS on shipping products."
					else
						p.Value = 0
					end if
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@CoreTeamID", 3, &H0001)
				p.Value = 2
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@RoadmapID",adInteger, &H0001)
				p.Value = 0
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@TargetDate", 135, &H0001)
				if request("txtTarget")="" then
					p.Value = null
		        else
					p.Value = CDate(request("txtTarget"))
				end if
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@TestDate", 135, &H0001)
				if trim(request("txtAvailDate")) <> "" and isdate(trim(request("txtAvailDate"))) then
					p.Value = CDate(trim(request("txtAvailDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@TestNote", 200, &H0001,35)
		        if strType = "5" then
					p.Value = left(request("cboMetricImpact"),35)
				else
					p.Value = left(request("txtAvailNotes"),35)
				end if
				cm.Parameters.Append p

				
	            
				Set p = cm.CreateParameter("@BTODate", 135, &H0001)
				if request("txtBTODate") <> "" then
					p.Value = CDate(request("txtBTODate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@CTODate", 135, &H0001)
				if request("txtCTODate") <> "" then
					p.Value = CDate(request("txtCTODate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

    			Set p = cm.CreateParameter("@Notify", 200, &H0001, 8000)
				p.Value = left(strTxtNotifyFin,8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OnStatus", 16, &H0001)
				'if lcase(trim(request("chkReports"))) = "on" then
				'	p.Value = 1
				'else
					p.Value = 0
				'end if
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@Priority", 16, &H0001)
				if request("lstPriority") = "" then
					p.Value = 0
				else
					p.Value = clng(request("lstPriority"))
				end if
				cm.Parameters.Append p
	
		        Set p = cm.CreateParameter("@Commercial", 11, &H0001)
				if request("chkCommercial") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", Commercial"
				else
					p.Value = false
				end if
				cm.Parameters.Append p


		        Set p = cm.CreateParameter("@Consumer", 11, &H0001)
				if request("chkConsumer") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", Consumer"
				else
					p.Value = false
				end if
				cm.Parameters.Append p


		        Set p = cm.CreateParameter("@SMB", 11, &H0001)
				if request("chkSMB") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", SMB"
				else
					p.Value = false
				end if
				cm.Parameters.Append p
        		        
                Set p = cm.CreateParameter("@NA", 11, &H0001)
				if request("chkNA") = "on" then
					p.Value = True
					strRegions = strRegions & ", NA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@LA", 11, &H0001)
				if request("chkLA") = "on" then
					p.Value = True
					strRegions = strRegions & ", LA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@APJ", 11, &H0001)
				if request("chkAPJ") = "on" then
					p.Value = True
					strRegions = strRegions & ", APJ"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@EMEA" ,11, &H0001)
				if request("chkEMEA") = "on" then
					p.Value = True
					strRegions = strRegions & ", EMEA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@AddChange", 11, &H0001)
				if request("chkAdd") = "on" then
					p.Value = True
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ModifyChange", 11, &H0001)
				if request("chkModify") = "on" then
					p.Value = True
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@RemoveChange", 11, &H0001)
				if request("chkRemove") = "on" then
					p.Value = True
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ImageChange", 11, &H0001)
				p.Value = bSoftware
				if bSoftware then
					strChangeType = strChangeType & ",Image"
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@CategoryBiosChange", 11, &H0001)
				p.Value = bCategoryBiosChange
				if bCategoryBiosChange then
					strChangeType = strChangeType & ",Bios"
				end if
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@CommodityChange", 11, &H0001)
				p.Value = bCommodity
				if bCommodity then
					strChangeType = strChangeType & ",Commodity"
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@DocChange", 11, &H0001)
				p.Value = bDocs
				if bDocs then
					strChangeType = strChangeType & ",Doc"
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@SKUChange", 11, &H0001)
    			p.Value = bSku
				if bSku then
					strChangeTypeCat = strChangeTypeCat & ",SKU"
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ReqChange", 11, &H0001)
                p.Value = bRequirement
                if bRequirement Then
					strChangeTypeCat = strChangeTypeCat & ",Requirement"
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@BiosChange", 11, &H0001)
				p.Value = bBios
				if bBios Then
				    strChangeTypeCat = strChangeTypeCat & ",Bios"
				end if
				
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@SwChange", 11, &H0001)
				p.Value = bSCR
				if bSCR Then
				    strChangeTypeCat = strChangeTypeCat & ",SW"
				end if
				
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@IDChange", 11, &H0001)
				p.Value = bID
				if bID Then
				    strChangeTypeCat = strChangeTypeCat & ",ID"
				end if
				
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@OtherChange", 11, &H0001)
				p.Value = bOther
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@PendingImplementation", 11, &H0001)
				if request("chkClosed") = "on" then
					p.Value = true
				else
					p.Value = false
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Summary", 200, &H0001, 120)
				p.Value = left(request("txtSummary"),120)
				cm.Parameters.Append p
            
				Set p = cm.CreateParameter("@Description", 200, &H0001, 8000)
				p.Value = left(strDescription,8000)
				cm.Parameters.Append p
           
				Set p = cm.CreateParameter("@Justification", 200, &H0001, 8000)
                if CurrentUserPartner = "1" then
				    p.Value = left(request("txtJustification"),8000)
                else
                    p.Value = null
                end if
				cm.Parameters.Append p


				Set p = cm.CreateParameter("@Details", 200, &H0001, 8000)
				if bID then
                    p.Value= ""
                elseif bBios Or bSCR then
    				p.Value = left(sBiosNotes & vbcrlf & strOSList & vbCrLf & strLanguageList,8000)
				else
	    			p.Value = left(strOSList & vbCrLf & strLanguageList,8000)
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@LastUpdUser", 200, &H0001, 200)
				p.Value = CurrentUserName
				cm.Parameters.Append p
	    
				Set p = cm.CreateParameter("@Actions", 201, &H0001, 2147483647)
				p.Value = request("txtActions")
				cm.Parameters.Append p
	    
		        Set p = cm.CreateParameter("@Resolution", 201, &H0001, 2147483647)
				'if request("txtFormType") = "Change" then
		        '    p.Value = request("txtActionItems")
				'else
		            p.Value = ""
				'end if
				cm.Parameters.Append p
	    
				Set p = cm.CreateParameter("@ZsrpReadyTargetDt", 135, &H0001)
				if request("txtZsrpReadyTargetDt") <> "" then
					p.Value = CDate(request("txtZsrpReadyTargetDt"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

    			Set p = cm.CreateParameter("@ZsrpReadyActualDt", 135, &H0001)
				if request("txtZsrpReadyActualDt") <> "" then
					p.Value = CDate(request("txtZsrpReadyActualDt"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ZsrpRequired", 11, &H0001)
				p.Value = bZsrpReadyRequired
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@RTPDate", 135, &H0001)
				if trim(request("txtRTPDate")) <> "" and isdate(trim(request("txtRTPDate"))) then
					p.Value = CDate(trim(request("txtRTPDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@RASDiscoDate", 135, &H0001)
				if trim(request("txtRASDiscoDate")) <> "" and isdate(trim(request("txtRASDiscoDate"))) then
					p.Value = CDate(trim(request("txtRASDiscoDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p
				
                'Harris, Valerie -  02/8/2016 - PBI 15660/ Task 16234 - If Type 3, Add Release parameter to capture Product's selected Releases
                Set p =  cm.CreateParameter("@ProductVersionRelease", 200, &H0001, 8000)
                If CInt(strType) = 3 Then
                    p.value = sProductRelease 
                Else
                    p.value = ""
                End If
				cm.Parameters.Append p

                'Harris, Valerie -  02/8/2016 - PBI 15660/ Task 16234 - If Type 3, use Change Request ID parameter so products submitted on one change request are displayed toghether in edit mode
                Set p = cm.CreateParameter("@ChangeRequestID", 	20, &H0001)
				If CInt(strType) = 3 Then
                    If iChangeRequestID <> 0 Then
                        p.value = iChangeRequestID 
                    Else
                        p.value = null
                    End If
                Else
                    p.value = null
                End If
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@AVRequired", 11, &H0001)
				p.Value = bAVRequired
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@QualificationRequired", 11, &H0001)
				p.Value = bQualificationRequired
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@TargetApprovalDate", 135, &H0001)
				if trim(request("txtTargetApprovalDate")) <> "" and isdate(trim(request("txtTargetApprovalDate"))) then
					p.Value = CDate(trim(request("txtTargetApprovalDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@Important", 11, &H0001)
				p.Value = bImportant
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@NewID", 3,  &H0002)
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@Attachment1",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath1"),500)
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Attachment2",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath2"),500)
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Attachment3",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath3"),500)
	            cm.Parameters.Append p
	    
                Set p = cm.CreateParameter("@Attachment4",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath4"),500)
	            cm.Parameters.Append p

                Set p = cm.CreateParameter("@Attachment5",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath5"),500)
	            cm.Parameters.Append p

                set p =  cm.CreateParameter("@DeliverableIssueID", 3, &H0001)
				p.value = Displayedid
				cm.Parameters.Append p

				cm.Execute RowsEffected
				
				if cn.Errors.Count > 1 then
					strproducts = ""
					Errors = true
				end if
				strIDOutput = strIDOutput & "<TR bgcolor=cornsilk><TD><font size=2 face=verdana color=black>" & strproductname & "</font></TD><TD><font size=2 face=verdana color=black>" & cm("@NewID") & "</FONT></TD><TD><font size=2 face=verdana color=black>" & OwnerName & "</FONT></TD></TR>"
				Displayedid = cm("@NewID")	
				set cm = nothing				
                
				strApproverIDs = ""
				strApproverEmails = ""
				strApprovers = ""               
				CommodityPMEmail = ""
                strPCEmail = ""				
				
				If (not Errors) and clng(strDcrAutoOpen) >= 2 and cint(strType) = 3 then
					If clng(strDcrAutoOpen) = 2 Then ''''System Team
	    				rs.open "usp_ListSystemTeamApprovers " & strproductid,cn,adOpenForwardOnly
					
					    do while not rs.eof
                           ' TAKE ALL ROLES THAT ARE PRIMARY TEAM MEMBERS and NOT LISTED YET
						    If (trim(rs("PrimaryTeam") & "") = "1") and InStr(strApproverEmails, rs("Email")) = 0 then								
                                strApproverIDs = strApproverIDs & "," & rs("ID")
								strApproverEmails = strApproverEmails & rs("Email") & ";"
								strApprovers = strApprovers & rs("Name") & " - Requested<BR>"                                      
						    end if
                            if trim(rs("Role") & "") = "Commodity PM" then
							    CommodityPMEmail = rs("Email")
						    end if
						    if trim(rs("Role") & "") = "Program Coordinator" then
						        strPCEmail = rs("Email")
						    end if
						    rs.movenext	
					    loop
					    rs.close
                    ElseIf clng(strDcrAutoOpen) = 4 Then ''''System Team but no ODM
	    				rs.open "usp_ListSystemTeamApprovers " & strproductid,cn,adOpenForwardOnly
					    do while not rs.eof
						    If (trim(rs("PrimaryTeam") & "") = "1") and (InStr(strApproverEmails, rs("Email")) = 0) and (InStr(UCase(rs("Role")), "ODM") = 0) then								
                                strApproverIDs = strApproverIDs & "," & rs("ID")
								strApproverEmails = strApproverEmails & rs("Email") & ";"
								strApprovers = strApprovers & rs("Name") & " - Requested<BR>"                                      
						    end if
                            if trim(rs("Role") & "") = "Commodity PM" then
							    CommodityPMEmail = rs("Email")
						    end if
						    if trim(rs("Role") & "") = "Program Coordinator" then
						        strPCEmail = rs("Email")
						    end if
						    rs.movenext	
					    loop
					    rs.close
					ElseIf clng(strDcrAutoOpen) = 3 Then ''''manual list from "Edit Product"
					    rs.open "spListProductDcrApprovers " & strProductId, cn, adOpenForwardOnly
					    
					    Do While Not rs.eof                            
                            strApproverIDs = strApproverIDs & "," & rs("ID")
                            strApproverEmails = strApproverEmails & rs("Email") & ";"
                            strApprovers = strApprovers & rs("Name") & " - Requested<BR>"                             
                            rs.movenext
					    loop
					    rs.close
					    
					    if bScr And trim(strDeliverableRootID) <> "0" Then
					        rs.open "usp_SelectDeliverableRootScrApprovers " & strDeliverableRootID, cn, adOpenStatic
					    
					        Do While Not rs.Eof
                               strApproverIDs = strApproverIDs & "," & rs("ID")
                                strApproverEmails = strApproverEmails & rs("Email") & ";"
                                strApprovers = strApprovers & rs("Name") & " - Requested<BR>" 
                               rs.MoveNext
					        Loop
					        rs.Close
					    end if
					end if
	
					if len(strApproverIDs) > 0 then
						strApproverIDs = mid(strApproverIDs,2)
					end if
        
					if request("chkDocChange") = "on" then
						if len(strApproverEmails) > 0 then
							strApproverEmails = strApproverEmails & "houdcrdocs@hp.com;"
						else
							strApproverEmails = "houdcrdocs@hp.com;" 
						end if						
					end if					
					ApproverArray = split(strApproverIDs,",")
					for each Approver in ApproverArray
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spAddApprover"
		

						Set p = cm.CreateParameter("@ID", 3, &H0001)
						p.Value = DisplayedID
						cm.Parameters.Append p
	
						Set p = cm.CreateParameter("@ApproverID", 3, &H0001)
						p.Value = clng(Approver)
						cm.Parameters.Append p

						cm.Execute RowsEffected
						Set cm=nothing					
					next	
				end if
				
								
				if strGroupsAdded <> "" then
					for i = lbound(GroupAddArray) to ubound(GroupAddArray)
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spAddActionGroup"
		
						Set p = cm.CreateParameter("@ID", 3, &H0001)
						p.Value = clng( Displayedid)
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@GroupID", 3, &H0001)
						p.Value = clng( GroupAddArray(i))
						cm.Parameters.Append p
	
						cm.Execute rowseffected
						Set cm=nothing
							
						if rowseffected <> 1 then
							Errors= true
							exit for
						end if
					next
				end if				
				
				
				
				strGroupList=""
				if request("lstFunctionalGroup") <> "" then
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
				end if
								
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetEmployeeByID"
		

				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = OwnerID
				cm.Parameters.Append p
	

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				strPMEmail = rs("Email") & ""		
				rs.Close
				

				if trim(strProductEmail) = "" then
					strProductEmail = strPMEmail
				end if

				if strPMEmail <> "" then
					select case strtype
					case "1"
						strTypeName = "Issue"
					case "2"
						strTypeName = "Action Item"
					case "3"
						strTypeName = "Change Request"
					case "4"
						strTypeName = "Status Note"
					case "5"
						strTypeName = "Improvement Opportunity"
					case "6"
						strTypeName = "Test Request"
					case else
						strTypeName = "Item"
					end select
					
					if strChangeTypeCat <> "" then
						strChangeTypeCat = " (" & mid(strChangeTypeCat,2) & ")"
					end if
			        
                    strExcaliburHPLink = ""
                    strExcaliburODMLink = ""
					strExcaliburHPLink = "<font face=Arial size=2><a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & strtype & "&id=" & DisplayedID & """>Open this " & strtypename & "</a><br>"
					strExcaliburHPLink = strExcaliburHPLink & "<a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/Excalibur.asp"">Open Pulsar Today Page</a></font><br><br>"
                    strExcaliburODMLink =strExcaliburHPLink

           If bImportant Then
					    strBody = "<font face=Arial size=4 color=red><b>This Change Request has been identified as IMPORTANT!. Please handle quickly.</b></font><br><br><font face=Arial size=2>"
            else
              strBody = "<font face=Arial size=2>"
           End If	

					
					strBody = strBody & "<b>NUMBER:</B> " & DisplayedID & "<BR>"
					strBody = strBody & "<b>TYPE:</b> " & strtypename & "<BR>"
					strBody = strBody & "<b>SUBMITTER:</b> " & currentusername & "<BR>"
					strBody = strBody & "<b>PRODUCT:</b> " & strproductname & "<BR>"
                    
                    'Harris, Valerie -  02/12/2016 - BUG 15660/ Task 16234 - If Type 3, Add Release section email to capture Product's selected Releases for DCR
                    If CInt(strType) = 3 And sProductRelease <> "" Then
                        strBody = strBody & "<b>RELEASE:</b> " & sProductRelease & "<BR>"
                    End If
					
					if strtype = "5" then
						strBody = strBody & "<b>ISSUE/ACCOMPLISHMENT:</B> " & replace(server.HTMLEncode(request("txtSummary")),"""","&QUOT;") & "<BR>"
					else
						strBody = strBody & "<b>SUMMARY:</B> " & replace(server.HTMLEncode(request("txtSummary")),"""","&QUOT;") & "<BR>"
					end if
					if strDcrAutoOpen >= 2 and cint(strType) = 3 then
						strBody = strBody & "<b>STATUS:</b> Investigating<BR>"
					else
						strBody = strBody & "<b>STATUS:</b> Open<BR>"
					end if
					strBody = strBody & "<b>OWNER:</b> " & OwnerName & "<BR>"
			        if strBusiness <> "" then
				        strBody = strBody & "<b>BUSINESS:</b> " & Mid(strBusiness, 3) & "<BR>"
			        end if
			        if strRegions <> "" Then
			            strBody = strBody & "<b>REGIONS:</b> " & Mid(strRegions, 3) & "<br>"
			        end if

					if request("txtDescription") <> "" then
       					if trim(strType) = "5" then
       						strDescription = replace(server.HTMLEncode(request("txtDescription")) & "",vbcrlf,"<BR>")
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
							strBody = strBody & strDescription & "<BR>" & "<BR>"
						else
							strBody = strBody & "<b>DESCRIPTION:</b> " & "<BR>" & replace(server.HTMLEncode(request("txtDescription")) & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						end if
					end if

                    If trim(strType) = "3" Then
                        if bBios Then
                            sBiosNotes = replace(sBiosNotes, vbcrlf, "<BR>")
                            strBody = strBody & "<b>DETAILS:</b> " & "<BR>" & sBiosNotes & "<BR>" & strOSList & "<BR>" & strLanguageList & "<BR>" & "<BR>"
                        Elseif not bID then
                            strBody = strBody & "<b>DETAILS:</b> " & "<BR>" & strOSList & "<BR>" & strLanguageList & "<BR>" & "<BR>"
                        End If
                    End If
        
                    'use the same email body for the ODM user upto this point
                    strODMBody = strBody            

                    if CurrentUserPartner = "1" then
                        if request("txtJustification") <> "" then
						    if trim(strtype) = "5" then
							    strBody = strBody & "<b>ROOT CAUSE:</b> " & "<BR>" & replace(server.HTMLEncode(request("txtJustification")) & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						    else
							    strBody = strBody & "<b>JUSTIFICATION:</b> " & "<BR>" & replace(server.HTMLEncode(request("txtJustification")) & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						    end if
					    end if
                    end if
					'do not add justification to the ODM's. add everything else
					If strtypename <> "Status Note" and request("txtActions") <> ""  Then
       					if trim(strType) = "5" then
       						strDescription = replace(request("txtActions"),vbcrlf,"<BR>")
							StringArray = split(strDescription,chr(1))
							if ubound(StringArray) > -1 then
								if trim(StringArray(0)) <> "" then
									strDescription = "<b>CORRECTIVE ACTIONS:</b><br>" & StringArray(0)
								else
									strDescription = ""				
								end if
							end if
							if ubound(StringArray) > 0 then
								if trim(StringArray(0)) <> "" and trim(StringArray(1)) <> ""  then
									strDescription = strDescription & "<BR><BR>"
								end if
								if trim(StringArray(1)) <> "" then
									strDescription = strDescription & "<b>PREVENTIVE ACTIONS:</b><br>" & StringArray(1)
								end if
							end if
							strBody = strBody & strDescription & "<BR>" & "<BR>"
                            strODMBody = strODMBody & strDescription & "<BR>" & "<BR>"
						else
							strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & replace(server.HTMLEncode(request("txtActions")),vbcrlf,"<BR>") & "</font>" & "<BR>" & "<BR>"
                            strODMBody = strODMBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & replace(server.HTMLEncode(request("txtActions")),vbcrlf,"<BR>") & "</font>" & "<BR>" & "<BR>"
						end if
					End If
					
       		        if trim(strType) = "5" then
						select case trim(request("lstPriority"))
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
                        strODMBody = strODMBody & "<b>IMPACT:</b> " & strPriority & "<BR>"
					end if 
					
       		        if trim(strType) = "5" then
						strBody = strBody & "<b>METRIC IMPACTED:</b> " & left(request("cboMetricImpact"),35)	 & "<BR>"
                        strODMBody = strODMBody & "<b>METRIC IMPACTED:</b> " & left(request("cboMetricImpact"),35)	 & "<BR>"
					end if
					
       		        if trim(strType) = "5" then
						if request("cboNetAffect")=1  then
							strCustomers = "Positive"
						elseif request("cboNetAffect")=0 then
							strCustomers = "&nbsp;"
						else
							strCustomers = "Negative"
						end if
						strBody = strBody & "<b>NET AFFECT: </b> " & strCustomers & "<BR>"
                        strODMBody = strODMBody & "<b>NET AFFECT: </b> " & strCustomers & "<BR>"
					end if        
					
					If bZsrpReadyRequired Then
					    strBody = strBody & "<b>ZSRP READY TARGET: </b> " & request("txtZsrpReadyTargetDt") & "<br>"
                        strODMBody = strODMBody & "<b>ZSRP READY TARGET: </b> " & request("txtZsrpReadyTargetDt") & "<br>"
					    strBody = strBody & "<b>ZSRP READY ACTUAL: </b> " & request("txtZsrpReadyActualDt") & "<br><br><br>"
                        strODMBody = strODMBody & "<b>ZSRP READY ACTUAL: </b> " & request("txtZsrpReadyActualDt") & "<br><br><br>"
                    End If					    
					
                    If sBios = "" and not bID Then
                        strBody = strBody & "<b>CUSTOMER IMPACT: </b>" & strCustomerImpact & "<BR>"
                        strODMBody = strODMBody & "<b>CUSTOMER IMPACT: </b>" & strCustomerImpact & "<BR>"
                    End If
        			if trim(request("txtAvailDate")) <> "" then
                        if bID then
		        		    strBody = strBody & "<font color=red><b>DEADLINE: </b> " & request("txtAvailDate") & "</font><BR>"
                            strODMBody = strODMBody & "<font color=red><b>DEADLINE: </b> " & request("txtAvailDate") & "</font><BR>"
                        else
		        		    strBody = strBody & "<font color=red><b>SAMPLES AVAILABLE: </b> " & request("txtAvailDate") & "</font><BR>"
                            strODMBody = strODMBody & "<font color=red><b>SAMPLES AVAILABLE: </b> " & request("txtAvailDate") & "</font><BR>"
                        end if
			        end if

                    if cint(strType) = 3 then
                        if trim(request("txtTargetApprovalDate")) <> "" then
                            strBody = strBody & "<font color=red><b>Target Approval Date:</b> " & request("txtTargetApprovalDate") & "</font><BR>"
                            strODMBody = strODMBody & "<font color=red><b>Target Approval Date:</b> " & request("txtTargetApprovalDate") & "</font><BR>"
                        end if
                    end if

					strBody = strBody & "<b>NOTIFY ON CLOSURE: </b> " & strTxtNotifyFin & "<BR>"
                    strODMBody = strODMBody & "<b>NOTIFY ON CLOSURE: </b> " & strTxtNotifyFin & "<BR>"
					
					if strGroupList <> "" then
						strBody = strBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
                        strODMBody = strODMBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
					end if
					
					if strApprovers <> "" then
						strBody = strBody & "<b>APPROVALS:</b><font color=teal><BR>" & strApprovers & "</font><BR><BR>"
                        strODMBody = strODMBody & "<b>APPROVALS:</b><font color=teal><BR>" & strApprovers & "</font><BR><BR>"
					end if
					
					if clng(strDcrAutoOpen) >= 2 and cint(strType) = 3 then
						if strApprovers <> "" then
						    If strDcrAutoOpen = 2 Then
							    strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The System Team members have been added as approvers for this DCR.</font></b><BR><BR>" & strExcaliburHPLink & strBody
							    strODMBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The System Team members have been added as approvers for this DCR.</font></b><BR><BR>" & strExcaliburODMLink & strODMBody
							Else 
							    strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The specified people have been added as approvers for this DCR.</font></b><BR><BR>" & strExcaliburHPLink & strBody
							    strODMBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The specified people have been added as approvers for this DCR.</font></b><BR><BR>" & strExcaliburODMLink & strODMBody
							End If
						else
							strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".</font></b><BR><BR>" & strBody
						    strODMBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".</font></b><BR><BR>" & strODMBody
						end if

					else
						strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added and assigned to " & OwnerName & ".</font></b><BR><BR>" & strBody
                        strODMBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added and assigned to " & OwnerName & ".</font></b><BR><BR>" & strODMBody
					end if
					strBody = strBody & "<br><font size=1 color=red face=verdana>HP Restricted</font>"
					
                    dim strAllEmails
                    strAllEmails = ""
                    if strApproverEmails <> "" then
						strAllEmails=strProductEmail & ";" & strApproverEmails
					else
						strAllEmails=strProductEmail
					end if
        
                    'separate the hp and odm email addresses
                    strHPEmailAddress = ""
                    strODMEmailAddress = ""
                    EmailArray = getUniqueItems(split(strAllEmails,";"))
				    for each emailaddress in EmailArray
                        if Len(Trim(emailaddress)) > 0 then                    
					        if instr(UCase(emailaddress), "@HP.COM") = 0 then
                                strODMEmailAddress = strODMEmailAddress & ";" & emailaddress
                            else
                                strHPEmailAddress = strHPEmailAddress & ";" & emailaddress
                            end if		
                        end if
				    next

                    if len(strHPEmailAddress) > 0 then
						strHPEmailAddress = mid(strHPEmailAddress,2)
					end if

                   'email to HP users
					Set oMessage = New EmailQueue 
					oMessage.From = CurrentUserEmail
					if strProgramMail = "1"  then
						oMessage.To= strHPEmailAddress						
						oMessage.Subject = strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
					else
						oMessage.To= CurrentUserEmail 
						oMessage.Subject = "TEST MAIL: " & strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
					end if
				
					if trim(strProgramMail) <> "1"   then
                        strHPBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strHPEmailAddress & "</font><BR><BR>" & strBody
					else
                        strHPBody = strBody
					end if
					
				    oMessage.HtmlBody = strHPBody

					'If (strProductStatusID = "3" or strProductStatusID = "4") And strBusinessID = 1 Then				
					'    oMessage.Cc = "houportpreinpm@hp.com"
					'End If
					
					oMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
					Set oMessage = Nothing 
                    
                    if len(strODMEmailAddress) > 0 then
						strODMEmailAddress = mid(strODMEmailAddress,2)
					end if
                    'email to ODM users
				    Set oODMMessage = New EmailQueue					

					oODMMessage.From = CurrentUserEmail
					if strProgramMail = "1"  then
						if strODMEmailAddress <> "" then
							oODMMessage.To= strODMEmailAddress				
						    oODMMessage.Subject = strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
                            oODMMessage.HtmlBody = strODMBody
                            oODMMessage.SendWithOutCopy  'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
                        end if				
					else
                        if strODMEmailAddress <> "" then
						    oODMMessage.To= CurrentUserEmail 
						    oODMMessage.Subject = "TEST MAIL: " & strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
                            oODMMessage.HtmlBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strODMEmailAddress & "</font><BR><BR>" & strODMBody					
                            oODMMessage.SendWithOutCopy   'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.				
                        end if
					end if				
										
					Set oODMMessage = Nothing
            	end if				
				
			loop
			
			'if Errors then
			'	cn.RollbackTrans
			'else
			'	cn.CommitTrans
			'end if
		end if
		cn.Close
		set cn = nothing
		
		if errors then
			'if request("txtFormType") = "Change" then
				response.write "<label style=""Display:none"" ID=Results><b><font face=verdana>Unable to submit your " & strTypeName & " at this time.</font></b><BR><BR>An Unexpected Error Occurred.<BR><BR></label>"
			'else
			'	reponse.write "Unable to submit you Issue/Risk."
			'end if
		else
				Response.Write "<label style=""Display:none"" ID=Results><h3><font face=verdana>" & strTypeName & " submitted.  Thank you for your input.</font></h3>"
				'response.write "<a href=""changeform.asp"" >Add Another Change Request</A><BR>"
				'Response.Write "<a href=""default.asp"" >Done Adding Change Requests</A><BR><BR>"
				Response.Write "<font face=verdana size=2><a href=""javascript:window.print();"" >Print This Window</A> | "
				Response.Write "<a href=""javascript:window.returnValue = 1;window.parent.close();"" >Close This Window</A><BR><BR></font>"
				Response.Write "<font size=2 face=verdana><b>" & request("txtSummary") & "</b></br></font><br>"
				Response.Write "<TABLE borderColor=tan cellSpacing=1 cellPadding=1 width=400 bgColor=wheat border=1><TR><TD width=270><FONT face=verdana size=2 color=black><b>Product</b></FONT></TD><TD width=200><FONT face=verdana size=2 color=black><b>ID Number</b></FONT></TD><TD width=270><FONT face=verdana size=2 color=black><b>Owner Assigned</b></FONT></TD></TR>"
				Response.Write strIDOutput & "</TABLE></label>"

		end if
	else 'Editing
			'Pull old record
			strProductID = left(strproducts,instr(strproducts,",")-1)
			strproducts = mid(strproducts,instr(strproducts,",")+1)

			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetActionProperties"
			

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("txtID")
			cm.Parameters.Append p
	

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			if not(rs.EOF and rs.BOF) then
				oldSummary = rs("Summary") & ""
				oldProduct = rs("ProductVersionID") & ""
				oldStatus = rs("status") & ""
				oldOwner = rs("OwnerID") & ""
				oldCoreTeam = rs("CoreTeamRep") & ""
				oldTargetDate = rs("TargetDate") & ""
				oldNotify = rs("Notify") & ""
				oldDescription = rs("Description") & ""
				oldAction = rs("Actions") & ""
				oldResolution = rs("Resolution") & ""
				oldOnStatus = rs("OnStatusReport") & ""
				blnReleaseNotification = rs("ReleaseNotification") & ""
				oldZsrpTargetDt = rs("ZsrpReadyTargetDt") & ""
				oldZsrpActualDt = rs("ZsrpReadyActualDt") & ""
				oldZsrpRequired = rs("ZsrpRequired") & ""
                productVersionRelease= rs("ProductVersionRelease") & ""
                productName = rs("dotsname") & ""
                oldConsumer = rs("Consumer") & ""
                oldSMB = rs("SMB") & ""
                oldCommercial=rs("Commercial") & ""
                oldNA=rs("NA") & ""
                oldLA=rs("LA") & ""
                oldEMEA=rs("EMEA") & ""
                oldAPJ=rs("APJ") & ""
                oldJustification=rs("Justification") & ""
			end if
			rs.Close
	
	        If oldZsrpRequired <> "" Then 
	            oldZsrpRequired = CBool(oldZsrpRequired)
	        Else
	            oldZsrpRequired = false
	        End If
	
			if (request("cboApproverStatus") = "2" OR request("cboApproverStatus") = "5") AND request("commentsonly") = "" then 'Approved
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spVerifyAutoApprove"
		
				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				'rs.Open "spVerifyAutoApprove " & request("txtID"), cn,adOpenForwardOnly
				if not(rs.EOF and rs.BOF) then
					if rs("Verified") = 1 then	
						if clng(strType) = 3 then				
							NewStatus = 4
						else
							NewStatus = 2
						end if
					end if
				end if
				rs.Close
			elseif request("cboApproverStatus") = "3" AND request("commentsonly") = "" then 'Disapproved
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spVerifyAutoApprove"
		
				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				if not(rs.EOF and rs.BOF) then
					if rs("DisapprovedCount") = 0 then	
						blnSendDisapprovedEmail = true
					end if
				end if
				rs.Close
			end if
		
		
                '''get Program Coordinator's Email, PCEmail 
        		set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 1
				cm.CommandText = "SELECT  ISNULL(ui.Email, '') AS PCEmail FROM ProductVersion AS v INNER JOIN BusinessSegment bs ON v.BusinessSegmentID = bs.BusinessSegmentID LEFT OUTER JOIN UserInfo ui ON v.PCID = ui.UserID WHERE v.TypeID IN (1, 3) AND v.ID = " + cstr(cint(strProductID)) + " AND bs.BusinessID=1 and bs.Operation =1; "

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				if rs.EOF and rs.BOF then
					strPCEmail = ""
				else
					strPCEmail = rs("PCEmail")	
				end if
		
				rs.Close
                '''get Program Coordinator's Email, PCEmail 

                '''Add PCEmail into Notify email list, strTxtNotify
                if len(strPCEmail)>5 and instr(strTxtNotify,strPCEmail)<1 then
                    strTxtNotifyFin = strPCEmail + ";" + strTxtNotify
                else
                    strTxtNotifyFin = strTxtNotify
                end if

         
				'Save updates
				set cm = server.CreateObject("ADODB.command")
		
				cm.ActiveConnection = cn
				cm.CommandText = "spUpdateDeliverableActionWeb2"
				cm.CommandType =  &H0004
	
				'cn.BeginTrans
				
				Set p = cm.CreateParameter("@ID", 3,  &H0001)
				p.value = clng( request("txtID"))
				cm.Parameters.Append p

				set p =  cm.CreateParameter("@ProductID", 3, &H0001)
				p.value = strProductID
				cm.Parameters.Append p
	            
				set p =  cm.CreateParameter("@DeliverableRootId", 3, &H0001)
				If trim(strDeliverableRootId) = "" Then
				    p.value = 0
				Else
				    p.value = strDeliverableRootId
				End If
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Type", 16, &H0001)
				p.Value = clng(strType)
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Status", 16, &H0001)
				p.Value = NewStatus
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Submitter", 200, &H0001, 50)
				p.Value = left(request("txtSubmitter"),50)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SubmitterID", 3, &H0001)
				if trim(request("txtSubmitterID")) = "" then
                    p.Value = 0
                else
                    p.Value = clng(request("txtSubmitterID"))
				end if
                cm.Parameters.Append p
		
				set p =  cm.CreateParameter("@CategoryID", 3, &H0001)
				p.value = 1 
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OwnerID", 3, &H0001)
				p.Value = request("cboOwner")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@PreinstallOwner", 3, &H0001)
				p.Value = clng(request("cboPreinstallApprover"))
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@AffectsCustomers", 16, &H0001)
		        if strType = "5" then
					p.value=request("cboNetAffect")
					strCustomerImpact = "Affects Images And/Or BIOS on shipping products."
		        else
					if request("chkCustomers") = "on" then
						p.Value = 1
					else
						p.Value = 0
					end if
				end if
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@CoreTeamID", 3, &H0001)
				p.Value = clng(request("cboCoreTeam"))
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@RoadmapID",adInteger, &H0001)
				p.Value = 0
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@TargetDate", 135, &H0001)
				if request("txtTarget")="" then
					p.Value = null
		        else
			        p.Value = CDate(request("txtTarget"))
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@ECNDate", 135, &H0001)
				if NewStatus = "4" And strType = "3" Then
				    p.Value = NOW()
				elseif trim(request("txtECNDate")) <> "" and isdate(request("txtECNDate")) then
					p.Value = CDate(request("txtECNDate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@TestDate", 135, &H0001)
				if trim(request("txtAvailDate")) <> "" and isdate(request("txtAvailDate")) then
					p.Value = CDate(request("txtAvailDate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@TestNote", 200, &H0001,35)
		        if strType = "5" then
					p.Value = left(request("cboMetricImpact"),35)
				else
					p.Value = left(request("txtAvailNotes"),35)
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@BTODate", 135, &H0001)
				if request("txtBTODate") <> "" then
					p.Value = CDate(request("txtBTODate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@CTODate", 135, &H0001)
				if request("txtCTODate") <> "" then
					p.Value = CDate(request("txtCTODate"))
				else
					p.value = null
				end if
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Notify", 200, &H0001, 8000)
				p.Value = left(strTxtNotifyFin,8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OnStatus", 16, &H0001)
				if lcase(trim(request("chkReports"))) = "on" then
					p.Value = 1
				else
					p.Value = 0
				end if
				cm.Parameters.Append p
	

				Set p = cm.CreateParameter("@Priority", 16, &H0001)
				if request("lstPriority") = "" then
					p.Value = 0
				else
					p.Value = clng(request("lstPriority"))
				end if
				cm.Parameters.Append p
	
		        Set p = cm.CreateParameter("@Commercial", 11, &H0001)
				if request("chkCommercial") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", Commercial"
				else
					p.Value = false
				end if
				cm.Parameters.Append p


		        Set p = cm.CreateParameter("@Consumer", 11, &H0001)
				if request("chkConsumer") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", Consumer"
				else
					p.Value = false
				end if
				cm.Parameters.Append p


		        Set p = cm.CreateParameter("@SMB", 11, &H0001)
				if request("chkSMB") = "on" then
					p.Value = true
					strBusiness = strBusiness & ", SMB"
				else
					p.Value = false
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@NA", 11, &H0001)
				if request("chkNA") = "on" then
					p.Value = True
					strRegions = strRegions & ", NA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@LA", 11, &H0001)
				if request("chkLA") = "on" then
					p.Value = True
					strRegions = strRegions & ", LA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@APJ", 11, &H0001)
				if request("chkAPJ") = "on" then
					p.Value = True
					strRegions = strRegions & ", APJ"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@EMEA" ,11, &H0001)
				if request("chkEMEA") = "on" then
					p.Value = True
					strRegions = strRegions & ", EMEA"
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@AddChange", 11, &H0001)
				if request("chkAdd") = "on" then
					p.Value = True
				else
					p.Value = False
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ModifyChange", 11, &H0001)
				p.Value = (request("chkModify") = "on")
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@RemoveChange", 11, &H0001)
				p.Value = (request("chkRemove") = "on")
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ImageChange", 11, &H0001)
				p.Value = bSoftware
				cm.Parameters.Append p

      		    Set p = cm.CreateParameter("@CategoryBiosChange", 11, &H0001)
				p.Value = bCategoryBiosChange
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@CommodityChange", 11, &H0001)
				p.Value = bCommodity
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@DocChange", 11, &H0001)
				p.Value = bDocs
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@SKUChange", 11, &H0001)
				p.Value = bSku
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ReqChange", 11, &H0001)
				p.Value = bRequirement
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@BiosChange", 11, &H0001)
				p.Value = bBios
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@SwChange", 11, &H0001)
				p.Value = bSCR
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@IDChange", 11, &H0001)
				p.Value = bID
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@OtherChange", 11, &H0001)
				p.Value = bOther
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@PendingImplementation", 11, &H0001)
				if request("chkClosed") = "on" then
					p.Value = true
				else
					p.Value = false
				end if
				cm.Parameters.Append p
            
				Set p = cm.CreateParameter("@Summary", 200, &H0001, 120)
				p.Value = left(request("txtSummary"),120)
				cm.Parameters.Append p
            
				Set p = cm.CreateParameter("@Description", 200, &H0001, 8000)
				p.Value = left(strDescription,8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Justification", 200, &H0001, 8000)
                if CurrentUserPartner = "1" then
			        p.Value = left(request("txtJustification"),8000)
                else
                    p.Value = null
                end if
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Details", 200, &H0001, 8000)
				p.Value = left(request("txtDetails"),8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@LastUpdUser", 200, &H0001, 200)
				p.Value = CurrentUserName
				cm.Parameters.Append p
	    
				Set p = cm.CreateParameter("@Actions", 201, &H0001, 2147483647)
				p.Value = request("txtActions")
				cm.Parameters.Append p
	    
		        Set p = cm.CreateParameter("@Resolution", 201, &H0001, 2147483647)
				'if request("txtFormType") = "Change" then
		            p.Value = request("txtResolution")
				'else
		        '    p.Value = ""
				'end if
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@ZsrpReadyTargetDt", 135, &H0001)
				if request("txtZsrpReadyTargetDt") <> "" then
					p.Value = CDate(request("txtZsrpReadyTargetDt"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

    			Set p = cm.CreateParameter("@ZsrpReadyActualDt", 135, &H0001)
				if request("txtZsrpReadyActualDt") <> "" then
					p.Value = CDate(request("txtZsrpReadyActualDt"))
				else
					p.value = null
				end if
				cm.Parameters.Append p

		        Set p = cm.CreateParameter("@ZsrpRequired", 11, &H0001)
				p.Value = bZsrpReadyRequired
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@RTPDate", 135, &H0001)
				if trim(request("txtRTPDate")) <> "" and isdate(trim(request("txtRTPDate"))) then
					p.Value = CDate(trim(request("txtRTPDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@RASDiscoDate", 135, &H0001)
				if trim(request("txtRASDiscoDate")) <> "" and isdate(trim(request("txtRASDiscoDate"))) then
					p.Value = CDate(trim(request("txtRASDiscoDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p                        

                Set p = cm.CreateParameter("@AVRequired", 11, &H0001)
				p.Value = bAVRequired
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@QualificationRequired", 11, &H0001)
				p.Value = bQualificationRequired
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@TargetApprovalDate", 135, &H0001)
				if trim(request("txtTargetApprovalDate")) <> "" and isdate(trim(request("txtTargetApprovalDate"))) then
					p.Value = CDate(trim(request("txtTargetApprovalDate")))
				else
					p.value = null
				end if
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@Important", 11, &H0001)
				p.Value = bImportant
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@Attachment1",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath1"),500)
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Attachment2",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath2"),500)
	            cm.Parameters.Append p

	            Set p = cm.CreateParameter("@Attachment3",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath3"),500)
	            cm.Parameters.Append p
	    
                Set p = cm.CreateParameter("@Attachment4",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath4"),500)
	            cm.Parameters.Append p

                Set p = cm.CreateParameter("@Attachment5",200, &H0001,500)
	            p.Value = left(request("txtAttachmentPath5"),500)
	            cm.Parameters.Append p

				cm.Execute RowsEffected
				
				set cm = nothing
				
				if cn.Errors.Count > 0 then
					'cn.RollbackTrans
					Errors = true
					Response.Write "Failed"
				else
					Response.Write "Continuing"
					blnUpdateApprovals = false
					'Add New Approvers
					strApproverList = request("Approvers2Add")
					
					if right(strApproverList,1) <> "," and len(strApproverList) > 0 then
						strApproverList = strApproverList & ","
					end if
					
					strApproverEmails = ""
					'Convert to Email List here

					do while instr(strApproverList,",")> 0
						Response.Write request("txtID") & "," & left(strApproverList,instr(strApproverList,",")-1) & "<BR>"
						
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spAddApprover"
		

						Set p = cm.CreateParameter("@ID", 3, &H0001)
						p.Value = request("txtID")
						cm.Parameters.Append p
	
						Set p = cm.CreateParameter("@ApproverID", 3, &H0001)
						p.Value = left(strApproverList,instr(strApproverList,",")-1)
						cm.Parameters.Append p

						cm.Execute RowsEffected
						Set cm=nothing


						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spGetEmployeeByID"

						Set p = cm.CreateParameter("@ID", 3, &H0001)
						p.Value = left(strApproverList,instr(strApproverList,",")-1)
						cm.Parameters.Append p

						rs.CursorType = adOpenForwardOnly
						rs.LockType=AdLockReadOnly
						Set rs = cm.Execute 
						Set cm=nothing

					    if not(rs.EOF and rs.BOF) then
							strApproverEmails= strApproverEmails &  rs("Email") & ";"		
						end if
						
						rs.Close
											
						strApproverList = mid(strApproverList,instr(strApproverList,",")+1)
						blnUpdateApprovals = true
					loop
					
					Response.Write "ApproverAdded"
					response.write "." & request("cboApproverStatus") & "."
					'Update Approver Status
					if request("txtSaveApproval") <> "" and request("txtSaveApproval") <> "0" then
						set cm = server.CreateObject("ADODB.command")
		
						cm.ActiveConnection = cn
						cm.CommandText = "spUpdateApproval"
						cm.CommandType =  &H0004
	
						Set p = cm.CreateParameter("@ApprovalID", 3,  &H0001)
						p.value = clng( request("txtSaveApproval"))
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@Status", 16,  &H0001)
						If request("cboApproverStatus") <> "" Then
						    p.value = clng(request("cboApproverStatus"))
						end if
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@Comments", 200, &H0001, 300)
						p.Value = left(request("txtApproverComments"),300)
						cm.Parameters.Append p

	
						cm.execute Rowseffected
						if rowseffected <> 1 then
							Errors= true
						else
							blnUpdateApprovals = true											
						end if
					end if
						
					if strGroupsAdded <> "" then
						for i = lbound(GroupAddArray) to ubound(GroupAddArray)
							set cm = server.CreateObject("ADODB.Command")
							Set cm.ActiveConnection = cn
							cm.CommandType = 4
							cm.CommandText = "spAddActionGroup"
		
							Set p = cm.CreateParameter("@ID", 3, &H0001)
							p.Value = clng( request("txtID"))
							cm.Parameters.Append p

							Set p = cm.CreateParameter("@GroupID", 3, &H0001)
							p.Value = clng( GroupAddArray(i))
							cm.Parameters.Append p
	
							cm.Execute rowseffected
							Set cm=nothing
							
							if rowseffected <> 1 then
								Errors= true
								exit for
							end if
						next
					end if

					if strGroupsRemoved <> "" then
						for i = lbound(GroupRemoveArray) to ubound(GroupRemoveArray)
							set cm = server.CreateObject("ADODB.Command")
							Set cm.ActiveConnection = cn
							cm.CommandType = 4
							cm.CommandText = "spRemoveActionGroup"
		
							Set p = cm.CreateParameter("@ID", 3, &H0001)
							p.Value = clng( request("txtID"))
							cm.Parameters.Append p

							Set p = cm.CreateParameter("@GroupID", 3, &H0001)
							p.Value = clng( GroupRemoveArray(i))
							cm.Parameters.Append p
	
							cm.Execute rowseffected
							Set cm=nothing
							
							if rowseffected <> 1 then
								Errors= true
								exit for
							end if
						next
					end if
					
					if (not errors) and blnUpdateApprovals then
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spSetApprovalList"
		
						Set p = cm.CreateParameter("@ID", 3, &H0001)
						p.Value = clng( request("txtID"))
						cm.Parameters.Append p
	
						cm.Execute 
						Set cm=nothing

						'cn.Execute "spSetApprovalList " & clng( request("txtID"))
					end if
					
				'	if not errors then
				'		cn.CommitTrans
				'	end if
				end if


		    if cint(newStatus & "") <> cint(oldstatus & "") then
			    'insert into history
				
			    set cm = server.CreateObject("ADODB.Command")
			    Set cm.ActiveConnection = cn
			    cm.CommandType = 1
			    cm.CommandText = "INSERT INTO DeliverableIssuesHistory (ChangeDt,ID,ProductVersionID,Type, Status,OwnerID,ActualDate,LastUpdUser) SELECT top 1 GETDATE(), ID , ProductVersionID , 3 ," & cint(newStatus & "") & " , OwnerID , ActualDate , " & clng(currentuserid ) & "  FROM DeliverableIssues where ID = " & clng( request("txtID"))
			   
			    cm.Execute 
			    Set cm=nothing
		    end if
				
	
		Response.Write "Done"  'Done Save updates
'
'   Start Email Section
'
		Dim oMessage 

		if errors then
				response.write "<label style=""Display:"" ID=Results><b><font face=verdana>Unable to submit your " & strTypeName & " at this time.</font></b><BR><BR>An Unexpected Error Occurred.<BR><BR></label>"
		else
				Response.Write "<label style=""Display:"" ID=Results><h3><font face=verdana>" & strTypeName & " updated.</font></h3>"
		
		dim strNewOwnerEmail
		dim strSubmitterEmail
        dim strSETestLead
		dim strDivision
        dim strOperation
		Response.Write request("cboOwner") & "<BR>"
		
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetEmployeeByID"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("cboOwner")
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
        if request("txtSubmitterID") = "" then
		    p.Value = clng(currentuserid)
        else
		    p.Value = clng(request("txtSubmitterID"))
		end if
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
		p.Value = strproductid
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		strProductEmail = rs("Distribution") & ""		
		strPM = rs("PMName") & ""
        strPMEmail = rs("PMEmail") & ""
        strSETestLead = rs("SETestLead") & ""
		strDivision = rs("Division") & ""
        strOperation = rs("Operation") & ""
		strBusinessID = rs("BrandBusinessID")
		rs.Close
		
		DisplayedID = request("txtID")


        'Lookup SETestLeadEmail
        if strSETestLead = 0 then
            strSETestLead=""
        else
		    set cm = server.CreateObject("ADODB.Command")
		    Set cm.ActiveConnection = cn
		    cm.CommandType = 4
		    cm.CommandText = "spGetEmployeeByID"
		

		    Set p = cm.CreateParameter("@ID", 3, &H0001)
		    p.Value = clng(strSETestLead)
		    cm.Parameters.Append p
	

		    rs.CursorType = adOpenForwardOnly
		    rs.LockType=AdLockReadOnly
		    Set rs = cm.Execute 
		    Set cm=nothing

		    strSETestLead = rs("Email") & ""
		    rs.Close
		end if
		
		'Notifications
		
	    dim notifycount
	    notifycount = 0
	
		strTO = ""

		if request("chkDocChange") = "on" then
			if request("txtID") <> "" then 'Only notify them when editing
				if strTo = "" then
					strTO = "houdcrdocs@hp.com;"
				else
					strTO = strTO & "houdcrdocs@hp.com;"
				end if
			end if
		end if
	
		'---------
		
		strGroupList=""
		if request("lstFunctionalGroup") <> "" then
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
		end if

	    '
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
			strBody = "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br><br>"
		else
		
			If strType = "4" And rs("Status") = "4" Then
                strTo = strTo & ";houprtdcrnotif@hp.com;"
			End If
						
			TypeID = strType
			Select Case strType
			Case "1"
				strtype = "Issue"
			Case "2"
				strtype = "Action Item"
			Case "3"
				strtype = "Change Request"
			Case "4"
				strtype = "Status Note"
			Case "5"
				strtype = "Improvement Opportunity"
			Case "6"
				strtype = "Test Request"
			End Select
			
			strProgramName = rs("Program")
			strProgramMail=  rs("EmailActive")
			
			
			strBusiness = ""
			if (not isnull(rs("Consumer"))) and (not isnull(rs("Commercial"))) and (not isnull(rs("SMB"))) then
				if rs("Consumer") then
					strBusiness = strBusiness & ", Consumer"  
				end if
				if rs("Commercial") then
					strBusiness = strBusiness & ", Commercial"  
				end if
				if rs("SMB") then
					strBusiness = strBusiness & ", SMB"  
				end if
			end if
			
			strExcaliburHPLink = ""
            strExcaliburODMLink = ""
			strExcaliburHPLink = "<font face=Arial size=2><a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br>"	
			strExcaliburHPLink = strExcaliburHPLink & "<a href=""https://" & Application("Excalibur_ODM_ServerName") & "/excalibur/Excalibur.asp"">Open Pulsar Today Page</a></font><br><br>"
            strExcaliburODMLink = strExcaliburHPLink
       
           If bImportant Then
				      strBody = "<font face=Arial size=4 color=red><b>This Change Request has been identified as IMPORTANT!. Please handle quickly.</b></font><br><br><font face=Arial size=2>"
                else
              strBody = "<font face=Arial size=2>"
           End If	

			strBody = strBody & "<b>NUMBER:</B> " & DisplayedID & "<BR>"
			strBody = strBody & "<b>SUBMITTER:</b> " & rs("Submitter") & "<BR>"
			If IsNull(rs("Created")) Then
				strBody = strBody & "<b>SUBMITTED:</b> N/A" & "<BR>"
			Else
				strBody = strBody & "<b>SUBMITTED:</b> " & formatdatetime(rs("Created"), vbshortdate) & "<BR>"
			End If
			strProductName = rs("Program") & ""
			strBody = strBody & "<b>PROGRAM:</b> " & rs("Program") & "<BR>"
			
            'Harris, Valerie -  02/12/2016 - BUG 15660/ Task 16234 - If Type 3, Add Release section email to capture Product's selected Releases for DCR
            If CInt(TypeID) = 3 And rs("ProductVersionRelease") <> "" Then
                strBody = strBody & "<b>RELEASE:</b> " & rs("ProductVersionRelease") & "<BR>"
            End If

			If bScr Then
    			strBody = strBody & "<b>DELIVERABLE ROOT:</b> " & rs("DeliverableRoot") & "<BR>"
			End If
			
			strBody = strBody & "<b>TYPE:</b> " & strtype & "<BR>"
			if trim(TypeID) = "5" then
				strBody = strBody & "<b>ISSUE/ACCOMPLISHMENT:</B> " & replace(request("txtSummary"),"""","&QUOT;") & "<BR>"
			else
                 if (trim(oldSummary) <> trim(request("txtSummary")))  then
			        strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			        Response.Write "<BR>DCR Summary Changed: " & strTO &  "<BR>"
			        notifycount = notifycount + 1
			        strSubject = strSubject & "(DCR Updated)"			        
                    strBody = strBody & "<font face=Arial size=2 color=red><b>SUMMARY:</B> " & replace(server.HTMLEncode(request("txtSummary")),"""","&QUOT;") & "</font><BR>"                    
                else
                    strBody = strBody & "<b>SUMMARY:</B> " & replace(server.HTMLEncode(request("txtSummary")),"""","&QUOT;") & "<BR>"
		        end if				
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
			strNewOwner = request("cboOwner")
			strOwnerName = rs("Owner") & ""
			if strBusiness <> "" then
                  consumer=""
                  SMB=""
                  Commercial=""
                  if request("chkConsumer") = "on" then
                    consumer="True"
                  else
                    consumer="False"
                  end if
                  if request("chkSMB") = "on" then
                    SMB="True"
                  else
                    SMB="False"
                  end if
                  if request("chkCommercial") = "on" then
                    Commercial="True"
                  else
                    Commercial="False"
                  end if

                  if (oldConsumer <> consumer or oldSMB <> SMB or  oldCommercial <> Commercial) then
                        strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			            Response.Write "<BR>DCR Business Changed: " & strTO &  "<BR>"
			            notifycount = notifycount + 1
			            strSubject = strSubject & "(DCR Updated)"			            
                        strBody = strBody & "<font face=Arial size=2 color=red><b>BUSINESS:</b> " & Mid(strBusiness, 3) & "</font><BR>"                        
                  else
                        strBody = strBody & "<b>BUSINESS:</b> " & Mid(strBusiness, 3) & "<BR>"
                  end if
				
			end if
			if strRegions <> "" Then
                  NA=""
                  LA=""
                  APJ=""
                  EMEA=""
                  if request("chkNA") = "on" then
                    NA="True"
                  else
                    NA="False"
                  end if
                  if request("chkLA") = "on" then
                    LA="True"
                  else
                    LA="False"
                  end if
                  if request("chkAPJ") = "on" then
                    APJ="True"
                  else
                    APJ="False"
                  end if
                  if request("chkEMEA") = "on" then
                    EMEA="True"
                  else
                    EMEA="False"
                  end if
                 if (oldNA <> NA or oldLA <> LA or oldAPJ <> APJ or oldEMEA <> EMEA) then
                        strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			            Response.Write "<BR>DCR Region Changed: " & strTO &  "<BR>"
			            notifycount = notifycount + 1
			            strSubject = strSubject & "(DCR Updated)"			            
                        strBody = strBody & "<font face=Arial size=2 color=red><b>REGIONS:</b> " & Mid(strRegions, 3) & "</font><br>"                        
                  else
                        strBody = strBody & "<b>REGIONS:</b> " & Mid(strRegions, 3) & "<br>"
                  end if			               
			end if
			if rs("Description") & "" <> "" then
				strDescription = replace(server.HTMLEncode(rs("Description"))& "",vbcrlf,"<BR>")      
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
                    if (trim(oldDescription) <> trim(request("txtDescription"))) then
                        strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			            Response.Write "<BR>DCR Description Changed: " & strTO &  "<BR>"
			            notifycount = notifycount + 1
			            strSubject = strSubject & "(DCR Updated)"			            
                        strBody = strBody & "<font face=Arial size=2 color=red><b>DESCRIPTION:</b> " & "<BR>" & strDescription & "</font><BR>" & "<BR>"                        
                    else
                        strBody = strBody & "<b>DESCRIPTION:</b> " & "<BR>" & strDescription & "<BR>" & "<BR>"
                    end if
				end if
			end if
			If trim(TypeID) = "3" Then
			    strBody = strBody & "<b>DETAILS:</b> "  & "<BR>" & Replace(server.HTMLEncode(request("txtDetails")), VbCrLf, "<BR>") & "<BR>"
			End If

            'use same email body as above for ODM
            strODMBody = strBody

			if rs("Justification") & "" <> "" then
				if trim(TypeID) = "5" then
					strBody = strBody & "<b>ROOT CAUSE:</b><BR>" & replace(server.HTMLEncode(rs("Justification")),vbcrlf,"<BR>") & "<BR><BR>"
				else					
                   if (trim(oldJustification) <> trim(request("txtJustification"))) then
                        strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			            Response.Write "<BR>DCR Justification Changed: " & strTO &  "<BR>"
			            notifycount = notifycount + 1
			            strSubject = strSubject & "(DCR Updated)"			            
                        strBody = strBody & "<font face=Arial size=2 color=red><b>JUSTIFICATION:</b><BR>" & replace(server.HTMLEncode(rs("Justification")),vbcrlf,"<BR>") & "</font><BR><BR>"
                        strODMBody = strODMBody & "<font face=Arial size=2 color=red><b>JUSTIFICATION:</b><BR>" & replace(server.HTMLEncode(rs("Justification")),vbcrlf,"<BR>") & "</font><BR><BR>"                        
                  else
                        strBody = strBody & "<b>JUSTIFICATION:</b><BR>" & replace(server.HTMLEncode(rs("Justification")),vbcrlf,"<BR>") & "<BR><BR>"
                  end if
				end if
			end if

			If strtype <> "Status Note" Then
			
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
				
                If bZsrpReadyRequired Then
                    strBody = strBody & "<b>ZSRP READY TARGET: </b> " & request("txtZsrpReadyTargetDt") & "<br>"
                    strBody = strBody & "<b>ZSRP READY ACTUAL: </b> " & request("txtZsrpReadyActualDt") & "<br><br><br>"
                    strODMBody = strODMBody & "<b>ZSRP READY TARGET: </b> " & request("txtZsrpReadyTargetDt") & "<br>"
                    strODMBody = strODMBody & "<b>ZSRP READY ACTUAL: </b> " & request("txtZsrpReadyActualDt") & "<br><br><br>"
                End If					    

				if trim(TypeID) = "5" then
					strBody = strBody & "<font color=red>" & server.HTMLEncode(strActions) & "</font>" & "<BR>" & "<BR>"
                    strODMBody = strODMBody & "<font color=red>" & server.HTMLEncode(strActions) & "</font>" & "<BR>" & "<BR>"
				else
					strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & server.HTMLEncode(strActions) & "</font>" & "<BR>" & "<BR>"
                    strODMBody = strODMBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & server.HTMLEncode(strActions) & "</font>" & "<BR>" & "<BR>"
				end if
				strBody = strBody & "<b>RESOLUTION:</b> " & "<BR>" & replace(server.HTMLEncode(rs("Resolution")) & "",vbcrlf,"<BR>") & "<BR>"
                strODMBody = strODMBody & "<b>RESOLUTION:</b> " & "<BR>" & replace(server.HTMLEncode(rs("Resolution")) & "",vbcrlf,"<BR>") & "<BR>"
			End If

			if trim(request("txtAvailDate")) <> "" then
				strBody = strBody & "<font color=red><b>SAMPLES AVAILABLE:</b> " & request("txtAvailDate") & "</font><BR>"
                strODMBody = strODMBody & "<font color=red><b>SAMPLES AVAILABLE:</b> " & request("txtAvailDate") & "</font><BR>"
			end if

            if trim(TypeID) = "3" then
                if trim(request("txtTargetApprovalDate")) <> "" then
                    strBody = strBody & "<font color=red><b>Target Approval Date:</b> " & request("txtTargetApprovalDate") & "</font><BR>"
                    strODMBody = strODMBody & "<font color=red><b>Target Approval Date:</b> " & request("txtTargetApprovalDate") & "</font><BR>"
                end if
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
                strODMBody = strODMBody & "<b>IMPACT:</b> " & strPriority & "<BR>"
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
                strODMBody = strODMBody & "<b>NET AFFECT:</b> " & strCustomers & "<BR>"
			end if

			if trim(TypeID) = "5"  and trim(rs("AvailableNotes") & "") <> "" then
				strBody = strBody & "<b>METRIC IMPACTED:</b> " & rs("AvailableNotes") & "<BR>"
                strODMBody = strODMBody & "<b>METRIC IMPACTED:</b> " & rs("AvailableNotes") & "<BR>"
			end if


			strCoreTeam = rs("CoreTeamRep")
			if (not isnull(rs("BTODate"))) and (rs("Distribution") = "BTO" or rs("Distribution") = "BOTH") then
		        strBody = strBody & "<b>BTO-IMPLEMENT BY:</b> " & rs("BTODate") & "<BR>"
                strODMBody = strODMBody & "<b>BTO-IMPLEMENT BY:</b> " & rs("BTODate") & "<BR>"
			end if
			if (not isnull(rs("CTODate"))) and (rs("Distribution") = "CTO" or rs("Distribution") = "BOTH") then
		        strBody = strBody & "<b>CTO-IMPLEMENT BY:</b> " & rs("CTODate") & "<BR>"
                strODMBody = strODMBody & "<b>CTO-IMPLEMENT BY:</b> " & rs("CTODate") & "<BR>"
			end if

			If strtype <> "Status Note" Then
				strBody = strBody & "<b>NOTIFY ON CLOSURE:</b> " & rs("Notify") & "<BR>"
                strODMBody = strODMBody & "<b>NOTIFY ON CLOSURE:</b> " & rs("Notify") & "<BR>"
			End If
			strNotify = rs("Notify") & ""
			if rs("Approvals") & "" <> "" then
				strBody = strBody & "<b>APPROVALS:</b><font color=teal><BR>" & replace(rs("Approvals"),vbcrlf,"<BR>") & "</font><BR><BR>"
                strODMBody = strODMBody & "<b>APPROVALS:</b><font color=teal><BR>" & replace(rs("Approvals"),vbcrlf,"<BR>") & "</font><BR><BR>"
			end if
			if strGroupList <> "" then
				strBody = strBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
                strODMBody = strODMBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
			end if

			strBody = strBody & "<br><font size=1 color=red face=verdana>HP Restricted</font>"
            strODMBody = strODMBody & "<br><font size=1 color=red face=verdana>HP Restricted</font>"

		end if
		rs.Close

		
		'
		' Get the PC's Email Address
		'
		rs.open "spListSystemTeam " & strproductid,cn,adOpenForwardOnly
					
        do while not rs.eof
            if trim(rs("Role") & "") = "Program Coordinator" then
                strPCEmail = rs("Email")
            end if
           rs.movenext	
		loop
       rs.close
        
        
        set rs = nothing
		cn.Close
		set cn = nothing
		Response.Write "<font size=2 face =verdana>"
		Response.Write "<font face=verdana size=3><b>Notifications Sent:</b></font>"
		if NewStatus = 3 and oldStatus <> "3" then 'More input needed
			strTo = strTo & ";" & strNewOwnerEmail & ";" & strSubmitterEmail & ";"
			Response.Write "<BR>More Information Needed: " & strTO & "<BR>"		
			notifycount = notifycount + 1
            strSubject = strSubject & "(Need more info)"
			strBody = "<font face=Arial size=2 color=red><b>Please provide more information for this " & strtype & "</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>Please provide more information for this " & strtype & "</font></b><BR><BR>" & strODMBody 
		end if

		if blnSendDisapprovedEmail then 'First Disapprover Found
			strTo = strTO & ";" & strNewOwnerEmail & ";"
			Response.Write "<BR>First time disapproved by Approver: " & strTO & "<BR>"		
			notifycount = notifycount + 1
		    strSubject = strSubject & "(Disapproved by Approver)"
			strBody = "<font face=Arial size=2 color=red><b>This item has been disapproved by an approver. Further disapprovals will not generate and email notification.</font></b><BR><BR>" & strBody
            strODMBody = "<font face=Arial size=2 color=red><b>This item has been disapproved by an approver. Further disapprovals will not generate and email notification.</font></b><BR><BR>" & strODMBody 
		end if


		if strApproverEmails <> "" then 'Approvers Added
			strTo = strTO & strApproverEmails 
			Response.Write "<BR>New Approvers: " & strTO & "<BR>"		
			notifycount = notifycount + 1
			if newStatus = 6 and oldStatus <> "6" then
					strCC = strCC & strProductEmail & ";"
			end if
			strSubject = strSubject & "(Approval Requested)"
			strBody = "<font face=Arial size=2 color=red><b>You have been added to the Approval List</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>You have been added to the Approval List</font></b><BR><BR>" & strODMBody 
		end if

		if newStatus = 6 and oldStatus <> "6" then 'Set to Investigating
            strTo = strTO & ";" & strProductEmail & ";" & strCurrentApproverEmails & ";"
			If trim(strDivision) = "1" And (Not bBios) And (Not bSCR) and (not bID) and LCase(Request.Form("hidAddDCRNotificationList")) = "true" Then
                if trim(strOperation) = "0" then 
				    strTo = strTO & ";NotebookDCRNotification@hp.com;"
                end if
			End If
				
			Response.Write "<BR>Investigating: " & strTO & "<BR>"		
			notifycount = notifycount + 1
			strSubject = strSubject & "(Investigating)"
			strBody = "<font face=Arial size=2 color=red><b>Status set to Investigating</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>Status set to Investigating</font></b><BR><BR>" & strODMBody
		end if

		if trim(oldOwner) <> trim(strNewOwner) then 'New owner assigned
			strTo = strTO & strNewOwnerEmail & ";"
			Response.Write "<BR>New Owner Assignment: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(Reassigned)"
			strBody = "<font face=Arial size=2 color=red><b>This " & strType & " has been assigned to " & strOwnerName & ".</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>This " & strType & " has been assigned to " & strOwnerName & ".</font></b><BR><BR>" & strODMBody
		end if
		
		if oldZsrpRequired And NOT bZsrpReadyRequired Then ' ZSRP Required Bit Changed
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strPCEmail & ";" & strCurrentApproverEmails & ";"
			Response.Write "<BR>ZSRP Requirement Removed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Requirement Removed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Requirement Removed.</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>ZSRP Requirement Removed.</font></b><BR><BR>" & strODMBody 
		end if
		
		if (trim(oldZsrpTargetDt) <> trim(request("txtZsrpReadyTargetDt"))) And bZsrpReadyRequired Then ' ZSRP Ready Target Date Changed
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strPCEmail & ";" & strCurrentApproverEmails & ";"
			if LCase(Request.Form("hidAddDCRNotificationList")) = "true" then
			    if trim(strOperation) = "0" then
				    strTo = strTO & ";NotebookDCRNotification@hp.com;"
                end if
			end if
			Response.Write "<BR>ZSRP Target Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Target Date Changed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Target Date Updated.</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>ZSRP Target Date Updated.</font></b><BR><BR>" & strODMBody
		end if
		
		if (trim(oldZsrpActualDt) <> trim(request("txtZsrpReadyActualDt"))) And bZsrpReadyRequired Then ' ZSRP Ready Actual Date Changed
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strPCEmail & ";" & strCurrentApproverEmails & ";"
			if LCase(Request.Form("hidAddDCRNotificationList")) = "true" then
			    if trim(strOperation) = "0" then
				    strTo = strTO & ";NotebookDCRNotification@hp.com;"
                end if
			end if
			Response.Write "<BR>ZSRP Actual Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Actual Date Changed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Actual Date Updated.</font></b><BR><BR>" & strBody
            strODMBody = "<font face=Arial size=2 color=red><b>ZSRP Actual Date Updated.</font></b><BR><BR>" & strODMBody  
		end if
		
		if (trim(oldTargetDate) <> trim(request("txtTarget"))) and request("txtID") <> "" then 'TargetDate Updated
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";" & strCurrentApproverEmails & ";"
			Response.Write "<BR>Target Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(Target Changed)"
			strBody = "<font face=Arial size=2 color=red><b>Target Date Updated.</font></b><BR><BR>" & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>Target Date Updated.</font></b><BR><BR>" & strODMBody
		end if

      
      

     
      
     
		
		If (newStatus = 2 or NewStatus = 4 or NewStatus = 5) and (oldstatus <> trim(cstr(newStatus)) ) Then ' Status Changed
			Response.Write strtype & ":" & strProductEmail & "<BR>"
			if TypeID = "3" then
				strTO = strTO & strNewOwnerEmail  & ";" & strProductEmail & ";" & strSubmitterEmail & ";" & strCurrentApproverEmails & ";"
				if strTxtNotifyFin <> "" then
						strTO = strTO & strTxtNotifyFin & ";" 
				end if
                if NewStatus = 4 then
                    strTo = strTO & strSETestLead & ";"
                end if

				if trim(strDivision) = "1"  And (Not bBios) And (Not bSCR) and (not bID) and LCase(Request.Form("hidAddDCRNotificationList")) = "true" then
					if trim(strOperation) = "0" then
				        strTo = strTO & ";NotebookDCRNotification@hp.com;"
                    end if
				end if
				Response.Write "<BR>Item Closed: " & strTO & "<BR>"
			else
				strTO = strTO & strNewOwnerEmail  & ";" & strSubmitterEmail  & ";"
				if strTxtNotifyFin <> "" then
					if strTO = "" then
						strTO = strTxtNotifyFin & ";"
					else
						strTO = strTO & strTxtNotifyFin & ";"
					end if
				end if
				strTo = replace(StrTo,CurrentUserEmail & ";","")
				Response.Write "<BR>Item Closed: " & strTO &  "<BR>"
			end if
			notifycount = notifycount + 1
			strSubject = strSubject & "(" & strStatus & ")"
			strBody = "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strExcaliburHPLink & strBody 
            strODMBody = "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strExcaliburODMLink & strODMBody
		end if


		end if
		
		if notifycount = 0 then
			Response.Write "<BR>None."
		end if
		
        'strTo = Replace(strTo,CurrentUserEmail & ";","")

        if strCoreTeam = "Sustaining System Team" then
            strTO = strTo & ";" & CCTMailList & ";"
        end if
        
        Do Until InStr(strTo, ";;") = 0
            strTo = Replace(strTo, ";;", ";")
        Loop

		if notifycount > 0 And strTo <> "" then

	            if notifycount > 1 then
	                If instr(strSubject, "Approved") > 0 And instr(strSubject, "Disapproved") = 0 Then
	                    strSubject = "(Approved)"
	                elseif instr(strSubject, "(Disapproved)") > 0 Then
	                    strSubject = "(Disapproved)"
	                elseif instr(strSubject, "(Disapproved by Approver)") > 0 Then
	                    strSubject = "(Disapproved by Approver)"
	                elseif instr(strSubject, "ZSRP") > 0 Then
	                    strSubject = "(ZSRP Updates)"
	                else
	                    strSubject = "(Multiple Changes)"
	                end if
	            End If
	            
	            strSubject = strtype & " " & DisplayedID &  " " & strSubject & " : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
	            
                'separate the hp and odm email addresses
                strHPEmailAddress = ""
                strODMEmailAddress = ""
                EmailArray = getUniqueItems(split(strTo,";"))
               
				for each emailaddress in EmailArray
                    if Len(Trim(emailaddress)) > 0 then  
					    if instr(UCase(emailaddress), "@HP.COM") = 0 then
                            strODMEmailAddress = strODMEmailAddress & ";" & emailaddress
                        else
                            strHPEmailAddress = strHPEmailAddress & ";" & emailaddress
                        end if
                    end if		
				next

                if len(strHPEmailAddress) > 0 then
					strHPEmailAddress = mid(strHPEmailAddress,2)
				end if

	            if strProgramMail <> "1" then
	                strSubject = "TEST MAIL: " & strSubject
	                strBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strHPEmailAddress & " " & strCC & "</font><BR><BR>" & strBody
                    strTo = CurrentUserEmail
                else
                    strTo = strHPEmailAddress
	            end if

                strODMCC = ""
                strHPCC = ""
                EmailArray = getUniqueItems(split(strCC,";"))
				for each emailaddress in EmailArray
                    if Len(Trim(emailaddress)) > 0 then  
					    if instr(UCase(emailaddress), "@HP.COM") = 0 then
                            strODMCC = strODMCC & ";" & emailaddress
                        else
                            strHPCC = strHPCC & ";" & emailaddress
                        end if
                    end if		
				next
	            
                'email to hp users
				Set oMessage = New EmailQueue 
				oMessage.From = CurrentUserEmail
				oMessage.To= strTo
				If Len(Trim(strHPCC)) > 0 Then
				    oMessage.Cc = strHPCC
				End If
				oMessage.Subject =  strSubject
				oMessage.HtmlBody = strBody

     
				oMessage.SendWithOutCopy 'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				Set oMessage = Nothing 	
               
                '//RTP and EOM Date updated in Schedule Tab
if NewStatus = 4 then 
	dim OpenWorkFlowCount
    dim WorkFlowId
    dim workflowUserID                
    dim workflowPVID
    dim workflowDCRID
    strBody=""   
   
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 60
	cn.Open

   set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "usp_DCRWorkflowCheck"

    Set p = cm.CreateParameter("@DCRID", 3, &H0001)
	p.Value = request("txtID")
	cm.Parameters.Append p
	
	set rs = server.CreateObject("ADODB.recordset")
	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	if not(rs.EOF and rs.BOF) then
		OpenWorkFlowCount = rs("OpenWorkflowCount")			
	end if
	Set cm=nothing
	rs.Close
	
	if OpenWorkFlowCount >= 1 then 
		if productVersionRelease <> "" then 
            releaseNames = Split(productVersionRelease,",")
            
			for each releaseName in releaseNames                
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
				cm.CommandType = 4
				cm.CommandText = "usp_GetScheduleRTPandEOMDatebyProductID"

				Set p = cm.CreateParameter("@ProductId", 3, &H0001)
				p.Value = strProductID
				cm.Parameters.Append p

                Set p = cm.CreateParameter("@ProductReleaseName", 200, &H0001, 80)
	            p.Value = releaseName
	            cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing
					if not(rs.EOF and rs.BOF) then
						rtpProjectedDate=rs("Projected_Start_dt")	
                        eomProjectedDate=rs("Projected_Start_dt")                        
                        strBody = "All,<BR><BR>DCR " & request("txtID") & " is approved for " & productName & "(" & releaseName & ") to change the entire Product Schedule.<BR><BR> As a result, the Current Commitment column on the Product Schedule Tab will be modified to reflect this change.<BR><BR>Dates BEFORE this change were: <BR><BR>RTP: " & rtpProjectedDate & " <BR><BR>End of Manufacturing:" & eomProjectedDate & " <BR><BR> The RTP Date will be " & request("txtRTPDate") & ".<BR><BR> The End of Manufacturing Date will be " & request("txtRASDiscoDate") & ".<BR><BR>A Workflow Milestone is was added to the DCR and is being processed now.  Another email will be automatically sent when the Milestone is complete."
						Set scheduleOMessage = New EmailQueue 
						scheduleOMessage.From = CurrentUserEmail
						scheduleOMessage.To= strTo
						If Len(Trim(strHPCC)) > 0 Then
							scheduleOMessage.Cc = strHPCC
						end if
						scheduleOMessage.Subject =  "Product Schedule Change"
						scheduleOMessage.HtmlBody = strBody                        
						scheduleOMessage.SendWithOutCopy 'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
						Set scheduleOMessage = Nothing 	
					end if					
			next
		end if 
      rs.Close     
	end if
	cn.close
end if


                'email to ODM users
                if strODMEmailAddress <> "" then
                    strODMEmailAddress = mid(strODMEmailAddress,2)					
                    if strProgramMail <> "1" then
	                    strSubject = "TEST MAIL: " & strSubject
	                    strODMBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strODMEmailAddress & " " & strODMCC & "</font><BR><BR>" & strODMBody
	                    strTo = CurrentUserEmail
                    else
                        strTo = strODMEmailAddress
	                end if
	            
                    'email to ODM users
				    Set oODMMessage = New EmailQueue
				    oODMMessage.From = CurrentUserEmail
				    oODMMessage.To= strTo
                    If Len(Trim(strODMCC)) > 0 Then
				        oODMMessage.Cc = strODMCC
				    End If
                    
				    oODMMessage.Subject =  strSubject
				    oODMMessage.HtmlBody = strODMBody
                   
				    oODMMessage.SendWithOutCopy 'No need to send copy to the "From", because it is in the list of "To" or forbidden to get.
				    Set oODMMessage = Nothing 	
             end if
		end if
		
		Response.Write "</font>"	
	end if	

	'end if	



	
		
	if request("txtID") = "" then 'Adding
  %>
  <input type="text" style="display: none" id="DisplayedID" name="DisplayedID" value="">
  <%else%>
  <input type="text" style="display: none" id="DisplayedID" name="DisplayedID" value="<%=DisplayedID%>">
  <%end if%>
  <input type="hidden" id="layout" value="<%=Request("layout")%>" />
</body>
</html>
