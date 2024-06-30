<%@  language="VBScript" %>
<%Response.Buffer = true %>
<!-- #include file="../includes/emailwrapper.asp" -->
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <title>ECR Action Save</title>
    <script type="text/javascript">
<!--
        function window_onload() {
            var SavePrompt = document.getElementById("SavePropmt");
            var Results = document.getElementById("Results");

            if (document.getElementById("DisplayedID").value == "") {
                if (SavePrompt != null) { SavePrompt.style.display = "none"; }
                if (Results != null) { Results.style.display = ""; }
                window.returnValue = 1;
            }
            else {

                if (SavePrompt != null) { SavePrompt.style.display = "none"; }
                if (Results != null) { Results.style.display = ""; }
                window.returnValue = 1;
                window.parent.opener = 'X';
                window.parent.open('', '_parent', '')
                window.parent.close();
            }
        }

//-->
    </script>

<style type="text/css">
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
</head>
<body style="background-color:FFFFF0" onload="return window_onload()">
<%

	Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim TypeID : TypeID = regEx.Replace(Request.Form("txtType"), "")
    Dim ProdID : ProdID = regEx.Replace(Request.QueryString("ProdID"), "")
    Dim IssueID : IssueID = regEx.Replace(Request.QueryString("ID"), "")
    Dim CategoryID : CategoryID = regEx.Replace(Request.QueryString("CAT"), "")  
    
If cint(TypeID) <> 7 Then
    Response.Write "<h1 style=""color:red"">Incorrect Type Submission</h1>"
    Response.End
End If

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
	dim strPMEmail
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
	dim strCustomerImpact
	dim strDeliverableRootID
	dim strDeliverableRootName
	dim strDeliverableManagerID
	Dim strNotifyEmail
	
	strNotifyEmail = ""
	
	strChangeTypeCat = ""
	strCustomerImpact = "None"
	blnSendDisapprovedEmail = false
	
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
	case "7"
	    strTypeName = "ECR"
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
	

	Set p = cm.CreateParameter("@UserName", adVarchar, adParamInput, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", adVarchar, adParamInput, 30)
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
	end if
	rs.Close


	strDescription = request("txtDescription")
	if trim(request("txtType")) = "4" then
		strDescription = StripHTMLTag(strDescription)
	end if

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

	strProducts = request("lstProducts") & ","

	dim NewStatus
	NewStatus = clng(request("cboStatus"))
	
	if request("txtID") = "" then 'Adding
		if strproducts <> "," then 'Just to make sure we don't get into an infinite loop
			
			cn.BeginTrans
			'loopcount = 0
				strChangeTypeCat=""
				strBusiness = ""
				strRegions = ""
    			
				'strProductID = left(strProducts,instr(strProducts,",")-1)
				'strProducts = mid(strProducts,instr(strProducts,",")+1)
				strProductId = Request.Form("hidProdId")

				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetProductVersionName"

				Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
				p.Value = strProductId
				cm.Parameters.Append p

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				strProductStatusID = rs("ProductStatusID") & ""
				strProductName = rs("Name") & ""

				rs.Close		

				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "usp_GetSBA"

                Response.Write strProductId & "<br>"

				Set p = cm.CreateParameter("@p_ProductVersionId", adInteger, adParamInput)
				p.Value = strProductId
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing
                
                If Not rs.EOF Then
				OwnerID = rs("ID")
				OwnerName = rs("Name")
				rs.Close
		        Else
                    Response.Write "<h2>Bom Analyst Not Found</h2><p>Save ECR Failed</p>"
                    Response.End
                End If
			
				set cm = server.CreateObject("ADODB.command")
		
				cm.ActiveConnection = cn
				'cm.CommandText = "spAddDeliverableActionWeb2"
				cm.CommandText = "spAddDeliverableActionEcr"
				cm.CommandType =  &H0004
				cm.NamedParameters = True

				set p =  cm.CreateParameter("@ProductID", adInteger, adParamInput)
				p.value = strproductid
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Type", adTinyInt, adParamInput)
				p.Value = clng(strType)
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Status", adTinyInt, adParamInput)
				p.Value = 1
				cm.Parameters.Append p
		
				Set p = cm.CreateParameter("@Submitter", adVarchar, adParamInput, 50)
				p.Value = left(currentusername,50)
				cm.Parameters.Append p
	
				Set p = cm.CreateParameter("@OwnerID", adInteger, adParamInput)
				p.Value = OwnerID
				cm.Parameters.Append p

    			Set p = cm.CreateParameter("@Notify", adVarchar, adParamInput, 255)
				p.Value = left(request("txtnotify"),255)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Summary", adVarchar, adParamInput, 120)
				p.Value = left(request("txtSummary"),120)
				cm.Parameters.Append p
            
				Set p = cm.CreateParameter("@Description", adVarchar, adParamInput, 8000)
				p.Value = left(request("txtAttachment"),8000)
				cm.Parameters.Append p
           
				Set p = cm.CreateParameter("@Details", adVarchar, adParamInput, 8000)
   				p.Value = left(request("txtDetails"),8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@LastUpdUser", adVarchar, adParamInput, 200)
				p.Value = CurrentUserName
				cm.Parameters.Append p
	    
				Set p = cm.CreateParameter("@Initiator", adInteger, adParamInput)
				p.Value = Request("selInitiator")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpareKitPn", adVarchar, adParamInput, 500)
				p.Value = Request("txtSpareKitNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SubAssemblyPn", adVarchar, adParamInput, 500)
				p.Value = Request("txtSaNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@InventoryDisposition", adInteger, adParamInput)
				p.Value = Request("selInventorDisposition")
				cm.Parameters.Append p
			
				Set p = cm.CreateParameter("@QSpecSubmitted", adDate, adParamInput)
				If IsDate(Request("txtQSpecDt")) Then
				    p.Value = Request("txtQSpecDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@CompEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtCompEcoDt")) Then
				    p.Value = Request("txtCompEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@CompEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtCompEcoNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SaEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtSaEcoDt")) Then
				    p.Value = Request("txtSaEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SaEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtSaEcoNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpsEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtSpsEcoDt")) Then
				    p.Value = Request("txtSpsEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpsEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtSpsEcoNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@NewID", adInteger,  &H0002)
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
				
'
' Add GPLM as approver
'				
                If (not Errors) And cint(strType) = 7 then
    				rs.open "usp_GetGPLM " & strproductid,cn,adOpenForwardOnly
				
				    do while not rs.eof
					    if InStr(strApproverIDs, rs("ID")) = 0 then
						    strApproverIDs = strApproverIDs & "," & rs("ID")
						    strApproverEmails = strApproverEmails & rs("Email") & ";"
						    strApprovers = strApprovers & rs("Name") & " - Requested<BR>"  
					    end if
					    rs.movenext	
				    loop
				    rs.close
	
					if len(strApproverIDs) > 0 then
						strApproverIDs = mid(strApproverIDs,2)
					end if
					
					ApproverArray = split(strApproverIDs,",")
					for each Approver in ApproverArray
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spAddApprover"
		

						Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
						p.Value = DisplayedID
						cm.Parameters.Append p
	
						Set p = cm.CreateParameter("@ApproverID", adInteger, adParamInput)
						p.Value = clng(Approver)
						cm.Parameters.Append p

						cm.Execute RowsEffected
						Set cm=nothing

						
						if RowsEffected <> 1 then
							cn.RollbackTrans
							Errors = true
							exit for
						end if					
					next
				end if
				
								
'				if strGroupsAdded <> "" then
'					for i = lbound(GroupAddArray) to ubound(GroupAddArray)
'						set cm = server.CreateObject("ADODB.Command")
'						Set cm.ActiveConnection = cn
'						cm.CommandType = 4
'						cm.CommandText = "spAddActionGroup"
'		
'						Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
'						p.Value = clng( Displayedid)
'						cm.Parameters.Append p
'
'						Set p = cm.CreateParameter("@GroupID", adInteger, adParamInput)
'						p.Value = clng( GroupAddArray(i))
'						cm.Parameters.Append p
'	
'						cm.Execute rowseffected
'						Set cm=nothing
'							
'						if rowseffected <> 1 then
'							Errors= true
'							exit for
'						end if
'					next
'				end if				
				
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetEmployeeByName"

				Set p = cm.CreateParameter("@Name", adVarchar, adParamInput,80)
				p.Value = OwnerName
				cm.Parameters.Append p

				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				strPMEmail =  rs("Email") & ";"		
				rs.Close

'				if trim(strProductEmail) = "" then
'					strProductEmail = strPMEmail
'				end if

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
                case "7"
                    strTypeName = "Service ECR"
				case else
    				strTypeName = "Item"
				end select

			do while instr(strProducts,",")
				strProdID = left(strProducts,instr(strProducts,",")-1)
				strProducts = mid(strProducts,instr(strProducts,",")+1)

    				rs.open "usp_GetGPLM " & strProdID, cn, adOpenForwardOnly
				
				    do while not rs.eof
					    if InStr(strApproverIDs, rs("ID")) = 0 then
						    strPMEmail = strPMEmail & rs("Email") & ";"
						    response.Write rs("Email") & ";"
					    end if
					    rs.movenext	
				    loop
				    rs.close

			loop

				if trim(strProductEmail) = "" then
					strProductEmail = strPMEmail
				end if

				if strPMEmail <> "" then
					
					if strChangeTypeCat <> "" then
						strChangeTypeCat = " (" & mid(strChangeTypeCat,2) & ")"
					end if
			        'strBody = strBody & "<b>TARGET CLOSURE:</b> " & request("txtTarget") & "<BR>"
					strBody = "<span style=""font-weight:bold"">Internal Links:</span><br /><font face=Arial size=2><a href=""http://" & Application("Excalibur_ServerName") & "/mobilese/today/action.asp?Type=" & strtype & "&id=" & DisplayedID & """>Open this " & strtypename & "</a><br>"
					strBody = strBody & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur.asp"">Open Pulsar Today Page</a></font><br><br>"
					strBody = strBody & "<span style=""font-weight:bold"">External Links:</span><br /><font face=Arial size=2><a href=""https://prp.atlanta.hp.com/excalibur/mobilese/today/action.asp?Type=" & strtype & "&id=" & DisplayedID & """>Open this " & strtypename & "</a><br>"
					strBody = strBody & "<a href=""https://prp.atlanta.hp.com/excalibur/excalibur.asp"">Open Pulsar Today Page</a></font><br><br>"
					strBody = strbody & "<font face=Arial size=2>"
					strBody = strBody & "<b>NUMBER:</B> " & DisplayedID & "<BR>"
					strBody = strBody & "<b>TYPE:</b> " & strtypename & "<BR>"
					strBody = strBody & "<b>SUBMITTER:</b> " & currentusername & "<BR>"
					strBody = strBody & "<b>PROGRAM:</b> " & strproductname & "<BR>"
					if strtype = "5" then
						strBody = strBody & "<b>ISSUE/ACCOMPLISHMENT:</B> " & replace(request("txtSummary"),"""","&QUOT;") & "<BR>"
					else
						strBody = strBody & "<b>SUMMARY:</B> " & replace(request("txtSummary"),"""","&QUOT;") & "<BR>"
					end if
					if strDcrAutoOpen >= 2 and (cint(strType) = 3 Or Cint(strType) = 7) then
						strBody = strBody & "<b>STATUS:</b> Investigating<BR>"
					else
						strBody = strBody & "<b>STATUS:</b> Open<BR>"
					end if
					strBody = strBody & "<b>OWNER:</b> " & ownername & "<BR>"
			        if strBusiness <> "" then
				        strBody = strBody & "<b>BUSINESS:</b> " & Mid(strBusiness ,3) & "<BR>"
			        end if
			        if strRegions <> "" Then
			            strBody = strBody & "<b>REGIONS:</b> " & Mid(strRegions, 3) & "<br>"
			        end if

					if request("txtDescription") <> "" then
       					if trim(strType) = "5" then
       						strDescription = replace(request("txtDescription") & "",vbcrlf,"<BR>")
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
							strBody = strBody & "<b>DESCRIPTION:</b> " & "<BR>" & replace(request("txtDescription") & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						end if
					end if

                    If trim(strType) = "3" Then
                        if bBios Then
                            sBiosNotes = replace(sBiosNotes, vbcrlf, "<BR>")
                            strBody = strBody & "<b>DETAILS:</b> " & "<BR>" & sBiosNotes & "<BR>" & sDistribution & "<BR>" & strOSList & "<BR>" & strLanguageList & "<BR>" & "<BR>"
                        Else
                            strBody = strBody & "<b>DETAILS:</b> " & "<BR>" & sDistribution & "<BR>" & strOSList & "<BR>" & strLanguageList & "<BR>" & "<BR>"
                        End If
                    End If
                    
                    if request("txtJustification") <> "" then
						if trim(strtype) = "5" then
							strBody = strBody & "<b>ROOT CAUSE:</b> " & "<BR>" & replace(request("txtJustification") & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						else
							strBody = strBody & "<b>JUSTIFICATION:</b> " & "<BR>" & replace(request("txtJustification") & "",vbcrlf,"<BR>") & "<BR>" & "<BR>"
						end if
					end if
					
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
						else
							strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & replace(request("txtActions"),vbcrlf,"<BR>") & "</font>" & "<BR>" & "<BR>"
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
					end if 
					
       		        if trim(strType) = "5" then
						strBody = strBody & "<b>METRIC IMPACTED:</b> " & left(request("cboMetricImpact"),35)	 & "<BR>"
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
					end if        
					
					If bZsrpReadyRequired Then
					    strBody = strBody & "<b>ZSRP READY TARGET: </b> " & request("txtZsrpReadyTargetDt") & "<br>"
					    strBody = strBody & "<b>ZSRP READY ACTUAL: </b> " & request("txtZsrpReadyActualDt") & "<br><br><br>"
                    End If					    
					
                    If sBios = "" Then
                        strBody = strBody & "<b>CUSTOMER IMPACT: </b>" & strCustomerImpact & "<BR>"
                    End If
        			if trim(request("txtAvailDate")) <> "" then
		        		strBody = strBody & "<font color=red><b>SAMPLES AVAILABLE: </b> " & request("txtAvailDate") & "</font><BR>"
			        end if

					strBody = strBody & "<b>NOTIFY ON CLOSURE: </b> " & request("txtNotify") & "<BR>"
					
					if strGroupList <> "" then
						strBody = strBody & "<b>SUB-GROUP OWNERS:</b> " & strGroupList & "<BR><BR>"
					end if
					
					if strApprovers <> "" then
						strBody = strBody & "<b>APPROVALS:</b><font color=teal><BR>" & strApprovers & "</font><BR><BR>"
					end if
					
					if clng(strDcrAutoOpen) >= 2 and cint(strType) = 3 then
						if strApprovers <> "" then
						    If strDcrAutoOpen = 2 Then
							    strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The System Team members have been added as approvers for this DCR.</font></b><BR><BR>" & strBody
							Else 
							    strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".<BR><BR>The specified people have been added as approvers for this DCR.</font></b><BR><BR>" & strBody
							End If
						else
							strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added, set to investigating, and assigned to " & OwnerName & ".</font></b><BR><BR>" & strBody
						end if

					else
						strBody = "<font face=Arial size=2 color=red><b>This " & strTypename & " has been added and assigned to " & OwnerName & ".</font></b><BR><BR>" & strBody
					end if
					strBody = strBody & "<br><font size=1 color=red face=verdana>HP Restricted</font>"
					
					Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
					'Set oMessage.Configuration = Application("CDO_Config")

					oMessage.From = CurrentUserEmail
					if strProgramMail = "1" then
						if strApproverEmails <> "" then
							oMessage.To=strProductEmail & ";" & strApproverEmails
						else
							oMessage.To=strProductEmail
						end if
						oMessage.Subject = strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
					else
						oMessage.To= CurrentUserEmail 
						oMessage.Subject = "TEST MAIL: " & strtypename & strChangeTypeCat & " " & DisplayedID &  " (Added) : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
					end if
				
					if trim(strProgramMail) = "1" then
						oMessage.HTMLBody = strBody
					else
						oMessage.HTMLBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strProductEmail & ";" & strApproverEmails & "</font><BR><BR>" & strBody
					end if				
					''oMessage.Bcc = "kenneth.berntsen@hp.com"
					oMessage.Send 
					Set oMessage = Nothing 
				
				end if				
			
			if Errors then
				cn.RollbackTrans
			else
				cn.CommitTrans
			end if
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
'			if request("txtFormType") = "Change" then
				Response.Write "<label style=""Display:none"" ID=Results><h3><font face=verdana>" & strTypeName & " submitted.  Thank you for your input.</font></h3>"
				'response.write "<a href=""changeform.asp"" >Add Another Change Request</A><BR>"
				'Response.Write "<a href=""default.asp"" >Done Adding Change Requests</A><BR><BR>"
				Response.Write "<font face=verdana size=2><a href=""javascript:window.print();"" >Print This Window</A> | "
				Response.Write "<a href=""javascript:window.returnValue = 1;window.parent.close();"" >Close This Window</A><BR><BR></font>"
				Response.Write "<font size=2 face=verdana><b>" & request("txtSummary") & "</b></br></font><br>"
				Response.Write "<TABLE borderColor=tan cellSpacing=1 cellPadding=1 width=400 bgColor=wheat border=1><TR><TD width=270><FONT face=verdana size=2 color=black><b>Product</b></FONT></TD><TD width=200><FONT face=verdana size=2 color=black><b>ID Number</b></FONT></TD><TD width=270><FONT face=verdana size=2 color=black><b>Owner Assigned</b></FONT></TD></TR>"
				Response.Write strIDOutput & "</TABLE></label>"

'			else
'				Response.Write "<h3>Issue/Risk submitted.  Thank you for your input.</h3>"
'				response.write "<a href=""programinput.asp?Type=Risk"" >Add Another Issue/Risk</A><BR>"
'				Response.Write "<a href=""default.asp"" >Done Adding Issues/Risks</A><BR>"
'				Response.Write "<BR><TABLE borderColor=skyblue cellSpacing=1 cellPadding=1 width=400 bgColor=#e6f7ff border=1><TR><TD width=200 bgColor=#006697><FONT color=white>Product</FONT></TD><TD bgColor=#006697><FONT color=white>Issue/Risk Number</FONT></TD></TR>"
'				Response.Write strIDOutput & "</TABLE>"
'			end if
		end if
	else 'Editing
			'Pull old record
			strProductID = left(strproducts,instr(strproducts,",")-1)
			strproducts = mid(strproducts,instr(strproducts,",")+1)

			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetActionProperties"
			

			Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
			p.Value = request("txtID")
			cm.Parameters.Append p
	

			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			'rs.Open "spGetActionProperties " & request("txtID"),cn,adOpenForwardOnly
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
			end if
			rs.Close
	
	        If oldZsrpRequired <> "" Then 
	            oldZsrpRequired = CBool(oldZsrpRequired)
	        Else
	            oldZsrpRequired = false
	        End If
	
			if request("cboApproverStatus") = "2" AND request("commentsonly") = "" then 'Approved
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spVerifyAutoApprove"
		
				Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

				'rs.Open "spVerifyAutoApprove " & request("txtID"), cn,adOpenForwardOnly
				if not(rs.EOF and rs.BOF) then
					if rs("Verified") = 1 then	
							'NewStatus = 4
					end if
				end if
				rs.Close
			elseif request("cboApproverStatus") = "3" AND request("commentsonly") = "" then 'Disapproved
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spVerifyAutoApprove"
		
				Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
				p.Value = request("txtID")
				cm.Parameters.Append p
	
				rs.CursorType = adOpenForwardOnly
				rs.LockType=AdLockReadOnly
				Set rs = cm.Execute 
				Set cm=nothing

'				rs.Open "spVerifyAutoApprove " & request("txtID"), cn,adOpenForwardOnly
				if not(rs.EOF and rs.BOF) then
					if rs("DisapprovedCount") = 0 then	
						blnSendDisapprovedEmail = true
					end if
				end if
				rs.Close
			end if
		
			if true then 'oldOnStatus <> replace(request("chkReports"),"on","1") or trim(oldResolution) <> trim(request("txtResolution")) or trim(oldAction) <> trim(request("txtActions")) or trim(oldDescription) <> trim(left(request("txtDescription"),8000)) or trim(oldNotify) <> trim(left(request("txtnotify"),255)) or trim(oldTargetDate) <> trim(request("txtTarget")) or trim(oldCoreTeam) <> trim(request("cboCoreTeam")) or trim(oldOwner) <> trim(request("cboOwner")) or trim(oldProduct) <> trim(strproductid)  or trim(oldStatus) <> trim(cstr(NewStatus)) or trim(oldSummary) <> trim(left(request("txtSummary"),80)) then 
				'Save updates
				set cm = server.CreateObject("ADODB.command")
		
				cm.ActiveConnection = cn
				cm.CommandText = "spUpdateDeliverableActionEcr"
				cm.CommandType =  &H0004
	
				cn.BeginTrans
				
				Set p = cm.CreateParameter("@ID", adInteger,  adParamInput)
				p.value = clng( request("txtID"))
				cm.Parameters.Append p

				set p =  cm.CreateParameter("@ProductID", adInteger, adParamInput)
				p.value = strproductid
				cm.Parameters.Append p
	            
				Set p = cm.CreateParameter("@Status", adTinyInt, adParamInput)
				p.Value = CINT(NewStatus)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@OwnerID", adInteger, adParamInput)
				p.Value = request("cboOwner")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Notify", adVarchar, adParamInput, 255)
				p.Value = left(request("txtnotify"),255)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Summary", adVarchar, adParamInput, 120)
				p.Value = left(request("txtSummary"),120)
				cm.Parameters.Append p
            
				Set p = cm.CreateParameter("@Description", adVarchar, adParamInput, 8000)
				p.Value = left(request("txtAttachment"),8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Details", adVarchar, adParamInput, 8000)
				p.Value = left(request("txtDetails"),8000)
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@LastUpdUser", adVarchar, adParamInput, 200)
				p.Value = CurrentUserName
				cm.Parameters.Append p
	    
				Set p = cm.CreateParameter("@Initiator", adInteger, adParamInput)
				p.Value = Request("selInitiator")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpareKitPn", adVarchar, adParamInput, 500)
				p.Value = Request("txtSpareKitNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SubAssemblyPn", adVarchar, adParamInput, 500)
				p.Value = Request("txtSaNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@InventoryDisposition", adInteger, adParamInput)
				p.Value = Request("selInventorDisposition")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@QSpecSubmitted", adDate, adParamInput)
				If IsDate(Request("txtQSpecDt")) Then
				    p.Value = Request("txtQSpecDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p
				
				Set p = cm.CreateParameter("@CompEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtCompEcoDt")) Then
				    p.Value = Request("txtCompEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@CompEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtCompEcoNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SaEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtSaEcoDt")) Then
				    p.Value = Request("txtSaEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SaEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtSaEcoNo")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpsEcoSubmitted", adDate, adParamInput)
				If IsDate(Request("txtSpsEcoDt")) Then
				    p.Value = Request("txtSpsEcoDt")
				Else
				    p.Value = Null
				End If
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@SpsEcoNo", adVarchar, adParamInput, 50)
				p.Value = Request("txtSpsEcoNo")
				cm.Parameters.Append p

				cm.Execute RowsEffected
				
				set cm = nothing
				
				if cn.Errors.Count > 0 then
					cn.RollbackTrans
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
                    response.Write "<br>" & strApproverList & "<br>" 
					do while instr(strApproverList,",")> 0
						Response.Write request("txtID") & "," & left(strApproverList,instr(strApproverList,",")-1) & "<BR>"
						
						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spAddApprover"
		

						Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
						p.Value = request("txtID")
						cm.Parameters.Append p
	
						Set p = cm.CreateParameter("@ApproverID", adInteger, adParamInput)
						p.Value = left(strApproverList,instr(strApproverList,",")-1)
						cm.Parameters.Append p

						cm.Execute RowsEffected
						Set cm=nothing

						if RowsEffected <> 1 then
							cn.RollbackTrans
							Errors = true
							exit do
						end if	

						set cm = server.CreateObject("ADODB.Command")
						Set cm.ActiveConnection = cn
						cm.CommandType = 4
						cm.CommandText = "spGetEmployeeByID"

						Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
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
					
					'Update Approver Status
					if request("txtSaveApproval") <> "" and request("txtSaveApproval") <> "0" then
						set cm = server.CreateObject("ADODB.command")
		
						cm.ActiveConnection = cn
						cm.CommandText = "spUpdateApproval"
						cm.CommandType =  &H0004
	
						Set p = cm.CreateParameter("@ApprovalID", adInteger,  adParamInput)
						p.value = clng( request("txtSaveApproval"))
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@Status", adTinyInt,  adParamInput)
						If request("cboApproverStatus") <> "" Then
						    p.value = clng(request("cboApproverStatus"))
						end if
						cm.Parameters.Append p

						Set p = cm.CreateParameter("@Comments", adVarchar, adParamInput, 300)
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
		
							Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
							p.Value = clng( request("txtID"))
							cm.Parameters.Append p

							Set p = cm.CreateParameter("@GroupID", adInteger, adParamInput)
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
		
							Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
							p.Value = clng( request("txtID"))
							cm.Parameters.Append p

							Set p = cm.CreateParameter("@GroupID", adInteger, adParamInput)
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
		
						Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
						p.Value = clng( request("txtID"))
						cm.Parameters.Append p
	
						cm.Execute 
						Set cm=nothing

						'cn.Execute "spSetApprovalList " & clng( request("txtID"))
					end if
					
					if not errors then
						cn.CommitTrans
					end if
				end if
		end if 'Save				
		Response.Write "Done"
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
		dim strPM
		dim strDivision
		Response.Write request("cboOwner") & "<BR>"
		
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetEmployeeByID"
		

		Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
		p.Value = request("cboOwner")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		
		'rs.Open "spGetEmployeeByID " & request("cboOwner"),cn,adOpenForwardOnly
		strNewOwnerEmail = rs("Email") & ""		
		rs.Close

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetEmployeeByName"
		

		Set p = cm.CreateParameter("@Name", adVarchar, adParamInput,80)
		p.Value = left(request("txtSubmitter"),80)
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

'		rs.Open "spGetEmployeeByName '" & request("txtSubmitter") & "'",cn,adOpenForwardOnly
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
		

		Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
		p.Value = strproductid
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

'		rs.Open "spGetProductVersion  " & strproductid ,cn,adOpenForwardOnly
		strProductEmail = rs("Distribution") & ""		
		strPM = rs("PMName") & ""
		strDivision = rs("Division") & ""
		rs.Close
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetEmployeeByName"
		

		Set p = cm.CreateParameter("@ID", adVarchar, adParamInput,80)
		p.Value = strPM
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spGetEmployeeByName '" & strPM & "'",cn,adOpenForwardOnly
		strPMEmail = rs("Email") & ""		
		rs.Close
		DisplayedID = request("txtID")


		
		
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
	
		strTO = ""
'		if request("chkCommodityChange") = "on" then
'			if request("txtID") <> "" then 'Only notify them when editing
'				if trim(request("txtCommodityManagerID")) = "" or trim(request("txtCommodityManagerID")) = "0" then
'					strTO = strTO & "lester.williams@hp.com;fred.massoudian@hp.com;t.kutyba@hp.com;j.wong@hp.com;franklin.lee@hp.com;jim.tso@hp.com;ernest.chang@hp.com;kennon.skyvara@hp.com;steve.bachmeier@hp.com;"
'				else
'					rs.open "spGetEmployeeByID " & clng(request("txtCommodityManagerID")) ,cn,adOpenForwardOnly
'					if rs.eof and rs.bof then
'						strTO = strTO & "lester.williams@hp.com;fred.massoudian@hp.com;t.kutyba@hp.com;j.wong@hp.com;franklin.lee@hp.com;jim.tso@hp.com;ernest.chang@hp.com;kennon.skyvara@hp.com;steve.bachmeier@hp.com;"
'					else
'						strTO = strTO & rs("Email") & ";"
'					end if
'					rs.close
'				end if
'			end if
'		end if


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
		
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetAction4Mail"
		

		Set p = cm.CreateParameter("@ID", adInteger, adParamInput)
		p.Value = DisplayedID
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

'	    rs.Open "spGetAction4Mail " & DisplayedID, cn, adOpenForwardOnly

		
		if rs.EOF and rs.BOF then
			strBody = "<span style=""font-weight:bold"">Internal Link:</span><br /><a href=""http://" & Application("Excalibur_ServerName") & "/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br><br>"
			strBody = strBody & "<span style=""font-weight:bold"">External Link:</span><br /><a href=""https://prp.atlanta.hp.com/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br><br>"
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
			Case "7"
			    strType = "Service ECR"
			End Select
			
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
			end if
			
			strBody = "<span style=""font-weight:bold"">Internal Links:</span><br /><font face=Arial size=2><a href=""http://" & Application("Excalibur_ServerName") & "/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br>"
			strBody = strBody & "<a href=""http://" & Application("Excalibur_ServerName") & "/excalibur.asp"">Open Pulsar Today Page</a><br><br></font>"
			strBody = strBody & "<span style=""font-weight:bold"">External Links:</span><br /><font face=Arial size=2><a href=""https://prp.atlanta.hp.com/excalibur/mobilese/today/action.asp?Type=" & typeid & "&id=" & DisplayedID & """>Open this " & strtype & "</a><br>"
			strBody = strBody & "<a href=""https://prp.atlanta.hp.com/excalibur/excalibur.asp"">Open Pulsar Today Page</a><br><br></font>"
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
			
			If bScr Then
    			strBody = strBody & "<b>DELIVERABLE ROOT:</b> " & rs("DeliverableRoot") & "<BR>"
			End If
			
			strBody = strBody & "<b>TYPE:</b> " & strtype & "<BR>"
			if trim(TypeID) = "5" then
				strBody = strBody & "<b>ISSUE/ACCOMPLISHMENT:</B> " & replace(request("txtSummary"),"""","&QUOT;") & "<BR>"
			else
				strBody = strBody & "<b>SUMMARY:</B> " & replace(request("txtSummary"),"""","&QUOT;") & "<BR>"
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
				strBody = strBody & "<b>BUSINESS:</b> " & Mid(strBusiness, 3) & "<BR>"
			end if
			if strRegions <> "" Then
			    strBody = strBody & "<b>REGIONS:</b> " & Mid(strRegions, 3) & "<br>"
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
			    strBody = strBody & "<b>DETAILS:</b> "  & "<BR>" & Replace(request("txtDetails"), VbCrLf, "<BR>") & "<BR>"
			End If

			if rs("Justification") & "" <> "" then
				if trim(TypeID) = "5" then
					strBody = strBody & "<b>ROOT CAUSE:</b><BR>" & replace(rs("Justification"),vbcrlf,"<BR>") & "<BR><BR>"
				else	
					strBody = strBody & "<b>JUSTIFICATION:</b><BR>" & replace(rs("Justification"),vbcrlf,"<BR>") & "<BR><BR>"
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
                End If					    

				if trim(TypeID) = "5" then
					strBody = strBody & "<font color=red>" & strActions & "</font>" & "<BR>" & "<BR>"
				else
					strBody = strBody & "<b>ACTIONS NEEDED:</b> <font color=red>" & "<BR>" & strActions & "</font>" & "<BR>" & "<BR>"
				end if
				strBody = strBody & "<b>RESOLUTION:</b> " & "<BR>" & replace(rs("Resolution") & "",vbcrlf,"<BR>") & "<BR>"
			End If

			if trim(request("txtAvailDate")) <> "" then
				strBody = strBody & "<font color=red><b>SAMPLES AVAILABLE:</b> " & request("txtAvailDate") & "</font><BR>"
			end if
			
'			if trim(request("txtAvailNote")) <> "" then
'				strBody = strBody & "<font color=red><b>AVAILABILITY NOTE:</b> " & request("txtAvailNote") & "</font><BR>"
'			end if
    
    


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

			if trim(TypeID) = "5"  and trim(rs("AvailableNotes") & "") <> "" then
				strBody = strBody & "<b>METRIC IMPACTED:</b> " & rs("AvailableNotes") & "<BR>"
			end if

'			strBody = strBody & "<b>CORE TEAM REP:</b> " & rs("CoreTeamRep") & "<BR>"
			strCoreTeam = rs("CoreTeamRep")
'			If strtype <> "Status Note" Then
'		        strBody = strBody & "<b>TARGET CLOSURE:</b> " & rs("TargetDate") & "<BR>"
'			End If
			if (not isnull(rs("BTODate"))) and (rs("Distribution") = "BTO" or rs("Distribution") = "BOTH") then
		        strBody = strBody & "<b>BTO-IMPLEMENT BY:</b> " & rs("BTODate") & "<BR>"
			end if
			if (not isnull(rs("CTODate"))) and (rs("Distribution") = "CTO" or rs("Distribution") = "BOTH") then
		        strBody = strBody & "<b>CTO-IMPLEMENT BY:</b> " & rs("CTODate") & "<BR>"
			end if

			If strtype <> "Status Note" Then
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
		set rs = nothing
		
		cn.Close
		set cn = nothing
		
'		strBody = replace(strBody,CurrentUserName, "<B><font color=blue size=2 face=verdana>" & currentusername & "</font></b>")
		
		'----------
		
		dim strCC
		
		
		Response.Write "<font size=2 face =verdana>"
		Response.Write "<font face=verdana size=3><b>Notifications Sent:</b></font>"
		if NewStatus = 3 and oldStatus <> "3" then 'More input needed
			strTo = strTo & strNewOwnerEmail & ";" & strSubmitterEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if
			Response.Write "<BR>More Information Needed: " & strTO & "<BR>"		
			notifycount = notifycount + 1
            strSubject = strSubject & "(Need more info)"
			strBody = "<font face=Arial size=2 color=red><b>Please provide more information for this " & strtype & "</font></b><BR><BR>" & strBody 
		end if

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
		    strSubject = strSubject & "(Disapproved by Approver)"
			strBody = "<font face=Arial size=2 color=red><b>This item has been disapproved by an approver. Further disapprovals will not generate and email notification.</font></b><BR><BR>" & strBody 
		end if


		if strApproverEmails <> "" then 'Approvers Added
			strTo = strTO & strApproverEmails 
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>New Approvers: " & strTO & "<BR>"		
			notifycount = notifycount + 1
			if newStatus = 6 and oldStatus <> "6" then
				if strCoreTeam = "Sustaining System Team" then
					strCC = strProductEmail & ";" & CCTMailList& ";"
				else
					strCC = strProductEmail & ";"
				end if			
			end if
			strSubject = strSubject & "(Approval Requested)"
			strBody = "<font face=Arial size=2 color=red><b>You have been added to the Approval List</font></b><BR><BR>" & strBody 
		end if

		if newStatus = 6 and oldStatus <> "6" then 'Set to Investigating
			if strApproverEmails = "" then
				If trim(strDivision) = "1" And LCase(Request.Form("hidAddDCRNotificationList")) = "true" Then
					strTo = strTO & strProductEmail & ";NotebookDCRNotification@hp.com;"
				Else
					strTo = strTO & strProductEmail & ";"				
				End If
				strTo = replace(StrTo,CurrentUserEmail & ";","")
				if strCoreTeam = "Sustaining System Team" then
					if strTO = "" then
						strTO = CCTMailList & ";"
					else
						strTO = strTo & CCTMailList & ";"
					end if
				end if
			else
				if trim(strDivision) = "1" and LCase(Request.Form("hidAddDCRNotificationList")) = "true" then
					strTo = "NotebookDCRNotification@hp.com;"			
				end if
			end if
			Response.Write "<BR>Investigating: " & strTO & "<BR>"		
			notifycount = notifycount + 1
			strSubject = strSubject & "(Investigating)"
			strBody = "<font face=Arial size=2 color=red><b>Status set to Investigating</font></b><BR><BR>" & strBody 
		end if

		if trim(oldOwner) <> trim(strNewOwner) then 'New owner assigned
			strTo = strTO & strNewOwnerEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>New Owner Assignment: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(Reassigned)"
			strBody = "<font face=Arial size=2 color=red><b>This " & strType & " has been assigned to " & strOwnerName & ".</font></b><BR><BR>" & strBody 
		end if
		
		if oldZsrpRequired And NOT bZsrpReadyRequired Then
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>ZSRP Requirement Removed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Requirement Removed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Requirement Removed.</font></b><BR><BR>" & strBody 
		end if
		
		if (trim(oldZsrpTargetDt) <> trim(request("txtZsrpReadyTargetDt"))) And bZsrpReadyRequired Then ' ZSRP Ready Target Date Changed
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>ZSRP Actual Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Target Date Changed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Target Date Updated.</font></b><BR><BR>" & strBody 
		end if
		
		if (trim(oldZsrpActualDt) <> trim(request("txtZsrpReadyActualDt"))) And bZsrpReadyRequired Then ' ZSRP Ready Actual Date Changed
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>ZSRP Actual Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(ZSRP Actual Date Changed)"
			strBody = "<font face=Arial size=2 color=red><b>ZSRP Actual Date Updated.</font></b><BR><BR>" & strBody 
		end if
		
		if (trim(oldTargetDate) <> trim(request("txtTarget"))) and request("txtID") <> "" then 'TargetDate Updated
			strTo = strTO & ";" & strProductEmail & ";" & strPMEmail & ";"
			strTo = replace(StrTo,CurrentUserEmail & ";","")
			if strCoreTeam = "Sustaining System Team" then
				if strTO = "" then
					strTO = CCTMailList & ";"
				else
					strTO = strTo & CCTMailList & ";"
				end if
			end if			
			Response.Write "<BR>Target Date Changed: " & strTO &  "<BR>"
			notifycount = notifycount + 1
			strSubject = strSubject & "(Target Changed)"
			strBody = "<font face=Arial size=2 color=red><b>Target Date Updated.</font></b><BR><BR>" & strBody 
		end if
		
		if (newStatus = 2 or NewStatus = 4 or NewStatus = 5) and (oldstatus <> trim(cstr(newStatus)) ) then
			Response.Write strtype & ":" & strProductEmail & "<BR>"
			if TypeID = "3" then
				strTO = strTO & strNewOwnerEmail  & ";" & strProductEmail & ";" & strSubmitterEmail & ";"
				if request("txtnotify") <> "" then
					if strTO = "" then
						strTO = request("txtnotify") & ";" 
					else
						strTO = strTO & request("txtnotify") & ";" 
					end if
				end if

				if trim(strDivision) = "1" and LCase(Request.Form("hidAddDCRNotificationList")) = "true" then
					if strTO = "" then
						strTO = "NotebookDCRNotification@hp.com;"
					else
						strTO = strTO & "NotebookDCRNotification@hp.com;"
					end if
				end if
				Response.Write "<BR>Item Closed: " & strTO & "<BR>"
			else
				strTO = strTO & strNewOwnerEmail  & ";" & strSubmitterEmail  & ";"
				if request("txtnotify") <> "" then
					if strTO = "" then
						strTO = request("txtnotify") & ";"
					else
						strTO = strTO & request("txtnotify") & ";"
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
			
			notifycount = notifycount + 1
			strSubject = strSubject & "(" & strStatus & ")"
			strBody = "<font face=Arial size=2 color=red><b>Status changed to " & strStatus & ".</font></b><BR><BR>" & strBody 
		end if

'			if TypeID = "3" and trim(strDivision) = "1" then
'				strTO = "NotebookDCRNotification@hp.com;"
'				notifycount = notifycount + 1
'				Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
'				'Set oMessage.Configuration = Application("CDO_Config")
'
'				oMessage.From = CurrentUserEmail
'				if strProgramMail = "1" then
'					oMessage.To=strTo
'					oMessage.Subject = "DCR Notification: " & strtype & " " & DisplayedID  & " (" & strStatus & ") : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
'				else
'					oMessage.To=CurrentUserEmail
'					oMessage.Subject = "TEST MAIL-Notebook List Notification: " & "DCR Notification: " &  strtype & " " & DisplayedID  &  " (" & strStatus & ") : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
'				end if	
'				
'				if strProgramMail = "1" then
'					oMessage.HTMLBody = "<font face=Arial size=2 color=red><b>FOR NOTIFICATION PURPOSES ONLY.<BR><BR>Status changed to " & strStatus & ".</font></b><BR><BR>" & strBody 
'				else
'					oMessage.HTMLBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strTo & "</font><BR><BR>" & "<font face=Arial size=2 color=red><b>FOR NOTIFICATION PURPOSES ONLY.<BR><BR>Status changed to " & strStatus & ".</font></b><BR><BR>" & strBody 
'				end if
'	
'				oMessage.Send 
'				Set oMessage = Nothing 
'			end if



		end if
		if notifycount = 0 then
			Response.Write "<BR>None."
		end if
		
		if notifycount > 0 And strTo <> "" then

	            if notifycount > 1 then
	                If instr(strSubject,"Approved") > 0 And instr(strSubject,"Disapproved") = 0 Then
	                    strSubject = "(Approved)"
	                elseif instr(strSubject, "ZSRP") > 0 Then
	                    strSubject = "(ZSRP Updates)"
	                else
	                    strSubject = "(Multiple Changes)"
	                end if
	            End If
	            
	            strSubject = strtype & " " & DisplayedID &  " " & strSubject & " : " & strProductName &  " : " & replace(request("txtSummary"),"""","'")
	            
	            if strProgramMail <> "1" then
	                strSubject = "TEST MAIL: " & strSubject
	                strBody = "<font face=Arial size=2><b>Email Not Enabled For This Product.<BR>Mail Would Have Been Sent To:</B> " & strTo & " " & strCC & "</font><BR><BR>" & strBody
	                strTo = CurrentUserEmail
	            end if
	            
				Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
	
				oMessage.From = CurrentUserEmail
				oMessage.To= strTo
				If Len(Trim(strCC)) > 0 Then
				    oMessage.Cc = strCC
				End If
				oMessage.Subject =  strSubject
				oMessage.HTMLBody = strBody

				'oMessage.Bcc = "kenneth.berntsen@hp.com"
				oMessage.Send 
				Set oMessage = Nothing 			

		end if
		
		Response.Write "</font>"	
	end if	

	Dim displayedIdValue
	if request("txtID") = "" then 'Adding
	    displayedIdValue = ""
	Else
	    displayedIdValue = DisplayedID
	End If
%>
    <input type="hidden" id="DisplayedID" name="DisplayedID" value="<%=displayedIdValue %>" />
</body>
</html>
