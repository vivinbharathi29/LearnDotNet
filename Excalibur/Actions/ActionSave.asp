<%@ Language=VBScript %>
<!-- #include file="../includes/emailwrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
	var OutArray = new Array();
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
		{
		    if (IsFromPulsarPlus()) {
		        if (GetQueryStringValue("component") == "apperror") {
		            window.parent.parent.parent.ApplicationErrorCallback(txtSuccess.value);
		        }
		        else
		        {
		            window.parent.parent.parent.ActionsCallback(txtSuccess.value);
		        }
		        ClosePulsarPlusPopup();
		    }
		    else {
		        //window.parent.close();
		        if (parent.window.parent.document.getElementById('modal_dialog')) {
		            parent.window.parent.modalDialog.cancel(true);
		        } else {
		            window.returnValue = 1;
		            window.close();
		        }
		    }
		    
			//OutArray[0]= txtOut1.value;
			//OutArray[1]= txtOut2.value;

			//window.returnValue = 1; //OutArray;
			//window.parent.opener=self;
			//window.parent.close();
			
		    //window.parent.opener='X';
		    //window.parent.open('','_parent','')
		    //window.parent.close();				
			}
		//else
		//	document.write ("<BR><font size=2 face=verdana>Unable to update item.</font>");
		}
	//else
	//	document.write ("<BR><font size=2 face=verdana>Unable to update item.</font>");
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

ID: <%=request("txtDisplayedID")%><BR>
UserID: <%=request("txtCurrentUserID")%><BR>


<%


	dim cn
	dim cm
	dim rs
	dim p
	dim i
	dim strNewID
	dim strSuccess
	dim strSubmitterName
    dim strSubmitterID
	dim strSubmitterEmail
	strSubmitterName = ""
    strSubmitterID=0
	strSubmitterEmail = ""

	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.CommandTimeout = 240
	cn.Open
	
	'Lookup Submitter Name
	rs.Open "spGetEmployeeByID " & clng(request("cboFrom")),cn,adOpenStatic
	strSubmitterName = rs("Name") & ""
    strSubmitterID = clng(request("cboFrom"))
	strSubmitterEmail = rs("Email") & ""
	rs.Close
	set rs = nothing
	
	'Save Action Item
	set cm = server.CreateObject("ADODB.Command")
		
	cm.ActiveConnection = cn
	cm.CommandType = &H0004
	if trim(request("txtDisplayedID"))="" or trim(request("txtDisplayedID"))="0" then
		cm.CommandText = "spAddDeliverableActionWeb"
	else
		strNewID = request("txtDisplayedID")
		
		cm.CommandText = "spUpdateDeliverableActionWeb"
		
		Set p = cm.CreateParameter("@ID",adInteger, &H0001)
		p.Value = clng(request("txtDisplayedID"))
		cm.Parameters.Append p

	end if

	Set p = cm.CreateParameter("@ProductID",adInteger, &H0001)
	p.Value = clng(request("cboProject"))
	cm.Parameters.Append p
			
	Set p = cm.CreateParameter("@Type",adTinyInt, &H0001)
	p.Value = clng(request("txtType"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Status",adTinyInt, &H0001)
	p.Value = clng(request("cboStatus"))
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@Submitter",adVarChar, &H0001,50)
	p.Value = left(strSubmitterName,50)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SubmitterID",adInteger, &H0001)
	p.Value = clng(strSubmitterID)
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@CategoryID",adInteger, &H0001)
	p.Value = 1
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OwnerID",adInteger, &H0001)
	p.Value = clng(request("cboOwner"))
	cm.Parameters.Append p
			
	Set p = cm.CreateParameter("@PreinstallOwner",adInteger, &H0001)
	if request("chkCopySubmitter") = "1" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
			
	Set p = cm.CreateParameter("@AffectsCustomers",adTinyInt, &H0001)
	if request("chkCopyMe") = "1" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@CoreTeamID",adInteger, &H0001)
	p.Value = 1
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@RoadmapID",adInteger, &H0001)
	p.Value = clng(request("cboRoadmap"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SponsorID",adInteger, &H0001)
	p.Value = clng(request("cboSponsor"))
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@Duration",adInteger, &H0001)
	if trim(request("txtDuration")) = "" then
		p.Value = null
	else
		p.Value = clng(request("txtDuration"))
	end if
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@DisplayOrder",adInteger, &H0001)
	if trim(request("txtOrder")) = "" then
		p.Value = 1
	else
		p.Value = clng(request("txtOrder"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TargetDate",adDBTimeStamp, &H0001)
	if isDate(request("txtTargetDate")) then
		p.Value = cdate(request("txtTargetDate"))
	else
		p.Value = null
	end if
	cm.Parameters.Append p
	
	if trim(request("txtDisplayedID"))<>"" and trim(request("txtDisplayedID"))<>"0" then
		Set p = cm.CreateParameter("@ECNDate",adDBTimeStamp, &H0001)
		p.Value = null
		cm.Parameters.Append p
	end if			

	Set p = cm.CreateParameter("@TestDate",adDBTimeStamp, &H0001)
	p.Value = null
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TestNote",adVarChar, &H0001,35)
	p.Value = ""
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@BTODate",adDBTimeStamp, &H0001)
	p.Value = null
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CTODate",adDBTimeStamp, &H0001)
	p.Value = null
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Distribution",adVarChar, &H0001,4)
	p.Value = "N/A"
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Notify",adVarChar, &H0001,255)
	p.Value = left(request("txtNotify"),255)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OnStatus",adTinyInt, &H0001)
	if request("chkReviewInput") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Priority",adTinyInt, &H0001)
	p.Value = clng(request("cboPriority"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Commercial",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Consumer",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SMB",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@APD",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CKK",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@EMEA",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@LA",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@NA",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@GCD",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@AddChange",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@MofityChange",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@RemoveChange",adBoolean, &H0001)
	p.Value = False
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ImageChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@CommodityChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DocChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@SKUChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ReqChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OtherChange",adBoolean, &H0001)
	p.Value = 0
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PendingImplementation",adBoolean, &H0001)
	if request("chkWorking") = "on" then
		p.Value = 1
	else
		p.Value = 0
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Summary",adVarChar, &H0001,120)
	p.Value = left(request("txtSummary"),120)
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@Description",adVarChar, &H0001,8000)
	p.Value = left(request("txtDetails"),8000)
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Justification",adVarChar, &H0001,8000)
	p.Value = ""
	cm.Parameters.Append p
					
	Set p = cm.CreateParameter("@Actions",adLongVarChar, &H0001,2147483647)
	p.Value = ""
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Resolution",adLongVarChar, &H0001,2147483647)
	p.Value = request("txtResolution")
	cm.Parameters.Append p
	
    Set p = cm.CreateParameter("@StatusNotes",adLongVarChar, &H0001,2147483647)
	p.Value = request("txtStatusNotes")
	cm.Parameters.Append p

	if trim(request("txtDisplayedID"))="" or trim(request("txtDisplayedID"))="0" then
		Set p = cm.CreateParameter("@NewID",adInteger, &H0002)
		cm.Parameters.Append p
	end if					
		
	cm.Execute rowsupdated

	if trim(request("txtDisplayedID"))=""  or trim(request("txtDisplayedID"))="0" then
		strNewID = cm("@NewID")
	end if

	Set cm = Nothing
	
	
	if cn.Errors.count > 0 then
		strSuccess = ""
	else
		strSuccess = "1"
	end if
	

	'Send Notifications if necessary
	if trim(request("cboOwner")) <> trim(request("tagOwner")) or trim(request("cboStatus")) <> trim(request("tagStatus")) then

		dim strBody
		dim strFrom
		dim strSubject

		dim strProduct
		dim strStatus
		dim strOwner
		dim strSummary
		dim strSubmitter
		dim strRoadmap
		dim strDetails
		dim strType
		dim strOwnerEmail
		dim strResolution

		strProduct = ""
		strStatus = ""
		strOwner = ""
		strSummary = ""
		strSubmitter = ""
		strDetails = ""
		strResolution = ""
		
		if trim(request("txtType")) = "1" then
			strType = "Issue"
		else
			strType = "Action"
		end if	
		
		set rs = server.CreateObject("ADODB.Recordset")
		rs.Open "spGetActionProperties4Print " & strNewID,cn,adOpenStatic
			strProduct = trim(rs("ProductName") & "")
			strStatus = trim(rs("Status") & "")
			if trim(strStatus) = "1" then
				strStatus = "Open"
			elseif trim(strStatus) = "2" then
				strStatus = "Closed"
			elseif trim(strStatus) = "3" then
				strStatus = "Blocked"
			end if
			strOwner = trim(rs("Owner") & "")
			strSummary = trim(rs("Summary") & "")
			strOwnerEmail = trim(rs("OwnerEmail") & "")
			strSubmitter = trim(rs("Submitter") & "")
			strDetails = trim(rs("Description") & "&nbsp;")
			strResolution = trim(rs("Resolution") & "")
		rs.Close
		set rs = nothing

		strFrom  = request("txtCurrentUserEmail")
		strBody = ""
		if trim(request("txtEmailNote")) <> "" then
			strBody = replace(request("txtEmailNote"),vbcrlf,"<BR>") & "<BR><BR>"
		end if
		if strResolution <> "" then
			strBody = strBody & "<font size=2 face=verdana><b>Resolution:</b> <BR><UL><LI>" & replace(strResolution,vbcrlf,"</li><li>") & "</li></ul><BR><BR></font>"
		end if
		strBody = strBody & "<TABLE WIDTH=""100%"" BORDER=1 CELLSPACING=0 CELLPADDING=2 borderColor=LightGrey bgColor=ivory>"
		strBody = strBody & "<TR><TD valign=top><b>Summary:</b></TD><TD colspan=3>" & strSummary & "</TD></TR>"
		strBody = strBody & "<TR><TD width=120><b>Product:</b></TD><TD width=120 nowrap>" & strProduct & "</TD><TD width=120><b>Owner:</b></TD><TD width=""100%"">" & strOwner & "</TD></TR>"
		strBody = strBody & "<TR><TD width=120><b>Status:</b></TD><TD width=120>" & strStatus & "</TD><TD width=120><b>Submitter:</b></TD><TD width=""100%"">" & strSubmitter & "</TD></TR>"
		strBody = strBody & "<TR><TD width=120><b>ID:</b></TD><TD width=120>" & strNewID & "</TD><TD width=120><b>Links:</b></TD><TD nowrap width=""100%"">" & "<a href=""http://16.81.19.70/Actions/Action.asp?ID=" & strNewID & "&Working=0&Type=" & request("txtType") & "&ProdID=" & request("cboProject") & """>Open This " & strType & "</a>" & "&nbsp;|&nbsp;" & "<a href=""http://16.81.19.70/Excalibur.asp"">Open Pulsar</a>" & "</TD></TR>"
		strBody = strBody & "<TR><TD width=120 vAlign=top><b>Details:</b></TD><TD colspan=3>" & strDetails & "</TD></TR>"
		strBody = strBody & "</Table>"
	
		strSubject = ""
		strNotes = ""
		strTo = ""
		strCC = ""
				
		'Owner Changed
		if trim(request("cboOwner")) <> trim(request("tagOwner")) and trim(request("txtCurrentUserID")) <> trim(request("cboOwner")) then
			'ASSERT: Owner Changed and current user is not owner
			strTo = strOwnerEmail
			if trim(request("tagOwner")) = "" then
				strNotes = "<b><font size=2 face=verdana>This task has been assigned to you.</font></b><BR>"
				strSubject = strSubject & "assigned to you"
			else
				strNotes = "<b><font size=2 face=verdana>This task has been reassigned to you.</font></b><BR>"
				strSubject = strSubject & "reassigned to you"
			end if
		end if

		'Item Closed
		if trim(request("cboStatus")) = 2 then
			if trim(request("txtCurrentUserID")) <> trim(request("cboOwner")) then
				strTo = strOwnerEmail & ";" & left(request("txtNotify"),255)
			else
				strTo = left(request("txtNotify"),255)
			end if
			if trim(request("txtCurrentUserID")) <> trim(request("cboFrom")) and request("chkCopySubmitter") = "1" then
				if trim(strCC) = "" then
					strCC = strSubmitterEmail
				else
					strCC = strSubmitterEmail & ";" & strCC
				end if
			end if

			if request("chkCopyMe") = "1" and strFrom <> "" then
				if trim(strCC) = "" then
					strCC = strFrom
				else
					strCC = strFrom & ";" & strCC
				end if
			end if
			
			strNotes = strNotes & "<b><font size=2 face=verdana>This task has been closed.</font></b><BR>"
			if strSubject = "" then
				strSubject = "closed"
			else
				strSubject = strSubject & " and closed"
			end if
		elseif trim(request("cboStatus")) = "1" and trim(request("tagStatus")) = "2" and  trim(request("txtCurrentUserID")) <> trim(request("cboOwner")) then
			strTo = strOwnerEmail
			strNotes = strNotes & "<b><font size=2 face=verdana>This task has been reopened.</font></b><BR>"
			if strSubject = "" then
				strSubject = "reopened"
			else
				strSubject = strSubject & " and reopened"
			end if
		end if
		
		strSubject = strProduct & " " & strType & " " & strSubject & " : " & strSummary
		
		if strTo <> "" then
			strBody = "<HTML><STYLE>TD{FONT-FAMILY: Verdana;FONT-SIZE: x-small;}</STYLE><BODY><font size=2 face=verdana>" & strNotes & "<BR>" & strBody & "</font></BODY></HTML>"
			Set oMessage = New EmailWrapper 'CreateObject("CDO.Message")
			'Set oMessage.Configuration = Application("CDO_Config")		
			oMessage.From = strFrom
			oMessage.To=strTO
			oMessage.CC=strCC
			oMessage.Subject = strSubject
			oMessage.HTMLBody = strBody
			oMessage.DSNOptions = cdoDSNFailure
			oMessage.Send 
			Set oMessage = Nothing 
		end if
	end if


	'Close AppError record if requested
	if strSuccess = "1" and trim(request("txtAppError")) <> "" then
		cn.Execute "spUpdateAppError " & clng(request("txtAppError")) & ",'Converted to action item " & strNewID & ".'"  
	end if

	'Link this action to the ticket it came from
	if strSuccess = "1" and trim(request("txtTicketID")) <> "" then
		cn.Execute "spSupportTicketConverted2Action " & clng(request("txtTicketID")) & "," & clng(strNewID )
	end if


	cn.Close
	set cn=nothing

	Response.Write strBody
%>
<INPUT type="hidden" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>
