<%@ Language="VBScript" %>

<%
  Dim AppRoot : AppRoot = Session("ApplicationRoot")
		
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>

<html>
<head>
<title>Excalibur</title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }


function PrintWindow(){
	PrintLink.style.display = "none";
	chkNA.disabled = false;
	chkEMEA.disabled = false;
	chkCKK.disabled = false;
	chkAPD.disabled = false;
	chkGCD.disabled = false;
	chkLA.disabled = false;
	chkCustomer.disabled = false;
	chkAdd.disabled = false;
	chkModify.disabled = false;
	chkRemove.disabled = false;
	chkStatus.disabled = false;
	window.print();
	window.close();
}

function MailWindow(){
	if (frmSend.txtTo.value == "")
		{
			window.alert("Please enter email recipients first.");
			frmSend.txtTo.focus();
			return;
		}
		
		
	PrintLink.style.display = "none";
	frmSend.txtEmailBody.value = ItemDetails.innerHTML;
	
	frmSend.submit();
}
function window_onload() {
	if (typeof(frmSend.txtTo) != "undefined")	
		frmSend.txtTo.focus();
}

function AddAddress(strEmail,strBox){
	var strText;
	
	if (strBox == 1)
		{
		strText = trim(frmSend.txtTo.value);
		if (strText == "")
			frmSend.txtTo.value = strEmail;
		else
			{
			if (strText.charAt(strText.length-1) != ";")
				frmSend.txtTo.value = strText + ";" + strEmail;
			else
				frmSend.txtTo.value = strText + strEmail;
			}
		}
	else
		{
		strText = trim(frmSend.txtCC.value);
		if (strText == "")
			frmSend.txtCC.value = strEmail;
		else
			{
			if (strText.charAt(strText.length-1) != ";")
				frmSend.txtCC.value = strText + ";" + strEmail;
			else
				frmSend.txtCC.value = strText + strEmail;
			}
		}
}

//-->
</script>
</head>
<body onload="return window_onload()">

<font size="2" face="verdana">
<% dim strType
	Select Case request("Type")
	case "1"
		strType = "Issue"
	case "3"
		strType = "Change Request"
	case "4"
		strType = "Status Note"
	case "5"
		strType = "Improvement Opportunity"
	case "6"
		strType = "Test Request"
	case else
		strType = "Action Item"
	end select
%>
<h3 align=center><%=strType & " " & request("ID")%></h3>
<Span ID=PrintLink>
<a href="javascript:PrintWindow();">Print</a>
<%if request("Action") = "1" then%>
<font size=2 face=verdana>&nbsp;|&nbsp;<a href="javascript:MailWindow();">Send Email</a>
<%end if%>
<BR></span>
<%
	dim rs 
	dim cn
	dim cm
	dim p
	dim strproducts
	dim strOwners
	dim DisplayForAdd
	dim DisplayForChangeOnly
	dim JustificationTemplate
	dim strID
	dim strPMID
	dim strSummary
	dim strRep
	dim strReps
	dim strSubmitter
	dim strSubmitted
	dim strTarget
	dim strNotify
	dim strAction
	dim strJustification
	dim strDescription
	dim strResolution
	dim strProgramID
	dim strOwnerID
	dim strNA
	dim strLA
	dim strEMEA
	dim strCKK
	dim strAPD
	dim strGCD
	dim strCoreTeamRep
	dim strStatus
	dim strStatuses	
	dim ClosureLabel
	dim strDisplayReport
	dim strOnlineReports
	dim strReportValue
	dim strApprovals
	dim strStatusText
	dim NoApprovals
	dim strSaveApprovals
	dim strApproverComments
	dim strDistribution
	dim strCTODate
	dim strBTODate
	dim DisplayDistribution
	dim BTOYes
	dim CTOYes
	dim BTONo
	dim CTONo
	dim DisplayBTODate
	dim DisplayCTODate 
	dim strAddChange
	dim strModifyChange
	dim strRemoveChange
	dim ApproversLoaded
	dim ApproversPending
	dim DescriptionHeight
	dim DescriptionTemplate
	dim DisplayRestore
	dim LanguageList
	dim strPriority
	dim strPriorityOptions
	dim strEditSubmitter
	dim strCustomers
	dim blnSubmitterFound
	dim  blnPreinstallApprover
	dim PreinstallOwnerID
	dim strPreinstallOwnerList
	dim blnProdFound
	dim strProductName
	dim strActionStatusText
	dim strOwner 
	dim strCoreTeamRepText
	dim strAvailableForTest
	dim strCoreTeamEmail
	dim strConsumer
	dim strCommercial
	dim strSMB
	dim ImpactArray
	dim strDetails
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	
	strPMID = ""
	strCoreTeamEmail = ""
	blnPreinstallApprover = false
	PreinstallOwnerID = ""
	strPreinstallOwnerList = ""
	strID = ""
	strSummary=""
	strDescription = ""
	strRep = ""
	strSubmitter = ""
	strSubmitted = ""
	strtarget = ""
	strNotify = ""
	strDescription = DescriptionTemplate
	strAction = ""
	strResolution = ""
	strJustification = ""
	strProgramId = ""
	strOwnerID = ""
	strNA = ""
	strLA = ""
	strEMEA = ""
	strCKK = ""
	strAPD = ""
	strGCD = ""
	strCoreTeamRep = ""
	strStatus = ""
	strStatuses = ""
	strPMID = ""
	strOnlineReports = ""
	strReportValue = ""
	strApprovals = ""
	strStatusText = ""
	strSaveApprovals = "0"
	strDistribution = ""
	strCTODate = ""
	strBTODate = ""
	BTOYes = ""
	CTOYes = ""
	BTONo = ""
	CTONo = ""
	DisplayBTODate = "none"
	DisplayCTODate = "none"
	strAddChange = ""
	strModifyChange = ""
	strRemoveChange = ""
	strOwner = ""	
	strPriority = ""
	LanguageList = ""
	strPriorityOptions	= ""
	strEditSubmitter = ""
	strCustomers = ""
	strProductName = ""
	strActionStatusText = ""
	strOwner = ""
	strCoreTeamRepText = ""
	strAvailableForTest = ""
	strAvailableNotes = ""
	strConsumer = ""
	strCommercial = ""
	strSMB = ""
	strDetails = ""
	
	On Error Resume Next
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetActionProperties4Print"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = Request("ID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing


		'rs.Open "spGetActionProperties4Print  " & Request("ID"),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strID = Request("ID")
			strAction  = rs("Actions") & ""
			strResolution  = rs("Resolution") & ""
			strSummary  = replace(rs("Summary") & "","""","&QUOT;")
			strDescription = rs("Description") & ""
			strRep = rs("CoreTeamRep") & ""
			strSubmitter = rs("Submitter") & ""
			strSubmitted = rs("Created") & ""
			strTarget = rs("TargetDate") & ""
			strNotify = rs("Notify") & "&nbsp;"
			strJustification = rs("Justification") & ""
			strProgramId = rs("ProductVersionID") & ""
			strOwnerId = rs("OwnerID") & ""
			If rs("Americas") Then strAmericas = "checked"
			If rs("APJ") Then strAPJ = "checked"
			If rs("EMEA") Then strEMEA = "checked"
			strStatus  = rs("Status") & ""
			strPMID = rs("PMID") & ""
			strOnlineReports = rs("OnlineReports") & ""
			strReportValue = replace(Replace(rs("OnStatusReport") & "","1","checked"),"0","")
			strDistribution = rs("Distribution") & ""
			strCTODate = rs("CTODate") & ""
			strBTODate = rs("BTODate") & ""
			strAddChange = replace(Replace(rs("AddChange") & "","Yes","checked"),"No","")
			strModifyChange = replace(Replace(rs("ModifyChange") & "","Yes","checked"),"No","")
			strRemoveChange = replace(Replace(rs("RemoveChange") & "","Yes","checked"),"No","")
			strPriority = rs("Priority") & ""
			if trim(request("TYPE")) = "5" then
				if rs("AffectsCustomers")=1  then
					strCustomers = "Positive"
				elseif rs("AffectsCustomers")=0 then
					strCustomers = "&nbsp;"
				else
					strCustomers = "Negative"
				end if
			else
				strCustomers = replace(Replace(rs("AffectsCustomers") & "","1","checked"),"0","")
			end if
			PreinstallOwnerID = rs("PreinstallOwnerID") & ""
			strProductname = rs("Productname") & ""
			strCoreTeamEmail = rs("CoreTeamEmail") & ""
			strOwner = rs("Owner") & ""
			strCoreTeamRepText = rs("CoreTeamRepText") & ""
			strAvailableForTest = rs("AvailableForTest") & ""
			strAvailableNotes = rs("AvailableNotes") & ""
			strConsumer = replace(replace(rs("Consumer") & "","True","checked"),"False","")
			strCommercial = replace(replace(rs("Commercial") & "","True","checked"),"False","")
			strSMB = replace(replace(rs("SMB") & "","True","checked"),"False","")
			strDetails = rs("Details") & ""
		end if
		rs.Close
	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovals"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = Request("ID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spListApprovals " & Request("ID"),cn,adOpenForwardOnly
		do while not rs.EOF
			strStatusText = rs("Status")
			select case strStatusText
			case "1"
				strStatusText = "Approval Requested"
			case "2"
				strStatusText = "Approved"
			case "3"
				strStatusText = "Disapproved"
			case "4"
				strStatusText = "Cancelled"
			end select
			
			strApprovals = strApprovals & "<TR><TD nowrap><font size=1 face=verdana>" & rs("Approver") & "</font></td><TD nowrap><font face=verdana size=1>" & strStatusText & "</font></td><TD><font face=verdana size=1>" & rs("Comments") & "&nbsp;" & "</font></TD></TR>" 
			rs.MoveNext
		loop			
		rs.Close

		if trim(strApprovals) = "" then
			strApprovals = "No Approvers Assigned"
		else
			strApprovals = "<TABLE bgcolor=ivory border=1 cellpadding=2 cellspacing=0><TR bgcolor=Gray><TD><font color=white size=1 face=verdana><b>Approver</b></font></TD><TD><font  color=white size=1 face=verdana><b>Status</b></font></TD><TD><font color=white size=1 face=verdana><b>Comments</b></font></TD></TR>" & strApprovals & "</TABLE>"
		end if
		
		
	dim CurrentUser
	dim CurrentUserEmail
	CurrentUserEmail = ""

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
		CurrentUserEmail = rs("Email") & ""
	end if
	rs.Close
		
		
	set rs = nothing
	cn.Close
	set cn = nothing

	Select case strStatus
	case "1"
		strActionStatusText = "Open"
	case "2"
		strActionStatusText = "Closed"
	case "3"
		strActionStatusText = "Need More Information"
	case "4"
		strActionStatusText = "Approved"
	case "5"
		strActionStatusText = "Disapproved"
	case "6"
		strActionStatusText = "Investigating"
	case else
		strActionStatusText = "&nbsp;"
	end select
	
	select case strPriority
	case "1"
		strPriority="High"
	case "2"
		strPriority="Medium"
	case "3"
		strPriority="Low"
	case else
		strPriority="&nbsp;"
	end select
	
	
	if trim(CurrentUserEmail) = "" and trim(request("Action")) = "1" then
		Response.Write "<BR><BR>Only registered Excalibur users can send email using this function."
		%>
		<SCRIPT>
		PrintLink.style.display = "none";
		</SCRIPT>
		<%
	else
%>
<BR />
	<form name="frmSend" method="post" action="ActionEmail.asp">
	<%if request("Action") = "1" then%>
		<font size=1 face=verdana color=green>Enter SMTP Email addresses only (i.e., john.doe@hp.com)</font><BR>
		<Table bgcolor=ivory border=1 cellspacing=0 cellpadding=2 width=100%>
		<TR>
			<TD valign=top width=150><font face=verdana size=1><b>Send To:</b>
			<%if strCoreTeamEmail <> "" then%>
				<BR>&nbsp;&nbsp;<a href="javascript:AddAddress('<%=strCoreTeamEmail%>',1)">System Team</a>
			<%end if%>
			</font></TD>
			<TD valign=top>
				<INPUT style="Width=100%" type="text" id=txtTo name=txtTo>
			</TD>
		</TR>
		<TR>
			<TD valign=top width=150><font face=verdana size=1><b>CC:</b>
			<%if strCoreTeamEmail <> "" then%>
				<BR>&nbsp;&nbsp;<a href="javascript:AddAddress('<%=strCoreTeamEmail%>',2)">System Team</a>
			<%end if%>
			
			</font></TD>
			<TD valign=top>
				<INPUT style="Width=100%" type="text" id=txtCC name=txtCC>
			</TD>
		</TR>
		<TR>
			<TD valign=top><font face=verdana size=1><b>Subject:</b></font></TD>
			<TD valign=top><INPUT style="Width=100%" type="text" id=txtSubject name=txtSubject value="<%=strType & " " & request("ID")%>"></TD>
		</TR>
		<TR>
			<TD valign=top><font face=verdana size=1><b>Notes:</b></font></TD>
			<TD valign=top><TEXTAREA rows=5 style="Width=100%" cols=20 id=txtNotes name=txtNotes></TEXTAREA></TD>
		</TR>
		</Table><INPUT type="hidden" id=txtFrom name=txtFrom value="<%=CurrentUserEmail%>">
		<TEXTAREA style="Display:none" rows=2 cols=20 id=txtEmailBody name=txtEmailBody></TEXTAREA>
		<font face=verdana size=2><BR><b><%=strType%> Details:</b></font>
		
	<%end if%>
		</form>
	<span ID=ItemDetails>
	<Table border=1 cellspacing=0 cellpadding=2 width=100%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Submitter:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strSubmitter%></font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Date Submitted:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strSubmitted%></font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Program:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strProductname%></font></TD>
	</TR>
	<TR>
		<TD width=150 valign=top><font face=verdana size=1><b>Summary:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strSummary%></font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Status:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strActionStatusText%></font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Owner:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strOwner%></font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Business:</b></font></TD>
		<TD valign=top><font face=verdana size=1>
		<INPUT disabled type="checkbox" id=chkConsumer name=chkConsumer <%=strConsumer%>>&nbsp;Consumer&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=chkCommercial name=chkCommercial <%=strCommercial%>>&nbsp;Commercial&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=chkSMB name=chkSMB <%=strSMB%>>&nbsp;SMB&nbsp;&nbsp;
		</font></TD>
	</TR>
	<%if trim(request("TYPE")) = "3" then%> 
	<TR>
		<TD valign=top><font face=verdana size=1><b>GEOS:</b></font></TD>
		<TD valign=top><font face=verdana size=1>
		<INPUT disabled type="checkbox" id=Checkbox1 name=chkAmericas <%=strAmericas%>>&nbsp;Americas&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=Checkbox2 name=chkEMEA <%=strEmea%>>&nbsp;EMEA&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=Checkbox3 name=chkAPJ <%=strAPJ%>>&nbsp;APJ&nbsp;&nbsp;
		</font></TD>
	</TR>
    <%end if%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Description:</b></font></TD>
		<% 
			if trim(request("TYPE")) = "5" then
				ImpactArray = split(strDescription,chr(1))
				if ubound(ImpactArray) > -1 then
					if trim(ImpactArray(0)) <> "" then
						strDescription = "<b>Positive Impact:</b><br>" & ImpactArray(0)
					else
						strDescription = ""				
					end if
				end if
				if ubound(ImpactArray) > 0 then
					if trim(ImpactArray(0)) <> "" and trim(ImpactArray(1)) <> ""  then
						strDescription = strDescription & "<BR><BR>"
					end if
					if trim(ImpactArray(1)) <> "" then
						strDescription = strDescription & "<b>Negative Impact:</b><br>" & ImpactArray(1)
					end if
				end if
			end if
		%>
			<TD valign=top><font face=verdana size=1><%=strDescription%>&nbsp;</font></TD>
	</TR>
    <tr>
        <td valign="top"><font face=verdana size=1><b>Details:</b></font></td>
        <td valign=top><font face=verdana size=1><%=strDetails %>&nbsp;</font></td>
    </tr>    
	<%if strJustification <> "" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Justification:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strJustification%>&nbsp;</font></TD>
	</TR>
	<%end if%>
	<%if strResolution <> "" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Resolution:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strResolution%>&nbsp;</font></TD>
	</TR>
	<%end if%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Actions Needed:</b></font></TD>
		<% 
			if trim(request("TYPE")) = "5" then
				ImpactArray = split(strAction,chr(1))
				if ubound(ImpactArray) > -1 then
					if trim(ImpactArray(0)) <> "" then
						strAction = "<b>Corrective Actions:</b><br>" & ImpactArray(0)
					else
						strAction = ""				
					end if
				end if
				if ubound(ImpactArray) > 0 then
					if trim(ImpactArray(0)) <> "" and trim(ImpactArray(1)) <> ""  then
						strAction = strAction & "<BR><BR>"
					end if
					if trim(ImpactArray(1)) <> "" then
						strAction = strAction & "<b>Preventive Actions:</b><br>" & ImpactArray(1)
					end if
				end if
			end if
		%>
		<TD valign=top><font face=verdana size=1><%=strAction%>&nbsp;</font></TD>
	</TR>




	<%if trim(strPriority) <> "" and trim(request("TYPE")) = "5" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Impact:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strPriority%></font></TD>
	</TR>
	<% end if%>	
	<%if trim(request("TYPE")) = "5" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Net Affect:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strCustomers%></font></TD>
	</TR>
	<% end if%>	
	<%if trim(strAvailableNotes) <> "" and trim(request("TYPE")) = "5" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Metric Impacted:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strAvailableNotes%></font></TD>
	</TR>
	<% end if%>	
	<%if trim(strAvailableForTest) <> "" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Samples Available:</b></font></TD>
		<TD valign=top><font color=red face=verdana size=1><%=strAvailableForTest%></font></TD>
	</TR>
	<% end if%>
	<%if trim(strAvailableNotes) <> "" and trim(request("TYPE")) <> "5" then%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Availability Notes:</b></font></TD>
		<TD valign=top><font color=red face=verdana size=1><%=strAvailableNotes%></font></TD>
	</TR>
	<% end if%>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Customer Impact:</b></font></TD>
		<TD valign=top><font face=verdana size=1>
		<INPUT disabled type="checkbox" id=chkCustomer name=chkCustomer <%=strCustomers%> disabled>&nbsp;Affects images and/or BIOS on shipping products
		</font></TD>
	</TR>
	<TR>
		<TD nowrap valign=top><font face=verdana size=1><b>Notify On Approval:</b></font></TD>
		<TD valign=top><font face=verdana size=1><%=strNotify%></font></TD>
	</TR>
	<%if trim(request("TYPE")) = "3" then%> 
	<%else%>
	<TR style="Display:none"><TD colspan=2>
		<INPUT disabled type="checkbox" id=Checkbox4 name=chkAmericas <%=strAmericas%>>&nbsp;Americas&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=Checkbox5 name=chkEMEA <%=strEmea%>>&nbsp;EMEA&nbsp;&nbsp;
		<INPUT disabled type="checkbox" id=Checkbox6 name=chkAPJ <%=strAPJ%>>&nbsp;APJ&nbsp;&nbsp;
	</TD></TR>
	<%end if%>
	
	<TR>
		<TD valign=top><font face=verdana size=1><b>Status Report:</b></font></TD>
		<TD valign=top><font face=verdana size=1>
		<INPUT disabled type="checkbox" id=chkStatus name=chkStatus <%=strReportValue%> disabled>&nbsp;Display in Online Status Report
		</font></TD>
	</TR>
	<TR>
		<TD valign=top><font face=verdana size=1><b>Approvals:</b></font></TD>
		<TD valign=top><font face=verdana size=1>
		<%=strApprovals%>
		</font></TD>
	</TR>
	
</Table>
<BR /><BR /><BR />
<font size=1 face=verdana>Report Generated <%=Date()%></font>

</font>
</span>
<%end if%>
</body>
</html>
