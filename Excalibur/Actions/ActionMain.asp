<%@ Language=VBScript %>
<%
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
%>
<!DOCTYPE html>
<HTML>
<head>  
<meta charset="utf-8">
<title>Actions</title>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../_ScriptLibrary/jsrsClient.js" language="javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

$(function() {
	$("#txtTargetDate").datepicker();


	var CheckIllegalChars = function (e) {
		var pastedData;

		if (e.originalEvent.clipboardData === undefined)//IE
			pastedData = clipboardData.getData('text');
		else
			pastedData = e.originalEvent.clipboardData.getData('text');

		var clean = pastedData.replace(/[^\x20-\x7E\r\n]/g, '_');//replace any nont printable character

		if (clean != pastedData) {
			var msg = 'Invalid characters detected(location marked with _):\n'
			alert(msg + clean);
		}
	};


	var ua = window.navigator.userAgent;
	var msie = ua.indexOf("MSIE ");
	if (msie > 0) // If Internet Explorer
	{
		$("body").off('paste');
		$("body").on('paste', function (e) { CheckIllegalChars(e); });
	}
	else {
		$(document).on('paste', function (e) { CheckIllegalChars(e); });
	}



  });

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					break;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}


function Left(str, n)
{
    if (n <= 0)     // Invalid bound, return blank string
	    return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
}


//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs     
//*****************************************************************
function window_onload() {
	frmUpdate.txtSummary.value=frmUpdate.tagSummary.value;
    frmUpdate.txtNotify.value=frmUpdate.txtNotifyList.value;
	//frmUpdate.txtSummary.value=frmUpdate.tagSummary.value;
    if (txtAccessOK.value == "1") {
        frmUpdate.txtSummary.focus();
    } else {
        window.parent.frames["LowerWindow"].cmdOK.disabled = true;
    }

    //Add modal dialog code to body tag: ---
    modalDialog.load();

    //Add datepicker to date fields
    load_datePicker();
}

//*****************************************************************
//Description:  Close Dialog Window
//Function:     CloseDialog();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs      
//*****************************************************************
function CloseDialog() {
    //close child dialog window
    modalDialog.cancel();
}


function myCallback( returnstring ){
	CellRoadmap.innerHTML = returnstring; 
	frmUpdate.txtNotify.value=frmUpdate.txtDefaultNotify.value;
}

function cboProject_onchange() {

	var	strID = event.srcElement.value;
	if (event.srcElement.value !="")
		{
	      jsrsExecute("ActionRSget.asp", myCallback, "getItem", strID);
		}
}

//*****************************************************************
//Description:  Open Email Address
//Function:     cmdAdd_onclick();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367      
//*****************************************************************
function cmdAdd_onclick() {
    var strResult;
    modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../Email/AddressBook.asp?AddressList=' + frmUpdate.txtNotify.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: false, dialogDraggable: true });
    globalVariable.save('txtNotify', 'email_field');
}
//-->
</SCRIPT>
</HEAD>

<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<%

	dim cn
	dim rs
	dim i
	dim cm
	dim p
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserEmail
	dim strID
	dim strSummary
	dim strProduct
	dim strTimeframe
	dim strStatus
	dim strMilestone
	dim strNotes
	dim strDetails
	dim blnFound
	dim strOrder
	dim strPriority
	dim strOwner
	dim strPM
	dim strWorking
	dim strReviewInput
	dim strType
	dim strSubmitter
	dim strRoadmapID
	dim strResolution
	dim strSponsorID
	dim strTargetDate
	dim strDuration
	dim strCopyMe
	dim strCopySubmitter
	dim strDefaultproduct
	dim strDefaultNotify
    dim ProductType
    dim strProductName
    dim strStatusNotesUpdatedDt
    dim strOriginalTargetDt
    dim strStatusNotes
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
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
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("Email") & ""
		CurrentUserDefaultProduct = rs("DefaultWorkingListProduct") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	if trim(Request("ProdID")) = "" or trim(Request("ProdID"))="0" then
		if CurrentUserDefaultProduct <> "" and isnumeric(CurrentUserDefaultProduct) then
			strProduct = CurrentUserDefaultProduct
		else
			strProduct = 0 '235 'Excalibur
		end if
	else
		strProduct = trim(Request("ProdID"))
	end if


	strID = trim(Request("ID"))
	strType = trim(request("Type"))
    if strtype = "" then
        strtype = 2
    end if
	strSummary = ""
	strWorking = ""
	strReviewInput = ""
	strStatus = ""
	strDetails = ""
	strPriority = 0
	strSponsorID=0
	strDuration = ""
	strTargetDate =""
	strOwner = ""
	strPM = ""
	strOrder = ""
	strSubmitter = ""
	strNotify = ""
	strDefaultNotify = ""
	strRoadmapID = ""
	strResolution = ""
	strCopyMe = ""
	strCopySubmitter =""
    strProductName = ""
    strStatusNotesUpdatedDt = ""
    strOriginalTargetDt = ""
    strStatusNotes = ""
	
	if trim(request("Working") ) = "1" then
		strWorking = "checked"
	else
		strWorking = ""
	end if
	if (strID = ""  or trim(strID) = "0") and request("RoadmapID") <> "" then
		strRoadmapID = request("RoadmapID")
	end if
    strTicketID = ""
	if (trim(strID)="0" or trim(strID)="") and request("TicketID") <> "" then
        strTicketID = clng(request("TicketID"))
        rs.open "spSupportTicketSelect " & clng(request("TicketID")),cn,adOpenForwardOnly
	    if not (rs.eof and rs.bof) then
            strSubmitter = rs("SubmitterName") & ""
            strOwner = rs("OwnerID") & ""
            strDetails = rs("Details") & ""
            strResolution = rs("Resolution") & ""
            strSummary = rs("Summary") & ""
            strProduct= rs("actionprojectid") & ""
            strWorking = "checked"
        end if
        rs.close
    elseif trim(strID)="0" and request("AppErrorID") <> "" then
		rs.Open "spGetAppError " & clng(request("AppErrorID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strDetails = "Research application error:" & vbcrlf
			strDetails = strDetails & "ERROR_ID: " & request("AppErrorID") & vbcrlf
			strDetails = strDetails & "FOUND: " & rs("AuthUser") & " - " & rs("ErrorDateTime") & vbcrlf
			strDetails = strDetails & "FILE: " & rs("ErrFile") & vbcrlf
			strDetails = strDetails & "LINE: " & rs("ErrLine") & vbcrlf
			strDetails = strDetails & "COLUMN: " & rs("ErrColumn") & vbcrlf
			strDetails = strDetails & "DESCRIPTION: " & rs("ErrDescription") & vbcrlf
			strDetails = strDetails & "INPUT: "
			strForm = trim(rs("RequestForm") & "")
			if strForm = "" then
				strForm = trim(rs("RequestQueryString") & "")
			else
				strForm = trim(rs("RequestQueryString") & "") & "&" & trim(rs("RequestForm") & "")	
			end if 
			strForm = replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(replace(strForm,"%25","%"),"%2F","/"),"%3D","="),"%22",""""),"%27","'"),"%3C","<"),"%3E",">"),"%3B",";"),"%5C","\"),"%2C",","),"+"," "),"&",vbcrlf)
			strDetails = strDetails & strForm & vbcrlf
			if strRoadmapID = "" then
				strRoadmapID = 21
			end if
			if trim(rs("Submitter") & "") <> "" then
				strSubmitter = trim(rs("Submitter") & "")
			end if
			strWorking = "checked"
		end if
		rs.Close
	end if
	
	if strID <> "" and strID <> "0" then
		rs.Open "spGetActionProperties " & clng(strID),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			blnFound = false
            ProductType = 0
		else
			blnFound = true
			strProduct = rs("ProductVersionID") & ""
            strProductName = rs("DotsName") & ""
			strStatus = rs("Status") & ""
			strSummary = rs("Summary") & ""
			strSubmitter = rs("Submitter") & ""
			strOwner = rs("OwnerID") & ""
			strNotify = rs("Notify") & ""
			strType = rs("Type") & ""
            ProductType = rs("TypeID")
			strDuration = rs("Duration") & ""
			strRoadmapID = rs("ActionRoadmapID") & ""
			strSponsorID = rs("SponsorID") & ""
			strPM = rs("PMID") & ""
			strTargetDate = rs("targetDate") & ""
			strPriority = trim(rs("Priority") & "")
			if rs("OnStatusReport") then
				strReviewInput = "checked"
			else
				strReviewInput = ""
			end if
			if rs("PendingImplementation") then
				strWorking = "checked"
			else
				strWorking = ""
			end if
			if rs("AffectsCustomers") = 1 then
				strCopyMe = "checked"
			else
				strCopyMe = ""
			end if
			if rs("PreinstallOwnerID") = 1 then
				strCopySubmitter = "checked"
			else
				strCopySubmitter = ""
			end if
			strOrder = rs("DisplayOrder") & ""
			strResolution = rs("Resolution") & ""
			strDetails = rs("Description") & ""
			strStatusNotesUpdatedDt = trim(rs("StatusNotesUpdatedDt") & "")
            strOriginalTargetDt = trim(rs("OriginalTargetDt") & "")
            strStatusNotes = trim(rs("StatusNotes") & "")
		end if
		rs.Close
	else
		if isnumeric(trim(strProduct)) and trim(strProduct) <> "" and trim(strProduct) <> "0" then
			rs.Open "spGetProductVersion " & clng(strProduct),cn,adOpenStatic
			if rs.eof and rs.bof then
                ProductType = 0
                strNotify = ""
			    strDefaultNotify = ""
            elseif rs("TypeID") = 1 then
                ProductType = rs("TypeID")
                strNotify = rs("Distribution") & ""
			    strDefaultNotify = strNotify
                strProductName = rs("DotsName") & ""
            else
                ProductType = rs("TypeID")
                strNotify = rs("ActionNotifyList") & ""
			    strDefaultNotify = strNotify
	            strProductName = rs("Name") & ""
    		end if
            rs.Close
		end if
		blnFound = true
	end if

'if strOrder = "0" and strID <> "" then
'    strOrder = "1"
'end if
if not blnFound then
	Response.Write "Unable to find the requested task."
else
	
	'Check to see if Update is OK
	dim blnToolsPM
	dim strToolAccessList
	dim blnActionOwner
	
	blnToolsPM = false
	blnActionOwner = false
	strToolsAccessList = ""
	
	rs.Open "spListToolsPMs",cn,adOpenKeyset
	do while not rs.EOF
		if trim(CurrentUserID) = trim(rs("ID")) then
			blnToolsPM = true
			exit do
		end if
		rs.MoveNext
	loop
	rs.Close	
	
	if isnumeric(trim(strProduct)) and trim(strProduct) <> "" and trim(strProduct) <> "0" then
		rs.Open "spGetProductVersion " & clng(strProduct),cn,adOpenStatic
		strToolAccessList = rs("ToolAccessList") & ""
		rs.Close

		if instr("," & strToolAccessList & ",","," & trim(CurrentUserID) & ",")> 0 then
			blnActionOwner = true
		else
			rs.Open "spListToolsProjectOwners " & clng(strProduct),cn,adOpenKeyset
			do while not rs.EOF
				if trim(CurrentUserID) = trim(rs("ID")) then
					blnActionOwner = true
					exit do
				end if
				rs.MoveNext
			loop
			rs.Close
		end if	
	end if

		
	
%>

<font face=verdana size=><b>
<label ID="lblTitle">
<%if strID = "" or strID="0" then%>
	Add
<%else%>
	Update
<%end if%>
<%if strType = "1" then%>
Issue
<%else%>
Task
<%end if%>
<%
if strID <> "" and strID<>"0" and isnumeric(strid) then
	Response.Write clng(strID)
end if
%>
</label></b></font>

<%if blnActionOwner or blnToolsPM or trim(strProduct) = "" or trim(strProduct) = "0" or trim(strID)="0" or trim(strID)="" then %>
	<div>
<%else%>
	<div disabled>
<%end if%>
<form id="frmUpdate" method="post" action="ActionSave.asp">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<td valign=top width=120 nowrap><b>Summary:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD colspan=3>
            <INPUT style="width:100%" type="text" id=txtSummary name=txtSummary maxlength=120 mytag="<%=replace(strSummary,"""","&quot;")%>" value="">
            <INPUT style="width:100%" type="hidden" id="tagSummary" maxlength=120 value="<%=replace(strSummary,"""","&quot;")%>">
        </TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Project:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD width=50%><SELECT style="width:100%" id=cboProject name=cboProject  LANGUAGE=javascript onchange="return cboProject_onchange()" onkeydown="return combo_onkeydown()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeypress="return combo_onkeypress()">
				<%
					Response.Write "<OPTION selected value=""""></OPTION>"
                    response.write "<optgroup label=""---Tools------------""></optgroup>"
					rs.Open "spGetProducts 2",cn,adOpenForwardOnly
                    blnSelected = false
					do while not rs.EOF
						if trim(rs("ID")) = trim(strProduct) then
							Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
							strPM = rs("PMID") & ""
                            blnSelected = true
						else
							Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
						end if
						rs.MoveNext
					loop
					rs.Close

                    response.write "<optgroup label=""---Products------------""></optgroup>"
					rs.Open "spGetProducts 1",cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(rs("ID")) = trim(strProduct) then
							Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("name") & " " & rs("Version") & "</OPTION>"
							strPM = rs("PMID") & ""
                            blnSelected = true
						else
							Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("name") & " " & rs("Version") & "</OPTION>"
						end if
						rs.MoveNext
					loop
					rs.Close




                    if request("ProdID") <> "" and (not blnSelected) then
    					Response.Write "<OPTION selected value=""" & clng(request("ProdID")) & """>" & strProductName & "</OPTION>"
                    end if
				%>
			</SELECT>
			</TD>
			<td valign=top width=120 nowrap><b>Owner:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD width=50%><SELECT style="width:100%" id=cboOwner name=cboOwner LANGUAGE=javascript onkeydown="return combo_onkeydown()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeypress="return combo_onkeypress()">
				<%
					rs.Open "spGetEmployees",cn,adOpenForwardOnly
					do while not rs.EOF
						if rs("ID") <> 646 then
							if trim(rs("ID")) = trim(strOwner) or (trim(CurrentUserID) = trim(rs("ID")) and strID="0" ) then
								Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
								if strSubmitter = "" then
									strSubmitter = rs("name") & ""
								end if
							elseif rs("Active") = 1 then
								Response.Write "<OPTION value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
							end if
						end if
						rs.MoveNext
					loop
					rs.Close
				%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Status:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD><SELECT style="width:100%" id=cboStatus name=cboStatus>
				<%
					rs.Open "spListActionStatuses 3",cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(rs("ID")) = trim(strStatus) or (trim(strStatus)="" and rs("ID")=1) then
							Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						else
							Response.Write "<OPTION value=" & rs("ID") & ">" & rs("name") & "</OPTION>"
						end if
						rs.MoveNext
					loop
					rs.Close
				%>
			</SELECT>
		</TD>
		<td valign=top width=110 nowrap><b>Submitter:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD nowrap>
		<SELECT style="width:100%" id=cboFrom name=cboFrom LANGUAGE=javascript onkeydown="return combo_onkeydown()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeypress="return combo_onkeypress()">
		<%
					rs.Open "spGetEmployees",cn,adOpenForwardOnly
					do while not rs.EOF
						if rs("ID") <> 646 then
							if trim(rs("Name")) = trim(strSubmitter) then
								Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
							elseif rs("Active") = 1 then
								Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
							end if
						end if
						rs.MoveNext
					loop
					rs.Close
		
		%>
		</select>
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Priority:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD><SELECT style="width:100%" id=cboPriority name=cboPriority>
				<%
					for i = 0 to 3
						if strPriority = trim(i) then
							Response.Write "<option selected value=""" & i & """>" & i & "</option>"					
						else
							Response.Write "<option value=""" & i & """>" & i & "</option>"					
						end if
					next
				%>
			</SELECT>
		</TD>
		<td valign=top width=110 nowrap><b>Working&nbsp;List&nbsp;Order:</b></td>
		<TD nowrap><INPUT type="text" style="width:100%" id=txtOrder name=txtOrder value="<%=strOrder%>"></TD>
	</TR>
	<TR>
	    <td valign=top width=110 nowrap><b>Hours&nbsp;Required:</b></td>
		<TD nowrap><INPUT type="text" style="width:100%" id=txtDuration name=txtDuration value="<%=strDuration%>"></TD>
		<td valign=top width=120 nowrap><strong>Target&nbsp;Date:</strong></td>
		<TD nowrap>
			<INPUT style="width:80%" type="text" id=txtTargetDate name=txtTargetDate value="<%=strTargetDate%>" <% if blnActionOwner or blnToolsPM then %> class="dateselection" <%end if%>>
		</TD>
	</TR>
	<TR height=30>
		<td valign=top width=110 nowrap><b>Notes&nbsp;Updated&nbsp;Date:</b></td>
          <TD nowrap>&nbsp;<%=strStatusNotesUpdatedDt%></TD>
		<td valign=top width=120 nowrap><b>Original&nbsp;Target&nbsp;Date:</b></td>
		  <TD>&nbsp;<%=strOriginalTargetDt%></TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Status&nbsp;Notes:</b></td>
		<TD colspan=3><TEXTAREA rows=4 style="Width:100%" id=txtStatusNotes name=txtStatusNotes><%=strStatusNotes%></TEXTAREA>
	</TD>
	<TR>
		<td valign=top width=120 nowrap><b>Resolution:</b></td>
		<TD colspan=3><TEXTAREA rows=4 style="Width:100%" id=txtResolution name=txtResolution><%=strResolution%></TEXTAREA>
	</TD>
	</TR>
	<TR>
		<td valign=top width=110 nowrap><b>Notify:</b></td>
		<TD colspan=3 nowrap>
			<TABLE width=100% cellpadding=0 cellspacing=0 border=0><TR><TD width=100%><INPUT type="text" style="width:100%" id=txtNotify name=txtNotify value=""></TD><TD><INPUT type="button" value="Add" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()"></TD></TR></TABLE>
			<INPUT  <%=strCopySubmitter%> type="checkbox" id=chkCopySubmitter name=chkCopySubmitter value="1"> Copy Submitter&nbsp;&nbsp;<INPUT <%=strCopyMe%> type="checkbox" id=chkCopyMe name=chkCopyMe value="1"> Copy Me
		</TD>
	</TR>
    <%if ProductType = 2 then%>
	    <TR>
    <%else %>
        <TR style="display:none">
    <%end if%>
		<td valign=top width=120 nowrap><b>Roadmap&nbsp;Item:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD colspan=3 ID=CellRoadmap><SELECT style="width:100%" id=cboRoadmap name=cboRoadmap>
				<OPTION selected value=0>TBD</OPTION>
				<%
					rs.Open "spListActionRoadmap " & clng(strProduct),cn,adOpenForwardOnly
					do while not rs.EOF
						if trim(strRoadmapID) = trim(rs("ID")) then
							Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("DisplayOrder") & ". " & rs("Summary") & "</OPTION>"
						else
							Response.Write "<OPTION value=" & rs("ID") & ">" & rs("DisplayOrder") & ". " & rs("Summary") & "</OPTION>"
						end if
						rs.MoveNext
					loop
					rs.Close
				%>
			</SELECT>
			<INPUT type="hidden" id=txtDefaultNotify name=txtDefaultNotify value="<%=strDefaultNotify%>">
		</TD>
	</TR>
    <%if ProductType = 2 then%>
	    <TR>
    <%else %>
        <TR style="display:none">
    <%end if%>
		<td valign=top width=120 nowrap><b>Sponsor:</b><strong><font color="red" size="1">&nbsp;*</font></strong></td>
		<TD colspan=3 ID=CellSponsor><SELECT style="width:100%" id=cboSponsor name=cboSponsor>
				<Option value=0 selected>None</Option>
				<%
					rs.Open "spListActionSponsor",cn,adOpenForwardOnly
					do while not rs.EOF
						if rs("ID") <> 0 then
							if trim(strSponsorID) = trim(rs("ID")) then
								Response.Write "<OPTION selected value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
							else
								Response.Write "<OPTION value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
							end if
						end if
						rs.MoveNext
					loop
					rs.Close
				%>
			</SELECT>
			
		</TD>
	</TR>
	
	<TR>
		<td valign=top width=120 nowrap><b>Working&nbsp;List:</b></td>
		<TD colspan=3><INPUT <%=strWorking%> type="checkbox" id=chkWorking name=chkWorking>&nbsp;Show this item in the owner's working list.</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Review&nbsp;Input:</b></td>
		<TD colspan=3><INPUT <%=strReviewInput%> type="checkbox" id=chkReviewInput name=chkReviewInput>&nbsp;Flag this item as important for next year's FPR Input.</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Details:</b></td>
		<TD colspan=3><TEXTAREA rows=11 style="Width:100%" id=txtDetails name=txtDetails><%=strDetails%></TEXTAREA>
	</TD>
	<TR>
		<td valign=top width=120 nowrap><b>Email Note:</b></td>
		<TD colspan=3>
			<TEXTAREA rows=3 style="Width:100%" id=txtEmailNote name=txtEmailNote></TEXTAREA>
			<font size=1 face=verdana color=green>Note: This text is only added to any emails generated by this update.</font>
	</TD>
	</TR>
</table>


<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=strID%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtCurrentUserEmail name=txtCurrentUserEmail value="<%=lcase(trim(CurrentUserEmail))%>">
<INPUT type="hidden" id=tagDisplayOrder name=tagDisplayOrder value="<%=strPriority%>">
<INPUT type="hidden" id=txtType name=txtType value="<%=strType%>">
<INPUT type="hidden" id=tagStatus name=tagStatus value="<%=trim(strStatus)%>">
<INPUT type="hidden" id=tagOwner name=tagOwner value="<%=trim(strOwner)%>">
<INPUT type="hidden" id=txtNotifyList name=txtNotifyList value="<%=trim(strNotify)%>">


<INPUT type="hidden" id=txtAppError name=txtAppError value="<%=request("AppErrorID")%>">
<INPUT type="hidden" id=txtTicketID name=txtTicketID value="<%=strTicketID%>">
</form>
</div>
<%

end if

	set rs = nothing
	cn.Close
	set cn = nothing

dim AccessOK
if blnActionOwner or blnToolsPM or trim(strProduct) = "" or trim(strProduct) = "0" or trim(strID)  = "0" or trim(strID) = "" then
	AccessOK = "1"
else
	AccessOK = "0"
end if
%>
<INPUT type="hidden" id=txtAccessOK name=txtAccessOK value="<%=AccessOK%>">
</BODY>
</HTML>


