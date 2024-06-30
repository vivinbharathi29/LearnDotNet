<%@ Language=VBScript %>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "../../_ScriptLibrary/sort.js" -->

var KeyString = "";

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


function cmdDate_onclick(FieldID) {
	var strID;
		
	strID = window.showModalDialog("../../mobilese/today/caldraw1.asp", $("#txtAccessoryDate").val(),"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		$("#txtAccessoryDate").val(strID);
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


function window_onload() {
	frmStatus.cboStatus.focus();
}

function cboStatus_onchange() {
    var StatusVal = $("#cboStatus").val();
    var strRequired = $("#txtDateRequired").val().indexOf("," + StatusVal + ",");
	var strShow = $("#txtDateShow").val().indexOf("," + StatusVal + ",");
	var strStatus = $("#cboStatus option:selected").text();

	var strCommentsRequired = $("#txtCommentsRequired").val().indexOf("," + StatusVal + ",");
	
	if (strShow != -1)
		{
		$("#DateRow").show();
		if (strRequired != -1)
			$("#DateStar").show();
		else
			$("#DateStar").hide();
		}
	else
		{
		$("#DateRow").hide();
		$("#txtAccessoryDate").val("");
		$("#DateStar").hide();
		}

	if (strCommentsRequired == -1)
		$("#CommentStar").hide();
	else
		$("#CommentStar").show();

	var StatusSelectedIndex = $("#cboStatus option:selected").index();
	if (StatusSelectedIndex == 0 && $("#txtQualStatus").val() == "Not Used")
		$("#SupportRow").show();
	else
		$("#SupportRow").hide();
	
	if (StatusSelectedIndex > 1)
		$("#KitStar").show();
	else
	    $("#KitStar").hide();
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

.DelTable TBODY TD{
	BORDER-TOP: gray thin solid;
}


</STYLE>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<%
	dim cn
	dim rs
	dim cm
	dim p
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim strPilotStatus
	dim strDate
	dim strDeliverable
	dim strID
	dim strPartNumber
	dim strProduct
	dim strVendor
	dim blnAdmin
	dim blnFound
	dim strStatusList
	dim strStatusText
	dim strQualStatus
	dim strComments
	dim strHW
	dim strFW
	dim strRev
	dim strModel
	dim strStatusSelected
	dim CurrentUserEmail
	dim strQCompleteSubject
	dim strQCompleteBody
	dim strFailSubject
	dim strFailBody
	dim strPMEmail
	dim strAccessoryStatusID
	dim strDevEmail
	'dim strQCompleteCount
	dim strDevCenter
	dim strRows
	dim blnLeveraged
	dim blnShowDateinLinkedStatus
	dim strKitDescription
	
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("email") & ""
	else
		CurrentUserID = 0
		CurrentUserEmail = "max.yu@hp.com"
	end if
	rs.Close
	
	if currentuserid=31 or currentuserid=8 then
		blnadmin=true
	else
		blnadmin=false
	end if
	
	strID=""
	strPilotStatus = ""
	strDate = ""
	strVendor=""
	strPartNumber=""
	strDeliverable = ""
	strProduct= ""
	strStatusList = ""
	strDevStatus = ""
	strComments = ""
	strHW = "&nbsp;"
	strFW = "&nbsp;"
	strRev = "&nbsp;"
	strModel = "&nbsp;"
	blnFound = false
	strStatusSelected = ""
	strAccessoryStatusID=""
	strFailSubject = ""
	strFailBody = ""
	strPMEmail = ""
	strDevEmail = ""
	'strQCompleteCount = ""
	strDevCenter = ""
	strRows = ""
	blnLeveraged = false
	blnShowDateinLinkedStatus = false
	strKitDescription = ""
	
	
	if request("ProdID") = "" or request("VersionID") = "" then
		Response.Write "Not enough information supplied to process your request."
	else
		rs.Open "spGetAccessoryStatus " & clng(request("ProdID")) & "," & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strNumber = ""
		else
			blnFound = true
			
			strID = rs("ID") & ""
			strProduct = rs("Product") & ""
			strDevCenter = trim(rs("DevCenter") & "")
			strDeliverable = rs("DeliverableName") & ""
			strPilotStatus = rs("PilotStatus") & ""
			strAccessoryStatusID = trim(rs("AccessoryStatusID") & "")
			if strAccessoryStatusID = ""then
				strAccessoryStatusID = "0"
			end if
			strPartNumber = rs("PartNumber") & ""
			strModel = rs("ModelNumber") & "&nbsp;"
			strKitNumber = rs("KitNumber") & ""
			strKitDescription = rs("KitDescription") & ""
			strDate = rs("AccessoryDate") & ""
			if isdate(strDate) then
				strDate = formatdatetime(strDate,vbshortdate)
			end if
			strVendor=rs("Vendor") & ""
			strComments = rs("AccessoryNotes") & ""
			strHW = rs("Version") & "&nbsp;"
			strFW = rs("Revision") & "&nbsp;"
			strRev = rs("Pass") & "&nbsp;"
			if not isnull(rs("accessoryleveraged")) then
				blnLeveraged = rs("accessoryleveraged") 
			end if
			strQualStatus = trim(rs("TestStatus") & "")
			if strQualStatus = "" then
				strQualStatus = "Not Used"
			elseif strQualStatus = "Date" then
				blnShowDateinLinkedStatus = true
				strQualStatus = rs("TestDate") & "&nbsp;"
			end if
			if strPilotStatus = "" then
				strPilotStatus = "Not Required"
			elseif strPilotStatus = "Date" then
				strPilotStatus = rs("PilotDate") & "&nbsp;"
			end if
		end if
		rs.Close
		
		strStatusSelected = ""
		strDateShow = ""
		strDateRequired = ""
		strCommentsRequired = ""
		rs.Open "spListAccessoryStatus",cn,adOpenForwardOnly
		do while not rs.EOF
			strStatusText = rs("Name") & ""
			if trim(rs("ID")) = trim(strAccessoryStatusID) then
				strStatusList = strStatusList & "<option selected value=""" & rs("ID") & """>" & strStatusText & "</option>"
				strStatusSelected = strStatusText
				if rs("ID") = 2 and blnShowDateinLinkedStatus then
					strStatusSelected = strStatusSelected & " (" & strDate & ")"
				end if
			else	
				strStatusList = strStatusList & "<option value=""" & rs("ID") & """>" & strStatusText & "</option>"
			end if
			if trim(rs("DateField")) & "" = "2" then
				strDateShow = strDateShow & "," & trim(rs("ID"))
			elseif	trim(rs("DateField")) & "" = "1" then
				strDateShow = strDateShow & "," & trim(rs("ID"))
				strDateRequired = strDateRequired & "," & trim(rs("ID"))
			end if
			if rs("CommentsRequired") then
				strCommentsRequired = strCommentsRequired & "," & trim(rs("ID"))
			end if
			
			rs.movenext
		loop
		
		rs.Close
		
		if blnLeveraged then
			strStatusList = strStatusList & "<option selected value=""-1"">Link to Commodity Status</option>"
		else
			strStatusList = strStatusList & "<option value=""-1"">Link to Commodity Status</option>"
		end if
		
		strDateShow = strDateShow & ","
		strDateRequired = strDateRequired & ","
		strCommentsRequired = strCommentsRequired & ","



		strDevEmail = ""
'		rs.open "spGetDeliverableDeveloper " & clng(request("VersionID")),cn,adOpenStatic
'		if not (rs.EOF and rs.BOF) then
'			if trim(rs("DeveloperEmail") & "") <> "" then
'				strDevEmail = strDevEmail & ";" & rs("DeveloperEmail")
'			end if
'			if trim(rs("DevManagerEmail") & "") <> "" then
'				strDevEmail = strDevEmail & ";" & rs("DevManagerEmail")
'			end if
'		end if
'		rs.Close
'		if trim(strDevEmail) = "" then
			strDevEmail = "max.yu@hp.com"
'		else
'			strDevEmail = mid(strDevEmail,2)
'		end if


		strPMEmail = ""
'		rs.Open "spListCommodityPMs4Version " & clng(request("VersionID")),cn,adOpenStatic
'		do while not rs.EOF
'			strPMEmail = strPMEmail & ";" & rs("Email")
'			rs.MoveNext
'		loop
'		rs.Close
'		if trim(strPMEmail) = "" then
			strPMEmail = "max.yu@hp.com"
'		else
'			strPMEmail = mid(strPMEmail,2)
'		end if
		
		
	end if
	
	if 	blnFound then
%>



<font face=verdana size=2><b>
<label ID="lblTitle"><%=strVendor%>&nbsp;<%=strDeliverable%> on <%=strProduct%></label></b></font>
<% 

	dim strVersion
	
	strVersion = ""
	if trim(strHW) <> "&nbsp;" then
		strVersion =  strHW
	end if
	if trim(strFW) <> "&nbsp;" then
		strVersion = strVersion & "," & strFW
	end if
	if trim(strRev) <> "&nbsp;" then
		strVersion = strVersion & "," & strRev
	end if




%>

<form id="frmStatus" method="post" action="AccessoryStatusSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">

<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td valign=top width=120 nowrap><b>Qual.&nbsp;Status:</b>&nbsp;</td>
		<td>
			<%=strQualStatus%>
		</td>
		<td valign=top width=60 nowrap><b>HW&nbsp;Ver:</b>&nbsp;</td>
		<td><%=strHW%></td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Pilot Status:</b>&nbsp;</td>
		<td>
			<%=strPilotStatus%>
		</td>
		<td valign=top width=60 nowrap><b>FW&nbsp;Ver:</b>&nbsp;</td>
		<td><%=strFW%></td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Model:</b>&nbsp;</td>
		<td>
			<%=strModel%>
		</td>
		<td valign=top width=60 nowrap><b>Rev:</b>&nbsp;</td>
		<td><%=strRev%></td>
	</tr>
	<tr><%
			dim strRequireKit
			if trim(strAccessoryStatusID) > 1 then 	
				strRequireKit = ""		
			else
				strRequireKit = "none"			
			end if
		%>
		<td valign=top width=60 nowrap><b>Kit&nbsp;Number:</b>&nbsp;<font style="Display:<%=strRequireKit%>" color="red" size="1" ID=KitStar>*</font>&nbsp;
		</td>
		<td><INPUT type="text" id=txtKitNumber name=txtKitNumber style="width:200" value="<%=strKitNumber%>" maxlength=20></td>
		<td valign=top width=120 nowrap><b>Part&nbsp;Number:</b>&nbsp;</td>
		<td>
			<%=strPartNumber%>&nbsp;
		</td>
	</tr>



	<tr>
		<td valign=top width=120 nowrap><b>Accessory&nbsp;Status:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan=3>
			<SELECT style="width:200" id=cboStatus name=cboStatus language=javascript onchange="cboStatus_onchange();">
			<%=strStatusList%>
			</SELECT>
			<%
				if (trim(strAccessoryStatusID)="0" or trim(strAccessoryStatusID)="") and strQualStatus="Not Used" then
					strSupportDisplay = ""
				else
					strSupportDisplay = "none"
				end if
				
				if blnLeveraged then
					strStatusDisplay = ""
				else
					strStatusDisplay = "none"
				end if
			%>
			<span id=SupportRow style=display:<%=strSupportDisplay%>><INPUT type="checkbox" id=chkDelete name=chkDelete>Completely remove support</span><BR>
			<span id=StatusRow style=display:<%=strStatusDisplay%>><b>Current Status:</b> <%=strStatusSelected%></span>
		</td>
	</tr>
	
	<%if instr(strDateShow,"," & trim(strAccessoryStatusID) & ",") > 0 and not blnLeveraged then %>
		<tr ID=DateRow>
	<%else%>
		<tr ID=DateRow style="display:none">
	<%end if%>
		<td valign=top width=120 nowrap><b>Date:</b>
		<%if instr(strDateRequired,"," & trim(strAccessoryStatusID) & ",") > 0 then %>
			&nbsp;<font color="red" size="1" ID=DateStar>*</font>&nbsp;
		<%else%>
			&nbsp;<font style="Display:none" color="red" size="1" ID=DateStar>*</font>&nbsp;
		<%end if%>
		</td>
		<td colspan=3>
			<INPUT type="text" id=txtAccessoryDate name=txtAccessoryDate value="<%=strDate%>">&nbsp;<a href="javascript: cmdDate_onclick()"><img ID="picTarget" SRC="../../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a>
		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Comments:</b>
		<%if instr(strCommentsRequired,"," & trim(strAccessoryStatusID) & ",") > 0 then %>
			<font color="red" size="1" ID=CommentStar>*</font>&nbsp;
		<%else%>
			<font style="Display:none" color="red" size="1" ID=CommentStar>*</font>&nbsp;
		<%end if%>
		
		</td>
		<td colspan=7>	<TEXTAREA rows=3 style="WIDTH:100%" id=txtComments name=txtComments><%=strComments%></TEXTAREA>
			
			
		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Kit Description:</b></td>
		<td colspan=7>	
		<INPUT type="text" style="WIDTH:100%" id=txtKitDescription name=txtKitDescription value="<%=strKitDescription%>" maxlength=120>
		</td>
	</tr>
</table>


<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=strID%>">
<INPUT style="Display:none" type="text" id=txtUserID name=txtUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtUserName name=txtUserName value="<%=CurrentDomain & "_" & CurrentUser%>">
<INPUT type="hidden" id=txtUserEmail name=txtUserEmail value="<%=CurrentUserEmail%>">
<INPUT type="hidden" id=txtQualStatus name=txtQualStatus value="<%=strQualStatus%>">
<INPUT type="hidden" id=txtStatusText name=txtStatusText value="">
<INPUT type="hidden" id=txtVendor name=txtVendor value="<%=strVendor%>">
<INPUT type="hidden" id=txtDeliverable name=txtDeliverable value="<%=strDeliverable%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtProdID name=txtProdID value="<%=request("ProdID")%>">
<INPUT type="hidden" id=txtVersion name=txtVersion value="<%=strVersion%>">
<INPUT type="hidden" id=txtModel name=txtModel value="<%=strModel%>">
<INPUT type="hidden" id=txtPartNumber name=txtPartNumber value="<%=strPartNumber%>">
<INPUT type="hidden" id=txtStatusLoaded name=txtStatusLoaded value="<%=trim(strPilotStatus)%>">
<INPUT type="hidden" id=txtProduct name=txtProduct value="<%=strProduct%>">
<INPUT type="hidden" id=txtPMEmail name=txtPMEmail value="<%=strPMEmail%>">
<INPUT type="hidden" id=txtDevEmail name=txtDevEmail value="<%=strDevEmail%>">
<input type="hidden" id=txtTodayPageSection name=txtTodayPageSection value="<%=Request("TodayPageSection")%>">
<input type="hidden" id="txtRowID" name="txtRowID" value="<%=Request("RowID")%>" />
</form> 
<%end if

	cn.Close
	set cn = nothing
	set rs = nothing


%>
<INPUT type="hidden" id=txtDateShow name=txtDateShow value="<%=strDateShow%>">
<INPUT type="hidden" id=txtDateRequired name=txtDateRequired value="<%=strDateRequired%>">
<INPUT type="hidden" id=txtCommentsRequired name=txtCommentsRequired value="<%=strCommentsRequired%>">
 <INPUT style="Display:none" type="hidden" id=hdnApp name=hdnApp value="<%=Request.form("app")%>">
</BODY>
</HTML>


