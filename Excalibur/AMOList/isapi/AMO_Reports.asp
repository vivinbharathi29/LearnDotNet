<%@ Language=VBScript %>
<% Option Explicit 
	
    Response.Buffer = False
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"

    Server.ScriptTimeout = 6000
%>
<!------------------------------------------------------------------- 
'Description: AMO DATA
'----------------------------------------------------------------- //-->    
<!-- #include file="../includes/incConstants.inc" -->
<!-- #include file="../data/oDataConnection.asp" -->
<!-- #include file="../data/oDataAMO.asp" -->
<!-- #include file="../data/oDataAVL.asp" -->
<!-- #include file="../data/oDataPermission.asp" -->
<!-- #include file="../data/oDataGeneral.asp" -->
<!-- #include file="../data/oDataMOLCategory.asp" -->
<!-- #include file="../data/oDataWebCategory.asp" -->
<!-- #include file="../library/includes/MOL_CategoryRs.inc" -->
<!-- #include file="../library/includes/CategoryRs.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO PERMISSIONS 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Roles.inc" -->
<!-- #include file="../library/includes/cookies.inc" -->
<!-- #include file="../library/includes/SessionValidation.inc" -->
<!-- #include file="../library/includes/ErrHandler.inc" -->

<!------------------------------------------------------------------- 
'Description: AMO HTML 
'----------------------------------------------------------------- //--> 
<!-- #include file="../library/includes/Grid.inc" -->
<!-- #include file="../library/includes/ListboxRs.inc" -->
<!-- #include file="../library/includes/DualListBoxRs.inc" -->
<!-- #include file="../library/includes/lib_debug.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->

<!------------------------------------------------------------------- 
'Description: Initialize AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/openDBConnection.asp" -->
<%
Call ValidateSession

dim sHeader, sHelpFile, sDescHTML, sErr, sSaved, sSchURLHTML, sIURURLHTML
dim sScheduleURL, sIURURL, sObj, oSvr, oErr, nBusSelected
dim nNumRequest, nNumTotalRequest, nMode, nStatusID, nCol, nNumCol, nUCLink, nCLink, nUserCommitID, nMol
dim irowcount
dim bScopeChangeUpdate, bPMUpdate
dim oRsSel, oRsCommit
dim lbxOwnerHTML
dim aCfg(), aUCLink(3), aCLink(0, 4)
dim RsAMOPublishs
dim sBusSegIDs1, sBusSegHTML1, oRsBusSeg, sBusSegIDs2, sBusSegHTML2
dim sFilter, strEOLDate, strPubEOLDate, nMolCheck, rsBS, rsBSSelected
dim bAMOCreate, bAMOView, bAMOUpdate, bAMODelete
dim arrColors
arrColors = Array("#CCCCCC", "white")

GetRights2 Application("AMO_Permission"), bAMOCreate, bAMOView, bAMOUpdate, bAMODelete

sHeader = "After Market Options List"
sHelpFile = "../help/HELP_AMO_Reports.asp"


'business segment filter 1
'get the cookie. If we didn't get it default it
if Request.Form("chkBusSeg1") = "" then
	'get the cookie. If we didn't get it default it
	sBusSegIDs1 = GetDBCookie( "AMO chkBusSeg1")
else
	sBusSegIDs1 = Request.Form("chkBusSeg1")
end if
	'store the cookie
Call SaveDBCookie( "AMO chkBusSeg1", sBusSegIDs1 )	


nMol = Request.QueryString("nMol")
if nMol = 1 Then
	nMolCheck = "checked"
else
	nMolCheck = ""
end if


sBusSegHTML1 = ""
nBusSelected = ""

'set oErr = GetMOLCategory(oRsBusSeg, 28)
set oRsBusSeg = GetMOLCategory(34)	
if oRsBusSeg is Nothing then
	Response.Write("Recordset error: oRsBusSeg")
	Response.End()
end if


sBusSegHTML1 =GetBusSegHTML (oRsBusSeg, sBusSegIDs1, 1)

if Request.Form("txtEOLDate") = "" then
	'get the cookie. If we didn't get it default it
	strEOLDate = GetDBCookie( "AMO txtEOLDate")
	if trim(strEOLDate) = "" then
		'set default day to 1 month prior
		strEOLDate = dateAdd("m", -1, date)
	end if
else
	strEOLDate = Request.Form("txtEOLDate")
end if
'store the cookie
Call SaveDBCookie( "AMO txtEOLDate", strEOLDate )


if Request.Form("txtPubEOLDate") = "" then
	'get the cookie. If we didn't get it default it
	strPubEOLDate = GetDBCookie( "AMO txtPubEOLDate")
	if trim(strPubEOLDate) = "" then
		'set default day to 1 month prior
		strPubEOLDate = dateAdd("m", -1, date)
	end if
else
	strPubEOLDate = Request.Form("txtPubEOLDate")
end if
'store the cookie
Call SaveDBCookie( "AMO txtPubEOLDate", strPubEOLDate )

'business segment filter 2
'get the cookie. If we didn't get it default it
if Request.Form("chkBusSeg2") = "" then
	'get the cookie. If we didn't get it default it
	sBusSegIDs2 = GetDBCookie( "AMO chkBusSeg2")
else
	sBusSegIDs2 = Request.Form("chkBusSeg2")
end if
	'store the cookie
Call SaveDBCookie( "AMO chkBusSeg2", sBusSegIDs2 )	

sBusSegHTML2 =GetBusSegHTML (oRsBusSeg, sBusSegIDs2, 2)

'Get only the business segments the user is in
'GetCategory rsBS, 44, session("USERID")

set rsBSSelected = Nothing

nBusSelected = GetDBCookie( "AMO BusSelected")
	

'call sbusseg function
nNumRequest = 0
'set oSvr = server.CreateObject("JF_S_AMO.ISAMO")
set oSvr = New ISAMO
	
set RsAMOPublishs = oSvr.ViewAll_AMO_SCM_PublishList(Application("REPOSITORY"), sBusSegIDs2)
if RsAMOPublishs is nothing then
	sErr = "Missing required parameters.  Unable to complete your request."
	Response.Write(sErr)
	Response.End()
end if
%>
<!DOCTYPE html>
<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta content="<%=sHeader%>" name="description">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Pragma" content="no-cache"> 
<meta http-equiv="Expires" content="-1">
<title><%=sHeader%> - Reports</title>
<meta http-equiv="X-UA-Compatible" content="IE=EDGE,chrome=1">
<link rel="stylesheet" type="text/css" href="../style/cupertino/jquery-ui.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/simple.css">
<link rel="stylesheet" type="text/css" href="../library/stylesheets/popup.css">
<link rel="stylesheet" type="text/css" href="../style/wizard%20style.css">
<link rel="stylesheet" type="text/css" href="../style/amo.css" />
<script type="text/javascript" src="../scripts/jquery-1.10.2.min.js" ></script>
<script type="text/javascript" src="../scripts/jquery-ui-1.11.4.min.js" ></script>
<script type="text/javascript" src="../library/scripts/formChek.js"></script>
<script type="text/javascript" src="../library/scripts/calendar.js"></script>
<script type="text/javascript" src="../scripts/amo.js"></script>
<SCRIPT ID=clientEventHandlersJS type="text/javascript" LANGUAGE=javascript>
<!--
function ValidateData() {
	var i;
	var nMol, slink;
	var sRASDiscontinueDate, sDivisionIds;
	sRASDiscontinueDate = thisform.txtPubEOLDate.value;
	
	if (!checkDate (thisform.txtPubEOLDate, "Publish options with End of Manufacturing (EM)", true))
		return false;
		
	SelectAll(thisform.lbxSelectedDivision);
	if (thisform.lbxSelectedDivision.value == "") {
		alert("Please select at least one business segment before proceeding");
		return false;
	} 	

	sDivisionIds = "";
	for (i=0; i < thisform.lbxSelectedDivision.options.length; i++)
		sDivisionIds = sDivisionIds + "," + thisform.lbxSelectedDivision[i].value;	
	if (sDivisionIds != "")
		sDivisionIds=sDivisionIds.slice(1);
		
	slink = "AMO_ValidateData.asp?BusSeg=" + sDivisionIds + "&RASDiscontinueDate=" + sRASDiscontinueDate;
	
	thisform.action = slink;
	thisform.target = ''
	thisform.submit();
}

function PublishSCM() {
	var i;
	var nMol, slink;
	var sEOLDate, sDivisionIds;
	sEOLDate = thisform.txtPubEOLDate.value;
	
	if (!checkDate (thisform.txtPubEOLDate, "Publish options with End of Manufacturing (EM)", true))
		return false;
		
	SelectAll(thisform.lbxSelectedDivision);
	if (thisform.lbxSelectedDivision.value == "") {
		alert("Please select at least one business segment before proceeding");
		return false;
	} 	

	sDivisionIds = "";
	for (i=0; i < thisform.lbxSelectedDivision.options.length; i++)
		sDivisionIds = sDivisionIds + "," + thisform.lbxSelectedDivision[i].value;	
	if (sDivisionIds != "")
		sDivisionIds=sDivisionIds.slice(1);
		
	if(thisform.chkAlternatePub.checked) 
	    slink = "/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?nPublish=1&BusSeg=" + sDivisionIds + "&nMol=1&bsnapshot=1&txtEOLDate=" + sEOLDate + "&nFormat=1";
	else
	    slink = "/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?nPublish=1&BusSeg=" + sDivisionIds + "&nMol=1&bsnapshot=1&txtEOLDate=" + sEOLDate + "&nFormat=0";
	
	thisform.action = slink;
	thisform.target = ''
	thisform.submit();
}

function ShowSCMExcel() {
	var i;
	var strTemp, nMol;
	strTemp = "";
	var sEOLDate;
	sEOLDate = thisform.txtEOLDate.value;
	if (!checkDate (thisform.txtEOLDate, "Show options with End of Manufacturing (EM)", true))
		return false;

	var coll = document.getElementsByName("chkBusSeg1");
	for (i=0;i< coll.length; i++) {
		if (coll[i].checked){
			strTemp += coll[i].value + ",";
		}
	} 
	if (strTemp.length > 0)
		strTemp = strTemp.substring(0,strTemp.length-1)
	
	if (document.getElementById("chkhfMol").checked)
	   nMol = 1;
	else
	   nMol = 0;
	   
	if(thisform.chkAlternateView.checked) 
	    window.open("/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?nMol=" + nMol + "&BusSeg=" + strTemp + "&txtEOLDate=" + sEOLDate + "&nFormat=1")
	else
	    window.open("/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?nMol=" + nMol + "&BusSeg=" + strTemp + "&txtEOLDate=" + sEOLDate + "&nFormat=0")
}

function CompareSCM() {
    var i;
    var strTemp;
    var NumChecked;
    var strBusSegIDs;
    var strSCMIDs;
    var coll;
    var sEOLDate;

    NumChecked =0;
    coll = document.getElementsByName("chkCompareSCM");
    for (i=0;i< coll.length; i++) {
        if (coll[i].checked) {
            NumChecked += 1;
        }
    } 
    if (NumChecked != 2) {
        alert("Please select two published SCMs for comparison.");
        return false;
    } else {
        sEOLDate = thisform.txtEOLDate.value;

        coll = document.getElementsByName("chkCompareSCM");
        for (i=0;i< coll.length; i++) {
            if (coll[i].checked) {
                NumChecked += 1;
                strTemp += coll[i].value + ",";
            }
        } 
        if (strTemp.length > 0){
            strTemp = strTemp.substring(0,strTemp.length-1);
        }
        strSCMIDs = strTemp.replace("undefined", "");

        var coll = document.getElementsByName("chkBusSeg1");
        for (i=0;i< coll.length; i++) {
            if (coll[i].checked){
                strTemp += coll[i].value + ",";
            }
        } 
        if (strTemp.length > 0){
            strTemp = strTemp.substring(0,strTemp.length-1);
        }
        strBusSegIDs = strTemp.replace("undefined", "");

        window.open('/IPulsar/Reports/AMO/AMO_SCM_Comparison.aspx?SCMID='+strSCMIDs+'&BusSeg='+strBusSegIDs+'&CompareCurrent=0&EOLDate='+sEOLDate+'');
        //thisform.action = "AMO_SCM_Comparison.asp";
        //thisform.target = '_blank'
        //thisform.submit();
        return true;
    }
}

function CompareSCMCurrent() {
    var i;
    var strTemp;
    var NumChecked;
    var strBusSegIDs;
    var strSCMIDs;
    var coll;
    var sEOLDate;
    
    NumChecked =0;
    coll = document.getElementsByName("chkCompareSCM");
    for (i=0;i< coll.length; i++) {
        if (coll[i].checked) {
            NumChecked += 1;
        }
    } 

    if (NumChecked != 1) {
        alert("Please select one published SCM for comparison with current SCM.");
        return false;
    } else {
        sEOLDate = thisform.txtEOLDate.value;

        coll = document.getElementsByName("chkCompareSCM");
        for (i=0;i< coll.length; i++) {
            if (coll[i].checked) {
                NumChecked += 1;
                strTemp += coll[i].value + ",";
            }
        } 
        if (strTemp.length > 0){
            strTemp = strTemp.substring(0,strTemp.length-1);
        }
        strSCMIDs = strTemp.replace("undefined", "");

        var coll = document.getElementsByName("chkBusSeg1");
        for (i=0;i< coll.length; i++) {
            if (coll[i].checked){
                strTemp += coll[i].value + ",";
            }
        } 
        if (strTemp.length > 0){
            strTemp = strTemp.substring(0,strTemp.length-1);
        }
        strBusSegIDs = strTemp.replace("undefined", "");

        window.open('/IPulsar/Reports/AMO/AMO_SCM_Comparison.aspx?SCMID='+strSCMIDs+'&BusSeg='+strBusSegIDs+'&CompareCurrent=1&EOLDate='+sEOLDate+'');
        //thisform.action = "AMO_SCM_Comparison.asp?CompareCurrent=1";
        //thisform.target = '_blank'
        //thisform.submit();
        return true;
    }
}
function lbxGo_onchange() {
	thisform.action = "AMO_Reports.asp";
	thisform.target = ''
	thisform.submit();
}

function window_onload() {
<%if len(request.querystring("SCMID")) > 0 then%>
	<%	if request.querystring("Norecords") ="1" Then%>
		alert("No After Market Option found that matches filter for Publishing")	
	<%else%>
			<%if request.querystring("nFormat") ="1" Then%>
				window.open("/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?BusSeg=<%=nBusSelected%>&SCMID=<%=request.querystring("SCMID")%>&PublishDate=<%=date()%>&nFormat=1")
			<%else%>
				window.open("/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?BusSeg=<%=nBusSelected%>&SCMID=<%=request.querystring("SCMID")%>&PublishDate=<%=date()%>&nFormat=0")
	<%end if end if%>
<%end if %>
}
//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript  onload = "return window_onload();">
<h1 class="page-title"><%=sHeader%></h1>
<FORM name=thisform method=post>
<%
'insert the header, global navigation and overview links

Response.Write ""
WriteTabs "Reports"
%>

<TABLE border=0 cellPadding=1 cellSpacing=2 width="100%">
	<tr>
		<td colspan=2><h2>View SCM Report</h2></td></tr>	
	<tr>
		<td width=20%>Business Segment</td>
		<td width=80%><%=sBusSegHTML1%></td>
	</tr>
	<tr>
		<td><INPUT type='checkbox' id=chkhfMol NAME=chkhfMol <%=nMolCheck%>>Include Hide from PRL</td>
		<td colspan=2>Show options with End of Manufacturing (EM)(<i>base unit</i>) on or after &nbsp;<input type=text maxlength=10 size=10 id=txtEOLDate NAME=txtEOLDate value=<%= strEOLDate %> class="filter-dateselection">
			(MM/DD/YYYY)</td>
	</tr>
	
	<tr>
		<td colspan=2><b>SCM Reports are created in Excel format, Microsoft Office 2007 or higher is needed to run the reports.</b></td></tr>
	<tr>
		<td><br><Input type=button name=btnSCMReport ID=btnSCMReport style="width:170" value="View SCM Report"  onclick="return ShowSCMExcel();"></td>
		<td><br><INPUT type='checkbox' id=chkAlternateView NAME=chkAlternateView>Use Alternate Format</td>
		</tr>
		
	<%if bAMOCreate or bAMOUpdate then%>	
	
	<TR>
		<TD colspan=2 ><HR></TD></TR>
		
	<tr>
		<td colspan=2><h2>Publish SCM Report</h2></td>
	</tr>	
	
	<tr>
		<td width=20%>User's Business Segments</td>
		<td width=80%>
		<%	
			DualListboxRs_GetHTML6_Write oRsBusSeg, "SegmentName", "BusinessSegmentID", rsBSSelected, _
			"SegmentName", "BusinessSegmentID", true, true, nBusSelected, "Available", "Selected", _
			"Division", true, 130, 250, false, true, false, 350, 13		
		%>
		
		
		
		<br></td>
	</tr>
	
	<tr>
		<td colspan=2>Publish options with End of Manufacturing (EM)(<i>base unit</i>) on or after &nbsp;<input type=text maxlength=10 size=10 id=txtPubEOLDate NAME=txtPubEOLDate value=<%= strPubEOLDate %> class="filter-dateselection">
			(MM/DD/YYYY)</td>
	</tr>
	<tr>
		<td colspan=2><b>SCM Reports are created in Excel format, Microsoft Office 2007 or higher is needed to run the reports.</b></td></tr>
	
	<tr><td colspan=3>
		<table>
		<tr>
			<td><br><Input type=button name=btnValidate ID=btnValidate style="width:170" value="Validate Data"  onclick="return ValidateData();"></td>
			<td><br><Input type=button name=btnPublish ID=btnPublish style="width:170" value="Publish SCM Report"  onclick="return PublishSCM();"></td>
			<td><br><INPUT type='checkbox' id=chkAlternatePub NAME=chkAlternatePub>Use Alternate Format</td>
		</tr>
		</table>
	</td>
	</tr>
	
	<%end if %>
		
	<TR>
		<TD colspan=2 ><HR></TD></TR>
		
	<tr>
		<td colspan=2><h2>Published After Market Option SCM Reports</h2></td></tr>	
	<tr>
		<td width=20%>Business Segment</td>
		<td width=80%><%=sBusSegHTML2%></td>
	</tr>
	<tr>
		<td colspan=2 align=left><INPUT type='button' value="Filter List" id=btnRefresh NAME=btnRefresh LANGUAGE=javascript onclick="return lbxGo_onchange();"></td></tr>
</table>



<%if RsAMOPublishs.recordcount > 0 then%>
<TABLE id="tblAMOList" border="1" CELLSPACING="0" CELLPADDING="3" width="60%">
	<%	irowcount=0 %>
	<tr class="tblrow-pulsar"> 
		<th align=center>Select</th>
		<th align=center>Date Published</th>
		<th align=center>Publisher</th>
		<th align=center>Format</th>
		<th align=center>Business Segment</th>
	</tr>
	<%do while not RsAMOPublishs.eof   %>
		<%irowcount = irowcount + 1 %>
		<tr  bgcolor="<%= arrColors(irowcount mod 2) %>" > 
			<td align=center><input type=checkbox name =chkCompareSCM ID =chkCompareSCM  value =<%=RsAMOPublishs("SCMID")%> ></td>
			<% if RsAMOPublishs("FormatType") = "Standard" then%>
				<td align=center><A href= "<%="/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?BusSeg=" & GetDBCookie( "AMO chkBusSeg2") & "&SCMID=" & RsAMOPublishs("SCMID")& "&PublishDate=" & cdate(RsAMOPublishs("Datecreated")) & "&nFormat=0"%>" target=new ><%=RsAMOPublishs("Timecreated")%></A>	</td>
			<%else%>
				<td align=center><A href= "<%="/IPulsar/Reports/AMO/AMO_SCM_Excel_Alter_Report.aspx?BusSeg=" & RsAMOPublishs("CompatibilityDivisionIds") & "&SCMID=" & RsAMOPublishs("SCMID")&"&PublishDate=" & cdate(RsAMOPublishs("Datecreated")) & "&nFormat=1"%>" target=new ><%=RsAMOPublishs("Timecreated")%></A>	</td>
			<%end if%>
			<td align=center><%=RsAMOPublishs("Creator")   %>	</td>
			<td align=center><%=RsAMOPublishs("FormatType")%>	</td>
			<td align=center><%=RsAMOPublishs("BusSegments") %>	</td>
		</tr>	
		<%RsAMOPublishs.movenext%>
	<%loop%>			
</table>
<%end if %>		

<%if RsAMOPublishs.recordcount > 0 then%>
	<p>	
	<%if RsAMOPublishs.recordcount > 1 then%>
		<input style="width:150" type=button name =btnCompare ID=btnCompare value = "Compare Published"  onclick= "return CompareSCM();">
	<%end if %>
	<input style="width:150" type=button name =btnCompare ID=btnCompare value = "Compare with Current"  onclick= "return CompareSCMCurrent();">
	</p>
<% else %>
	<p>No Published SCM Reports found for the above filter.</p>
<%end if %>	
		
<%

%>
    <input type="hidden" id="inpUserID" value="<%=Session("AMOUserID")%>" />
</FORM>
</BODY>
</HTML>
<script type="text/javascript">
    //*****************************************************************
    //Description:  OnLoad, on page load instantiate functions
    //*****************************************************************
    $(window).load(function () {
        load_datePicker();
    });
</script>
<%
function GetBusSegHTML (oRsBusSeg, byref sBusSegIDs2, sectionnumber)	
	dim sBusSegHTML2
	if not oRsBusSeg is nothing then
		if oRsBusSeg.RecordCount > 0 then
			if sBusSegIDs2 = "" then
				'make every one checked
					oRsBusSeg.MoveFirst
					do until oRsBusSeg.EOF
						sBusSegIDs2 = sBusSegIDs2 & oRsBusSeg("BusinessSegmentID").Value & ","
						oRsBusSeg.MoveNext
					loop
					if right(sBusSegIDs2, 1) = "," then
						sBusSegIDs2 = left(sBusSegIDs2, len(sBusSegIDs2)-1)
					end if
					'store the cookie
					Call SaveDBCookie( "AMO chkBusSeg" & sectionnumber, sBusSegIDs2 )
			end if
			oRsBusSeg.MoveFirst
			Do until oRsBusSeg.EOF
				if sBusSegHTML2 <> "" then 
					sBusSegHTML2 = sBusSegHTML2 & "&nbsp;"
				end if
				if instr(1, sBusSegIDs2, oRsBusSeg("BusinessSegmentID").Value) > 0 then 
					sBusSegHTML2 = sBusSegHTML2 & "<INPUT type='checkbox' id=chkBusSeg" & sectionnumber & " NAME=chkBusSeg" & sectionnumber & "  value=" & oRsBusSeg("BusinessSegmentID").Value & " checked "  & ">" & oRsBusSeg("SegmentName").Value
				else
					sBusSegHTML2 = sBusSegHTML2 & "<INPUT type='checkbox' id=chkBusSeg" & sectionnumber & "  NAME=chkBusSeg" & sectionnumber & "  value=" & oRsBusSeg("BusinessSegmentID").Value & ">" & oRsBusSeg("SegmentName").Value
				end if
				oRsBusSeg.MoveNext
			Loop
		end if
	end if
	
	GetBusSegHTML = sBusSegHTML2
end function	
%>
<!------------------------------------------------------------------- 
'Description: Close AMO DB Connection
'----------------------------------------------------------------- //-->  
<!-- #include file="../data/closeDBConnection.asp" -->