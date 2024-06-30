<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="includes/bundleConfig.inc" -->
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.close();
    }
}

function cmdOK_onclick() {
	var ID = window.frmMain.txtID.value;
	var SectionList = "";
//	var	NewLeft = (screen.width - 655)/2;
//	var NewTop = (screen.height - 650)/2;

	if (window.frmMain.chkChangeOpened.checked)
		SectionList = SectionList + ",12"
	if (window.frmMain.chkChangeClosed.checked)
		SectionList = SectionList + ",13"
	if (window.frmMain.chkBcrChangeOpened.checked)
		SectionList = SectionList + ",20"
	if (window.frmMain.chkBcrChangeClosed.checked)
		SectionList = SectionList + ",21"
	if (window.frmMain.chkDeliverables.checked)
		SectionList = SectionList + ",6"
	if (window.frmMain.chkOTSOpened.checked)
		SectionList = SectionList + ",14"
	if (window.frmMain.chkOTSClosed.checked)
		SectionList = SectionList + ",15"
	if (window.frmMain.chkSchedule.checked)
		SectionList = SectionList + ",17"
	if (window.frmMain.chkAgency.checked)
		SectionList = SectionList + ",16"
	if (window.frmMain.chkCountry.checked)
		SectionList = SectionList + ",18"
	if (window.frmMain.chkLocalization.checked)
		SectionList = SectionList + ",19"

	if (SectionList == "")
	    alert("No report sections selected");
	else {
	    SectionList = SectionList.substr(1);
	    //MainBody.innerHTML = "Processing.  Please Wait...";
	    //window.open("ProductStatus.asp?ID=" + ID + "&ReportDays=7&ReportTitle= - This Week&Sections=" + SectionList,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes"); 
	    window.open("ProductStatus.asp?ID=" + ID + "&ReportDays=7&ReportTitle= - Change Log&Sections=" + SectionList + "&StartDt=" + window.frmMain.txtStartDt.value + "&EndDt=" + window.frmMain.txtEndDt.value, "_blank", "Width=655,Height=500,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes,location=yes");
	    var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
	    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	        // For Closing current popup if Called from pulsarplus
	        parent.window.parent.closeExternalPopup();
	    }
	    else if (parent.window.parent.document.getElementById('modal_dialog')) {
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.close();
	    }
	}
}

function chkChange_onclick() {
	window.frmMain.chkChangeOpened.checked = window.frmMain.chkChange.checked;
	window.frmMain.chkChangeClosed.checked = window.frmMain.chkChange.checked;
}

function chkBcrChange_onclick() {
	window.frmMain.chkBcrChangeOpened.checked = window.frmMain.chkBcrChange.checked;
	window.frmMain.chkBcrChangeClosed.checked = window.frmMain.chkBcrChange.checked;
}

function chkOTS_onclick() {
	window.frmMain.chkOTSOpened.checked = window.frmMain.chkOTS.checked;
	window.frmMain.chkOTSClosed.checked = window.frmMain.chkOTS.checked;
}

function chkOTSSub_onclick() {
	window.frmMain.chkOTS.indeterminate=0;
	if (window.frmMain.chkOTSClosed.checked && window.frmMain.chkOTSOpened.checked)
		window.frmMain.chkOTS.checked = true;
	else if(window.frmMain.chkOTSClosed.checked==false && window.frmMain.chkOTSOpened.checked==false) 
		window.frmMain.chkOTS.checked = false;
	else
		window.frmMain.chkOTS.indeterminate=-1;
}


function chkBcrChangeSub_onclick() {
	window.frmMain.chkBcrChange.indeterminate=0;
	if (window.frmMain.chkBcrChangeClosed.checked && window.frmMain.chkBcrChangeOpened.checked)
		window.frmMain.chkBcrChange.checked = true;
	else if(window.frmMain.chkBcrChangeClosed.checked==false && window.frmMain.chkBcrChangeOpened.checked==false) 
		window.frmMain.chkBcrChange.checked = false;
	else
		window.frmMain.chkBcrChange.indeterminate=-1;
}

function chkChangeSub_onclick() {
	window.frmMain.chkChange.indeterminate=0;
	if (window.frmMain.chkChangeClosed.checked && window.frmMain.chkChangeOpened.checked)
		window.frmMain.chkChange.checked = true;
	else if(window.frmMain.chkChangeClosed.checked==false && window.frmMain.chkChangeOpened.checked==false) 
		window.frmMain.chkChange.checked = false;
	else
		window.frmMain.chkChange.indeterminate=-1;
}

function cmdDate_onclick(FieldID) {
	var strID;
		
	strID = window.showModalDialog("./mobilese/today/caldraw1.asp",FieldID,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) == "undefined")
		return
	
	window.frmMain.elements(FieldID).value = strID;
}

//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  10/12/2016 - Harris, Valerie      
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();

    //Add datepicker
    load_datePicker();
}
//-->
</SCRIPT>
</HEAD>
<STYLE>
BODY
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}
TD
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}

H1
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: small;
}
.DateInput
{
	font-family: Verdana;
	font-size: x-small;
	height: 20;
	width: 80;
	border: solid 1px gray;
}
</STYLE>
<body style="background-color:ivory;" id="MainBody" onload="window_onload();">
<form action="" method="post" id="frmMain" name="frmMain">
<h1>Change Log</h1>
<span style="font-size:x-small; font-weight:bold">Date Range</span>
<table width="100%" style=" background-color:cornsilk; border: tan double;" cellspacing="0" cellpadding="2">
<tr><td style="white-space:nowrap"><font size="2">Start Date:&nbsp;</font></td>
<td width="100%"><input type="text" class="DateInput, dateselection" id=txtStartDt name=txtStartDt value="<%= FormatDateTime(Now()-7,2)%>" />
</td></tr>
<tr><td nowrap><font size=2>End Date:</font></td><td><input type="text" class="DateInput, dateselection" id=txtEndDt name=txtEndDt value="<%= FormatDateTime(Now(),2)%>" />
</td></tr></table><br>
<font size=2><b>Report Sections</b><BR></font>
<table width="100%" bgcolor=cornsilk style="BORDER-LEFT-COLOR: tan; BORDER-BOTTOM-COLOR: tan; BORDER-TOP-STYLE: double; BORDER-TOP-COLOR: tan; BORDER-RIGHT-STYLE: double; BORDER-LEFT-STYLE: double; BORDER-RIGHT-COLOR: tan; BORDER-BOTTOM-STYLE: double" cellSpacing=0 cellPadding=0>
<tr><td><input type="checkbox" id="chkChange" name="chkChange" checked onclick="return chkChange_onclick()" />&nbsp;</td><td>Change Requests (DCR)</td></tr>
	<tr><td>&nbsp;</td><td width="100%"><input type="checkbox" id=chkChangeOpened name=chkChangeOpened checked onclick="return chkChangeSub_onclick()" />&nbsp;Opened</td></tr>
	<tr><td>&nbsp;</td><td><input type="checkbox" id=chkChangeClosed name=chkChangeClosed checked onclick="return chkChangeSub_onclick()" />&nbsp;Closed</td></tr>
<tr><td><input type="checkbox" id=chkBcrChange name=chkBcrChange checked onclick="return chkBcrChange_onclick()" />&nbsp;</td><td>Change Requests (BCR)</td></tr>
	<tr><td>&nbsp;</td><td width="100%"><input type="checkbox" id="chkBcrChangeOpened" name="chkBcrChangeOpened" checked onclick="return chkBcrChangeSub_onclick()" />&nbsp;Opened</td></tr>
	<tr><td>&nbsp;</td><td><input type="checkbox" id="chkBcrChangeClosed" name="chkBcrChangeClosed" checked onclick="return chkBcrChangeSub_onclick()" />&nbsp;Closed</td></tr>
<tr><td><input type="checkbox" id=chkDeliverables name=chkDeliverables checked />&nbsp;</td><td>Deliverable Matrix Updates</td></tr>
<tr><td><input type="checkbox" id=chkOTS name=chkOTS checked onclick="return chkOTS_onclick()" />&nbsp;</td><td>Observations</td></tr>
	<tr><td>&nbsp;</td><td><input type="checkbox" id=chkOTSOpened name=chkOTSOpened checked onclick="return chkOTSSub_onclick()" />&nbsp;Opened</td></tr>
	<tr><td>&nbsp;</td><td><input type="checkbox" id=chkOTSClosed name=chkOTSClosed checked onclick="return chkOTSSub_onclick()" />&nbsp;Closed</td></tr>
<tr><td><input type="checkbox" id=chkSchedule name=chkSchedule checked />&nbsp;</td><td>Schedule Changes</td></tr>
<tr><td><input type="checkbox" id=chkAgency name=chkAgency checked />&nbsp;</td><td>Agency Changes</td></tr>
<tr><td><input type="checkbox" id=chkCountry name=chkCountry checked />&nbsp;</td><td>Country Changes</td></tr>
<tr><td><input type="checkbox" id=chkLocalization name=chkLocalization checked />&nbsp;</td><td>Localization Changes</td></tr>
</table>

<input type="hidden" id=txtID name=txtID value="<%=request("ID")%>" />
<input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
</FORM>
</BODY>
</HTML>
