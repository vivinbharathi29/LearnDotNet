<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
	var ID = txtID.value;
	var SectionList = "";
//	var	NewLeft = (screen.width - 655)/2;
//	var NewTop = (screen.height - 650)/2;
	
	if (chkChangeOpened.checked)
		SectionList = SectionList + ",12"
	if (chkChangeClosed.checked)
		SectionList = SectionList + ",13"
	if (chkDeliverables.checked)
		SectionList = SectionList + ",6"
	if (chkOTSOpened.checked)
		SectionList = SectionList + ",14"
	if (chkOTSClosed.checked)
		SectionList = SectionList + ",15"
	if (chkAgency.checked)
		SectionList = SectionList + ",16"
	
	if (SectionList == "")
		alert("No report sections selected");
	else
		{
		SectionList = SectionList.substr(1);
		//MainBody.innerHTML = "Processing.  Please Wait...";
		//window.open("ProductStatus.asp?ID=" + ID + "&ReportDays=7&ReportTitle= - This Week&Sections=" + SectionList,"_blank","Left=" + NewLeft + ",Top=" + NewTop + ",Width=655,Height=650,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes"); 
		window.open("ProductStatus.asp?ID=" + ID + "&ReportDays=7&ReportTitle= - This Week&Sections=" + SectionList,"_blank","Width=655,Height=500,menubar=yes,toolbar=yes,resizable=Yes,scrollbars=yes,status=yes,location=yes"); 
	
		window.close();
		}

}

function chkChange_onclick() {
	chkChangeOpened.checked = chkChange.checked;
	chkChangeClosed.checked = chkChange.checked;
}

function chkOTS_onclick() {
	chkOTSOpened.checked = chkOTS.checked;
	chkOTSClosed.checked = chkOTS.checked;
}

function chkOTSSub_onclick() {
	chkOTS.indeterminate=0;
	if (chkOTSClosed.checked && chkOTSOpened.checked)
		chkOTS.checked = true;
	else if(chkOTSClosed.checked==false && chkOTSOpened.checked==false) 
		chkOTS.checked = false;
	else
		chkOTS.indeterminate=-1;
}



function chkChangeSub_onclick() {
	chkChange.indeterminate=0;
	if (chkChangeClosed.checked && chkChangeOpened.checked)
		chkChange.checked = true;
	else if(chkChangeClosed.checked==false && chkChangeOpened.checked==false) 
		chkChange.checked = false;
	else
		chkChange.indeterminate=-1;

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
</STYLE>
<BODY bgcolor=ivory id=MainBody>


<H1>Build Employee List</H1>
<font size=2><b>Select Employees By Manager</b><BR></font>
<table width="100%" bgcolor=cornsilk style="BORDER-LEFT-COLOR: tan; BORDER-BOTTOM-COLOR: tan; BORDER-TOP-STYLE: double; BORDER-TOP-COLOR: tan; BORDER-RIGHT-STYLE: double; BORDER-LEFT-STYLE: double; BORDER-RIGHT-COLOR: tan; BORDER-BOTTOM-STYLE: double" cellSpacing=0 cellPadding=0>
<tr><td></td></tr>
</table>
<HR>
<table width="100%">
	<TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>
</table>
</BODY>
</HTML>