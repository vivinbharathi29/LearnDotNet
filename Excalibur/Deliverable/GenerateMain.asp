<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<TITLE>Generate Files</TITLE>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
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
		
		for (i=event.srcElement.length-1;i>=0;i--)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
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

function replace(sampleStr,replaceChars,replaceWith){

	var replaceRegExp = new RegExp(replaceChars,'g');

	return sampleStr.replace(replaceRegExp,replaceWith);
}

function window_onload() {
	if (txtOSInd.value == "True")
		{
			window.parent.location.replace ("NoOSInd.asp");
			return;
		}

	WaitLabel.style.display="none";
	
	var i;
	var strID;
	var strName;
	var strModelNames="";
	
	CurrentState =  "Softpaq";
	ProcessState();
	FormLoading = false;	
	
	
}


function EditField(strField, RootID, VerID) {
	if (strField=="Title" || strField=="Category" || strField=="Description")
		strID = window.showModalDialog("../root.asp?ID=" + RootID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
	else
		strID = window.showModalDialog("../WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + VerID,"","dialogWidth:700px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 


//	if (typeof(strID) != "undefined")
		
	

}

function Replace(strsstrfind){

}

function FormatDate(MyDate){
	var strDate = "";
	var strDay = "";
	var strYear = "";

	
	months = "January,February,March,April,May,June,July,August,September,October,November,December".split(',') 
	
	if (MyDate.getDate() < 10)
		strDay =  String("0" + MyDate.getDate()).substring(0,2);
	else
		strDay = MyDate.getDate();
	
	if (MyDate.getFullYear() < 1970)
		strYear = MyDate.getFullYear() + 100
	else
		strYear = MyDate.getFullYear() 
	
	strDate = months[MyDate.getMonth()] + " " + strDay + ", " + strYear;
	return strDate;
}


function SplitRows(strRows){
	var i;
	var j;
	var RowLength= 0;
	var strNewRows="";
	var strRowBuffer="";
	var RowArray = strRows.split("\r\n");
	var TokenArray

	for  (i=0;i<RowArray.length;i++)
		{
		strRowBuffer = "";
		TokenArray=RowArray[i].split(" ");
		for (j=0;j<TokenArray.length;j++)
			{
			if (TokenArray[j].length + strRowBuffer.length + 1 > 80)
				{
				strNewRows = strNewRows + strRowBuffer + "\r\n"
				strRowBuffer="";
				}
			if (strRowBuffer!="")
				strRowBuffer = strRowBuffer + " ";
				 
			strRowBuffer = strRowBuffer + TokenArray[j];
			}
		
		strNewRows = strNewRows + strRowBuffer
		if (i != RowArray.length-1)
			 strNewRows = strNewRows + "\r\n";
		}
	return strNewRows;

}

function EnumerateEnhancements(strText){

	var strOutput = "";
	var RowArray = strText.split("\r\n");
	var i;
	
	for (i=0;i<RowArray.length;i++)
		{
		strOutput = strOutput + "Enh" + (i+1) + "=" + RowArray[i] + "\r\n"
		}

	return strOutput;
}


function GenerateFiles(){
	var strSoftpaq = "";
	var strWelcome = "";
	var strCVA = "";
	var strBuffer = "";
	var strCVABuffer = "";
	var strCVABuffer2 = "";
	var strLangList = "";
	var i;
	var EffectiveDate = new Date(frmSoftpaq.txtEffective.value);

	strCVA="[CVA File Information]\r\n"
	strCVA= strCVA + "CVA Version=" + frmSoftpaq.txtCVAVersion.value + "\r\n\r\n"

	strCVA= strCVA + "[General]\r\n"
	strCVA= strCVA + "Category=" + frmSoftpaq.cboCategory.options[frmSoftpaq.cboCategory.selectedIndex].text + "\r\n"
	strCVA= strCVA + "Version=" + frmSoftpaq.txtDelVersion.value + "\r\n"
	strCVA= strCVA + "Revision=" + frmSoftpaq.txtRevision.value + "\r\n"
	strCVA= strCVA + "Pass=" + frmSoftpaq.txtPass.value + "\r\n\r\n"

	strCVA= strCVA + "[Install Execution]\r\n"
	strCVA= strCVA + "ARCDInstall=" + frmSoftpaq.txtARCD.value + "\r\n"
	strCVA= strCVA + "SilentInstall=" + frmSoftpaq.txtSilent.value + "\r\n\r\n"

	strCVA= strCVA + "[DetailedFileInformation]\r\n"
	strCVA= strCVA + "TBD\r\n"
	strCVA= strCVA + "\r\n" //blank line

	strSoftpaq = SplitRows("TITLE: " + frmSoftpaq.txtTitle.value + "\r\n\r\n");
	strSoftpaq = strSoftpaq + "VERSION: " + frmSoftpaq.txtVersion.value + "\r\n";
	strSoftpaq = strSoftpaq + SplitRows("DESCRIPTION:\r\n" + frmSoftpaq.txtDescription.value + "\r\n\r\n");
	strSoftpaq = strSoftpaq + "PURPOSE: " + frmSoftpaq.cboPurpose.options[frmSoftpaq.cboPurpose.selectedIndex].text + "\r\n";
	if (frmSoftpaq.txtMultiLanguage.value == "1")
		{
		strSoftpaq = strSoftpaq + "SOFTPAQ NUMBER: " + frmSoftpaq.txtSoftpaqNumber.value + "\r\n";
		strWelcome = strSoftpaq;
		if (trim(frmSoftpaq.txtSoftpaqSupersedes.value) == "")
			strSoftpaq = strSoftpaq + "SUPERSEDES: N/A\r\n";
		else
			strSoftpaq = strSoftpaq + "SUPERSEDES: " + frmSoftpaq.txtSoftpaqSupersedes.value + "\r\n";
		
		strCVA = strCVA + "[Softpaq]\rSoftpaqNumber=" + frmSoftpaq.txtSoftpaqNumber.value + "\r\n"
		strCVA = strCVA + "SupercededSoftpaqNumber=" + frmSoftpaq.txtSoftpaqSupersedes.value + "\r\n\r\n"
		}
	else
		{
		strSoftpaq = strSoftpaq + "SOFTPAQ NUMBER: [VARIES BY LANGUAGE]\r\n";
		strWelcome = strSoftpaq;
		strSoftpaq = strSoftpaq + "SUPERSEDES: [VARIES BY LANGUAGE]\r\n";
		strCVA = strCVA + "[Softpaq]\r\nSoftpaqNumber=[VARIES BY LANGUAGE]\r\n"
		strCVA = strCVA + "SupercededSoftpaqNumber=[VARIES BY LANGUAGE]\r\n\r\n"
		}

	strBuffer= "EFFECTIVE DATE: " + FormatDate(EffectiveDate) + "\r\n";
	strSoftpaq = strSoftpaq + strBuffer;
	strWelcome = strWelcome + strBuffer;

	strBuffer = "CATEGORY: " + frmSoftpaq.cboCategory.options[frmSoftpaq.cboCategory.selectedIndex].text + "\r\n";
	strSoftpaq = strSoftpaq + strBuffer;
	strWelcome = strWelcome + strBuffer;

	if (frmSoftpaq.chkSSM.checked)
		strBuffer = "SSM SUPPORTED: Yes\r\n\r\n";
	else
		strBuffer = "SSM SUPPORTED: No\r\n\r\n";
	
	strSoftpaq = strSoftpaq + strBuffer;
	strWelcome = strWelcome + strBuffer;

	strBuffer = "";
	for (i=0;i<frmSoftpaq.lstType.length;i++)
		if (frmSoftpaq.lstType[i].checked)
			strBuffer = strBuffer + ", " + frmSoftpaq.lstType[i].value;
	if (strBuffer.length > 0)
		strBuffer = strBuffer.substring(2);

	strSoftpaq = strSoftpaq + SplitRows("PRODUCT TYPE(S):\r\n" + strBuffer + "\r\n\r\n");


	strSoftpaq = strSoftpaq + SplitRows("PRODUCT MODEL(S):\r\n" + "TBD" + "\r\n\r\n");

	strCVABuffer = "";
	intCounter = 0;
	for (i=0;i<frmSoftpaq.lstProduct.length;i++)
		if (frmSoftpaq.lstProduct[i].checked)
			{
			intCounter = intCounter + 1
			strCVABuffer = strCVABuffer + "SysID" + intCounter + "=0x" + frmSoftpaq.txtSystemID[i].value.substring(0,frmSoftpaq.txtSystemID[i].value.length-1) + "\r\n";
			}

	strCVA = strCVA + "[System Information]\r\n" + strCVABuffer + "\r\n"

	strBuffer = "";
	strCVABuffer = "";
	for (i=0;i<frmSoftpaq.lstPNPDevices.length;i++)
		{
		strBuffer = strBuffer + frmSoftpaq.lstPNPDevices[i].text.substring(frmSoftpaq.lstPNPDevices[i].text.lastIndexOf("=") + 1) + "\r\n";
		strCVABuffer = strCVABuffer + frmSoftpaq.lstPNPDevices[i].text.substring(3) + "\r\n"
		}
	if (strBuffer.length > 0)
		strSoftpaq = strSoftpaq + SplitRows("DEVICES SUPPORTED:\r\n" + strBuffer + "\r\n");
	if (trim(strCVABuffer) != "")
		strCVA = strCVA + "[Devices]\r" + strCVABuffer + "\r\n"


	strBuffer = "";
	strCVABuffer = "";
	for (i=0;i<frmSoftpaq.lstOS.length;i++)
		if (frmSoftpaq.lstOS[i].checked)
			{
			strBuffer = strBuffer + frmSoftpaq.lstOS[i].value +"\r\n";
			strCVABuffer = strCVABuffer + frmSoftpaq.txtOSKey[i].value + "=" + frmSoftpaq.cboMinLevel[i].value + "\r\n";
			}

	strSoftpaq = strSoftpaq + "OPERATING SYSTEM(S):\r\n" + strBuffer + "\r\n";
	strCVA = strCVA + "[Operating Systems]\r\n" + strCVABuffer + "\r\n";

	strCVA = strCVA + "[Software Title]\r\n"

	if (frmSoftpaq.txtMultiLanguage.value == "1")
	{
		strBuffer = "";
		strCVABuffer = "";
		strCVABuffer2 = "";
		for (i=0;i<frmSoftpaq.lstLanguage.length;i++)
			if (frmSoftpaq.lstLanguage[i].checked)
				{
				strBuffer = strBuffer + frmSoftpaq.txtLanguage[i].value +"\r\n";
				strCVABuffer = strCVABuffer + "," + frmSoftpaq.txtLanguage[i].value.substring(0,2)
				strCVABuffer2 = strCVABuffer2 + "[" + frmSoftpaq.txtLanguage[i].value.substring(0,2) + ".Software Description]\r\n" + replace(frmSoftpaq.txtCVADescription[i].value,"\r\n"," ")  + "\r\n\r\n"
				strCVA=strCVA + frmSoftpaq.txtLanguage[i].value.substring(0,2) + "=" + frmSoftpaq.txtCVATitle[i].value  + "\r\n"
				}
		strLangList = strLangList + ",LI=" + trim(frmSoftpaq.txtSoftpaqNumber.value) + "=" + trim(frmSoftpaq.txtSoftpaqSupersedes.value);
		strSoftpaq = strSoftpaq + "LANGUAGE(S):\r\n" + strBuffer + "\r\n";
	}
	else
	{
		strBuffer = "";
		strCVABuffer = "";
		for (i=0;i<frmSoftpaq.txtSoftpaqNumber.length;i++)
			if (trim(frmSoftpaq.txtSoftpaqNumber[i].value) != "")
				{
				strBuffer = strBuffer + frmSoftpaq.txtLanguage[i].value +"\r\n";
				strCVABuffer = strCVABuffer + "," + frmSoftpaq.txtLanguage[i].value.substring(0,2)
				strCVABuffer2 = strCVABuffer2 + "[" + frmSoftpaq.txtLanguage[i].value.substring(0,2) + ".Software Description]\r\n" + replace(frmSoftpaq.txtCVADescription[i].value,"\r\n"," ")  + "\r\n\r\n"
				strCVA=strCVA + frmSoftpaq.txtLanguage[i].value.substring(0,2) + "=" + frmSoftpaq.txtCVATitle[i].value  + "\r\n"
				strLangList = strLangList + "," + frmSoftpaq.txtLanguage[i].value.substring(0,2) + "=" + trim(frmSoftpaq.txtSoftpaqNumber[i].value) + "=" + trim(frmSoftpaq.txtSoftpaqSupersedes[i].value);
				}
		strSoftpaq = strSoftpaq + "LANGUAGE(S):\r\n" + strBuffer + "\r\n";
	}
	
	frmSoftpaq.txtSoftpaqList.value = strLangList.substring(1);
	if (strCVABuffer.length> 0)
		strCVABuffer = strCVABuffer.substr(1);
	strCVA = strCVA + "\r\n[SupportedLanguages]\r\nLanguages=" + strCVABuffer + "\r\n\r\n";

	if (strCVABuffer2.length> 0)
		strCVA = strCVA +  strCVABuffer2;


	if (trim(frmSoftpaq.txtEnhancements.value) != "")
		{
		strBuffer = SplitRows("ENHANCEMENTS:\r\n" + frmSoftpaq.txtEnhancements.value + "\r\n\r\n");
		strSoftpaq = strSoftpaq + strBuffer;
		strWelcome = strWelcome + strBuffer;
		strCVA = strCVA + "[US.Enhancements]\r\n" + EnumerateEnhancements(frmSoftpaq.txtEnhancements.value);
		}
		
	if (trim(frmSoftpaq.txtFixes.value) != "")
		{
		strBuffer = SplitRows("FIXES:\r\n" + frmSoftpaq.txtFixes.value + "\r\n\r\n");
		strSoftpaq = strSoftpaq + strBuffer;
		strWelcome = strWelcome + strBuffer;
		}
		
	if (trim(frmSoftpaq.txtPrerequsites.value) != "")
		{
		strBuffer = SplitRows("PREREQUISITES:\r\n" + frmSoftpaq.txtPrerequsites.value + "\r\n\r\n");
		strSoftpaq = strSoftpaq + strBuffer;
		strWelcome = strWelcome + strBuffer;
		}
		
	strBuffer = SplitRows("HOW TO USE:\r\n" + frmSoftpaq.txtHowToUse.value + "\r\n\r\n");
	strSoftpaq = strSoftpaq + strBuffer;
	strWelcome = strWelcome + strBuffer;

	strBuffer = "Copyright (c) 2019 HP Development Company, L.P."
	strSoftpaq = strSoftpaq + strBuffer;
	strWelcome = strWelcome + strBuffer;


	frmSoftpaq.txtPreview.value= strSoftpaq 
	frmSoftpaq.txtPreviewWelcome.value= strWelcome
	frmSoftpaq.txtPreviewCVA.value= strCVA
	
}

var CurrentState;
var FormLoading = true;


function ProcessState() {
	var steptext;
	var strPreview;
	
	switch (CurrentState)
	{
		case "Softpaq":
		
		lblTitle.innerText = "Softpaq Fields - Step 1 of 3";
		lblInstructions.innerText = "Enter Softpaq Information.";

		tabSoftpaq.style.display="";
		tabCVA.style.display="none";
		tabPreview.style.display="none";
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = true;
			window.parent.frames["LowerWindow"].cmdNext.disabled = false;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
		}
		window.scrollTo(0,0);
		break;

		case "CVA":
		
		lblTitle.innerText = "CVA Fields - Step 2 of 3";
		lblInstructions.innerText = "Enter the remaining CVA information.";

		tabSoftpaq.style.display="none";
		tabCVA.style.display="";
		tabPreview.style.display="none";
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			window.parent.frames["LowerWindow"].cmdNext.disabled = false;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
		}
		window.scrollTo(0,0);
		break;


		case "Preview":
		
		GenerateFiles();

		lblTitle.innerText = "Review Softpaq File Format - Step 3 of 3";
		lblInstructions.innerText = "Review softpaq file format and click \"Finish\" button to generate the files.";

		tabSoftpaq.style.display="none";
		tabCVA.style.display="none";
		tabPreview.style.display="";
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			window.parent.frames["LowerWindow"].cmdNext.disabled = true;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = false;
		}
		
		//for (i=0;i<frmSoftpaq.chkSelected.length;i++)
			
		
//		frmSoftpaq.txtPreview.value = frmSoftpaq.txtPreview1.value  + "\r" + frmSoftpaq.txtSoftpaqLines.value  + frmSoftpaq.txtPreview2.value + "\r" + "PRODUCT MODEL(S):\r" + txtSelectedProducts.value + "\r" + frmSoftpaq.txtPreview3.value;
//		frmPMR.txtPreview.value = strPreview;
		frmSoftpaq.txtPreview.focus();
		window.scrollTo(0,0);
		break;
	}
}

function cmdEffective_onclick() {
	var strID;
	var i;
	
	strID = window.showModalDialog("../Mobilese/today/calDraw1.asp",frmSoftpaq.txtEffective.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) != "undefined")
		{
			frmSoftpaq.txtEffective.value = strID;
		}
}


function mouseover_Column(){
	event.srcElement.style.color="red";
	event.srcElement.style.cursor="hand";
	
}
function mouseout_Column(){
	event.srcElement.style.color="black";
}

function mouseover_Cell(){
	event.srcElement.parentElement.style.color="red";
	event.srcElement.parentElement.style.cursor="hand";
}

function mouseout_Cell(){
	event.srcElement.parentElement.style.color="black";
}

function onclick_Cell(chkBox, strID){
	var i;
	if(event.srcElement.name==chkBox(0).name)
		return;
	for (i=0;i<chkBox.length;i++)
		{
			if (chkBox(i).value==strID)
				{
				if (chkBox(i).checked)
					{
					chkBox(i).checked = false;
					}
				else
					{
					chkBox(i).checked = true;
					}
				}
		}
}

function cmdEditDevice_onclick() {
	lstPNPDevices_ondblclick();
}


function onclick_CheckAll(chkBox,chkBoxAll) {
	for (i=0;i<chkBox.length;i++)
		{
		chkBox(i).checked = chkBoxAll.checked;
		}
}
function lstPNPDevices_ondblclick() {
	var rc;
	if (frmSoftpaq.lstPNPDevices.selectedIndex>-1)
		{
		rc = window.showModalDialog  ("../Devices.asp",frmSoftpaq.lstPNPDevices.options(frmSoftpaq.lstPNPDevices.selectedIndex).text,"dialogWidth=28;dialogHeight=14;edge: Raised; center: Yes; help: No; resizable: No; status: No;Scroll: No;");
		if (typeof(rc) == "undefined")
			{
			
			}
		else
			{
			frmSoftpaq.lstPNPDevices.options(frmSoftpaq.lstPNPDevices.selectedIndex).innerText = rc;
			frmSoftpaq.txtPNPDevices.value = "";
			for (i=0;i<frmSoftpaq.lstPNPDevices.length;i++)
				{
					frmSoftpaq.txtPNPDevices.value = frmSoftpaq.txtPNPDevices.value + frmSoftpaq.lstPNPDevices.options(i).text + "\r";
				}
			
			}
	}
}

function cmdDeleteDevice_onclick() {
	if (frmSoftpaq.lstPNPDevices.selectedIndex>-1)
		{
		frmSoftpaq.lstPNPDevices.options.remove(frmSoftpaq.lstPNPDevices.selectedIndex);
		}
	frmSoftpaq.txtPNPDevices.value = "";
	for (i=0;i<frmSoftpaq.lstPNPDevices.length;i++)
		{
			frmSoftpaq.txtPNPDevices.value = frmSoftpaq.txtPNPDevices.value  + frmSoftpaq.lstPNPDevices.options(i).text + "\r";
		}
		
}

function cmdNewDevice_onclick() {
	var oOption = document.createElement("OPTION");
	var rc;
	var Args;
	var i;
	Args = ""
	rc = window.showModalDialog  ("../Devices.asp",Args,"dialogWidth=28;dialogHeight=14;edge: Raised; center: Yes; help: No; resizable: No; status: No;Scroll: No;");
	if (typeof(rc) == "undefined")
		{
		
		}
	else
		{
		frmSoftpaq.lstPNPDevices.options.add(oOption);
		oOption.innerText =rc;
		frmSoftpaq.txtPNPDevices.value = "";
		for (i=0;i<frmSoftpaq.lstPNPDevices.length;i++)
			{
				frmSoftpaq.txtPNPDevices.value = frmSoftpaq.txtPNPDevices.value  + frmSoftpaq.lstPNPDevices.options(i).text + "\r";
			}
		}
	oOption.Value = "0";

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
<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font size=3 face=verdana color=red>This page is under development.  Do not use the files generated by this process.</font><BR><BR>


<%

if request("VersionID") = "" then
	Response.Write "No Version ID Supplied."
else
	dim blnOSInd
	dim cn
	dim rs
	dim cm
	dim p
	dim blnFound 
	dim i
	dim strTitle
	dim strCategories
	dim CurrentUser
	dim CurrentUserID
	dim CurrentWorkgroupID
	dim strSEPMID
	dim strPMID
	dim blnPOR
	dim blnEditOK
	dim strShowEditBoxes
	dim strDelList
	dim strVersion
	dim strProductName
	dim strEmployees
	dim strDeliverable
	dim strFilename
	dim strRootID
	dim strDescription
	dim strPurpose
	dim strHowToUse
	dim strPrerequsites
	dim strFixes
	dim strEnhancements
	dim strEffective
	dim strOS
	dim strMultiLanguage
	dim strSoftpaqNumber
	dim strSupersedes
	dim strSSM
	dim strCategory
	dim strPNPDevices
	dim MinLevelArray
	dim Min95LevelArray
	dim TypeArray
	dim strSoftpaqType
	dim strARCDInstall
	dim strSilentInstall
	dim strPass
	dim strFileInfo
	dim RootID
	dim strCVAVersion
	dim strDelVersion
	dim strRevision
			
	MinLevelArray=split("OEM,SP1,SP2,SP3,SP4,SP5,SP6,SP7",",")
	Min95LevelArray=split("OSR0,OSR1,OSR2,OSR21,OSR25",",")
	TypeArray = split("Notebooks,Desktops,Workstations,Thin Clients,Monitors,Projectors,Handhelds,Printers,Personal Audio",",")
	strDelList = ""	
	strProductName = ""
	strDeliverable = ""
	strFilename = ""
	strHowToUse = ""
	strPrerequsites = ""
	strFixes = ""
	strEnhancements = ""
	strEffective=""
	strMultiLanguage = ""
	strSupersedes	= ""
	strSoftpaqNumber = ""
	strSSM = ""
	strCategory = ""
	strPNPDevices = ""
	strSoftpaqType = ""
	strSilentInstall = ""
	strARCDInstall = ""
	strPass = ""
	blnOSInd = false
	strFileInfo = ""
	strCVAVersion = ""
	strDelVersion = ""
	strRevision = ""
	
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

	set cm=nothing
	
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentWorkgroupID = rs("WorkgroupID") & ""
	end if
	rs.Close
	
	strDescription = ""
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetSoftpaqText"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("VersionID")
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing
	
	'rs.Open "spGetSoftpaqText " & request("VersionID"),cn,adOpenForwardOnly
		strTitle = trim(rs("title") & "")
		strVersion = trim(rs("version") & "")
		strDescription = trim(rs("DESCRIPTION") & "")
		strPurpose = trim(rs("Purpose") & "")
		strHowToUse = trim(rs("HowToUse") & "")
		strFilename = trim(rs("Filename") & "")
		strPrerequsites = trim(rs("Prerequisites") & "")
		strFixes = trim(rs("Fixes") & "")
		strEnhancements = trim(rs("Enhancements") & "")
		strEffective=trim(rs("EffectiveDate") & "")
		strMultiLanguage = trim(rs("MultiLanguage") & "")
		strSupersedes = trim(rs("Supersedes") & "")
		strSoftpaqNumber = trim(rs("SoftpaqNumber") & "")
		strSSM = replace(replace(trim(rs("SSM") & ""),"True","checked"),"False","")
		strCategory = trim(rs("SoftpaqCategory") & "")
		strPNPDevices = trim(rs("PNPDevices") & "")
		strSoftpaqType = trim(rs("SoftpaqType") & "")
		strSilentInstall = trim(rs("SilentInstall") & "")
		strARCDInstall = trim(rs("ARCDInstall") & "")
		strDelVersion = trim(rs("DelVersion") & "")
		strRevision = trim(rs("Revision") & "")
		strPass = trim(rs("Pass") & "")
		strFileInfo = trim(rs("DetailedFileInfo") & "" )
		RootID = trim(rs("RootID") & "")
		strCVAVersion = trim(rs("CVAVersion") & "")
	rs.Close


if 	strSoftpaqType = "" then
	strSoftpaqType = "Notebooks"
end if

if strCVAVersion = "" then
	strCVAVersion = 1
else
	strCVAVersion = strCVAVersion + 1
end if

%>

<font size=4 face=verdana><b><%=strTitle & " " & strVersion%></b></font><BR><BR>
<h4>
<label ID="lblTitle"></label>
</h3>

<font size="2">
<label ID="lblInstructions"></label>
</font>
<%
	Response.Write "<font size=2 face=verdana><label ID=WaitLabel>Loading Product List.  Please wait...</label></font>"
	'Response.Flush
		
	
%>

<form id="frmSoftpaq" method="post" action="GenerateSave.asp">
<span ID=tabSoftpaq style="Display:none">

<table width=100% cellpadding=2 cellspacing=0 border=1 bordercolor=tan bgcolor=cornsilk>
<Tr>
	<TD><b>Title:</b></td>
	<TD width=100%><%=strTitle%>
	<INPUT type="hidden" id=txtTitle name=txtTitle value= "<%=strTitle%>">
	</td>
</tr>
<Tr>
	<TD><b>Version:</b></td>
	<TD width=100%><%=strVersion%>
	<INPUT type="hidden" id=txtVersion name=txtVersion value= "<%=strVersion%>">
	</td>
</tr>
<tr>
	<TD valign=top><b>Description:</b></td>
	<TD width=100%><TEXTAREA rows=4 style="width:100%"  id=txtDescription name=txtDescription><%=strDescription%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>Category:</b></td>
	<TD width=100%>
	<SELECT id=cboCategory name=cboCategory>
	<%
		rs.Open "spListSoftpaqCategories",cn,adOpenForwardOnly
		do while not rs.EOF
			if rs("ID") <> 1 then 'remove "N/A"
				if strCategory = trim(rs("Name") & "") then
					Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>"
				else
					Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>"
				end if
			end if
			rs.MoveNext
		loop
		rs.Close
	%>
	</SELECT>
	</td>
</tr>

<tr>
	<td width=160 nowrap><b>Purpose:</b></td>
	<td colspan="10">
			<select id="cboPurpose" name="cboPurpose" style="WIDTH: 170px">
			<%if strPurpose = "" or strPurpose = "0" then %>
					<option selected value="0"></option>
			<%else%>
					<option value="0"></option>
			<%end if%>
			<%if strPurpose = "1" then %>
					<option selected value="1">Routine Release</option>
			<%else%>
					<option value="1">Routine Release</option>
			<%end if%>
			<%if strPurpose = "2" then %>
					<option selected value="2">Recommended Update</option>
			<%else%>
					<option value="2">Recommended Update</option>
			<%end if%>
			<%if strPurpose = "3" then %>
				<option selected value="3">Critical Update</option>
			<%else%>
				<option value="3">Critical Update</option>
			<%end if%>
			
			</select>
		</td>


</tr>
	<%if trim(strMultiLanguage) <> "1" then%>
<tr>
	<TD valign=top><b>Softpaq&nbsp;Numbers:</b></td>
	<TD width=100%>
	<INPUT type="hidden" id=txtMultiLanguage name=txtMultiLanguage value="0">
	
	<font size=1 face=verdana color=green>Enter softpaq number for each supported language.</font>
	<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIV1>
		<TABLE ID=TableSoftpaqNumbers width=100%>
			<THEAD bgcolor=LightSteelBlue  ><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Language/Region&nbsp;</TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Number&nbsp;</TD><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">Supercedes&nbsp;</TD></THEAD>
		<% 
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetSelectedLanguages"
	

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p


			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
			
			'rs.Open "spGetSelectedLanguages " & request("VersionID"),cn,adOpenForwardOnly
			do while not rs.eof
			%>
			<TR>
				<TD nowrap><%=rs("Abbreviation") & " - " & rs("Name")%></TD>			
				<TD nowrap><INPUT style="Width=100%" type=text name=txtSoftpaqNumber value="<%=rs("SoftpaqNumber")%>" id="txtSoftpaqNumber"><INPUT type="hidden" id=txtLanguage name=txtLanguage value="<%=rs("Abbreviation") & " - " & rs("Name")%>"></TD>			
				<TD nowrap><INPUT style="Width=100%" id=txtSoftpaqSupersedes type=text name=txtSoftpaqSupersedes value="<%=rs("Supersedes")%>"></TD>
			</TR>
			<%
				rs.MoveNext
			loop
			rs.Close
			%>
		</TABLE>    
	</div>	
	
	
	
	</td>
</tr>
<%else%>
<tr>
	<TD valign=top><b>Softpaq&nbsp;Number:</b></td>
	<TD width=100%>
	<INPUT type="hidden" id=txtMultiLanguage name=txtMultiLanguage value="1">
	<INPUT style="width:170px" id=txtSoftpaqNumber type=text name=txtSoftpaqNumber value="<%=strSoftpaqNumber%>"></td>
</tr>
<tr>
	<TD valign=top><b>Supercedes:</b></td>
	<TD width=100%><INPUT style="width:170px"  id=txtSoftpaqSupersedes type=text name=txtSoftpaqSupersedes value="<%=strSupersedes%>"></td>
</tr>

<%end if%>

<TR>
	<TD><font size=2 face=verdana><b>Effective Date:</b></td>
	<td width=100%>
		<table cellpadding=0 cellspacing=0>
			<TR>
				<TD><input id="txtEffective" name="txtEffective" style="WIDTH: 170px; HEIGHT: 22px" size="28" value="<%=strEffective%>"></TD>
				<TD>&nbsp;<a href="javascript: cmdEffective_onclick()"><img ID="picTarget" SRC="../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21"></a></font></td>
			</tr>
		</TABLE>
	</TD>
</TR>
<tr>
	<TD valign=top><b>SSM:</b></td>
	<TD width=100%><INPUT type="checkbox" id=chkSSM name=chkSSM <%=strSSM%>>&nbsp;SSM Compliant</td>
</tr>

<tr>
	<TD valign=top><b>Product&nbsp;Types:</b></td>
	<TD width="100%">
	<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 170; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 90px; BACKGROUND-COLOR: white" id=DIV2>
		<TABLE ID=TableTypes width=100%>
			<THEAD bgcolor=LightSteelBlue ><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT <%=CheckAllSelected%> id=chkTypeAll type=checkbox name=chkTypeAll LANGUAGE=javascript onclick="return onclick_CheckAll(lstType,chkTypeAll)"></TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Type&nbsp;</TD></THEAD>
			<%
			for i = lbound(TypeArray) to ubound(TypeArray)%>
			<TR onclick="onclick_Cell(lstType,'<%=TypeArray(i)%>');" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();">
					<TD>
					<% if instr("," & strSoftpaqType & ",","," & TypeArray(i) & "," ) > 0 then%>
						<INPUT id=lstType checked type=checkbox name=lstType value="<%=TypeArray(i)%>"></TD>
					<%else%>
						<INPUT id=lstType type=checkbox name=lstType value="<%=TypeArray(i)%>"></TD>
					<% end if%>
					<TD nowrap><%=TypeArray(i)%></TD>			
			</TR>

			<%next%>
		</TABLE>    
	</div>	

	</td>
</tr>

<tr>
	<TD valign=top><b>Products:</b></td>
	<TD width=100%>
		<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 150px; BACKGROUND-COLOR: white" id=DIV2>
		<TABLE ID=TableProducts width=100%>
			<THEAD bgcolor=LightSteelBlue ><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT checked id=chkProductAll type=checkbox name=chkProductAll LANGUAGE=javascript onclick="return onclick_CheckAll(lstProduct,chkProductAll)"></TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Product&nbsp;</TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;SystemID&nbsp;</TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Models&nbsp;</TD></THEAD>
			<%
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetProductsForRoot"
			
		
			Set p = cm.CreateParameter("@RootID", 3, &H0001)
			p.Value = RootID
			cm.Parameters.Append p
		
		
			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
			
			'rs.Open "spGetProductsForRoot " & RootID,cn,adOpenForwardOnly
			do while not rs.EOF
			%>
			<TR nowrap onclick="onclick_Cell(lstProduct,'<%=rs("ID")%>');" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();">
					<TD>
						<INPUT id=lstProduct checked type=checkbox name=lstProduct value="<%=rs("ID")%>"></TD>
					<TD><%=rs("Name") & " " & rs("Version")%></TD>			
					<TD><%=rs("SystemBoardID")%><INPUT type="hidden" id=txtSystemID name=txtSystemID value="<%=rs("SystemBoardID")%>"></TD>			
					<TD nowrap><%="TBD"%></TD>			
					
			</TR>

			<%
				rs.MoveNext
			loop
			rs.Close			
			%>
		</TABLE>    
	</div>	
	
</td>	
</tr>

<tr>
	<TD valign=top><b>Operating&nbsp;Systems:</b></td>
	<TD width=100%>
	
	<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIV1>
		<TABLE ID=TableDash width=100%>
			<THEAD bgcolor=LightSteelBlue  ><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT checked id=chkOSAll type=checkbox name=chkOSAll LANGUAGE=javascript onclick="return onclick_CheckAll(lstOS,chkOSAll)"></TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Official&nbsp;Name&nbsp;</TD><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">Minimum&nbsp;Level</TD></THEAD>
		<% 
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetSelectedOS"
	

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p


			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

'			rs.Open "spGetSelectedOS " & request("VersionID"),cn,adOpenForwardOnly
			
			do while not rs.eof
			%>
			<TR>
				<TD nowrap>
				<% if 1 then%>
					<INPUT id=lstOS checked type=checkbox name=lstOS value="<%=rs("OfficialName")%>"></TD>
				<%else%>
					<INPUT id=lstOS type=checkbox name=lstOS value="<%=rs("OfficialName")%>"></TD>			
				<%end if%>
				<TD nowrap onclick="onclick_Cell(lstOS,'<%=rs("OfficialName")%>');" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();"><%=rs("OfficialName")%></TD>
				<TD>
				<%
				if rs("ID") = 16 then
					blnOSInd = true
				end if
				response.write "<SELECT style=""width:70"" id=cboMinLevel name=cboMinLevel>"
				if rs("ID") = 1  then
					for i = lbound(Min95LevelArray) to ubound(Min95LevelArray)
						if Min95LevelArray(i) = "OSR0" then
							response.write "<OPTION selected value=""" & Min95LevelArray(i) & """>" & Min95LevelArray(i) & "</OPTION>"
						else
							response.write "<OPTION value=""" & Min95LevelArray(i) & """>" & Min95LevelArray(i) & "</OPTION>"
						end if
					next
				else
					for i = lbound(MinLevelArray) to ubound(MinLevelArray)
						if MinLevelArray(i) = "OEM" then
							response.write "<OPTION selected value=""" & MinLevelArray(i) & """>" & MinLevelArray(i) & "</OPTION>"
						else
							response.write "<OPTION value=""" & MinLevelArray(i) & """>" & MinLevelArray(i) & "</OPTION>"
						end if
					next
				end if
				Response.Write "</SELECT>"
				%>
				<INPUT type="hidden" id=txtOSKey name=txtOSKey value="<%=rs("CVAKey") & ""%>">
				</TD>
				
			</TR>
			<%
				rs.MoveNext
			loop
			rs.Close
			%>
		</TABLE>    
	</div>	
	
	</td>
</tr>

<%if trim(strMultiLanguage) = "1" then%>
<tr>
	<TD valign=top><b>Languages:</b></td>
	<TD width=100%>
	<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIV1>
		<TABLE ID=TableLanguage width=100%>
			<THEAD bgcolor=LightSteelBlue  ><TD nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT checked id=chkLanguageAll type=checkbox name=chkLanguageAll LANGUAGE=javascript onclick="return onclick_CheckAll(lstLanguage,chkLanguageAll)"></TD><TD style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Language&nbsp;</TD></THEAD>
		<% 
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetSelectedLanguages"
	

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = request("VersionID")
			cm.Parameters.Append p


			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing
			'rs.Open "spGetSelectedLanguages " & request("VersionID"),cn,adOpenForwardOnly
			
			do while not rs.eof
			%>
			<TR>
				<TD nowrap>
					<INPUT id=lstLanguage checked type=checkbox name=lstLanguage value="<%=rs("ID")%>"></TD>
				<TD nowrap onclick="onclick_Cell(lstLanguage,'<%=rs("ID")%>');" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();"><%=rs("Abbreviation") & " - " & rs("name")%>
				<INPUT type="hidden" id=txtLanguage name=txtLanguage value="<%=rs("Abbreviation") & " - " & rs("name")%>">
				</TD>
			</TR>
			<%
				rs.MoveNext
			loop
			rs.Close
			%>
		</TABLE>    
	</div>	
	
	
	</td>
</tr>
<%end if%>

<tr>
	<TD valign=top><b>Enhancements:</b></td>
		<TD width=100%><TEXTAREA rows=3 style="width:100%" id=txtEnhancements name=txtEnhancements><%=strEnhancements%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>Fixes:</b></td>
	<TD width=100%><TEXTAREA rows=3 style="width:100%" id=txtFixes name=txtFixes><%=strFixes%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>Prerequsites:</b></td>
	<TD width=100%><TEXTAREA rows=3 style="width:100%" id=txtPrerequsites name=txtPrerequsites><%=strPrerequsites%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>How&nbsp;To&nbsp;Use:</b></td>
	<TD width=100%><TEXTAREA rows=3 style="width:100%" id=txtHowToUse name=txtHowToUse><%=strHowToUse%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>Copyright:</b></td>
	<TD width=100%>Copyright (c) 2019 HP Development Company, L.P.</td>
</tr>
</table>
</span>


<span ID=tabCVA style="Display:none">
<table width=100% cellpadding=2 cellspacing=0 border=1 bordercolor=tan bgcolor=cornsilk>
<tr>
	<TD valign=top><b>CVA&nbsp;Version:</b></td>
	<TD width=100%><%=strCVAVersion%>
	<INPUT type="hidden" id=txtCVAVersion name=txtCVAVersion value="<%=strCVAVersion%>">
	</td>
</tr>
<tr>
	<TD valign=top><b>Title/Description:</b></td>
	<TD width=100%>
		<font size=1 face=verdana color=green>US title and description will be used for other languages if none provided.</font>
	<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 143px; BACKGROUND-COLOR: white" id=DIV1>
		<TABLE bgcolor=ivory ID=TableTranslations width=100%>
		<% 
			set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spGetLanguagesForRoot"
	

			Set p = cm.CreateParameter("@ID", 3, &H0001)
			p.Value = rootid
			cm.Parameters.Append p


			rs.CursorType = adOpenForwardOnly
			rs.LockType=AdLockReadOnly
			Set rs = cm.Execute 
			Set cm=nothing

			'rs.Open "spGetLanguagesForRoot " & rootid,cn,adOpenForwardOnly
			do while not rs.eof
			%>
			<TR width=100%>
				<TD nowrap width=120><b>Language:</b></td>
				<td><%=rs("name")%></td></tr>
					<tr><td><b>Title:</b></td><td><INPUT style="Width=100%" type=text name=txtCVATitle value="<%=trim(rs("Title") & "")%>&nbsp;" id="txtCVATitle"></TD></tr>			
					<tr><td valign=top><b>Description:</b></td><td>
					<TEXTAREA rows=3 style="width:100%" id=txtCVADescription name=txtCVADescription><%=trim(rs("Description") & "")%>&nbsp;</TEXTAREA>
					
					</TD></tr>
					
			</TR>
			<tr><td colspan=2><hr></td></tr>
			<%
				rs.MoveNext
			loop
			rs.Close
			%>
		</TABLE>    
	</div>	

	</td>
</tr>
<tr>
	<TD valign=top><b>Software&nbsp;Version:</b></td>
	<TD width=100%><%=strDelVersion%></td>
	</td>
</tr>
<tr>
	<TD valign=top><b>Software&nbsp;Revision:</b></td>
	<TD width=100%><%=strRevision%></td>
	</td>
</tr>
<tr>
	<TD valign=top><b>Software&nbsp;Pass:</b></td>
	<TD width=100%><%=strPass%></td>
	</td>
</tr>

<tr>
	<TD valign=top><b>Detailed&nbsp;File&nbsp;Info:</b></td>
	<TD width=100%>
	<TEXTAREA rows=3 style="width=100%" id=txtDetaileFileInfo name=txtDetailedFileInfo><%=strFileInfo%></TEXTAREA></td>
</tr>
<tr>
	<TD valign=top><b>Devices&nbsp;Supported:</b></td>
		<td width="100%" align="right">
			<textarea id="txtPNPDevices" style="DISPLAY:none;WIDTH: 100; HEIGHT: 100px" name="txtPNPDevices" rows="10" cols="62"><%=strPNPDevices%></textarea>
			<table border=0><tr>
			<td>
			<input style="WIDTH=50" type="button" value="Add" id="cmdNewDevice" name="cmdNewDevice" LANGUAGE="javascript" onclick="return cmdNewDevice_onclick()"><BR>
			<input style="WIDTH=50" type="button" value="Edit" id="cmdEditDevice" name="cmdEditDevice" LANGUAGE="javascript" onclick="return cmdEditDevice_onclick()"><BR>
			<input style="WIDTH=50" type="button" value="Delete" id="cmdDeleteDevice" name="cmdDeleteDevice" LANGUAGE="javascript" onclick="return cmdDeleteDevice_onclick()">			
			</td>
			<td width=100%>
			<select size="2" id="lstPNPDevices" name="lstPNPDevices" style="WIDTH: 100%; HEIGHT: 70px" LANGUAGE="javascript" ondblclick="return lstPNPDevices_ondblclick()">
			<%
				do while instr(strPNPDevices,vbcrlf) > 0
					Response.Write "<option>" & left(strPNPDevices,instr(strPNPDevices,vbcrlf)-1) & "</option>"
					strPNPDevices = mid(strPNPDevices,instr(strPNPDevices,vbcrlf)+ 2)
				loop
				if strPNPDevices <> "" then
					Response.Write "<option>" & strPNPDevices & "</option>"
				end if
			%>	
			</select>
			</td>
			</tr></table>
		</td>
</tr><tr>
	<TD valign=top><b>ARCD&nbsp;Install&nbsp;Execution:</b></td>
	<TD width=100%><input id="txtARCD" name="txtARCD" style="WIDTH: 100%; HEIGHT: 22px" size="28" value="<%=strARCDInstall%>"></td>
</tr>
<tr>
	<TD valign=top><b>Silent&nbsp;Install&nbsp;Execution:</b></td>
	<TD width=100%><input id="txtSilent" name="txtSilent" style="WIDTH: 100%; HEIGHT: 22px" size="28" value="<%=strSilentInstall%>"></td>
</tr>

</table>
</span>



<%
	



	cn.Close
	set cn = nothing
	set rs = nothing
	%>


<span ID=tabPreview style="Display:none">
Softpaq
<TEXTAREA style="width:100%;height=200" id=txtPreview name=txtPreview readOnly></TEXTAREA>
Welcome
<TEXTAREA style="width:100%;height=200" id=txtPreviewWelcome name=txtPreviewWelcome readOnly></TEXTAREA>
CVA
<TEXTAREA style="width:100%;height=200" id=txtPreviewCVA name=txtPreviewCVA readOnly></TEXTAREA>
</span>


<!--<TEXTAREA style="width:100%;height=200;Display:none" id=txtPreview1 name=txtPreview1 readOnly><%=strSoftpaqTop %></TEXTAREA>
<TEXTAREA style="width:100%;height=200;Display:none" id=txtPreview2 name=txtPreview2 readOnly><%=strSoftpaqMiddle%></TEXTAREA>
<TEXTAREA style="width:100%;height=200;Display:none" id=txtPreview3 name=txtPreview3 readOnly><%=strSoftpaqBottom%></TEXTAREA>
<TEXTAREA style="width:100%;height=200;Display:none" id=txtSoftpaqLines name=txtSoftpaqLines readOnly><%=strSoftpaqLines%></TEXTAREA>-->
<INPUT type="hidden" id=txtFilename name=txtFilename value= "<%=strFilename%>">
<INPUT type="hidden" id=txtID name=txtID value= "<%=request("VersionID")%>">
<INPUT type="hidden" id=txtDelVersion name=txtDelVersion value= "<%=strDelVersion%>">
<INPUT type="hidden" id=txtRevision name=txtRevision value= "<%=strRevision%>">
<INPUT type="hidden" id=txtPass name=txtPass value= "<%=strPass%>">
<INPUT type="hidden" id=txtSoftpaqList name=txtSoftpaqList value="">
</form>

<%end if%>
<TEXTAREA style="width:100%;Display:none" id=txtSelectedProducts name=txtSelectedProducts></TEXTAREA>

<INPUT type="hidden" id=txtOSInd name=txtOSInd value="<%=blnOSInd%>">
</BODY>
</HTML>


