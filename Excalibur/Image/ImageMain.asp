<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../includes/bundleConfig.inc" -->
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var CurrentState;
var States = new Array(2);
var FormLoading = true;

function ProcessState() {
	var steptext;

	switch (CurrentState)
	{
		case "Regions":
		if (! DisplayedID == "")
			steptext = "";
		else
			steptext = " (Step 2 of 3)";
	
		DeleteLink.style.display = "none";
		AllRegionsLink.style.display = "";
	    ProductRegionsLink.style.display = "";
	    ImageRegionsLink.style.display = "";
		
		lblTitle.innerText = "Regions" + steptext;
		tabGeneral.style.display="none";
		tabRegions.style.display="";
		
		//tabDeliverables.style.display="none";
		tabPreview.style.display="none";

		lblInstructions.innerText = "Enter the Priority for each Region. Leave unsupported Regions empty.";
		//AddImage.txtEndUser.focus();
		
		//for (i=1;i<txtTabRowCount.value;i++)
		//    {
		//    if(AddImage.cboOS.value=="19")
		//        window.document.all("TabRow" + i).style.display="";
		//    }
		window.scrollTo(0,0);		
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			window.parent.frames["LowerWindow"].cmdNext.disabled = false;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
		}

		break;

/*		case "Deliverables":
		if (! DisplayedID == "")
			steptext = "";
		else
			steptext = " (Step 3 of 4)";
		lblTitle.innerText = "Root Deliverables in Image" + steptext;
		tabGeneral.style.display="none";
		tabRegions.style.display="none";
		tabDeliverables.style.display="";
		tabPreview.style.display="none";

		lblInstructions.innerText = "Select the Root Deliverables for this Image Definition.";
		
		
		window.scrollTo(0,0);		
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			window.parent.frames["LowerWindow"].cmdNext.disabled = false;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
		}
		break;
*/
		case "Preview":
		if (! DisplayedID == "")
			steptext = "";
		else
			steptext = " (Step 3 of 3)";
		lblTitle.innerText = "Preview" + steptext;
		tabGeneral.style.display="none";
		tabRegions.style.display="none";
		//tabDeliverables.style.display="none";
		tabPreview.style.display = "";

		AllRegionsLink.style.display = "none";
		ProductRegionsLink.style.display = "none";
		ImageRegionsLink.style.display = "none";

		lblInstructions.innerText = "Review the Information you entered for this Image Definition.";
		//AddImage.txtEndUser.focus();
		
		window.scrollTo(0,0);		
		if (! FormLoading)
		{
			window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			window.parent.frames["LowerWindow"].cmdNext.disabled = true;
			window.parent.frames["LowerWindow"].cmdFinish.disabled = false;
		}
		break;

	    default:

	        AllRegionsLink.style.display = "none";
	        ProductRegionsLink.style.display = "none";
	        ImageRegionsLink.style.display = "none";
	        
	        if (!DisplayedID == "")
	            steptext = "";
	        else
	            steptext = " (Step 1 of 3)";

	        if (txtDeleteOK.value == "True")
	            DeleteLink.style.display = "";
	        else
	            DeleteLink.style.display = "none";

	        lblTitle.innerText = "General" + steptext;
	        tabGeneral.style.display = "";
	        tabRegions.style.display = "none";
	        //tabDeliverables.style.display="none";
	        tabPreview.style.display = "none";

	        lblInstructions.innerText = "Enter General Information for this Image Definition.";
	        //AddImage.txtEndUser.focus();

	        window.scrollTo(0, 0);
	        if (!FormLoading) {
	            window.parent.frames["LowerWindow"].cmdPrevious.disabled = true;
	            window.parent.frames["LowerWindow"].cmdNext.disabled = false;
	            window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
	        }
	        RegionDisplayType = "";
	        break;


	}
}
function window_onload() {
	var i;
	var strID;
	var strName;
	var hidDisplayTab = document.getElementById("hidDisplayTab");

	DisplayedID = AddImage.txtDisplayedID.value;
	Loaded = AddImage.txtDisplayedID.value;
	//CurrentState =  "General";
	//ProcessState();
	SelectTab(hidDisplayTab.value);
    DisplayRegions("Product");
    FormLoading = false;

    //add datepicker to date fields
    load_datepicker();

    //PBI 30442:TICKET#: 10852 - Unable to edit the RTM Date text box when Copy Image in Products
    try {
        // set the RTMDate be able to edit. because the jquery datepicker already set it readonly.
        $("#txtRTMDate").removeAttr('readonly');
    } catch (ee) {
    }

}

function SelectTab(strStep) {
	var i;
    if (strStep == "")
        strStep = "General";
	//Reset all tabs
	document.all("CellGeneralb").style.display="none";
	document.all("CellGeneral").style.display="";
	document.all("CellRegionsb").style.display="none";
	document.all("CellRegions").style.display="";
//	document.all("CellDeliverablesb").style.display="none";
//	document.all("CellDeliverables").style.display="";

	//Highight the selected tab
	document.all("Cell"+strStep).style.display="none";
	document.all("Cell"+strStep+"b").style.display="";
    
	CurrentState = strStep;
	ProcessState();
}

function DisplayRegions(type)
{
    RegionDisplayType = type;
    //SelectTab("Regions");

    switch (type) {
        case "All":
            linkDisplayAll.style.display = "none";
            linkDisplayAllText.style.display = "";
            linkDisplayProduct.style.display = "";
            linkDisplayProductText.style.display = "none";
            linkDisplayImage.style.display = "";
            linkDisplayImageText.style.display = "none";
    		break;
        case "Product":
            linkDisplayAll.style.display = "";
            linkDisplayAllText.style.display = "none";
            linkDisplayProduct.style.display = "none";
            linkDisplayProductText.style.display = "";
            linkDisplayImage.style.display = "";
            linkDisplayImageText.style.display = "none";
            break;
        case "Image":
            linkDisplayAll.style.display = "";
            linkDisplayAllText.style.display = "none";
            linkDisplayProduct.style.display = "";
            linkDisplayProductText.style.display = "none";
            linkDisplayImage.style.display = "none";
            linkDisplayImageText.style.display = "";
            break;
    }

    var regionTable = document.getElementById("regionTable");
    var rows = regionTable.getElementsByTagName("tr");

    for (var i = 0; i < rows.length; i++) {
        if ((rows[i].className != "Hidden") && (rows[i].className != "Header")) {
            if (type == "All") {
                rows[i].style.display = "";
            } else if (type == "Product") {
                if (rows[i].className == "Product" || rows[i].className == "Image" || rows[i].className == "NotSupported") {
                    rows[i].style.display = "";
                } else {
                    rows[i].style.display = "none";
                }
            } else if (type == "Image") {
                if (rows[i].className == "Image" || rows[i].className == "NotSupported") {
                    rows[i].style.display = "";
                } else {
                    rows[i].style.display = "none";
                }
            }
        }
    }
}

/*function lstAvailableRoots_ondblclick() {
	if (AddImage.lstAvailableRoots.options.selectedIndex > -1)
		{
		AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.length] = new Option(AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.options.selectedIndex].text,AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.options.selectedIndex].value);
		AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.options.selectedIndex] = null;
		}

}

function lstSelectedRoots_ondblclick() {
	if (AddImage.lstSelectedRoots.options.selectedIndex > -1)
		{
			AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.length] = new Option(AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.options.selectedIndex].text,AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.options.selectedIndex].value);
			AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.options.selectedIndex] = null;
		}
}


function cmdAddRoot_onclick() {
	var i;
	for (i=0;i<AddImage.lstAvailableRoots.length;i++)
		{
			if(AddImage.lstAvailableRoots.options[i].selected)
				{
					AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.length] = new Option(AddImage.lstAvailableRoots.options[i].text,AddImage.lstAvailableRoots.options[i].value);
				}
		}
	for (i=AddImage.lstAvailableRoots.length-1;i>=0;i--)
		{
			if(AddImage.lstAvailableRoots.options[i].selected)
					AddImage.lstAvailableRoots.options[i] = null;
		}


}

function cmdAddAllRoot_onclick() {
	var i;
	
	for (i=0;i<AddImage.lstAvailableRoots.length;i++)
		{
			AddImage.lstSelectedRoots.options[AddImage.lstSelectedRoots.length] = new Option(AddImage.lstAvailableRoots.options[i].text,AddImage.lstAvailableRoots.options[i].value);
		}
	for (i=AddImage.lstAvailableRoots.length-1;i>=0;i--)
		{
			AddImage.lstAvailableRoots.options[i] = null;
		}

}

function cmdRemoveRoot_onclick() {
	var i;
	
	for (i=0;i<AddImage.lstSelectedRoots.length;i++)
		{
			if(AddImage.lstSelectedRoots.options[i].selected)
				{
					AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.length] = new Option(AddImage.lstSelectedRoots.options[i].text,AddImage.lstSelectedRoots.options[i].value);
				}
		}
	for (i=AddImage.lstSelectedRoots.length-1;i>=0;i--)
		{
			if(AddImage.lstSelectedRoots.options[i].selected)
					AddImage.lstSelectedRoots.options[i] = null;
		}

}

function cmdRemoveAllRoot_onclick() {
	var i;
	
	for (i=0;i<AddImage.lstSelectedRoots.length;i++)
		{
			AddImage.lstAvailableRoots.options[AddImage.lstAvailableRoots.length] = new Option(AddImage.lstSelectedRoots.options[i].text,AddImage.lstSelectedRoots.options[i].value);
		}
	for (i=AddImage.lstSelectedRoots.length-1;i>=0;i--)
			AddImage.lstSelectedRoots.options[i] = null;

}

*/

function BuildPreview(){
	var strPreview = "";
	
	if (AddImage.txtDCRRequired.value == "")
		strPreview = strPreview + "APPROVED DCR: " + AddImage.cboDCR.options[AddImage.cboDCR.selectedIndex].text + "\r\r";

	if (AddImage.txtSKU.value == "")
		strPreview = strPreview + "SKU NUMBER: Not Specified\r";
	else
		{
		if (AddImage.txtSKUDigit.value == "")
			strPreview = strPreview + "SKU NUMBER: " + AddImage.txtSKU.value + "\r";
		else
			strPreview = strPreview + "SKU NUMBER: " + AddImage.txtSKU.value +  "-xx" + AddImage.txtSKUDigit.value + "\r";
				
		}
	strPreview = strPreview + "BRAND: " + AddImage.cboBrand.options[AddImage.cboBrand.selectedIndex].text + "\r";
	strPreview = strPreview + "OS: " + AddImage.cboOS.options[AddImage.cboOS.selectedIndex].text + "\r";
	strPreview = strPreview + "SOFTWARE: " + AddImage.cboSW.options[AddImage.cboSW.selectedIndex].text + "\r";
	strPreview = strPreview + "TYPE: " + AddImage.cboType.options[AddImage.cboType.selectedIndex].text + "\r";
	strPreview = strPreview + "STATUS: " + AddImage.cboStatus.options[AddImage.cboStatus.selectedIndex].text + "\r\r";
	strPreview = strPreview + "RTM DATE: " + AddImage.txtRTMDate.value + "\r\r";
	strPreview = strPreview + "COMMENTS: " + AddImage.txtComments.value + "\r\r";
	
	for (i=0;i<AddImage.txtDisplay.length;i++)
		if (AddImage.cboPriority[i].options[AddImage.cboPriority[i].selectedIndex].text != "")
			strPreview = strPreview + AddImage.txtDisplay[i].value + ": " + AddImage.cboPriority[i].options[AddImage.cboPriority[i].selectedIndex].text + "\r";
	

	AddImage.txtPreview.value = strPreview; 
}


function cboDCR_onchange() {
	var strShowValues = "";
	var strShowEditBoxes = "";
	
	DeleteImage.DelDCRID.value = AddImage.cboDCR.value;
	if (AddImage.cboDCR.value != "" && AddImage.txtDisplayedID.value != "")
		txtDeleteOK.value = "True";
	else
		txtDeleteOK.value = "False";

	if (AddImage.txtDisplayedID.value == "")
		return;
	
	if (AddImage.cboDCR.selectedIndex == 0)
		{
			strShowValues = "";
			strShowEditBoxes = "none";
		}
	else
		{
			strShowValues = "none";
			strShowEditBoxes = "";
		}
	
	AddImage.txtSKU.style.display = strShowEditBoxes;
	lblSKU.style.display = strShowValues;

	AddImage.cboBrand.style.display = strShowEditBoxes;
	lblBrand.style.display = strShowValues;

	AddImage.cboOS.style.display = strShowEditBoxes;
	lblOS.style.display = strShowValues;

	AddImage.cboSW.style.display = strShowEditBoxes;
	lblSW.style.display = strShowValues;
	
	AddImage.cboStatus.style.display = strShowEditBoxes;
	lblStatus.style.display = strShowValues;
	
	AddImage.cboType.style.display = strShowEditBoxes;
	lblType.style.display = strShowValues;
	
	AddImage.txtRTMDate.style.display = strShowEditBoxes;
	lblRTMDate.style.display = strShowValues;

	AddImage.txtComments.style.display = strShowEditBoxes;
	lblComments.style.display = strShowValues;

	DeleteLink.style.display = strShowEditBoxes;
	//RegionClearLink.style.display = strShowEditBoxes;
	
	//Regions
	
	for (i=0;i<AddImage.txtDisplay.length;i++)
		{
		AddImage.cboPriority[i].style.display = strShowEditBoxes;
		lblPriority[i].style.display = strShowValues;
		
		}
}

function DisableImageDef(){
	if (window.confirm ("Are you sure you want to delete this Image Definition?\r\rWARNING: This will remove all deliverables from these images in Excalibur.  If you decide to add these images back later, you will have to manually relink the deliverables to the new images.") == true)
		{

			window.parent.frames["LowerWindow"].cmdOK.disabled =true;
			window.parent.frames["LowerWindow"].cmdEditCancel.disabled =true;

			DeleteImage.Auth.value = "DeLeTeOk";
			DeleteImage.submit();
		}
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
		
		for (i=0;i<event.srcElement.length;i++)
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


function cboStatus_onchange() {
	AddImage.txtStatusText.value = AddImage.cboStatus.options[AddImage.cboStatus.selectedIndex].text;

	if (AddImage.cboStatus.options[AddImage.cboStatus.selectedIndex].value == "2" && AddImage.txtDevCenter.value != "2")
		divImagesValidated.style.display = "";
	else
		divImagesValidated.style.display = "none";
	
}

function cboBrand_onchange() {
	AddImage.txtBrandText.value = AddImage.cboBrand.options[AddImage.cboBrand.selectedIndex].text;
}

function cboOS_onchange() {
	AddImage.txtOSText.value = AddImage.cboOS.options[AddImage.cboOS.selectedIndex].text;
}

function cboSW_onchange() {
	AddImage.txtSWText.value = AddImage.cboSW.options[AddImage.cboSW.selectedIndex].text;
}

function cboType_onchange() {
	AddImage.txtTypeText.value = AddImage.cboType.options[AddImage.cboType.selectedIndex].text;
}

function cmdRTMDate_onclick(strID){
	var strRC;
	var strRelease;
	
	
	strRC = window.showModalDialog("../mobilese/today/caldraw1.asp",AddImage.txtRTMDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strRC) != "undefined")
		AddImage.txtRTMDate.value=strRC;
	
	

}

function cmdAddOS_onlick(){
	var strOutput;
	var i;
	var j;
	var OptionArray;

	strOutput = window.showModalDialog("AddproductOS.asp?ProductID=" + AddImage.txtProdID.value,"","dialogWidth:600px;dialogHeight:550px;edge: Raised;center:Yes; help: No;resizable: Yes;status: No"); 
	if (typeof(strOutput) != "undefined")
		{
		    for (i=0;i<strOutput.length;i++)
		        {
		        OptionArray = strOutput[i].split("^");
		        AddImage.cboOS.options[AddImage.cboOS.length] = new Option(OptionArray[1],OptionArray[0]);
		        AddImage.cboOS[AddImage.cboOS.length-1].selected = true;
		        }
		}
}

function msc_onclick(ImageId) {
    var strOutput;
    strOutput = window.showModalDialog("SelectMstrSkuCompFrame.asp?ID=" + ImageId, "", "dialogWidth:350px;dialogHeight:150px;edge: Raised;center:Yes; help: No;resizable: Yes;status: No");
    if (typeof (strOutput) != "undefined") {
        var link = document.getElementById("msc" + ImageId);
        link.innerHTML = strOutput;
    }

}

function cboImageType_onchange(){
    var i;
    var blnFound = false;

    if (AddImage.cboImageType.selectedIndex==1)
        {
        //Select NONE SW
        for (i=0;i<AddImage.cboSW.options.length;i++)
            {
                if (AddImage.cboSW.options[i].value == 32)
                    {
                    blnFound=true;
                    if (AddImage.cboSW.selectedIndex==0)
                        AddImage.cboSW.selectedIndex=i;
                    }
            }
        if (!blnFound)
            {
            AddImage.cboSW.options[AddImage.cboSW.options.length]=new Option("None", "32", true, true)
            }

        //Select All OS
        blnFound = false;
        for (i=0;i<AddImage.cboOS.options.length;i++)
            {
                if (AddImage.cboOS.options[i].value == 56)
                    {
                    blnFound=true;
                    if (AddImage.cboOS.selectedIndex==0)
                        AddImage.cboOS.selectedIndex=i;
                    }
            }
        if (!blnFound)
            {
                AddImage.cboOS.options[AddImage.cboOS.options.length] = new Option("All OS", "56", true, true)
            }

        //Select NONE Brand
        blnFound = false;
        for (i=0;i<AddImage.cboBrand.options.length;i++)
            {
                if (AddImage.cboBrand.options[i].value == 74)
                    {
                    blnFound=true;
                    if (AddImage.cboBrand.selectedIndex==0)
                        AddImage.cboBrand.selectedIndex=i;
                    }
            }
        if (!blnFound)
            {
            AddImage.cboBrand.options[AddImage.cboBrand.options.length]=new Option("None", "74", true, true)
            }

        //Select NONE Type
        blnFound = false;
        for (i=0;i<AddImage.cboType.options.length;i++)
            {
                if (AddImage.cboType.options[i].text == "None")
                    {
                    blnFound=true;
                    if (AddImage.cboType.selectedIndex==0)
                        AddImage.cboType.selectedIndex=i;
                    }
            }
        if (!blnFound)
            {
            AddImage.cboType.options[AddImage.cboType.options.length]=new Option("None", "None", true, true)
            }

        }
}

function PriorityChange(ID) {
    var myChk = document.getElementById("chkPublish" + ID);
    if (event.srcElement.selectedIndex == 0) {
        myChk.style.display = "none";
        myChk.disabled = true;
    }
    else {
        myChk.style.display = "";
        myChk.disabled = false;

    }

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
.NotSupported
{
	background-color: MistyRose;
}
.All
{
	color:Gray;
}

</STYLE>
<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()" style="overflow:auto">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">




<%

	dim cn
	dim rs
	dim cm
	dim p
	dim strBrand
	dim strBrandID
	dim strBrandList
    dim strImageTypeList
	dim strBrands
	dim strOS
	dim strOSID
	dim strOSList
	dim strSW
	dim strSWID
	dim strSWList
	dim strType
	dim blnFound 
	dim strSKU
	dim strRegionMatrix
	dim strPriorityList
	dim i
	dim strLastGeo
	dim strAllRoots
	dim strImageIDList
	dim strImageNameList 
	dim strImageTag
	dim OSID
	dim SWID
	dim BrandID
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserEmail
	dim CurrentWorkgroupID
	dim strPMID
	dim strSEPMID
	dim strShowValues
	dim strShowEditBoxes
	dim strShowDCR
	dim blnPOR
	dim strDCRs
	dim blnEditOK
	dim strStatus
	dim strStatusID
	dim strStatusList
	dim mstrCopyTag
	dim strRegionList
	dim PriorityArray
	dim blnDeleteOK
	dim strDisplayDelete
	dim blnSaveEditValue

	dim strAllRegions
	dim strProductRegions
	dim strImageRegions
	
	dim strSKUBase
	dim strSKUDigit
	dim blnMarketingAdmin
	dim strRTMDate
	dim strComments
	dim strActiveColor
	dim CurrentUserSysAdmin
	dim strDevCenter
	dim blnTabletOS
	dim strTabRowID
	dim TabRowIndex
	dim strProductOSList
	dim CurrentUserPinPm
	dim DriveDefinitionList
	dim DriveDefinitionId
	dim DriveDefinitionName
	dim ImageMasterSkuComp
	dim strTagPublish
    dim strImageTypeID
    dim strImageTypeName

	CurrentUserSysAdmin = false
	CurrentUserPinPm = false
	blnTabletOS = false
	blnOSIndependent = false
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
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
	if (rs.EOF and rs.BOF) then
		set rs = nothing
       	set cn=nothing
       	Response.Redirect "../NoAccess.asp?Level=1"
	else
        CurrentUserPartner = rs("PartnerID")	
		CurrentUserID = rs("ID") & ""
		CurrentUserEmail = rs("Email") & ""
		CurrentWorkgroupID = rs("WorkgroupID") & ""
		CurrentUserSysAdmin = rs("SystemAdmin")
	end if
	rs.Close

	
	'See if the user is a superuser
	rs.open "spListPMsActive 3",cn,adOpenForwardOnly
	do while not rs.eof
		if trim(CurrentUserID) = trim(rs("ID")) then
			CurrentUserSysAdmin = true
			exit do
		end if
		rs.movenext	
	loop
	rs.close	
	
	'See if the user is a Pin PM
	rs.open "spListPmsActive 5", cn, adOpenStatic
	do while not rs.eof
	    if trim(CurrentUserId) = trim(rs("ID")) then
	        CurrentUserPinPm = true
	        exit do
	    end if
	    rs.movenext
	loop
	rs.close
	
	blnPOR = false
	blnMarketingAdmin = false
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetProductVersion"
		

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ProdID")
	cm.Parameters.Append p
	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

'	rs.Open "spGetProductVersion " & request("ProdID"),cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../NoAccess.asp?Level=1"
			end if
		end if
	
	
		strSEPMID = rs("SEPMID") & ""
		strDevCenter = trim(rs("DevCenter") & "")
		strPMID = rs("PMID") & ""
        '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
		if trim(rs("ComMarketingID")& "") = trim(CurrentUserID) or trim(rs("SMBMarketingID")& "") = trim(CurrentUserID) or trim(rs("ConsMarketingID")& "") = trim(CurrentUserID) then
			blnMarketingAdmin = true
		end if
	end if
	rs.Close

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	if (request("ID") <> "" and (blnMarketingAdmin or CurrentUserSysAdmin or CurrentWorkgroupID = 15 or strSEPMID = CurrentUSerID or strPMID = CurrentUSerID)) or request("ID") = "" then
		blnEditOK = true
	else
		blnEditOK = false
	end if
	strShowDCR = "none"
	if (not (blnEditOK) or (blnPOR)) and request("ID") <> "" then
		strShowValues = ""
		strShowEditBoxes = "none"
		blnDeleteOK = false
	else
		strShowValues = "none"
		strShowEditBoxes = ""
		if request("CopyID") = "" and request("ID") <> "" then
			blnDeleteOK = true
		else
			blnDeleteOK = false	
		end if
	end if
	
	if blnDeleteOK then
		strDisplayDelete = ""
		strAllRegions = "none"
	    strProductRegions = "none"
	    strImageRegions = "none"
	else
		strDisplayDelete = "none"
		strAllRegions = ""
	    strProductRegions = ""
	    strImageRegions = ""
	end if

	if blnPOR and blnEditOK then
		strShowDCR = "" 
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovedDCRs"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ProdID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		
		'rs.Open "spListApprovedDCRs " & request("ProdID"),cn,adOpenForwardOnly
		strDCRs = "<Option selected></option>"
		do while not rs.EOF
			strDCRs = strDCRs & "<option value=""" & rs("ID") & """>" & rs("ID") & " - " & rs("Summary") & "</option>"
			rs.MoveNext
		loop
		rs.Close
	else
		strShowDCR = "none" 
	end if

		
	'Load Current Values
	strBrand = ""
	strBrandID = ""
	strSKU = ""
	strOS = ""
	strOSID = ""
	strSW = ""
	strSWID = ""
	strType = ""
	strLastGeo = ""
	strRTMDate = ""
	strComments = ""
	OSID = 0 
	SWID = 0
	BrandID = 0
	strStatus = ""
	strStatusID = ""
	strSKUBase = ""
	strSKUDigit = ""
	strBrands = ""
	strProductOSList = ""
	DriveDefinitionId = ""
	DriveDefinitionName = ""
	DriveDefinitionList = ""
	
	if Request("ID") <> "" or Request("CopyID") <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetImageDefinition"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		if Request("CopyID") <> "" then
			p.Value = Request("CopyID")
		else
			p.Value = Request("ID")
		end if
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
	
		'if Request("CopyID") <> "" then
		'	rs.Open "spGetImageDefinition " & Request("CopyID"),cn,adOpenForwardOnly
		'else
		'	rs.Open "spGetImageDefinition " & Request("ID"),cn,adOpenForwardOnly
		'end if
		if not (rs.EOF and rs.BOF) then
			strBrand = rs("Brand") & ""
			strBrandID = rs("BrandID") & ""
			strStatus = rs("Status") & ""
			strStatusID = rs("StatusID") & ""
			strSKU = rs("SKUNUmber") & ""
			strOS = rs("OS") & ""
			strOSID = rs("OSID") & ""
			strSW = rs("SW") & ""
			strSWID = rs("SWID") & ""
			strType = rs("ImageType") & ""
            strImageTypeID = rs("ImageTypeID") & ""
            strImageTypeName = rs("ImageTypeName") & ""
			strRTMDate = rs("RTMDate") & ""
			strComments = trim(rs("Comments") & "")
			DriveDefinitionId = trim(rs("ImageDriveDefinitionId") & "")
			DriveDefinitionName = trim(rs("DriveName") & "")
		end if
		rs.Close
	end if
	
	if request("CopyID") <> "" then
		strSKU = ""
	end if

	
'	'Load Drive Definitions
	rs.Open "usp_ListImageDriveDefinitions",cn,adOpenStatic
	blnFound = false
	Do While Not rs.EOF
	    if trim(DriveDefinitionId) = trim(rs("ID")) then
	        DriveDefinitionList = DriveDefinitionList & "<option selected value=""" & rs("ID") & """>" & rs("DriveName") & "</option>"
	        blnFound = true
	    else
	        DriveDefinitionList = DriveDefinitionList & "<option value=""" & rs("ID") & """>" & rs("DriveName") & "</option>"
	    end if
	    rs.MoveNext
	Loop
	rs.Close
	if (not blnfound) and request("ID") <> "" then
		DriveDefinitionList = DriveDefinitionList & "<Option selected value=""" & DriveDefinitionId & """>" & DriveDefinitionName & "</Option>" 
	end if
	
    'ImageTypes
    strImageTypeList = ""
	rs.Open "spListImageTypes" ,cn,adOpenForwardOnly
    do while not rs.eof
        if trim(strImageTypeID) = trim(rs("ID")) then
            strImageTypeList = strImageTypeList & "<option selected value=""" & rs("ID") & """>" & rs("name") & "</option>"
        elseif rs("active") then
            strImageTypeList = strImageTypeList & "<option value=""" & rs("ID") & """>" & rs("name") & "</option>"
        end if
        rs.movenext 
    loop
    rs.close
    
	'Load Brands
	'rs.Open "spListBrands" ,cn,adOpenForwardOnly
	rs.Open "spListBrands4Product " & clng(request("ProdID")) ,cn,adOpenForwardOnly
    blnFound = false
    strBrandList = ""
    if trim(strBrandID) = "69" then
    	strBrandList = strBrandList & "<Option selected value=""69"">All Supported Brands</Option>" 
		blnFound = true
		BrandID = "69" 'rs("ID")		
	else
		strBrandList = strBrandList & "<Option value=""69"">All Supported Brands</Option>" 
    end if
	do while not rs.EOF
		if trim(strBrandID) = trim(rs("ID")) then
			strBrandList = strBrandList & "<Option selected value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
			blnFound = true
			BrandID = rs("ID")			
		else
			strBrandList = strBrandList & "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
		end if
		rs.MoveNext
	loop
	rs.Close
	if (not blnfound) and request("ID") <> "" then
		strBrandList = strBrandList & "<Option selected value=""" & strBrandID & """>" & strBrand & "</Option>" 
	end if



	'Load Status
	rs.Open "spListImageStatus",cn,adOpenForwardOnly
	strStatusList = ""
	blnFound = false
	do while not rs.EOF
		if trim(strStatusID) = trim(rs("ID")) then
			strStatusList = strStatusList & "<Option selected value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
			blnFound = true
			StatusID = rs("ID")			
		else
			strStatusList = strStatusList & "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
		end if
		rs.MoveNext
	loop
	rs.Close
	if (not blnfound) and request("ID") <> "" then
		strStatusList = strStatusList & "<Option selected value=""" & strStatusID & """>" & strStatus & "</Option>" 
	end if

	'Load OS
	    'get OS list for product
    strProductOSList = ""
    rs.open "spListProductOS " & clng(request("ProdID")) & ",1",cn,adOpenForwardOnly
    do while not rs.eof
        strProductOSList = strProductOSList & "," & trim(rs("ID"))
        rs.movenext
    loop
    rs.close
    if strProductOSList <> "" then
        strProductOSList = strProductOSList & ","
    end if
	    

	    'load os list into combo
	
	rs.Open "spGetOS null, null",cn,adOpenForwardOnly
	strOSList = ""
	if trim(strOSID) = "19" then
	    blnTabletOS = true
	else
	    blnTabletOS = false
	end if
	blnFound = false
	do while not rs.EOF
		'if trim(rs("ID")) <> "16" then
			if trim(strOSID) = trim(rs("ID")) then
				strOSList = strOSList & "<Option selected value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
				blnFound = true
				OSID = rs("ID")
				if trim(strOSID) = 16 then
					blnOSIndependent = true
				end if
			elseif instr(strProductOSList,"," & trim(rs("ID")) & ",") > 0 then 'show all, not only active OS's
				strOSList = strOSList & "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
			end if
		'end if
		rs.MoveNext
	loop
	rs.Close
	if blnOSIndependent = false then
		strOSList = "<Option value=16>(OS Independent)</Option>" & strOSList
	end if

	if (not blnfound) and request("ID") <> "" then
		strOSList = strOSList & "<Option selected value=""" & strOSID & """>" & strOS & "</Option>" 
	end if

	'Load SW
	rs.Open "spListImageSWType",cn,adOpenForwardOnly
	strSWList = ""
	blnFound = false
	do while not rs.EOF
		if trim(strSWID) = trim(rs("ID")) then
			strSWList = strSWList & "<Option selected value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
			blnFound = true
			SWID = rs("ID")
		else
			strSWList = strSWList & "<Option value=""" & rs("ID") & """>" & rs("Name") & "</Option>" 
		end if
		rs.MoveNext
	loop
	rs.Close
	if (not blnfound) and request("ID") <> "" then
		strSWList = strSWList & "<Option selected value=""" & strSWID & """>" & strSW & "</Option>" 
	end if
	

	'Load Region List for Priority dropdowns
    strRegionList = ""
    if trim(request("ID")) = "" then
	    rs.Open "spListRegions",cn,adOpenForwardOnly
	    blnFound = false
	    do while not rs.EOF
		    if instr("," & strRegionList & ",","," & rs("Dash") & ",") = 0 then
			    strRegionList = strRegionList & "," &  rs("Dash")
		    end if
		    rs.MoveNext
	    loop
	    rs.Close

    elseif trim(request("ID") <> "") then
        rs.open "spListRegionsForImageDefIncludingInactive " & clng(request("ID") ),cn
        do while not rs.eof
    		if instr("," & strRegionList & ",","," & rs("Dash") & ",") = 0 then
                strRegionList = strRegionList & "," &  rs("Dash")
            end if
            rs.movenext 
        loop
        rs.close    
    end if
	 strRegionList = "0,1,2,3,4,5,6,7,8,9,10" & strRegionList
	PriorityArray  = split(strRegionList,",")
	
	'Build Deliverable List
'	strAllRoots = ""
'	strSelectedRoots = ""
'	rs.Open "spGetDelRoot",cn,adOpenForwardOnly
'	do while not rs.EOF
'		strAllRoots = strAllRoots & "<Option value=""" & rs("ID") & """>" & rs("name")& "</option>"
'		rs.MoveNext	
'	loop
'	rs.Close
'	

    'Build Product Localization List 
     rs.Open "usp_SelectProdBrandConfigs " & clng(request("ProdID")) ,cn,adOpenForwardOnly

     strConfigs = ""

     do while not rs.EOF
        if strConfigs = "" then
            strConfigs = rs("OptionConfig")
        else
	        strConfigs = strConfigs & "," & rs("OptionConfig")
	    end if
	    rs.MoveNext
	 loop 
	 rs.Close



	'Build Region Matrix
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListRegionsForImage"
		

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	if request("CopyID") <> "" then
		p.Value = request("CopyID")
	elseif request("ID") = "" then
		p.Value = 0
	else
		p.Value = request("ID")
	end if
	cm.Parameters.Append p
	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	
	'if request("CopyID") <> "" then
	'	rs.open "spListRegionsForImage " & request("CopyID"),cn,adOpenForwardOnly
	'elseif request("ID") = "" then
	'	rs.open "spListRegionsForImage 0",cn,adOpenForwardOnly
	'else
	'	rs.open "spListRegionsForImage " & request("ID"),cn,adOpenForwardOnly
	'end if
	strRegionMatrix = "<TABLE  bgcolor=""Cornsilk"" bordercolor=""tan"" border=""1"" cellpadding=""1"" cellspacing=""0"" width=""100%"" id=""regionTable"">"
	strLastGeo = ""
	strImageIDList = ""
	strImageNameList = ""
	strImageTag = ""
	strCopyTag = ""
	strActiveColor =""
	strTabRowID = ""
	TabRowIndex = 0
	ImageMasterSkuComp = ""
    strTagPublish = ""

	do while not rs.EOF
        strIssues = ""
		if rs("Active") or trim(rs("Priority") & "" ) <> "" then
			strBrands = ""
			strImageIDList = strImageIDList & "," & rs("ID")
			strImageNameList = strImageNameList & "," & rs("Name")
			if rs("dash") & "" <> "" then
			strImageNameList = strImageNameList & " (" & rs("Dash") & ")"
			end if
			if rs("Geo") & "" <> strLastGeo then
				strRegionMatrix = strRegionMatrix & "<TR bgcolor=""wheat"" class=""Header""><TD colspan=11><font color=black size=2 face=verdana><b>" & rs("Geo") & "</b></font></td></tr>"
				strLastGeo = rs("GEO") & ""
				strRegionMatrix = strRegionMatrix & "<TR class=""Header""><TD valign=bottom width=60><font size=2 face=verdana><b>Priority</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom nowrap><font size=2 face=verdana><b>Name</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>#Code</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Dash</b></font></TD>"
				'strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Cons</b></font></TD>"
				'strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>ENT</b></font></TD>"
				'strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>TAB</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>OS&nbsp;Lang</b></font></TD>"
				if CurrentUserPinPm Then
				    strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Master SKU Comp.</b></font></TD>"
			'	Else
			'	    strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Country</b></font></TD>"
			'	    strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Keyboard</b></font></TD>"
			'	    strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Power Cord</b></font></TD>"
                End If
				'strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>LocID</b></font></TD>"
				'strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>ImgID</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Issues&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Publish&nbsp;</b></font></TD>"
				strRegionMatrix = strRegionMatrix & "</TR>"
			end if
			if rs("Active") then
				strActiveColor = " bgcolor=""cornsilk"" "
            else
				strActiveColor = " bgcolor=""mistyrose"" " 'grey
			    strIssues = strIssues & "<BR>Localization is Inactive"
            end if
            
			strPriorityList = "<Select onchange=""javascript: PriorityChange(" & rs("ID") & ");""  ID=cboPriority Name=cboPriority style=""WIDTH:60;Display:" & strShowEditBoxes & """><OPTION selected></OPTION>"
			blnFound = false
			for i = lbound(PriorityArray) to ubound(PriorityArray)
				if trim(PriorityArray(i)) = trim(rs("Priority") & "" ) and not blnfound then
					strPriorityList = strPriorityList & "<OPTION selected>" & trim(PriorityArray(i)) & "</OPTION>"
					strImageTag = strImageTag & "," & trim(PriorityArray(i))
					blnFound = true
				else
					strPriorityList = strPriorityList & "<OPTION>" & trim(PriorityArray(i)) & "</OPTION>"
				end if
			next 

			
            if not blnFound and trim(rs("Priority") & "" ) <> "" then
			    strPriorityList = strPriorityList & "<OPTION selected>" & trim(rs("Priority") & "" )  & "</OPTION>"
                strImageTag = strImageTag & "," & trim(rs("Priority") & "" )
                strIssues = strIssues & "<BR>Selected Dash is Inactive"
            elseif not blnFound then
				strImageTag = strImageTag & ", "
            end if

			if rs("Consumer") then
				strBrands = strBrands & "<TD align=middle>X</TD>"
			else
				strBrands = strBrands & "<TD align=middle>&nbsp;</TD>"
			end if

			if rs("Commercial") then
				strBrands = strBrands & "<TD align=middle>X</TD>"
			else
				strBrands = strBrands & "<TD align=middle>&nbsp;</TD>"
			end if

			if rs("Tablet") then
				strBrands = strBrands & "<TD align=middle>X</TD>"
			else
				strBrands = strBrands & "<TD align=middle>&nbsp;</TD>"
			end if

			
			strCopyTag = strCopyTag & ", "
			strPriorityList = strPriorityList & "</Select>"
			
			'if rs("Tablet") and (not rs("Consumer")) and (not rs("Commercial")) then
            '	TabRowIndex = TabRowIndex + 1
	    	'	strTabRowID = " ID=TabRow" & trim(TabRowIndex) & " "
			'else
	    	'	strTabRowID = " "
			'end if
			
			Priority = trim(rs("Priority") & "" )
			Config = rs("OptionConfig")
			
			if (clng(strDevCenter) = 2 and (rs("Consumer") or (rs("Tablet") and blnTabletOS))) or (clng(strDevCenter) <> 2 and (rs("Commercial") or (rs("Tablet") and blnTabletOS))) or blnFound then
			    if Priority > "" and instr(strConfigs, Config) = 0 then
			        StrRegionClass = "NotSupported"
			    elseif trim(Priority) > "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Image"
			    elseif Priority = "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Product"
			    else 
                    StrRegionClass = "All"
			    end if
                if instr(strConfigs, Config) = 0 then
       			    strIssues = strIssues & "<BR>Localization is not supported"
                end if
			    strRegionMatrix = strRegionMatrix & "<TR" & strActiveColor & " id=""regionRow"" class=" & StrRegionClass & ">"
			else
			    strRegionMatrix = strRegionMatrix & "<TR" & strActiveColor & "style=""display:none"" class=""Hidden"">"
			end if
			strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=1>" & "<LABEL ID=lblPriority style=""Display:" & strShowValues & """>" & rs("Priority")  & "&nbsp;</LABEL>" & strPriorityList & "</font></TD>"
			strRegionMatrix = strRegionMatrix & "<TD><INPUT type=""hidden"" id=txtDisplay name=txtDisplay value=""" & rs("DisplayName") & """><font face=verdana size=2 nowrap>" & rs("Name") & "</font></TD>"
			strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("OptionConfig") & "&nbsp;</font></TD>"
			strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("Dash") & "</font></TD>"
			'strRegionMatrix = strRegionMatrix & strBrands
			if trim(rs("OtherLanguage") & "") <> "" then
				strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u>," & rs("OtherLanguage") & "</font></TD>"
			else
				strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u> </font></TD>"
			end if
			If CurrentUserPinPm and rs("ImageId") & "" <> "" Then
			    ImageMasterSkuComp = rs("DriveName") & ""
			    If ImageMasterSkuComp = "" Then ImageMasterSkuComp = "[ Default ]"
			    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2><a href=""#"" id=""msc" & rs("ImageId") & """ onclick=""msc_onclick(" & rs("ImageId") & ");"">" & ImageMasterSkuComp & "</a>&nbsp;</font></TD>"
			ElseIf CurrentUserPinPm Then
			    strRegionMatrix = strRegionMatrix & "<TD>&nbsp;</TD>"
			'Else
			'    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("CountryCode") & "</font></TD>"
			'    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("Keyboard") & "</font></TD>"
			'    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("PowerCord") & "&nbsp;</font></TD>"
			End If
'		    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("Id") & "&nbsp;</font></TD>"
'		    strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & rs("ImageId") & "&nbsp;</font></TD>"
		    if left(strIssues,4) = "<BR>" then
                strIssues = mid(strIssues,5)
            end if
            strRegionMatrix = strRegionMatrix & "<TD><font face=verdana size=2>" & strIssues & "&nbsp;</font></TD>"
            strRegionMatrix = strRegionMatrix & "<TD align=""center"">"
            if not rs("Published") then
                if trim(rs("Priority") & "") <> "" then
                    strRegionMatrix = strRegionMatrix & "<input id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                else
                    strRegionMatrix = strRegionMatrix & "<input disabled style=""display:none"" id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                end if
            else
                if trim(rs("Priority") & "") <> "" then
                    strRegionMatrix = strRegionMatrix & "<input checked style=""display:" & strShowEditBoxes & """ id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                else
                    strRegionMatrix = strRegionMatrix & "<input style=""display:none"" checked id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                end if
            end if
            strRegionMatrix = strRegionMatrix & "&nbsp;</TD>"
			strRegionMatrix = strRegionMatrix & "</TR>"
		end if
		rs.MoveNext
	loop
	rs.Close
	strRegionMatrix = strRegionMatrix & "</TABLE>"
	
	
	cn.Close
	set cn = nothing
	set rs = nothing
	
    if blnEditOK then
        strShowStatusEdit = ""
        strShowStatusValue = "none"
    else
        strShowStatusEdit = "none"
        strShowStatusValue = ""
    end if
    'You can't edit the image definition for RTM images, so override normal display
    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    if trim(StatusID) = "2" then
        blnEditOK = false
        blnDeleteOK = false
   		strShowValues = ""
		strShowEditBoxes = "none"
    end if


	if strSKU <> "" then
		if instr(strSKU,"-xx")> 0 then
			strSKUBase = left(strSKU,instr(strSKU,"-xx")-1)
			strSKUDigit = mid(strSKU,instr(strSKU,"-xx")+3)
		else	
			strSKUBase = strSKU
			strSKUDigit = ""
		end if
	else
		strSKUBase = ""
		strSKUDigit = ""
	end if
	
	dim strTitleColor
	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0
	
%>


<%if request("ID") <> "" then%>
<table Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0">
	<tr bgcolor="<%=strTitleColor%>">
		<td id="CellGeneral" style="Display:none" width="10"><font size="2" color="black"><b>&nbsp;<a href="javascript:SelectTab('General')">General</a>&nbsp;</b></font></td>
		<td id="CellGeneralb" style="Display:" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;General&nbsp;</b></font></td>
		<td id="CellRegions" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Regions')">Regions</a>&nbsp;</b></font></td>
		<td id="CellRegionsb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Regions&nbsp;</b></font></td>
<!--		<td id="CellDeliverables" style="Display:" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Deliverables')">Deliverables</a>&nbsp;</b></font></td>
		<td id="CellDeliverablesb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Deliverables&nbsp;</b></font></td>-->
	</tr>
</table>
<hr color="Tan">
<%else%>
<table><tr><td style="Display:none" id="CellGeneral"><td style="Display:none" id="CellGeneralb"><td style="Display:none" id="CellRegions"><td style="Display:none" id="CellRegionsb"></td></tr></table>
<%end if%>

<font face=verdana size=4><b>
<label ID="lblTitle"></label></b></font>

<form id="AddImage" method="post" action="saveimage.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">


<font size="2">
<label ID="lblInstructions"></label>
</font>

<Table width=100% border=0>
  <TR>
    <td nowrap id=ProductRegionsLink style="Display:<%=strProductRegions%>" align=left>
        <div id="linkDisplayProduct"  style="font-size:xx-small; font-family:Verdana"><a href="javascript:DisplayRegions('Product');">Product</a></div>
        <div id="linkDisplayProductText" style="display:none; font-size:xx-small; font-family:Verdana">Product</div>
    </td>
    <td nowrap id=ImageRegionsLink style="Display:<%=strImageRegions%>" align=left>
        <div id="linkDisplayImage" style="font-size:xx-small; font-family:Verdana">&nbsp;|&nbsp;<a href="javascript:DisplayRegions('Image');">Image</a></div>
        <div id="linkDisplayImageText" style="display:none; font-size:xx-small; font-family:Verdana">&nbsp;|&nbsp;Image</div>
    </td>
    <td nowrap id=AllRegionsLink style="Display:<%=strAllRegions%>" align=left>
        <div id="linkDisplayAll" style="display:none; font-size:xx-small; font-family:Verdana">&nbsp;|&nbsp;<a href="javascript:DisplayRegions('All');">All</a></div>
        <div id="linkDisplayAllText" style="font-size:xx-small; font-family:Verdana">&nbsp;|&nbsp;All</div>
    </td>
    <td width=100% align=right><font size=1 face=verdana><a target="_blank" href="http://teams1.sharepoint.hp.com/teams/BNBSE/Shared%20Documents/Forms/AllItems.aspx?RootFolder=%2fteams%2fBNBSE%2fShared%20Documents%2fImage%20Map%20Recovery&FolderCTID=&View=%7b48EB9556%2d6618%2d44D6%2d8C98%2dAF63EF347B20%7d">Image/Recovery Map</a></font></td>
    <td nowrap id=DeleteLink style="Display:<%=strDisplayDelete%>" align=right><font size=1 face=verdana>&nbsp;|&nbsp;<a href="javascript:DisableImageDef();">Delete Image Definition</a></font></td>
  </tr>
</table>
<table ID="tabGeneral" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
<col style="width:150px" /><col />
	<tr style="Display:<%=strShowDCR%>">
		<td nowrap><b>Approved&nbsp;DCR:&nbsp;</b><font color="#ff0000" size="1">*</font></td>
		<td>
		<SELECT style="WIDTH:100%" id=cboDCR name=cboDCR LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()" onchange="return cboDCR_onchange()"><%=strDCRs%></SELECT>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Image&nbsp;Number:</b></td>
		<td>
			<LABEL ID=lblSKU Style="Display:<%=strShowValues%>"><%=strSKU%></LABEL>
			<INPUT style="WIDTH:70px;Display:<%=strShowEditBoxes%>" type="text" id=txtSKU name=txtSKU value="<%=strSKUBase%>" maxlength=20><font size=2 face=verdana><b style="Display:<%=strShowEditBoxes%>">&nbsp;-xx&nbsp;</b></font><INPUT style="width=20;Display:<%=strShowEditBoxes%>" type="text" maxlength=1 id=txtSKUDigit name=txtSKUDigit value="<%=strSKUDigit%>">
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagSku name=tagSKU value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagSku name=tagSKU value="<%=strSKU%>">
			<%end if%>
		</td>
	</tr>
    <%  '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
        'if currentuserid = 31 or currentuserid = 8 or currentuserid = 674 or currentuserid=685 or currentuserid = 3082 then  
    %>
	    <!--//<tr>//-->
    <%'else%>
        <tr style="display:none">
    <%'end if%>
		<td nowrap><b>Image&nbsp;Type:</b></td>
		<td>
			<LABEL ID=lblImageType Style="Display:<%=strShowValues%>"><%=strImageTypeName%></LABEL>
			<SELECT style="WIDTH:300px;Display:<%=strShowEditBoxes%>"  id=cboImageType name=cboImageType onchange="javascript: cboImageType_onchange();">
				<%=strImageTypeList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagImageType name=tagImageType value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagImageType name=tagImageType value="<%=strImageTypeID%>">
			<%end if%>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Brand:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<LABEL ID=lblBrand Style="Display:<%=strShowValues%>"><%=strBrand%></LABEL>
			<SELECT style="WIDTH:300px;Display:<%=strShowEditBoxes%>"  id=cboBrand name=cboBrand LANGUAGE=javascript onchange="return cboBrand_onchange()">
				<OPTION></OPTION>
				<%=strBrandList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagBrand name=tagBrand value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagBrand name=tagBrand value="<%=BrandID%>">
			<%end if%>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Operating&nbsp;System:</b>&nbsp;<font color="#ff0000" size="1">*</font>&nbsp;</td>
		<td>
			<LABEL ID=lblOS Style="Display:<%=strShowValues%>"><%=strOS%></LABEL>
			<SELECT style="WIDTH:300px;Display:<%=strShowEditBoxes%>"  id=cboOS name=cboOS LANGUAGE=javascript onchange="return cboOS_onchange()">
				<OPTION></OPTION>
				<%=strOSList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagOS name=tagOS value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagOS name=tagOS value="<%=OSID%>">
			<%end if%>		
            <input style="display:<%=strShowEditBoxes%>" id="cmdAddOS" name="cmdAddOS" type="button" value="Add" onclick="cmdAddOS_onlick();">
		</td>
	</tr>
	<tr>
		<td nowrap><b>Software:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<LABEL ID=lblSW Style="Display:<%=strShowValues%>"><%=strSW%></LABEL>
			<SELECT style="WIDTH:300px;Display:<%=strShowEditBoxes%>"  id=cboSW name=cboSW LANGUAGE=javascript onchange="return cboSW_onchange()">
				<OPTION></OPTION>
				<%=strSWList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagSW name=tagSW value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagSW name=tagSW value="<%=SWID%>">
			<%end if%>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Build&nbsp;Type:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<LABEL ID=lblType Style="Display:<%=strShowValues%>"><%=strType%></LABEL>
			<SELECT style="WIDTH:150px;Display:<%=strShowEditBoxes%>"  id=cboType name=cboType LANGUAGE=javascript onchange="return cboType_onchange()">
				<OPTION selected></OPTION>
				<%if strType = "BTO" then%>
					<OPTION selected>BTO</OPTION>
				<%else%>
					<OPTION>BTO</OPTION>
				<%end if%>
				<%if strType = "CTO" then%>
					<OPTION selected>CTO</OPTION>
				<%else%>
					<OPTION>CTO</OPTION>
				<%end if%>
				<%if strType = "BTO/CTO" or  strType = "CTO/BTO" then%>
					<OPTION selected>CTO/BTO</OPTION>
				<%else%>
					<OPTION>CTO/BTO</OPTION>
				<%end if%>
				<%if strType = "None" then%>
					<OPTION selected>None</OPTION>
				<%end if%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagType name=tagType value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagType name=tagType value="<%=strType%>">
			<%end if%>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Status:</b> <font color="#ff0000" size="1">*</font></td>
		<td nowrap>
			<LABEL ID=lblStatus Style="Display:<%=strShowStatusValue%>"><%=strStatus%></LABEL>
			<SELECT style="WIDTH:150px;Display:<%=strShowStatusEdit%>" id=cboStatus name=cboStatus LANGUAGE=javascript onchange="return cboStatus_onchange()">
				<OPTION></OPTION>
				<%=strStatusList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagStatus name=tagStatus value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagStatus name=tagStatus value="<%=strStatusID%>">
			<%end if%>
			<DIV ID=divImagesValidated style="display:none">
			<%if strDevCenter <> "2" and strStatusID <> "2" then%>
				<INPUT type="checkbox" id=chkImagesValidated name=chkImagesValidated>
			<%else%>
				<INPUT checked type="checkbox" id=chkImagesValidated name=chkImagesValidated>
			<%end if%>
			I have <a target=_blank" href="CompareImage.asp?ImageDefinitionID=<%=request("ID")%>&PINTest=0&ProdID=<%=request("ProdID")%>">verified</a> these images are 100% accurate in Conveyor.
			</DIV>
		</td>
	</tr>
	<tr>
		<td nowrap><b>RTM Date:</b></td>
		<td>
			<LABEL ID=lblRTMDate Style="Display:<%=strShowValues%>"><%=strRTMDate%></LABEL>
			<INPUT type="text" style="WIDTH:100px;Display:<%=strShowEditBoxes%>" id=txtRTMDate name=txtRTMDate value="<%=strRTMDate%>" class="dateselection">
			<!--<a href="javascript: cmdRTMDate_onclick(<%=request("ID")%>)"><img style="Display:<%=strShowEditBoxes%>" ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a>-->
			
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagRTMDate name=tagRTMDate value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagRTMDate name=tagRTMDate value="<%=strRTMDate%>">
			<%end if%>
			
		</td>
	</tr>
	<tr>
	    <td style="white-space:nowrap;font-weight:bold;">Mstr. SKU Comp.:</td>
	    <td>
	        <label id="lblDriveDefinition" style="display:<%=strShowValues%>"><%=DriveDefinitionName %></label>
	        <select style="display:<%=strShowEditBoxes%>" id="cboDriveDefinition" name="cboDriveDefinition">
	            <option></option>
	            <%=DriveDefinitionList %>
	        </select>
	        <input type="hidden" id="tagDriveDefinition" name="tagDriveDefinition" value="<%If Request("CopyID") = "" Then Response.Write DriveDefinitionId %>" />
	    </td>
	</tr>
	<tr>
		<td nowrap><b>Comments:</b></td>
		<td>
			<LABEL ID=lblComments Style="Display:<%=strShowValues%>"><%=strComments%></LABEL>
			<INPUT type="text" style="WIDTH:100%;Display:<%=strShowEditBoxes%>" id=txtComments name=txtComments value="<%=strComments%>">
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagComments name=tagComments value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagComments name=tagComments value="<%=strComments%>">
			<%end if%>
			
		</td>
	</tr>
</table>
<table ID="tabRegions" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="0" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<!--<td width=80 nowrap valign=top><b>Regions:</b><BR><br><font size=1 face=verdana><a style="Display:<%=strShowEditBoxes%>" ID=RegionClearLink href="javascript: AddImage.reset();">Clear Regions</a></font></td>-->
		<td>
			<%=strRegionMatrix%>
		</TD>
	</tr>
</table>

<input style="Display:none" type="text" id="ID" name="ID" value="<%=request("ID")%>">
<!--<table ID="tabDeliverables" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td valign=top nowrap><b>Available Deliverables:</b><BR>
			<SELECT style="WIDTH:277px;" id=cboFilterRoots name=cboFilterRoots>
				<OPTION Selected>[Show Preinstalled Roots for this Product]</OPTION>
				<OPTION>[Show All Root Deliverables]</OPTION>
			</SELECT><HR>
			<SELECT style="WIDTH:277px;HEIGHT:360" size=22 id=lstAvailableRoots name=lstAvailableRoots multiple LANGUAGE=javascript ondblclick="return lstAvailableRoots_ondblclick()">
			<%=strAllRoots %>
			</SELECT>
		</td>
		<td valign=top width=10><BR>
			<INPUT style="Width=30" type="button" value=">" id=cmdAddRoot name=cmdAddRoot title="Add Selected" LANGUAGE=javascript onclick="return cmdAddRoot_onclick()"><BR>
			<INPUT style="Width=30" type="button" value=">>" id=cmdAddAllRoot name=cmdAddAllRoot title="Add All" LANGUAGE=javascript onclick="return cmdAddAllRoot_onclick()"><BR><BR>
			<INPUT style="Width=30" type="button" value="<" id=cmdRemoveRoot name=cmdRemoveRoot title="Remove Selected" LANGUAGE=javascript onclick="return cmdRemoveRoot_onclick()"><BR>
			<INPUT style="Width=30" type="button" value="<<" id=cmdRemoveAllRoot name=cmdRemoveAllRoot title="Remove All" LANGUAGE=javascript onclick="return cmdRemoveAllRoot_onclick()">
		</td>
		<td valign=top nowrap><b>Selected Deliverables:</b><BR>
			<SELECT style="WIDTH:277px;HEIGHT:400px" size=22 id=lstSelectedRoots name=lstSelectedRoots multiple LANGUAGE=javascript ondblclick="return lstSelectedRoots_ondblclick()">
			</SELECT>
		</td>
	</tr>
</table>-->



<table ID="tabPreview" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td nowrap><b>Preview:</b><br><textarea id="txtPreview" style="WIDTH: 100%; HEIGHT: 400px" name="txtPreview" cols="92"></textarea></td>
	</tr>
</table>
<%
	if len(strImageTag) > 0 then
		strImageTag = mid(strImageTag,2)
	end if

	if len(strCopyTag) > 0 then
		strCopyTag = mid(strCopyTag,2)
	end if

	if len(strImageIDList) > 0 then
		strImageIDList = mid(strImageIDList,2)
	end if

	if len(strImageNameList) > 0 then
		strImageNameList = mid(strImageNameList,2)
	end if


%>

<%if request("CopyID") <> "" then%>
	<INPUT type="hidden" style="WIDTH:100%" id=txtTag name=txtTag value="<%=strCopyTag%>">
<%else%>
	<INPUT type="hidden" style="WIDTH:100%" id=txtTag name=txtTag value="<%=strImageTag%>">
<%end if%>
<INPUT type="hidden" id=txtImageIDList name=txtImageIDList value="<%=strImageIDList%>">
<INPUT type="hidden" id=txtImageNameList name=txtImageNameList value="<%=strImageNameList%>">
<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtCopyID name=txtCopyID value="<%=request("CopyID")%>">
<INPUT type="hidden" id=txtProdID name=txtProdID value="<%=request("ProdID")%>">
<INPUT type="hidden" id=txtDCRRequired name=txtDCRRequired value="<%=strShowDCR%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtCurrentUserEmail name=txtCurrentUserEmail value="<%=CurrentUserEmail%>">
<INPUT type="hidden" id=txtStatusTag name=txtStatusTag value="<%=strStatus%>">
<INPUT type="hidden" id=txtStatusText name=txtStatusText value="<%=strStatus%>">
<INPUT type="hidden" id=txtBrandTag name=txtBrandTag value="<%=strBrand%>">
<INPUT type="hidden" id=txtBrandText name=txtBrandText value="<%=strBrand%>">
<INPUT type="hidden" id=txtOSTag name=txtOSTag value="<%=strOS%>">
<INPUT type="hidden" id=txtOSText name=txtOSText value="<%=strOS%>">
<INPUT type="hidden" id=txtSWTag name=txtSWTag value="<%=strSW%>">
<INPUT type="hidden" id=txtSWText name=txtSWText value="<%=strSW%>">
<INPUT type="hidden" id=txtTypeTag name=txtTypeTag value="<%=strType%>">
<INPUT type="hidden" id=txtTypeText name=txtTypeText value="<%=strType%>">
<INPUT type="hidden" id=txtDevCenter name=txtDevCenter value="<%=trim(strDevCenter)%>">

</form>

<%
if strTagPublish<> "" then
    strTagPublish = mid(replace(strTagPublish," ",""),2)
end if

 %>

<form ID=DeleteImage method=post action="ImageDelete.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>">
	<INPUT type="hidden" id=Auth name=Auth value="">
	<INPUT type="hidden" id=DelImageID name=DelImageID value="<%=request("ID")%>">
	<INPUT type="hidden" id=txtDelUserID name=txtDelUserID value="<%=CurrentUserID%>">
	<INPUT type="hidden" id=DelDCRID name=DelDCRID value="">
</form>
<INPUT type="hidden" id=txtDeleteOK name=txtDeleteOK value="<%=blnDeleteOK%>">
<input type="hidden" id="txtTabRowCount" name="txtTabRowCount" value="<%=TabRowIndex%>">
<INPUT type="hidden" id="hidDisplayTab" name="hidDisplayTab" value="<%=request("Tab")%>">
<INPUT type="hidden" id="hidDisplayMode" name="hidDisplayMode" value="<%=request("Mode")%>">
<INPUT type="hidden" id="tagPublish" name="tagPublish" value="<%=strTagPublish%>">

</BODY>
</HTML>
