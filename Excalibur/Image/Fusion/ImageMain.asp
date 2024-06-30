<%@ Language=VBScript %>
<HTML>
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="Pragma" content="no-cache"> 
    <meta http-equiv="Expires" content="0">
    <title>Image Pulsar</title>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link type="text/css"  href="../../style/shared.css" rel="stylesheet" />
    <link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">
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
        body
        {
            background-color:ivory;
        }
        .disabled 
        {
            opacity: 0.7;
            -ms-filter:"progid:DXImageTransform.Microsoft.Alpha(Opacity=70)";
            filter: alpha(opacity=70);
            -moz-opacity:0.7;
            -khtml-opacity: 0.7;
        }
        .hide 
        {
            display: none !important;
        }
        .show 
        {
            display: inline !important;
        }
    </STYLE>
    <!-- #include file="../../includes/bundleConfig.inc" -->
    <script type="text/javascript" src="../../includes/client/json2.js"></script>
    <script type="text/javascript" src="../../includes/client/json_parse.js"></script>
    <script src="scripts/imagemain_pulsar.js" type="text/javascript"></script>
    <script src="/Pulsar/Scripts/spin/spin.js"></script>
    <script src="/Pulsar/Scripts/spin/jquery.spin.js"></script>
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
		
		    tabPreview.style.display="none";

		    lblInstructions.innerText = "Select Regions.";
		    window.scrollTo(0,0);		
		    if (! FormLoading)
		    {
			    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
			    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
			    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
		    }

		    break;

		    case "Preview":
		    if (! DisplayedID == "")
			    steptext = "";
		    else
			    steptext = " (Step 3 of 3)";
		    lblTitle.innerText = "Preview" + steptext;
		    tabGeneral.style.display="none";
		    tabRegions.style.display="none";
		    tabPreview.style.display = "";

		    AllRegionsLink.style.display = "none";
		    ProductRegionsLink.style.display = "none";
		    ImageRegionsLink.style.display = "none";

		    lblInstructions.innerText = "Review the Information you entered for this Image Definition.";
		
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
	            tabPreview.style.display = "none";

	            lblInstructions.innerText = "Enter General Information for this Image Definition.";

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
	    SelectTab(hidDisplayTab.value);
        DisplayRegions("Product");
        FormLoading = false;

        //Add modal dialog code to body tag: ---
        modalDialog.load();

        //add jquery datepicker to date fields
        load_datePicker();

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

	    //Highight the selected tab
	    document.all("Cell"+strStep).style.display="none";
	    document.all("Cell"+strStep+"b").style.display="";
    
	    CurrentState = strStep;
	    ProcessState();
    }

    function DisplayRegions(type)
    {
        RegionDisplayType = type;

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


    function BuildPreview(){
        var strPreview = "";
        var strBrand = "";
        var i;

        if (typeof(AddImage.chkBrands) == "undefined")
            strBrand = "Not Specified";
        else if (typeof (AddImage.chkBrands.length) == "undefined") {
            if (AddImage.chkBrands.checked)
                strBrand = AddImage.txtBrandsText.value;
        }
        else {
            for (i=0;i<AddImage.chkBrands.length;i++)
                if (AddImage.chkBrands[i].checked)
                    strBrand = strBrand + ", " + AddImage.txtBrandsText[i].value;
        }
        if (strBrand.substring(0, 2) == ", ")
            strBrand = strBrand.substring(2);

	    if (AddImage.txtDCRRequired.value == "")
		    strPreview = strPreview + "APPROVED DCR: " + AddImage.cboDCR.options[AddImage.cboDCR.selectedIndex].text + "\r\r";

	    if (AddImage.txtProductDrop.value == "")
		    strPreview = strPreview + "PRODUCT DROP: Not Specified\r";
	    else
		    {
	        strPreview = strPreview + "PRODUCT DROP: " + AddImage.txtProductDrop.value + "\r";
		    }
	    strPreview = strPreview + "BRANDS: " + strBrand + "\r";
	    strPreview = strPreview + "OS: " + AddImage.cboOS.options[AddImage.cboOS.selectedIndex].text + "\r";
	    strPreview = strPreview + "STATUS: " + AddImage.cboStatus.options[AddImage.cboStatus.selectedIndex].text + "\r\r";
	    strPreview = strPreview + "RTM DATE: " + AddImage.txtRTMDate.value + "\r\r";
	    strPreview = strPreview + "COMMENTS: " + AddImage.txtComments.value + "\r\r";
	    strPreview = strPreview + "Releases for Operating System: " + AddImage.cboOSRelease.options[AddImage.cboOSRelease.selectedIndex].text + "\r\r";
	
	    for (i=0;i<AddImage.txtDisplay.length;i++)
		    if (AddImage.chkRegion[i].checked)
			    strPreview = strPreview + AddImage.txtDisplay[i].value + "\r";
	

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
	
	    AddImage.txtProductDrop.style.display = strShowEditBoxes;
	    lblProductDrop.style.display = strShowValues;


	    AddImage.cboOS.style.display = strShowEditBoxes;
	    lblOS.style.display = strShowValues;
	
	    AddImage.cboStatus.style.display = strShowEditBoxes;
	    lblStatus.style.display = strShowValues;
	
	    AddImage.txtRTMDate.style.display = strShowEditBoxes;
	    lblRTMDate.style.display = strShowValues;

	    AddImage.txtComments.style.display = strShowEditBoxes;
	    lblComments.style.display = strShowValues;

	    DeleteLink.style.display = strShowEditBoxes;
	
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

    function cboOSRelease_onchange() {
        AddImage.txtOSReleaseText.value = AddImage.cboOSRelease.options[AddImage.cboOSRelease.selectedIndex].text;
    }

    function cmdRTMDate_onclick(strID){
	    var strRC;
	    var strRelease;
	
	    strRC = window.showModalDialog("../../mobilese/today/caldraw1.asp",AddImage.txtRTMDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	    if (typeof(strRC) != "undefined")
		    AddImage.txtRTMDate.value=strRC;
    }

    function cmdAddOS_onlick(){
        modalDialog.open({ dialogTitle: 'Add Product OS', dialogURL: '../AddproductOS.asp?ProductID=' + AddImage.txtProdID.value + '', dialogHeight: 400, dialogWidth: 475, dialogResizable: false, dialogDraggable: true });
    }

    function cmdAddOSResult() {
        var strOutput;
        var i;
        var j;
        var OptionArray;

        strOutput = modalDialog.getArgument('productOS_query_array');
        strOutput = JSON.parse(strOutput);

        if (typeof (strOutput) != "undefined") {
            for (i = 0; i < strOutput.length; i++) {
                OptionArray = strOutput[i].split("^");
                AddImage.cboOS.options[AddImage.cboOS.length] = new Option(OptionArray[1], OptionArray[0]);
                AddImage.cboOS[AddImage.cboOS.length - 1].selected = true;
            }
        }
    }

    function msc_onclick(ImageId) {
        var strOutput;
        strOutput = window.showModalDialog("../SelectMstrSkuCompFrame.asp?ID=" + ImageId, "", "dialogWidth:350px;dialogHeight:150px;edge: Raised;center:Yes; help: No;resizable: Yes;status: No");
        if (typeof (strOutput) != "undefined") {
            var link = document.getElementById("msc" + ImageId);
            if (strOutput == "-- Use Default Value --")
                link.innerHTML = "[ Default ]";
            else
                link.innerHTML = strOutput;
        }

    }

    function cboImageType_onchange(){
        var i;
        var blnFound = false;

        if (AddImage.cboImageType.selectedIndex==1)
            {
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
            }
    }

        function PriorityChange(ID) {
            var myChk = document.getElementById("chkPublish" + ID);
            var myRow = document.getElementById("regionRow" + ID);
            if (!event.srcElement.checked) {
                myChk.style.display = "none";
                myChk.disabled = true;
                myRow.className = "Product";
            }
            else {
                myChk.style.display = "";
                myChk.disabled = false;
                myRow.className = "Image";
            }

        }
    //-->
    </SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()" style="overflow:auto">
<%

	dim cn
	dim rs
	dim cm
	dim p
	dim strBrandList
    dim strImageTypeList
	dim strOS
	dim strOSID
	dim strOSList
	dim blnFound 
	dim strProductDrop
	dim strRegionMatrix
	dim strPriorityList
	dim i
	dim strAllRoots
	dim strImageIDList
	dim strImageNameList 
	dim strImageTag
	dim OSID
	dim SWID
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
	dim PriorityArray
	dim blnDeleteOK
	dim strDisplayDelete
	dim blnSaveEditValue

	dim strAllRegions
	dim strProductRegions
	dim strImageRegions
	
	dim blnMarketingAdmin
	dim strRTMDate
	dim strComments
	dim strActiveColor
	dim CurrentUserSysAdmin
	dim strDevCenter
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
    dim strBrandsLoaded
    dim strBrandIDsLoaded

    dim osReleaseId
    dim strOSRelease
    dim strOSReleaseList

    'Harris, Valerie (5/31/2016) - PBI 19513/ Task 20989
    Dim bCopyWithTarget
    If Request("CopyTarget") <> "" Then
        If Request("CopyTarget") = "1" Then
            bCopyWithTarget = True
	    Else
		    bCopyWithTarget = False
	    End If
    Else
       bCopyWithTarget = False 
    End If

	CurrentUserSysAdmin = false
	CurrentUserPinPm = false
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
       	Response.Redirect "../../NoAccess.asp?Level=1"
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

	if not (rs.EOF and rs.BOF) then
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../../NoAccess.asp?Level=1"
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
	strProductDrop = ""
	strOS = ""
	strOSID = ""
	strType = ""
	strRTMDate = ""
	strComments = ""
	OSID = 0 
	strStatus = ""
	strStatusID = ""
	strProductOSList = ""
	DriveDefinitionId = ""
	DriveDefinitionName = ""
	DriveDefinitionList = ""
    strOSRelease = ""
    osReleaseId = 0
	
	if Request("ID") <> "" or Request("CopyID") <> "" then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetImageDefinitionFusion"
		

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
	
		if not (rs.EOF and rs.BOF) then
			strStatus = rs("Status") & ""
			strStatusID = rs("StatusID") & ""
			strProductDrop = rs("ProductDrop") & ""
			strOS = rs("OS") & ""
			strOSID = rs("OSID") & ""
            strImageTypeID = rs("ImageTypeID") & ""
            strImageTypeName = rs("ImageTypeName") & ""
			strRTMDate = rs("RTMDate") & ""
			strComments = trim(rs("Comments") & "")
			DriveDefinitionId = trim(rs("ImageDriveDefinitionId") & "")
			DriveDefinitionName = trim(rs("DriveName") & "")
            osReleaseId = rs("OSReleaseId")
            strOSRelease = rs("OSReleaseName")
		end if
		rs.Close
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
    




	'Load Status
	rs.Open "spListImageStatus",cn,adOpenForwardOnly
	strStatusList = ""
	blnFound = false
	do while not rs.EOF
		if request("CopyID") = "" and trim(strStatusID) = trim(rs("ID")) or ( (request("ID") = "" or request("CopyID") <> "") and rs("ID") = 1) then
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
    if request("ID") = "" or request("CopyID") <> "" then
        strStatus = "Not Released"
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
			elseif rs("Active") and instr(strProductOSList,"," & trim(rs("ID")) & ",") > 0 then
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

   'Load Brands
    strBrandsLoaded = ""
    strBrandIDsLoaded = ""
    if request("CopyID") <> "" then
        rs.Open "spListImageDefinitionBrandsAll " & clng(request("ProdID")) & "," & clng(request("CopyID")) ,cn,adOpenForwardOnly
    elseif request("ID") <> "" then
        rs.Open "spListImageDefinitionBrandsAll " & clng(request("ProdID")) & "," & clng(request("ID")) ,cn,adOpenForwardOnly
    else
        rs.Open "spListImageDefinitionBrandsAll " & clng(request("ProdID")) & ",0" ,cn,adOpenForwardOnly
    end if
    strBrandList = ""
    do while not rs.EOF
	    if rs("Selected") then
            strBrandList = strBrandList & "<tr><TD><INPUT checked type=""checkbox"" id=chkBrands name=chkBrands value="""  & rs("ID") & """>" & rs("Brand") & "<input id=""txtBrandsText"" type=""hidden"" value=""" & rs("Brand") & """ /></TD></tr>"
                strBrandIDsLoaded = strBrandIDsLoaded & "," & trim(rs("ID"))
                strBrandsLoaded = strBrandsLoaded & "," & trim(rs("Brand"))
        else
	        strBrandList = strBrandList & "<tr><TD><INPUT type=""checkbox"" id=chkBrands name=chkBrands value="""  & rs("ID") & """>" & rs("Brand") & "<input id=""txtBrandsText"" type=""hidden"" value=""" & rs("Brand") & """ /></TD></tr>"
        end if
	    rs.MoveNext
    loop
    rs.Close

    if strBrandIDsLoaded <> "" then
        strBrandIDsLoaded = mid(strBrandIDsLoaded,2)
    end if
    if strBrandsLoaded <> "" then
        strBrandsLoaded = mid(strBrandsLoaded,2)
    end if


    if strBrandList = "" then
        strBrandList = "<tr><td><font color=red>Warning:  You must select Brands on the Product properties screen first.</font></td></tr>"
    end if

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
	cm.CommandText = "spListRegionsForImage2"
		

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

	
	strRegionMatrix = "<TABLE  bgcolor=""cornsilk"" bordercolor=""tan"" border=""1"" cellpadding=""1"" cellspacing=""0"" width=""100%"" id=""regionTable"">"
	strImageIDList = ""
	strImageNameList = ""
	strImageTag = ""
	strCopyTag = ""
	strActiveColor =""
	strTabRowID = ""
	TabRowIndex = 0
	ImageMasterSkuComp = ""
    strTagPublish = ""
    strRegionMatrixBottom = ""
    strLastGeo = ""
	do while not rs.EOF
        strIssues = ""
        strTemp = ""
		if rs("Active") or trim(rs("Priority") & "" ) <> "" then
			strImageIDList = strImageIDList & "," & rs("ID")
			strImageNameList = strImageNameList & "," & rs("Name")
			if rs("optionconfig") & "" <> "" then
			    strImageNameList = strImageNameList & " (" & rs("optionconfig") & ")"
			end if
			if rs("Active") then
				strActiveColor = " bgcolor=""cornsilk"" "
            else
				strActiveColor = " bgcolor=""mistyrose"" " 'grey
			    strIssues = strIssues & "<BR>Localization is Inactive"
            end if

			if rs("Geo") & "" <> strLastGeo then
				strRegionMatrix = strRegionMatrix & "<TR bgcolor=""wheat"" class=""Header""><TD colspan=11><font color=black size=2 face=verdana><b>" & rs("Geo") & "</b></font></td></tr>"
		        strRegionMatrix = strRegionMatrix & "<TR bgcolor=""cornsilk"" class=""Header""><TD valign=bottom width=10>&nbsp;</TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>#Code</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom nowrap><font size=2 face=verdana><b>Name</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>OS&nbsp;Lang</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>GM&nbsp;</b></font></TD>"
		        if CurrentUserPinPm Then
			        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Master SKU Comp.</b></font></TD>"
                End If
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Issues&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Publish&nbsp;</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "</TR>"
            end if
			strLastGeo = rs("GEO") & ""
                
			Config = trim(rs("OptionConfig") & "")
            Priority = trim(rs("Priority") & "")
			if (clng(strDevCenter) = 2 and rs("Consumer")) or (clng(strDevCenter) <> 2 and rs("Commercial") )  or (trim(rs("Priority") & "" ) <> "") then
			    if trim(Priority) <> "" and instr(strConfigs, Config) = 0 then
			        StrRegionClass = "NotSupported"
			    elseif trim(Priority) <> "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Image"
			    elseif trim(Priority) = "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Product"
			    else 
                    StrRegionClass = "All"
			    end if
                if instr(strConfigs, Config) = 0 then
       			    strIssues = strIssues & "<BR>Localization is not supported"
                end if
                if ((clng(strDevCenter) <> 2 and rs("Consumer") and not rs("Commercial")) ) and (trim(rs("Priority") & "" ) <> "") then
                    strActiveColor = " bgcolor=""mistyrose"" " 
                    strIssues = strIssues & "<BR>Consumer Image Selected"
                elseif ((clng(strDevCenter) = 2 and rs("Commercial") and not rs("Consumer") )) and (trim(rs("Priority") & "" ) <> "") then
                    strActiveColor = " bgcolor=""mistyrose"" " 
                    strIssues = strIssues & "<BR>Commercial Image Selected"
                end if
			    strTemp = strTemp & "<TR" & strActiveColor & " id=""regionRow" & rs("ID") & """ class=""" & StrRegionClass & """>"
            else
			    strTemp = strTemp & "<TR" & strActiveColor & "style=""display:none"" class=""Hidden"">"
			end if
			if trim(rs("priority") & "") = "1" then
                strTemp = strTemp & "<TD><input onclick=""javascript: PriorityChange(" & rs("ID") & ");"" checked id=""chkRegion"" name=""chkRegion"" type=""checkbox"" value=""" & rs("ID") & """ /></TD>"
                strImageTag = strImageTag & "," & trim(rs("ID"))        
            else
                strTemp = strTemp & "<TD><input onclick=""javascript: PriorityChange(" & rs("ID") & ");"" id=""chkRegion"" name=""chkRegion"" type=""checkbox"" value=""" & rs("ID") & """ /></TD>"
            end if
			strTemp = strTemp & "<TD><font face=verdana size=2>" & rs("OptionConfig") & "&nbsp;</font></TD>"
			strTemp = strTemp & "<TD><INPUT type=""hidden"" id=txtDisplay name=txtDisplay value=""" & rs("DisplayName") & """><font face=verdana size=2 nowrap>" & rs("Name") & "</font></TD>"
			if trim(rs("OtherLanguage") & "") <> "" then
				strTemp = strTemp & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u>," & rs("OtherLanguage") & "</font></TD>"
			else
				strTemp = strTemp & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u> </font></TD>"
			end if
			strTemp = strTemp & "<TD><font face=verdana size=2>" & rs("GMCode") & "&nbsp;</font></TD>"
			If CurrentUserPinPm and rs("ImageId") & "" <> "" Then
			    ImageMasterSkuComp = rs("DriveName") & ""
			    If ImageMasterSkuComp = "" Then ImageMasterSkuComp = "[ Default ]"
			    strTemp = strTemp & "<TD><font face=verdana size=2><a href=""#"" id=""msc" & rs("ImageId") & """ onclick=""msc_onclick(" & rs("ImageId") & ");"">" & ImageMasterSkuComp & "</a>&nbsp;</font></TD>"
			ElseIf CurrentUserPinPm Then
			    strTemp = strTemp & "<TD>&nbsp;</TD>"
			End If
		    if left(strIssues,4) = "<BR>" then
                strIssues = mid(strIssues,5)
            end if
            strTemp = strTemp & "<TD><font face=verdana size=2>" & strIssues & "&nbsp;</font></TD>"
            strTemp = strTemp & "<TD align=""center"">"
            if not rs("Published") then
                if trim(rs("Priority") & "") <> "" then
                    strTemp = strTemp & "<input id=""chkPublish" & trim(rs("ID")) & """ value=""" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                else
                    strTemp = strTemp & "<input disabled style=""display:none"" id=""chkPublish" & trim(rs("ID")) & """ value=""" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                end if
            else
                if trim(rs("Priority") & "") <> "" then
                    strTemp = strTemp & "<input checked style=""display:" & strShowEditBoxes & """ id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" value=""" & trim(rs("ID")) & """ type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                else
                    strTemp = strTemp & "<input style=""display:none"" checked id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" value=""" & trim(rs("ID")) & """ type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                end if
            end if
            strTemp = strTemp & "&nbsp;</TD>"
			strTemp = strTemp & "</TR>"
			strRegionMatrix = strRegionMatrix & strTemp
		end if
		rs.MoveNext
	loop
	rs.Close

	strRegionMatrix = strRegionMatrix & "</TABLE>"

    'load Operating System Release list
    strOSReleaseList = ""
	rs.Open "SELECT Id, Description FROM OSRelease WHERE State = 1" ,cn,adOpenForwardOnly
	do while not rs.EOF
		if (osReleaseId = rs("Id")) then
			strOSReleaseList = strOSReleaseList & "<Option selected value=""" & rs("Id") & """>" & rs("Description") & "</Option>"
		else
			strOSReleaseList = strOSReleaseList & "<Option value=""" & rs("Id") & """>" & rs("Description") & "</Option>"
		end if
		rs.MoveNext
	loop
	rs.Close

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
    <td width=100% align=right><font size=1 face=verdana>&nbsp;</font></td>
    <td nowrap id=DeleteLink style="Display:<%=strDisplayDelete%>" align=right><font size=1 face=verdana>&nbsp;<a href="javascript:DisableImageDef();">Delete Image Definition</a></font></td>
  </tr>
</table>
<table ID="tabGeneral" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr style="Display:<%=strShowDCR%>">
		<td nowrap><b>Approved&nbsp;DCR:&nbsp;</b><font color="#ff0000" size="1">*</font></td>
		<td>
		<SELECT style="WIDTH:100%" id=cboDCR name=cboDCR LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()" onchange="return cboDCR_onchange()"><%=strDCRs%></SELECT>
		</td>
	</tr>
	<tr>
		<td nowrap><b>Product&nbsp;Drop:</b></td>
		<td width="100%">
			<LABEL ID=lblProductDrop Style="Display:<%=strShowValues%>"><%=strProductDrop%></LABEL>
			<INPUT style="WIDTH:200px;Display:<%=strShowEditBoxes%>" type="text" id=txtProductDrop name=txtProductDrop value="<%=strProductDrop%>">
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagProductDrop name=tagProductDrop value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagProductDrop name=tagProductDrop value="<%=strProductDrop%>">
			<%end if%>
		</td>
	</tr>

	
    <tr style="display:none">

		<td nowrap><b>Image&nbsp;Type:</b></td>
		<td>
			<LABEL ID=lblImageType Style="Display:<%=strShowValues%>"><%=strImageTypeName%></LABEL>
			<SELECT style="WIDTH:200px;Display:<%=strShowEditBoxes%>"  id=cboImageType name=cboImageType onchange="javascript: cboImageType_onchange();">
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
		<td valign="top" nowrap><b>Brands:</b> <font color="#ff0000" size="1">*</font></td>
		<td>

 
                                <table cellpadding=0 cellspacing=0>
                            <%
                            	                     response.write strBrandList

                            %>
                        </table>




		</td>
	</tr>
	<tr>
		<td nowrap><b>Operating&nbsp;System:</b>&nbsp;<font color="#ff0000" size="1">*</font>&nbsp;</td>
		<td>
			<LABEL ID=lblOS Style="Display:<%=strShowValues%>"><%=strOS%></LABEL>
			<SELECT style="WIDTH:200px;Display:<%=strShowEditBoxes%>"  id=cboOS name=cboOS LANGUAGE=javascript>
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
		<td nowrap><b>Status:</b> <font color="#ff0000" size="1">*</font></td>
		<td nowrap>
			<LABEL ID=lblStatus Style="Display:<%=strShowStatusValue%>"><%=strStatus%></LABEL>
			<SELECT style="WIDTH:200px;Display:<%=strShowStatusEdit%>" id=cboStatus name=cboStatus LANGUAGE=javascript onchange="return cboStatus_onchange()">
				<OPTION></OPTION>
				<%=strStatusList%>
			</SELECT>
			<%if request("CopyID") <> "" or request("ID") = "" then%>
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
			I have <a target=_blank" href="CompareFusionImage.asp?ImageDefinitionID=<%=request("ID")%>&PINTest=0&ProdID=<%=request("ProdID")%>">verified</a> these images are 100% accurate in IRS.
			</DIV>
		</td>
	</tr>
	<tr>
		<td nowrap><b>RTM Date:</b></td>
		<td>
			<LABEL ID=lblRTMDate Style="Display:<%=strShowValues%>"><%=strRTMDate%></LABEL>
			<INPUT type="text" style="WIDTH:200px;Display:<%=strShowEditBoxes%>" id=txtRTMDate name=txtRTMDate value="<%=strRTMDate%>" class="dateselection">
			
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
	        <select style="width:200px;display:<%=strShowEditBoxes%>" id="cboDriveDefinition" name="cboDriveDefinition">
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
				<INPUT style="WIDTH:350px" type="hidden" id=tagComments name=tagComments value="">
			<%else%>
				<INPUT style="WIDTH:350px" type="hidden" id=tagComments name=tagComments value="<%=strComments%>">
			<%end if%>
			
		</td>
	</tr>
    <tr>
        <td><b>Releases for Operating System: </b></td>
        <td>
            <LABEL ID=lblOSRelease Style="Display:<%=strShowValues%>"><%=strOSRelease%></LABEL>
			<SELECT style="WIDTH:200px;Display:<%=strShowEditBoxes%>"  id=cboOSRelease name=cboOSRelease LANGUAGE=javascript onchange="return cboOSRelease_onchange()">
				<OPTION></OPTION>
				<%=strOSReleaseList%>
			</SELECT>
			<%if request("CopyID") <> "" then%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagOSRelease name=tagOSRelease value="">
			<%else%>
				<INPUT style="WIDTH:200px" type="hidden" id=tagOSRelease name=tagOSRelease value="<%=osReleaseId%>">
			<%end if%>		
        </td>
    </tr>
</table>
<table ID="tabRegions" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="0" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td>
			<%=strRegionMatrix%>
		</TD>
	</tr>
</table>

<input style="Display:none" type="text" id="ID" name="ID" value="<%=request("ID")%>">



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

if strTagPublish<> "" then
    strTagPublish = mid(replace(strTagPublish," ",""),2)
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
<%if request("ID") = "" or request("CopyID") <> "" then %>
    <INPUT type="hidden" id=txtStatusTag name=txtStatusTag value="">
<%else%>
    <INPUT type="hidden" id=txtStatusTag name=txtStatusTag value="<%=strStatus%>">
<%end if%>
<INPUT type="hidden" id=txtStatusText name=txtStatusText value="<%=strStatus%>">
<INPUT type="hidden" id=txtBrandTag name=txtBrandTag value="<%=strBrandsLoaded%>">
<!--//<INPUT type="hidden" id=txtBrandIDTag name=txtBrandIDTag value="<%=strBrandIDsLoaded%>">//-->
<%if request("CopyID") <> "" then%>
    <INPUT type="hidden" id=txtBrandIDTag name=txtBrandIDTag value="">
<%else %>
    <INPUT type="hidden" id=txtBrandIDTag name=txtBrandIDTag value="<%=strBrandIDsLoaded%>">
<%end if %>
<INPUT type="hidden" id=txtBrandText name=txtBrandText value="<%=strBrandsLoaded%>">
<INPUT type="hidden" id=txtOSTag name=txtOSTag value="<%=strOS%>">
<INPUT type="hidden" id=txtOSText name=txtOSText value="<%=strOS%>">
<INPUT type="hidden" id=txtDevCenter name=txtDevCenter value="<%=trim(strDevCenter)%>">
<INPUT type="hidden" id="tagPublish" name="tagPublish" value="<%=strTagPublish%>">
<input type="hidden" id="inpCopyWithTarget" name="inpCopyWithTarget" value="<%=bCopyWithTarget%>" />
<INPUT type="hidden" id=txtOSReleaseTag name=txtOSReleaseTag value="<%=strOSRelease%>">
<INPUT type="hidden" id=txtOSReleaseText name=txtOSReleaseText value="<%=strOSRelease%>">
<INPUT type="hidden" id=txtOSReleaseId name=txtOSReleaseId value="<%=osReleaseId%>">
</form>


<form ID=DeleteImage method=post action="ImageDelete.asp?pulsarplusDiv=<%=Request("pulsarplusDivId")%>">
	<INPUT type="hidden" id=Auth name=Auth value="">
	<INPUT type="hidden" id=DelImageID name=DelImageID value="<%=request("ID")%>">
	<INPUT type="hidden" id=txtDelUserID name=txtDelUserID value="<%=CurrentUserID%>">
	<INPUT type="hidden" id=DelDCRID name=DelDCRID value="">
</form>
<INPUT type="hidden" id=txtDeleteOK name=txtDeleteOK value="<%=blnDeleteOK%>">
<input type="hidden" id="txtTabRowCount" name="txtTabRowCount" value="<%=TabRowIndex%>">
<INPUT type="hidden" id="hidDisplayTab" name="hidDisplayTab" value="<%=request("Tab")%>">
<INPUT type="hidden" id="hidDisplayMode" name="hidDisplayMode" value="<%=request("Mode")%>">
</BODY>
</HTML>
