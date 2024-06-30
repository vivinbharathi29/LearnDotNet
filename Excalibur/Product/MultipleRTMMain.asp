<%@ Language=VBScript %>

<%
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
    Dim AppRoot : AppRoot = Session("ApplicationRoot")	  
%>
<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<TITLE>Product RTM</TITLE>
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../includes/client/json2.js"></script>
<script type="text/javascript" src="../includes/client/json_parse.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    var SelectedBIOSRow;
    var CurrentState;
    var FormLoading = true;
    var intAlertStatus;
    var intProducts;

    intAlertStatus = 0;

    function UploadZip(ID){
        //save ID for return function: ---
        globalVariable.save(ID, 'main_uploadzip_ID');

        var sURL = "<%=AppRoot %>/PMR/SoftpaqFrame.asp?Title=Upload SCMX File&Page=<%=AppRoot %>/common/fileupload.aspx&KeepLocal=true";
        modalDialog.open({ dialogTitle: 'Upload SCMX File', dialogURL: '' + sURL + '', dialogHeight: 250, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
    }

    function UploadZip_return(strPath) {
        var strServer;
        var PathArray;
        var ID = globalVariable.get('main_uploadzip_ID');

        if (typeof (strPath) != "undefined") {
            PathArray = strPath.split("|");

            $("#UploadAddLinks" + ID).hide();
            $("#UploadRemoveLinks" + ID).show();
            $("#txtUploadPath" + ID).val(PathArray[0].substr(0,PathArray[0].lastIndexOf("\\")));
            $("#txtAttachmentPath" + ID).val(PathArray[1]);
            $("#UploadPath" + ID).text(PathArray[0].substr(PathArray[0].lastIndexOf("\\") + 1, PathArray[0].length));
        }

    }

    function RemoveUpload(ID){
        $("#UploadAddLinks" + ID).show();
        $("#UploadRemoveLinks" + ID).hide();
        $("#txtAttachmentPath" + ID).val("");
        $("#txtUploadPath" + ID).val("");
        $("#UploadPath" + ID).text("");
    }

    function txtRTMComments_onfocus() {
        frmMain.txtRTMComments.style.fontStyle = "normal";
        frmMain.txtRTMComments.style.color="black";
        if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value)
            frmMain.txtRTMComments.select();
    }


    function txtRTMComments_onblur() {
        if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value)
        {
            frmMain.txtRTMComments.style.fontStyle = "italic";
            frmMain.txtRTMComments.style.color="blue";
        }
    }

    function txtRestoreComments_onfocus() {
        frmMain.txtRestoreComments.style.fontStyle = "normal";
        frmMain.txtRestoreComments.style.color="black";
        if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value)
            frmMain.txtRestoreComments.select();
    }


    function txtRestoreComments_onblur() {
        if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value)
        {
            frmMain.txtRestoreComments.style.fontStyle = "italic";
            frmMain.txtRestoreComments.style.color="blue";
        }
    }

    function txtImageComments_onfocus() {
        frmMain.txtImageComments.style.fontStyle = "normal";
        frmMain.txtImageComments.style.color="black";
        if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value)
            frmMain.txtImageComments.select();
    }

    function txtImageComments_onblur() {
        if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value)
        {
            frmMain.txtImageComments.style.fontStyle = "italic";
            frmMain.txtImageComments.style.color="blue";
        }
    }

    function txtBIOSComments_onfocus() {
        frmMain.txtBIOSComments.style.fontStyle = "normal";
        frmMain.txtBIOSComments.style.color="black";
        if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value)
            frmMain.txtBIOSComments.select();
    }


    function txtBIOSComments_onblur() {
        if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value)
        {
            frmMain.txtBIOSComments.style.fontStyle = "italic";
            frmMain.txtBIOSComments.style.color="blue";
        }
    }

    function txtPatchComments_onfocus() {
        frmMain.txtPatchComments.style.fontStyle = "normal";
        frmMain.txtPatchComments.style.color="black";
        if (frmMain.txtPatchComments.value == frmMain.txtPatchCommentsTemplate.value)
            frmMain.txtPatchComments.select();
    }


    function txtPatchComments_onblur() {
        if (frmMain.txtPatchComments.value == frmMain.txtPatchCommentsTemplate.value)
        {
            frmMain.txtPatchComments.style.fontStyle = "italic";
            frmMain.txtPatchComments.style.color="blue";
        }
    }

    function cmdAdd_onclick(ProdIdx) {
        var strResult;
        var textboxNotify;
        textboxNotify = $("divPreview" + ProdIdx.toString()).find("#txtNotify");
        modalDialog.open({ dialogTitle: 'Add Email Address', dialogURL: '../Email/AddressBook.asp?AddressList=' + textboxNotify.value + '', dialogHeight: 350, dialogWidth: 540, dialogResizable: true, dialogDraggable: true });
        globalVariable.save('txtNotify', 'email_field');
    }

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

    var isDate = function (date) {
        return ((new Date(date)).toString() !== "Invalid Date") ? true : false;
    }

    function ValidateTab(strState){
        var blnSuccess;
        var FocusOn;
        var i;
        var blnFound;
        var intCount;
        var strNewVersion="";
        var BIOSArray=new Array(); 
        var RestoreArray=new Array();
        var PatchArray=new Array();
    
        blnSuccess = true;
	
        switch (strState)
        {
            case "General":
                blnFound = false;

                for (var j=0; j < intProducts; j++){
                    for (i=0;i<cboTitles.length;i++){
                        if (cboTitles[i].getAttribute("prodIdx") == j.toString()){
                            if (cboTitles[i].text.toLowerCase().replace(/^\s+|\s+$/g,"") == document.getElementById("txtRTMName" + j.toString()).value.toLowerCase().replace(/^\s+|\s+$/g,""))
                            {
                                blnFound = true;
                                FocusOn = document.getElementById("txtRTMName" + j.toString());
                                break;
                            }
                        }
                    }
                }

                var blnRtmBlank = false;
                for (var j=0; j < intProducts; j++){
                    if (document.getElementById("txtRTMName" + j.toString()).value == ""){
                        blnRtmBlank = true;
                        FocusOn = document.getElementById("txtRTMName" + j.toString());
                        break;
                    }
                }
                    
                if (blnRtmBlank && blnSuccess)
                {
                    window.alert("RTM Title is required.");
                    //FocusOn = frmMain.txtRTMName0;
                    blnSuccess = false;					
                }
                else if ((!blnRtmBlank && blnFound) && blnSuccess)
                {
                    window.alert("The RTM tile you entered was used on a previous RTM for these Products.");
                    //FocusOn = frmMain.txtRTMName0;
                    blnSuccess = false;					
                }
                else if ((frmMain.txtRTMDate.value == "") && blnSuccess)
                {
                    window.alert("RTM Date is required.");
                    FocusOn = frmMain.txtRTMDate;
                    blnSuccess = false;					
                }
                else if ((!isDate(frmMain.txtRTMDate.value)) && blnSuccess)
                {
                    window.alert("RTM Date must be a valid Date format.");
                    FocusOn = frmMain.txtRTMDate;
                    blnSuccess = false;					
                }
                else if ((!frmMain.chkBIOS.checked && !frmMain.chkRestore.checked && !frmMain.chkImages.checked && !frmMain.chkPatch.checked) && blnSuccess)
                {
                    window.alert("You must select the items to RTM.");
                    FocusOn = frmMain.txtRTMName0;
                    blnSuccess = false;					
                }
                else if ((frmMain.chkBIOS.checked && (!frmMain.optPhaseIn[0].checked) && (!frmMain.optPhaseIn[1].checked) && (!frmMain.optPhaseIn[2].checked)  ) && blnSuccess)
                {
                    window.alert("You must select BIOS Affectivity.");
                    FocusOn = frmMain.txtRTMName0;
                    blnSuccess = false;					
                }
                break;

            case "BIOS":
    
                blnFound = true;
                var blnEach =false;
                for(var j =0; j< intProducts; j++){
                    blnEach =false;
                    var tblBIOS = document.getElementById("tableBIOS" + j.toString());
                    var allChkBIOSList = new Array();
                    if (typeof($(tblBIOS).find('#chkBIOSList')) == "undefined"){
                        // no BIOS on this product
                        blnEach = true;
                    }else{
                        if (typeof($(tblBIOS).find('#chkBIOSList').length) == "undefined"){
                            allChkBIOSList[0] = $(tblBIOS).find('#chkBIOSList');
                        }else{
                            allChkBIOSList = $(tblBIOS).find('#chkBIOSList');
                        }
                        for (k=0; k<allChkBIOSList.length; k++ ){
                            if (allChkBIOSList[k].checked)
                            {
                                blnEach = true;;            
                            }
                        }
                    }
                    if (!blnEach){
                        blnFound = false;
                    }
                }


                if (!blnFound)
                {
                    window.alert("You must select at least one BIOS version for each product."); 
                    FocusOn = window.document;
                    blnSuccess = false;					
                }

		   
		    
                break;
            case "Restore":
    
                if (typeof(frmMain.chkRestoreList.length) == "undefined")
                    RestoreArray[0] = frmMain.chkRestoreList;
                else
                    RestoreArray = frmMain.chkRestoreList;
                blnFound = false;
                for (i=0;i<RestoreArray.length;i++)
                {
                    if (RestoreArray[i].checked)
                    {
                        blnFound = true;
                        break;
                    }
                }
		    
                if (!blnFound)
                {
                    window.alert("You must select at least one Restore Media version.");
                    FocusOn = window.document;
                    blnSuccess = false;					
                }

		    
                break;
            case "Patches":
                blnFound = true;
                var blnEach =false;
                for(var j =0; j< intProducts; j++){
                    blnEach =false;
                    var tblPatch = document.getElementById("tablePatch" + j.toString());
                    var allChkPatchList = new Array();
                    if (!tblPatch){
                        // no patch on this product
                        blnEach = true;
                    }else{
                        if (typeof($(tblPatch).find('#chkPatchList').length) == "undefined"){
                            allChkPatchList[0] = $(tblPatch).find('#chkPatchList');
                        }else{
                            allChkPatchList = $(tblPatch).find('#chkPatchList');
                        }
                        for (k=0; k<allChkPatchList.length; k++ ){
                            if (allChkPatchList[k].checked)
                            {
                                blnEach = true;;            
                            }
                        }
                    }
                    if (!blnEach){
                        blnFound = false;
                    }
                }


                if (!blnFound)
                {
                    window.alert("You must select at least one Patch for each product."); //Herb
                    FocusOn = window.document;
                    blnSuccess = false;					
                }

		    
                break;
            case "Alerts":
                var intShow;
                intShow = (intProducts - intAlertStatus - 1);
                var divAlertPage;
                divAlertPage = document.getElementById("divProductAlerts" + intShow.toString());
    
                if (!$(divAlertPage).find('#chkBuildLevel').is(":checked"))
                {
                    window.alert("You must signoff on the Build Level Alerts.");
                    FocusOn = $(divAlertPage).find('#chkBuildLevel');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkDistribution').is(":checked"))
                {
                    window.alert("You must signoff on the Distribution Alerts.");
                    FocusOn = $(divAlertPage).find('#chkDistribution');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkCertification').is(":checked"))
                {
                    window.alert("You must signoff on the Certification Alerts.");
                    FocusOn = $(divAlertPage).find('#chkCertification');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkWorkflow').is(":checked"))
                {
                    window.alert("You must signoff on the Workflow Alerts.");
                    FocusOn = $(divAlertPage).find('#chkWorkflow');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkAvailability').is(":checked"))
                {
                    window.alert("You must signoff on the Availability Alerts.");
                    FocusOn = $(divAlertPage).find('#chkAvailability');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkDeveloper').is(":checked"))
                {
                    window.alert("You must signoff on the Developer Alerts.");
                    FocusOn = $(divAlertPage).find('#chkDeveloper');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkRoot').is(":checked"))
                {
                    window.alert("You must signoff on the Root Deliverable Alerts.");
                    FocusOn = $(divAlertPage).find('#chkRoot');
                    blnSuccess = false;					
                }
                else if (!$(divAlertPage).find('#chkOTSPrimary').is(":checked"))
                {
                    window.alert("You must signoff on the Primary OTS Alerts.");
                    FocusOn = $(divAlertPage).find('#chkOTSPrimary');
                    blnSuccess = false;					
                }
                                
                break;	    
            case "Images":
     
                blnFound = true;
                var blnEach =false;
                var intImgEachCount = 0;
                var intTmgEachCheck = 0;
                var arrImg;

                var allChkImage0 = document.getElementsByName("chkImage");
                var allChkImage;

                if (typeof(allChkImage0.length) == "undefined"){
                    allChkImage = new Array();
                    allChkImage[0] = allChkImage0;
                }else{
                    allChkImage = allChkImage0;
                }

                for(var j =0; j< intProducts; j++){
                    blnEach =false;

                    intImgEachCount = 0;
                    intImgEachCheck = 0;

                    arrImg = new Array();
                    var strProdId = document.getElementById("divPreview" + j.toString()).getAttribute("prodId");
                    for (var k =0; k<allChkImage.length; k++){
                        if((allChkImage[k].getAttribute("prodId") == strProdId)){
                            intImgEachCount++;
                            if (allChkImage[k].checked){
                                intImgEachCheck++;
                            }
                        }
                    }


                    if (intImgEachCount==0){
                        // no patch on this product
                        blnEach = true;
                    }else{
                        if (intImgEachCheck > 0 )
                        blnEach = true;            
                    }

                    if (!blnEach){
                        blnFound = false;
                    }
                }


                if (!blnFound)
                {
                    window.alert("You must select at least one Image for each product."); 
                    FocusOn = window.document;
                    blnSuccess = false;					
                }

		    
                break;
        }


        if (blnSuccess == false)
        {
            if (CurrentState != strState)
            {
                CurrentState = strState;
                ProcessState();
            }
            try{//Herb
                FocusOn.focus();
            }catch(err){
                try{
                    FocusOn[0].focus();
                }catch(err2){
                    //Herb
                }
            }
            
        }


        return blnSuccess;
    }

    function window_onload() {
	
        var i;
        var strID;
        var strName;
        if(typeof(frmMain) != "undefined")
        {
            setProductsNumber();//Herb
            setAlertPageNumber();

            if (txtImageCount.value == "0")
            {
                frmMain.chkImages.checked = false;
                frmMain.chkImages.disabled = true;
                ImagesDisabled.innerHTML="&nbsp;"//"&nbsp;(None&nbsp;Available)"
                ImagesTextColor.color = "darkgray"
            }
            else
            {
                frmMain.chkImages.disabled = false;
                ImagesDisabled.innerHTML=""
                ImagesTextColor.color = "black"
            }	    
	   
            if (txtBIOSCount.value == "0")
            {
                frmMain.chkBIOS.checked = false;
                frmMain.chkBIOS.disabled = true;
                BIOSDisabled.innerHTML="&nbsp;"//"&nbsp;(None&nbsp;Available)"
                BIOSTextColor.color = "darkgray"
            }
            else
            {
                frmMain.chkBIOS.disabled = false;
                BIOSDisabled.innerHTML=""
                BIOSTextColor.color = "black"
            }	    
	        
            if (txtRestoreCount.value == "0")
            {
                frmMain.chkRestore.checked = false;
                frmMain.chkRestore.disabled = true;
                RestoreDisabled.innerHTML="&nbsp;"//"&nbsp;(None&nbsp;Available)"
                RestoreTextColor.color = "darkgray"
            }
            else
            {
                frmMain.chkRestore.disabled = false;
                RestoreDisabled.innerHTML=""
                RestoreTextColor.color = "black"
            }	    


            if (txtPatchCount.value == "0")
            {
                frmMain.chkPatch.checked = false;
                frmMain.chkPatch.disabled = true;
                PatchDisabled.innerHTML="&nbsp;"//"&nbsp;(None&nbsp;Available)"
                PatchTextColor.color = "darkgray"
            }
            else
            {
                frmMain.chkPatch.disabled = false;
                PatchDisabled.innerHTML=""
                PatchTextColor.color = "black"
            }	    

            CurrentState =  "General";
            ProcessState();
            FormLoading = false;	
        }else{
            window.parent.frames["LowerWindow"].cmdNext.disabled=true;
        }

        //Add modal dialog code to body tag: ---
        modalDialog.load();

        //load date picker: ---
        load_datePicker();
    
    }

    function LoadAlerts(strType){
        return;//herb 12/23
        if (!frmMain.chkBuildLevel.checked)
            document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value + strType;
        if (!frmMain.chkDistribution.checked)
            document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value + strType;
        if (!frmMain.chkCertification.checked)
            document.all.CertificationIFrame.src = frmMain.txtCertificationIFramesrc.value + strType;
        if (!frmMain.chkWorkflow.checked)
            document.all.WorkflowIFrame.src = frmMain.txtWorkflowIFramesrc.value + strType;
        if (!frmMain.chkAvailability.checked)
            document.all.AvailabilityIFrame.src = frmMain.txtAvailabilityIFramesrc.value + strType;
        if (!frmMain.chkDeveloper.checked)
            document.all.DeveloperIFrame.src = frmMain.txtDeveloperIFramesrc.value + strType;
        if (!frmMain.chkRoot.checked)
            document.all.RootIFrame.src = frmMain.txtRootIFramesrc.value + strType;
        if (!frmMain.chkOTSPrimary.checked)
            document.all.OTSPrimaryIFrame.src = frmMain.txtOTSPrimaryIFramesrc.value + strType;
        //if (!frmMain.chkOTSRelated.checked)
        //    document.all.OTSRelatedIFrame.src = frmMain.txtOTSRelatedIFramesrc.value + strType;
    
        if (strType == "&ReportType=2")
        {
            // OTSRelatedTypeText.innerText = "System BIOS deliverables"
            OTSPrimaryTypeText.innerText = "System BIOS deliverables"
            RootTypeText.innerText = "(System BIOS deliverables)"
            DeveloperTypeText.innerText = "(System BIOS deliverables)"
            AvailabilityTypeText.innerText = "(System BIOS deliverables)"
            WorkflowTypeText.innerText = "(System BIOS deliverables)"
            CertificationTypeText.innerText = "(System BIOS deliverables)"
            DistributionTypeText.innerText = "(System BIOS deliverables)"
            BuildLevelTypeText.innerText = "(System BIOS deliverables)"
        }
        else
        {
            //  OTSRelatedTypeText.innerText = "SW, FW, and Doc deliverables"
            OTSPrimaryTypeText.innerText = "SW, FW, and Doc deliverables"
            RootTypeText.innerText = "(SW, FW, and Doc deliverables)"
            DeveloperTypeText.innerText = "(SW, FW, and Doc deliverables)"
            AvailabilityTypeText.innerText = "(SW, FW, and Doc deliverables)"
            WorkflowTypeText.innerText = "(SW, FW, and Doc deliverables)"
            CertificationTypeText.innerText = "(SW, FW, and Doc deliverables)"
            DistributionTypeText.innerText = "(SW, FW, and Doc deliverables)"
            BuildLevelTypeText.innerText = "(SW, FW, and Doc deliverables)"
        }
    }

    
    function LoadAlertsByProduct(intAlertPageNumber,strType){
        
        strType = strType || "";
        var divAlertPage;
        var strIdx;
        strIdx = intAlertPageNumber.toString();
        divAlertPage = document.getElementById("divProductAlerts" + strIdx);
        
        if (!$(divAlertPage).find('#chkBuildLevel').is(":checked"))
            $(divAlertPage).find('#BuildLevelIFrame').attr("src",$(divAlertPage).find('#txtBuildLevelIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkDistribution').is(":checked"))
            $(divAlertPage).find('#DistributionIFrame').attr("src", $(divAlertPage).find('#txtDistributionIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkCertification').is(":checked"))
            $(divAlertPage).find('#CertificationIFrame').attr("src", $(divAlertPage).find('#txtCertificationIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkWorkflow').is(":checked"))
            $(divAlertPage).find('#WorkflowIFrame').attr("src", $(divAlertPage).find('#txtWorkflowIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkAvailability').is(":checked"))
            $(divAlertPage).find('#AvailabilityIFrame').attr("src", $(divAlertPage).find('#txtAvailabilityIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkDeveloper').is(":checked"))
            $(divAlertPage).find('#DeveloperIFrame').attr("src", $(divAlertPage).find('#txtDeveloperIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkRoot').is(":checked"))
            $(divAlertPage).find('#RootIFrame').attr("src", $(divAlertPage).find('#txtRootIFramesrc').val() + strType);
        if (!$(divAlertPage).find('#chkOTSPrimary').is(":checked"))
            $(divAlertPage).find('#OTSPrimaryIFrame').attr("src", $(divAlertPage).find('#txtOTSPrimaryIFramesrc').val() + strType);
       
        if (strType == "&ReportType=2")
        {
            var biostitle = "System BIOS deliverables";           
            $(divAlertPage).find('#OTSPrimaryTypeText').val(biostitle);
            $(divAlertPage).find('#RootTypeText').val(biostitle);
            $(divAlertPage).find('#DeveloperTypeText').val(biostitle);
            $(divAlertPage).find('#AvailabilityTypeText').val(biostitle);
            $(divAlertPage).find('#WorkflowTypeText').val(biostitle);
            $(divAlertPage).find('#CertificationTypeText').val(biostitle);
            $(divAlertPage).find('#DistributionTypeText').val(biostitle);
            $(divAlertPage).find('#BuildLevelTypeText').val(biostitle);
        }
        else
        {
             var biostitle = "SW, FW, and Doc deliverables";
            $(divAlertPage).find('#OTSPrimaryTypeText').val(biostitle);
            $(divAlertPage).find('#RootTypeText').val(biostitle);
            $(divAlertPage).find('#DeveloperTypeText').val(biostitle);
            $(divAlertPage).find('#AvailabilityTypeText').val(biostitle);
            $(divAlertPage).find('#WorkflowTypeText').val(biostitle);
            $(divAlertPage).find('#CertificationTypeText').val(biostitle);
            $(divAlertPage).find('#DistributionTypeText').val(biostitle);
            $(divAlertPage).find('#BuildLevelTypeText').val(biostitle);
        }
    }

    //Herb
    function LoadAlertsByProductAll(strType){
        
        strType = strType || "";
        
        for (var i =0; i< intProducts; i++){
            LoadAlertsByProduct(i,strType);
        }
    }

    //Herb
    function ProcessAlertDisplay() {
       
        var intShow;
        intShow = 0;
        //from 0 to intProducts-1
        if (CurrentState == "Alerts"){
            for (var i =0; i< intProducts; i++){
                intShow = intProducts - intAlertStatus - 1
                if (intShow == i){
                    document.getElementById("divProductAlerts" + i.toString()).style.display="";
                    //LoadAlertsByProduct(i);
                }else{
                    document.getElementById("divProductAlerts" + i.toString()).style.display="none";
                }

            }
        }else{
            for (var i =0; i< intProducts; i++){
                document.getElementById("divProductAlerts" + i.toString()).style.display="none";
            }
        }

    }

    //Herb
    function setAlertPageNumber() {
        intAlertStatus = parseInt(document.getElementById("txtProductCount").value) -1;
    }

    //Herb
    function setProductsNumber(){
        intProducts=parseInt(document.getElementById("txtProductCount").value);
    }

    function ProcessState() {
        
        switch (CurrentState)
        {
            case "General":
                lblTitle.innerText = "Enter General RTM information";
                if (frmMain.txtRTMComments.value == "")
                {
                    frmMain.txtRTMComments.value = frmMain.txtRTMCommentsTemplate.value;
                    frmMain.txtRTMComments.style.fontStyle = "italic";
                    frmMain.txtRTMComments.style.color="blue";
                }
                tabGeneral.style.display="";
                tabPatch.style.display="none";
                tabBIOS.style.display="none";
                tabRestore.style.display="none";
                tabImages.style.display="none";
                tabAlerts.style.display="none";
                tabPreview.style.display="none";
                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = true;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                frmMain.txtRTMName0.focus();
                window.scrollTo(0,0);
                break;

            case "BIOS":
                lblTitle.innerText = "Select System BIOS to RTM";

                if (frmMain.txtBIOSComments.value == "")
                {
                    frmMain.txtBIOSComments.value = frmMain.txtBIOSCommentsTemplate.value;
                    frmMain.txtBIOSComments.style.fontStyle = "italic";
                    frmMain.txtBIOSComments.style.color="blue";
                }

                tabGeneral.style.display="none";
                tabBIOS.style.display="";
                tabPatch.style.display="none";
                tabRestore.style.display="none";
                tabImages.style.display="none";
                tabAlerts.style.display="none";
                tabPreview.style.display="none";
                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtBIOSComments.focus();
                window.scrollTo(0,0);
                break;

            case "Patches":
                lblTitle.innerText = "Select Patches to RTM";

                if (frmMain.txtPatchComments.value == "")
                {
                    frmMain.txtPatchComments.value = frmMain.txtPatchCommentsTemplate.value;
                    frmMain.txtPatchComments.style.fontStyle = "italic";
                    frmMain.txtPatchComments.style.color="blue";
                }

                tabGeneral.style.display="none";
                tabBIOS.style.display="none";
                tabPatch.style.display="";
                tabRestore.style.display="none";
                tabImages.style.display="none";
                tabAlerts.style.display="none";
                tabPreview.style.display="none";
                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtBIOSComments.focus();
                window.scrollTo(0,0);
                break;

            case "Restore":

                //document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                //document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;
		
                lblTitle.innerText = "Select Restore Media to RTM";

                if (frmMain.txtRestoreComments.value == "")
                {
                    frmMain.txtRestoreComments.value = frmMain.txtRestoreCommentsTemplate.value;
                    frmMain.txtRestoreComments.style.fontStyle = "italic";
                    frmMain.txtRestoreComments.style.color="blue";
                }

                tabGeneral.style.display="none";
                tabBIOS.style.display="none";
                tabPatch.style.display="none";
                tabRestore.style.display="";
                tabImages.style.display="none";
                tabAlerts.style.display="none";
                tabPreview.style.display="none";
                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtRestoreComments.focus();
                window.scrollTo(0,0);
                break;

            case "Images":

                //document.all.BuildLevelIFrame.src = frmMain.txtBuildLevelIFramesrc.value;
                //document.all.DistributionIFrame.src = frmMain.txtDistributionIFramesrc.value;
		
                lblTitle.innerText = "Select Images to RTM";

                if (frmMain.txtImageComments.value == "")
                {
                    frmMain.txtImageComments.value = frmMain.txtImageCommentsTemplate.value;
                    frmMain.txtImageComments.style.fontStyle = "italic";
                    frmMain.txtImageComments.style.color="blue";
                }

                tabGeneral.style.display="none";
                tabBIOS.style.display="none";
                tabRestore.style.display="none";
                tabPatch.style.display="none";
                tabImages.style.display="";
                tabAlerts.style.display="none";
                tabPreview.style.display="none";
                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                frmMain.focus();
                //frmMain.txtImageComments.focus();
                window.scrollTo(0,0);
                break;

            case "Alerts":
                //Temp
                /*
                frmMain.chkAvailability.checked=true;
                frmMain.chkBuildLevel.checked=true;
                frmMain.chkCertification.checked=true;
                frmMain.chkDeveloper.checked=true;
                frmMain.chkDistribution.checked=true;
                frmMain.chkOTSPrimary.checked=true;
                frmMain.chkRoot.checked=true;
                frmMain.chkWorkflow.checked=true;
                */
                //temp
                lblTitle.innerText = "Review Alerts";

                tabGeneral.style.display="none";
                tabBIOS.style.display="none";
                tabRestore.style.display="none";
                tabPatch.style.display="none";
                tabImages.style.display="none";
                tabAlerts.style.display="";
                tabPreview.style.display="none";

                ProcessAlertDisplay();//Herb

                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = false;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
                }
                window.document.focus();
                window.scrollTo(0,0);
                break;

            case "Preview":
                lblTitle.innerText = "Review Selected Information";
                
        
                tabGeneral.style.display="none";
                tabBIOS.style.display="none";
                tabRestore.style.display="none";
                tabImages.style.display="none";
                tabPatch.style.display="none";
                tabAlerts.style.display="none";
                tabPreview.style.display="";

                PopulatePreview();

                if (! FormLoading)
                {
                    window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
                    window.parent.frames["LowerWindow"].cmdNext.disabled = true;
                    window.parent.frames["LowerWindow"].cmdFinish.disabled = false;
                }

                //frmMain.txtPreview.focus();
                window.document.focus();
                window.scrollTo(0,0);
                break;
        }
    }

    function PopulatePreview(){

       
        for(var j =0; j< intProducts; j++){

            var strPreview="";
            var i;
            var strVersions="";
            var ImageValueArray;
            var ControlArray=new Array(); //BIOS
            var RestoreArray=new Array(); 
            var PatchArray=new Array(); 
            var ImageArray=new Array(); 
            var isFusion=false;  
            
            var divAlertPage = document.getElementById("divProductAlerts" + j.toString());
            var strDivAlertPageId = "divProductAlerts" + j.toString();
            var divPreview = document.getElementById("divPreview" + j.toString());
            var strProdName = divPreview.getAttribute("prodName");
            var strProdId = divPreview.getAttribute("prodId");

            strPreview="";

            if (frmMain.txtRTMComments.value == frmMain.txtRTMCommentsTemplate.value)
                frmMain.txtRTMComments.value = "";

            if (typeof(frmMain.txtBIOSComments) != "undefined")
                if (frmMain.txtBIOSComments.value == frmMain.txtBIOSCommentsTemplate.value)
                    frmMain.txtBIOSComments.value = "";

            if (typeof(frmMain.txtRestoreComments) != "undefined")
                if (frmMain.txtRestoreComments.value == frmMain.txtRestoreCommentsTemplate.value)
                    frmMain.txtRestoreComments.value = "";

            if (typeof(frmMain.txtImageComments) != "undefined")
                if (frmMain.txtImageComments.value == frmMain.txtImageCommentsTemplate.value)
                    frmMain.txtImageComments.value = "";


            //General
            strPreview = "<font size=2 face=verdana><b>General RTM Information</b></font><table class=EmbeddedTable bgcolor=white width=100% border=1 cellpadding=2 cellspacing=0>";
            strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>RTM Title:</b></td><td width='100%'>" + document.getElementById("txtRTMName" + j.toString()).value + "&nbsp;</td>";
            strPreview = strPreview + "<td nowrap valign=top bgcolor=gainsboro><b>RTM Date:</b></td><td width='120'>" + frmMain.txtRTMDate.value + "</td></tr>";
        
            if (frmMain.txtAttachmentPath1){
            if ((frmMain.txtAttachmentPath1.value !="") && (!frmMain.chkImages.checked)){
                RemoveUpload(1);
            }}

            if (!UploadPath1.innerText){
                //if (document.getElementById("chkSCMX").checked){
                //    strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>SCMx File:</b></td><td colspan=3 width='100%'>" + "None" + "&nbsp;</td></tr>";
                //}
            }else if (UploadPath1.innerText != ""){
                strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>SCMx File:</b></td><td colspan=3 width='100%'>" + "<a target='_blank' href='" + frmMain.txtUploadPath1.value + "'>" + UploadPath1.innerText +"</a>" + "&nbsp;</td></tr>";
            }

    
            if (frmMain.txtRTMComments.value)
                strPreview = strPreview + "<tr><td nowrap valign=top bgcolor=gainsboro><b>RTM Comments:</b></td><td colspan=3>" + frmMain.txtRTMComments.value.replace(/\r\n/g,"<BR>") + "</td></tr>";
            strPreview = strPreview + "</table>";
        

        


            //BIOS
            var strBIOSPreview = "";
            if (frmMain.chkBIOS.checked)
            {
                strBIOSPreview = "";
    
                var tblBIOS = document.getElementById("tableBIOS" + j.toString());
                
                if (!$(tblBIOS)){
                    strBIOSPreview = "";
                }else if($(tblBIOS).find('#chkBIOSList')){
                    strPreview = strPreview + "<BR><font size=2 face=verdana><b>System BIOS to RTM</b></font><BR>" ;
                    if (frmMain.txtBIOSComments.value)
                    {
                        strPreview = strPreview + "<font size=1 face=verdana color=black><BR><i>" + frmMain.txtBIOSComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
                    }                

                    if (typeof($(tblBIOS).find('#chkBIOSList').length) == "undefined"){
                        ControlArray[0] = $(tblBIOS).find('#chkBIOSList');
                    }else{
                        ControlArray = $(tblBIOS).find('#chkBIOSList');
                    }

                    for (i=0;i<ControlArray.length;i++){
                        if (ControlArray[i].checked){
                            strBIOSPreview = strBIOSPreview + "<tr><td>" + $(ControlArray[i]).attr('previewid')  + "</td><td>" + $(ControlArray[i]).attr('previewname') + "</td><td>" + $(ControlArray[i]).attr('previewversion') + "</td>"
                            if (frmMain.optCutIn.checked){
                                strBIOSPreview = strBIOSPreview + "<td>Immediate (Rework All Units)</td></tr>";
                            }else if (frmMain.optWebOnly.checked){
                                strBIOSPreview = strBIOSPreview + "<td>Web Release Only</td></tr>";
                            }else{
                                strBIOSPreview = strBIOSPreview + "<td>Phase-in</td></tr>";
                            }
                        }
                    }


                    strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Affectivity</b></td></tr>"
                    strPreview = strPreview + strBIOSPreview + "</table>";
                }
            }
        

            //Patch
            var strPatchPreview = "";
            if (frmMain.chkPatch.checked)
            {
                var tblPatch = document.getElementById("tablePatch" + j.toString());
                if (!$(tblPatch)){
                    // no patch on this product
                    strPatchPreview = "";
                }else{
                    strPatchPreview = strPatchPreview + "<BR><font size=2 face=verdana><b>Patches to RTM</b></font><BR>" ;
                    if (frmMain.txtPatchComments.value){
                        strPatchPreview = strPatchPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtPatchComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
                    }

                    if (typeof($(tblPatch).find('#chkPatchList')) == "undefined"){
                        PatchArray[0] = $(tblPatch).find('#chkPatchList');
                    }else{
                        PatchArray = $(tblPatch).find('#chkPatchList');
                    }
                    strPatchPreview = strPatchPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>ID</b></td><td><b>Name</b></td><td><b>Version</b></td><td><b>Patch&nbsp;Contents</b></td><td><b>Images</b></td></tr>";
                    
                    var intChecked =0;
                    for (i=0;i<PatchArray.length;i++){
                        if (PatchArray[i].checked)
                        {
                            intChecked = intChecked + 1;
                            strPatchPreview = strPatchPreview + "<tr><td valign=top>" + $(PatchArray[i]).attr('previewid')  + "</td><td valign=top>" + $(PatchArray[i]).attr('previewname') + "</td><td valign=top>" + $(PatchArray[i]).attr('previewversion')  + "</td><td valign=top>" + $(PatchArray[i]).attr('previewcontents')  + "&nbsp;</td><td valign=top><a target=_blank href='../Image/PatchImages.asp?ProdID=" + $(PatchArray[i]).attr('productid') + "&DelID=" + $(PatchArray[i]).attr('previewid')  + "'>View</a>&nbsp;</td></tr>"
                        }
                    }
                    strPatchPreview = strPatchPreview + "</table>";

                    if(intChecked == 0){
                        strPatchPreview ="";
                    }

                }
                strPreview = strPreview + strPatchPreview;

           
            }

            //Image
            var strImgPreview = "";
            if (frmMain.chkImages.checked)
            {
                strImgPreview = "";
                if (typeof(frmMain.txtImagePreview.length) == "undefined")
                    ImageArray[0] = frmMain.txtImagePreview;
                else
                    ImageArray = frmMain.txtImagePreview;

                for (i=0;i<ImageArray.length;i++)
                {
                    ImageValueArray = ImageArray[i].value.split("|")
                    if (ImageArray[i].getAttribute("prodId").toString() == strProdId){

                        if (ImageValueArray[0] == "FUSION")
                        {
                            isFusion = true;
                            if (document.all("chkImage" + ImageValueArray[1]).checked)
                            {
                                strImgPreview = strImgPreview + "<tr><td>" + ImageValueArray[2]  + "&nbsp;</td><td>" + ImageValueArray[7] + "&nbsp;-&nbsp;" + ImageValueArray[6] + "&nbsp;</td><td>" + ImageValueArray[3]   + "&nbsp;</td><td>" + ImageValueArray[4]   + "&nbsp;</td><td>" + ImageValueArray[5] + "&nbsp;</td></tr>"
                            }
                        }
                        else
                        {
                            if (document.all("chkImage" + ImageValueArray[0]).checked)
                            {
                                strImgPreview = strImgPreview + "<tr><td>" + ImageValueArray[1]  + "&nbsp;</td><td>" + ImageValueArray[8] + "&nbsp;-&nbsp;" + ImageValueArray[7] + "&nbsp;</td><td>" + ImageValueArray[2]   + "&nbsp;</td><td>" + ImageValueArray[3]   + "&nbsp;</td><td>" + ImageValueArray[4] + "&nbsp;</td><td>" + ImageValueArray[5]  + "&nbsp;</td></tr>"
                            }
                        }
                    }
                }
                strPreview = strPreview + "<BR><font size=2 face=verdana><b>Images to RTM</b></font><BR>" ;
                if (frmMain.txtImageComments.value)
                {
                    strPreview = strPreview + "<font face=verdana size=1 color=black><BR><i>" + frmMain.txtImageComments.value.replace(/\r\n/g, '<br>') + "</i><BR><br></font>"
                }
                if (isFusion)
                    strPreview = strPreview + "<div id=divPvImgLstTbl><table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>Product&nbsp;Drop</b></td><td><b>Region</b></td><td><b>Brands</b></td><td><b>OS</b></td><td><b>Comments</b></td></tr>"
                else
                    strPreview = strPreview + "<div id=divPvImgLstTbl><table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>SKU</b></td><td><b>Region</b></td><td><b>Model</b></td><td><b>OS</b></td><td><b>Apps Bundle</b></td><td><b>BTO/CTO</b></td></tr>"
                strPreview = strPreview + strImgPreview + "</table></div>";
            }
        
            //Alert
            strPreview = strPreview + "<div id=AlertPreviewSection><BR><font size=2 face=verdana><b>Alerts Reviewed</b></font><BR>" ;
            strPreview = strPreview + "<table class=EmbeddedTable bgcolor=white width='100%' cellspacing=0 cellpadding=2 border=1><tr bgcolor=gainsboro><td><b>Alert</b></td><td><b>Count</b></td><td width='100%'><b>Comments</b></td></tr>";
            try{
                strPreview = strPreview + "<tr><td nowrap>Build Level</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"BuildLevelIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"BuildLevelIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtBuildLevelComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Build Level</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Distribution</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"DistributionIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"DistributionIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtDistributionComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Distribution</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Certification</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"CertificationIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"CertificationIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtCertificationComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Certification</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Workflow</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"WorkflowIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"WorkflowIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtWorkflowComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Workflow</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Availability</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"AvailabilityIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"AvailabilityIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtAvailabilityComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Availability</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Developer</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"DeveloperIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"DeveloperIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtDeveloperComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Developer</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>Root</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"RootIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"RootIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtRootComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>Root</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
            try{
                strPreview = strPreview + "<tr><td nowrap>OTS Primary</td><td nowrap align=center><a target=_blank href='../Product/MilestoneSignoffAlertPreview.asp?ID=" + GetElementInsideContainer(strDivAlertPageId ,"OTSPrimaryIFrame").contentWindow.RecordID.innerText + "'>" + GetAlertCount(GetElementInsideContainer(strDivAlertPageId ,"OTSPrimaryIFrame").contentWindow.document.body.innerHTML) + "</a></td><td>" + GetElementInsideContainer(strDivAlertPageId ,"txtOTSPrimaryComments").value + "&nbsp;</td></tr>";
            }catch(err){
                //strPreview = strPreview + "<tr><td nowrap>OTS Primary</td><td>Error</td><td>" + err.message.toString() +   "</td></tr>";
            }
   
            // if Patch and Image are no data, no RTM and display Msg
            if ((strPatchPreview + strImgPreview + strBIOSPreview) ==""){
                document.getElementById("lblPreviewMsg" + j.toString()).innerHTML = " - [ This product will not RTM. Because no BIOS, image or patch was selected. ]";
            }



            strPreview = strPreview + "</table></div>";

            
            $(divPreview).find("#PreviewPage").html(strPreview);

        }

    }


    function GetAlertCount(strString){
        if (strString.indexOf("<TD>None Found</TD></TR>") >-1 && strString.toLowerCase().split(/<tr>/g).length==1)
            return 0;
        else
            return strString.toLowerCase().split(/<tr>/g).length-1;
    }

    function chkBIOS_onclick(){
        if (frmMain.chkBIOS.checked)
            BIOSAffectivityRow.style.display = "";
        else
            BIOSAffectivityRow.style.display = "none";
    }


//////////////////////////////////////////////////////////
    function GetElementInsideContainer(containerID, childID) {
        var elm = {};
        var elms = document.getElementById(containerID).getElementsByTagName("*");
        for (var i = 0; i < elms.length; i++) {
            if (elms[i].id === childID) {
                elm = elms[i];
                break;
            }
        }
        return elm;
    }

    //////////////////////////
    //include jQuery
    // post RTM data by each product
    function submitMutipleRTM(){
        for(var j =0; j< intProducts; j++){
            submitRTM(rtmPackage(j));
        }

    }


    function submitRTM(rtmPack){

        if ((rtmPack.chkPatch ==0) && (rtmPack.chkImages ==0) && (rtmPack.chkBIOS ==0)){
            alert( "[ " + rtmPack.txtProductName + " ] \n \n No Image or patch was selected on this product. \n This product is skipped." );
        }else{
            $.ajax({
                type: 'POST',
                url: 'MultipleRTMSave.asp',
                data: rtmPack,
                success: function() {
                            alert( "Success: [ " + rtmPack.txtProductName + " ]");
                         },
                async:false
            })        
            .fail(function() {
                alert( "Error: [ " + rtmPack.txtProductName + " ]" );
            })


//          $.post('MultipleRTMSave.asp', rtmPack, function() {
//              alert( "Success: [ " + rtmPack.txtProductName + " ]");
//          })
//          .fail(function() {
//              alert( "Error: [ " + rtmPack.txtProductName + " ]" );
//          })
//          ;
        };
    }


    //Object
    function rtmPackage( productIndex ){
        var strDivAlertPageId = "divProductAlerts" + productIndex.toString();
        var strDivPreviewId = "divPreview" + productIndex.toString();

        var intChkPatch =0;
        var strChkPatchList ="";
        if(frmMain.chkPatch.checked){
            if (!(document.getElementById("tablePatch" + productIndex.toString()))){
                // no patch on this product
                intChkPatch =0;
            }else{
                intChkPatch =1;
                strChkPatchList = getStrProductPatchList(productIndex);
            }
        }
        if(strChkPatchList == ""){
            intChkPatch =0;
        }


        var intChkImges =0;
        var strChkImgs ="";
        if(frmMain.chkImages.checked){
            intChkImges =1;
            strChkImgs = getStrProductImageList(productIndex);
        }
        if(strChkImgs == ""){
            intChkImges =0;
        }


        var intChkBIOS =0;
        var strChkBIOSList ="";
        if(frmMain.chkBIOS.checked){
            intChkBIOS =1;
            strChkBIOSList = getStrProductBIOSList(productIndex);
        }
        if(strChkBIOSList == ""){
            intChkBIOS =0;
        }
        var intOptPhaseIn =0;
        try{
 
            if (frmMain.optCutIn.checked){
                intOptPhaseIn =0;
            }else if (frmMain.optWebOnly.checked){
                intOptPhaseIn =2;
            }else{
                intOptPhaseIn =1;
            }

        }catch(e){
            intChkBIOS =0;
        };


        var strEmailText ="";
        var strEmailReplaceMsg = "<p>About the RTM Image list, Please refer to the Excalibur website.</P>";
        strEmailText = GetElementInsideContainer(strDivPreviewId, "PreviewPage").innerHTML;
        if(strEmailText.length > 90000){
            strEmailText = strEmailText.replace(GetElementInsideContainer(strDivPreviewId, "divPvImgLstTbl").innerHTML,strEmailReplaceMsg);
        }
        var strTxtProductID = document.getElementById(strDivPreviewId).getAttribute("prodID");
        var strTxtProductName = document.getElementById(strDivPreviewId).getAttribute("prodName");
        var strTxtRTMComments = frmMain.txtRTMComments.value;
        var strTxtRTMDate = frmMain.txtRTMDate.value;
        var strTxtRTMName = document.getElementById("txtRTMName" + productIndex.toString()).value;
        var strTxtCurrentUserEmail = frmMain.txtCurrentUserEmail.value;
        var strTxtNotify = GetElementInsideContainer(strDivPreviewId, "txtNotify").value;
        var strTxtImageComments = "";
        if(frmMain.txtImageComments){
            strTxtImageComments = frmMain.txtImageComments.value;
        };
        var strTxtUploadPath1 = ""; 
        if(intChkImges != 0){
            if(frmMain.txtUploadPath1){
                strTxtUploadPath1 = frmMain.txtUploadPath1.value;  
            }
        };
        var strTxtAttachmentPath1 = ""; 
        if(intChkImges != 0){
            if(frmMain.txtAttachmentPath1){
                strTxtAttachmentPath1 = frmMain.txtAttachmentPath1.value ;    
            }
        };
        var strTxtPatchComments = "";
        if(frmMain.txtPatchComments){
            strTxtPatchComments = frmMain.txtPatchComments.value;
        };
        var strTxtBIOSComments = "";
        if(frmMain.txtBIOSComments){
            strTxtBIOSComments = frmMain.txtBIOSComments.value;
        }
        var strTxtRestoreComments = "";
        if(frmMain.txtRestoreComments){
            strTxtRestoreComments = frmMain.txtRestoreComments.value;
        }



        var strTxtBuildLevelComments = GetElementInsideContainer(strDivAlertPageId ,"txtBuildLevelComments").value;
        var strTxtDistributionComments = GetElementInsideContainer(strDivAlertPageId ,"txtDistributionComments").value;
        var strTxtCertificationComments = GetElementInsideContainer(strDivAlertPageId ,"txtCertificationComments").value;
        var strTxtWorkflowComments = GetElementInsideContainer(strDivAlertPageId ,"txtWorkflowComments").value;
        var strTxtAvailabilityComments = GetElementInsideContainer(strDivAlertPageId ,"txtAvailabilityComments").value;
        var strTxtDeveloperComments = GetElementInsideContainer(strDivAlertPageId ,"txtDeveloperComments").value;
        var strTxtRootComments =GetElementInsideContainer(strDivAlertPageId ,"txtRootComments").value;
        var strTxtOTSPrimaryComments =GetElementInsideContainer(strDivAlertPageId ,"txtOTSPrimaryComments").value;
       
        var objRTM = {
            txtProductID : strTxtProductID,
            txtProductName : strTxtProductName,
            txtRTMComments :strTxtRTMComments,
            txtRTMDate : strTxtRTMDate,
            txtRTMName : strTxtRTMName,

            txtCurrentUserEmail : strTxtCurrentUserEmail,
            txtNotify : strTxtNotify,
            txtEmailPreview : strEmailText,

            chkImages : intChkImges,
            chkImage : strChkImgs,
            txtImageComments : strTxtImageComments,
            txtUploadPath1 : strTxtUploadPath1,  
            txtAttachmentPath1 : strTxtAttachmentPath1,

            chkPatch : intChkPatch,
            chkPatchList : strChkPatchList,
            txtPatchComments : strTxtPatchComments,

            chkBIOS : intChkBIOS,
            chkBIOSList : strChkBIOSList,
            optPhaseIn : intOptPhaseIn,
            txtBIOSComments : strTxtBIOSComments,


            //chkRestore : ,
            //chkRestoreList : ,
            txtRestoreComments : strTxtRestoreComments,

            //chkAvailability : ,
            //chkBuildLevel : ,
            //chkCertification : ,
            //chkDeveloper : ,
            //chkDistribution : ,
            //chkRoot : ,
            //chkWorkflow : ,
            //chkOTSPrimary : ,      
                
            txtBuildLevelComments : strTxtBuildLevelComments,
            txtDistributionComments : strTxtDistributionComments, 
            txtCertificationComments : strTxtCertificationComments ,
            txtWorkflowComments : strTxtWorkflowComments,
            txtAvailabilityComments : strTxtAvailabilityComments,  
            txtDeveloperComments : strTxtDeveloperComments,
            txtRootComments : strTxtRootComments,
            txtOTSPrimaryComments : strTxtOTSPrimaryComments 
        }
        return objRTM;
    }

    //json
    function jsonRtmPackage( productIndex ){
        var strRTM = "{}";
  
        return strRTM;
    }

    function getStrProductImageList(ProductIdx){
        var strResult = "";
        var arrImg = new Array();
        var strProdId = document.getElementById("divPreview" + ProductIdx.toString()).getAttribute("prodId");

        var allChkImage0 = document.getElementsByName("chkImage");
        var allChkImage;

        if (typeof(allChkImage0.length) == "undefined"){
            allChkImage = new Array();
            allChkImage[0] = allChkImage0;
        }else{
            allChkImage = allChkImage0;
        }

        for (var k =0; k<allChkImage.length; k++){
            if((allChkImage[k].getAttribute("prodId") == strProdId) && (allChkImage[k].checked)){
                arrImg.push(allChkImage[k].value.toString());
            }
        }

        if(arrImg.length >0){
            strResult = arrImg.toString();
        }else{
            strResult ="";
        }
        
        return strResult;
    }

    function getStrProductPatchList(ProductIdx){
        var strResult = "";
        var arrPatch = new Array();
        var tblPatch = document.getElementById("tablePatch" + ProductIdx.toString());
        var allChkPatchList = new Array();

        if (!$(tblPatch)){
            strResult = "";
        }else{
            if (typeof($(tblPatch).find("#chkPatchList").length) == "undefined"){
                allChkPatchList[0] = $(tblPatch).find("#chkPatchList");
            }else{
                allChkPatchList = $(tblPatch).find("#chkPatchList");
            }
        }
        for (k=0; k<allChkPatchList.length; k++ ){
            if (allChkPatchList[k].checked)
            {
                arrPatch.push(allChkPatchList[k].value.toString());            
            }
        }


        if(arrPatch.length >0){
            strResult = arrPatch.toString();
        }else{
            strResult ="";
        }
        
        return strResult;
    }

    function getStrProductBIOSList(ProductIdx){
        var strResult = "";
      
        var arrBIOS = new Array();
        var tblBIOS = document.getElementById("tableBIOS" + ProductIdx.toString());
        var allChkBIOSList = new Array();

        if (typeof($(tblBIOS).find("#chkBIOSList")) == "undefined"){
            strResult = "";
        }else{
            if (typeof($(tblBIOS).find("#chkBIOSList").length) == "undefined"){
                allChkBIOSList[0] = $(tblBIOS).find("#chkBIOSList");
            }else{
                allChkBIOSList = $(tblBIOS).find("#chkBIOSList");
            }
        }
        for (k=0; k<allChkBIOSList.length; k++ ){
            if (allChkBIOSList[k].checked)
            {
                arrBIOS.push(allChkBIOSList[k].value.toString());            
            }
        }

        if(arrBIOS.length >0){
            strResult = arrBIOS.toString();
        }else{
            strResult ="";
        }

        return strResult;
    }

    //////////////////////////

    function chkAlert_onClick(intAlertPageNumber,chkAlert){
       
        var divAlertPage;
        var strDivId;
        var strIdx;
        var strAlertName;

        strAlertName = chkAlert.getAttribute("alertname").toString();
        strDivId = "divProductAlerts" + intAlertPageNumber.toString();

        if (chkAlert.checked)
        {
            GetElementInsideContainer(strDivId,strAlertName + "AlertDetails").style.display="none";
            GetElementInsideContainer(strDivId,"txt" + strAlertName + "Comments").focus();  
        }
        else{
            GetElementInsideContainer(strDivId,strAlertName + "AlertDetails").style.display="";
        }
            
    }

//    function chkBuildLevel_onclick(){
//        return;
//        if (frmMain.chkBuildLevel.checked)
//        {
//            BuildLevelAlertDetails.style.display="none";
//            frmMain.txtBuildLevelComments.focus();
//           
//        }
//        else
//            BuildLevelAlertDetails.style.display="";
//    }
//
//    function chkDistribution_onclick(){
//        return;
//        if (frmMain.chkDistribution.checked)
//        {
//            DistributionAlertDetails.style.display="none";
//            frmMain.txtDistributionComments.focus();
//        }
//        else
//            DistributionAlertDetails.style.display="";
//    }
//
//    function chkCertification_onclick(){
//        return;
//        if (frmMain.chkCertification.checked)
//        {
//            CertificationAlertDetails.style.display="none";
//            frmMain.txtCertificationComments.focus();
//        }
//        else
//            CertificationAlertDetails.style.display="";
//    }
//    
//    function chkWorkflow_onclick(){
//        return;
//        if (frmMain.chkWorkflow.checked)
//        {
//            WorkflowAlertDetails.style.display="none";
//            frmMain.txtWorkflowComments.focus();
//        }
//        else
//            WorkflowAlertDetails.style.display="";
//    }
//
//    function chkAvailability_onclick(){
//        return;
//        if (frmMain.chkAvailability.checked)
//        {
//            AvailabilityAlertDetails.style.display="none";
//            frmMain.txtAvailabilityComments.focus();
//        }
//        else
//            AvailabilityAlertDetails.style.display="";
//    }
//    
//    function chkDeveloper_onclick(){
//        return;
//        if (frmMain.chkDeveloper.checked)
//        {
//            DeveloperAlertDetails.style.display="none";
//            frmMain.txtDeveloperComments.focus();
//        }
//        else
//            DeveloperAlertDetails.style.display="";
//    }
//
//    function chkRoot_onclick(){
//        return;
//        if (frmMain.chkRoot.checked)
//        {
//            RootAlertDetails.style.display="none";
//            frmMain.txtRootComments.focus();
//        }
//        else
//            RootAlertDetails.style.display="";
//    }
//
//    function chkOTSPrimary_onclick(){
//        return;
//        if (frmMain.chkOTSPrimary.checked)
//        {
//            OTSPrimaryAlertDetails.style.display="none";
//            frmMain.txtOTSPrimaryComments.focus();
//        }
//        else
//            OTSPrimaryAlertDetails.style.display="";
//    }
//
    /////////////////////////////////////////////////////////

    function cmdDate_onclick(FieldID) {
        var strID;
	
		
        strID = window.showModalDialog("../mobilese/today/caldraw1.asp",frmMain.txtRTMDate.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
        if (typeof(strID) != "undefined")
            frmMain.txtRTMDate.value = strID;
    }


    function BIOSRow_onclick() {
        var RowElement;
	
        RowElement = window.event.srcElement;
        while (RowElement.className != "Row")
        {
            RowElement = RowElement.parentElement;
        }
        if(RowElement.style.backgroundColor=="cornflowerblue")//lightgoldenrodyellow
            RowElement.style.backgroundColor="";
        else
            RowElement.style.backgroundColor="cornflowerblue";//lightgoldenrodyellow

        if 	(typeof(SelectedBIOSRow) != "undefined")
        {
            if (SelectedBIOSRow!=RowElement)
                SelectedBIOSRow.style.backgroundColor="";
        }
        SelectedBIOSRow=RowElement;

    }

    function BaseRow_onmouseover(strID){
        if (window.event.srcElement.id == strID)
        {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("BaseRow" + strID).bgColor="lightsteelblue";	
    }

    function DropRow_onmouseover(strID){
        if (window.event.srcElement.id == strID)
        {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("DropRow" + strID).bgColor="lightsteelblue";	
    }

    function BaseRow_onmouseout(strID){
        document.all("BaseRow" + strID).bgColor="cornsilk";	
    }

    function DropRow_onmouseout(strID){
        document.all("DropRow" + strID).bgColor="cornsilk";	
    }

    function BaseRow_onclick(strID){

        if (window.event.srcElement.id == strID)
            return;


        if (document.all("ImageRow" + strID).style.display == "" )
            document.all("ImageRow" + strID).style.display="none";	
        else
            document.all("ImageRow" + strID).style.display="";	
    }

    function DropRow_onclick(strID){

        if (window.event.srcElement.id == strID)
            return;


        if (document.all("DropContents" + strID).style.display == "" )
            document.all("DropContents" + strID).style.display="none";	
        else
            document.all("DropContents" + strID).style.display="";	
    }


    function chkAll_onclick(){
        var i;
	
        if(typeof(frmMain.chkImage.length)=="undefined")
        {
            if (frmMain.chkImage.indeterminate) // && frmMain.chkAll.checked
            {
                frmMain.chkImage.indeterminate=0;
                //document.all("Lang" + frmMain.chkImage.value).innerText = document.all("Lang" + frmMain.chkImage.value).className;
                document.all("Row" + frmMain.chkImage.value).bgColor = "ivory";
            }
            frmMain.chkImage.checked = frmMain.chkAll.checked;
        }
        else
        {
            for (i=0;i<frmMain.chkImage.length;i++)
            {
                if (frmMain.chkImage(i).indeterminate) //&& frmmain.chkAll.checked
                {
                    frmMain.chkImage(i).indeterminate=0;
                    //document.all("Lang" + frmMain.chkImage(i).value).innerText = document.all("Lang" + frmMain.chkImage(i).value).className;
                    document.all("Row" + frmMain.chkImage(i).value).bgColor = "ivory";
                }
                frmMain.chkImage(i).checked = frmMain.chkAll.checked;
                if (document.all("Base" + frmMain.chkImage(i).className).indeterminate) //&& frmMain.chkAll.checked
                    document.all("Base" + frmMain.chkImage(i).className).indeterminate=0;			
                document.all("Base" + frmMain.chkImage(i).className).checked = frmMain.chkAll.checked;
            }
        }

    }

    function chkBase_onclick(){
        var i;
	
        if(typeof(frmMain.chkImage.length)=="undefined")
        {
            if (frmMain.chkImage.indeterminate && window.event.srcElement.checked)
            {
                frmMain.chkImage.indeterminate=0;
                //document.all("Lang" + frmMain.chkImage.value).innerText = document.all("Lang" + frmMain.chkImage.value).className;
                document.all("Row" + frmMain.chkImage.value).bgColor = "ivory";
            }
            frmMain.chkImage.checked = window.event.srcElement.checked;
        }
        else
        {
            for (i=0;i<frmMain.chkImage.length;i++)
            {
                if (frmMain.chkImage(i).className == window.event.srcElement.id)
                {
                    if (frmMain.chkImage(i).indeterminate && window.event.srcElement.checked)
                    {
                        frmMain.chkImage(i).indeterminate=0;
                        //document.all("Lang" + frmMain.chkImage(i).value).innerText = document.all("Lang" + frmMain.chkImage(i).value).className;
                        document.all("Row" + frmMain.chkImage(i).value).bgColor = "ivory";
                    }
                    frmMain.chkImage(i).checked = window.event.srcElement.checked;
                }
            }
        }	

    } 

    function chkDrop_onclick(){
        var i;
	
        if(typeof(frmMain.chkImage.length)=="undefined")
        {
            frmMain.chkImage.checked = window.event.srcElement.checked;
        }
        else
        {
            for (i=0;i<frmMain.chkImage.length;i++)
            {
                if (frmMain.chkImage(i).className == window.event.srcElement.id)
                {
                    frmMain.chkImage(i).checked = window.event.srcElement.checked;
                }
            }
        }	

    } 

    function UpdateBase(chkClicked){
        var i;
        var blnAllSame=true;
	
        for (i=0;i<frmMain.chkImage.length;i++)
        {
		
            if (frmMain.chkImage(i).className != "")
                if (frmMain.chkImage(i).className == chkClicked.className)
                {
                    if ((frmMain.chkImage(i).checked != chkClicked.checked) || frmMain.chkImage(i).indeterminate)
                    {
                        blnAllSame = false;	
                    }
                }
        }
	
        if (blnAllSame)
        {
            document.all("Base" + chkClicked.className).indeterminate=0;
            document.all("Base" + chkClicked.className).checked = chkClicked.checked;
        }
        else
            document.all("Base" + chkClicked.className).indeterminate=-1;

        if (chkClicked.checked)
        {
            //document.all("Lang" + chkClicked.value).innerText = document.all("Lang" + chkClicked.value).className;
            if (document.all("Row" + chkClicked.value)!=null)
                document.all("Row" + chkClicked.value).bgColor = "ivory";
        }

    }



    function chkImage_onclick(){
        UpdateBase(window.event.srcElement);
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

.EmbeddedTable TBODY TD{
	FONT-FAMILY: Verdana;
}
.EmbeddedTable TBODY TD{
	Font-Size: xx-small;
}

input
{
    FONT-SIZE: 10pt;	
    FONT-FAMILY: Verdana;	
}
textarea
{
    FONT-SIZE: 10pt;	
    FONT-FAMILY: Verdana;	
}
.ImageTable TBODY TD{
	BORDER-TOP: gray thin solid;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}
.ImageTable TH{
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}

.imagerows TBODY TD{
	BORDER-TOP: none;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}

.imagerows THEAD TD{
	BORDER-TOP: none;
	FONT-SIZE: xx-small;
	FONT-FAMILY: verdana;
}
</STYLE>
<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<%
	dim cn
	dim rs
	dim blnFound 
	dim i
    dim j
	dim cm
	dim p
	dim strCategories
	dim CurrentUser
	dim CurrentUserID
	dim CurrentWorkgroupID
	dim strSEPMID
	dim strPMID
	dim strTestLeadID
	dim blnPOR
	dim blnEditOK
	dim strShowEditBoxes
	dim strVersion
	dim strProductName
    dim blnFusion
	dim strEmployees
	dim strDevCenter
	dim strLastRoot
	dim BIOSCount
	dim RestoreCount
    dim PatchCount
	dim CurrentUserEmail
	dim strPMRDate
	dim RTMCommentsTemplate
	dim BIOSCommentsTemplate
	dim RestoreCommentsTemplate
	dim ImageCommentsTemplate
    dim strCDPartNumber
    dim strCDPartNumber2
	dim strProdIds
    dim arrProdIDs '//Herb
    dim lngProdId '//Herb
    
	
'	RTMCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
'							"1. Reason for RTM" & vbcrlf & _
'							"2. Special circumstances, patches, schedule updates, future expectations." & vbcrlf & _
'							"3. Any factory holds affected/lifted and actions surrounding the issue." & vbcrlf & _
'							"4. Any special rules set by DCR/AVs/EAs/Factory Memos/Etc." & vbcrlf & _
'							"5. Any platform specific additional comments."

    RTMCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                          "1. Reason for RTM" & vbcrlf & _
                          "2. Special circumstances, patches, schedule updates, future expectations." & vbcrlf & _
                          "3. Any factory holds affected/lifted and actions surrounding the issue." & vbcrlf & _
                          "4. Any special rules set by DCR/AVs/EAs/Factory Memos/Etc." & vbcrlf & _
                          "5. Any platform specific additional comments."

	
	BIOSCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
							"1. Any special instructions regarding BIOS cut-ins, staggered schedules for release, parallel releases, risk releases." & vbcrlf & _
							"2. Any information regarding updates to VBIOS, ME, AMT, MRC, etc that are noteworthy." & vbcrlf & _
							"3. Any platform BIOS specific additional comments." 

'	RestoreCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
'							"1. Any part numbers for older media that may not be present in wizard." & vbcrlf & _
'							"2. Verification of media transfer to replication houses and date of arrival." & vbcrlf & _
'							"3. Any other restore media specific additional comments." 
	RestoreCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                              "1. Any part numbers for older media that may not be present in wizard." & vbcrlf & _
                              "2. Verification of media transfer to replication houses and date of arrival." & vbcrlf & _
                              "3. Name of deliverable (i.e. ODIE_1.0_Win7_DRDVD)" & vbcrlf & _
                              "4. Part Number for media to be RTM’d." & vbcrlf & _
                              "5. Version of restore media (i.e 1.28 A,1)" & vbcrlf & _
                              "6. PMR ID for restore media" & vbcrlf & _
                              "7. Any other restore media specific additional comments."


'	ImageCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
'							"1. Final Rev numbers of the images being released." & vbcrlf & _
'							"2. Verification of image transfer to the production servers or schedule for final arrival." & vbcrlf & _
'							"3. Name of the production servers on which the images are housed." & vbcrlf & _
'							"4. Any special requirements and/or dependencies for the images being release (BIOS/FW/Patches/MSCU/etc)." & vbcrlf & _
'							"5. System WHQL completion dates per required OS." & vbcrlf & _
'							"6. MDA log completion dates per required OS." & vbcrlf & _
'							"7. 2PP requirements for all supported images." & vbcrlf & _
'							"8. Marketing Name." & vbcrlf & _
'							"9. PCR File Tested (Not for production)." & vbcrlf & _
'							"10. Any other image specific additional comments." 

	ImageCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                            "1. Final Rev numbers of the images being released." & vbcrlf & _
                            "2. Verification of image transfer to the production servers or schedule for final arrival." & vbcrlf & _
                            "3. Name of the production servers on which the images are housed." & vbcrlf & _
                            "4. Any special requirements and/or dependencies for the images being release (BIOS/FW/Patches/MSCU/etc)." & vbcrlf & _
                            "5. System WHQL completion dates per required OS." & vbcrlf & _
                            "6. MDA log completion dates per required OS." & vbcrlf & _
                            "7. 2PP requirements for all supported images." & vbcrlf & _
                            "8. Marketing Name." & vbcrlf & _
                            "9. PCR File Tested (Not for production)." & vbcrlf & _
                            "10.Be sure to attach the .SCMX file" & vbcrlf & _
                            "11. Provide the SCMX file name." & vbcrlf & _
                            "12. Provide the PCR file name" & vbcrlf & _
                            "13. Provide the released ML numbers after updating the final version online." & vbcrlf & _
                            "14. Any other image specific additional comments."


	PatchCommentsTemplate = "Should include the following information if present:" & vbcrlf & _
                            "1.	Patch Name (i.e. CSI Patch – 3PP Addon)" & vbcrlf & _
                            "2.	CSI Patch Part Number" & vbcrlf & _
                            "3.	CSI Patch version (i.e. 1.00 A,5)" & vbcrlf & _
                            "4.	CSI Patch PRISM Revision number"
 

	BIOSCount = 0
	RestoreCount = 0
    PatchCount=0
	
	strProductName = ""
    blnFusion = false
	strLastRoot = ""
	
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
		CurrentWorkgroupID = rs("WorkgroupID") & ""
		CurrentUserEmail = rs("Email") & ""
	end if
	rs.Close
	
	if request("ID") = "" then
		blnprodFound = false
	else
		rs.Open "spGetProductVersion " & clng(request("ID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strSEPMID = rs("SEPMID") & ""
			strTestLeadID = rs("SeTestLead") & ""
			strPMID = rs("PMID") & ""
            blnFusion = rs("Fusion") & ""
			strPMEmail = rs("SEPMEmail") & ""
			strPMName = rs("SEPMName") & ""
			strProductName = rs("Name") & " " & rs("Version")
			strDevCenter = rs("DevCenter") & ""
			if trim(rs("RTMNotifications") & "") = "" then
			    strDistribution = CurrentUserEmail & ";" & trim(rs("Distribution") & "")
			elseif instr(trim(rs("RTMNotifications") & ""),CurrentUserEmail) < 1 then
			    strDistribution = CurrentUserEmail & ";" & trim(rs("RTMNotifications") & "")
            else
			    strDistribution = trim(rs("RTMNotifications") & "")
			end if
			blnProdFound = true
			'strMSG = "Please SMR the following Deliverables for " & strProductName & " as soon as possible. The versions listed must be released because they are required to support the factory images. Additional “upgrade” versions may be released at your discretion." & vbcrlf & vbcrlf & "[Their Deliverables Listed Here]" & vbcrlf & vbcrlf
		else
			blnProdFound = false
		end if
		rs.Close
		
	end if
	
    '//Herb
    if request("IDS") = "" then
		blnprodFound = false
    else
        blnprodFound = true
        strProdIds = trim(request("IDS"))
        arrProdIds = Split(strProdIds,",")
    end if

    if not blnprodFound then
        response.End
    end if

    '//Herb
    i = ubound(arrProdIds)
    dim arrProdNames()
    redim arrProdNames(i)
    dim strProdNames
    dim arrDistribution()
    redim arrDistribution(i)
    i = 0
    '''// workarround
    rs.Open "SELECT ID, DOTSName,Distribution,RTMNotifications  from ProductVersion where ID in (" & strProdIds & ") order by DOTSName; ",cn
    if not (rs.EOF and rs.BOF) then
        do while not rs.eof
            arrProdNames(i) = rs("DOTSName")
            arrProdIds(i) = rs("ID")

    		if trim(rs("RTMNotifications") & "") = "" then
			    arrDistribution(i) = CurrentUserEmail & ";" & trim(rs("Distribution") & "")
			elseif instr(trim(rs("RTMNotifications") & ""),CurrentUserEmail) < 1 then
			    arrDistribution(i) = CurrentUserEmail & ";" & trim(rs("RTMNotifications") & "")
            else
			    arrDistribution(i) = trim(rs("RTMNotifications") & "")
			end if

            i = i+ 1
            rs.movenext
        loop
    else
		blnProdFound = false
	end if
	rs.Close
    strProdNames = Join(arrProdNames,", ")

    if not blnProdFound then
        response.write "Unable to find the requested product."
    else
%>

<font size=4 face=verdana><b>Multiple Product RTM Wizard for  <%=strProdNames %> </b></font><BR><BR> 
<font size=2 face=verdana><b><label ID="lblTitle">Enter General RTM Information:</label></b></font>

<form id="frmMain" method="post" action="MultipleRTMSave.asp">
    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
    <input type="hidden" id="txtProductCount" name="txtProductCount" value="<%=ubound(arrProdIds)+1 %>">
    <div id="tabGeneral" style="display: inline;">
        <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
            <% 
                for i = lbound(arrProdIds) to ubound(arrProdIds)
                %>
            <tr>
                <td width="120">
                    <font size="2" face="verdana"><b>RTM&nbsp;Title</b></font><br />
                    <font size="2" face="verdana"><b>(<%=arrProdNames(i) %>) :</b></font><font color="red" size="1">*</font></td>
                <td>
                    <input style="width: 100%" id="txtRTMName<%=cstr(i) %>" name="txtRTMName<%=cstr(i) %>" type="text" value="" maxlength="120">
                </td>
            </tr>
            <% 
                next
                %>
            <tr>
                <td width="120"><font size="2" face="verdana"><b>RTM&nbsp;Date:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                <td>
                    <input style="width: 120px" id="txtRTMDate" name="txtRTMDate" type="text" class="dateselection" value="<%=formatdatetime(now,vbshortdate)%>">
                </td>
            </tr>
            <tr>
                <td width="120"><font size="2" face="verdana"><b>Items&nbsp;to&nbsp;RTM:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                <td>
                    <table>
                        <tr>
                            <td>
                                <input id="chkBIOS" value="1" name="chkBIOS" type="checkbox" language="javascript" onclick="chkBIOS_onclick();">&nbsp;<font id="BIOSTextColor" color="black">System&nbsp;BIOS</font><font color="red" size="1" face="verdana" id="BIOSDisabled">&nbsp;</font>
                            </td>
                            <td style="display: none;">
                                <input id="chkRestore" value="1" name="chkRestore" type="checkbox">&nbsp;<font id="RestoreTextColor" color="black">Restore&nbsp;Media</font>&nbsp;<font color="red" size="1" face="verdana" id="RestoreDisabled">&nbsp;</font>
                            </td>
                            <td>
                                <input id="chkImages" value="1" name="chkImages" type="checkbox">&nbsp;<font id="ImagesTextColor" color="black">Images</font>&nbsp;<font color="red" size="1" face="verdana" id="ImagesDisabled">&nbsp;</font>
                                <input id="chkPatch" value="1" name="chkPatch" type="checkbox">&nbsp;<font id="PatchTextColor" color="black">Patches</font>&nbsp;<font color="red" size="1" face="verdana" id="PatchDisabled">&nbsp;</font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="BIOSAffectivityRow" style="display: none;">
                <td width="120" valign="top"><font size="2" face="verdana"><b>BIOS&nbsp;Affectivity:</b></font>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
                <td>
                    <input id="optCutIn" name="optPhaseIn" type="radio" value="0"> Immediate (Rework All Units)
                    <input id="optPhaseIn" name="optPhaseIn" type="radio" value="1"> Phase-in
                    <input id="optWebOnly" name="optPhaseIn" type="radio" value="2"> Web Release Only
                </td>
            </tr>
            <tr>
                <td width="120" valign="top"><font size="2" face="verdana"><b>Comments:</b></font></td>
                <td>
                    <textarea style="width: 100%; color: blue; font-style: italic" id="txtRTMComments" name="txtRTMComments" cols="80" rows="7" onfocus="return txtRTMComments_onfocus()" onblur="return txtRTMComments_onblur()"><%=RTMCommentsTemplate%></textarea>
                    <textarea style="display: none" id="txtRTMCommentsTemplate" name="txtRTMCommentsTemplate"><%=RTMCommentsTemplate%></textarea>
                </td>
            </tr>
            <%
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
            'if currentuserid = 30 then                 
            %>
            <!--//<tr>
                <td width="120" valign="top"><font size="2" face="verdana"><b>SCMX:</b></font></td>
                <td>
                    <a href="">Upload</a>
                </td>
            </tr>//-->
           <%'end if 
               ' 03/11/2016, Herb, Merged with PBI 16007 and changeset 15097
               %>
        </table>


    </div>
    
    
    <div id="tabBIOS" style="display: none">

    <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%">
        <tr>
            <td width="120" valign="top">
                <font size="2" face="verdana">
                    <b>Comments:&nbsp;&nbsp;&nbsp;</b>
                </font>
            </td>
            <td width="100%">
                <textarea style="width: 100%; color: blue; font-style: italic" id="txtBIOSComments" name="txtBIOSComments" cols="80" rows="7" onfocus="return txtBIOSComments_onfocus()" onblur="return txtBIOSComments_onblur()"><%=BIOSCommentsTemplate%></textarea>
                <textarea style="display: none" id="txtBIOSCommentsTemplate" name="txtBIOSCommentsTemplate"><%=BIOSCommentsTemplate%></textarea>
            </td>
        </tr>
    </table>
    <br>

<%
    for i = lbound(arrProdIds) to ubound(arrProdIds)
%>
    <p>&nbsp;</p>
    <p><b><%=arrProdNames(i) %> &nbsp;</b></p>
    <table ID="tableBIOS<%=cstr(i) %>"  style="border-left: gray thin solid; border-right: gray thin solid; border-bottom: gray thin solid" class="ImageTable" cellpadding="2" cellspacing="0" bgcolor="cornsilk" width="100%">
<%
        rs.open "spListBIOSVersions4Productrtm " & arrProdIds(i) & ",3,0",cn

        if (rs.eof and rs.bof) then

%>
        <tr>
            <td>
                <font size="2" color="red" face="verdana">There are no System BIOS deliverables available to RTM on this product. </font>
            </td>
        </tr>
<%
        end if
%>

<%
            strLastRoot = ""
            do while not rs.eof
        	    BIOSCount = BIOSCount + 1
                if trim(strLastRoot) <> trim(rs("DeliverableName") & "") then
                    response.Write "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=5 bgcolor=wheat><b>&nbsp;" & rs("DeliverableName") & "</b></td></tr>" & vbcrlf            
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""100%""><b>&nbsp;Notes&nbsp;</b></TD></tr>" & vbcrlf
                end if
                strLastRoot = trim(rs("DeliverableName") & "")
                response.Write "<tr bgcolor=ivory>"    
                if rs("targeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSList"" name=""chkBIOSList"" type=""checkbox"" checked></td>"        
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("DeliverableName") & """ PreviewVersion=""" & rs("Version") & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkBIOSList"" name=""chkBIOSList"" type=""checkbox""></td>"        
                end if
                response.Write "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("Version") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>&nbsp;" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>" & "</tr>" & vbcrlf        
                rs.movenext
            loop

            rs.close
        %>

    </table>


<%

    next
       
%>

    </div>

<div ID=tabPatch style="Display:none">
 
                <Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
    	        <TR>
	    	    <TD width=120 valign=top><font size=2 face=verdana><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
		        <td width="100%">
                    <textarea style="width:100%;color: blue; font-style: italic" id="txtPatchComments" name="txtPatchComments" cols="80" rows="5" onfocus="return txtPatchComments_onfocus()" onblur="return txtPatchComments_onblur()"><%=PatchCommentsTemplate%></textarea>
                    <textarea style="display: none" id="txtPatchCommentsTemplate" name="txtPatchCommentsTemplate"><%=PatchCommentsTemplate%></textarea>
                    </TD>
	            </TR>
	            </table>
<%
    

    for i = lbound(arrProdIds) to ubound(arrProdIds)
        'response.Write arrProdIds(i) + ","
   
%>
    <p><b>Patch deliverables on <%=arrProdNames(i) %></b></p>

            <%
            rs.open "spListPatches4ProductRTM " & arrProdIds(i)& ",0" ,cn
            if not(rs.eof and rs.bof) then
            %>
                
			    <TABLE style="BORDER-Left: gray thin solid;BORDER-RIGHT: gray thin solid;BORDER-BOTTOM: gray thin solid" class=ImageTable width=100% ID="tablePatch<%=cstr(i) %>" cellspacing=0 cellpadding=2>
            <%
            else
            %>
                <font size=2 color=red face=verdana>There are no Patch deliverables available to RTM on this product.</font>
            <%
            end if
            strLastRoot = ""
            do while not rs.eof
        	    PatchCount = PatchCount + 1
                if trim(strLastRoot) <> trim(rs("Name") & "") then
                    response.Write vbcrlf & "<tr class=""Row"">"        
                    response.Write "<td valign=top colspan=5 bgcolor=wheat><b>&nbsp;" & rs("Name") & "</b></td></tr>" & vbcrlf     
			        response.Write "<tr bgcolor=cornsilk><TD>&nbsp;</TD><TD>&nbsp;<b>ID</b></TD><TD><b>&nbsp;Version&nbsp;</b></TD><TD><b>&nbsp;TGT&nbsp;</b></TD><TD width=""100%""><b>&nbsp;Notes&nbsp;</b></TD></tr>" & vbcrlf
                end if
                strLastRoot = trim(rs("Name") & "")
                response.Write "<tr bgcolor=ivory>" & vbcrlf 
                
                strPatchContents = "" 
               	set rs2 = server.CreateObject("ADODB.recordset")
                rs2.open "spGetSelectedDepends " & clng(rs("ID")),cn
                do while not rs2.eof
                    if strPatchContents <> "" then
                        strPatchContents = strPatchContents & "<BR>" 
                    end if
                    strPatchContents = strPatchContents & rs2("Name") & " [" & rs2("Version")
                    if trim(rs2("revision")&"") <> "" then
                        strPatchContents = strPatchContents & "," & rs2("revision")
                    end if
                    if trim(rs2("pass")&"") <> "" then
                        strPatchContents = strPatchContents & "," & rs2("pass")
                    end if
                    rs2.movenext
                loop
                rs2.close   
            	set rs2 = nothing

                if trim(strPatchContents) = "" then
                    strPatchContents = "&nbsp;"
                else
                    strPatchContents = strPatchContents & "]"
                    strPatchContents = server.HTMLEncode(strPatchContents)
                end if
                  
                if rs("targeted") then      
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchList"" name=""chkPatchList"" type=""checkbox"" productId="& arrProdIds(i) &" checked></td>"    
                else
                    response.Write "<td valign=top><input PreviewName=""" & rs("Name") & """ PreviewVersion=""" & rs("Version") & """ PreviewContents=""" & strPatchContents & """ PreviewID=""" & rs("ID") & """ value=""" & rs("ID") & """ id=""chkPatchList"" name=""chkPatchList"" type=""checkbox"" productId="& arrProdIds(i) &"></td>"      
                end if
                response.Write vbcrlf & "<td valign=top>&nbsp;" & rs("ID") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("Version") & "&nbsp;&nbsp;&nbsp;</td>"        
                response.Write "<td valign=top align=center>&nbsp;" & replace(replace(trim(rs("targeted")&""),"False","&nbsp;"),"True","X") & "</td>"        
                response.Write "<td valign=top>&nbsp;" & rs("TargetNotes") & "&nbsp;</td>"  & vbcrlf & "</tr>" & vbcrlf        
                rs.movenext
            loop
            if not(rs.eof and rs.bof) then
            %>
			    </table>
            <%
            end if
            rs.close
            %>
<%

    next
       
%>
    <br />

</div>

<div ID=tabRestore style="Display:none">
     
                <Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
    	        <TR>
	    	    <TD width=120 valign=top><font size=2 face=verdana><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
		        <td width="100%">
                    <textarea style="width:100%;color: blue; font-style: italic" id="txtRestoreComments" name="txtRestoreComments" cols="80" rows="9" onfocus="return txtRestoreComments_onfocus()" onblur="return txtRestoreComments_onblur()"><%=RestoreCommentsTemplate%></textarea>
                    <textarea style="display: none" id="txtRestoreCommentsTemplate" name="txtRestoreCommentsTemplate"><%=RestoreCommentsTemplate%></textarea>
                    </TD>
	            </TR>
	            </table><br>
			    <TABLE style="BORDER-Left: gray thin solid;BORDER-RIGHT: gray thin solid;BORDER-BOTTOM: gray thin solid" class=ImageTable width=100% ID=TableRestore cellspacing=0 cellpadding=2>

                    <tr><td>

                    <font size=2 color=red face=verdana>There are no Restore Media deliverables available to RTM on this product.</font></td>
	            </tr>
</table>
         

</div>
<div ID=tabImages style="Display:none">
<%
    dim imagecount
    dim strSQL
    dim strLastProductId
    strLastProductId = ""

    imagecount = 0
	strAllImages = ""
    'strSQL = "spListImagesForProduct2RTM " & clng(request("ID"))
    strSQL = "spListImagesForProducts2RTMs '" & strProdIds & "';"

    if blnFusion and false then 'This was the method where they had to pick a whole product drop
        'not really running

	    rs.open strSQL,cn,adOpenForwardOnly
	    if rs.EOF and rs.BOF then
		    strAllImages = vbcrlf & "<TR><TD colspan=11><FONT size=1 face=verdana>No images defined for this product.</font></td></TR>" & vbcrlf 
	    else
            strLastProductDrop = ""
            do while not rs.eof
                if lcase(trim(strLastProductDrop)) <> lcase(trim(rs("ProductDrop") & "")) then
'                    strAllImages = strAllImages & "<tr><td><input id=""chkProductDrop"" type=""checkbox"" /></td><td colspan=10>" & rs("ProductDrop") & "</td></TR>"
			        if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				        RowStyle = "style=display:none"
			        else
    				    rowStyle=""
			        end if
                    if strLastProductDrop <> "" then
                        strAllImages= strAllImages & "</table><br>" & vbcrlf 
                    end if
			         strAllImages= strAllImages &  "<TR " & rowStyle & " id=DropRow" & rs("ProductDropID") & " LANGUAGE=javascript onmouseover=""return DropRow_onmouseover(" & rs("ProductDropID") & ")"" onmouseout=""return DropRow_onmouseout(" & rs("ProductDropID") & ")"" onclick=""return DropRow_onclick(" & rs("ProductDropID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("ProductDropID") & """ name=""Drop" & rs("ProductDropID") & """ type=""checkbox"" class=""chkDrop"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkDrop_onclick()""></TD>" 
                     strAllImages= strAllImages & "<td>" & rs("ProductDrop") & "&nbsp;&nbsp;&nbsp;</td>"
                     strAllImages= strAllImages & "<td colspan=10>" & rs("OSList") & "</td>"
			         strAllImages= strAllImages &  "<TR style=""display:none"" id=DropContents" & rs("ProductDropID") & " bgcolor=cornsilk><td>&nbsp;</td>"
                     strAllImages= strAllImages & "<td colspan=10>" 
                     strAllImages= strAllImages & "<BR><table width=""100%"" cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD><b>Region</b></TD><TD><b>Brands</b></TD><TD><b>OS</b></TD><TD><b>Comments</b></TD></THEAD>" & vbcrlf 
                end if

                if trim(rs("ProductDrop") & "") =  "" then
                    strimagePreview = "FUSION|" & rs("ID") & "|No Product Drop Number Defined|"
                else
                    strimagePreview = "FUSION|" & rs("ID") & "|" & ucase(rs("ProductDrop")) & "|" 
	            end if
                strimagePreview = server.HTMLEncode(strimagePreview & rs("Brands") & "|" & rs("OS") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))

                strAllImages= strAllImages  & vbcrlf & "<tr>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & "<INPUT class=""" & trim(rs("ProductDropID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " style=""display:none"" name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16"">" & rs("OptionConfig") & "&nbsp;-&nbsp;" & rs("Region") & "<input style=""display:none"" id=""txtImagePreview"" type=""text"" value=""" & strImagePreview & """>" &  "</td>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & rs("Brands") & "</td>"
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & rs("OS") & "</td>"
                strCommentCollection = ""
                if trim(rs("DefinitionComments") & "" ) <> "" then
                    strCommentCollection = trim(rs("DefinitionComments") & "" )
                end if
                if trim(rs("Comments") & "" ) <> "" and strCommentCollection <> "" then
                    strCommentCollection = strCommentCollection & "<br>" & trim(rs("DefinitionComments") & "" )
                elseif trim(rs("Comments") & "" ) <> ""  then
                    strCommentCollection = trim(rs("DefinitionComments") & "" )
                end if
                strAllImages= strAllImages & "<td bgcolor=""ivory"">" & strCommentCollection & "&nbsp;</td>"
                strAllImages= strAllImages & "</TR>" 
                strLastProductDrop= rs("ProductDrop") & "" 
                imagecount = imagecount + 1
                rs.movenext
            loop
            if imagecount > 0 then
                strAllImages= strAllImages & "</td></TR></table>"  & vbcrlf 
            end if
        end if
        rs.close
    elseif blnFusion then ' This section is for converged notebooks
       
	    rs.open strSQL,cn,adOpenForwardOnly
	    lastDefinition = ""
	    strLastProductId = ""    

	    if rs.EOF and rs.BOF then
		    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></TR>" & vbcrlf 
	    else
		    do while not rs.EOF
			    imagecount = imagecount + 1

                if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
				    strAllImages = strAllImages & strBase
				    strAllImages = strAllImages &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Config</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></TR>" & vbcrlf 
				    strRows = ""
				    YesCount = 0
				    NoCount = 0
				    MixedCount=0
			    end if
			    lastdefinition = rs("DefinitionID")
    
                if (strLastProductId <> rs("ProductVersionId")) and (strLastProductId <>"") then
                    strAllImages = strAllImages & vbcrlf & " <TR style='FONT-WEIGHT: bold; BACKGROUND-COLOR: wheat'> <TD></TD> <TD align=left >ID</TD> <TD align=left >Product Drop</TD> <TD align=left >Product Version&nbsp;</TD> <TD align=left >Brands</TD>  <TD align=left >OS</TD> <TD align=left colspan=6>Comments</TD> </TR> " & vbcrlf
                end if
                strLastProductId = rs("ProductVersionId")	
    

			    strLanguageList = rs("OSLanguage")
			    if trim(rs("OtherLanguage") & "") <> "" then
				    strLanguageList = strLanguageList & "," & rs("OtherLanguage")
			    end if	
			
		
            
                if trim(rs("ProductDrop") & "") =  "" then
                    strimagePreview = "FUSION|" & rs("ID") & "|No Product Drop Number Defined|"
                else
                    strimagePreview = "FUSION|" & rs("ID") & "|" & ucase(rs("ProductDrop")) & "|" 
	            end if
                strimagePreview = server.HTMLEncode(strimagePreview & rs("Brands") & "|" & rs("OS") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
			    strCellColor = "ivory"
			    if instr(strImages,", " & rs("ID") & ",") > 0 or not blnImages then
			
					    strCellColor = "ivory"
					    YesCount = YesCount + 1
					    if request("Type") = "1" then
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()""  prodId=""" & rs("ProductVersionId") &""" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
					    else
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()""  prodId=""" & rs("ProductVersionId") &""" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagepreview & """></TD>"							
					    end if
			    else
				    if request("Type") = "1" then
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  prodId=""" & rs("ProductVersionId") &""" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
				    else
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  prodId=""" & rs("ProductVersionId") &""" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
				    end if
				    NoCount = NoCount + 1
			    end if
			
			    if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				    RowStyle = "style=display:none"
			    else
				    rowStyle=""
			    end if
			    if YesCount = 0  and MixedCount=0 then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
			        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      end if
			    elseif NoCount=0 and MixedCount=0  then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      end if
			      TotalImageDefsChecked= TotalImageDefsChecked + 1
			    else
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      else 
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      end if
			      TotalImageDefsChecked = TotalImageDefsChecked + 1
			    end if
			    strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("ProductDrop") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                strBase = strBase &  "<TD nowrap>" & rs("ProductVersionName") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("Brands") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" &  rs("OS")  & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
'			    strBase = strBase &  "<TD nowrap>" & rs("brand") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("Comments") & "&nbsp;</TD>"
			    'strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
			    strBase = strBase &  "</TR>" & vbcrlf 
			
			    strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
'			    strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
			    strRows = strRows & "<TD width=130>" & rs("Region") & "</TD>"
			    strRows = strRows & "<TD>" & rs("countrycode") & "</TD>"

			    if trim(rs("OtherLanguage") & "") <> "" then
			    '  if request("Type") = "1" then
			        strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			     ' else
				 '   strRows = strRows & "<TD ID=""Row" & rs("ID") & """ bgcolor=" & strCellColor & " width=70>" & strSavedLanguages & "</a></TD>"
			     ' end if
			    else
				    strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			    end if				
			    strRows = strRows & "<TD width=65>" & rs("Keyboard") & "</TD>"
			    strRows = strRows & "<TD width=75>" & rs("powercord") & "</TD>"
			    strRows = strRows & "</TR>" & vbcrlf 

			    rs.MoveNext
		    loop
		    if imagecount = 0 then
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No active images defined for this product.</font></td></TR>" & vbcrlf 
		    end if
		    strAllImages = strAllImages & strBase
		    strAllimages = strAllImages & "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Config</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></TR>" & vbcrlf 
		    strRows = ""
		

	    end if	
	    rs.Close


    else ' This section is for legacy notebooks
        
	    rs.open strSQL,cn,adOpenForwardOnly
	    lastDefinition = ""
        strLastProductId = ""
	    
	    if rs.EOF and rs.BOF then
		    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></TR>" & vbcrlf 
	    else
		    do while not rs.EOF
			    imagecount = imagecount + 1

			    if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
				    strAllImages = strAllImages & strBase
				    strAllImages = strAllImages &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></TR>" & vbcrlf 
				    strRows = ""
				    YesCount = 0
				    NoCount = 0
				    MixedCount=0
			    end if
			    lastdefinition = rs("DefinitionID")
			
                if (strLastProductId <> rs("ProductVersionId")) and (strLastProductId <>"") then
                    strAllImages = strAllImages & vbcrlf & " <TR style='FONT-WEIGHT: bold; BACKGROUND-COLOR: wheat'> <TD>&nbsp; </TD> <TD align=left>ID</TD> <TD align=left>SKU</TD> <TD align=left>Product Version&nbsp;</TD> <TD align=left>Model</TD> <TD align=left>OS</TD> <TD align=left>Apps&nbsp;Bundle</TD> <TD align=left>BTO/CTO</TD>  <TD align=left>Comments</TD> </TR> " & vbcrlf
                end if
                strLastProductId = rs("ProductVersionId")	


			    strLanguageList = rs("OSLanguage")
			    if trim(rs("OtherLanguage") & "") <> "" then
				    strLanguageList = strLanguageList & "," & rs("OtherLanguage")
			    end if	
			
			    strSavedLanguages = getLanguages(strImages,rs("ID"))
			    if strSavedLanguages = "" then
				    strSavedLanguages = strLanguageList
			    end if
            
                if trim(rs("Skunumber") & "") =  "" then
                    strimagePreview = server.HTMLEncode(rs("ID") & "|No SKU Defined|" & rs("Brand") & "|" & rs("OS") & "|" & rs("SW") & "|" & rs("ImageType") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
                else
                    strimagePreview = server.HTMLEncode(rs("ID") & "|" & 	replace(ucase(rs("SKUNumber")) & "","XX",mid(rs("Dash"),2,2)) & "|" & rs("Brand") & "|" & rs("OS") & "|" & rs("SW") & "|" & rs("ImageType") & "|" & rs("DefinitionComments") & "|" & rs("Region") & "|" & rs("OptionConfig"))
	            end if
			    strCellColor = "ivory"
			    if instr(strImages,", " & rs("ID") & ",") > 0 or not blnImages then
				    if strLanguageList <> strSavedLanguages then
					    strCellColor = "mistyrose"
					    MixedCount = MixedCount + 1
					    if request("Type") = "1" then
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""   LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
					    else
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""   LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strimagepreview & """></TD>"
					    end if
				    else
					    strCellColor = "ivory"
					    YesCount = YesCount + 1
					    if request("Type") = "1" then
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""   LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
					    else
						    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""   LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagepreview & """></TD>"							
					    end if
				    end if
			    else
				    if request("Type") = "1" then
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT Disabled class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""  value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
				    else
					    strRows = strRows & "<TR bgcolor=ivory><TD><INPUT class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage prodId=""" & rs("ProductVersionId") &"""  value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""><input style=""display:none"" id=""txtImagePreview"" type=""text"" prodId=""" & rs("ProductVersionId") &""" value=""" & strImagePreview & """></TD>"
				    end if
				    NoCount = NoCount + 1
			    end if
			
			    if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or rs("StatusID") = 2) then
				    RowStyle = "style=display:none"
			    else
				    rowStyle=""
			    end if
			    if YesCount = 0  and MixedCount=0 then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
			        strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      end if
			    elseif NoCount=0 and MixedCount=0  then
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      else
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()""></TD>" 
			      end if
			      TotalImageDefsChecked= TotalImageDefsChecked + 1
			    else
			      if request("Type") = "1" then
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT Disabled id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      else 
				    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkBase_onclick()"" indeterminate=-1></TD>"
			      end if
			      TotalImageDefsChecked = TotalImageDefsChecked + 1
			    end if
			    strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
                strBase = strBase &  "<TD nowrap>" & rs("ProductVersionName") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("brand") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
			    strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
			    strBase = strBase &  "</TR>" & vbcrlf 
			
			    strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
			    strRows = strRows & "<TD width=130>" & rs("Region") & "</TD>"
			    strRows = strRows & "<TD width=50>" & rs("CountryCode") & "</TD>"
			    strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
			    if trim(rs("OtherLanguage") & "") <> "" then
			      if request("Type") = "1" then
			        strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			      else
				    strRows = strRows & "<TD ID=""Row" & rs("ID") & """ bgcolor=" & strCellColor & " width=70>" & strSavedLanguages & "</a></TD>"
			      end if
			    else
				    strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
			    end if				
			    'strRows = strRows & "<TD width=70>" & strLanguageList & "</TD>"
			    strRows = strRows & "<TD width=65>" & rs("Keyboard") & "</TD>"
			    strRows = strRows & "<TD width=75>" & rs("Powercord") & "</TD>"
			    strRows = strRows & "</TR>" & vbcrlf 

			    rs.MoveNext
		    loop
		    if imagecount = 0 then
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No active images defined for this product.</font></td></TR>" & vbcrlf 
		    end if
		    strAllImages = strAllImages & strBase
		    strAllimages = strAllImages & "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></TR>" & vbcrlf 
		    strRows = ""
		

	    end if	
	    rs.Close
    end if

	if TotalImageDefsChecked > 0 then
		strAllChecked="checked"
	else
		strAllChecked=""
	end if

    if imagecount <> 0 then
%>
                <Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
                <tr>
    	    	    <TD width="120" valign="top"><font size=2 face=verdana><b>File&nbsp;Upload:&nbsp;&nbsp;&nbsp;</b></font></td>
                    <td>
                        <div id="UploadAddLinks1"><a href="javascript: UploadZip(1);">Upload</a></div>
                        <div id="UploadRemoveLinks1" style="display:none"><a href="javascript: UploadZip(1);">Change</a> | <a href="javascript: RemoveUpload(1);">Remove</a>&nbsp;&nbsp;<b>File:&nbsp;</b><label id="UploadPath1"></label></div>
                        <input id="txtAttachmentPath1" name="txtAttachmentPath1" type="hidden" value="" />
                        <input id="txtUploadPath1" name="txtUploadPath1" type="hidden" value="" />
                    </td>
                </tr>
    	        <TR>
	    	    <TD width=120 valign=top><font size=2 face=verdana><b>Comments:&nbsp;&nbsp;&nbsp;</b></font></td>
		        <td width="100%">
                    <textarea style="width:100%;color: blue; font-style: italic" id="txtImageComments" name="txtImageComments" cols="80" rows="18"  onfocus="return txtImageComments_onfocus()" onblur="return txtImageComments_onblur()"><%=ImageCommentsTemplate%></textarea>
                    <textarea style="display: none" id="txtImageCommentsTemplate" name="txtImageCommentsTemplate"><%=ImageCommentsTemplate%></textarea>
                    </TD>
	            </TR>
	            </table><br>
                <%if blnFusion then %>
		            <table class="ImageTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			            <THEAD bgcolor=Wheat>
				            <% if request("Type") = "1" then%>
				            <TH align=left><INPUT disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				            <%else%>
				            <TH align=left><INPUT type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
                            <%end if%>
				            <TH align=left >ID</TH>
				            <TH align=left >Product Drop</TH>
                            <TH align=left >Product Version&nbsp;</TH>
				            <TH align=left >Brands</TH>
				            <TH align=left >OS</TH>
				            <TH align=left colspan="6" width="100%">Comments</TH>
			            </thead>
			            <%=strAllImages%>

		            </table>

                <%else%>
		            <table class="ImageTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			            <THEAD bgcolor=Wheat>
				            <% if request("Type") = "1" then%>
				            <TH align=left><INPUT disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				            <%else%>
				            <TH align=left><INPUT type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
                            <%end if%>
				            <TH align=left>ID</TH>
				            <TH align=left>SKU</TH>
                            <TH align=left>Product Version&nbsp;</TH>
				            <TH align=left>Model</TH>
				            <TH align=left>OS</TH>
				            <TH align=left>Apps&nbsp;Bundle</TH>
				            <TH align=left>BTO/CTO</TH>
				            <TH align=left>Comments</TH>
			            </thead>
			            <%=strAllImages%>

		            </table>
               <%end if%>
		<%else
                response.write "<label id=""UploadPath1""></label>"
                Response.write "<font size=2 color=red face=verdana>There are no Images available to RTM on this product.</font>"
		end if%>


<INPUT style="Display:none" type="checkbox" id=chkAllChecked name=chkAllChecked>
</div>

<div ID=tabAlerts style="Display:none">

<%
   

    for i = lbound(arrProdIds) to ubound(arrProdIds)
        'response.Write arrProdIds(i) + ","
        lngProdId = cLng(arrProdIds(i))
%>
    <div ID="divProductAlerts<%=cStr(i) %>" name="divProductAlerts" style="Display:none">
    <p><b>Alerts of <%=arrProdNames(i) %></b></p>
	
	
	
    <font size=2 face=verdana><b>Build Level Alerts: </b>&nbsp;&nbsp;<font id=BuildLevelTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkBuildLevel" name="chkBuildLevel" type="checkbox" value="1" onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="BuildLevel"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtBuildLevelComments" class="txtBuildLevelComments" name="txtBuildLevelComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id="BuildLevelAlertDetails" class="BuildLevelAlertDetails">
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtBuildLevelIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=1&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=BuildLevelIFrame name=BuildLevelIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>
<br>
    <font size=2 face=verdana><b>Distribution Alerts: </b>&nbsp;&nbsp;<font id=DistributionTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkDistribution" name="chkDistribution" type="checkbox" value="1" onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Distribution"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtDistributionComments" name="txtDistributionComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=DistributionAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtDistributionIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=2&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=DistributionIFrame name=DistributionIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>

<br>
    <font size=2 face=verdana><b>Certification Alerts: </b>&nbsp;&nbsp;<font id=CertificationTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkCertification" name="chkCertification" type="checkbox" value="1" onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Certification"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtCertificationComments" name="txtCertificationComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=CertificationAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtCertificationIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=3&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=CertificationIFrame name=CertificationIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>

<br>
    <font size=2 face=verdana><b>Workflow Alerts: </b>&nbsp;&nbsp;<font id=WorkflowTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkWorkflow" name="chkWorkflow" type="checkbox" value="1" onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Workflow"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></td>
    <td>        
        <textarea id="txtWorkflowComments" name="txtWorkflowComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=WorkflowAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtWorkflowIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=4&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=WorkflowIFrame name=WorkflowIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>


<br>
    <font size=2 face=verdana><b>Availability Alerts: </b>&nbsp;&nbsp;<font id=AvailabilityTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkAvailability" name="chkAvailability" type="checkbox" value="1" onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Availability"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtAvailabilityComments" name="txtAvailabilityComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=AvailabilityAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtAvailabilityIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=5&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=AvailabilityIFrame name=AvailabilityIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>

<br>
    <font size=2 face=verdana><b>Developer Alerts: </b>&nbsp;&nbsp;<font id=DeveloperTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkDeveloper" name="chkDeveloper" type="checkbox" value="1"  onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Developer"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtDeveloperComments" name="txtDeveloperComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=DeveloperAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtDeveloperIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=6&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=DeveloperIFrame name=DeveloperIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>

<br>
    <font size=2 face=verdana><b>Root Deliverable Alerts: </b>&nbsp;&nbsp;<font id=RootTypeText color=green></font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkRoot" name="chkRoot" type="checkbox" value="1"  onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="Root"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtRootComments" name="txtRootComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=RootAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtRootIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=7&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=RootIFrame name=RootIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>

<br>
    <font size=2 face=verdana><b>OTS Primary Alerts - <%=arrProdNames(i)%>:&nbsp;&nbsp;</b><font id=OTSPrimaryFilterText color=green>(P0/P1 observations for <label id=OTSPrimaryTypeText></label>)</font><BR></font>
<Table border=1 cellpadding=2 cellspacing=0 bgcolor=cornsilk bordercolor=tan width=100%>
	<TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Signoff:</b></font></td>
	<TD>
        <input id="chkOTSPrimary" name="chkOTSPrimary" type="checkbox" value="1"  onclick="chkAlert_onClick(<%=cStr(i) %>,this);" alertname="OTSPrimary"> I have reviewed these alerts.
    </td>
    </TR>
    <TR>
	<TD width=120 valign=top><font size=2 face=verdana><b>Comments:</b></font></td>
    <td>        
        <textarea id="txtOTSPrimaryComments" name="txtOTSPrimaryComments" rows="4" cols="90" style="width:100%"></textarea>
    </TD>
	</TR>
	<TR id=OTSPrimaryAlertDetails>
		<TD width=120 valign=top><font size=2 face=verdana><b>Alerts:</b></font></td>
		<td>
            <textarea id="txtOTSPrimaryIFramesrc" style="display:none" cols="20" rows="2">../ReadinessReport.asp?ProdID=<%=lngProdId%>&Sections=8&TableOnly=1&RTMSignoff=1&RTMID=0</textarea>
            <iframe id=OTSPrimaryIFrame name=OTSPrimaryIFrame marginwidth=0 width=100% src="../maint/blank_loading.htm">
            </iframe>
		</TD>
	</TR>

</table>
</div>
<%

    next
    
%>
</div>


<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=trim(clng(request("ID")))%>">
<INPUT type="hidden" id=txtProductName name=txtProductName value="<%=strProductName%>">
<INPUT type="hidden" id=txtCurrentUserEmail name=txtCurrentUserEmail value="<%=CurrentUserEmail%>">



<div ID=tabPreview style="Display:none">
<%
    for i = lbound(arrProdIds) to ubound(arrProdIds)
        'response.Write arrProdIds(i) + ","
        lngProdId = cLng(arrProdIds(i))
%>
    <hr />
    <br />
    <div id="divPreview<%=cStr(i) %>" prodID="<%=lngProdId %>" prodName="<%=arrProdNames(i) %>">
        <p><font size=2 face=verdana><b><%=arrProdNames(i) %></b></font><label id="lblPreviewMsg<%=cStr(i) %>" style="color: red; font-size: x-small;"></label></p>
        <table width="100%" border="0">
            <tr>
                <td valign="top"><b>Email:&nbsp;&nbsp;</b></td>
                <td width="100%">
                    <textarea style="width: 100%" id="txtNotify" name="txtNotify" rows="3"><%=arrDistribution(i)%></textarea></td>
                <td valign="top">
                    <button style="height: 50" id="cmdAdd" name="cmdAdd" language="javascript" onclick="return cmdAdd_onclick(<%=cStr(i) %>)">Address<br>
                        Book</button>
                </td>
            </tr>
        </table>


        <font size=2 face=verdana><b>Preview:</b></font>
        <div style="padding-left:10;padding-right:10;padding-bottom:10;padding-top:10;border-right: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: auto; BORDER-LEFT: steelblue 1px solid; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 100%; BACKGROUND-COLOR: white" id=PreviewPage>
        </div>

    </div>
<% 
    next  
%>
</div>

    <textarea style="display:none" id="txtEmailPreview" name="txtEmailPreview"></textarea>


    
</form>
<%
        dim strExistingRTMTitles
        'Load existing RTm titles
        strExistingRTMTitles = ""
        for i = lbound(arrProdIds) to ubound(arrProdIds)
	        rs.Open "spListProductRTMTitles " & clng(arrProdIds(i)),cn,adOpenForwardOnly
            do while not rs.eof
                strExistingRTMTitles = strExistingRTMTitles & "<option prodId=" & arrProdIds(i) & " prodIdx=" & cstr(i) & " >" & rs("Title") & "</option>" & vbcrlf
                rs.movenext
            loop
            rs.close
        next
    end if

	cn.Close
	set cn = nothing
	set rs = nothing
%>

<%
function GetLanguages(strImages, strID)
	dim strTemp
	
	if instr(strImages,trim(strID) & "=") = 0 then
		GetLanguages = ""
	else
		strTemp = mid(strimages,instr(strImages,trim(strID) & "=")+ len(trim(strID) & "="))
		strTemp = mid(strTemp,1,instr(strTemp,")") -1) 'Strip off )...
		GetLanguages = strTemp
	end if
	
end function
%>
    <select style="display:none" id="cboTitles" name="cboTitles">
        <%=strExistingRTMTitles%>
    </select>
    <input style="display:none" id="txtImageCount" type="text" value="<%=imagecount%>">
    <input style="display:none" id="txtBIOSCount" type="text" value="<%=BIOScount%>">
    <input style="display:none" id="txtRestoreCount" type="text" value="<%=Restorecount%>">
    <input style="display:none" id="txtPatchCount" type="text" value="<%=Patchcount%>">
</BODY>
</HTML>


