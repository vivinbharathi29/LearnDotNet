<%@ Language=VBScript %>
<html>
<head>
<title>ProgramButtons</title>
    <link href="../../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../../Scripts/shared_functions.js"></script>
    <script src="../../Scripts/verifyEmailAddress.js"></script>
    <script src="../../Scripts/PulsarPlus.js"></script>
    <script src="../../Scripts/Pulsar2.js"></script>
    <script type="text/javascript">
        $(function () {
            if (document.getElementById("preferredLayout").value == "" && document.getElementById("preferredLayout").value != 'pulsar2') {
                $('#cmdEditCancel').css('display', 'none');
            }
            $("input:button").button();
        });
    </script>

<script type="text/javascript">
<!--

    <!-- #include file = "../../includes/Date.asp" -->

   
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

    var followMKTName;
    function VerifySave(){
        var blnSuccess = true;
        var blnFoundRelease = false;
        var blnFoundProductName = false;
        var blnFoundArchiveProductName = false;
        var blnDesktopProduct = false;
        var i;
        var phase = window.parent.frames["UpperWindow"].ProgramInput.cboPhase.value;
        var factory = window.parent.frames["UpperWindow"].ProgramInput.cboFactory.value;
        var intDCROwner = 0;
        intDCROwner = parseInt(window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.value); //1,2
        followMKTName = window.parent.frames["UpperWindow"].ProgramInput.hdnEnableFollowMarketingName.value;
        //check if a brand was checked
        if (followMKTName == 0) { 
        var strBrand ="";
            for(i=0;i<window.parent.frames["UpperWindow"].ProgramInput.chkBrands.length;i++)
            {
                if (window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].checked)
                {
                    if (strBrand=="")
                        strBrand = window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].value;
                    else
                        strBrand = strBrand + ", " + window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].value;
                }
            }
        }
        var productId;
        var isClone = 0;

        if (window.location != undefined) {
            var queryString = window.location.search.replace("?", "");
            var pieces = queryString.split("&");
            if (pieces.length > 1) {
                for (var i = 0; i <= pieces.length - 1; i++) {
                    if (pieces[i] != undefined && pieces[i].indexOf("ID=") > -1) {
                        var val = pieces[i];
                        productId = val.replace("ID=", "");
                    } else if (pieces[i] != undefined && pieces[i].indexOf("clone=") > -1) {
                        var val = pieces[i];
                        isClone = val.replace("clone=", "");
                    }
                }
            }
        }

        var VersionText=trim(window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value);
        if (VersionText.length >= 3)
        {VersionText = VersionText.substr(VersionText.length-3,3).toUpperCase();
        }
       
        var strProductName = trim(window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value);
        var stroldProductName=trim(window.parent.frames["UpperWindow"].ProgramInput.tagProductNameBase.value);
        var strExistingproductName="";       
       
        for (i=0;i<window.parent.frames["UpperWindow"].ProgramInput.cboCompleteProductList.length;i++) {
            strExistingproductName = window.parent.frames["UpperWindow"].ProgramInput.cboCompleteProductList.options[i].text.toUpperCase();
            if ((productId == 0 || productId == undefined || isClone == "1" || stroldProductName.toUpperCase() !=strProductName.toUpperCase() ) &&   strExistingproductName == strProductName.toUpperCase() )
            {
                blnFoundProductName = true;
                break;
            }
        }

        for (i=0;i<window.parent.frames["UpperWindow"].cboInactiveProducts.length;i++)
            if (window.parent.frames["UpperWindow"].cboInactiveProducts.options[i].text.toUpperCase() == window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value.toUpperCase() + " " + window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value.toUpperCase() )
                blnFoundArchiveProductName = true;
        if(window.parent.frames["UpperWindow"].ProgramInput.cboBusinessSegmentID.options[window.parent.frames["UpperWindow"].ProgramInput.cboBusinessSegmentID.selectedIndex].innerHTML.toLowerCase().indexOf("desktop") > -1) {
            blnDesktopProduct = true;
        }
        if (trim(window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value) =="")
        {
            window.alert("Please enter Product Name"); 
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.focus();
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.select();
        }
        else if (blnFoundProductName)
        {
            window.alert("A product with that name already exists.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.focus();
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.select();
        }
        else if (blnFoundArchiveProductName)
        {
            window.alert("An inactive product with that name already exists.  Please reactivate the existing product or choose another name.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.focus();
            window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.select();
        }
        else if ( window.parent.frames["UpperWindow"].ProgramInput.txtID == "" && (VersionText=="PRS" || VersionText=="PAV" || VersionText=="SMB" || VersionText=="ENT" || VersionText=="WKS"  || VersionText=="TAB") )
        {
            window.alert("Please contact Dave Whorton before adding PRS, PAV or SMB products.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtVersion.focus();
            window.parent.frames["UpperWindow"].ProgramInput.txtVersion.select();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboBusinessSegmentID.selectedIndex == 0)
        {
            window.alert("Business Segment is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboBusinessSegmentID.focus();
        }
        /* LY BEGINNNING OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM */
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboProductLine.selectedIndex == 0
            && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2
            )
        {
            window.alert("Product Line is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboProductLine.focus();
        }
        else if (window.parent.frames["UpperWindow"].ProgramInput.txtProductRelease.value == ""
            && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("Product Release is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
        }
            /* LY END OF CHANGE - ADD PRODUCT LINE TEXT FIELD TO FORM */
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboFamily.selectedIndex == 0)
        {
            window.alert("Product Family is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboFamily.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboSM.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("System Manager is required.    This may be the same person as the PM if your orgainzation does not have System Managers.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboSM.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.hdnIsDesktop.value == "" && window.parent.frames["UpperWindow"].ProgramInput.hdnIsCommercial.value == "YES" && window.parent.frames["UpperWindow"].ProgramInput.cboPM.selectedIndex == 0)
        {
            window.alert("Configuration Manager is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboPM.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.hdnIsDesktop.value == "" && window.parent.frames["UpperWindow"].ProgramInput.hdnIsCommercial.value == "" && window.parent.frames["UpperWindow"].ProgramInput.cboTDCCM.selectedIndex == 0)
        {
            window.alert("Program Office Program Manager is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboTDCCM.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.hdnIsDesktop.value == "YES" && window.parent.frames["UpperWindow"].ProgramInput.cboSCMOwner.selectedIndex == 0)
        {
            window.alert("SCM Owner is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboSCMOwner.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.hdnIsDesktop.value == "YES" && window.parent.frames["UpperWindow"].ProgramInput.cboEngineeringDataManagement.selectedIndex == 0)
        {
            window.alert("Engineering Data Management is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboEngineeringDataManagement.focus();
        }

        else if(window.parent.frames["UpperWindow"].ProgramInput.cboSEPM.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("System Engineering PM is required.  This may be the same person as the PM if your orgainzation does not have System Engineering PMs.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
            window.parent.frames["UpperWindow"].ProgramInput.cboSEPM.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboToolsPM.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value == 2)
        {
            window.alert("Project Manager is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboToolsPM.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboDevCenter.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2 && !blnDesktopProduct)
        {
            window.alert("Development Center is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboDevCenter.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboPartner.selectedIndex == 0  && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("ODM Partner is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboPartner.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboPreinstall.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2 && !blnDesktopProduct)
        {
            window.alert("Preinstall Team is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboPreinstall.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.cboReleaseTeam.selectedIndex == 0 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2 && !blnDesktopProduct)
        {
            window.alert("Release Team is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.cboReleaseTeam.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.value == "" && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("Product Distribution list is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.value.length > 1000 && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value != 2)
        {
            window.alert("Product Distribution list can not be longer than 1000 characters.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.focus();
        }
        else if((window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.value.length > 1) && (!VerifyEmail(window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.value)))
        {// verify email address for Product Distribution list
            window.alert("You must enter a valid SMTP email address for Product Distribution list.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.select();
            window.parent.frames["UpperWindow"].ProgramInput.txtDistribution.focus();
        }
        else if((window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.value.length > 1) && (!VerifyEmail(window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.value)))
        {// verify email address for txtActionNotifyList
            window.alert("You must enter a valid SMTP email address for Notify On Closure.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.select();
            window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.focus();
        }
        else if((window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.value.length > 1) && (!VerifyEmail(window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.value)))
        {// verify email address for txtDCRApproverList
            window.alert("You must enter a valid SMTP email address for DCR Approvers.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.select();
            window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.value.length > 1000  && window.parent.frames["UpperWindow"].ProgramInput.cboType.options[window.parent.frames["UpperWindow"].ProgramInput.cboType.selectedIndex].value == 2)
        {
            window.alert("Notify on Closure list can not be longer than 1000 characters.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtActionNotifyList.focus();
        }
        else if(window.parent.frames["UpperWindow"].ProgramInput.txtDescription.value == "")
        {
            window.alert("Description is required.");
            blnSuccess = false;
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDescription.focus();
        }
        else if(phase == "2" || phase == "3" || phase =="4")
            {
                if(strBrand == "")
                {
                    alert("Brand is required to be selected when product Phase is changed to Development, Production or Post-Production.");
                    blnSuccess = false;
                }
        }
        if (window.parent.frames["UpperWindow"].ProgramInput.Radio2.checked && window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.value == "") {
            blnSuccess = false;
            window.alert("You must enter a valid email address for DCR Approver List.");
            window.parent.frames["UpperWindow"].SelectTab("General");
            window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.select();
            window.parent.frames["UpperWindow"].ProgramInput.txtDCRApproverList.focus();
        }
        if(phase != "1"){
            //if Product is not in Definition, Default DCR Owner is required
            if(intDCROwner ==1 &&  window.parent.frames["UpperWindow"].ProgramInput.cboPM.selectedIndex == 0) {
                blnSuccess = false;
                window.parent.frames["UpperWindow"].SelectTab("General");
                window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.focus();   
                if(window.confirm("[Default DCR Owner]\n\nConfiguration Manager has not been assigned in System Team. \n\nDo you want to assign Program Office Program Manager to Default DCR Owner?\n")){
                    window.parent.frames["UpperWindow"].SelectTab("General");
                    window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.selectedIndex = 1; //automatically select Program Office Manager
                    window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.focus();    
                }else{
                    window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
                    window.parent.frames["UpperWindow"].ProgramInput.cboPM.focus();                        
                }
            }	 
            if(intDCROwner == 2 && window.parent.frames["UpperWindow"].ProgramInput.cboTDCCM.selectedIndex == 0) {
                blnSuccess = false;
                window.parent.frames["UpperWindow"].SelectTab("General");
                window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.focus();   
                if(window.confirm("[Default DCR Owner]\n\nProgram Office Program Manager has not been assigned in System Team. \n\nDo you want to assign Configuration Manager to Default DCR Owner?\n")){
                    window.parent.frames["UpperWindow"].SelectTab("General");
                    window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.selectedIndex = 0; //automatically select Configuration Manager
                    window.parent.frames["UpperWindow"].ProgramInput.cboDCRDefaultOwner.focus(); 	
                }else{
                    window.parent.frames["UpperWindow"].SelectTab("SystemTeam");
                    window.parent.frames["UpperWindow"].ProgramInput.cboTDCCM.focus(); 
                }
            } 
        }

        return blnSuccess;
    }

    function cmdEditCancel_onclick(Type) {
        var sDialogView = globalVariable.get('product_prop_view');
        if (isFromPulsar2()) {
            closePulsar2Popup(false);
        }
        else if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else 
        {
            if (Type==1) //close the jquery pop up when editing or cloning
            {
                if (CheckOpener() === false) {
                    parent.window.parent.ClosePropertiesDialog();
                } else {
                    window.parent.close();
                }
            } else { //close the jquery pop up when adding new product
                if (CheckOpener() === false && sDialogView == 'add') {
                    //the ClosePropertiesDialog is initiated from leftmenu's Add New link
                    parent.parent.window.parent.ClosePropertiesDialog();
                }else{
                    window.parent.close();
                }
            }
        }
    }

    function CheckOpener(){
        //If True, page opened with showModalDialog
        //if False, page opened with JQuery Modal Dialog
        var oWindow = window.dialogArguments;
        return (oWindow == null)?false:true;
    }

    function cmdClear_onclick() {
        window.parent.frames["UpperWindow"].ProgramInput.reset();
        window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.fontStyle = "italic";
        window.parent.frames["UpperWindow"].ProgramInput.txtJustification.style.color="blue";
    }

    function cmdSubmit_onclick() {
        var i;
        var strOutput="", strBrandNamesWOFormula="";
        var bSaveData = true;
       
        if (window.location.hostname.indexOf("/pulsarplus/") < 0) {

            if (VerifySave())
            {
                if (followMKTName == 0) {
                     //Build List of Brands
                    for(i=0;i<window.parent.frames["UpperWindow"].ProgramInput.chkBrands.length;i++)
                    {
                        if (window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].checked)
                        {
                            if (strOutput=="")
                                strOutput =  $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('brandname');
                            else
                                strOutput = strOutput + ", " + $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('brandname');
                            if ($('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('bnwoformula')!= "")
                            {
                                strBrandNamesWOFormula = strBrandNamesWOFormula == "" ? $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('brandname') + " :\n\t" + $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('bnwoformula').substring(0, $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('bnwoformula').length -1 ).replace(/,/g,"\n\t")
                                    : strBrandNamesWOFormula + "\n" + $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('brandname') + " :\n\t" +  $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('bnwoformula').substring(0, $('#chkBrands', window.parent.frames["UpperWindow"].ProgramInput).get(i).getAttribute('bnwoformula').length -1 ).replace(/,/g,"\n\t");                                                         
                            }
                        }
                    }    
                }                      
                if (strBrandNamesWOFormula != "" &&  !window.parent.frames["UpperWindow"].ProgramInput.chkEnableFollowMarketingName.checked ) 
                {
                    bSaveData = confirm("Brands listed below do not have all formulas created in Brand Admin. The Brand Name will be used in the names with missing formulas.\n\nThe brands with the missing formulas are:\n" + strBrandNamesWOFormula + "\n\nClick 'OK' to continue or 'Cancel' to go back to the previous screen.");
                }
                if (bSaveData)
                {
                    window.parent.frames["UpperWindow"].ProgramInput.txtBrands.value = strOutput;
	
                    //Submit form
                    if (document.getElementById("preferredLayout").value == "" && document.getElementById("preferredLayout").value != 'pulsar2') {
                        cmdEditCancel.disabled = true;
                    }
                    cmdSubmit.disabled =true;
                    cmdClear.disabled =true;
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value=="")
                        window.parent.frames["UpperWindow"].ProgramInput.txtProductName.value = window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value;
                    else
                        window.parent.frames["UpperWindow"].ProgramInput.txtProductName.value = window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value + " " + window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value;

                    window.parent.frames["UpperWindow"].ProgramInput.submit();
                }
            }
        }
        else
        {
  
            if (VerifySave())
            {
                if (followMKTName == 0) {
                    //Build List of Brands
                    for (i = 0; i < window.parent.frames["UpperWindow"].ProgramInput.chkBrands.length; i++) {
                        if (window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].checked) {
                            if (strOutput == "")
                                strOutput = window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BrandName;
                            else
                                if (window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BNWOFormula != "") {
                                    strBrandNamesWOFormula = strBrandNamesWOFormula == "" ? window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BrandName + " :\n\t" + window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BNWOFormula.substring(0, window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BNWOFormula.length - 1).replace(/,/g, "\n\t")
                                        : strBrandNamesWOFormula + "\n" + window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BrandName + " :\n\t" + window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BNWOFormula.substring(0, window.parent.frames["UpperWindow"].ProgramInput.chkBrands[i].BNWOFormula.length - 1).replace(/,/g, "\n\t");
                                }
                        }

                    }
                }
                if (strBrandNamesWOFormula != "" &&  !window.parent.frames["UpperWindow"].ProgramInput.chkEnableFollowMarketingName.checked)
                {
                    bSaveData = confirm("Brands listed below do not have all formulas created in Brand Admin. The Brand Name will be used in the names with missing formulas.\n\nThe brands with the missing formulas are:\n" + strBrandNamesWOFormula + "\n\nClick 'OK' to continue or 'Cancel' to go back to the previous screen.");
                }
                if (bSaveData)
                {
                    window.parent.frames["UpperWindow"].ProgramInput.txtBrands.value = strOutput;
	
                    //Submit form
                    cmdEditCancel.disabled =true;
                    cmdSubmit.disabled =true;
                    cmdClear.disabled =true;
                    if (window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value=="")
                        window.parent.frames["UpperWindow"].ProgramInput.txtProductName.value = window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value;
                    else
                        window.parent.frames["UpperWindow"].ProgramInput.txtProductName.value = window.parent.frames["UpperWindow"].ProgramInput.txtProductNameBase.value + " " + window.parent.frames["UpperWindow"].ProgramInput.txtVersion.value;

                    window.parent.frames["UpperWindow"].ProgramInput.submit();
                }
            }
        }
     
    }

    function cmdOK_onclick(strFunction) {

        var blnAll = true;
        var i;
        var blnFailed= false;
        if (strFunction == 1)
        {
            if (window.parent.frames["UpperWindow"].ProgramInput.txtServiceEndDate.value!="")
            {
                if (!isDate(window.parent.frames["UpperWindow"].ProgramInput.txtServiceEndDate.value))
                {
                    alert("The End of Service Life must be a date if it is entered.");
                    blnFailed =true;
                }
            }
            else if (window.parent.frames["UpperWindow"].ProgramInput.cboWWAN.selectedIndex==0)
            {
                alert("You must specify if this is a WWAN product.");
                blnFailed =true;
            }
        }
        if (!blnFailed)
        {			
            cmdEditCancel.disabled =true;
            cmdSubmit.disabled =true;
            window.parent.frames["UpperWindow"].ProgramInput.submit();
        }
        if (parent.window.parent.loadDatatodiv != undefined) {
            parent.window.parent.closeExternalPopup();
        } 

    }



    //-->
</script>
</head>
<body style="border-top:2px solid #b2b2b2">
<div style="text-align:right;">
		<%if request("ID") <> "" then%>
			<%if trim(request("Commodity")) = "1" then %>
				<INPUT type="button" value="OK" id=cmdSubmit name=cmdSubmit LANGUAGE=javascript onclick="return cmdOK_onclick(1)">
			<%elseif trim(request("FactoryEngineer")) = "1" or trim(request("Accessory")) = "1" then %>
				<INPUT type="button" value="OK" id=Button1 name=cmdSubmit LANGUAGE=javascript onclick="return cmdOK_onclick(2)">
			<%elseif trim(request("HWPM")) = "1" then %>
				<INPUT type="button" value="OK" id=Button2 name=cmdSubmit LANGUAGE=javascript onclick="return cmdOK_onclick(3)">
			<%else%>
				<INPUT type="button" value="OK" id=cmdSubmit name=cmdSubmit LANGUAGE=javascript onclick="return cmdSubmit_onclick()">
			<%end if%>
			<INPUT style="Display:none" type="button" value="Clear Form" id=cmdClear name=cmdClear  LANGUAGE=javascript onclick="return cmdClear_onclick()">
			<INPUT type="button" value="Cancel" id=cmdEditCancel name=cmdEditCancel  LANGUAGE=javascript onclick="return cmdEditCancel_onclick(1)"  >
		<%else%>
			<INPUT type="button" value="OK" id=cmdSubmit name=cmdSubmit LANGUAGE=javascript onclick="return cmdSubmit_onclick()">
			<INPUT type="button" value="Clear Form" id=cmdClear name=cmdClear  LANGUAGE=javascript onclick="return cmdClear_onclick()">
         <%if Request.Cookies("PreferredLayout2") <> "pulsar2" then%>
          <INPUT type="button" value="Cancel" id=cmdEditCancel name=cmdEditCancel  LANGUAGE=javascript onclick="return cmdEditCancel_onclick(2)"  >
         <%end if%>
		<%end if%>
        <input type="hidden" id="preferredLayout" value="<%=Request.Cookies("PreferredLayout2")%>" />
</div>
</body>
</html>