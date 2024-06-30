<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script src="/excalibur/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="/excalibur/includes/client/jquery-ui.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

    function cmdCancel_onclick() {
        try {
            if ('<%=Request.QueryString("pulsarplusDivId")%>' != '') {
                parent.window.parent.closeExternalPopup();
            }
            else {
                window.parent.ClosePlatFormDetail("Dialog1", false);
            }
        }
        catch (e) {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.returnValue = 0;
                window.close();
            }
        }
}

function cmdOK_onclick() {
    var ID = window.parent.frames["UpperWindow"].frmMain.txtProductionVersionID.value;
    var validHEXChars = "012345678ABCDEFabcdef"
    var introYear = window.parent.frames["UpperWindow"].frmMain.txtIntroYear.value
    var introYearLength = introYear.length;
    var isNotaNumber = isNaN(introYear);

    var productFamily = window.parent.frames["UpperWindow"].frmMain.txtProductFamily.value;
    var pcaName = window.parent.frames["UpperWindow"].frmMain.txtPCAName.value;
    var chassis = window.parent.frames["UpperWindow"].frmMain.ddlChassis.value;
    var brand = window.parent.frames["UpperWindow"].frmMain.ddlBrand.value;
    var genericName = window.parent.frames["UpperWindow"].frmMain.txtGenericName.value;
    var marketingName = window.parent.frames["UpperWindow"].frmMain.txtMarketingName.value;
    var fullMarketingName = window.parent.frames["UpperWindow"].frmMain.txtFullMarketingName.value;
    var systemID = window.parent.frames["UpperWindow"].frmMain.txtSystemID.value;
    //'Yong removed the code (as business function moved to PRL) instead of comment it out to make the page clean; 
    //if for some reason we need to put this back, we just need to look at tfs history
    var ModelNumber = window.parent.frames["UpperWindow"].frmMain.txtProductModelNumber.value;
    var AllowedCharacters = /^[[0-9a-zA-Z \-]+$/;
    var MKTAllowedChars = /^[0-9a-zA-Z \- \.]+$/; 
    var screenSizeAllowedChars = /^[0-9\.]+$/; 
    var followMKTName = parseInt('<%=Request.QueryString("FollowMKTName")%>');
    var MKTBrand, showDCRDiv,  tagMKTName, curMKTName, curPhwFamilyName,strBrandNamesWOFormula="", bSaveData
    var SMBIOSFamily = window.parent.frames["UpperWindow"].frmMain.txtBFName.value;
    
    if (followMKTName == 1){
        MKTBrand = window.parent.frames["UpperWindow"].frmMain.txtBrandsLoaded.value;
        showDCRDiv = window.parent.frames["UpperWindow"].frmMain.showDCRDiv.value; 
        tagMKTName = window.parent.frames["UpperWindow"].frmMain.tagMktName.value;
        curMKTName = window.parent.frames["UpperWindow"].frmMain.hidMktName.value;
        curPhwFamilyName = window.parent.frames["UpperWindow"].frmMain.hidphwFamilyName.value;
        curPMS = window.parent.frames["UpperWindow"].frmMain.hidPMS.value;
        curSentPMSReq = window.parent.frames["UpperWindow"].frmMain.hidSentPMSReq.value;
        curScreenSize = window.parent.frames["UpperWindow"].document.getElementById("txtScreenSize" + MKTBrand).value;

        if (MKTBrand > 0) { 
            if ($('#chkBrands' + MKTBrand , window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('bnwoformula')!= "")
            {
                strBrandNamesWOFormula = strBrandNamesWOFormula == "" ? $('#chkBrands' + MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('brandname') + " :\n\t" + $('#chkBrands' + MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('bnwoformula').substring(0, $('#chkBrands'+ MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('bnwoformula').length -1 ).replace(/,/g,"\n\t")
                    : strBrandNamesWOFormula + "\n" + $('#chkBrands' + MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('brandname') + " :\n\t" +  $('#chkBrands'+ MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('bnwoformula').substring(0, $('#chkBrands'+ MKTBrand, window.parent.frames["UpperWindow"].frmMain).get(0).getAttribute('bnwoformula').length -1 ).replace(/,/g,"\n\t");                                                         
            }
        }
    }


  if (introYear == "") {
      alert("You must enter intro year.");
      return;
    }
  if (isNotaNumber == true) {
        alert("You enter an invalid intro year.");
        return;
    }
  else {
      if (introYearLength < 4) {
          alert("Intro year must have four digits.");
          return;
      }
      else {
          if (introYear < 2000) {
              alert("Intro year must be greater than 2000.");
              return;
          }
      }
  }


   if (productFamily == "") {
        alert("You must enter product family name.");
        return;
    }
   else if (pcaName == "") {
        alert("You must enter PCA name.");
        return;
   }
   else if (ModelNumber != "" && !(ModelNumber.match(AllowedCharacters)))
   {
        alert("Model Number can only contain alphanumeric and dashes");
        return;
   }
   else if (chassis == "") {
        alert("You must select a chassis.");
        return;
    }
   else if (brand == "" && followMKTName == 0) {
        alert("You must select a brand.");
        return;
    }
   else if (genericName == "" && followMKTName == 0) { 
       alert("You must enter generic name.");
       return;
   }
   else if (marketingName == "" && followMKTName == 0) { 
       alert("You must enter marketing name.");
       return;
   }
   else if (followMKTName == 1 && MKTBrand == "") {
       alert("You must select a brand.");
       return;
   }
   else if (followMKTName == 1 && (curMKTName=="")) {
       alert("Marketing Name can't be blank. Please change the settings for Brand.") 
        return;
   }
   else if (followMKTName == 1 && (curMKTName.length > 60)) {
       alert("Marketing Name cannot exceed 60 characters. Please contact the platform CM / SCM owner to correct the Branding formula");
        return;
   }
   else if (followMKTName == 1 && (curPhwFamilyName=="")) {
       alert("PHWeb Family Name can't be blank. Please change the settings for Brand.") 
        return;
   }
   else if (followMKTName == 1 && showDCRDiv == "" && window.parent.frames["UpperWindow"].frmMain.cboDCR.value =="" && tagMKTName != curMKTName) { 
        alert("You must select a DCR.");
        return;
   }
   else if (SMBIOSFamily !="" && SMBIOSFamily.indexOf("103C_")== -1) {
        alert("SMBIOS Family Name Override filed must be a string that starts with '103C_'.");
        return;
   }
   else if (followMKTName == 1 && curPMS == "No Match" && curSentPMSReq == 0 && (showDCRDiv=="" || window.parent.frames["UpperWindow"].frmMain.hidPdStatus.value >1 )) {
       alert("Marketing name does not match Product Master. Please click the button 'Request New Product Master Series'");
       return;
   }
   else if (followMKTName == 1 && (curMKTName != "" && !curMKTName.match(MKTAllowedChars))) {
        alert("Marketing Name can only use A-Z, a-z, space, hyphan and period.") 
        return;
   }
   else if (followMKTName == 1 && curScreenSize !="" && !curScreenSize.match(screenSizeAllowedChars)) {
        alert("Screen Size can only contain numeric values.");
        return;
   }
   else if (followMKTName == 1 && strBrandNamesWOFormula != "") {
       bSaveData = confirm("Brands listed below do not have all formulas created in Brand Admin. The Brand Name will be used in the names with missing formulas.\n\nThe brands with the missing formulas are:\n" + strBrandNamesWOFormula + "\n\nClick 'OK' to continue or 'Cancel' to go back to the previous screen.");
       if (!bSaveData)
           return;
       else {
           window.parent.frames["UpperWindow"].frmMain.submit();
           return true;
       }
   }
   else {
       window.parent.frames["UpperWindow"].frmMain.submit();
       return true;
   }
}
</SCRIPT>
</head>
<body bgcolor="ivory">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
        <td><font color="green" size="2">After adding a Base Unit Group, select Refresh on the Product Properties page if necessary to see the change.</font></td>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>