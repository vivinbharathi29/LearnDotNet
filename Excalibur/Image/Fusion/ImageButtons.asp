<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script ID="clientEventHandlersVBS" LANGUAGE="jscript">
<!--  
<!-- #include file = "../../includes/Date.asp" -->

    
function ValidateTab(strState){
	var blnSuccess;
	var blnFound;
	var FocusOn;
	var i;
	var BrandError="";

	blnSuccess = true;
	switch (strState)
	{
	    case "General":
            //Check Brands
	        if (typeof(window.parent.frames["UpperWindow"].AddImage.chkBrands) == "undefined")
	            BrandError = "No Brands specified for this product.";
	        else if (typeof (window.parent.frames["UpperWindow"].AddImage.chkBrands.length) == "undefined") {
	            if (!window.parent.frames["UpperWindow"].AddImage.chkBrands.checked)
	                BrandError = "You must select at least one Brand.";
            }
	        else {
	            BrandError = "You must select at least one Brand.";
	            for (i=0;i<window.parent.frames["UpperWindow"].AddImage.chkBrands.length;i++)
	                if (window.parent.frames["UpperWindow"].AddImage.chkBrands[i].checked)
	                    BrandError = "";
	        }

	        if(typeof(window.parent.frames["UpperWindow"].AddImage.cboRelease) != "undefined"  && (window.parent.frames["UpperWindow"].AddImage.cboRelease) != "undefined" &&((window.parent.frames["UpperWindow"].AddImage.cboRelease.value == "")||(window.parent.frames["UpperWindow"].AddImage.cboRelease.value == 0)))
	        {
	            BrandError = "Please select Release for Image.";
	        }

			//if (window.parent.frames["UpperWindow"].AddImage.txtDCRRequired.value == "")
			//	{
			//	if (window.parent.frames["UpperWindow"].AddImage.cboDCR.selectedIndex == 0)
			//		{
			//		window.parent.frames["UpperWindow"].CurrentState = "General";
			//		window.parent.frames["UpperWindow"].ProcessState();
			//		window.parent.frames["UpperWindow"].AddImage.cboDCR.focus();
			//		if (window.parent.frames["UpperWindow"].AddImage.txtDisplayedID.value != "")
			//			window.alert("Select an approved change request to update this image definition.");
			//		else
			//			window.alert("Select an approved change request to add this image definition.");
			//		blnSuccess = false;
			//		}
			//	}
			if (window.parent.frames["UpperWindow"].divImagesValidated.style.display =="" && ! window.parent.frames["UpperWindow"].AddImage.chkImagesValidated.checked)
				{
				window.parent.frames["UpperWindow"].CurrentState = "General";
				window.parent.frames["UpperWindow"].ProcessState();
				window.parent.frames["UpperWindow"].AddImage.chkImagesValidated.focus();
				window.alert("You must verify the images in Conveyor before you can release to factory.");
				blnSuccess = false;
				}			
			if (BrandError != "" && blnSuccess)
				{ 
				window.parent.frames["UpperWindow"].CurrentState = "General";
				window.parent.frames["UpperWindow"].ProcessState();
				window.alert(BrandError);
				blnSuccess = false;
				}
                
	        //Harris, Valerie (3/16/2016) - PBI 17835/ Task 18059  - validate selected index is greater than "0" for Operating System only when Copy With Targeting isn't selected
			if(typeof(window.parent.frames["UpperWindow"].AddImage.inpCopyWithTarget) != "undefined"){
			    if (window.parent.frames["UpperWindow"].AddImage.inpCopyWithTarget.value == "False" && window.parent.frames["UpperWindow"].AddImage.cboOS.selectedIndex == 0  && blnSuccess)
				    {
				    window.parent.frames["UpperWindow"].CurrentState = "General";
				    window.parent.frames["UpperWindow"].ProcessState();
				    window.parent.frames["UpperWindow"].AddImage.cboOS.focus();
				    window.alert("Operating System is required.");
				    blnSuccess = false;
				}
            }else{
                if (window.parent.frames["UpperWindow"].AddImage.cboOS.selectedIndex == 0  && blnSuccess)
			    {
			        window.parent.frames["UpperWindow"].CurrentState = "General";
			        window.parent.frames["UpperWindow"].ProcessState();
			        window.parent.frames["UpperWindow"].AddImage.cboOS.focus();
			        window.alert("Operating System is required.");
			        blnSuccess = false;
			    }
            }


			if (window.parent.frames["UpperWindow"].AddImage.cboStatus.selectedIndex == 0  && blnSuccess)
				{
				window.parent.frames["UpperWindow"].CurrentState = "General";
				window.parent.frames["UpperWindow"].ProcessState();
				window.parent.frames["UpperWindow"].AddImage.cboStatus.focus();
				window.alert("Image Status is required.");
				blnSuccess = false;
				}

			if (window.parent.frames["UpperWindow"].AddImage.txtRTMDate.value != "" && ! isDate(window.parent.frames["UpperWindow"].AddImage.txtRTMDate.value)  && blnSuccess)
				{
					window.parent.frames["UpperWindow"].CurrentState = "General";
					window.parent.frames["UpperWindow"].ProcessState();
					window.parent.frames["UpperWindow"].AddImage.txtRTMDate.focus();
			
					window.alert("RTM Date must be a valid date if supplied.");
					blnSuccess = false;					
				}


			if ((String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("]") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("*") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("\"") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("/") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("\\") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("?") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf(":") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("|") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("<") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf(">") != -1 || String(window.parent.frames["UpperWindow"].AddImage.txtProductDrop.value).indexOf("[") != -1) && blnSuccess)
				{
					window.parent.frames["UpperWindow"].CurrentState = "General";
					window.parent.frames["UpperWindow"].ProcessState();
					window.parent.frames["UpperWindow"].AddImage.txtProductDrop.focus();
			
					window.alert("Product Drop can not contain any of the following characters:\r\r\\  /  |  [  ]  *  :  ?  \"  <  >");
					blnSuccess = false;					
				}

		break;
	}
	return blnSuccess;
}

function instrAt(MyString,Find){
	return MyString.indexOf(Find,0);
	
}

function cmdNext_onclick() {
	var i;
	
	if (ValidateTab(window.parent.frames["UpperWindow"].CurrentState) == true)
		{
			
			switch (window.parent.frames["UpperWindow"].CurrentState)
			{
				case "General":
					window.parent.frames["UpperWindow"].CurrentState = "Regions";
				break;	
				case "Regions":
					window.parent.frames["UpperWindow"].BuildPreview();
					window.parent.frames["UpperWindow"].CurrentState = "Preview";
				break;
			}
		
			window.parent.frames["UpperWindow"].ProcessState();
		}
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (window.confirm ("Are you sure you want to exit this screen without saving your changes?") == true)
    {
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") 
        {
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
        }
        else
        {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }
}

function cmdPrevious_onclick() {
	switch (window.parent.frames["UpperWindow"].CurrentState)
	{
		case "Preview":
			window.parent.frames["UpperWindow"].CurrentState = "Regions";
		break;
		case "Regions":
			window.parent.frames["UpperWindow"].CurrentState = "General";
		break;
	}
	window.parent.frames["UpperWindow"].ProcessState();

}

function window_onload() {
	frameloaded=true
}

function cmdFinish_onclick() {
	var i;
	cmdFinish.disabled =true;
	cmdNext.disabled =true;
	cmdPrevious.disabled =true;
	cmdCancel.disabled =true;
	window.parent.frames["UpperWindow"].AddImage.submit();
}
-->
</script>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {

	var ValidationFailed = false;
	
	if (window.parent.frames["UpperWindow"].CellGeneral.style.display== ""  || window.parent.frames["UpperWindow"].CellGeneralb.style.display=="")
		{
		ValidationFailed = ! ValidateTab("General");
		}
	if (! ValidationFailed)
		{
		cmdFinish.disabled =true;
		cmdNext.disabled =true;
		cmdPrevious.disabled =true;
		cmdOK.disabled =true;
		cmdCancel.disabled =true;
		cmdEditCancel.disabled =true;
		//if (window.parent.frames["UpperWindow"].AddImage.txtDCRRequired.value == "none")
			window.parent.frames["UpperWindow"].AddImage.submit();
		//else if (window.parent.frames["UpperWindow"].AddImage.cboDCR.selectedIndex > 0)
		//	window.parent.frames["UpperWindow"].AddImage.submit();
		//else
		//    if (window.parent.frames["UpperWindow"]) {
		//        parent.window.parent.modalDialog.cancel(true);
		//    } else {
		//        window.parent.close();
		//    }
		}
}

function cmdEditCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") 
    {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else
    {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}


//-->
</SCRIPT>
</head>
<body bgcolor="ivory" LANGUAGE="javascript" onload="return window_onload()">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<!--<TD width=100%><font face=verdana size=1 color=blue>Select any field and Press F1 for more information</TD>-->
		<%if request("ID") <> "" then%>
			<TD><INPUT type="button" value="Save" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
			<TD><INPUT type="button" value="Cancel" id=cmdEditCancel name=cmdEditCancel  LANGUAGE=javascript onclick="return cmdEditCancel_onclick('<%=Request("pulsarplusDivId")%>')"></TD>
</TR></table><table><tr>
			<td><input style="Display:none" type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></td>
			<td width="10"></td>
			<td><input style="Display:none" type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input style="Display:none" type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input style="Display:none" type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>

		<%else%>
			<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></td> <!-- style="BORDER-LEFT-COLOR: tan; BORDER-BOTTOM-COLOR: tan; BORDER-TOP-COLOR: tan; BACKGROUND-COLOR: wheat; BORDER-RIGHT-COLOR: tan-->
			<td width="10"></td>
			<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>
		<%end if%>
	</tr>
</table>
</body>
</html>