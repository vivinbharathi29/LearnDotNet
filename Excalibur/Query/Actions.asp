<%@ Language=VBScript %>

	<%
	
 ' Response.Buffer = True
  'Response.ExpiresAbsolute = Now() - 1
  'Response.Expires = 0
  'Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<script language="JavaScript" src="../_ScriptLibrary/jsrsClient.js"></script>

<%if request("CAT") = "1" then%>
	<title>SKU Change Requests - Advanced Query</title>
<%else%>
	<title>Action Item - Advanced Query</title>
<%end if%>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
<!-- #include file = "../includes/Date.asp" -->

  function isNumeric(value) 
  {
    if (value == "")  { return false }
      start = 0;
    for (i=start; i<value.length; i++)
    {
      if (value.charAt(i) < "0") { return false; }
      if (value.charAt(i) > "9") { return false; }
    }
    return true;
  }

function VerifyFields(strType){
	var blnOK = true;
	var i;
	var blnSelected = false;
	
	
	if (ProgramInput.txtNumbers.value != "")
		{
		if (! isNumericWithCommas(ProgramInput.txtNumbers.value))
			{
			blnOK = false;
			window.alert("You can only enter a comma-seperated list of numbers into the ID field.");
			ProgramInput.txtNumbers.focus();
			}
		}
	
	
	
/*	
	if (strType == 1)
		{
			if (ProgramInput.txtSubject.value == "")
				{
					blnOK = false;
					window.alert("You must enter a subject when sending mail.");
					ProgramInput.txtSubject.focus();
				}
			else if(ProgramInput.txtNotes.value == "")
				{
					blnOK = false;
					window.alert("You must enter notes when sending mail.");
					ProgramInput.txtNotes.focus();
				}
		}

	if (blnOK)
		{
			if (ProgramInput.txtDaysInState.value == "")
				{
					blnOK = false;
					window.alert("Days In State is required.  Enter > 0 for all matching observations.");
					ProgramInput.txtDaysInState.focus();
				}
			else if (! isNumeric(ProgramInput.txtDaysInState.value))
				{
					blnOK = false;
					window.alert("Days In State must be a positive number.");
					ProgramInput.txtDaysInState.focus();
				}
		}
	if (blnOK)
		{
			blnSelected = false;
			for (i=0;i<ProgramInput.lstProducts.length;i++)
				{
					if (ProgramInput.lstProducts.options(i).selected)
						blnSelected = true;
				}
			if (! blnSelected)
				{
					if (ProgramInput.txtNumbers.value == "")
						{
						blnOK = false;
						window.alert("You must select at least one product or enter at least one Observation Number.");
						ProgramInput.lstStates.focus();
						}
				}		
			
		}
		
*/
	return blnOK;
	
}

  function isNumericWithCommas(value) 
  {
    if (value == "")  { return false }
    start = 0;
    var FoundComma = false;
    for (i=start; i<value.length; i++)
    {
	  if (value.charAt(i) == ",")
	   {
		if (FoundComma)
			return false;
		else
			FoundComma = true;
	   }
	  else if (value.charAt(i) != " ")
		{
		FoundComma=false;
		if (value.charAt(i) < "0") { return false; }
		if (value.charAt(i) > "9") { return false; }
		}
    }
    return true;
  }


function cmdReport_onclick() {
	if (VerifyFields(1))
		{
		ProgramInput.txtFunction.value = "1"		
		ProgramInput.submit();
		}
}


function VerifyDateFields(){
	var blnOK = true;
	if (ProgramInput.cboDaysOpenCompare.selectedIndex == 3)
		{
			if (ProgramInput.txtOpenRange1.value != "")
				{
					if (! isDate(ProgramInput.txtOpenRange1.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtOpenRange1.focus();
						}				
				}
			if (ProgramInput.txtOpenRange2.value != "" && blnOK)
				{
					if (! isDate(ProgramInput.txtOpenRange2.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtOpenRange2.focus();
						}				
				}
				
		}
	else
		{
		if (ProgramInput.txtDaysOpen.value == "")
			{
				blnOK = false;
				window.alert("Date Open is required.  Enter >= 0 for all matching observations.");
				ProgramInput.txtDaysOpen.focus();
			}
		else if (! isNumeric(ProgramInput.txtDaysOpen.value))
			{
				blnOK = false;
				window.alert("Date Open must be a positive number.");
				ProgramInput.txtDaysOpen.focus();
			}
		}
		
		
		
	if (blnOK)
		{
		if (ProgramInput.cboDaysClosedCompare.selectedIndex == 3)
		{
			if (ProgramInput.txtClosedRange1.value != "")
				{
					if (! isDate(ProgramInput.txtClosedRange1.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtClosedRange1.focus();
						}				
				}
			if (ProgramInput.txtClosedRange2.value != "" && blnOK)
				{
					if (! isDate(ProgramInput.txtClosedRange2.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtClosedRange2.focus();
						}				
				}
				
		}
		else
		{
		if (ProgramInput.txtDaysClosed.value == "")
			{
				blnOK = false;
				window.alert("Date Closed is required.  Enter >= 0 for all matching observations.");
				ProgramInput.txtDaysClosed.focus();
			}
		else if (! isNumeric(ProgramInput.txtDaysClosed.value))
			{
				blnOK = false;
				window.alert("Date Closed must be a positive number.");
				ProgramInput.txtDaysClosed.focus();
			}
		}
		
		
		}


	if (blnOK) //Target Date
		{
		if (ProgramInput.cboDaysTargetCompare.selectedIndex == 3)
		{
			if (ProgramInput.txtTargetRange1.value != "")
				{
					if (! isDate(ProgramInput.txtTargetRange1.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtTargetRange1.focus();
						}				
				}
			if (ProgramInput.txtTargetRange2.value != "" && blnOK)
				{
					if (! isDate(ProgramInput.txtTargetRange2.value))
						{
						blnOK = false;
						window.alert("Dates in Range must be a valid date format.");
						ProgramInput.txtTargetRange2.focus();
						}				
				}
				
		}
		else
		{
		if (ProgramInput.txtDaysTarget.value == "")
			{
				blnOK = false;
				window.alert("Target Date is required.  Enter >= 0 for all matching observations.");
				ProgramInput.txtDaysTarget.focus();
			}
		else if (! isNumeric(ProgramInput.txtDaysTarget.value))
			{
				blnOK = false;
				window.alert("Target Date must be a positive number.");
				ProgramInput.txtDaysTarget.focus();
			}
		}
		
		
		}
	
	return blnOK;		
		
}

function cmdDetails_onclick() {
	//window.alert("Not implemented yet.  Check back soon.");
	if (VerifyFields(2))
		{
		ProgramInput.txtFunction.value = "2"		
		ProgramInput.submit();
		}
}


function cmdReset_onclick() {
	spnOpenCount.style.display="";
	spnOpenRange.style.display="none";
	spnClosedCount.style.display="";
	spnClosedRange.style.display="none";
	spnTargetCount.style.display="";
	spnTargetRange.style.display="none";
	ProgramInput.reset();
}


function window_onload() {
	lblLoad.style.display = "none";
}

function cboDaysOpenCompare_onchange() {
	if (ProgramInput.cboDaysOpenCompare.selectedIndex ==3)
		{
		spnOpenCount.style.display="none";
		spnOpenRange.style.display="";
		}
	else
		{
		spnOpenCount.style.display="";
		spnOpenRange.style.display="none";
		}
}

function cboDaysClosedCompare_onchange() {
	if (ProgramInput.cboDaysClosedCompare.selectedIndex ==3)
		{
		spnClosedCount.style.display="none";
		spnClosedRange.style.display="";
		}
	else
		{
		spnClosedCount.style.display="";
		spnClosedRange.style.display="none";
		}
}

function cboDaysTargetCompare_onchange() {
	if (ProgramInput.cboDaysTargetCompare.selectedIndex ==3)
		{
		spnTargetCount.style.display="none";
		spnTargetRange.style.display="";
		}
	else
		{
		spnTargetCount.style.display="";
		spnTargetRange.style.display="none";
		}
}

function PickDate(intControl){

	var strDate;
	if (intControl==1)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtOpenRange1.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtOpenRange1.value = strDate;
			}
		}
	else if (intControl==2)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtOpenRange2.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtOpenRange2.value = strDate;
			}
		}
	else if (intControl==3)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtClosedRange1.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtClosedRange1.value = strDate;
			}
		}
	else if (intControl==4)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtClosedRange2.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtClosedRange2.value = strDate;
			}
		}
	else if (intControl==5)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtTargetRange1.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtTargetRange1.value = strDate;
			}
		}
	else if (intControl==6)
		{
		strDate = window.showModalDialog("../MobileSE/Today/calDraw1.asp",ProgramInput.txtTargetRange2.value,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strDate) != "undefined")
			{
			ProgramInput.txtTargetRange2.value = strDate;
			}
		}


}


function instr(MyString,Find){
	return MyString.indexOf(Find,0);
	
}


function cboProfile_onchange() {
	var strColumns;
	var strProducts;
	var strBuffer;
	var i;
	var strHeader;
		
	if (ProgramInput.cboProfile.selectedIndex > 1)
		{
			if (ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).CanEdit == "True")
				{
				ProfileOptionsUpdate.style.display = "";
				ProfileOptionsRename.style.display = "";		
				}
			else
				{
				ProfileOptionsUpdate.style.display = "none";
				ProfileOptionsRename.style.display = "none";		
				}

			if (ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).CanDelete == "True")
				ProfileOptionsDelete.style.display = "";		
			else
				ProfileOptionsDelete.style.display = "none";		

			if (ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).PrimaryOwner == "")
				{
				ProfileOptionsRemove.style.display = "none";		
				ProfileOptionsOwner.style.display = "none";		
				ProfileOptionsShare.style.display = "";		
				}
			else
				{
				if (ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).CanRemove=="0")
				    ProfileOptionsRemove.style.display = "none";
                else
				    ProfileOptionsRemove.style.display = "";
				ProfileOptionsShare.style.display = "none";		
				ProfileOptionsOwner.innerHTML = "<font size=1 face=verdana color=black><b>Profile Owner:</b> " + ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).PrimaryOwner + "</font>";		
				ProfileOptionsOwner.style.display = "";		
				}


			jsrsExecute("ActionsRSget.asp", myCallback1, "ProfileStrings", ProgramInput.cboProfile.value);			
			
		}
	else
		{
			ProfileOptionsUpdate.style.display = "none";		
			ProfileOptionsDelete.style.display = "none";		
			ProfileOptionsRename.style.display = "none";		
			ProfileOptionsOwner.style.display = "none";	
			ProfileOptionsRemove.style.display = "none";
			ProfileOptionsShare.style.display = "none";		
		
				
		}
}


function myCallback1( returnstring ){
	OutputArea.innerHTML = returnstring;
	if (txtQProfileName.value == "")
		{
			window.alert("Unable to read the selected profile.");			
		}			
	else
		{
			//if (txtQValue1.value != "")
			//{
			for (i=0;i<ProgramInput.lstProducts.length;i++)
				{
				if (instr("," + txtQValue1.value + "," , "," + ProgramInput.lstProducts.options[i].value + "," ) > -1)
					ProgramInput.lstProducts.options[i].selected = true;
				else
					ProgramInput.lstProducts.options[i].selected = false;
			}

			for (i=0;i<ProgramInput.lstProductsPulsar.length;i++)
			{
			    if (instr("," + txtQValue55.value + "," , "," + ProgramInput.lstProductsPulsar.options[i].value + "," ) > -1)
			        ProgramInput.lstProductsPulsar.options[i].selected = true;
			    else
			        ProgramInput.lstProductsPulsar.options[i].selected = false;
			}
			//}				
			//if (txtQValue2.value != "")
			//{
			for (i=0;i<ProgramInput.lstOwners.length;i++)
				{
				if (instr("," + txtQValue2.value + "," , "," + ProgramInput.lstOwners.options[i].value + "," ) > -1)
					ProgramInput.lstOwners.options[i].selected = true;
				else
					ProgramInput.lstOwners.options[i].selected = false;
				}
			//}				
		
			//if (txtQValue3.value != "")
			//{
			for (i=0;i<ProgramInput.lstApprovers.length;i++)
				{
				if (instr("," + txtQValue3.value + "," , "," + ProgramInput.lstApprovers.options[i].value + "," ) > -1)
					ProgramInput.lstApprovers.options[i].selected = true;
				else
					ProgramInput.lstApprovers.options[i].selected = false;
				}
			for (i=0;i<ProgramInput.lstSubmitter.length;i++)
				{
				if (instr("," + txtQValue26.value + "," , "," + ProgramInput.lstSubmitter.options[i].value + "," ) > -1)
					ProgramInput.lstSubmitter.options[i].selected = true;
				else
					ProgramInput.lstSubmitter.options[i].selected = false;
				}
			//}				


			if (txtQValue5.value != "")
				{
					ProgramInput.cboDaysOpenCompare.selectedIndex = txtQValue5.value;
				}				
				if (txtQValue7.value != "")
				{
					ProgramInput.cboApproverStatus.selectedIndex = txtQValue7.value;
				}				
			
			ProgramInput.txtDaysOpen.value = txtQValue6.value;

			if (txtQValue8.value != "")
				{
					ProgramInput.cboCategory.selectedIndex = txtQValue8.value;
				}				

			for (i=0;i<ProgramInput.lstStatus.length;i++)
				{
				if (instr("," + txtQValue9.value + "," , "," + ProgramInput.lstStatus.options[i].value + "," ) > -1)
					ProgramInput.lstStatus.options[i].selected = true;
					else
						ProgramInput.lstStatus.options[i].selected = false;
				}


			for (i=0;i<ProgramInput.lstType.length;i++)
				{
				if (instr("," + txtQValue10.value + "," , "," + ProgramInput.lstType.options[i].value + "," ) > -1)
					ProgramInput.lstType.options[i].selected = true;
				else
					ProgramInput.lstType.options[i].selected = false;
				}

			ProgramInput.txtSearch.value = txtQValue12.value;
			ProgramInput.txtTitle.value = txtQValue13.value;
			ProgramInput.txtNumbers.value = txtQValue14.value;

			if (txtQValue16.value != "")
			{
				ProgramInput.cboFormat.selectedIndex = txtQValue16.value;
			}				

			if (txtQValue17.value == "True")
				ProgramInput.chkSummarySearch.checked = true;
			else
				ProgramInput.chkSummarySearch.checked = false;
			if (txtQValue18.value == "True")
				ProgramInput.chkDescriptionSearch.checked = true;
			else
				ProgramInput.chkDescriptionSearch.checked = false;
			if (txtQValue19.value == "True")
				ProgramInput.chkActionSearch.checked = true;
			else
				ProgramInput.chkActionSearch.checked = false;

			if (txtQValue22.value != "")
			{
				ProgramInput.cboDaysClosedCompare.selectedIndex = txtQValue22.value;
			}				
				
			ProgramInput.txtDaysClosed.value = txtQValue23.value;

			if (txtQValue24.value != "")

			{
				ProgramInput.cboDaysTargetCompare.selectedIndex = txtQValue24.value;
			}				
			
			ProgramInput.txtDaysTarget.value = txtQValue25.value;

			ProgramInput.txtOpenRange1.value = txtQValue27.value;
			ProgramInput.txtOpenRange2.value = txtQValue28.value;
			ProgramInput.txtClosedRange1.value = txtQValue29.value;
			ProgramInput.txtClosedRange2.value = txtQValue30.value;
			ProgramInput.txtTargetRange1.value = txtQValue31.value;
			ProgramInput.txtTargetRange2.value = txtQValue32.value;


			if (txtQValue33.value == "True")
				ProgramInput.chkIncludeActions.checked = true;
			else
				ProgramInput.chkIncludeActions.checked = false;
	
			if (txtQValue34.value == "True")
				ProgramInput.chkIncludeDescription.checked = true;
			else
				ProgramInput.chkIncludeDescription.checked = false;

			if (txtQValue35.value == "True")
				ProgramInput.chkIncludeJustification.checked = true;
			else
				ProgramInput.chkIncludeJustification.checked = false;

			if (txtQValue36.value == "True")
				ProgramInput.chkIncludeResolution.checked = true;
			else
				ProgramInput.chkIncludeResolution.checked = false;

			if (txtQValue37.value == "True")
				ProgramInput.chkIncludeApprovers.checked = true;
			else
				ProgramInput.chkIncludeApprovers.checked = false;
				
			if (txtQValue38.value == "True")
			    ProgramInput.chkApproverComments.checked = true;
			else
			    ProgramInput.chkApproverComments.checked = false;
			    
			ProgramInput.txtApproverComments.value = txtQValue41.value;

			for (i=0;i<ProgramInput.lstProductGroups.length;i++)
				{
				if (ProgramInput.lstProductGroups.options[i].value != "" && instr("," + txtQValue42.value + "," , "," + ProgramInput.lstProductGroups.options[i].value + "," ) > -1)
					ProgramInput.lstProductGroups.options[i].selected = true;
				else
					ProgramInput.lstProductGroups.options[i].selected = false;
				}

			cboDaysOpenCompare_onchange();
			cboDaysClosedCompare_onchange();
			cboDaysTargetCompare_onchange();
		}			

}

function cboProfile_onkeydown() {
	//if(event.keyCode == 46 && DetailsForm.cboProfile.selectedIndex > 2)
	//	window.alert("Deleting");
}

var strNewName;

function RenameProfile(){
	strNewName = window.prompt("Enter new name for this profile.",ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex].text);
	
	if (strNewName != null)
		{
		jsrsExecute("ActionsRSupdate.asp", myCallback2, "ProfileStrings", Array(ProgramInput.cboProfile.value,"1",strNewName,"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""));			
		}
}
function myCallback2( returnstring ){
		if (returnstring != 1)
			window.alert("Unable to update this profile."); 
		else
			ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex].text = strNewName;
}

function UpdateProfile(){
	var i;
	var strProduct = "";
	var strProductPulsar = "";
	var strOwners = "";
	var strApprovers = "";
	var strSubmitter = "";
	var strType = "";
	var strStatus = "";
	var strProductGroup	= "";
	
	
	//Product List
	for(i=0;i<ProgramInput.lstProducts.length;i++)
		if (ProgramInput.lstProducts.options[i].selected)
			strProduct = strProduct + "," + ProgramInput.lstProducts.options[i].value;
	if (strProduct.length>0)
		strProduct = strProduct.substr(1);

    //Pulsar Product List
	for(i=0;i<ProgramInput.lstProductsPulsar.length;i++)
	    if (ProgramInput.lstProductsPulsar.options[i].selected)
	        strProductPulsar = strProductPulsar + "," + ProgramInput.lstProductsPulsar.options[i].value;
	if (strProductPulsar.length>0)
	    strProductPulsar = strProductPulsar.substr(1);

	//Owners List
	for(i=0;i<ProgramInput.lstOwners.length;i++)
		if (ProgramInput.lstOwners.options[i].selected)
			strOwners = strOwners + "," + ProgramInput.lstOwners.options[i].value;
	if (strOwners.length>0)
		strOwners = strOwners.substr(1);

	//Approvers List
	for(i=0;i<ProgramInput.lstApprovers.length;i++)
		if (ProgramInput.lstApprovers.options[i].selected)
			strApprovers = strApprovers + "," + ProgramInput.lstApprovers.options[i].value;
	if (strApprovers.length>0)
		strApprovers = strApprovers.substr(1);

	//Submitter List
	for(i=0;i<ProgramInput.lstSubmitter.length;i++)
		if (ProgramInput.lstSubmitter.options[i].selected)
			strSubmitter = strSubmitter + "," + ProgramInput.lstSubmitter.options[i].value;
	if (strSubmitter.length>0)
		strSubmitter = strSubmitter.substr(1);

	//Type List
	for(i=0;i<ProgramInput.lstType.length;i++)
		if (ProgramInput.lstType.options[i].selected)
			strType = strType + "," + ProgramInput.lstType.options[i].value;
	if (strType.length>0)
		strType = strType.substr(1);

	//Status List
	for(i=0;i<ProgramInput.lstStatus.length;i++)
		if (ProgramInput.lstStatus.options[i].selected)
			strStatus = strStatus + "," + ProgramInput.lstStatus.options[i].value;
	if (strStatus.length>0)
		strStatus = strStatus.substr(1);

	//ProductGroup List
	for(i=0;i<ProgramInput.lstProductGroups.length;i++)
		if (ProgramInput.lstProductGroups.options[i].selected)
			strProductGroup = strProductGroup + "," + ProgramInput.lstProductGroups.options[i].value;
	if (strProductGroup.length>0)
		{
		strProductGroup = strProductGroup.substr(1);
		}

	//Update
	if(window.confirm("Are you sure you want to update this profile?"))
		{
	    jsrsExecute("ActionsRSupdate.asp", myCallback3, "ProfileStrings", Array(ProgramInput.cboProfile.value,"3","",strProduct,strOwners,strApprovers,ProgramInput.cboDaysOpenCompare.selectedIndex.toString(),ProgramInput.txtDaysOpen.value.toString(),ProgramInput.cboApproverStatus.selectedIndex.toString(),ProgramInput.cboCategory.selectedIndex.toString(),strStatus.toString(),strType.toString(),"",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,"",ProgramInput.cboFormat.selectedIndex.toString(),ProgramInput.chkSummarySearch.checked.toString(),ProgramInput.chkDescriptionSearch.checked.toString(),ProgramInput.chkActionSearch.checked.toString(),"0","0",ProgramInput.cboDaysClosedCompare.selectedIndex.toString(),ProgramInput.txtDaysClosed.value.toString(),ProgramInput.cboDaysTargetCompare.selectedIndex.toString(),ProgramInput.txtDaysTarget.value.toString(),strSubmitter,ProgramInput.txtOpenRange1.value,ProgramInput.txtOpenRange2.value,ProgramInput.txtClosedRange1.value,ProgramInput.txtClosedRange2.value,ProgramInput.txtTargetRange1.value,ProgramInput.txtTargetRange2.value,ProgramInput.chkIncludeActions.checked.toString(),ProgramInput.chkIncludeDescription.checked.toString(),ProgramInput.chkIncludeJustification.checked.toString(),ProgramInput.chkIncludeResolution.checked.toString(),ProgramInput.chkIncludeApprovers.checked.toString(),ProgramInput.chkApproverComments.checked.toString(),"","",ProgramInput.txtApproverComments.value,strProductGroup,"","",strProductPulsar));			
		
		}
}

function myCallback3( returnstring ){
		if (returnstring != 1)
			window.alert("Unable to update this profile."); 
		else
			window.alert("Update Complete");
}



function AddProfile(){
	var i;
	var strProduct = "";
	var strProductPulsar = "";
	var strOwners = "";
	var strApprovers = "";
	var strSubmitter = "";
	var strType = "";
	var strStatus = "";
	var strProductGroup = "";

	strNewName = window.prompt("Enter a name for the new profile.","");
	if (strNewName != null)
		{
	
		//Product List
		for(i=0;i<ProgramInput.lstProducts.length;i++)
			if (ProgramInput.lstProducts.options[i].selected)
				strProduct = strProduct + "," + ProgramInput.lstProducts.options[i].value;
		if (strProduct.length>0)
			strProduct = strProduct.substr(1);
	
	    //Pulsar Product List
		for(i=0;i<ProgramInput.lstProductsPulsar.length;i++)
		    if (ProgramInput.lstProductsPulsar.options[i].selected)
		        strProductPulsar = strProductPulsar + "," + ProgramInput.lstProductsPulsar.options[i].value;
		if (strProductPulsar.length>0)
		    strProductPulsar = strProductPulsar.substr(1);

		//Owners List
		for(i=0;i<ProgramInput.lstOwners.length;i++)
			if (ProgramInput.lstOwners.options[i].selected)
				strOwners = strOwners + "," + ProgramInput.lstOwners.options[i].value;
		if (strOwners.length>0)
			strOwners = strOwners.substr(1);
	
		//Approvers List
		for(i=0;i<ProgramInput.lstApprovers.length;i++)
			if (ProgramInput.lstApprovers.options[i].selected)
				strApprovers = strApprovers + "," + ProgramInput.lstApprovers.options[i].value;
		if (strApprovers.length>0)
			strApprovers = strApprovers.substr(1);
	
		//Submitter List
		for(i=0;i<ProgramInput.lstSubmitter.length;i++)
			if (ProgramInput.lstSubmitter.options[i].selected)
				strSubmitter = strSubmitter + "," + ProgramInput.lstSubmitter.options[i].value;
		if (strSubmitter.length>0)
			strSubmitter = strSubmitter.substr(1);

		//Type List
		for(i=0;i<ProgramInput.lstType.length;i++)
			if (ProgramInput.lstType.options[i].selected)
				strType = strType + "," + ProgramInput.lstType.options[i].value;
		if (strType.length>0)
			strType = strType.substr(1);
	
		//Status List
		for(i=0;i<ProgramInput.lstStatus.length;i++)
			if (ProgramInput.lstStatus.options[i].selected)
				strStatus = strStatus + "," + ProgramInput.lstStatus.options[i].value;
		if (strStatus.length>0)
			strStatus = strStatus.substr(1);
	
		//ProductGroup List
		for(i=0;i<ProgramInput.lstProductGroups.length;i++)
			if (ProgramInput.lstProductGroups.options[i].selected)
				strProductGroup = strProductGroup + "," + ProgramInput.lstProductGroups.options[i].value;
		if (strProductGroup.length>0)
			{
			strProductGroup = strProductGroup.substr(1);
			}
	
		//Add
		
		//var objRS = RSGetASPObject("ActionsRSupdate.asp");
		//var objResult = objRS.updateProfile(ProgramInput.cboProfile.value,4,NewName,strProduct,strOwners,strApprovers,ProgramInput.cboDaysOpenCompare.selectedIndex,ProgramInput.txtDaysOpen.value,ProgramInput.cboApproverStatus.selectedIndex,ProgramInput.cboCategory.selectedIndex,strStatus,strType,"",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,"","",ProgramInput.cboFormat.selectedIndex,ProgramInput.chkSummarySearch.checked,ProgramInput.chkDescriptionSearch.checked,ProgramInput.chkActionSearch.checked,0,0,ProgramInput.cboDaysClosedCompare.selectedIndex,ProgramInput.txtDaysClosed.value,ProgramInput.cboDaysTargetCompare.selectedIndex,ProgramInput.txtDaysTarget.value,strSubmitter,ProgramInput.txtOpenRange1.value,ProgramInput.txtOpenRange2.value,ProgramInput.txtClosedRange1.value,ProgramInput.txtClosedRange2.value,ProgramInput.txtTargetRange1.value,ProgramInput.txtTargetRange2.value,ProgramInput.chkIncludeActions.checked,ProgramInput.chkIncludeDescription.checked,ProgramInput.chkIncludeJustification.checked,ProgramInput.chkIncludeResolution.checked,2,txtUserID.value);

		jsrsExecute("ActionsRSupdate.asp", myCallback4, "ProfileStrings", Array(ProgramInput.cboProfile.value,"4",strNewName,strProduct,strOwners,strApprovers,ProgramInput.cboDaysOpenCompare.selectedIndex.toString(),ProgramInput.txtDaysOpen.value,ProgramInput.cboApproverStatus.selectedIndex.toString(),ProgramInput.cboCategory.selectedIndex.toString(),strStatus,strType,"",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,"",ProgramInput.cboFormat.selectedIndex.toString(),ProgramInput.chkSummarySearch.checked.toString(),ProgramInput.chkDescriptionSearch.checked.toString(),ProgramInput.chkActionSearch.checked.toString(),"0","0",ProgramInput.cboDaysClosedCompare.selectedIndex.toString(),ProgramInput.txtDaysClosed.value,ProgramInput.cboDaysTargetCompare.selectedIndex.toString(),ProgramInput.txtDaysTarget.value,strSubmitter,ProgramInput.txtOpenRange1.value,ProgramInput.txtOpenRange2.value,ProgramInput.txtClosedRange1.value,ProgramInput.txtClosedRange2.value,ProgramInput.txtTargetRange1.value,ProgramInput.txtTargetRange2.value,ProgramInput.chkIncludeActions.checked.toString(),ProgramInput.chkIncludeDescription.checked.toString(),ProgramInput.chkIncludeJustification.checked.toString(),ProgramInput.chkIncludeResolution.checked.toString(),ProgramInput.chkIncludeApprovers.checked.toString(),ProgramInput.chkApproverComments.checked.toString(),"","",ProgramInput.txtApproverComments.value,strProductGroup,"2",txtUserID.value,strProductPulsar));			
		}
}	

function myCallback4( returnstring ){
	if (returnstring == "" || returnstring == 0 )
		window.alert("Unable to add this profile."); 
	else
		{
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.length] = new Option(strNewName,returnstring);
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.length-1].selected = true;
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.length-1].CanEdit="True";
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.length-1].CanDelete="True";
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.length-1].PrimaryOwner="";
		ProfileOptionsUpdate.style.display = "";		
		ProfileOptionsDelete.style.display = "";		
		ProfileOptionsRename.style.display = "";		
		ProfileOptionsOwner.style.display = "none";	
		ProfileOptionsRemove.style.display = "none";
		ProfileOptionsShare.style.display = "";		
		window.alert("Profile Added.");
		}
}


function DeleteProfile(){
	if(window.confirm("Are you sure you want to delete this profile?"))
		{
		jsrsExecute("ActionsRSupdate.asp", myCallback5, "ProfileStrings", Array(ProgramInput.cboProfile.value,"2","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""));			
//		var objResult = objRS.updateProfile(ProgramInput.cboProfile.value,2,"","","");

		}
}

function myCallback5( returnstring ){
		if (returnstring != 1)
			window.alert("Unable to delete this profile."); 
		else
			{
		    ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex] = null;
			ProfileOptionsUpdate.style.display = "none";		
			ProfileOptionsDelete.style.display = "none";		
			ProfileOptionsRename.style.display = "none";		
			ProfileOptionsOwner.style.display = "none";	
			ProfileOptionsRemove.style.display = "none";
			ProfileOptionsShare.style.display = "none";		
			cmdReset_onclick();
			}
}


function window_onload() {
	lblLoad.style.display = "none";
	lblInst.style.display = "";
}


function RemoveProfile(){
	if(window.confirm("Are you sure you want to stop receiving this shared profile?"))
		jsrsExecute("../OTSMailRSshare.asp", myCallbackRemoveSharing, "ProfileSharing", Array(ProgramInput.cboProfile(ProgramInput.cboProfile.selectedIndex).SharingID));			

}


function myCallbackRemoveSharing (returnstring){
	if (returnstring == "Error")
		alert("Unable to remove this profile.");
	else
		{
	    ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex] = null;
		ProfileOptionsUpdate.style.display = "none";		
		ProfileOptionsDelete.style.display = "none";		
		ProfileOptionsRename.style.display = "none";	
		ProfileOptionsOwner.style.display = "none";		
		ProfileOptionsRemove.style.display = "none";
		ProfileOptionsShare.style.display = "none";		
		cmdReset_onclick();
		}
}

function ShareProfile(){
	var strResult;
	strResult = window.showModalDialog("ProfileShare.asp?ID=" + ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex].value,"","dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
//		if (typeof(strResult) != "undefined")
	
}


function ActionCell_onmouseover() {
	window.event.srcElement.style.background="gainsboro";
	window.event.srcElement.style.cursor = "hand";
	window.event.srcElement.style.color = "black";	
}

function ActionCell_onmouseout() {

	window.event.srcElement.style.color = "white";	
	window.event.srcElement.style.background="#333333";
}



//-->
</script>
</head>
<STYLE>
TEXTAREA
{
    FONT-WEIGHT: normal;
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
}
A:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}
LEGEND
{
    FONT-WEIGHT: normal;
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	Color=black;
}

TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	Color=black;
}

TD.HeaderButton
{
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
	FONT-WEIGHT: bold;
	COLOR: White;
}

</STYLE>
<body bgcolor="ivory" LANGUAGE="javascript" onload="return window_onload()">
<font size="3" face="verdana"><b>
<%if request("CAT") = "1" then%>
SKU Change Requests - Advanced Query
<%else%>
Action Items - Advanced Query
<%end if%>
</b></font>

<span ID="lblLoad"><br><br><font size="2" face="verdana">Loading.  Please wait...</font></span>
<%
	

	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function		

	dim cn
	dim cm
	dim p
	dim rs
	dim strSQL
	dim strStates
	dim strProductOptions
    dim strProductOptionsPulsar
	dim strInList
	dim strcatOptions
	dim strOwners
	dim strApprovers
	dim CurrentUser
	dim CurrentUserID
	dim strDivision
	dim strSubmitters
	dim strProductGroupOptions

	strProductGroupOptions = ""

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
		strDivision = rs("Division") & ""
		CurrentUserID = rs("ID") & ""
	else
		strDivision = ""
		CurrentUserID = ""
	end if
	
	rs.Close

    rs.open "usp_getActionReportProductVersions " & strproductid,cn,adOpenForwardOnly
					
	strProductOptions = ""
    strProductOptionsPulsar = ""
	strInList = ""

	do while not rs.EOF
		if rs("DOTSName") <> "Test Product 1.0" then 
          if rs("FusionRequirements") = 0 then
			if trim(request("ID")) = trim(rs("ID")) then
				strProductOptions = strProductOptions &  "<Option selected value= """ & rs("ID") & """>" & rs("DOTSName") & "</OPTION>"
			else
				strProductOptions = strProductOptions &  "<Option value= """ & rs("ID") & """>" & rs("DOTSName") & "</OPTION>"
			end if
         else
            if trim(request("ID")) = trim(rs("ID")) then
				strProductOptionsPulsar = strProductOptionsPulsar &  "<Option selected value= """ & rs("ID") & """>" & rs("DOTSName") & "</OPTION>"
			else
				strProductOptionsPulsar = strProductOptionsPulsar &  "<Option value= """ & rs("ID") & """>" & rs("DOTSName") & "</OPTION>"
			end if
         end if
			strInList = strInList & "," & rs("ID") & ""
		end if
		rs.MoveNext
	loop
	rs.Close


'	if trim(currentuserpartner) = "1" then
		strSQL = "spListPartners 2"
	
		rs.Open strSQL,cn,adOpenForwardOnly
		strPartnerList = ""
		strProductGroupOptions = strProductGroupOptions & "<Option value="""">------------ODM-------------</Option>"
		do while not rs.EOF
			if rs("ID") <> 1 then
				strPartnerList = strPartnerList &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
				strProductGroupOptions= strProductGroupOptions &  "<Option value= ""1:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close
'	end if
	
	strSQL = "spListPrograms"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strProductGroupOptions = strProductGroupOptions & "<Option value="""">-----------Cycle------------</Option>"
	do while not rs.EOF
		'if rs("BusinessID") = 2 then
			strProductGroupOptions= strProductGroupOptions &  "<Option value= ""2:" & rs("ID") & """>" & rs("FullName") & "</OPTION>"
		'else
		'	strProductGroupOptions= strProductGroupOptions &  "<Option value= ""2:" & rs("ID") & """>BNB " & rs("Name") & "</OPTION>"
		'end if
		rs.MoveNext
	loop
	rs.Close
	
	
	strSQL = "spListDevCenters"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strProductGroupOptions = strProductGroupOptions & "<Option value="""">-------Dev. Center---------</Option>"
	do while not rs.EOF
		strProductGroupOptions= strProductGroupOptions &  "<Option value= ""3:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close

	strSQL = "spListProductStatuses"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strProductGroupOptions = strProductGroupOptions & "<Option value="""">-----Product Phase-----</Option>"
	do while not rs.EOF
		strProductGroupOptions= strProductGroupOptions &  "<Option value= ""4:" & rs("ID") & """>" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close
	


	strSQL = "spListActionOwners"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strOwners = ""
	do while not rs.EOF
		strOwners = strOwners &  "<Option value= """ & rs("ID") & """>" & rs("Owner") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close



	strSQL = "Select Distinct Submitter from DeliverableIssues with (NOLOCK) order by submitter;"
	strSubmitters = ""	
	rs.Open strSQL,cn,adOpenForwardOnly
	strSubmitters = ""
	do while not rs.EOF
		strSubmitters = strSubmitters &  "<Option value= """ & replace(rs("Submitter") & "",",","|") & """>" & rs("Submitter") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close

	

	
	strSQL = "spListActionApprovers"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strApprovers = ""
	do while not rs.EOF
		strApprovers = strApprovers &  "<Option value= """ & rs("ID") & """>" & rs("Approver") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close
	
	
	dim strProfileOptions
	
	strProfileOptions = ""
	
	if CurrentUserID <> "" then
		rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",2",cn,adOpenForwardOnly
		strProfileOptions = ""
		do while not rs.EOF
			strProfileOptions = strProfileOptions & "<Option SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			rs.MoveNext
		loop
		rs.Close
		
		rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",2",cn,adOpenForwardOnly
		do while not rs.EOF
			strProfileOptions = strProfileOptions & "<Option SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")& "")  &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			rs.MoveNext
		loop
		rs.Close

		rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",2",cn,adOpenForwardOnly
		do while not rs.EOF
			strProfileOptions = strProfileOptions & "<Option CanRemove=0 SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")& "")  &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			rs.MoveNext
		loop
		rs.Close
		
	end if		

	if strProfileOptions <> "" then
		strProfileOptions = "<option selected>Use Options Selected Below</option><option>----------------------------------------------------------</option>" & strProfileOptions	
	else
		strProfileOptions = "<option selected>Use Options Selected Below</option>"	
	end if	
	
	
%>
<form ACTION="ActionReport.asp" METHOD="post" NAME="ProgramInput" target="ActionReport">
<input type="hidden" id="txtDivision" name="txtDivision" value="<%=strDivision%>">
<table>
	<tr>
		<td colspan="8">
		<font face=verdana size="2"><b>Report Profile:&nbsp;</b></font>
		<select id="cboProfile" name="cboProfile" style="WIDTH: 250px" LANGUAGE=javascript onkeydown="return cboProfile_onkeydown()" onchange="return cboProfile_onchange()"><%=strProfileOptions%></select></font>
		<font size="2" face="verdana" color=blue>
		<font size=1 face=verdana><a href="javascript:AddProfile();">Add</a></font> 

		<span style="Display:none" ID=ProfileOptionsUpdate ><font size=1 face=verdana><a href="javascript:UpdateProfile();">Update</a></font> </span> 
		<span style="Display:none" ID=ProfileOptionsDelete ><font size=1 face=verdana><a href="javascript:DeleteProfile();">Delete</a></font> </span>
		<span style="Display:none" ID=ProfileOptionsRename ><font size=1 face=verdana><a href="javascript:RenameProfile();">Rename</a></font> </span>
		<span style="Display:none" ID=ProfileOptionsRemove><font size=1 face=verdana><a href="javascript:RemoveProfile();">Remove</a></font> </span>
		<span style="Display:none" ID=ProfileOptionsShare><font size=1 face=verdana><a href="javascript:ShareProfile();">Share</a></font> </span>
		<span style="Display:none" ID=ProfileOptionsOwner><font size=1 face=verdana color=black><b>Profile Owner:</b></font> </span>



		</td>
	</tr>
<TR><TD colspan=8><HR>
<TABLE border=0 cellpadding=3 cellspacing=2>
	<TR bgcolor=#333333 ID=HeaderRow>
		<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover()" onmouseout="return ActionCell_onmouseout()" onclick="return cmdReport_onclick()">&nbsp;&nbsp;Summary&nbsp;Report&nbsp;&nbsp;</TD>
		<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover()" onmouseout="return ActionCell_onmouseout()" onclick="return cmdDetails_onclick()">&nbsp;&nbsp;Detailed&nbsp;Report&nbsp;&nbsp;</TD>
		<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover()" onmouseout="return ActionCell_onmouseout()" onclick="return cmdReset_onclick();">&nbsp;&nbsp;Reset&nbsp;&nbsp;</TD>
	</TR>
</TABLE>
<%
	Response.Write "<label ID=lblInst style=""display:none"">"
	Response.Write "<font color=Green size=1 face=verdana>Use CTRL or SHIFT keys to select multiple items in lists</font></label>"
%>

</td></TR>

	<tr>
		<td valign="top"><font size="2" face="verdana"><b>Products (Legacy):</b></font><br><select style="WIDTH: 250px; HEIGHT: 145px" multiple id="lstProducts" name="lstProducts">
				<%=strProductOptions%>
			</select>
		</td>
        <td valign="top"><font size="2" face="verdana"><b>Products (Pulsar):</b></font><br><select style="WIDTH: 260px; HEIGHT: 145px" multiple id="lstProductsPulsar" name="lstProductsPulsar">
				<%=strProductOptionsPulsar%>
			</select>
		</td>
		<td valign="top"><font size="2" face="verdana"><b>Product&nbsp;Group:</b></font><br><select style="WIDTH: 250px; HEIGHT: 145px" multiple id="lstProductGroups" name="lstProductGroups">
				<%
					Response.write strProductGroupOptions
				%>
			</select>
		</td>
		<td><font size="2" face="verdana"><b>Owners:</b></font><br><select style="WIDTH: 250px; HEIGHT: 145px" multiple size="2" id="lstOwners" name="lstOwners">
				<%=strOwners%>
			</select>
		</td>
		<td><font size="2" face="verdana"><b>Approvers:</b></font><br><select style="WIDTH: 250px; HEIGHT: 145px" multiple size="2" id="lstApprovers" name="lstApprovers">
				<%=strApprovers%>
			</select>
		</td>
	</tr>
	<tr>
		<td valign=top><font size="2" face="verdana"><b>Status:</b></font><br>
			<select style="WIDTH: 250px; HEIGHT: 102px" multiple size="2" id="lstStatus" name="lstStatus">
				<option Value="1">New (Open)</option>
				<option Value="2">Closed (Not DCRs)</option>
				<option Value="3">Need More Info</option>
				<option Value="4">Approved</option>
				<option Value="5">Disapproved</option>
				<option Value="6">Investigating</option>
			</select>
		</td>
		<td valign="top"><font size="2" face="verdana"><b>Type:</b></font><br>
			<select style="WIDTH: 260px; HEIGHT: 102px" multiple size="2" id="lstType" name="lstType">
				<%if request("Type") = "1" then%>
					<option selected Value="1">Issue</option>
				<%Else%>
					<option Value="1">Issue</option>
				<%end if%>
				<%if request("Type") = "2" then%>
					<option selected Value="2">Action</option>
				<%Else%>
					<option Value="2">Action</option>
				<%end if%>
				<%if request("Type") = "3" then%>
					<option selected Value="3">Change Request (DCR)</option>
				<%Else%>
					<option Value="3">Change Request (DCR)</option>
				<%end if%>
				<%if request("Type") = "4" then%>
					<option selected Value="4">Status Note</option>
				<%Else%>
					<option Value="4">Status Note</option>
				<%end if%>
				<%if request("Type") = "5" then%>
					<option selected Value="5">Improvement Opportunity</option>
				<%Else%>
					<option Value="5">Improvement Opportunity</option>
				<%end if%>
				<%if request("Type") = "6" then%>
					<option selected Value="6">Test Request</option>
				<%Else%>
					<option Value="6">Test Request</option>
				<%end if%>
				<%if request("Type") = "7" then%>
					<option selected Value="7">Service ECR</option>
				<%Else%>
					<option Value="7">Service ECR</option>
				<%end if%>
				
			
			</select>
		</td>
		<td valign="top"><font size="2" face="verdana"><b>Submitter:</b></font><br>
			<select style="WIDTH: 250px; HEIGHT: 102px" multiple size="2" id="lstSubmitter" name="lstSubmitter">
				<%=strSubmitters%>
			</select>
		</td>
        <td></td>
        <td></td>
	</tr>
</table>
<table>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Report Title:<b></font></td>
	<td> <input type="text" id="txtTitle" name="txtTitle" value="Action Item Report" style="Width:460" maxlength=255></td>
  </tr>
	<TR>
		<TD nowrap width=120 valign=top><font size=2 face=verdana><b>ID&nbsp;Numbers:</b></font></TD>
		<td><INPUT id=txtNumbers name=txtNumbers style="WIDTH: 460; HEIGHT: 22px" size=46 maxlength=80><font size=1 face=verdana color=Green>&nbsp;(comma&nbsp;seperated)</font></td>
	</TR>  
  
  <tr>
	<td nowrap width="120"><font face="verdana" size="2"><b>Category:</b><BR></font></td>
	<td valign=center>
		<select style="Width:170" id="cboCategory" name="cboCategory">
			<option value=0 selected></option>
			<%if request("CAT") = "3" then%>
				<option selected value=3>Requirement Changes</option>
			<%else%>
				<option value=3>Requirement Changes</option>
			<%end if%>
			<%if request("CAT") = "1" then%>
				<option selected value=1>SKU Changes</option>
			<%else%>
				<option value=1>SKU Changes</option>
			<%end if%>
			<%if request("CAT") = "2" then%>
				<option selected value=2>Software Changes</option>
			<%else%>
				<option value=2>Software Changes</option>
			<%end if%>
			<%if request("CAT") = "6" then%>
				<option selected value=6>Commodity Changes</option>
			<%else%>
				<option value=6>Commodity Changes</option>
			<%end if%>
			<%if request("CAT") = "5" then%>
				<option selected value=5>Document Changes</option>
			<%else%>
				<option value=5>Document Changes</option>
			<%end if%>
			<%if request("CAT") = "4" then%>
				<option selected value=4>Other Changes</option>
			<%else%>
				<option value=4>Other Changes</option>
			<%end if%>
		</select>&nbsp;<font size=1 face=verdana color=green>(Change Requests only)</font>
	</td>
  </tr>
  <tr>
	<td nowrap width="120"><font face="verdana" size="2"><b>Report Format:<b></font></td>
	<td>
		<select style="Width:100" id="cboFormat" name="cboFormat">
			<option selected value=0>HTML</option>
			<option value=1>Excel</option>
			<option value=2>Word</option>
		</select>&nbsp;
	</td>
  </tr>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>ECN:&nbsp;<b></font></td>
	<td>
		<SELECT  style="Width:100" id=cboECN name=cboECN>
			<OPTION value="0" selected></OPTION>
			<OPTION value="1">Complete</OPTION>
			<OPTION value="2">Pending</OPTION>
		</SELECT>
	</td>
  </tr>  
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Search:<b></font></td>
	<td nowrap> <input type="text" id="txtSearch" name="txtSearch" value="" style="Width:134" maxlength=80>
	<font size=2 face=verdana>
	    Look&nbsp;In:&nbsp;
	    <INPUT type="checkbox" checked id=chkSummarySearch name=chkSummarySearch>Summary&nbsp;&nbsp;
	    <INPUT type="checkbox" id=chkDescriptionSearch name=chkDescriptionSearch>Description&nbsp;&nbsp;
	    <INPUT type="checkbox" id=chkActionSearch name=chkActionSearch>Actions&nbsp;&nbsp;
	    <input type="checkbox"" id=chkApproverComments name=chkApproverComments />Approver&nbsp;Comments
	    </font>
	
	</td>
  </tr>
  <tr>
	<td nowrap width="120"><font face="verdana" size="2"><b>Approver&nbsp;Status:<b></font></td>
	<td>
		<select style="Width:134" id="cboApproverStatus" name="cboApproverStatus">
			<option selected value=0></option>
			<option value=1>Requested</option>
			<option value=2>Approved</option>
			<option value=3>Disapproved</option>
		</select>&nbsp;<font size=1 face=verdana color=green>(applies to selected approvers only)</font>
	</td>
  </tr>
  <tr>
	<td nowrap width="120"><font face="verdana" size="2"><b>Approver&nbsp;Comments:<b></font></td>
	<td nowrap> <input type="text" id="txtApproverComments" name="txtApproverComments" value="" style="Width:134" maxlength=80>&nbsp;<font size=1 face=verdana color=green>(applies to selected approvers only)</font></td>
  </tr>
  <tr>
	<td nowrap width="120"><font face="verdana" size="2"><b>Date Opened:<b></font></td>
	<td>
		<select style="WIDTH:70" id="cboDaysOpenCompare" name="cboDaysOpenCompare" LANGUAGE="javascript" onchange="return cboDaysOpenCompare_onchange()">
			<option><=</option>
			<option>=</option>
			<option selected>&gt;=</option>
			<option>Range</option>
		</select>&nbsp;
		<span ID="spnOpenCount"><input style="width:55" type="text" id="txtDaysOpen" name="txtDaysOpen" value="0"> <font size="2" face="verdana">Days Ago</font></span>
		<span style="Display:none" ID="spnOpenRange"><font size="2" face="verdana">Between:&nbsp;<input style="width:100" type="text" id="txtOpenRange1" name="txtOpenRange1" maxlength=25>&nbsp;<a href="javascript:PickDate(1);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a>&nbsp;and&nbsp;<input style="width:100" type="text" id="txtOpenRange2" name="txtOpenRange2" maxlength=25>&nbsp;<a href="javascript:PickDate(2);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a></font></span>
		
	</td>
  </tr>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Date Closed:<b></font></td>
	<td>
		<select style="WIDTH:70" id="cboDaysClosedCompare" name="cboDaysClosedCompare" LANGUAGE="javascript" onchange="return cboDaysClosedCompare_onchange()">
			<option><=</option>
			<option>=</option>
			<option selected>&gt;=</option>
			<option>Range</option>
		</select>&nbsp;
		<span ID="spnClosedCount"><input style="width:55" type="text" id="txtDaysClosed" name="txtDaysClosed" value="0"> <font size="2" face="verdana">Days Ago</font></span>
		<span style="Display:none" ID="spnClosedRange"><font size="2" face="verdana">Between:&nbsp;<input style="width:100" type="text" id="txtClosedRange1" name="txtClosedRange1" maxlength=25>&nbsp;<a href="javascript:PickDate(3);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a>&nbsp;and&nbsp;<input style="width:100" type="text" id="txtClosedRange2" name="txtClosedRange2" maxlength=25>&nbsp;<a href="javascript:PickDate(4);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a></font></span>
		
	</td>
  </tr>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Target Date:<b></font></td>
	<td>
		<select style="WIDTH:70" id="cboDaysTargetCompare" name="cboDaysTargetCompare" LANGUAGE="javascript" onchange="return cboDaysTargetCompare_onchange()">
			<option><=</option>
			<option>=</option>
			<option selected>&gt;=</option>
			<option>Range</option>
		</select>&nbsp;
		<span ID="spnTargetCount"><input style="width:55" type="text" id="txtDaysTarget" name="txtDaysTarget" value="0"> <font size="2" face="verdana">Days From Now</font></span>
		<span style="Display:none" ID="spnTargetRange"><font size="2" face="verdana">Between:&nbsp;<input style="width:100" type="text" id="txtTargetRange1" name="txtTargetRange1" maxlength=25>&nbsp;<a href="javascript:PickDate(5);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a>&nbsp;and&nbsp;<input style="width:100" type="text" id="txtTargetRange2" name="txtTargetRange2" maxlength=25>&nbsp;<a href="javascript:PickDate(6);"><img SRC="../MobileSE/Today/images/calendar.gif" alt="Choose" border="0" align="absmiddle" WIDTH="26" HEIGHT="21"></a></font></span>
		
	</td>
  </tr>  
  <tr><td width="120"><span style="font-family:Verdana; font-size: x-small; font-weight:bold;">DCR Category:</span></td>
  <td style="font-family:Verdana; font-size:x-small;">
  <input type="radio" id="rdoDcrCategoryAll" name="rdoDcrCategory" checked value="NULL">All</input>
  <input type="radio" id="rdoDcrCategoryDcr" name="rdoDcrCategory" value="0">DCR</input>
  <input type="radio" id="rdoDcrCategoryBcr" name="rdoDcrCategory" value="1">BCR (BIOS)</input>
  <input type="radio" id="rdoDcrCategoryScr" name="rdoDcrCategory" value="2">SCR (SW)</input>
  </td></tr>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Include&nbsp;Columns:&nbsp;<b></font></td>
	<td>
		<INPUT type="checkbox" id=chkIncludeActions name=chkIncludeActions>&nbsp;<font face=verdana size=2>Actions</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkIncludeApprovers name=chkIncludeApprovers>&nbsp;<font face=verdana size=2>Approvers</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkIncludeDescription name=chkIncludeDescription>&nbsp;<font face=verdana size=2>Description</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkIncludeJustification name=chkIncludeJustification>&nbsp;<font face=verdana size=2>Justification</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkIncludeResolution name=chkIncludeResolution>&nbsp;<font face=verdana size=2>Resolution</font>&nbsp;&nbsp;
	</td>
  </tr>  
  
	<TR>
		<TD nowrap width=100 valign=top><font size=2 face=verdana><b ID=lblSort>Sort&nbsp;Order:</b></font></TD>
		<td colspan=5>
			<SELECT id=Sort1Column name=Sort1Column style="Width=110px">
				<OPTION selected></OPTION>
				<OPTION value="d.ID">ID</OPTION>
				<OPTION value="v.DOTSName">Product</OPTION>
				<OPTION value="d.Status" >Status</OPTION>
				<OPTION value="d.Created">Date Created</OPTION>
				<OPTION value="d.ActualDate" >Date Closed</OPTION>
				<OPTION value="d.targetDate">Target Date</OPTION>
				<OPTION value="d.submitter">Submitter</OPTION>
				<OPTION value="e.name">Owner</OPTION>
			</SELECT>
			 <SELECT style="Width=53px" id=Sort1Direction name=Sort1Direction>
				<OPTION selected></OPTION>
				<OPTION >Asc</OPTION>
				<OPTION>Desc</OPTION>
			 </SELECT><font face=verdana size=3><b>&nbsp;&nbsp;,&nbsp;&nbsp;</b></font>
						 
		
			<SELECT id=Sort2Column name=Sort2Column style="Width=110px">
				<OPTION selected></OPTION>
				<OPTION value="d.ID">ID</OPTION>
				<OPTION value="v.DOTSName">Product</OPTION>
				<OPTION value="d.Status" >Status</OPTION>
				<OPTION value="d.Created">Date Created</OPTION>
				<OPTION value="d.ActualDate" >Date Closed</OPTION>
				<OPTION value="d.targetDate">Target Date</OPTION>
				<OPTION value="d.submitter">Submitter</OPTION>
				<OPTION value="e.name">Owner</OPTION>
			 </SELECT>
			 <SELECT style="Width=53px" id=Sort2Direction name=Sort2Direction>
				<OPTION selected></OPTION>
				<OPTION >Asc</OPTION>
				<OPTION >Desc</OPTION>
			 </SELECT>
		
						 <font face=verdana size=3><b>&nbsp;&nbsp;,&nbsp;&nbsp;</b></font>
			<SELECT id=Sort3Column name=Sort3Column style="Width=110px">
				<OPTION selected></OPTION>
				<OPTION value="d.ID">ID</OPTION>
				<OPTION value="v.DOTSName">Product</OPTION>
				<OPTION value="d.Status" >Status</OPTION>
				<OPTION value="d.Created">Date Created</OPTION>
				<OPTION value="d.ActualDate" >Date Closed</OPTION>
				<OPTION value="d.targetDate">Target Date</OPTION>
				<OPTION value="d.submitter">Submitter</OPTION>
				<OPTION value="e.name">Owner</OPTION>
			</SELECT>
			<SELECT style="Width=53px" id=Sort3Direction name=Sort3Direction>
				<OPTION selected></OPTION>
				<OPTION >Asc</OPTION>
				<OPTION >Desc</OPTION>
			 </SELECT>
		</td>
	</TR>  
 
</table>
<!--<br>
<input type="button" value="Reset Form" id="cmdReset" name="cmdReset" LANGUAGE="javascript" onclick="return cmdReset_onclick()">
<input type="button" value="Summary Report" id="cmdReport" name="cmdReport" LANGUAGE="javascript" onclick="return cmdReport_onclick()">
<input type="button" value="Detailed Report" id="cmdDetails" name="cmdDetails" LANGUAGE="javascript" onclick="return cmdDetails_onclick()">
-->
<%
	cn.Close
	set rs = nothing
	set cn=nothing
	
	

%>

<input type="hidden" id="txtFunction" name="txtFunction">

</form>

<div ID=OutputArea>

</Div>

<INPUT type="hidden" id=txtUserID name=txtUserID value="<%=CurrentUserID%>">

</body>
</html>