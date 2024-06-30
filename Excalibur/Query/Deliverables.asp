<%@ Language=VBScript %>

	<%
	
 ' Response.Buffer = True
  'Response.ExpiresAbsolute = Now() - 1
  'Response.Expires = 0
  'Response.CacheControl = "no-cache"
	  
	%>


<html>
<head>
<script language="JavaScript" src="../_ScriptLibrary/jsrsClient.js"></script>
<title>Deliverables - Advanced Query</title>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
<!-- #include file = "../includes/Date.asp" -->

	var oPopup = window.createPopup();


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

function VerifyFields(TypeID){
	var blnOK = true;
	var i;
	var blnSelected = false;
	var blnFound = false;
	
	if (ProgramInput.txtNumbers.value != "")
		{
		if (! isNumericWithCommas(ProgramInput.txtNumbers.value))
			{
			blnOK = false;
			window.alert("You can only enter a comma-seperated list of numbers into the ID field.");
			ProgramInput.txtNumbers.focus();
			}
		}
    if (TypeID==8 && ProgramInput.cboFormat.selectedIndex != 0 /* HTML */ && ProgramInput.cboFormat.selectedIndex != 2 /* MS Word */)
    {
        blnOK = false;
		window.alert("The report format selected is currently not supported.");
		ProgramInput.cboFormat.focus();
    }

	if (blnOK)
		{
		for (i=0;i<ProgramInput.chkChangeType.length;i++)
			{
			if (ProgramInput.chkChangeType(i).checked)
				blnFound = true;
			}
		
		}
		
	if (!blnFound && TypeID==2)
		{
		blnOK = false;
		window.alert("You must select a change type to run the history report.");
		ProgramInput.chkChangeType(0).focus();
		}			
	else if (blnFound)
		{	//Since a history type is checked, they must enter valid start and end dates
		if (ProgramInput.cboHistoryRange.selectedIndex==3 && ProgramInput.txtStartDate.value == "")
			{
			blnOK = false;
			window.alert("You must enter a Start date for history reports.");
			ProgramInput.txtStartDate.focus();
			}			
		else if (ProgramInput.cboHistoryRange.selectedIndex==3 && ! isDate(ProgramInput.txtStartDate.value))
			{
			blnOK = false;
			window.alert("You must enter a valid Start date for history reports.");
			ProgramInput.txtStartDate.focus();
			}			
		else if (ProgramInput.cboHistoryRange.selectedIndex==3 && ProgramInput.txtEndDate.value == "")
			{
			blnOK = false;
			window.alert("You must enter an End date for history reports.");
			ProgramInput.txtEndDate.focus();
			}			
		else if (ProgramInput.cboHistoryRange.selectedIndex==3 && ! isDate(ProgramInput.txtEndDate.value))
			{
			blnOK = false;
			window.alert("You must enter a valid End date for history reports.");
			ProgramInput.txtEndDate.focus();
			}			
		else if (ProgramInput.cboHistoryRange.selectedIndex!=3 && ProgramInput.txtHistoryDays.value=="")
			{
			blnOK = false;
			window.alert("You must enter an number of days if you select Exactly, More Than, or Less Than in the Date Updated field.");
			ProgramInput.txtHistoryDays.focus();
			}			
		else if (ProgramInput.cboHistoryRange.selectedIndex!=3 && ! isNumeric(ProgramInput.txtHistoryDays.value))
			{
			blnOK = false;
			window.alert("You must enter an number of days if you select Exactly, More Than, or Less Than in the Date Updated field.");
			ProgramInput.txtHistoryDays.focus();
			}			
			
			
		}
/*
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

function BuildSQL(){
	var strResult;
	strResult = window.showModalDialog("DeliverableFields.asp",ProgramInput.txtAdvanced.value,"dialogWidth:400px;dialogHeight:150px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strResult) != "undefined")
		{
			if (ProgramInput.txtAdvanced.length > 0)
				ProgramInput.txtAdvanced.value +=  " " + strResult + " " ;
			else
				ProgramInput.txtAdvanced.value +=   strResult + " " ;
			ProgramInput.txtAdvanced.focus();
			if (ProgramInput.txtAdvanced.createTextRange) 
				{
				var r = ProgramInput.txtAdvanced.createTextRange();
				r.move('character',ProgramInput.txtAdvanced.length);
				r.select();
				r.moveStart('character', ProgramInput.txtAdvanced.value.length);
				r.collapse();
				r.select();
				}
			
		}
}
function cmdReport_onclick(strFunction) {
	var i;
	var strDefaultColumns="";
	if (VerifyFields(strFunction))
		{
		if (strFunction=="")
			strFunction="1";
		
		
		//Get the Deault Column List
		switch (strFunction) {
		    case 1:
			    for (i=0;i<ProgramInput.lstColumns.length;i++){
				    if (strDefaultColumns=="")
					    strDefaultColumns = ProgramInput.lstColumns.options[i].text;			
				    else
					    strDefaultColumns = strDefaultColumns + "," + ProgramInput.lstColumns.options[i].text;			
				}
			    ProgramInput.action = "DelReport.asp";	
			    ProgramInput.method = "post";
			    ProgramInput.txtDefaultColumns.value=strDefaultColumns;
			    ProgramInput.txtFunction.value = strFunction;		
                break;
            case 5:
                strDefaultColumns = "";
			    ProgramInput.action = "MdaCompliance.asp";	
			    ProgramInput.method = "post";
			    ProgramInput.txtDefaultColumns.value=strDefaultColumns;
			    ProgramInput.txtFunction.value = strFunction;		
			    break;
		    case 2015:
		        strDefaultColumns = "";
		        ProgramInput.action = "MdaCompliance2015.asp";	
		        ProgramInput.method = "post";
		        ProgramInput.txtDefaultColumns.value=strDefaultColumns;
		        ProgramInput.txtFunction.value = strFunction;		
		        break;
            case 8:
                strDefaultColumns = "";
			    ProgramInput.action = "DeliverableVersionDetails.aspx";	
			    ProgramInput.method = "post";
			    ProgramInput.txtDefaultColumns.value=strDefaultColumns;
			    ProgramInput.txtFunction.value = strFunction;		
		        break;
		    default:
			    strDefaultColumns = "";
			    ProgramInput.action = "DelReport.asp";	
			    ProgramInput.method = "post";
			    ProgramInput.txtDefaultColumns.value=strDefaultColumns;
			    ProgramInput.txtFunction.value = strFunction;		
		        break;
		}
        ProgramInput.submit();
	}
}

function cmdReleases_onclick() {
	if (VerifyFields(1))//report is 6, but uses the same query
		{
		ProgramInput.action = "DelReport.asp";	
		ProgramInput.method = "post";
		ProgramInput.txtFunction.value = "6";		
		ProgramInput.submit();
		}
}


function cmdCommodity_onclick() {
	if (VerifyFields(0))
        {
		ProgramInput.action = "../Deliverable/HardwareMatrix.asp"		
		ProgramInput.txtFunction.value = "";		
		ProgramInput.method = "post";
		ProgramInput.submit();
        }
}

function cmdDetails_onclick() {

	if (VerifyFields(5))
		{
		ProgramInput.txtFunction.value = "5"
		ProgramInput.method = "post";
		ProgramInput.submit();
		}
}


function cmdReset_onclick() {
	PilotLink.innerHTML = "<a href='javascript:GetSpecificChange(1);'>All Changes</a>"
	QualLink.innerHTML = "<a href='javascript:GetSpecificChange(2);'>All Changes</a>"
	AdminCriteriaRow.style.display = "none";
	ProgramInput.reset();
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

			jsrsExecute("DelRSget.asp", myCallback1, "ProfileStrings", ProgramInput.cboProfile.value);			
			
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
	var HistoryArray;
	var HistoryPartArray;
    var RangeArray;
	var i;

			OutputArea.innerHTML = returnstring; 
			if (txtQProfileName.value == "")
				{
					window.alert("Unable to read the selected filter.");			
				}			
			else
				{
					//list of legacy products
					for (i=0;i<ProgramInput.lstProducts.length;i++)
						{
						if (instr("," + txtQValue1.value + "," , "," + ProgramInput.lstProducts.options[i].value + "," ) > -1)
							ProgramInput.lstProducts.options[i].selected = true;
						else
							ProgramInput.lstProducts.options[i].selected = false;
						}


					for (i=0;i<ProgramInput.lstLanguage.length;i++)
						{
						if (instr("," + txtQValue2.value + "," , "," + ProgramInput.lstLanguage.options[i].value + "," ) > -1)
							ProgramInput.lstLanguage.options[i].selected = true;
						else
							ProgramInput.lstLanguage.options[i].selected = false;
						}
				
					for (i=0;i<ProgramInput.lstVendor.length;i++)
						{
						if (instr("," + txtQValue3.value + "," , "," + ProgramInput.lstVendor.options[i].value + "," ) > -1)
							ProgramInput.lstVendor.options[i].selected = true;
						else
							ProgramInput.lstVendor.options[i].selected = false;
						}

					for (i=0;i<ProgramInput.lstCategory.length;i++)
						{
						if (instr("," + txtQValue4.value + "," , "," + ProgramInput.lstCategory.options[i].value + "," ) > -1)
							ProgramInput.lstCategory.options[i].selected = true;
						else
							ProgramInput.lstCategory.options[i].selected = false;
						}

					if (txtQValue6.value != "")
					{
						ProgramInput.ReportFormat.selectedIndex = txtQValue6.value;
					}				


					for (i=0;i<ProgramInput.lstCommodityPM.length;i++)
						{
						if (instr("," + txtQValue7.value + "," , "," + ProgramInput.lstCommodityPM.options[i].value + "," ) > -1)
							{
							ProgramInput.lstCommodityPM.options[i].selected = true;
							}
						}


					for (i=0;i<ProgramInput.lstOS.length;i++)
						{
						if (instr("," + txtQValue15.value + "," , "," + ProgramInput.lstOS.options[i].value + "," ) > -1)
							ProgramInput.lstOS.options[i].selected = true;
						else
							ProgramInput.lstOS.options[i].selected = false;
						}
					for (i=0;i<ProgramInput.lstDeveloper.length;i++)
						{
						if (instr("," + txtQValue26.value + "," , "," + ProgramInput.lstDeveloper.options[i].value + "," ) > -1)
							ProgramInput.lstDeveloper.options[i].selected = true;
						else
							ProgramInput.lstDeveloper.options[i].selected = false;
						}
					//}				

                    var today = new Date();
				    if (txtQValue47.value=="")
                        txtQValue47.value = "3|" + today.getMonth() + "/" + today.getDate() + "/" + today.getFullYear() + "|" + (today.getMonth()+1) + "/" + today.getDate() + "/" + today.getFullYear() ;
                    
                    RangeArray = txtQValue47.value.split("|");
                    if (RangeArray[0] != "0" && RangeArray[0] != "1" && RangeArray[0] != "2" && RangeArray[0] != "3")
                        RangeArray[0] = "3";
                    
                    if (RangeArray.length==4)
                        {
                        ProgramInput.txtStartDate.value = RangeArray[1];
                        ProgramInput.txtEndDate.value = RangeArray[2];
                        ProgramInput.txtHistoryDays.value = RangeArray[3];
                        }
                    else
                        {
                        ProgramInput.txtStartDate.value = today.getMonth() + "/" + today.getDate() + "/" + today.getFullYear();
                        ProgramInput.txtEndDate.value = (today.getMonth()+1) + "/" + today.getDate() + "/" + today.getFullYear();
                        ProgramInput.txtHistoryDays.value = "";
                        }

                    ProgramInput.cboHistoryRange.selectedIndex = RangeArray[0];
                    cboHistoryRange_onchange();


					ProgramInput.txtSearch.value = txtQValue12.value;
					ProgramInput.txtTitle.value = txtQValue13.value;
					ProgramInput.txtNumbers.value = txtQValue14.value;

					if (txtQValue16.value != "")
					{
						ProgramInput.cboFormat.selectedIndex = txtQValue16.value;
					}				

					if (txtQValue17.value == "True")
						ProgramInput.chkNameSearch.checked = true;
					else
						ProgramInput.chkNameSearch.checked = false;
						
					if (txtQValue18.value == "True")
						ProgramInput.chkChangesSearch.checked = true;
					else
						ProgramInput.chkChangesSearch.checked = false;
						
					if (txtQValue19.value == "True")
						ProgramInput.chkDescriptionSearch.checked = true;
					else
						ProgramInput.chkDescriptionSearch.checked = false;

					if (txtQValue20.value == "True")
						ProgramInput.chkCommentsSearch.checked = true;
					else
						ProgramInput.chkCommentsSearch.checked = false;

					if (txtQValue21.value == "True")
						ProgramInput.chkDevelopment.checked = true;
					else
						ProgramInput.chkDevelopment.checked = false;

					if (txtQValue23.value != "")
					{
						ProgramInput.cboEOL.selectedIndex = txtQValue23.value;
					}				

					if (txtQValue24.value == "1")
						ProgramInput.chkSCRestricted.checked = true;
					else
						ProgramInput.chkSCRestricted.checked = false;



					if (txtQValue25.value != "")
					{
						ProgramInput.cboRohs.selectedIndex = txtQValue25.value;
					}				
					
					//Reset History Values
					ProgramInput.txtSpecificPilotStatus.value = "";
					ProgramInput.txtSpecificQualStatus.value = "";
					QualLink.innerHTML = "<a href='javascript:GetSpecificChange(2);'>All Changes</a>";
					PilotLink.innerHTML = "<a href='javascript:GetSpecificChange(1);'>All Changes</a>";
					ProgramInput.chkChangePilot.checked = false;
					ProgramInput.chkChangeQual.checked = false;
					
					if (txtQValue27.value != "")
					{
						HistoryArray = txtQValue27.value.split(";")
						for (i=0;i<HistoryArray.length;i++)
							{
							HistoryPartArray = HistoryArray[i].split("=");
							if(HistoryPartArray.length==2)
								{
								if (HistoryPartArray[0] == 21)
									{
									ProgramInput.txtSpecificQualStatus.value = HistoryPartArray[1];
									if (HistoryPartArray[1] != "")
										QualLink.innerHTML = "<a href='javascript:GetSpecificChange(2);'>Custom Filter</a>";
									ProgramInput.chkChangeQual.checked = true;
									}
								else if (HistoryPartArray[0] == 22)
									{
									ProgramInput.txtSpecificPilotStatus.value = HistoryPartArray[1];								
									if (HistoryPartArray[1] != "")
										PilotLink.innerHTML = "<a href='javascript:GetSpecificChange(1);'>Custom Filter</a>";
									ProgramInput.chkChangePilot.checked = true;
									}
								
								}
							}
					}

					if (txtQValue33.value == "True")
						ProgramInput.chkTest.checked = true;
					else
						ProgramInput.chkTest.checked = false;

					if (txtQValue34.value == "True")
						ProgramInput.chkRelease.checked = true;
					else
						ProgramInput.chkRelease.checked = false;

					if (txtQValue35.value == "True")
						ProgramInput.chkComplete.checked = true;
					else
						ProgramInput.chkComplete.checked = false;

					if (txtQValue36.value == "True")
						ProgramInput.chkTarget.checked = true;
					else
						ProgramInput.chkTarget.checked = false;

					if (txtQValue37.value == "True")
						ProgramInput.chkInImage.checked = true;
					else
						ProgramInput.chkInImage.checked = false;

					if (txtQValue38.value == "True")
						ProgramInput.chkFailed.checked = true;
					else
						ProgramInput.chkFailed.checked = false;

					for (i=0;i<ProgramInput.lstDevManager.length;i++)
						{
						if (instr("," + txtQValue39.value + "," , "," + ProgramInput.lstDevManager.options[i].value + "," ) > -1)
							ProgramInput.lstDevManager.options[i].selected = true;
						else
							ProgramInput.lstDevManager.options[i].selected = false;
						}
					for (i=0;i<ProgramInput.lstRoot.length;i++)
						{
						if (instr("," + txtQValue40.value + "," , "," + ProgramInput.lstRoot.options[i].value + "," ) > -1)
							ProgramInput.lstRoot.options[i].selected = true;
						else
							ProgramInput.lstRoot.options[i].selected = false;
						}

					for (i=0;i<ProgramInput.lstProductGroups.length;i++)
						{
						if (ProgramInput.lstProductGroups.options[i].value != "" && instr("," + txtQValue41.value + "," , "," + ProgramInput.lstProductGroups.options[i].value + "," ) > -1)
							ProgramInput.lstProductGroups.options[i].selected = true;
						else
							ProgramInput.lstProductGroups.options[i].selected = false;
						}
				
					for (i=0;i<ProgramInput.lstQualStatus.length;i++)
						{
						if (instr("," + txtQValue42.value + "," , "," + ProgramInput.lstQualStatus.options[i].value + "," ) > -1)
							ProgramInput.lstQualStatus.options[i].selected = true;
						else
							ProgramInput.lstQualStatus.options[i].selected = false;
						}
					ProgramInput.txtAdvanced.value = txtQValue43.value;
					
					if (txtQValue45.value != "")
						{
						ProgramInput.lstColumns.options.length = 0;
			
			            txtQValue45.value = AppendNewColumns(txtQValue45.value,ProgramInput.txtMasterColumns.value);

						ResultArray = txtQValue45.value.split(",");
						for (i=0;i<ResultArray.length;i=i+2)
							{
							ProgramInput.lstColumns.options[ProgramInput.lstColumns.length] = new Option(ResultArray[i+1],ResultArray[i+1]);			
							if (ResultArray[i]=="1")
								ProgramInput.lstColumns.options[ProgramInput.lstColumns.length-1].selected=true;
							}
						}

					for (i=0;i<ProgramInput.lstCoreTeam.length;i++)
						{
						if (instr("," + txtQValue46.value + "," , "," + ProgramInput.lstCoreTeam.options[i].value + "," ) > -1)
							ProgramInput.lstCoreTeam.options[i].selected = true;
						else
							ProgramInput.lstCoreTeam.options[i].selected = false;
						}
						
					if (txtAdmin.value != "1" )
						{
						if (txtQValue43.value == "")
							AdminCriteriaRow.style.display = "none";
						else
							{
							AdminCriteria.innerHTML = "<FIELDSET><LEGEND>The following criteria has been assigned to this profile by an administrator:</LEGEND>" + txtQValue43.value + "</FIELDSET>";
							AdminCriteriaRow.style.display = "";
							}
						}
                                                                  //list of Pulsar products
                    for (i=0;i<ProgramInput.lstProductsPulsar.length;i++)
						{
						if (instr("," + txtQValue55.value + "," , "," + ProgramInput.lstProductsPulsar.options[i].value + "," ) > -1)
							ProgramInput.lstProductsPulsar.options[i].selected = true;
						else
							ProgramInput.lstProductsPulsar.options[i].selected = false;
						}
		}
}

var strNewName;

function RenameProfile(){
	strNewName = window.prompt("Enter new name for this filter.",ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex].text);
	
	if (strNewName != null)
		{
		//var objRS = RSGetASPObject("DelRSupdate.asp");
		//var objResult = objRS.updateProfile(ProgramInput.cboProfile.value,1,strNewName,"","");
	
		jsrsExecute("DelRSupdate.asp", myCallback2, "ProfileStrings", Array(ProgramInput.cboProfile.value,"1",strNewName,"","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""));			
		}
}

function myCallback2( returnstring ){
	if (returnstring != 1)
		window.alert("Unable to update this filter."); 
	else
		{
		ProgramInput.cboProfile.options[ProgramInput.cboProfile.selectedIndex].text = strNewName;
		try{
			if (window.opener.name == "HardwareMatrix")
				window.opener.location.reload();

	  	   }
			catch(e)
				{
				
				}		
		}
}


function UpdateProfile(){
	var i;
	var strProduct = "";
	var strOS = "";
	var strLanguage = "";
	var strVendor = "";
	var strCategory = "";
	var strDeveloper = "";
	var strDevManager = "";
	var strRoot = "";
	var strProductGroup = "";
	var strQualStatus = "";
	var strColumnList = "";
	var strFilter = "";
	var strCommodityPM = "";
	var strHistoryFilter = "";
	var strRestricted = "";
	var strCoreTeamList = "";
	var strDateRange = "";
	var strProductPulsar = "";
	
	    //Product Legacy List
		for(i=0;i<ProgramInput.lstProducts.length;i++)
			if (ProgramInput.lstProducts.options[i].selected)
				strProduct = strProduct + "," + ProgramInput.lstProducts.options[i].value;
		        if (strProduct.length>0)
			        {
			        strProduct = strProduct.substr(1);
			        strFilter = "&lstProducts=" + strProduct;
		        }
   
         //Product Pulsar List
		 for(i=0;i<ProgramInput.lstProductsPulsar.length;i++)
		     if (ProgramInput.lstProductsPulsar.options[i].selected)
		                strProductPulsar = strProductPulsar + "," + ProgramInput.lstProductsPulsar.options[i].value;
		        if (strProductPulsar.length>0)
		        {
		            strProductPulsar = strProductPulsar.substr(1);
		            strFilter = strFilter + "&lstProductsPulsar=" + strProductPulsar;
		        }

		//OS List
		for(i=0;i<ProgramInput.lstOS.length;i++)
			if (ProgramInput.lstOS.options[i].selected)
				strOS = strOS + "," + ProgramInput.lstOS.options[i].value;
		if (strOS.length>0)
			{
			strOS = strOS.substr(1);
			strFilter = strFilter + "&lstOS=" + strOS;
			}
		//Language List
		for(i=0;i<ProgramInput.lstLanguage.length;i++)
			if (ProgramInput.lstLanguage.options[i].selected)
				strLanguage = strLanguage + "," + ProgramInput.lstLanguage.options[i].value;
		if (strLanguage.length>0)
			{
			strLanguage = strLanguage.substr(1);
			strFilter = strFilter + "&lstLanguage=" + strLanguage;
			}
	
		//Vendor List
		for(i=0;i<ProgramInput.lstVendor.length;i++)
			if (ProgramInput.lstVendor.options[i].selected)
				strVendor = strVendor + "," + ProgramInput.lstVendor.options[i].value;
		if (strVendor.length>0)
			{
			strVendor = strVendor.substr(1);
			strFilter = strFilter + "&lstVendor=" + strVendor;
			}
		//Category List
		for(i=0;i<ProgramInput.lstCategory.length;i++)
			if (ProgramInput.lstCategory.options[i].selected)
				strCategory = strCategory + "," + ProgramInput.lstCategory.options[i].value;
		if (strCategory.length>0)
			{
			strCategory = strCategory.substr(1);
			strFilter = strFilter + "&lstCategory=" + strCategory;
			}

		//CoreTeam
		for(i=0;i<ProgramInput.lstCoreTeam.length;i++)
			if (ProgramInput.lstCoreTeam.options[i].selected)
				strCoreTeamList = strCoreTeamList + "," + ProgramInput.lstCoreTeam.options[i].value;
		if (strCoreTeamList.length>0)
			{
			strCoreTeamList = strCoreTeamList.substr(1);
			strFilter = strFilter + "&lstCoreTeam=" + strCoreTeamList;
			}
	
	
		//Developer List
		for(i=0;i<ProgramInput.lstDeveloper.length;i++)
			if (ProgramInput.lstDeveloper.options[i].selected)
				strDeveloper = strDeveloper + "," + ProgramInput.lstDeveloper.options[i].value;
		if (strDeveloper.length>0)
			{
			strDeveloper = strDeveloper.substr(1);
			strFilter = strFilter + "&lstDeveloper=" + strDeveloper;
			}
		//DevManager List
		for(i=0;i<ProgramInput.lstDevManager.length;i++)
			if (ProgramInput.lstDevManager.options[i].selected)
				strDevManager = strDevManager + "," + ProgramInput.lstDevManager.options[i].value;
		if (strDevManager.length>0)
			{
			strDevManager = strDevManager.substr(1);
			strFilter = strFilter + "&lstDevManager=" + strDevManager;
			}
		
		//Column List
		strColumnList=""
		for(i=0;i<ProgramInput.lstColumns.length;i++)
			{
			if (ProgramInput.lstColumns.options[i].selected)
				strColumnList = strColumnList + ",1," + ProgramInput.lstColumns.options[i].text;
			else
				strColumnList = strColumnList + ",0," + ProgramInput.lstColumns.options[i].text;
			}
		if (strColumnList.length>0)
			{
			strColumnList = strColumnList.substr(1);
			}			
	
		//Root List
		for(i=0;i<ProgramInput.lstRoot.length;i++)
			if (ProgramInput.lstRoot.options[i].selected)
				strRoot = strRoot + "," + ProgramInput.lstRoot.options[i].value;
		if (strRoot.length>0)
			{
			strRoot = strRoot.substr(1);
			strFilter = strFilter + "&lstRoot=" + strRoot;
			}
		//ProductGroup List
		for(i=0;i<ProgramInput.lstProductGroups.length;i++)
			if (ProgramInput.lstProductGroups.options[i].selected)
				strProductGroup = strProductGroup + "," + ProgramInput.lstProductGroups.options[i].value;
		if (strProductGroup.length>0)
			{
			strProductGroup = strProductGroup.substr(1);
			strFilter = strFilter + "&lstProductGroups=" + strProductGroup;
			}
		//QualStatus List
		for(i=0;i<ProgramInput.lstQualStatus.length;i++)
			if (ProgramInput.lstQualStatus.options[i].selected)
				strQualStatus = strQualStatus + "," + ProgramInput.lstQualStatus.options[i].value;
		if (strQualStatus.length>0)
			{
			strQualStatus = strQualStatus.substr(1);
			strFilter = strFilter + "&lstQualStatus=" + strQualStatus;
			}

    ProgramInput.txtStartDate.value = ProgramInput.txtStartDate.value.replace("|","_")			
    ProgramInput.txtEndDate.value = ProgramInput.txtEndDate.value.replace("|","_")			
    ProgramInput.txtHistoryDays.value = ProgramInput.txtHistoryDays.value.replace("|","_")			
    strDateRange = ProgramInput.cboHistoryRange.selectedIndex + "|" + ProgramInput.txtStartDate.value + "|" + ProgramInput.txtEndDate.value + "|" + ProgramInput.txtHistoryDays.value;

	if (ProgramInput.chkSCRestricted.checked)
		strRestricted="1"
	else
		strRestricted="0"


	if (ProgramInput.lstCommodityPM.selectedIndex > 0)
		strFilter = strFilter + "&lstCommodityPM=" + ProgramInput.lstCommodityPM.options[ProgramInput.lstCommodityPM.selectedIndex].value;

	if (ProgramInput.txtTitle.value != "")
		strFilter = strFilter + "&txtTitle=" + ProgramInput.txtTitle.value;

    if (ProgramInput.cboHistoryRange.selectedIndex == 3)
        strFilter = strFilter + "&cboHistoryRange=Range";
    else
        strFilter = strFilter + "&cboHistoryRange=" + ProgramInput.cboHistoryRange.options[ProgramInput.cboHistoryRange.selectedIndex].value;

	if (ProgramInput.txtStartDate.value != "")
		strFilter = strFilter + "&txtStartDate=" + ProgramInput.txtStartDate.value;
	if (ProgramInput.txtEndDate.value != "")
		strFilter = strFilter + "&txtEndDate=" + ProgramInput.txtEndDate.value;
	if (ProgramInput.txtStartDate.value != "")
		strFilter = strFilter + "&txtHistoryDays=" + ProgramInput.txtHistoryDays.value;
	if (ProgramInput.txtSpecificPilotStatus.value != "")
		strFilter = strFilter + "&txtSpecificPilotStatus=" + ProgramInput.txtSpecificPilotStatus.value;
	if (ProgramInput.txtSpecificQualStatus.value != "")
		strFilter = strFilter + "&txtSpecificQualStatus=" + ProgramInput.txtSpecificQualStatus.value;
	if (ProgramInput.chkChangeType[0].checked != "")
		strFilter = strFilter + "&chkChangeType=" + ProgramInput.chkChangeType[0].value;
	if (ProgramInput.chkChangeType[1].checked != "")
		strFilter = strFilter + "&chkChangeType=" + ProgramInput.chkChangeType[1].value;
	
	if (ProgramInput.chkSCRestricted.checked)
		strFilter = strFilter + "&chkSCRestricted=1";


	if (ProgramInput.txtNumbers.value != "")
		strFilter = strFilter + "&txtNumbers=" + ProgramInput.txtNumbers.value;

	if (ProgramInput.cboEOL.selectedIndex > 0)
		strFilter = strFilter + "&cboEOL=" + ProgramInput.cboEOL.options[ProgramInput.cboEOL.selectedIndex].value;

	if (ProgramInput.cboRohs.selectedIndex > 0)
		strFilter = strFilter + "&cboRohs=" + ProgramInput.cboRohs.options[ProgramInput.cboRohs.selectedIndex].value;

	if (ProgramInput.ReportFormat.selectedIndex > 0)
		strFilter = strFilter + "&ReportFormat=" + ProgramInput.ReportFormat.options[ProgramInput.ReportFormat.selectedIndex].value;


		if (ProgramInput.lstCommodityPM.selectedIndex==-1)
			strCommodityPM = "0"
		else
			strCommodityPM = ProgramInput.lstCommodityPM.options[ProgramInput.lstCommodityPM.selectedIndex].value;


	strHistoryFilter = GetFilterString();
		
	if (strFilter.length>0)
		strFilter = strFilter.substr(1);

	//Update
	if(window.confirm("Are you sure you want to update this filter?"))
		{
		//var objRS = RSGetASPObject("DelRSupdate.asp");
		//var objResult = objRS.updateProfile(ProgramInput.cboProfile.value,3,"",strProduct,strLanguage,strVendor,strCategory,0,"",0,0,"","","",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,strOS,ProgramInput.cboFormat.selectedIndex,ProgramInput.chkNameSearch.checked,ProgramInput.chkChangesSearch.checked,ProgramInput.chkDescriptionSearch.checked,ProgramInput.chkCommentsSearch.checked,ProgramInput.chkDevelopment.checked,0,"",0,"",strDeveloper,"","","","","","",ProgramInput.chkTest.checked,ProgramInput.chkRelease.checked,ProgramInput.chkComplete.checked,ProgramInput.chkTarget.checked,ProgramInput.chkInImage.checked,ProgramInput.chkFailed.checked);
	    jsrsExecute("DelRSupdate.asp", myCallback3, "ProfileStrings", Array(ProgramInput.cboProfile.value,"3","",strProduct,strLanguage,strVendor,strCategory,"0",ProgramInput.ReportFormat.selectedIndex.toString(),strCommodityPM,txtReportType.value,"","","",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,strOS,ProgramInput.cboFormat.selectedIndex.toString(),ProgramInput.chkNameSearch.checked.toString(),ProgramInput.chkChangesSearch.checked.toString(),ProgramInput.chkDescriptionSearch.checked.toString(),ProgramInput.chkCommentsSearch.checked.toString(),ProgramInput.chkDevelopment.checked.toString(),"0",ProgramInput.cboEOL.selectedIndex.toString(),strRestricted,ProgramInput.cboRohs.selectedIndex.toString(),strDeveloper,strHistoryFilter,"","","","","",ProgramInput.chkTest.checked.toString(),ProgramInput.chkRelease.checked.toString(),ProgramInput.chkComplete.checked.toString(),ProgramInput.chkTarget.checked.toString(),ProgramInput.chkInImage.checked.toString(),ProgramInput.chkFailed.checked.toString(),strDevManager,strRoot,strProductGroup,strQualStatus,ProgramInput.txtAdvanced.value,"0",strColumnList,strCoreTeamList,strDateRange,"","","",strFilter,strProductPulsar));			
		}
}

function myCallback3( returnstring ){
	if (returnstring != 1)
		window.alert("Unable to update this filter."); 
	else
		{
		window.alert("Update Complete");
		try{
			if (window.opener.name == "HardwareMatrix")
				window.opener.location.reload();

	  	   }
			catch(e)
				{
				
				}		
		}
}

function cboProfile_onkeydown() {
	//if(event.keyCode == 46 && DetailsForm.cboProfile.selectedIndex > 2)
	//	window.alert("Deleting");
}

function AddProfile(){
	var i; //Save the Date range in fields 27 and 28
	var strProduct = "";
	var strOS = "";
	var strLanguage = "";
	var strVendor = "";
	var strCategory = "";
	var strDeveloper = "";
	var strDevManager = "";
	var strRoot = "";
	var strProductGroup = "";
	var strQualStatus = "";
	var strFilter = "";
	var strCommodityPM = "";
	var strHistoryFilter = "";
	var strRestricted = "";
	var strColumnList ="";	
	var strCoreTeamList = "";
	var strDateRange = "";
	var strProductPulsar =  "";

	strNewName = window.prompt("Enter a name for the new filter.","");
	if (strNewName != null)
		{
	
		//Product List
		for(i=0;i<ProgramInput.lstProducts.length;i++)
			if (ProgramInput.lstProducts.options[i].selected)
				strProduct = strProduct + "," + ProgramInput.lstProducts.options[i].value;
		if (strProduct.length>0)
			{
			strProduct = strProduct.substr(1);
			strFilter = "&lstProducts=" + strProduct;

		}

	    //Product Pulsar List
		for(i=0;i<ProgramInput.lstProductsPulsar.length;i++)
		    if (ProgramInput.lstProductsPulsar.options[i].selected)
		        strProductPulsar = strProductPulsar + "," + ProgramInput.lstProductsPulsar.options[i].value;
		if (strProductPulsar.length>0)
		{
		    strProductPulsar = strProductPulsar.substr(1);
		    strFilter = strFilter + "&lstProductsPulsar=" + strProductPulsar;
		}

		//OS List
		for(i=0;i<ProgramInput.lstOS.length;i++)
			if (ProgramInput.lstOS.options[i].selected)
				strOS = strOS + "," + ProgramInput.lstOS.options[i].value;
		if (strOS.length>0)
			{
			strOS = strOS.substr(1);
			strFilter = strFilter + "&lstOS=" + strOS;
			}
	
		//Language List
		for(i=0;i<ProgramInput.lstLanguage.length;i++)
			if (ProgramInput.lstLanguage.options[i].selected)
				strLanguage = strLanguage + "," + ProgramInput.lstLanguage.options[i].value;
		if (strLanguage.length>0)
			{
			strLanguage = strLanguage.substr(1);
			strFilter = strFilter + "&lstLanguage=" + strLanguage;
			}
	
		//Vendor List
		for(i=0;i<ProgramInput.lstVendor.length;i++)
			if (ProgramInput.lstVendor.options[i].selected)
				strVendor = strVendor + "," + ProgramInput.lstVendor.options[i].value;
		if (strVendor.length>0)
			{
			strVendor = strVendor.substr(1);
			strFilter = strFilter + "&lstVendor=" + strVendor;
			}

		//Category List
		for(i=0;i<ProgramInput.lstCategory.length;i++)
			if (ProgramInput.lstCategory.options[i].selected)
				strCategory = strCategory + "," + ProgramInput.lstCategory.options[i].value;
		if (strCategory.length>0)
			{		
			strCategory = strCategory.substr(1);
			strFilter = strFilter + "&lstCategory=" + strCategory;
			}

		//CoreTeam
		for(i=0;i<ProgramInput.lstCoreTeam.length;i++)
			if (ProgramInput.lstCoreTeam.options[i].selected)
				strCoreTeamList = strCoreTeamList + "," + ProgramInput.lstCoreTeam.options[i].value;
		if (strCoreTeamList.length>0)
			{
			strCoreTeamList = strCoreTeamList.substr(1);
			strFilter = strFilter + "&lstCoreTeam=" + strCoreTeamList;
			}
		
		//Developer List
		for(i=0;i<ProgramInput.lstDeveloper.length;i++)
			if (ProgramInput.lstDeveloper.options[i].selected)
				strDeveloper = strDeveloper + "," + ProgramInput.lstDeveloper.options[i].value;
		if (strDeveloper.length>0)
			{
			strDeveloper = strDeveloper.substr(1);
			strFilter = strFilter + "&lstDeveloper=" + strDeveloper;
			}
	
		//DevManager List
		for(i=0;i<ProgramInput.lstDevManager.length;i++)
			if (ProgramInput.lstDevManager.options[i].selected)
				strDevManager = strDevManager + "," + ProgramInput.lstDevManager.options[i].value;
		if (strDevManager.length>0)
			{
			strDevManager = strDevManager.substr(1);
			strFilter = strFilter + "&lstDevManager=" + strDevManager;
			}

		//Column List
		for(i=0;i<ProgramInput.lstColumns.length;i++)
			{
			if (ProgramInput.lstColumns.options[i].selected)
				strColumnList = strColumnList + ",1," + ProgramInput.lstColumns.options[i].text;
			else
				strColumnList = strColumnList + ",0," + ProgramInput.lstColumns.options[i].text;
			}
			
		if (strColumnList.length>0)
			{
			strColumnList = strColumnList.substr(1);
			}			

		//Root List
		for(i=0;i<ProgramInput.lstRoot.length;i++)
			if (ProgramInput.lstRoot.options[i].selected)
				strRoot = strRoot + "," + ProgramInput.lstRoot.options[i].value;
		if (strRoot.length>0)
			{
			strRoot = strRoot.substr(1);
			strFilter = strFilter + "&lstRoot=" + strRoot;
			}

		//ProductGroup List
		for(i=0;i<ProgramInput.lstProductGroups.length;i++)
			if (ProgramInput.lstProductGroups.options[i].selected)
				strProductGroup = strProductGroup + "," + ProgramInput.lstProductGroups.options[i].value;
		if (strProductGroup.length>0)
			{
			strProductGroup = strProductGroup.substr(1);
			strFilter = strFilter + "&lstProductGroups=" + strProductGroup;
			}

		//QualStatus List
		for(i=0;i<ProgramInput.lstQualStatus.length;i++)
			if (ProgramInput.lstQualStatus.options[i].selected)
				strQualStatus = strQualStatus + "," + ProgramInput.lstQualStatus.options[i].value;
		if (strQualStatus.length>0)
			{
			strQualStatus = strQualStatus.substr(1);
			strFilter = strFilter + "&lstQualStatus=" + strQualStatus;
			}

		if (ProgramInput.chkSCRestricted.checked)
			strRestricted="1"
		else
			strRestricted="0"

		if (ProgramInput.lstCommodityPM.selectedIndex > 0)
			strFilter = strFilter + "&lstCommodityPM=" + ProgramInput.lstCommodityPM.options[ProgramInput.lstCommodityPM.selectedIndex].value;

		if (ProgramInput.txtTitle.value != "")
			strFilter = strFilter + "&txtTitle=" + ProgramInput.txtTitle.value;

        if (ProgramInput.cboHistoryRange.selectedIndex == 3)
            strFilter = strFilter + "&cboHistoryRange=Range";
        else
            strFilter = strFilter + "&cboHistoryRange=" + ProgramInput.cboHistoryRange.options[ProgramInput.cboHistoryRange.selectedIndex].value;

		if (ProgramInput.txtStartDate.value != "")
			strFilter = strFilter + "&txtStartDate=" + ProgramInput.txtStartDate.value;
		if (ProgramInput.txtEndDate.value != "")
			strFilter = strFilter + "&txtEndDate=" + ProgramInput.txtEndDate.value;
    	if (ProgramInput.txtStartDate.value != "")
	    	strFilter = strFilter + "&txtHistoryDays=" + ProgramInput.txtHistoryDays.value;
		if (ProgramInput.txtSpecificPilotStatus.value != "")
			strFilter = strFilter + "&txtSpecificPilotStatus=" + ProgramInput.txtSpecificPilotStatus.value;
		if (ProgramInput.txtSpecificQualStatus.value != "")
			strFilter = strFilter + "&txtSpecificQualStatus=" + ProgramInput.txtSpecificQualStatus.value;
		if (ProgramInput.chkChangeType[0].checked != "")
			strFilter = strFilter + "&chkChangeType=" + ProgramInput.chkChangeType[0].value;
		if (ProgramInput.chkChangeType[1].checked != "")
			strFilter = strFilter + "&chkChangeType=" + ProgramInput.chkChangeType[1].value;

		if (ProgramInput.txtNumbers.value != "")
			strFilter = strFilter + "&txtNumbers=" + ProgramInput.txtNumbers.value;

		if (ProgramInput.cboEOL.selectedIndex > 0)
			strFilter = strFilter + "&cboEOL=" + ProgramInput.cboEOL.options[ProgramInput.cboEOL.selectedIndex].value;

		if (ProgramInput.cboRohs.selectedIndex > 0)
			strFilter = strFilter + "&cboRohs=" + ProgramInput.cboRohs.options[ProgramInput.cboRohs.selectedIndex].value;

		if (ProgramInput.ReportFormat.selectedIndex > 0)
			strFilter = strFilter + "&ReportFormat=" + ProgramInput.ReportFormat.options[ProgramInput.ReportFormat.selectedIndex].value;

		if (ProgramInput.chkSCRestricted.checked)
			strFilter = strFilter + "&chkSCRestricted=1";

		strHistoryFilter = GetFilterString();

		
		if (strFilter.length>0)
			strFilter = strFilter.substr(1);

		if (ProgramInput.lstCommodityPM.selectedIndex==-1)
			strCommodityPM = "0"
		else
			strCommodityPM = ProgramInput.lstCommodityPM.options[ProgramInput.lstCommodityPM.selectedIndex].value;

        ProgramInput.txtStartDate.value = ProgramInput.txtStartDate.value.replace("|","_")			
        ProgramInput.txtEndDate.value = ProgramInput.txtEndDate.value.replace("|","_")			
        ProgramInput.txtHistoryDays.value = ProgramInput.txtHistoryDays.value.replace("|","_")			
        strDateRange = ProgramInput.cboHistoryRange.selectedIndex + "|" + ProgramInput.txtStartDate.value + "|" + ProgramInput.txtEndDate.value + "|" + ProgramInput.txtHistoryDays.value;

		//Add
		
//		var objRS = RSGetASPObject("DelRSupdate.asp");
//		var objResult = objRS.updateProfile(ProgramInput.cboProfile.value,4,strNewName,strProduct,strLanguage,strVendor,strCategory,0,"",0,0,"","","",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,strOS,ProgramInput.cboFormat.selectedIndex,ProgramInput.chkNameSearch.checked,ProgramInput.chkChangesSearch.checked,ProgramInput.chkDescriptionSearch.checked,ProgramInput.chkCommentsSearch.checked,ProgramInput.chkDevelopment.checked,0,"",0,"",strDeveloper,"","","","","","",ProgramInput.chkTest.checked,ProgramInput.chkRelease.checked,ProgramInput.chkComplete.checked,ProgramInput.chkTarget.checked,ProgramInput.chkInImage.checked,ProgramInput.chkFailed.checked,3,txtUserID.value);

        jsrsExecute("DelRSupdate.asp", myCallback4, "ProfileStrings", Array(ProgramInput.cboProfile.value,"4",strNewName,strProduct,strLanguage,strVendor,strCategory,"0",ProgramInput.ReportFormat.selectedIndex.toString(),strCommodityPM,txtReportType.value,"","","",ProgramInput.txtSearch.value,ProgramInput.txtTitle.value,ProgramInput.txtNumbers.value,strOS,ProgramInput.cboFormat.selectedIndex.toString(),ProgramInput.chkNameSearch.checked.toString(),ProgramInput.chkChangesSearch.checked.toString(),ProgramInput.chkDescriptionSearch.checked.toString(),ProgramInput.chkCommentsSearch.checked.toString(),ProgramInput.chkDevelopment.checked.toString(),"0",ProgramInput.cboEOL.selectedIndex.toString(),strRestricted,ProgramInput.cboRohs.selectedIndex.toString(),strDeveloper,	strHistoryFilter,"","","","","",ProgramInput.chkTest.checked.toString(),ProgramInput.chkRelease.checked.toString(),ProgramInput.chkComplete.checked.toString(),ProgramInput.chkTarget.checked.toString(),ProgramInput.chkInImage.checked.toString(),ProgramInput.chkFailed.checked.toString(),strDevManager,strRoot,strProductGroup,strQualStatus,ProgramInput.txtAdvanced.value,"0",strColumnList,strCoreTeamList,strDateRange,"3",txtUserID.value,"",strFilter,strProductPulsar));			

		}
}	

function myCallback4( returnstring ){
	if (returnstring == "" || returnstring == 0 )
		window.alert("Unable to add this filter."); 
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

		window.alert("Filter Added.");
		
		try{
			if (window.opener.name == "HardwareMatrix")
				window.opener.location.reload();

	  	   }
			catch(e)
				{
				
				}		
		
		}
}


function DeleteProfile(){
	if(window.confirm("Are you sure you want to delete this filter?"))
		jsrsExecute("DelRSupdate.asp", myCallback5, "ProfileStrings", Array(ProgramInput.cboProfile.value,"2","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""));			
}

function myCallback5( returnstring ){
	if (returnstring != 1)
		window.alert("Unable to delete this filter."); 
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
		try{
			if (window.opener.name == "HardwareMatrix")
				window.opener.location.reload();

	  	   }
			catch(e)
				{
				
				}		
		}
}


function ActionCell_onmouseover(ActionNumber) {
	if(oPopup.isOpen) 
		oPopup.hide();
	window.focus();
//	ReportCell.style.color = "white";	
//	ReportCell.style.background="#333333";

	StatusCell.style.color = "white";	
	StatusCell.style.background="#333333";

//	if(ActionNumber==1)
//		cmdMail_onmouseover();
//	else if(ActionNumber==2)
//		cmdGraph_onmouseover();
		

	window.event.srcElement.style.background="gainsboro";
	window.event.srcElement.style.cursor = "hand";
	window.event.srcElement.style.color = "black";	
}

function ActionCell_onmouseout(ActionNumber) {
//	if(ActionNumber==1)
//		cmdMail_onmouseout();
//	else if(ActionNumber==2)
//		cmdGraph_onmouseout();
		

	window.event.srcElement.style.color = "white";	
	window.event.srcElement.style.background="#333333";
}

function CommodityCell_onmouseover() {
	if(oPopup.isOpen) 
		oPopup.hide();
	window.focus();

	StatusCell.style.color = "white";	
	StatusCell.style.background="#333333";


	window.event.srcElement.style.background="gainsboro";
	window.event.srcElement.style.cursor = "hand";
	window.event.srcElement.style.color = "black";	
}

function CommodityCell_onmouseout() {
	window.event.srcElement.style.color = "white";	
	window.event.srcElement.style.background="#333333";
}



function MenuCell_onmouseout() {
	//window.event.srcElement.style.color = "white";	
	//window.event.srcElement.style.background="#333333";
}

function MenuCell_onmouseover(MenuNumber) {
	window.focus();
	if (MenuNumber==1)
		{
		StatusCell.style.color = "white";	
		StatusCell.style.background="#333333";
		//cmdSummary_onmouseout();
		}

	
	window.event.srcElement.style.background="gainsboro";
	window.event.srcElement.style.color = "black";
	ShowMenu(MenuNumber);
}

function ShowMenu(MenuNumber) {
    var lefter = event.clientX - event.offsetX;
    var topper = (event.clientY - event.offsetY)+ event.srcElement.offsetHeight;
    var popupBody;
    
    if (MenuNumber==1)
		{
		popupBody = "<DIV STYLE=\"BORDER-RIGHT: black 1px solid; BORDER-TOP: black 1px solid; LEFT: 0px; BORDER-LEFT: black 1px solid; BORDER-BOTTOM: black 1px solid; POSITION: relative; TOP: 0px\">";

		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:cmdReport_onclick(3)'\" >&nbsp;&nbsp;&nbsp;All&nbsp;Observations</SPAN></FONT></DIV>";


		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:cmdReport_onclick(4)'\" >&nbsp;&nbsp;&nbsp;P0/P1&nbsp;Observations&nbsp;Only&nbsp;</SPAN></FONT></DIV>";
        
		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:cmdReport_onclick(2015)'\" >&nbsp;&nbsp;&nbsp;MDA&nbsp;Compliance&nbsp;2015</SPAN></FONT></DIV>";

		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN><HR width=95%></SPAN></DIV>";

		popupBody = popupBody + "<DIV onmouseover=\"this.style.background='RoyalBlue';this.style.cursor='hand';this.style.color='white'\" onmouseout=\"this.style.background='white';this.style.color='black'\">";
		popupBody = popupBody + "<FONT face=Arial size=2>";
		popupBody = popupBody + "<SPAN onclick=\"parent.location.href='javascript:cmdReport_onclick(5)'\" >&nbsp;&nbsp;&nbsp;MDA&nbsp;Compliance</SPAN></FONT></DIV>";

		popupBody = popupBody + "<DIV>";
		popupBody = popupBody + "<SPAN>&nbsp;</SPAN></DIV>";

		popupBody = popupBody + "</DIV>";
		oPopup.document.body.innerHTML = popupBody; 

		oPopup.show(lefter, topper, 130, 85, document.body);

		//Adjust window size
		if (oPopup.document.body.scrollHeight> 1 || oPopup.document.body.scrollWidth> 1)
			{
			NewHeight = oPopup.document.body.scrollHeight;
			NewWidth = oPopup.document.body.scrollWidth;
			oPopup.show(lefter, topper, NewWidth, NewHeight, document.body);
			}
		}
}

function window_onblur() {
	if(oPopup.isOpen) 
		oPopup.hide();

	StatusCell.style.color = "white";	
	StatusCell.style.background="#333333";

}

function window_onload() {
	lblLoad.style.display = "none";
	lblInst.style.display = "";
	AdjustPulsarProductsDropdownWidth();	
	
}

function cmdDate_onclick(FieldID) {
	var strID;
	var oldValue;
	
	if (FieldID==1)
		oldValue = ProgramInput.txtStartDate.value;
	else if (FieldID==2)
		oldValue = ProgramInput.txtEndDate.value;
	else if (FieldID==3)
		oldValue = ProgramInput.txtCompleteDateStart.value;
	else if (FieldID==4)
		oldValue = ProgramInput.txtCompleteDateEnd.value;
		
	strID = window.showModalDialog("../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strID) == "undefined")
		return;
	
	if (FieldID==1)
		ProgramInput.txtStartDate.value = strID;
	else if (FieldID==2)
		ProgramInput.txtEndDate.value = strID;
	else if (FieldID==3)
		ProgramInput.txtCompleteDateStart.value = strID;
	else if (FieldID==4)
		ProgramInput.txtCompleteDateEnd.value = strID;
}


function GetSpecificChange(TypeID){
	var strResult;
	
	if (TypeID==1)
		{
		strResult = window.showModalDialog("ChooseSpecificChange.asp?TypeID=1&Current=" + ProgramInput.txtSpecificPilotStatus.value,"","dialogWidth:450px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: No;status: No") 
		if (typeof(strResult) != "undefined") 
			{
			ProgramInput.txtSpecificPilotStatus.value = strResult[0];
			PilotLink.innerHTML=strResult[1];
			if (strResult[0].length > 0)
				ProgramInput.chkChangePilot.checked = true;
				
			}
		}
	else
		{
		strResult = window.showModalDialog("ChooseSpecificChange.asp?TypeID=2&Current=" + ProgramInput.txtSpecificQualStatus.value ,"","dialogWidth:450px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: No;status: No") 
		if (typeof(strResult) != "undefined") 
			{
			ProgramInput.txtSpecificQualStatus.value = strResult[0];
			QualLink.innerHTML=strResult[1];
			if (strResult[0].length > 0)
				ProgramInput.chkChangeQual.checked = true;
			}
		}
	
}

function GetFilterString (){
	var strField = "";
	
	if (ProgramInput.chkChangeQual.checked)
		strField = "21=" + ProgramInput.txtSpecificQualStatus.value;
	
	if (ProgramInput.chkChangePilot.checked)
		{
		if (strField != "")
			strField = strField + ";"
		strField = strField + "22=" + ProgramInput.txtSpecificPilotStatus.value;
		}
	return strField;
}


function cboHistoryRange_onchange() {
	if (ProgramInput.cboHistoryRange.selectedIndex ==3)
		{
		spnHistoryCount.style.display="none";
		spnHistoryRange.style.display="";
		}
	else
		{
		spnHistoryCount.style.display="";
		spnHistoryRange.style.display="none";
		}
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

/*function ResetColumns(){
	var strResult;
	var strColumns="ID,Name,Version,Vendor,Vendor Version,Part Number,Model,Category,Developer,Dev Manager,Software EOL,Factory EOA,Service EOA,Workflow,Product,Targeted,In Image,Softpaq,WHQL,OEM Ready,TTS,HW Qual Status,Code Name,Pilot Status,MIT Signoff,ODM Signoff,WWAN Signoff,Dev Signoff,MIT Samples,ODM Samples,WWAN Samples";
	var ResultArray;
	var i;
	if(window.confirm("Are you sure you want to delete your custom column list?"))
		{
		strResult = window.showModalDialog("ReorderColumnsSave.asp?optDefaultList=2&txtUserSettingsID=3&txtEmployeeID=" + txtUserID.value,"","dialogWidth:500px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
			if (typeof(strResult) != "undefined")
				{
				ProgramInput.lstColumns.options.length = 0;
			
				ResultArray = strColumns.split(",");
				for (i=0;i<ResultArray.length;i++)
					ProgramInput.lstColumns.options[ProgramInput.lstColumns.length] = new Option(ResultArray[i],ResultArray[i]);			
				}
		}
}
*/
function ReorderColumns(){
	var strResult;
	var strColumns="";
	var ResultArray;
	var i;
	var strSelected="";
	
	for (i=0;i<ProgramInput.lstColumns.length;i++)
		{
		if (ProgramInput.lstColumns.options[i].selected)
				strSelected = strSelected + "," + ProgramInput.lstColumns.options[i].text;			
		strSelected = strSelected + ",";
		
		if (strColumns=="")
			strColumns = ProgramInput.lstColumns.options[i].text;
		else
			strColumns = strColumns + "," + ProgramInput.lstColumns.options[i].text;
		}
	strResult = window.showModalDialog("ReorderColumns.asp?UserSettingsID=3&lstColumns=" + strColumns,"","dialogWidth:800px;dialogHeight:500px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strResult) != "undefined")
			{
			if (strResult.length > 0 )
				{
				ProgramInput.lstColumns.options.length = 0;
			
				ResultArray = strResult.split(",");
				for (i=0;i<ResultArray.length;i++)
					if (strSelected.indexOf("," + ResultArray[i] + ",")>-1)
						{
						ProgramInput.lstColumns.options[ProgramInput.lstColumns.length] = new Option(ResultArray[i],ResultArray[i]);			
						ProgramInput.lstColumns.options[i].selected=true;			
						}
					else
						ProgramInput.lstColumns.options[ProgramInput.lstColumns.length] = new Option(ResultArray[i],ResultArray[i]);			
				}
			}
}


    function AppendNewColumns(strSavedList,strMasterList){
        var SavedArray;
        var MasterArray;
        var i;
        var strOutput="";
        var strTemp="";
        
        SavedArray =  strSavedList.split(",");
        MasterArray = strMasterList.split(",");
        
        for (i=0;i< SavedArray.length;i++)
            if (SavedArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') != "")
                strOutput = strOutput + "," + SavedArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '');
 
        for (i=0;i< MasterArray.length;i++)
           if (MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') != "")
                {
                strTemp = "," + strOutput + "," 
                if (strTemp.indexOf("," + MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') + ",") == -1)
                    strOutput = strOutput + ",0," + MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '')

                }
        if (strOutput != "")
            strOutput = strOutput.substring(1);

        return strOutput;
} 

	//Added by PMV Pandian PBI 129405 Task 129466 - This is to adjust the width of Pulsar Products Dropdown list based on the values that are loaded  - Begin
	function AdjustPulsarProductsDropdownWidth() {
	     var maxWidth = 0;
	     var ddl =document.getElementById("lstProductsPulsar");
	    for (var i = 0; i < ddl.length; i++) {
	        if (ddl.options[i].text.length > maxWidth) {
	            maxWidth = ddl.options[i].text.length;
	        }
	    }
	    ddl.style.width = maxWidth * 7 + "px";
	    document.getElementById("lstDeveloper").style.width=maxWidth * 7 + "px";
	}
	//Added by PMV Pandian PBI 129405 Task 129466 - End
//-->
</script>
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
	Color:black;
}

TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	Color:black;
}

TD.HeaderButton
{
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana;
	FONT-WEIGHT: bold;
	COLOR: White;
}

</STYLE>
</head>

<body bgcolor="ivory" LANGUAGE="javascript" onload="return window_onload()" onblur="return window_onblur()">
<font size="3" face="verdana"><b>
<%if trim(request("HardwareMatrix")) = "1" then%>
	Create Custom Filter
<%else%>
	Deliverable Reports
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
	dim rs
	dim strSQL
	dim strStates
	dim strProductOptions
    dim strProductOptionsPulsar
	dim strInList
	dim strcatOptions
	dim strOSList
	dim strLanguages
	dim strQualStatus
	dim CurrentUser
	dim CurrentUserID
	dim strDivision
	dim strDevelopers
	dim strDevManagers
	dim strVendors
	dim strRoots
	dim blnCommodityMatrix
	dim blnAdmin
	dim CurrentUserPartner
	dim CurrentUserPartnerName
	dim CurrentUserPartnerType
	dim strProductGroupOptions
	dim strMasterColumnList
	dim strCoreTeams

	strProductGroupOptions = ""
	
	
	if trim(request("HardwareMatrix")) = "1" then
		blnCommodityMatrix = true
	else
		blnCommodityMatrix = false
	end if

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
		blnAdmin = rs("SystemAdmin")
		CurrentUserPartner = rs("PartnerID") & ""
	else
		strDivision = ""
		CurrentUserID = ""
		blnAdmin = false
		CurrentUserPartner = ""
	end if
	
	rs.Close


	if trim(CurrentUserPartner) = "9" then
		Response.Redirect "../mobilese/modusmain.asp"
	end if

'	if trim(CurrentUserPartner) = "1" then
'		CurrentUserPartnerName = "HP"
'	elseif trim(CurrentUserPartner) = "9" then
'		Response.Redirect "../mobilese/modusmain.asp"
'	else
'		rs.Open "spGetPartnerName " & CurrentUserPartner,cn,adOpenForwardOnly
'		if rs.EOF and rs.BOF then
'			CurrentUserPartnerName = ""
'		else
'			CurrentUserPartnerName = rs("Name") & ""
''		end if		
'		rs.Close
'	end if

	

    ''''Partner ODM Product Whitelist
    dim strWhitelistPartners
    dim arrWhitelistPartners
    dim isWhitelist
    strWhitelistPartners = ""
	rs.Open "SELECT ProductPartnerId FROM PartnerODMProductWhitelist WHERE UserPartnerId = " + CurrentUserPartner + ";",cn,adOpenForwardOnly
	do while not rs.EOF
		strWhitelistPartners = strWhitelistPartners + trim(rs("ProductPartnerId")) + ","
		rs.MoveNext
	loop
	rs.Close

    if trim(strWhitelistPartners) <> "" then
        strWhitelistPartners = left(strWhitelistPartners,(len(strWhitelistPartners) - 1) )
        arrWhitelistPartners = split(strWhitelistPartners,",")
    end if
  
	dim strLimitPartner
	if trim(CurrentUserPartner) <> "1" then
		rs.Open "spGetPartnerName " & CurrentUserPartner,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
    		CurrentUserPartnerType = "1"
		else
			CurrentUserPartnerType = rs("PartnerTypeID") & ""
		end if		
		rs.Close

        if trim(CurrentUserPartnerType) = "1" then
		    strLimitPartner = " pv.partnerid in (" & currentuserpartner  

            if strWhitelistPartners <> "" then
                strlimitpartner = strlimitpartner & "," & strWhitelistPartners
            end if

            strlimitpartner = strlimitpartner & ")"
        end if
	else
		strLimitPartner = ""
	end if

	'if blnCommodityMatrix then
	'	strSQL = "spListProductsOnCommodityMatrix"
'	elseif strDivision = "3" then
'		strSQL = "spGetProducts"
'	elseif strDivision = "2" then
'		strSQL = "spGetProductsByDivision 2"
	'else
	'	strSQL = "spGetProductsByDivision 1"
	'end if
    strSQL = "spListProductsOnCommodityMatrix"
	if strDivision <> "1" then
		Response.Write "<BR><BR><font size=2 face=verdana>This report only includes Mobile deliverables.<BR><BR></font><font size=1>Please look in <a target=""_blank"" href=""http://irsweb.usa.hp.com/"">IRS</a> for information on Desktop and Workstation deliverables.&nbsp;Contact <a href=""mailto: houcomirssupport@hp.com"">IRS Support</a> for assistance with IRS.</font><BR><BR>"
	end if
	rs.Open strSQL,cn,adOpenForwardOnly
	strProductOptions = ""
    strProductOptionsPulsar = ""
	strInList = ""
	do while not rs.EOF

        '''Herb, 11/08,2016, ODM White List, let Inventec and TNI cowork.
        isWhitelist = 0
        if strWhitelistPartners <> "" then
            for i = 0 to uBound(arrWhitelistPartners)
                if trim(rs("PartnerID")) = arrWhitelistPartners(i) then
                    isWhitelist = 1
                end if
            next
        end if

      if (rs("IsPulsarProduct") = 0) then 'list Legacy Products
		if (trim(CurrentUserPartner) = "1") or trim(CurrentUserPartnerType) <> "1"  or (trim(CurrentUserPartner) <> "9" and (trim(CurrentUserPartner)=trim(rs("PartnerID")) or isWhitelist ) ) then
			if rs("Name") & " " & rs("Version") <> "Test Product 1.0" then
				if trim(request("ID")) = trim(rs("ID")) then
					strProductOptions = strProductOptions &  "<Option selected value= """ & rs("ID") & """>" & rs("Dotsname")  & "</OPTION>"
				elseif rs("TypeID") = "1" or rs("TypeID") = "3" then
					strProductOptions = strProductOptions &  "<Option value= """ & rs("ID") & """>" & rs("Dotsname")  & "</OPTION>"
				end if
				strInList = strInList & "," & rs("ID") & ""
			end if
		end if
      end if
      dim productItem
      productItem =""
      productItem = rs("ID") &":"& rs("ProductReleaseID") 
      if (rs("IsPulsarProduct") = 1) then 'list Pulsar Products
		if (trim(CurrentUserPartner) = "1") or trim(CurrentUserPartnerType) <> "1"  or (trim(CurrentUserPartner) <> "9" and (trim(CurrentUserPartner)=trim(rs("PartnerID")) or isWhitelist ) ) then
			if rs("Name") & " " & rs("Version") <> "Test Product 1.0" then
                if blnCommodityMatrix then
				    if trim(request("ID")) = trim(rs("ID")) then
			    ' 05/05/2017 : PMV Pandian :PBI 129405- This is to pass the productid and its released id with a semicolon as a separator to HardwareMatrix.asp - Begin. 
					    strProductOptionsPulsar = strProductOptionsPulsar &  "<Option selected value= """ & rs("ProductReleaseID") & """>" & rs("Dotsname")  & "</OPTION>"
				    elseif rs("TypeID") = "1" or rs("TypeID") = "3" then
					    strProductOptionsPulsar = strProductOptionsPulsar &  "<Option value= """ & rs("ProductReleaseID") & """>" & rs("Dotsname")  & "</OPTION>"
                    	    ' 05/05/2017 : PMV Pandian :PBI 129405- This is to pass the productid and its released id with a semicolon as a separator to HardwareMatrix.asp - End.           
				    end if
				    
                else
                    if trim(request("ID")) = trim(rs("ID")) then
			    ' 05/05/2017 : PMV Pandian :PBI 129405- This is to pass the productid and its released id with a semicolon as a separator to HardwareMatrix.asp - Begin. 
					    strProductOptionsPulsar = strProductOptionsPulsar &  "<Option selected value= """ & productItem & """>" & rs("Dotsname")  & "</OPTION>"
				    elseif rs("TypeID") = "1" or rs("TypeID") = "3" then
					    strProductOptionsPulsar = strProductOptionsPulsar &  "<Option value= """ & productItem & """>" & rs("Dotsname")  & "</OPTION>"
                    	    ' 05/05/2017 : PMV Pandian :PBI 129405- This is to pass the productid and its released id with a semicolon as a separator to HardwareMatrix.asp - End.           
				    end if
                end if
                strInList = strInList & "," & rs("ID") & ""
			end if
		end if
     end if
		rs.MoveNext
	loop
	rs.Close

	if trim(currentuserpartner) = "1" then
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
	end if
	
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
	
	
	if not blnCommodityMatrix then
		strSQL = "spGetOS"
	
		rs.Open strSQL,cn,adOpenForwardOnly
		strOSList = ""
		do while not rs.EOF
			if rs("ID") <> 16 then
				strOSList = strOSList &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close

	end if

	strSQL = "spGetVendorList"
	
	rs.Open strSQL,cn,adOpenForwardOnly
	strVendors = ""
	do while not rs.EOF
		if rs("ID") <> 203 then
			strVendors = strVendors &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close

    

	strSQL = "spListDeliverableCoreTeams"
	rs.Open strSQL,cn,adOpenForwardOnly
	strCoreTeams = ""
	do while not rs.EOF
	    if rs("ID") <> 0 then
		    strCoreTeams = strCoreTeams &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close
	
		strSQL = "spListDevelopers"
	rs.Open strSQL,cn,adOpenForwardOnly
	strDevelopers = ""
	do while not rs.EOF
		strDevelopers = strDevelopers &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close

	strSQL = "spListDevManagers"
	rs.Open strSQL,cn,adOpenForwardOnly
	strDevManagers = ""
	do while not rs.EOF
		strDevManagers = strDevManagers &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close


	if blnCommodityMatrix then
		strSQL = "spListHWCategories"
	else
		strSQL = "spGetDeliverableCategories"
	end if
	rs.Open strSQL,cn,adOpenForwardOnly
	strCategories = ""
	strAbbr = ""
	do while not rs.EOF
		if rs("DeliverableTypeID") & "" = "1" then
			strAbbr = "HW - "
		elseif rs("DeliverableTypeID") & "" = "2" then
			strAbbr = "SW - "
		elseif rs("DeliverableTypeID") & "" = "3" then
			strAbbr = "FW - "
		elseif rs("DeliverableTypeID") & "" = "4" then
			strAbbr = "DOC - "
		elseif rs("DeliverableTypeID") & "" = "5" then
			strAbbr = "PACKAGING - "
		else
			strAbbr = ""
		end if
		if blnCommodityMatrix then
			strCategories = strCategories &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		else
			strCategories = strCategories &  "<Option value= """ & rs("ID") & """>" & strABBR & rs("Name") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close

	
	dim strCommodityPM
	
	if blnCommodityMatrix then
		rs.Open "spListCommodityPMs",cn,adOpenForwardOnly
		strCommodityPM = "<option value=""0"" selected></option>"
		do while not rs.EOF
			strCommodityPM = strCommodityPM &  "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
			rs.MoveNext
		loop
		rs.Close
	end if
	
	if blnCommodityMatrix then
		strSQL = "spListTestStatus"
			
		rs.Open strSQL,cn,adOpenForwardOnly
		strQualStatus = "<option value=0>Not Used</option>"
		do while not rs.EOF
			strQualStatus = strQualStatus &  "<Option value= """ & rs("ID") & """>" & rs("Status") & "</OPTION>"
			if rs("ID") = 5 then
                strQualStatus = strQualStatus & "<option value=-1>Risk Release</option>"
			end if
			rs.MoveNext
		loop
		rs.Close
	else
		strSQL = "spGetLanguages"
			
		rs.Open strSQL,cn,adOpenForwardOnly
		strLanguages = ""
		do while not rs.EOF
			if rs("ID") <> 58 then
				strLanguages = strLanguages &  "<Option value= """ & rs("ID") & """>" & rs("Abbreviation")  & "</OPTION>" '& " - " &  rs("language")
			end if
			rs.MoveNext
		loop
		rs.Close
	end if	
	
	dim strProfileOptions
	
	strProfileOptions = ""
	
	if Currentuserid <> "" then
		rs.Open "spListReportProfiles " & clng(CurrentUserID) & ",3",cn,adOpenForwardOnly
		strProfileOptions = ""
		do while not rs.EOF
			if (rs("Value8")<> 0 and request("HardwareMatrix") <> "") or request("HardwareMatrix") = "" then
				strProfileOptions = strProfileOptions & "<Option SharingID=0 PrimaryOwner="""" CanDelete=True CanEdit=True value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			end if
			rs.MoveNext
		loop
		rs.Close
		
		rs.Open "spListReportProfilesShared " & clng(CurrentUserID) & ",3",cn,adOpenForwardOnly
		do while not rs.EOF
			strProfileOptions = strProfileOptions & "<Option SharingID=""" & rs("SharingID") & """ PrimaryOwner=""" & shortname(rs("PrimaryOwner")& "")  &  """ CanDelete=" & rs("CanDelete") & " CanEdit=" & rs("CanEdit") & " value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
			rs.MoveNext
		loop
		rs.Close

		rs.Open "spListReportProfilesGroupShared " & clng(CurrentUserID) & ",3",cn,adOpenForwardOnly
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
<form ACTION="DelReport.asp" method=post NAME="ProgramInput" target="_blank">
<input type="hidden" id="txtDivision" name="txtDivision" value="<%=strDivision%>">
<table border=0>
	<tr>
		<td colspan="12">
		<font face=verdana size="2"><b>Saved Filters:&nbsp;</b></font>
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
<TR><TD colspan=9><HR>
<TABLE border=0 cellpadding=3 cellspacing=2>
	<TR bgcolor=#333333 ID=HeaderRow>
		<%if blnCommodityMatrix then%>
			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return CommodityCell_onmouseover()" onmouseout="return CommodityCell_onmouseout()" onclick="return cmdCommodity_onclick()">&nbsp;&nbsp;Filter&nbsp;Commodity&nbsp;Matrix&nbsp;Now&nbsp;&nbsp;</TD>
			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover()" onmouseout="return ActionCell_onmouseout()" onclick="javascript:ProgramInput.reset();">&nbsp;&nbsp;Clear&nbsp;This&nbsp;Page&nbsp;&nbsp;</TD></TR>
			<TD ID=StatusCell style="Display:none"> &nbsp;</TD>
		<%else%>
			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover(1)" onmouseout="return ActionCell_onmouseout(1)" onclick="return cmdReport_onclick(1)">&nbsp;&nbsp;Summary&nbsp;Report&nbsp;&nbsp;</TD>
			<%if trim(CurrentUserPartner) = "1" then%>
			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover(1)" onmouseout="return ActionCell_onmouseout(1)" onclick="return cmdReport_onclick(8)">&nbsp;&nbsp;Deliverable&nbsp;Details&nbsp;&nbsp;</TD>
			<TD ID=StatusCell class=HeaderButton LANGUAGE=javascript onmouseover="return MenuCell_onmouseover(1)" onmouseout="return MenuCell_onmouseout()">&nbsp;&nbsp;Deliverable&nbsp;Status&nbsp;&nbsp;</TD>
			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover(1)" onmouseout="return ActionCell_onmouseout(1)" onclick="return cmdReleases_onclick()">&nbsp;&nbsp;Releases&nbsp;&nbsp;</TD>
			<%else%>
			<TD ID=StatusCell style="Display:none"> &nbsp;</TD>
			<%end if%>
			<%if currentuserid <> 30 then%>
				<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover(2)" onmouseout="return ActionCell_onmouseout(2)" onclick="return cmdReport_onclick(2)">&nbsp;&nbsp;History&nbsp;Report&nbsp;&nbsp;</TD>
			<%end if%>
			

			<TD class=HeaderButton LANGUAGE=javascript onmouseover="return ActionCell_onmouseover()" onmouseout="return ActionCell_onmouseout()" onclick="return cmdReset_onclick();">&nbsp;&nbsp;Reset&nbsp;&nbsp;</TD></TR>
		<%end if%>
</TABLE>
<%
	Response.Write "<label ID=lblInst style=""display:none"">"
	Response.Write "<font color=Green size=1 face=verdana>Use CTRL or SHIFT keys to select multiple items in lists</font></label>"
%>

<tr>
		<td valign="top"><font size="2" face="verdana"><b>Product (Legacy):</b></font><br><select style="WIDTH: 200px; HEIGHT: 145px" multiple id="lstProducts" name="lstProducts">
				<%=strProductOptions%>
			</select>
		</td>
 
		<td valign="top"><font size="2" face="verdana"><b>Product (Pulsar):</b></font><br><select style="WIDTH: 200px; HEIGHT: 145px" multiple id="lstProductsPulsar" name="lstProductsPulsar">
				<%=strProductOptionsPulsar%>
			</select>
		</td>

		<td valign="top"><font size="2" face="verdana"><b>Product Group:</b></font><br><select style="WIDTH: 200px; HEIGHT: 145px" multiple id="lstProductGroups" name="lstProductGroups">
				<%=strProductGroupOptions%>
			</select>
		</td>
		<td valign=top><font size="2" face="verdana"><b>Vendor:</b></font><br>
			<select  style="WIDTH: 200px; HEIGHT: 145px" multiple size="2" id="lstVendor" name="lstVendor">
				<%=strVendors%>
			</select>
		</td>
		
		<td width=100% valign="top" colspan=6><font size="2" face="verdana"><b>Root&nbsp;Deliverable:</b></font><br>
			<select style="WIDTH: 100%; HEIGHT: 145px" multiple size="2" id="lstRoot" name="lstRoot">
<%
	if blnCommodityMatrix then
		strSQL = "spGetDelRootCommodities"
	else
		strSQL = "spGetDelRoot"
	end if
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		Response.Write   "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close


%>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><font size="2" face="verdana"><b>Manager/OTS PM:</b></font><br>
			<select style="WIDTH: 200px; HEIGHT: 102px" multiple size="2" id="lstDevManager" name="lstDevManager">
				<%=strDevManagers%>
			</select>
		</td>
		<td valign="top"><font size="2" face="verdana"><b>Developer:</b></font><br>
			<select style="WIDTH: 200px; HEIGHT: 102px" multiple size="2" id="lstDeveloper" name="lstDeveloper">
				<%=strDevelopers%>
			</select>
		</td>
		<%
		dim CategoryBoxWidth
		dim CategoryBoxColumns
		dim DisplayColumnBox
		if blnCommodityMatrix then
			CategoryBoxWidth="100%"
			CategoryBoxColumns="1"
			DisplayColumnBox="none"
		else
			CategoryBoxWidth="100%"
			CategoryBoxColumns="1"
			DisplayColumnBox=""
		end if
		
		
		%>
		<%if blnCommodityMatrix then 'and trim(currentuserpartner) = "1" then%>
			<select style="display:none" id="lstOS" name="lstOS">
		<%else%>

			<td><font size="2" face="verdana"><b>OS:</b></font><br><select style="WIDTH: 200px; HEIGHT: 102px" multiple size="2" id="lstOS" name="lstOS">
					<%=strOSList%>
				</select>
			</td>
		<%end if%>
		
		<%if blnCommodityMatrix then%>
			<td><font size="2" face="verdana"><b>Status:</b></font><br><select style="WIDTH: 200px; HEIGHT: 102px" multiple size="2" id="lstQualStatus" name="lstQualStatus">
					<%=strQualStatus%>
				</select>
				<select style="display:none" id="lstLanguage" name="lstLanguage">
			</td>
			<td style="display:none"><font size="2" face="verdana"><b>Core Team:</b></font><br>
			   <select style="WIDTH: 130px; HEIGHT: 102px" multiple size="2" id="lstCoreTeam" name="lstCoreTeam">
					<%=strCoreTeams%>
				</select>
				
			</td>
	
		<%else%>
			<td><font size="2" face="verdana"><b>Lang:</b></font><br><select style="WIDTH: 50px; HEIGHT: 102px" multiple size="2" id="lstLanguage" name="lstLanguage">
					<%=strLanguages%>
				</select>
				
				
				<select style="display:none" id="lstQualStatus" name="lstQualStatus">

			</td>
			<td><font size="2" face="verdana"><b>Core Team:</b></font><br>
			   <select style="WIDTH: 130px; HEIGHT: 102px" multiple size="2" id="lstCoreTeam" name="lstCoreTeam">
					<%=strCoreTeams%>
				</select>
				
			</td>

		<%end if%>

		<td width="20" valign="top"></td>

		<td style="WIDTH: <%=CategoryBoxWidth%>" colspan=<%=CategoryBoxColumns%> valign="top"><font size="2" face="verdana"><b>Category:</b></font><br>
			<select style="WIDTH: <%=CategoryBoxWidth%>; HEIGHT: 102px" multiple size="2" id="lstCategory" name="lstCategory">
				<%=strCategories%>
			</select>
		</td>
		<td style=display:<%=DisplayColumnBox%> width="20" valign="top"></td>
		<td style=display:<%=DisplayColumnBox%> width=120 valign="top"><font size="2" face="verdana"><b>Columns:</b>&nbsp;<font size=1><a href="javascript: ReorderColumns();">Reorder</a><!--&nbsp;&nbsp;<a href="javascript: ResetColumns();">Reset</a>--></font></font><br>
			<select style="WIDTH: 120; HEIGHT: 102px" multiple size="2" id="lstColumns" name="lstColumns">
			<%
    			dim ColumnArray
                '**********  Be sure to Update the IsProductColumn list on the report page ***************				
				'strMasterColumnList = "ID,Name,Version,Vendor,Vendor Version,Part Number,Model,Category,Core Team,Developer,Dev Manager,Software EOL,Factory EOA,Service EOA,Workflow,Product,Targeted,In Image,Softpaq,WHQL,TTS,HW Qual Status,Code Name,Pilot Status,MIT Signoff,MIT Samples,MIT Notes,ODM Signoff,ODM Samples,ODM Notes,WWAN Signoff,WWAN Samples,WWAN Notes,Dev Signoff,HW Version,FW Version,HW Rev,FCC ID,Anatel,ICASA,Secondary RF Kill"
				strMasterColumnList = "ID,Name,Version,HW Version,FW Version,HW Rev,Anatel,Category,Code Name,Core Team,Dev Manager,Dev Signoff,Developer,Device ID,Device ID String,Factory EOA,FCC ID,HW Qual Status,ICASA,In Image,IRS Part Number,MIT Samples,MIT Signoff,MIT Notes,Model,Part Number,Path,Pilot Status,Product,ODM Samples,ODM Signoff,ODM Notes,RF Kill Mechanism,Service EOA,Subsys Dev ID,Subsys Ven ID,Softpaq,Softpaq Numbers,Software EOL,Targeted,TTS,Vendor,Vendor ID,Vendor Version,WHQL,Workflow,WWAN Samples,WWAN Signoff,WWAN Notes"


				rs.Open "spGetDefaultProductFilter " & currentuserid & ",3",cn,adOpenStatic
				if rs.EOF and rs.BOF then
				    columnarray = split(strMasterColumnList,",")
				else
					columnarray = split(AppendNewColumns(rs("Setting") & "",strMasterColumnList),",")
				end if
				rs.Close

				for i = 0 to ubound(columnarray)
					Response.Write "<Option>" & columnarray(i) & "</Option>"
				next
			%>
			</select>
		</td>
	</tr>
</table>
<table width=100%>
  <tr>
	<td width="120"><font face="verdana" size="2"><b>Report Title:<b></font></td>
	<td width=100%>
		<%
			dim strTitle
			if blnCommodityMatrix then
				strTitle = ""
			else
				strTitle = "Deliverable Report"
			end if
		%>
		
		
	 <input type="text" id="txtTitle" name="txtTitle" value="<%=strTitle%>" style="Width:100%" maxlength=255></td>

	<%if blnCommodityMatrix then%>
		<TD nowrap width=50 valign=top><font size=2 face=verdana>&nbsp;&nbsp;&nbsp;<b>RoHS/Green:&nbsp;</b></font></TD>
		<td>
			<SELECT id=cboRohs name=cboRohs style="WIDTH:100;">
				<OPTION selected value=""></OPTION>
				<OPTION value=1>RoHS1</OPTION>
				<OPTION value=2>BFR/PVC</OPTION>
			</SELECT>&nbsp;&nbsp;&nbsp;
		</td>

		<TD nowrap width=120 valign=top><font size=2 face=verdana><b>Report&nbsp;Type:&nbsp;</b></font></TD>
		<td>
			<SELECT id=ReportFormat name=ReportFormat style="WIDTH:140;">
				<OPTION selected value="1">Qualification</OPTION>
				<OPTION value="2">Subassembly</OPTION>
				<OPTION value="3">Pilot</OPTION>
				<OPTION value="4">Accessory</OPTION>
				<OPTION value="5">Service</OPTION>
			</SELECT>
		</td>
	<%else%>
		<SELECT id=ReportFormat name=ReportFormat style="Display:none;"></SELECT>
		<SELECT id=cboRohs name=cboRohs style="Display:none;"></SELECT>
		
	<%end if%>

		
  </tr>
	<TR>
		<TD nowrap width=120 valign=top><font size=2 face=verdana><b>ID&nbsp;Numbers:</b></font><BR><font size=1 face=verdana color=Green>(comma&nbsp;separated)</font></TD>
		<td><INPUT id=txtNumbers name=txtNumbers style="WIDTH: 100%; HEIGHT: 22px" size=46 maxlength=80></td>
		
	<%if blnCommodityMatrix then%>
		<TD nowrap width=50 valign=top><font size=2 face=verdana><b>&nbsp;&nbsp;&nbsp;&nbsp;Inactive:&nbsp;</b></font></TD>
		<td>
			<SELECT id=cboEOL name=cboEOL style="WIDTH:100;">
				<OPTION selected value=""></OPTION>
				<OPTION value="0">Yes</OPTION>
				<OPTION value="1">No</OPTION>
			</SELECT>
		</td>

		<TD nowrap width=120 valign=top><font size=2 face=verdana><b>ODM&nbsp;PM:&nbsp;</b></font></TD>
		<td>
			<SELECT id=lstCommodityPM name=lstCommodityPM style="WIDTH:140;">
				<%=strCommodityPM%>
			</SELECT>
		</td>
	<%else%>
		<SELECT id=lstCommodityPM name=lstCommodityPM style="Display:none;"></SELECT>
		<SELECT id=cboEOL name=cboEOL style="Display:none;"></SELECT>
	<%end if%>

	</TR> 

	<%if blnCommodityMatrix then%>
		<TR style="Display:none">
	<%else%>
		<TR style="Display:none">
	<%end if%>
		<TD nowrap width=120 valign=top><font size=2 face=verdana><b>Planned&nbsp;Completion:&nbsp;</b></font></TD>	
		<TD><span ID="spnCompletionRange"><font size="2" face="verdana">Between:&nbsp;<INPUT id=txtCompleteDateStart name=txtCompleteDateStart style="WIDTH: 92px; HEIGHT: 22px" size=11>&nbsp;<a href="javascript: cmdDate_onclick(3)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a>&nbsp;and&nbsp;<INPUT id=txtCompleteDateEnd name=txtCompleteDateEnd style="WIDTH: 92px; HEIGHT: 22px" size=11>&nbsp;<a href="javascript: cmdDate_onclick(4)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a></font></span></TD>	
		</TR>

	<%if blnCommodityMatrix then%>
		<tr style="Display:none">
	<%else%>
		<tr>
	<%end if%>
	<td nowrap width="120"><font face="verdana" size="2"><b>Report Format:<b></font></td>
	<td>
		<select style="Width:80" id="cboFormat" name="cboFormat">
			<option selected value=0>HTML</option>
			<option value=1>Excel</option>
			<option value=2>Word</option>
		</select>&nbsp;
	</td>
  </tr>
	<%if blnCommodityMatrix then%>
		<tr style="Display:none">
	<%else%>
		<tr>
	<%end if%>
	<td width="120"><font face="verdana" size="2"><b>Search:<b></font></td>
	<td nowrap> <input type="text" id="txtSearch" name="txtSearch" value="" style="Width:134" maxlength=80>
	<font size=2 face=verdana>Look&nbsp;In:&nbsp;<INPUT type="checkbox" id=chkNameSearch name=chkNameSearch checked>Name&nbsp;&nbsp;<INPUT type="checkbox" id=chkChangesSearch name=chkChangesSearch>Changes&nbsp;&nbsp;<INPUT type="checkbox" id=chkDescriptionSearch name=chkDescriptionSearch>Description&nbsp;&nbsp;<INPUT type="checkbox" id=chkCommentsSearch  name=chkCommentsSearch>Comments</font>
	
	</td>
  </tr>
	<%if blnCommodityMatrix then%>
		<tr>
    	<td width="120"><font face="verdana" size="2"><b>Workflow Step:&nbsp;<b></font></td>
	    <td colspan=4>
		    <INPUT style="display:none" type="checkbox" id=chkDevelopment name=chkDevelopment>
		    <INPUT type="checkbox" id=chkTest name=chkTest>&nbsp;<font face=verdana size=2>Engineering Development</font>&nbsp;&nbsp;
		    <INPUT type="checkbox" id=chkRelease name=chkRelease>&nbsp;<font face=verdana size=2>Core Team</font>&nbsp;&nbsp;
		    <INPUT type="checkbox" id=chkComplete name=chkComplete checked>&nbsp;<font face=verdana size=2>Workflow Complete</font>&nbsp;&nbsp;
	    </td>
      </tr>  
	<%else%>
		<tr>
    	<td width="120"><font face="verdana" size="2"><b>Workflow Step:&nbsp;<b></font></td>
	    <td>
		    <INPUT type="checkbox" id=chkDevelopment name=chkDevelopment>&nbsp;<font face=verdana size=2>Development</font>&nbsp;&nbsp;
		    <INPUT type="checkbox" id=chkTest name=chkTest>&nbsp;<font face=verdana size=2>Functional Test</font>&nbsp;&nbsp;
		    <INPUT type="checkbox" id=chkRelease name=chkRelease>&nbsp;<font face=verdana size=2>Release Team</font>&nbsp;&nbsp;
		    <INPUT type="checkbox" id=chkComplete name=chkComplete>&nbsp;<font face=verdana size=2>Workflow Complete</font>&nbsp;&nbsp;
	    </td>
      </tr>  
	<%end if%>
	<%if blnCommodityMatrix then%>
		<tr style="Display:none">
	<%else%>
		<tr>
	<%end if%>
	<td width="120"><font face="verdana" size="2"><b>Status:&nbsp;<b></font></td>
	<td>
		<INPUT type="checkbox" id=chkTarget name=chkTarget>&nbsp;<font face=verdana size=2>Targeted</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkInImage name=chkInImage>&nbsp;<font face=verdana size=2>In Image</font>&nbsp;&nbsp;
		<INPUT type="checkbox" id=chkFailed name=chkFailed>&nbsp;<font face=verdana size=2>Failed</font>&nbsp;&nbsp;
	</td>
  </tr>  

	<TR>
		<td nowrap width="120"><font face="verdana" size="2"><b>Preference:<b></font></td>	
		<TD><INPUT type="checkbox" id=chkSCRestricted name=chkSCRestricted> Supply Chain Restriction</TD>
	</TR>
<%if blnadmin then%>  
	<TR>	
<%else%>
	<TR style="display:none">
<%end if%>
		<TD nowrap valign=top width=100><font size=2 face=verdana><b id=lblAdvanced>Other&nbsp;Criteria:</b></font>
		<!--<font size=1 face=verdana><BR><a target="_blank" href="DelSyntax.asp">Syntax</a>&nbsp;|&nbsp;<a href="javascript:BuildSQL();">Fields</a></font>-->
		</TD>
		<td  colspan=6 ><TEXTAREA id=txtAdvanced style="Font:Verdana; WIDTH: 100%; HEIGHT: 40px" name=txtAdvanced rows=2 cols=39><%=strLimitPartner%></TEXTAREA></td>
	</TR>
	<TR style="display:none" ID=AdminCriteriaRow>
		<TD nowrap valign=top width=100><font size=2 face=verdana><b id=lblAdvanced>Other&nbsp;Criteria:</b></font>
		</TD>
		<td  colspan=6 ID=AdminCriteria></td>
	</TR>


  
<%if currentuserid <> 30 then%>  
<%if trim(request("HardwareMatrix")) = "1" then%>
  <TR style="Display:">
 <%else%>
	<TR>
 <%end if%>
	<TD colspan=8>
	<FIELDSET>
		<LEGEND>Deliverable History</LEGEND>
		
		<span><b>Date Updated:&nbsp;</b></span>
		<select style="WIDTH:95" id="cboHistoryRange" name="cboHistoryRange" LANGUAGE="javascript" onchange="return cboHistoryRange_onchange()">
			<option value="<=">Less Than</option>
			<option value="=">Exactly</option>
			<option value=">=">More Than</option>
			<option selected>Range</option>
		</select>&nbsp;
		<span style="Display:none" ID="spnHistoryCount"><input style="width:55" type="text" id="txtHistoryDays" name="txtHistoryDays" value="1"> <font size="2" face="verdana">Days Ago</font></span>
		<span ID="spnHistoryRange"><font size="2" face="verdana">Between:&nbsp;<INPUT id=txtStartDate name=txtStartDate style="WIDTH: 92px; HEIGHT: 22px" size=11 value=<%=Date-30%>>&nbsp;<a href="javascript: cmdDate_onclick(1)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a>&nbsp;and&nbsp;<INPUT id=txtEndDate name=txtEndDate style="WIDTH: 92px; HEIGHT: 22px" size=11 value=<%=Date%>>&nbsp;<a href="javascript: cmdDate_onclick(2)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a></font></span>
			<BR><BR>
		
		<!--<b>Date Range:&nbsp;</b><INPUT id=txtStartDate1 name=txtStartDate1 style="WIDTH: 92px; HEIGHT: 22px" size=11 value=<%=Date-1%>>&nbsp;<a href="javascript: cmdDate_onclick(1)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a>
		to
		<INPUT id=txtEndDate1 name=txtEndDate1 style="WIDTH: 92px; HEIGHT: 22px" size=11 value=<%=Date%>>&nbsp;<a href="javascript: cmdDate_onclick(2)"><img align=absMiddle ID="picTarget" SRC="../mobilese/today/images/calendar.gif" alt="Choose Date" border=0 WIDTH=26 HEIGHT=21></a>
		<BR><BR>-->
			<TABLE width=100% border = 0 cellpadding=2 cellspacing=0>
				<TR bgcolor=gainsboro><TD><b>Change Type</b></TD><TD colspan=2><b>Specific Changes</b></TD></TR>
				<TR>
					<TD nowrap Style="BORDER-Top: gainsboro 1px solid;" valign=top><INPUT type="checkbox" id=chkChangePilot name=chkChangeType value="22"> Pilot Status Updated</TD>
					<TD Style="BORDER-Top: gainsboro 1px solid;" valign=top><INPUT type="hidden" id=txtSpecificPilotStatus name=txtSpecificPilotStatus><Span ID=PilotLink><a href="javascript:GetSpecificChange(1);">All Changes</a></Span></TD>
				</TR>
				<TR>
					<TD nowrap Style="BORDER-Top: gainsboro 1px solid;" valign=top><INPUT type="checkbox" id=chkChangeQual name=chkChangeType value="21"> Commodity Qualification Status Updated&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					<TD Style="WIDTH:100%;BORDER-Top: gainsboro 1px solid;" valign=top><INPUT type="hidden" id=txtSpecificQualStatus name=txtSpecificQualStatus><Span ID=QualLink><a href="javascript:GetSpecificChange(2);">All Changes</a></Span></TD>
				</TR>
				</Table>
	</FieldSet>

<!--
						<Table><tr><td>FROM:</TD><TD valign=top><a href="javascript:alert('Under Development');">Planning or Scheduled or On Hold</a></TD></TR>
								<TR><TD>TO:</TD><TD><a href="javascript:alert('Under Development');">Complete</A></TD></TR></TABLE>

-->
	
	</TD>
  </TR>
  <%end if%>
</table>
<br>
<%
	cn.Close
	set rs = nothing
	set cn=nothing
	
'Bug 26641/ Task 26642 - Harris, Valerie
'if Currentuserid = 31 then
	'Response.Write "<a href=""javascript: cmdDetails_onclick();"">Test</a>" 'cmdReport_onclick(7)
'end if	

%>

<input type="hidden" id="txtFunction" name="txtFunction">
<input type="hidden" id="txtDefaultColumns" name="txtDefaultColumns" value="">
<input type="hidden" id="txtMasterColumns" name="txtMasterColumns" value="<%=strMasterColumnList%>">
</form>

<div ID=OutputArea>

</Div>
<%if blnadmin then%>
	<input type="hidden" id="txtAdmin" name="txtAdmin" value="1">
<%else%>
	<input type="hidden" id="txtAdmin" name="txtAdmin" value="0">
<%end if%>

<INPUT type="hidden" id=txtUserID name=txtUserID value="<%=CurrentUserID%>">
<%if request("HardwareMatrix") = "" then%>
	<INPUT type="hidden" id=txtReportType name=txtReportType value="0">
<%else%>
	<INPUT type="hidden" id=txtReportType name=txtReportType value="1">
<%end if%>

<%
    function AppendNewColumns(strSavedList,strMasterList)
        dim SavedArray
        dim MasterArray
        dim i, j
        dim strOutput
        
        strOutput = ""
        
        SavedArray = split(strSavedList,",")
        MasterArray = split(strMasterList,",")
        
        for i = 0 to ubound(SavedArray)
            if trim(SavedArray(i)) <> "" then
                strOutput = strOutput & "," & trim(SavedArray(i))
            end if
        next

        for i = 0 to ubound(MasterArray)
            if trim(MasterArray(i)) <> "" then
                if instr("," & strOutput & ",","," & trim(MasterArray(i)) & "," ) = 0 then
                    strOutput = strOutput & "," & trim(MasterArray(i))
                end if
            end if
        next

        if strOutput <> "" then
            stroutput = mid(strOutput,2)
        end if
        AppendNewColumns = strOutput
    end function


%>
</body>
    <!--<script type="text/javascript">
        <!-- Added by PMV Pandian PBI 129405 Task 129466 - Begin
        window.onload = function() { 
            AdjustPulsarProductsDropdownWidth();
        };
        //Added by PMV Pandian PBI 129405 Task 129466 - End
    </script>-->
</html>
