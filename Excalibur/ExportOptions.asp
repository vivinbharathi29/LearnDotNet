<%@ Language=VBScript %>
<% response.ExpiresAbsolute = now() %>
<%
	'Response.Cookies("ExportProfileCount") = 5
	'Response.Cookies("ExportProfile0") = "Stuff Last Week"
	'Response.Cookies("ExportProfile1") = "Stuff Today"
	'Response.Cookies("ExportProfile2") = "Stuff Next Week"
	'Response.Cookies("ExportProfile1Columns") = "ID,Status"
	'Response.Cookies("ExportProfile0Columns") = "ID,Type"
	'Response.Cookies("ExportProfile0Products") = "Awards E 2.0"
	'Response.Cookies("ExportProfile1Products") = "Awards E 2.0,Boone EP 1.0"
	'Response.Cookies("ExportProfile2Products") = "Boone EP 1.0,Awards E 2.0"
	'Response.Cookies("ExportProfile1Header") = "1"

%>

<!-- #include file = "includes/noaccess.inc" -->
<!DOCTYPE html>
<html>
<head>
<title> Export Option</title>
<!-- #include file="includes/bundleConfig.inc" -->
<script type="text/javascript" src="Scripts/PulsarPlus.js"></script>
<script type="text/javascript" src="includes/client/json2.js"></script>
<script type="text/javascript" src="includes/client/json_parse.js"></script>
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function getCookieValue(cookieName)
{
	var cookieValue = document.cookie;
	var cookieStartsAt = cookieValue.indexOf(" " + cookieName + "=");
	if (cookieStartsAt == -1)
	{
		cookieStartsAt = cookieValue.indexOf(cookieName + "=");
	}
	if (cookieStartsAt == -1)
	{
		cookieValue="";
	}
	else
	{
		cookieStartsAt = cookieValue.indexOf("=",cookieStartsAt) + 1;
		var cookieEndsAt = cookieValue.indexOf(";",cookieStartsAt);
		if (cookieEndsAt == -1)
		{
			cookieEndsAt = cookieValue.length;
		}
		cookieValue=unescape(cookieValue.substring(cookieStartsAt,cookieEndsAt));
	}
	return cookieValue;
}


function cmdOK_onclick() {
    var i;
    var strProfileName;
    var strColumns;
    var strProducts;
    var strHeader;
    var CookieCount;
    if (DetailsForm.lstSelectedProd.length == 0 && DetailsForm.cboProfile.selectedIndex != 1) {
        window.alert("You must select at least one product.");
    }
    else if (DetailsForm.lstSelected.length == 0) {
        window.alert("You must select at least one column.");
    }
    else if (DetailsForm.cboProfile.selectedIndex == 1) {
        strProfileName = prompt("Enter a name for this profile. Please keep it as short as possible.", "Profile Name");
        if (strProfileName != null) {
            strColumns = "";
            for (i = 0; i < DetailsForm.lstSelected.length; i++) {
                strColumns = strColumns + DetailsForm.lstSelected.options[i].text + ","
            }

            strProducts = "";
            for (i = 0; i < DetailsForm.lstSelectedProd.length; i++) {
                strProducts = strProducts + DetailsForm.lstSelectedProd.options[i].text + ","
            }

            if (DetailsForm.chkHeader.checked)
                strHeader = "1";
            else
                strHeader = "";

            //var strID = new Array();
            //strID = window.showModalDialog("ExportQuery.asp?ActionType=" + txtActionType.value + "&EmployeeID=" + txtEmployeeID.value + "&Name=" + strProfileName + "&Products=" + strProducts + "&Columns=" + strColumns + "&Header=" + strHeader + "&Type=4", "", "dialogWidth:350px;dialogHeight:100px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            modalDialog.open({ dialogTitle: 'Export Option', dialogURL: 'ExportQuery.asp?ActionType=' + txtActionType.value + '&EmployeeID=' + txtEmployeeID.value + '&Name=' + strProfileName + '&Products=' + strProducts + '&Columns=' + strColumns + '&Header=' + strHeader + '&Type=4', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
            setTimeout(function () {
                strID = modalDialog.getArgument('export_query_array');
                strID = JSON.parse(strID);
            }, 1000);

            setTimeout(function () {
                if (typeof (strID) != "undefined") {
                    if (strID[0] != "") {
                        if (DetailsForm.cboProfile.length == 2) {
                            DetailsForm.cboProfile.options[DetailsForm.cboProfile.length] = new Option("----------------------------------------------------------", 0);
                        }

                        DetailsForm.cboProfile.options[DetailsForm.cboProfile.length] = new Option(strProfileName, strID[0]);

                        if (DetailsForm.lstSelectedProd.length == 0) //Add New Profile Only - No Excel
                        {
                            DetailsForm.cboProfile.selectedIndex = DetailsForm.cboProfile.length - 1
                            window.alert("Profile Saved");
                        }
                        else //Export to Excel
                        {
                            for (i = 0; i < DetailsForm.lstSelected.length; i++)
                                DetailsForm.lstSelected.options[i].selected = true;
                            for (i = 0; i < DetailsForm.lstSelectedProd.length; i++)
                                DetailsForm.lstSelectedProd.options[i].selected = true;
                            if ('<%=Request.QueryString("pulsarplusDivId")%>' != '') {
                                DetailsForm.Query.value = '<%=Request.QueryString("export_option_query")%>';
                            }
                            else {
                                DetailsForm.Query.value = parent.modalDialog.getArgument('export_option_query');
                            }
                            DetailsForm.submit();
                            if (window.location != window.parent.location) {
                                parent.modalDialog.cancel();
                            } else {
                                window.close();
                            }
                        }
                    } else {
                        window.alert("Unable to add profile.");
                    }
                } else {
                    window.alert("Unable to add profile.");
                }
            }, 1000);
        }
    }
    else if (DetailsForm.cboProfile.selectedIndex == 2) {
        DetailsForm.cboProfile.focus();
        window.alert("You must select a valid report option.");
    }
    else {
        for (i = 0; i < DetailsForm.lstSelected.length; i++)
            DetailsForm.lstSelected.options[i].selected = true;
        for (i = 0; i < DetailsForm.lstSelectedProd.length; i++)
            DetailsForm.lstSelectedProd.options[i].selected = true;
        if ('<%=Request.QueryString("pulsarplusDivId")%>' != '') {
            DetailsForm.Query.value = '<%=Request.QueryString("export_option_query")%>';
        }
        else {
            DetailsForm.Query.value = parent.modalDialog.getArgument('export_option_query');
        }
        //DetailsForm.Query.value = parent.modalDialog.getArgument('export_option_query');
        DetailsForm.submit();

        if ('<%Request.QueryString("pulsarplusDivId")%>' != undefined && '<%Request.QueryString("pulsarplusDivId")%>' != "") {
            parent.window.parent.closeExternalPopup();
        }
        else
        {
            if (window.location != window.parent.location) {
                parent.modalDialog.cancel();
            } else {
            window.close();
            }
        }
    }

}

function cmdRemoveAll_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstSelected.length;i++)
		{
			DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.length] = new Option(DetailsForm.lstSelected.options[i].text,DetailsForm.lstSelected.options[i].value);
		}
	for (i=DetailsForm.lstSelected.length-1;i>=0;i--)
			DetailsForm.lstSelected.options[i] = null;

}

function cmdRemove_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstSelected.length;i++)
		{
			if(DetailsForm.lstSelected.options[i].selected)
				{
					DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.length] = new Option(DetailsForm.lstSelected.options[i].text,DetailsForm.lstSelected.options[i].value);
				}
		}
	for (i=DetailsForm.lstSelected.length-1;i>=0;i--)
		{
			if(DetailsForm.lstSelected.options[i].selected)
					DetailsForm.lstSelected.options[i] = null;
		}
}

function cmdAddAll_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstAvailable.length;i++)
		{
			DetailsForm.lstSelected.options[DetailsForm.lstSelected.length] = new Option(DetailsForm.lstAvailable.options[i].text,DetailsForm.lstAvailable.options[i].value);
		}
	for (i=DetailsForm.lstAvailable.length-1;i>=0;i--)
		{
			DetailsForm.lstAvailable.options[i] = null;
		}
}

function cmdAdd_onclick() {
	var i;
	for (i=0;i<DetailsForm.lstAvailable.length;i++)
		{
			if(DetailsForm.lstAvailable.options[i].selected)
				{
					DetailsForm.lstSelected.options[DetailsForm.lstSelected.length] = new Option(DetailsForm.lstAvailable.options[i].text,DetailsForm.lstAvailable.options[i].value);
				}
		}
	for (i=DetailsForm.lstAvailable.length-1;i>=0;i--)
		{
			if(DetailsForm.lstAvailable.options[i].selected)
					DetailsForm.lstAvailable.options[i] = null;
		}		
}		


function lstAvailable_ondblclick() {
	if (DetailsForm.lstAvailable.options.selectedIndex > -1)
		{
		DetailsForm.lstSelected.options[DetailsForm.lstSelected.length] = new Option(DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.options.selectedIndex].text,DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.options.selectedIndex].value);
		DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.options.selectedIndex] = null;
		}
}

function lstSelected_ondblclick() {
	if (DetailsForm.lstSelected.options.selectedIndex > -1)
		{
			DetailsForm.lstAvailable.options[DetailsForm.lstAvailable.length] = new Option(DetailsForm.lstSelected.options[DetailsForm.lstSelected.options.selectedIndex].text,DetailsForm.lstSelected.options[DetailsForm.lstSelected.options.selectedIndex].value);
			DetailsForm.lstSelected.options[DetailsForm.lstSelected.options.selectedIndex] = null;
		}
}

function cmdRemoveAllProd_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstSelectedProd.length;i++)
		{
			DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.length] = new Option(DetailsForm.lstSelectedProd.options[i].text,DetailsForm.lstSelectedProd.options[i].value);
		}
	for (i=DetailsForm.lstSelectedProd.length-1;i>=0;i--)
			DetailsForm.lstSelectedProd.options[i] = null;

}

function cmdRemoveProd_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstSelectedProd.length;i++)
		{
			if(DetailsForm.lstSelectedProd.options[i].selected)
				{
					DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.length] = new Option(DetailsForm.lstSelectedProd.options[i].text,DetailsForm.lstSelectedProd.options[i].value);
				}
		}
	for (i=DetailsForm.lstSelectedProd.length-1;i>=0;i--)
		{
			if(DetailsForm.lstSelectedProd.options[i].selected)
					DetailsForm.lstSelectedProd.options[i] = null;
		}
}

function cmdAddAllProd_onclick() {
	var i;
	
	for (i=0;i<DetailsForm.lstAvailableProd.length;i++)
		{
			DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.length] = new Option(DetailsForm.lstAvailableProd.options[i].text,DetailsForm.lstAvailableProd.options[i].value);
		}
	for (i=DetailsForm.lstAvailableProd.length-1;i>=0;i--)
		{
			DetailsForm.lstAvailableProd.options[i] = null;
		}
}

function cmdAddProd_onclick() {
	var i;
	for (i=0;i<DetailsForm.lstAvailableProd.length;i++)
		{
			if(DetailsForm.lstAvailableProd.options[i].selected)
				{
					DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.length] = new Option(DetailsForm.lstAvailableProd.options[i].text,DetailsForm.lstAvailableProd.options[i].value);
				}
		}
	for (i=DetailsForm.lstAvailableProd.length-1;i>=0;i--)
		{
			if(DetailsForm.lstAvailableProd.options[i].selected)
					DetailsForm.lstAvailableProd.options[i] = null;
		}		
}		


function lstAvailableProd_ondblclick() {
	if (DetailsForm.lstAvailableProd.options.selectedIndex > -1)
		{
		DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.length] = new Option(DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.options.selectedIndex].text,DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.options.selectedIndex].value);
		DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.options.selectedIndex] = null;
		}
}

function lstSelectedProd_ondblclick() {
	if (DetailsForm.lstSelectedProd.options.selectedIndex > -1)
		{
			DetailsForm.lstAvailableProd.options[DetailsForm.lstAvailableProd.length] = new Option(DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.options.selectedIndex].text,DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.options.selectedIndex].value);
			DetailsForm.lstSelectedProd.options[DetailsForm.lstSelectedProd.options.selectedIndex] = null;
		}
}

function cmdCancel_onclick(pulsarplusDivId) {

    if (window.parent.location.pathname.indexOf('GetPmViewDetails') > 0 )
    {
        parent.closeExternalPopup();
    }
    else
    {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            if (window.location != window.parent.location) {
                window.parent.modalDialog.cancel();
            } else {
                window.close();
            }
        }
    }
}



function cboProfile_onkeydown() {
	if(event.keyCode == 46 && DetailsForm.cboProfile.selectedIndex > 2)
		window.alert("Deleting");
}



function Right(str, n){
	if (n <= 0)     // Invalid bound, return blank string
		return "";
	else if (n > String(str).length)   // Invalid bound, return
		return str;                     // entire string
	else 
		{ // Valid bound, return appropriate substring
			var iLen = String(str).length;
            return String(str).substring(iLen, iLen - n);
        }
}

function Replace(strString,strReplace,strWith){

  var arrTemp = strString.split(strReplace);
  strString = arrTemp.join(strWith);


	return strString
}

/*function cboProfile_onchangeCookieVersion() {
	var strColumns;
	var strProducts;
	var strBuffer;
	var i;
	var strHeader;
		
	if (DetailsForm.cboProfile.selectedIndex > 2)
		{
			ProfileOptions.style.display = "";
			strHeader = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Header");
			if (strHeader=="1")
				DetailsForm.chkHeader.checked = true;
			else
				DetailsForm.chkHeader.checked = false;
			cmdRemoveAll_onclick();
			cmdRemoveAllProd_onclick();

			for(i=0;i<DetailsForm.lstAvailable.length;i++)				
				DetailsForm.lstAvailable.options[i].selected=false;
			for(i=0;i<DetailsForm.lstAvailableProd.length;i++)				
				DetailsForm.lstAvailableProd.options[i].selected=false;

			strColumns = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Columns");
			if (Right(strColumns) != ",")
				strColumns=strColumns+",";
			strColumns = Replace(strColumns,"+"," ");
			while (strColumns.indexOf(",")> -1 )
				{
					strBuffer = strColumns.substring(0,strColumns.indexOf(","));
					
					for(i=0;i<DetailsForm.lstAvailable.length;i++)				
						{
							if(DetailsForm.lstAvailable.options[i].text==strBuffer)
								{
									DetailsForm.lstAvailable.options[i].selected=true;
								}
						}
					cmdAdd_onclick();
					strColumns = strColumns.substring(strColumns.indexOf(",")+1);
				}

			strProducts = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Products");

			if (Right(strProducts) != ",")
				strProducts=strProducts+",";
			strProducts = Replace(strProducts,"+"," ");
			while (strProducts.indexOf(",")> -1 )
				{
					strBuffer = strProducts.substring(0,strProducts.indexOf(","));
					
					for(i=0;i<DetailsForm.lstAvailableProd.length;i++)				
						{
							if(DetailsForm.lstAvailableProd.options[i].text==strBuffer)
								{
									DetailsForm.lstAvailableProd.options[i].selected=true;
								}
						}
					cmdAddProd_onclick();
					strProducts = strProducts.substring(strProducts.indexOf(",")+1);
				}
		}
	else
		{
			ProfileOptions.style.display = "none";		
		}
}*/

function cboProfile_onchange() {
	var strColumns;
	var strProducts;
	var strBuffer;
	var i;
	var strHeader;
		
	if (DetailsForm.cboProfile.selectedIndex > 2)
	{
	    ProfileOptions.style.display = "";
			
	    //var strID = new Array();
	    //window.showModalDialog("ExportQuery.asp?Type=1&ID=" + DetailsForm.cboProfile.value,"","dialogWidth:350px;dialogHeight:100px;edge: Raised;center:Yes; help: No;resizable: No;status: No");                               
	    modalDialog.open({ dialogTitle: 'Export Option', dialogURL: 'ExportQuery.asp?Type=1&ID=' + DetailsForm.cboProfile.value + '', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
	    setTimeout(function () {
	        strID = modalDialog.getArgument('export_query_array');
	        strID = JSON.parse(strID);
	    }, 1000);
			
	    setTimeout(function () {
	        if (typeof (strID) != "undefined") {
	            strProducts = strID[0];
	            strColumns = strID[1];
	            strHeader = strID[2];
	        }

	        //			strHeader = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Header");
	        if (strHeader == "1")
	            DetailsForm.chkHeader.checked = true;
	        else
	            DetailsForm.chkHeader.checked = false;
	        cmdRemoveAll_onclick();
	        cmdRemoveAllProd_onclick();

	        for (i = 0; i < DetailsForm.lstAvailable.length; i++)
	            DetailsForm.lstAvailable.options[i].selected = false;
	        for (i = 0; i < DetailsForm.lstAvailableProd.length; i++)
	            DetailsForm.lstAvailableProd.options[i].selected = false;

	        //			strColumns = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Columns");
	        if (Right(strColumns) != ",")
	            strColumns = strColumns + ",";
	        strColumns = Replace(strColumns, "+", " ");
	        while (strColumns.indexOf(",") > -1) {
	            strBuffer = strColumns.substring(0, strColumns.indexOf(","));

	            for (i = 0; i < DetailsForm.lstAvailable.length; i++) {
	                if (DetailsForm.lstAvailable.options[i].text == strBuffer) {
	                    DetailsForm.lstAvailable.options[i].selected = true;
	                }
	            }
	            cmdAdd_onclick();
	            strColumns = strColumns.substring(strColumns.indexOf(",") + 1);
	        }

	        //			strProducts = getCookieValue("ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Products");

	        if (Right(strProducts) != ",")
	            strProducts = strProducts + ",";
	        strProducts = Replace(strProducts, "+", " ");
	        while (strProducts.indexOf(",") > -1) {
	            strBuffer = strProducts.substring(0, strProducts.indexOf(","));

	            for (i = 0; i < DetailsForm.lstAvailableProd.length; i++) {
	                if (DetailsForm.lstAvailableProd.options[i].text == strBuffer) {
	                    DetailsForm.lstAvailableProd.options[i].selected = true;
	                }
	            }
	            cmdAddProd_onclick();
	            strProducts = strProducts.substring(strProducts.indexOf(",") + 1);
	        }
	    }, 1000);
	} else {
		ProfileOptions.style.display = "none";
	}
}


function RenameProfile_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function RenameProfile_onclick() {
	strProfileName = prompt("Enter new name for this profile. Please keep it as short as possible.",DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex].text);
	if (strProfileName != null && strProfileName != ""){
		//var strID = new Array();
		//strID = window.showModalDialog("ExportQuery.asp?NewName=" + strProfileName + "&Type=2&ID=" + DetailsForm.cboProfile.value,"","dialogWidth:350px;dialogHeight:100px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		modalDialog.open({ dialogTitle: 'Export Option', dialogURL: 'ExportQuery.asp?NewName=' + strProfileName + '&Type=2&ID=' + DetailsForm.cboProfile.value + '', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
		setTimeout(function () {
			strID = modalDialog.getArgument('export_query_array');
			strID = JSON.parse(strID);
		}, 1000);

		setTimeout(function () {
			if (typeof (strID) != "undefined") {
			    if (strID[0] == "1") {
			        DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex].text = strProfileName;
			    } else {
			        window.alert("Unable to rename this profile.");
			    }
			} else {
			    window.alert("Unable to rename this profile.");
			}
		}, 1000);
	}
}

function RenameProfile_onmouseout() {
	window.event.srcElement.style.color = "blue";
}

function RenameProfile_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function DeleteCookie(sName)
{
  document.cookie = sName + "=1" + "; expires=Fri, 31 Dec 1999 23:59:59 GMT;";
}


function DeleteProfile_onclick() {
    if (window.confirm("Are you sure you want to delete your " + DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex].text + " profile?")) {
        //var strID = new Array();
        //strID = window.showModalDialog("ExportQuery.asp?Type=3&ID=" + DetailsForm.cboProfile.value,"","dialogWidth:350px;dialogHeight:100px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
        modalDialog.open({ dialogTitle: 'Export Option', dialogURL: 'ExportQuery.asp?Type=3&ID=' + DetailsForm.cboProfile.value + '', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
        setTimeout(function () {
            strID = modalDialog.getArgument('export_query_array');
            strID = JSON.parse(strID);
        }, 1000);
        setTimeout(function () {
            if (typeof (strID) != "undefined") {
                if (strID[0] == "1") {
                    //Delete Profile from combo box
                    DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex] = null;

                    //Delete Profile Divider from combo box if necessary
                    if (DetailsForm.cboProfile.length == 3) {
                        DetailsForm.cboProfile.options[2] = null;
                    }

                    //Select First item in Combo 
                    DetailsForm.cboProfile.options[0].selected = true;
                } else {
                    window.alert("Unable to delete this profile.");
                }
            } else {
                window.alert("Unable to delete this profile.");
            }
        }, 1000);
    }
}
/*function DeleteProfile_onclickCookies() {
	if (window.confirm("Are you sure you want to delete your " + DetailsForm.cboProfile.options(DetailsForm.cboProfile.selectedIndex).value + " profile?"))
		{
			var CurrentIndex = DetailsForm.cboProfile.selectedIndex;
			DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex] = null;

			var CookieCount = getCookieValue("ExportProfileCount");
			if (CookieCount == "") 
				CookieCount = 0;

			var i;
			var expireDate = new Date();
			expireDate.setMonth(expireDate.getMonth()+12);

			for (i=(CurrentIndex -3);i<CookieCount-1;i++)
				{
				document.cookie = "ExportProfile" + i + "=" + getCookieValue("ExportProfile" + (i+1)) + ";expires=" + expireDate.toGMTString() + ";";
				document.cookie = "ExportProfile" + i + "Header=" + getCookieValue("ExportProfile" + (i+1) + "Header") + ";expires=" + expireDate.toGMTString() + ";";
				document.cookie = "ExportProfile" + i + "Columns=" + getCookieValue("ExportProfile" + (i+1) + "Columns") + ";expires=" + expireDate.toGMTString() + ";";
				document.cookie = "ExportProfile" + i + "Products=" + getCookieValue("ExportProfile" + (i+1) + "Products") + ";expires=" + expireDate.toGMTString() + ";";
				}

			//Delete Last Cookies
			if (CookieCount > 0)
				{
				DeleteCookie("ExportProfile" + (CookieCount -1));
				DeleteCookie("ExportProfile" + (CookieCount -1)+ "Columns" );
				DeleteCookie("ExportProfile" + (CookieCount -1) + "Products" );
				DeleteCookie("ExportProfile" + (CookieCount -1) + "Header" );
				}
			
			//Save CookieCount
			if (CookieCount > 0)
				CookieCount--;
			document.cookie = "ExportProfileCount=" + CookieCount + ";expires=" + expireDate.toGMTString() + ";";

			//Delete Profile from combo box
			if (CookieCount == 0)
				DetailsForm.cboProfile.options[2] = null;				
			DetailsForm.cboProfile.options(0).selected = true;
		}
}
*/
function DeleteProfile_onmouseout() {
	window.event.srcElement.style.color = "blue";
}

function DeleteProfile_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function UpdateProfile_onmouseover() {
	window.event.srcElement.style.color = "red";
	window.event.srcElement.style.cursor = "hand";
}

function UpdateProfile_onclick() {
	if (window.confirm("Are you sure you want to update your " + DetailsForm.cboProfile.options[DetailsForm.cboProfile.selectedIndex].text + " profile with the selected options?"))
		{
			var strColumns;
			var strProducts;
			var strHeader;
			
			strColumns = "";
			for (i=0;i<DetailsForm.lstSelected.length;i++)
				{
					strColumns = strColumns + DetailsForm.lstSelected.options[i].text + ","
				}

			strProducts = "";
			for (i=0;i<DetailsForm.lstSelectedProd.length;i++)
				{
					strProducts = strProducts + DetailsForm.lstSelectedProd.options[i].text + ","
				}
			if (DetailsForm.chkHeader.checked)
				strHeader = "1";
			else
				strHeader = "";



			//var strID = new Array();
			//strID = window.showModalDialog("ExportQuery.asp?ID=" + DetailsForm.cboProfile.value + "&Products=" + strProducts + "&Columns=" + strColumns + "&Header=" + strHeader + "&Type=5","","dialogWidth:350px;dialogHeight:100px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
			modalDialog.open({ dialogTitle: 'Export Option', dialogURL: 'ExportQuery.asp?ID=' + DetailsForm.cboProfile.value + '&Products=' + strProducts + '&Columns=' + strColumns + '&Header=' + strHeader + '&Type=5', dialogHeight: 175, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
			setTimeout(function () {
			    strID = modalDialog.getArgument('export_query_array');
			    strID = JSON.parse(strID);
			}, 1000);

			setTimeout(function () {
			    if (typeof (strID) != "undefined") {
			        if (strID[0] == "1") {
			            window.alert("Profile updated.");
			        } else {
			            window.alert("Unable to update profile.");
			        }
			    } else {
			        window.alert("Unable to update profile.");
			    }
			}, 1000);




/*			var expireDate = new Date();
			expireDate.setMonth(expireDate.getMonth()+12);

			document.cookie = "ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Products=" + strProducts + ";expires=" + expireDate.toGMTString() + ";";
			document.cookie = "ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Columns=" + strColumns + ";expires=" + expireDate.toGMTString() + ";";
			document.cookie = "ExportProfile" + (DetailsForm.cboProfile.selectedIndex - 3) + "Header=" + strHeader + ";expires=" + expireDate.toGMTString() + ";";
*/
		}
}

function UpdateProfile_onmouseout() {
	window.event.srcElement.style.color = "blue";
}

function lstSelected_onkeydown() {
	if(window.event.keyCode == 46)
		cmdRemove_onclick();
	
}

function lstSelectedProd_onkeydown() {
	if(window.event.keyCode == 46)
		cmdRemoveProd_onclick();
}

//*****************************************************************
//Description:  Use in parent page; closes modal dialog opend with modalDialog code
//*****************************************************************
function closeModalDialog(bReload) {
    modalDialog.cancel(bReload);
};

//*****************************************************************
//Description:  Initiaties modal dialog 
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();
}
//-->
</script>
</head>
<body onload="window_onload();" bgColor="ivory">
<%
	dim cn
	dim cm
	dim p
	dim rs
	dim strProducts
	dim strCurrentProduct

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn
	

	rs.Open "spgetproducts",cn,adOpenForwardOnly
	strproducts = ""
	do while not rs.EOF
		if isdate(rs("PDDReleased")) then
			if clng(request("ID")) <> rs("ID") then
				strproducts = strproducts &  "<OPTION Value=""" & rs("ID") & """>" & rs("Name") & " " & rs("Version") & "</OPTION>"
			else
				strCurrentProduct = "<OPTION Value=""" & rs("ID") & """>" & rs("Name") & " " & rs("Version") & "</OPTION>"
			end if
		end if
		rs.movenext
	loop
	rs.Close

	dim CurrentUser
	dim CurrentUserID

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
	
	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserName = rs("Name") & ""
	end if
	rs.Close


	dim strProfileOptions
	
	strProfileOptions = ""
	
	rs.Open "spListExcelProfiles " & clng(CurrentUserID),cn,adOpenForwardOnly
	strProfileOptions = ""
	do while not rs.EOF
		strProfileOptions = strProfileOptions & "<Option value=""" & rs("ID") & """>" & rs("ProfileName") & "</Option>"
		rs.MoveNext
	loop
	rs.Close
		

	if strProfileOptions <> "" then
		strProfileOptions = "<option selected>Use Options Selected Below</option><option>Save as New Profile</option><option>----------------------------------------------------------</option>" & strProfileOptions	
	else
		strProfileOptions = "<option selected>Use Options Selected Below</option><option>Save as New Profile</option>"	
	end if	

	cn.Close
	set rs=nothing
	set cn=nothing

%>




<br>
<font size="2" face="verdana">
<form id="DetailsForm" target=_blank method="post" action="ExportDetails.asp">
<input type="hidden" name="hidBios" value="<%= Request("Bios") %>" />
<input type="hidden" name="hidType" value="<%= Request("Type") %>" />
<input type="hidden" name="hidStatus" value="<%= Request("Status") %>" />
<table border="0" style="BORDER-RIGHT: 2px outset; BORDER-TOP: 2px outset; BORDER-LEFT: 2px outset; BORDER-BOTTOM: 2px outset; BACKGROUND-COLOR: LightGrey">
	<tr>
		<td colspan="4"><font size="2"><b>Report Profile:&nbsp;</b></font>
		<select id="cboProfile" name="cboProfile" style="WIDTH: 250px" LANGUAGE=javascript onkeydown="return cboProfile_onkeydown()" onchange="return cboProfile_onchange()"><%=strProfileOptions%></select></font>
		<font size="2" face="verdana" color=blue><div style="Display:none" ID=ProfileOptions align=right><font size=1 face=verdana><u ID=UpdateProfile LANGUAGE=javascript onclick="return UpdateProfile_onclick()" onmouseout="return UpdateProfile_onmouseout()" onmouseover="return UpdateProfile_onmouseover()">Update</u> <u ID=DeleteProfile LANGUAGE=javascript onclick="return DeleteProfile_onclick()" onmouseout="return DeleteProfile_onmouseout()" onmouseover="return DeleteProfile_onmouseover()">Delete</u> <u id=RenameProfile LANGUAGE=javascript onmouseover="return RenameProfile_onmouseover()" onclick="return RenameProfile_onclick()" onmouseout="return RenameProfile_onmouseout()">Rename</u></font></div>
		</td>
		
	</tr>
	<tr>
		<td colspan="4"><hr></td>
	</tr>
	<tr>
		<td colspan="2"><font size="2"><b>Available Columns:</b></font></td>
		<td><font size="2"><b>Selected Columns:</b></font></td>
	</tr>
<tr>
	<td>
		<select id="lstAvailable" style="WIDTH: 165px; HEIGHT: 150px" size="2" name="lstAvailable" multiple LANGUAGE="javascript" ondblclick="return lstAvailable_ondblclick()"> 
			<option value="Type">Type</option>
            <option value="Details">Details</option>
			<option value="ActualDate">Actual Date</option>
			<option value="CoreTeamRep">Core Team Rep</option>
			<option value="Notify">Notify</option>
			<option value="BTO">BTO</option>
			<option value="CTO">CTO</option>
			<option value="APD">APD</option>
			<option value="CKK">CKK</option>
			<option value="GCD">GCD</option>
			<option value="EMEA">EMEA</option>
			<option value="LA">LA</option>
			<option value="NA">NA</option>
			<option value="OnStatusReport">On Status Report</option>
			<option value="AffectsCustomers">Affects Customers</option>
			<option value="LastModified">Last Modified</option>
			<option value="BTODate">BTO Date</option>
			<option value="CTODate">CTO Date</option>			
		</select>
	</td>
	<td valign="top">
		<input type="button" value="&gt;" id="cmdAdd" name="cmdAdd" style="WIDTH: 25px; HEIGHT: 24px" size="28" LANGUAGE="javascript" onclick="return cmdAdd_onclick()"><br>
		<input type="button" value="&lt;" id="cmdRemove" name="cmdRemove" style="WIDTH: 25px; HEIGHT: 24px" size="27" LANGUAGE="javascript" onclick="return cmdRemove_onclick()"><br><br>
		<input type="button" value="&gt;&gt;" id="cmdAddAll" name="cmdAddAll" LANGUAGE="javascript" onclick="return cmdAddAll_onclick()"><br>
		<input type="button" value="&lt;&lt;" id="cmdRemoveAll" name="cmdRemoveAll" LANGUAGE="javascript" onclick="return cmdRemoveAll_onclick()"><br>
	</td>
	<td>
		<select id="lstSelected" style="WIDTH: 165px; HEIGHT: 150px" size="2" name="lstSelected" multiple LANGUAGE="javascript" ondblclick="return lstSelected_ondblclick()" onkeydown="return lstSelected_onkeydown()"> 
			<option value="ID">ID</option>
			<option value="Product">Product</option>
			<option value="Status">Status</option>
			<option value="Owner">Owner</option>
			<option value="Summary">Summary</option>
			<option value="Description">Description</option>
			<option value="Created">Created</option>
			<option value="Submitter">Submitter</option>
			<option value="TargetDate">Target Date</option>
			<option value="Justification">Justification</option>
			<option value="Approvals">Approvals</option>
			<option value="Actions">Actions</option>
			<option value="Resolution">Resolution</option>
            <option value="Release">Release</option>
		</select>
	</td>

</tr>

	<tr>
		<td colspan="2"><font size="2"><b>Available Products:</b></font></td>
		<td><font size="2"><b>Selected Products:</b></font></td>
	</tr>
<tr>
	<td>
		<select id="lstAvailableProd" style="WIDTH: 165px; HEIGHT: 150px" size="2" name="lstAvailableProd" multiple LANGUAGE="javascript" ondblclick="return lstAvailableProd_ondblclick()"> 
		<%=strProducts%>
		</select>
	</td>
	<td valign="top">
		<input type="button" value="&gt;" id="cmdAddProd" name="cmdAddProd" style="WIDTH: 25px; HEIGHT: 24px" size="28" LANGUAGE="javascript" onclick="return cmdAddProd_onclick()"><br>
		<input type="button" value="&lt;" id="cmdRemoveProd" name="cmdRemoveProd" style="WIDTH: 25px; HEIGHT: 24px" size="27" LANGUAGE="javascript" onclick="return cmdRemoveProd_onclick()"><br><br>
		<input type="button" value="&gt;&gt;" id="cmdAddAllProd" name="cmdAddAllProd" LANGUAGE="javascript" onclick="return cmdAddAllProd_onclick()"><br>
		<input type="button" value="&lt;&lt;" id="cmdRemoveAllProd" name="cmdRemoveAllProd" LANGUAGE="javascript" onclick="return cmdRemoveAllProd_onclick()"><br>
	</td>
	<td>
		<select id="lstSelectedProd" style="WIDTH: 165px; HEIGHT: 150px" size="2" name="lstSelectedProd" multiple LANGUAGE="javascript" ondblclick="return lstSelectedProd_ondblclick()" onkeydown="return lstSelectedProd_onkeydown()"> 
		<%=strCurrentProduct%>
		</select>
	</td>

</tr>



<tr><td colspan="2"><input checked id="chkHeader" type="checkbox" name="chkHeader"><font size="2"> Include Column Headers</font></td></tr>
<tr>
	<td colspan="3" align="right">
		<hr>
		<input type="button" value="OK" id="cmdOK" name="cmdOK" LANGUAGE="javascript" onclick="return cmdOK_onclick()">
		<input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%Request.QueryString("pulsarplusDivId")%>')">
	</td>
</tr>
</table>
</font>
<textarea ID="Query" name="Query" style="Display:none" rows="2" cols="20">
</textarea>

</form>
<INPUT type="hidden" id=txtActionType name=txtActionType value=<%=request("ActionType")%>>
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=CurrentUserID%>">

</body>
</html>