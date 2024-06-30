<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">



<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

var DisplayedID;

var KeyString = "";
var AddID = 0;

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
		if (event.srcElement.name == "cboDeliverable")
			cboDeliverable_onchange();
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

/*function cmdEditDeliverable_onclick() {
	var strID = new Array();
	strID = window.showModalDialog("HFCNAdd.asp","","dialogWidth:535px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strID) != "undefined")
		{
		window.alert(strID[0]);
		}

	AddHFCN.txtDeliverableName.value = AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.selectedIndex].text;
}*/

function cmdAddDeliverable_onclick() {
	var strID = new Array();
	var i;
	var blnFound;
		
	strID = window.showModalDialog("HFCNAdd.asp?CatID=" + AddHFCN.cboCategory.value,"","dialogWidth:535px;dialogHeight:320px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strID) != "undefined")
		{
			if (strID[0] != "")
			{
			blnFound = false;
			for (i=0;i<AddHFCN.cboCategory.length;i++)
				{
					if (AddHFCN.cboCategory.options[i].value == strID[1])
					{
						AddHFCN.cboCategory.options[i].selected = true;
						cboCategory_onchange();
						blnFound = true;
						break;
					}
				
				}
			if (! blnFound)
				{
				window.alert("Unable to find selected category.");
				return;
				}
				
			//Add Deliverable to Master Delioverable List, current list, and select it
			blnFound = false;
			
			for (i=0;i<AddHFCN.cboDeliverable.length;i++)
				{
					if (String(AddHFCN.cboDeliverable.options[i].text).toUpperCase()==String(strID[0]).toUpperCase() )
						{
							blnFound = true;
							AddHFCN.cboDeliverable.options(i).selected=true;						
							break;
						}
				}
				
			if (! blnFound)
				{
					AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.length] = new Option(strID[0],"ADD" + AddID + ":" + strID[1]);
					AddHFCN.cboAllDeliverables.options[AddHFCN.cboAllDeliverables.length] = new Option( strID[0],"ADD" + AddID + ":" + strID[1]);
					AddHFCN.cboDeliverable.options(AddHFCN.cboDeliverable.length-1).selected=true;						
					AddID++;
				}
			}
		}
	AddHFCN.txtDeliverableName.value = AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.selectedIndex].text;
		

}

function cmdAdd_onclick() {
	var strSummary;
	var strID;
	var strLen;
	var i;
	var exists;
	
	strID = AddHFCN.txtOTS.value;
	strLen = strID.length;
	strID = "0000000"  + strID;
	strID = strID.slice(strID.length -strLen -(7-strLen ));
	
	exists = false;
	for (i=0;i<AddHFCN.lstOTS.length;i++)
		{
			strOption = AddHFCN.lstOTS.options[i].text
			if(strOption.slice(0,7) == strID)
				exists = true;
		}
	if (exists == false)
		{
			strSummary = window.showModalDialog("../OTS.asp?ID=" + AddHFCN.txtOTS.value + "","","dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No"); 
			if (typeof(strSummary) != "undefined")
				{
					if (strSummary == "")
						{
						window.alert("Invalid Observation Number.");
						AddHFCN.txtOTS.value = "";
						}
					else
						{
						AddHFCN.lstOTS.options[AddHFCN.lstOTS.length] = new Option(strSummary,strSummary.substr(0,7));			
						AddHFCN.txtOTS.value = "";
						}
				}
			else
				window.alert("Unable to connect to OTS.");
		}
		else
			AddHFCN.txtOTS.value = "";
	
}

function txtOTS_onkeypress() {
	if (window.event.keyCode == 13)
		{
			window.event.keyCode = 0;
			cmdAdd_onclick();
		}
}

function cmdRemove_onclick() {
	
	for (i=AddHFCN.lstOTS.length-1;i>=0;i--)
		{
			if(AddHFCN.lstOTS.options[i].selected)
				AddHFCN.lstOTS.options[i] = null;
		}
	
}


function lstOTS_ondblclick() {
	if (AddHFCN.lstOTS.selectedIndex != -1)
		{
		AddHFCN.lstOTS.options[AddHFCN.lstOTS.selectedIndex] = null;
		}		
}

function cboDeliverable_onchange() {
	var i;
	var NewCat;
	//PMText.innerHTML = AddHFCN.lstPM.options[AddHFCN.cboDeliverable.selectedIndex].text + "&nbsp;";
	//VendorText.innerHTML = AddHFCN.lstVendor.options[AddHFCN.cboDeliverable.selectedIndex].text + "&nbsp;";
	//DescText.innerHTML = AddHFCN.lstDescription.options[AddHFCN.cboDeliverable.selectedIndex].text + "&nbsp;";

	NewCat = GetRight(AddHFCN.cboDeliverable.value,":");

	AddHFCN.txtDeliverableName.value = AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.selectedIndex].text;
	
	AddHFCN.cboCategory.options(0).selected = true;
	for (i=0;i<AddHFCN.cboCategory.length;i++)
		{
		if (NewCat==AddHFCN.cboCategory.options[i].value)
			{
				AddHFCN.cboCategory.options(i).selected = true;
				if (i==0)
					cboCategory_onchange();
				return;
			}
		}
}

function GetLeft(MyString,Find){
	return MyString.substr(0,MyString.indexOf(Find,0));	

}
function GetRight(MyString,Find){
	return MyString.substr(MyString.indexOf(Find,0)+1);	

}

function cboCategory_onchange() {
	var ID;
	var CategoryID;
	var i;
	var OldID;
	
	OldID = AddHFCN.cboDeliverable.value;
	
	//Clear out deliverables
	for (i=AddHFCN.cboDeliverable.options.length-1;i>0;i--)	
		AddHFCN.cboDeliverable.options[i] = null;

	
	ID= AddHFCN.cboCategory.options(AddHFCN.cboCategory.selectedIndex).value;
	
	for (i=0;i<AddHFCN.cboAllDeliverables.length;i++)
		{
			if (ID == GetRight(AddHFCN.cboAllDeliverables.options(i).value,":") || ID == "") 
				{
				AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.length] = new Option(AddHFCN.cboAllDeliverables.options[i].text,AddHFCN.cboAllDeliverables.options[i].value);
				if (OldID == GetLeft(AddHFCN.cboAllDeliverables.options[i].value,":"))
					AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.length-1].selected=true;
				}
		}
	AddHFCN.txtDeliverableName.value = AddHFCN.cboDeliverable.options[AddHFCN.cboDeliverable.selectedIndex].text;
}

function AddDev_onclick() {
	var strID = new Array();
	strID = window.showModalDialog("http://16.81.19.70/mobilese/today/employee.asp","","dialogWidth:460px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strID) != "undefined")
		{
			AddHFCN.cboDeveloper.options[AddHFCN.cboDeveloper.length] = new Option(strID[1],strID[0]);
			AddHFCN.cboDeveloper.selectedIndex = AddHFCN.cboDeveloper.length - 1
		}
}

function chkProdInd_onclick() {
	if (AddHFCN.chkProdInd.checked)
		{
		ProdRow1.style.display = "none";
		ProdRow2.style.display = "none";
		ProdHR.style.display = "none";
		cmdProdRemoveAll_onclick();		
		}
	else
		{
		ProdRow1.style.display = "";
		ProdRow2.style.display = "";
		ProdHR.style.display = "";
		};
}

function lstAvailableProd_ondblclick() {
	if (AddHFCN.lstAvailableProd.options.selectedIndex > -1)
		{
		AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.length] = new Option(AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.options.selectedIndex].text,AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.options.selectedIndex].value);
		AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.options.selectedIndex] = null;
		}

}

function lstSelectedProd_ondblclick() {
	if (AddHFCN.lstSelectedProd.options.selectedIndex > -1)
		{
			AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.length] = new Option(AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.options.selectedIndex].text,AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.options.selectedIndex].value);
			AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.options.selectedIndex] = null;
		}
}

function cmdProdAdd_onclick() {
	var i;
	for (i=0;i<AddHFCN.lstAvailableProd.length;i++)
		{
			if(AddHFCN.lstAvailableProd.options[i].selected)
				{
					AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.length] = new Option(AddHFCN.lstAvailableProd.options[i].text,AddHFCN.lstAvailableProd.options[i].value);
				}
		}
	for (i=AddHFCN.lstAvailableProd.length-1;i>=0;i--)
		{
			if(AddHFCN.lstAvailableProd.options[i].selected)
					AddHFCN.lstAvailableProd.options[i] = null;
		}


}

function cmdProdAddAll_onclick() {
	var i;
	
	for (i=0;i<AddHFCN.lstAvailableProd.length;i++)
		{
			AddHFCN.lstSelectedProd.options[AddHFCN.lstSelectedProd.length] = new Option(AddHFCN.lstAvailableProd.options[i].text,AddHFCN.lstAvailableProd.options[i].value);
		}
	for (i=AddHFCN.lstAvailableProd.length-1;i>=0;i--)
		{
			AddHFCN.lstAvailableProd.options[i] = null;
		}

}

function cmdProdRemove_onclick() {
	var i;
	
	for (i=0;i<AddHFCN.lstSelectedProd.length;i++)
		{
			if(AddHFCN.lstSelectedProd.options[i].selected)
				{
					AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.length] = new Option(AddHFCN.lstSelectedProd.options[i].text,AddHFCN.lstSelectedProd.options[i].value);
				}
		}
	for (i=AddHFCN.lstSelectedProd.length-1;i>=0;i--)
		{
			if(AddHFCN.lstSelectedProd.options[i].selected)
					AddHFCN.lstSelectedProd.options[i] = null;
		}

}

function cmdProdRemoveAll_onclick() {
	var i;
	
	for (i=0;i<AddHFCN.lstSelectedProd.length;i++)
		{
			AddHFCN.lstAvailableProd.options[AddHFCN.lstAvailableProd.length] = new Option(AddHFCN.lstSelectedProd.options[i].text,AddHFCN.lstSelectedProd.options[i].value);
		}
	for (i=AddHFCN.lstSelectedProd.length-1;i>=0;i--)
			AddHFCN.lstSelectedProd.options[i] = null;

}

function ProdText_onclick() {
	if (AddHFCN.chkProdInd.checked)
		AddHFCN.chkProdInd.checked = false;
	else
		AddHFCN.chkProdInd.checked = true;
	chkProdInd_onclick();
}


function FindProductChanges() {

	var strAddProducts;
	var strAddProductsNames;
	var strRemoveProducts;
	var strRemoveProductsNames;
	
	var i;
	strAddProducts = "";
	strAddProductsNames = "";
	strRemoveProducts = "";
	strRemoveProductsNames = "";


	if (lblProductsLoaded.innerText != "" && AddHFCN.lstSelectedProd.length == 0 )
		{
			//Don't add any products
		strAddProducts = "";
		strAddProductsNames = "";		
		}
	else

		for (i=0;i<AddHFCN.lstSelectedProd.length;i++)
			{
				//Check for add here
				if (instrAt("," + lblProductsLoaded.innerText,"," + AddHFCN.lstSelectedProd.options(i).value + ",")< 0)
					{
						strAddProducts = strAddProducts + AddHFCN.lstSelectedProd.options(i).value + ",";
						strAddProductsNames = strAddProductsNames + "," + AddHFCN.lstSelectedProd.options(i).text;
					}
			}

					
		
	if (lblProductsLoaded.innerText == "" && AddHFCN.lstSelectedProd.length > 0 )
		{
			//None Removed
			strRemoveProducts = ""
			strRemoveProductsNames = ""
		}
	else
		{
		for (i=0;i<AddHFCN.lstAvailableProd.length;i++)
			{
				//Check for remove here
				if (instrAt("," + lblProductsLoaded.innerText,"," + AddHFCN.lstAvailableProd.options(i).value + ",")>=0)
					{
					strRemoveProducts = strRemoveProducts + AddHFCN.lstAvailableProd.options(i).value + ",";
					strRemoveProductsNames = strRemoveProductsNames + AddHFCN.lstAvailableProd.options(i).text + ",";
					}
			}
		}		
	AddHFCN.AddProducts.value =  strAddProducts;
	AddHFCN.AddProductsNames.value =  strAddProductsNames;
	AddHFCN.RemoveProducts.value =  strRemoveProducts;
	AddHFCN.RemoveProductsNames.value =  strRemoveProductsNames;
}

function instrAt(MyString,Find){
	return MyString.indexOf(Find,0);
}

function FindOTSChanges() {


	var strAddOTS;
	var strRemoveOTS;
	
	var i;
	
	var strOTS = "";
	for(i=0;i<AddHFCN.lstOTS.length;i++)
		strOTS = strOTS + AddHFCN.lstOTS.options[i].text + "\r";
	AddHFCN.OTSSummary.value = strOTS;	
	
	
	strAddOTS = "";
	strRemoveOTS = "";
		
	strAddOTS = "";
	for (i=0;i<AddHFCN.lstOTS.length;i++)
		{
			//Check for add here
			if (instrAt("," + txtOTSLoaded.value ,"," + AddHFCN.lstOTS.options(i).value + ",")< 0)
				{
					strAddOTS = strAddOTS + AddHFCN.lstOTS.options(i).value + ",";
				}
		}
	
	strRemoveOTS = txtOTSLoaded.value;
	for (i=0;i<AddHFCN.lstOTS.length;i++)
		{
			//Check for remove here
			if (instrAt("," + txtOTSLoaded.value,"," + AddHFCN.lstOTS.options(i).value + ",")>=0)
				strRemoveOTS = replace(strRemoveOTS, AddHFCN.lstOTS.options(i).value + ",","");
		}
	AddHFCN.OTS2ADD.value =  strAddOTS;
	AddHFCN.OTS2REMOVE.value =  strRemoveOTS;
	
}


function cmdAddVendor_onclick() {
	var NewVendor = window.prompt("Please enter the name of the new vendor:","");
	var i;
	var buffer;
	var newitem;
	var blnFound;
	var strNewID;
		
	if (NewVendor != "" && NewVendor != null)
		{
			//See if it is already in the list
			newitem = NewVendor;
			newitem = newitem.toLowerCase();
			blnFound = false;
			for(i=0;i<AddHFCN.cboVendor.length;i++)	
				{
				buffer = AddHFCN.cboVendor.options[i].text;
				buffer = buffer.toLowerCase();
				if (buffer==newitem)
					{
					AddHFCN.cboVendor.selectedIndex = i;
					blnFound = true;
					}
				}
			//if not, add it and select it.
			if (blnFound == false)
				{
				strNewID = window.showModalDialog("../saveVendor.asp?Name=" + NewVendor,"","dialogHide:1;dialogWidth:300px;dialogHeight:20px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
				if (typeof(strNewID) != "undefined")
					{
					AddHFCN.cboVendor.options[AddHFCN.cboVendor.length] = new Option(NewVendor,strNewID);
					AddHFCN.cboVendor.selectedIndex = AddHFCN.cboVendor.length - 1			
					}
				else
					window.alert("Unable to add new vendor.");
				}
		}
}



function window_onload() {
	cboDeliverable_onchange();
}


//-->
</script>

</head>
<body bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">

<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

<%
	'Create Database Connection
	dim cn 
	dim rs
	dim cm
	dim p
	dim strSQL
	dim strEmployeeList
	dim strDelName
	'dim strDescription
	dim strCatID
	'dim	strVendorID
	'dim strDevmanID
	dim strTypeID
	dim strProductList
	dim workflowID
	dim strDisplayedID
	dim strDeliverables
	dim strSelectedProducts	
	dim strDisplayProduct
	dim strCategories
	dim strVendor
	'dim strDelDescription
	dim strDelCategory
	'dim strDelPM
	dim strAllDeliverables
	dim CurrentUser
	dim strUserID
	dim strAvailableProducts
	dim strProductsLoaded
	dim strDependencies
	dim strUserEmail
	dim strDisplayedRootID
	dim strCategoryName
	dim strUserName
		
	strDelName = ""
	'strDescription = ""
	strCatID = ""
	'strVendorID = ""
	'strDevmanID = ""
	strTypeID = ""
	strProductList = ""
	strWorkflowID = ""
	strDisplayProduct = ""
	strDeliverables = ""
	strCategories = ""
	strVendor = ""
	'strDelDescription = "<option selected></option>"
	strDelCategory = "<option selected></option>"
	'strDelPM = "<option selected></option>"
	strAllDeliverables = ""
	strAvailableProducts = ""
	strDependencies = ""
	strUserEmail = ""
	strDisplayedRootID = ""
	strcategoryName = ""
	strUserName = ""
	
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
	
	CurrentUserID = 0
	'rs.Open "spGetUserInfo '" & currentuser & "'",cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		strUserID = rs("ID")
		strUserEmail = rs("Email") & ""
		strUsername = rs("Name") & ""
	else
		strUserID = ""
		strUserEmail = ""
		strUserName = "the developer"
	end if
	
	rs.Close
		
	rs.Open "spGetHFCNDeliverables",cn,adOpenForwardOnly
	strDeliverables = "<option selected></option>"
	do while not rs.EOF
		if trim(Request("RootID")) = trim(rs("ID")) then
			strDeliverables = strDeliverables & "<option selected value=""" & rs("ID") & ":" & rs("CategoryID") & """>" & rs("Name") & "</option>"
		else
			strDeliverables = strDeliverables & "<option value=""" & rs("ID") & ":" & rs("CategoryID") & """>" & rs("Name") & "</option>"
		end if
		strAllDeliverables = strAllDeliverables & "<option value=""" & rs("ID") & ":" & rs("CategoryID") & """>" & rs("Name") & "</option>"
	'	strDelVendor = strDelVendor & "<option value=""" & rs("VendorID") & """>" & rs("Vendor") & "</option>" 
	'	strDelDescription = strDelDescription & "<option>" & rs("Description") & "</option>"
		strDelCategory = strDelCategory & "<option value=""" & rs("CategoryID") & """>" & rs("Category") & "</option>"
	'	strDelPM = strDelPM & "<option value=""" & rs("DevManagerID") & """>" & rs("DevManager") & "</option>"
		rs.MoveNext
	loop
	rs.Close
	
	strCategories = ""
		
	strSQL = "spListHFCNCategories"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		strCategories = strCategories & "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close
	
	strSQL = "spGetProducts4Web"
	rs.Open strSQL,cn,adOpenForwardOnly
	strSelectedProducts = ""
	strAvailableProducts = ""
	strProductList = ""
	strProductsLoaded = ""
	do while not rs.EOF
		if instr("," & strProductList,"," & rs("ID") & ",") = 0 then
			if rs("Division") & "" = "1" and rs("ID") <> 100 then '(rs("Active") & "" = "1" or rs("Sustaining") & "" = "1") and rs("Division") & "" = "1" and rs("ID") <> 100 then
				strAvailableProducts = strAvailableProducts &  "<Option value=" & rs("ID") & ">" & rs("Name") & " " &  rs("Version") & "</OPTION>"
			end if
		else
			strSelectedProducts = strSelectedProducts &  "<Option value=" & rs("ID") & ">" & rs("Name") & " " &  rs("Version") & "</OPTION>"
			strProductsLoaded = strProductsLoaded & rs("ID") & ","
		end if
		rs.MoveNext
	loop
	rs.Close
	
	
%>
<!--<h3><label ID="lblTitle">Submit HFCN</label></h3>

<font size="2">
<label ID="lblInstructions"><font face=verdana size=1 color=red><b>*</b></font><font face=verdana size=2> = Required Field</font></label>
</font>
<br><br>-->
<FONT size=2 color=red face=verdana><b>Warning:  This page is now obsolete (11/30/2005).  We will be sending out new instructions for how to enter HFCNs this evening.  Please only use this page to release deliverables for emergency releases only.</b></font><BR><BR><BR>
<form id="AddHFCN" method="post" action="HFCNSave.asp">
<input style="Display:none" type="text" id="ID" name="ID" value="<%=request("ID")%>">
<%if request("ID") = "" then%>
<font size=2 face=verdana><b>Step 1: Select a Deliverable</b></font><BR>
<%else%>
<font size=2 face=verdana><b>Deliverable Information</b></font><BR>
<%end if%>

<table ID="tabRoot" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="150" nowrap><b>Deliverable:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%">
		<SELECT id=cboDeliverable name=cboDeliverable style="WIDTH: 430px" LANGUAGE=javascript onchange="return cboDeliverable_onchange()" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<%=strDeliverables%>
		</SELECT>
		<SELECT id=cboAllDeliverables name=cboAllDeliverables style="display:none">
			<%=strAllDeliverables%>
		</SELECT>
		&nbsp;<input type="button" value="Add" id="cmdAddDeliverable" name="cmdAddDeliverable" LANGUAGE=javascript onclick="return cmdAddDeliverable_onclick()" tabindex="-1">
		<!--&nbsp;<input type="button" value="Edit" id="cmdEditDeliverable" name="cmdEditDeliverable" LANGUAGE=javascript onclick="return cmdEditDeliverable_onclick()" tabindex=-1>-->
		</td>
	</tr>
    <tr>
		<td nowrap><b>Category:</b> <font color="#ff0000" size="1">*</font></td>
		<td>
			<select id="cboCategory" name="cboCategory" style="WIDTH: 430px" LANGUAGE=javascript onchange="return cboCategory_onchange()" onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">			
				<option selected></option>
				<%=strCategories%>
			</select>
			<INPUT type="hidden" id=txtCategory name=txtCategory>
			
				<!--<font size=1 color=blue face=verdana>&nbsp;(Select Category to filer Deliverable List)</font>-->
			</td>
    </tr>
<!--    <tr>
		<td nowrap><b>Dev.&nbsp;Manager:</b></td>
		<td><font ID=PMText face=verdana size=2>&nbsp;</font></td>
    </tr>
    
    <tr>
		<td nowrap><b>Vendor:</b></td>
		<td colspan="10"><font ID=VendorText face=verdana size=2>&nbsp;</font></td>
	</tr>
    
    <tr>
		<td nowrap valign="top"><b>Description:</b></td>
		<td colspan="10"><font ID=DescText face=verdana size=2>&nbsp;</font></td>
	</tr>-->
	
    
</TABLE>
<!--<SELECT style="WIDTH=20%;HEIGHT=140px;Display:none" size=5 id=lstVendor name=lstVendor>
	<%=strDelVendor%>
</SELECT>-->
<!--<SELECT style="WIDTH=20%;HEIGHT=140px;Display:none" size=5 id=lstDescription name=lstDescription>
	<%=strDelDescription%>
</SELECT>-->
<SELECT style="WIDTH=20%;HEIGHT=140px;Display:none" size=5 id=lstCategory name=lstCategory>
	<%=	strDelCategory%>
</SELECT>
<!--<SELECT style="WIDTH=20%;HEIGHT=140px;Display:none" size=5 id=lstPM name=lstPM>
	<%=strDelPM%>
</SELECT>-->
<BR>
<%if request("ID") = "" then%>
<font size=2 face=verdana><b>Step 2: Enter Version Information</b></font><BR>
<%else%>
<font size=2 face=verdana><b>Version Information</b></font><BR>
<%end if%>
<table ID="tabVersion" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="150" nowrap><b>Title:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT type="text" id=txtTitle name=txtTitle maxlength=255 style="Width:100%">
		</td>
	</tr>
	<tr>
		<td width="150" nowrap><b>Version (Rev):</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT type="text" id=txtVersion name=txtVersion maxlength=20  style="WIDTH: 190px">
		</td>
	</tr>
	<tr>
		<td width="150" nowrap><b>Vendor:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%">
			<SELECT id=cboVendor name=cboVendor style="WIDTH: 190px" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
				<OPTION selected></OPTION>
		<%
		strSQL = "spGetVendorList"
		rs.Open strSQL,cn,adOpenForwardOnly
		do while not rs.EOF
			Response.Write "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			rs.MoveNext
		loop
		rs.Close
		%>
		</select>&nbsp;<input type="button" value="Add" id="cmdAddVendor" name="cmdAddVendor" LANGUAGE="javascript" onclick="return cmdAddVendor_onclick()">				
			<INPUT type="hidden" id=txtVendor name=txtVendor>
			
		</td>
	</tr>
	<tr>
		<td width="150" nowrap><b>Vendor&nbsp;Version:</b>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT type="text" id=txtVendorVersion name=txtVendorVersion  style="WIDTH: 190px" maxlength=30>
		</td>
	</tr>
    <tr>
		<td nowrap><b>Developer:</b> <font color="#ff0000" size="1">*</font></td>
		<td><select id="cboDeveloper" style="WIDTH: 190px" name="cboDeveloper" LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
		 <option selected></option> <!--LANGUAGE=javascript onmouseover="return cboDeveloper_onmouseover()" onmouseout="return cboDeveloper_onmouseout()"-->
      
      
		<%
		strSQL = "spGetEmployees"
		rs.Open strSQL,cn,adOpenForwardOnly
		do while not rs.EOF
			if trim(strUserID) = trim(rs("ID") & "") then
				Response.Write "<Option selected Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			else
				Response.Write "<Option Value=" & rs("ID") & ">" & rs("Name") & "</OPTION>"
			end if
			rs.MoveNext
		loop
		rs.Close
		%>
      
      
      </select>
      <input type="button" value="Add" id="cmdAddDev" name="cmdAddDev" LANGUAGE="javascript" onclick="return AddDev_onclick()">
      </td>
    </tr>
	
	<tr>
		<td width="150" nowrap><b>Send Email To:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT style="width:100%" type="text" id=txtEmail name=txtEmail value="" maxlength=255>
		</td>
	</tr>
	<tr>
		<td width="150" nowrap><b>Location:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td colspan="10" width="100%"><INPUT  style="width:100%" type="text" id=txtLocation name=txtLocation maxlength=1000>
		</td>
	</tr>
	<tr valign=top>
		<td><b>Observations&nbsp;Fixed:&nbsp;</b></td><td><input type="text" style="Width:100" id="txtOTS" name="txtOTS" LANGUAGE="javascript" onkeypress="return txtOTS_onkeypress()"> <input type="button" value="Add" id="cmdAdd" name="cmdAdd" LANGUAGE="javascript" onclick="return cmdAdd_onclick()">
		<input type="button" value="Remove" id="cmdRemove" name="cmdRemove" LANGUAGE="javascript" onclick="return cmdRemove_onclick()">
		<br>
		<select size="2" style="Width=100%;Height=108px;" id="lstOTS" name="lstOTS" multiple LANGUAGE="javascript" ondblclick="return lstOTS_ondblclick()">
		<%=strSelectedOTS%>
		</select>
		</td>
	</tr>
	<tr>
		<td valign=top><b>Other&nbsp;Changes:</b></td>
		<td>
			<textarea id="txtChanges" style="WIDTH: 100%; HEIGHT: 90px" name="txtChanges" rows="10" cols="62"><%=strChanges%></textarea>
		</td>
	</tr>
	<tr>
		<td valign=top><b>Dependencies:</b></td>
		<td>
			<textarea id="txtDependencies" style="WIDTH: 100%; HEIGHT: 90px" name="txtDependencies" rows="10" cols="62"><%=strDependencies%></textarea>
		</td>
	</tr>
	<tr>
		<td valign=top><b>Special Notes:</b></td>
		<td>
			<textarea id="txtSpecial" style="WIDTH: 100%; HEIGHT: 90px" name="txtSpecial" rows="10" cols="62"></textarea>
		</td>
	</tr>
  <tr>
		<td width="150" nowrap vAlign="top"><b>Products:</b>&nbsp;<font color="red" size="1">*</font></td>		<td colspan="10" width="100%">
      <table cellSpacing="1" cellPadding="1" width="460" border="0">
		<tr><td colspan="4"><input <%=strProductInd%> type="checkbox" id="chkProdInd" name="chkProdInd" LANGUAGE="javascript" onclick="return chkProdInd_onclick()"><label id="ProdText" LANGUAGE="javascript" onclick="return ProdText_onclick()">Product Independent</label><hr style="Display:<%=strDisplayProdLists%>" ID="ProdHR" color="Tan"></td></tr>
		<tr ID="ProdRow1" style="Display:<%=strDisplayProdLists%>"><td><strong>Not Supported</strong></td><td></td><td><strong>Supported</strong></td><tr>       
        <tr ID="ProdRow2" style="Display:<%=strDisplayProdLists%>">
          <td width="220"><select id="lstAvailableProd" style="WIDTH: 213px; HEIGHT: 120px" size="2" name="lstAvailableProd" LANGUAGE="javascript" ondblclick="return lstAvailableProd_ondblclick()" multiple><%=strAvailableProducts%> 
            </select></td>
          <td width="20"><input id="cmdProdADD" style="WIDTH: 25px" type="button" width="25" value="&gt;" name="cmdProdADD" LANGUAGE="javascript" onclick="return cmdProdAdd_onclick()"><br><input id="cmdProdRemove" style="WIDTH: 25px" type="button" width="20" value="&lt;" name="cmdProdRemove" LANGUAGE="javascript" onclick="return cmdProdRemove_onclick()"><br><br><input id="cmdProdAddAll" type="button" value="&gt;&gt;" name="cmdProdAddAll" LANGUAGE="javascript" onclick="return cmdProdAddAll_onclick()"><br><input id="cmdProdRemoveAll" type="button" value="&lt;&lt;" name="cmdProdRemoveAll" LANGUAGE="javascript" onclick="return cmdProdRemoveAll_onclick()"></td>
          <td width="220"><select id="lstSelectedProd" style="WIDTH: 213px; HEIGHT: 120px" size="2" name="lstSelectedProd" LANGUAGE="javascript" ondblclick="return lstSelectedProd_ondblclick()" multiple> 
              <%=strSelectedProducts%></select></td></tr></div></table></td></tr>
 	

	
</table>

<input type="hidden" id="OTS2ADD" name="OTS2ADD" value>
<input type="hidden" id="OTS2REMOVE" name="OTS2REMOVE" value>
<input type="hidden" id="txtDeliverableName" name="txtDeliverableName">
<input type="hidden" id="AddProducts" name="AddProducts">
<input type="hidden" id="AddProductsNames" name="AddProductsNames">
<input type="hidden" id="RemoveProducts" name="RemoveProducts">
<input type="hidden" id="RemoveProductsNames" name="RemoveProductsNames">
<input type="hidden" id="OTSSummary" name="OTSSummary">
<input type="hidden" id="CurrentUserEmail" name="CurrentUserEmail" value="<%=strUserEmail%>">
<input type="hidden" id="CurrentUserName" name="CurrentUserName" value="<%=strUserName%>">


</form>

<input type="hidden" id="txtOTSLoaded" name="txtOTSLoaded">
<label id="lblDisplayedID" style="Display:none"><%=request("ID")%></label>
<label id="lblProductsLoaded" style="Display:none"><%=strProductsLoaded%></label>

	<select style="Display:none" size="2" id="lstRoots" name="lstRoots">
	<%
	strSQL = "spGetDelRoot"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
		if lcase(trim(strDelName)) <> lcase(trim(rs("Name"))) then
			Response.Write "<OPTION>" & replace(rs("name"),"""","&quot;") & "</OPTION>"
		end if
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
	set cn = nothing
	%>

</select>	

</body>
</html>