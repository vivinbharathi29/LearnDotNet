<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link href="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
<script src="<%= Session("ApplicationRoot") %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%= Session("ApplicationRoot") %>/includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var CurrentTab=0;
var TabArray = Array("Tab1","Tab2","Tab3")
var FiltersDirty = true;

function cmdAdd_onclick() {
	var strResult;
	strResult = window.showModalDialog("../../Email/AddressBook.asp?AddressList=" + frmUpdate.txtCC.value,"","dialogWidth:400px;dialogHeight:600px;edge: Sunken;maximize:Yes;center:Yes; help: No;resizable: Yes;status: No"); 
	if (typeof(strResult) != "undefined")
		//if (frmUpdate.txtNotify.value == "")
			frmUpdate.txtCC.value = strResult;
		//else
		//	frmUpdate.txtNotify.value = frmUpdate.txtNotify.value + ";" + strResult;

}


function SelectProducts(ProdList){

	var ProdArray = ProdList.split(",");
	var i;
	var j;
	var UpdateCount = 0;
	var ProdCount=0;

	FiltersDirty=true;
	
	for (i=0;i<ProdArray.length;i++)
		{
		ProdCount++;
		for (j=0;j<frmUpdate.lstProducts.length;j++)
			{
			if (ProdArray[i] == frmUpdate.lstProducts[j].value)
				{
				frmUpdate.lstProducts[j].selected=true;
				UpdateCount++;
				break;
				}
			}
		}
	if (UpdateCount == ProdCount && UpdateCount==1)
		alert("Automatically selected the product defined for this cycle.\r\rPlease verify the product list was updated correctly.");
	else if (UpdateCount == ProdCount)
		alert("Automatically selected all " + UpdateCount + " products defined for this cycle.\r\rPlease verify the product list was updated correctly.");
	else if (ProdCount==1)
		alert("Unable to find the only product defined for this cycle.  No products have been selected.");
	else
		alert("Automatically selected only 1 of the " + ProdCount + " products defined for this cycle.\r\rPlease verify the product list was updated correctly.");
}

function PopulateVersions(){
//	alert(CatID +":" + frmUpdate.cboVendor.options[frmUpdate.cboVendor.selectedIndex].value);
    var ActionID = $('input[name=optRestriction]:checked').val();
	var RootID = $("#txtRootID").val();
	var CatID = $("#txtCategoryID").val();

	var arrayListValue = ListValue().split(',');
	var PulsarProductIDList = "";
	var LegacyProductIDList = "";

	for (var i = 0; i < arrayListValue.length; i++) {
	    var arrayID = arrayListValue[i].split(';');

	    if (arrayID[1] > 0) {
	        if (PulsarProductIDList == "") {
	            PulsarProductIDList = ":" + arrayID[0];
	        }
	        else {
	            PulsarProductIDList += ":";
	            if (PulsarProductIDList.indexOf(":" + arrayID[0] + ":") == -1)
	                PulsarProductIDList += arrayID[0];
	            else
	                PulsarProductIDList = PulsarProductIDList.slice(0, PulsarProductIDList.length - 1);
	        }
	    }
	    else {
	        if (LegacyProductIDList == "")
	            LegacyProductIDList = ":" + arrayID[0];
	        else {
	            LegacyProductIDList += ":";
	            if (LegacyProductIDList.indexOf(":" + arrayID[0] + ":") == -1)
	                LegacyProductIDList += arrayID[0];
	            else
	                LegacyProductIDList = LegacyProductIDList.slice(0, LegacyProductIDList.length - 1);
	        }
	    }
	}

	PulsarProductIDList = PulsarProductIDList.slice(1, PulsarProductIDList.length);
	LegacyProductIDList = LegacyProductIDList.slice(1, LegacyProductIDList.length);
    //jsrsExecute("RestrictVersionReleasesRSget.asp", myCallback1, "VersionList", Array(CatID.toString(), ProductIDList, $("#cboVendor").val(), ActionID, RootID, ReleaseIDList));
	var ajaxurl = "RestrictVersionReleasesRSget.asp?PulsarProductIDList=" + PulsarProductIDList + "&LegacyProductIDList=" + LegacyProductIDList + "&VendorID=" + $("#cboVendor").val() + "&RootID=" + RootID + "&ActionID=" + ActionID + "&CatID=" + CatID.toString();
	$.ajax({
	    url: ajaxurl,
	    type: "POST",
	    success: function (data) {
	        myCallback1(data);
	    },
	    error: function (xhr, status, error) {
	        alert(error);
	    }

	});
}

function myCallback1(returnstring) {
	if (FiltersDirty)
	{
		VersionListDisplay.innerHTML = returnstring;
	}

	FiltersDirty = false;
	var strVendor = $("#cboVendor option:selected").text();
	var strAction;
	var strAction2;

	if ($('input[name=optRestriction]:checked').val() == 0)
	{
		strAction = "Restrict "
		strAction2 = "have been restricted"
    }
	else
	{
		strAction = "Unrestrict "
		strAction2 = "are no longer restricted"
	}

	var intCategoryID = $("#txtCategoryID").val();
	if (intCategoryID=="2")
		strCategory = "hard drives";
	else if (intCategoryID == "3")
		strCategory = "memory modules";
	else if (intCategoryID == "15")
		strCategory = "optical drives";
	else if (intCategoryID == "71")
		strCategory = "batteries";
	else if (intCategoryID=="33")
		strCategory = "input devices";
	else if (intCategoryID=="36")
		strCategory = "AC adapters";
	else if (intCategoryID=="40")
		strCategory = "floppy drives";
	else if (intCategoryID=="131")
		strCategory = "panels";
	else
		strCategory = $("#txtCategory").val();
	
	ReviewAction.innerText = strAction + strVendor + " " + strCategory + " on the selected products."
	ReviewProducts.innerText = replace(ListText(), ",", ", ");
	$("#txtProductNames").val(replace(ListText(), ",", ", "));
	$("#txtActionText").val("The following " + strVendor + " " + strCategory + " " + strAction2 + " on the selected products.");
	if (VersionListDisplay.innerText == "No Versions Found.")
		window.parent.frames["LowerWindow"].cmdFinish.disabled=true;

}

function replace(sampleStr,replaceChars,replaceWith){

	var replaceRegExp = new RegExp(replaceChars,'g');

	return sampleStr.replace(replaceRegExp,replaceWith);
}


function ListValue(){
	var i;
	var strList;
	strList = "";
	$('#lstProducts :selected').each(function (i, selected) {
	    if (strList == "")
	        strList = $(selected).val() + ";" + $(selected).attr("ReleaseID");
	    else
	        strList = strList + "," + $(selected).val() + ";" + $(selected).attr("ReleaseID");
	});
	return strList;
}

function ListText(){
	var i;
	var strList;
	strList = "";
	$('#lstProducts :selected').each(function (i, selected) {
	    if (strList == "")
	        strList = $(selected).text();
	    else
	        strList = strList + "," + $(selected).text();
	});
	return strList;
}

function ToggleVersions(){
	var i;
	if (typeof(frmUpdate.chkVersions.length) == "undefined")
		frmUpdate.chkVersions.checked= frmUpdate.chkAllVersions.checked;
	else
		for (i=0;i<frmUpdate.chkVersions.length;i++)
			frmUpdate.chkVersions[i].checked= frmUpdate.chkAllVersions.checked;
}

function ValidateTab(TabID){
	var strSuccess = true;
	var blnFound;
	var i;
	
	if (TabID==0) //Moving to Tab 1 (second tab)
		{
		//No Validation Required
		strSuccess = true;
		}
	else if (TabID==1) //Moving to Tab 2 (third tab)
	{
		if (ListText()=="")
		{
			alert("You must select at least one product.");
			$("#lstProducts").focus();
			strSuccess = false;
        }
		else if ($("#cboVendor option:selected").index() < 1)
		{
			alert("You must select a vendor.");
			$("#lstProducts").focus();
			strSuccess = false;
        }
		else
			PopulateVersions();
		}
	else if (TabID == 2) //Saving
	{
	    if (VersionListDisplay.innerText == "No Versions Found.") {
	        alert("Your filters did not find any versions to update.");
	        strSuccess = false;
	    }
	    else {
	        blnFound = false;
	        if (typeof ($('input[type="checkbox"]:checked').length) == "undefined") {
	            if ($('input[name="chkVersions"]').is(':checked'))
	                blnFound = true;
	        }
	        else {
	            if ($('input[type="checkbox"]:checked').length > 0)
	                blnFound = true;
	        }

	        if (!blnFound) {
	            alert("You must select at least one deliverable version.");
	            strSuccess = false;
	        }
	        else if (!VerifyEmail($("#txtCC").val())) {
	            alert("The email address specified in the CC field does not appear to be valid.");
	            $("#txtCC").focus();
	            strSuccess = false;
	        }
	        else {
	            var strProductList = ListValue();
	            var arrProductList = strProductList.split(',');
	            var strLegacyProductList = "";
	            var strPulsarProductList = "";
	            var strReleaseList = "";
	            for (var i in arrProductList) {
	                var arrIDs = arrProductList[i].split(';');
	                if (arrIDs[1] == 0) {
	                    if (strLegacyProductList == "") {
	                        strLegacyProductList += ",";
	                        strLegacyProductList += arrIDs[0];
	                    }
	                    else {
	                        strLegacyProductList += ",";
	                        if (strLegacyProductList.indexOf("," + arrIDs[0] + ",") == -1)
	                            strLegacyProductList += arrIDs[0];
	                        else
	                            strLegacyProductList = strLegacyProductList.slice(0, strLegacyProductList.length - 1);
	                    }
	                }
	                else {
	                    if (strPulsarProductList == "") {
	                        strPulsarProductList += ",";
	                        strPulsarProductList += arrIDs[0];
	                    }
	                    else {
	                        strPulsarProductList += ",";
	                        if (strPulsarProductList.indexOf("," + arrIDs[0] + ",") == -1)
	                            strPulsarProductList += arrIDs[0];
	                        else 
	                            strPulsarProductList = strPulsarProductList.slice(0, strPulsarProductList.length - 1);
	                    }	                    

	                    if (strReleaseList != "") {
	                        strReleaseList += ",";
	                        strReleaseList += arrIDs[1];
	                    }
	                    else {
	                        strReleaseList += ",";
	                        if (strReleaseList.indexOf("," + arrIDs[1] + ",") == -1)
	                            strReleaseList += arrIDs[1];
	                        else
	                            strReleaseList = strReleaseList.slice(0, strReleaseList.length - 1);
	                    }
	                }
	            }

	            strPulsarProductList = strPulsarProductList.slice(1, strPulsarProductList.length);
	            strLegacyProductList = strLegacyProductList.slice(1, strLegacyProductList.length);
	            strReleaseList = strReleaseList.slice(1, strReleaseList.length);

	            $("#txtPulsarProductList").val(strPulsarProductList);
	            $("#txtLegacyProductList").val(strLegacyProductList);
	            $("#txtReleaseList").val(strReleaseList);

	            frmUpdate.submit();
	        }
	    }
	}
	return strSuccess;
}


function VerifyEmail(src) {
 	 var blnOK=true;
 	 var i;
 	 var AddressArray;
     var emailReg = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
     var regex = new RegExp(emailReg);
     AddressArray = src.split(";");
     for (i=0;i<AddressArray.length;i++)
		if (! regex.test(trim(AddressArray[i])))
			if ( ! (i == AddressArray.length -1 && trim(AddressArray[i])== "" ) )
				blnOK=false;
     return blnOK;
  }
  
function ButtonClicked(ButtonID){
	if (ButtonID==1)
		{
		window.parent.opener = self;
		window.parent.close();
		}
	else if (ButtonID==3)//Next
		{
		
		if (ValidateTab(CurrentTab))
			{
			if (CurrentTab < 2)
				{
				window.document.all(TabArray[CurrentTab]).style.display="none";
				CurrentTab = CurrentTab + 1
				window.document.all(TabArray[CurrentTab]).style.display="";
				if (CurrentTab==2)
					{
					window.parent.frames["LowerWindow"].cmdNext.disabled=true;
					window.parent.frames["LowerWindow"].cmdFinish.disabled=false;
					}
				}

			window.parent.frames["LowerWindow"].cmdPrevious.disabled=false;
			}
		
		}
	else if (ButtonID==2)//Previous
		{
		if (CurrentTab > 0)
			{
			window.document.all(TabArray[CurrentTab]).style.display="none";
			CurrentTab = CurrentTab - 1
			window.document.all(TabArray[CurrentTab]).style.display="";
			if (CurrentTab==0)
				window.parent.frames["LowerWindow"].cmdPrevious.disabled=true;

			}


		window.parent.frames["LowerWindow"].cmdNext.disabled=false;
		window.parent.frames["LowerWindow"].cmdFinish.disabled=true;

		}
	else if (ButtonID==4)//Finish
	{
		if (ValidateTab(CurrentTab))
		{
			if (CurrentTab == 2)
			{
		        window.parent.frames["LowerWindow"].cmdPrevious.disabled=true;
				window.parent.frames["LowerWindow"].cmdNext.disabled=true;
				window.parent.frames["LowerWindow"].cmdFinish.disabled=true;
				window.parent.frames["LowerWindow"].cmdCancel.value="Close";
				window.parent.frames["LowerWindow"].cmdCancel.disabled=false;
			}
		}	
	}
}

function lstProducts_onchange() {
	FiltersDirty=true;
}

function cboVendor_onchange() {
	FiltersDirty=true;
}

function optRestriction_onchange() {
	FiltersDirty=true;

}

function optRestriction_onclick() {
	FiltersDirty=true;

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

function GroupHeaderClicked(ID) {
    if (document.all("ProgramGroupList" + ID).style.display == "none")
        document.all("ProgramGroupList" + ID).style.display = "";
    else
        document.all("ProgramGroupList" + ID).style.display = "none";
}

function SwitchRelease(RootID, VersionID, ProductID, ReleaseID) {
    document.location = "RestrictMainPulsar.asp?RootID=" + RootID + "&ProductID=" + ProductID + "&VersionID=" + VersionID + "&ReleaseID=" + ReleaseID;
    window.parent.RepositionPopup();
}
//-->
</SCRIPT>
</HEAD>
<STYLE>
BODY
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana
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
h3
{
    font-family: Verdana;
    font-size: small;
    color: Black;
}
td
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana
}
.VersionList TD
{
	FONT-SIZE: xx-small;
	FONT-FAMILY: Verdana
}
</STYLE>
<BODY  bgcolor=Ivory>
<LINK href="../style/wizard style.css" type=text/css rel=stylesheet >
<font size=3 face=verdana><b>Set Deliverable Preference</b><BR></font>

<%
	if trim(Request("RootID")) = "" or not isnumeric(trim(Request("RootID"))) then
		Response.write "Not enough information supplied to complete this transaction."
	else
		dim cn
		dim rs
		dim strType
		dim strTypeID
		dim strCategory
		dim strCategoryID
		dim strVendorID
	    dim strReleaseLink
        dim strSql
        dim intReleaseID

        intReleaseID = 0
		strType = ""
		strTypeID = ""
		strCategory = ""
		strCategoryID = ""
		
		set cn = server.CreateObject("ADODB.connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString")
		cn.Open

		blnLoadFailed = false

		set rs = server.CreateObject("ADODB.recordset")

		rs.Open "spGetRootProperties " & clng(Request("RootID")),cn,adOpenStatic
		if rs.EOF and rs.BOF then
			Response.Write "Deliverable not found."		
		else
			strType = rs("DeliverableType")
			strTypeID = rs("TypeID")
			strCategory = rs("Category")
			strCategoryID = rs("CategoryID")
		end if
		rs.Close
		
		strVendorID = ""
		if trim(request("VersionID")) <> "" then
			if isnumeric(request("VersionID") ) then
				rs.Open "spGetVersionProperties4Web " & request("VersionID"),cn,adOpenStatic
				if not (rs.EOF and rs.BOF) then
					strVendorID = rs("VendorID")
				end if		
				rs.Close
			end if
		end if
		
		if clng(strTypeID) <> 1 then
			Response.Write "You can only restrict hardware deliverables."
		else
		
		dim strCycleProductLinksCons
		dim strCycleProductLinksComm
		strCycleProductLinksCons = ""
		strCycleProductLinksComm = ""
		dim strProductIDLinks
		strProductIDLinks = ""
		
		dim strBusinessProductLinks
		
		
%>		
		<form id=frmUpdate action="RestrictSavePulsar.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
        
		<Table width=100%>
			<TR ID=Tab1><TD>
				<font size=2 face=verdana>Step 1 of 3 - Select Action<BR><BR></font>
				<table bgcolor=cornsilk border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
					<TR>
					<TD valign=top><b>Action:</b>&nbsp;&nbsp;&nbsp;</TD>
					<TD width="100%">
						<INPUT type="radio" id=optRestriction name=optRestriction checked value=0 LANGUAGE=javascript onchange="return optRestriction_onchange()" onclick="return optRestriction_onclick()"> Add Restriction<BR>
						<INPUT type="radio" id=optRestriction name=optRestriction value=1 LANGUAGE=javascript onchange="return optRestriction_onchange()" onclick="return optRestriction_onclick()"> Remove Restriction
					</TD>
					</TR>
				</TABLE><BR>
			</TD></TR>
			
			<TR ID=Tab2 style="display:none"><TD>
				<font size=2 face=verdana>Step 2 of 3 - Select Filters<BR><BR></font>
				<table bgcolor=cornsilk border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
					<TR>
					<TD valign=top><b>Category:</b></TD>
					<TD width="100%"><%=strCategory%></TD>
					</TR>
					<TR>
						<TD valign=top><b>Vendor:</b></TD>
						<TD width="100%">
						<SELECT style="width:200" id=cboVendor name=cboVendor LANGUAGE=javascript onchange="return cboVendor_onchange()">
							<OPTION value=""></OPTION>
							<%
							rs.open "spListVendors4Category " & clng(strCategoryID),cn,adOpenStatic
							do while not rs.EOF
								if trim(strVendorID) = trim(rs("ID")) then
									Response.Write "<Option selected value=""" & rs("ID") & """>" & rs("Vendor") & "</option>"
								else
									Response.Write "<Option value=""" & rs("ID") & """>" & rs("Vendor") & "</option>"
								end if
								rs.MoveNext
							loop
							rs.Close
							%>
							</Select>
						</TD>
					</TR>

					<TR>
						<TD valign=top><b>Products:</b></TD>
						<TD width="100%" valign=top>
						<TABLE><TR><TD valign=top>
						<SELECT style="width:300" id=lstProducts name=lstProducts size=12 multiple LANGUAGE=javascript onchange="return lstProducts_onchange()">
							<%
							rs.open "Select ProductVersionID = v.ID, ID = case isnull(pv_r.ID,0) when 0 then v.ID else pv_r.ID end, ReleaseID = isnull(pvr.ID,0), name = case isnull(pv_r.ID,0) when 0 then v.DOTSName else v.DOTSName + ' (' + pvr.Name + ')' end, c.name as DevCenter " &_
                                    "from productversion v with (NOLOCK) " &_
                                    "inner join devcenter c with (NOLOCK) on c.id = v.devcenter " &_
                                    "left outer join productversion_release pv_r with (NOLOCK) on pv_r.ProductVersionID = v.ID " &_
                                    "left outer join productversionrelease pvr with (NOLOCK) on pv_r.releaseid = pvr.id " &_
                                    "where v.oncommoditymatrix=1 order by c.name, v.dotsname, pvr.ReleaseYear desc, pvr.ReleaseMonth desc;", cn, adOpenStatic
                                
                          
							dim strProdIDList
							strProdIDList = ""
                            dim LastDevCenter
                            LastDevCenter = ""
							do while not rs.EOF
                                if LastDevCenter <> rs("DevCenter") then
                                    if LastDevCenter = "" then
                                        response.write "</optgroup>"
                                    end if
                                    response.write "<optgroup label=""" & rs("DevCenter") & """>"
                                end if
                                LastDevCenter = rs("DevCenter")
								if trim(rs("ProductVersionID")) = trim(request("ProductID")) and rs("ReleaseID") = 0 then
									Response.Write "<Option selected value=""" & rs("ID") & """ ReleaseID=""" & rs("ReleaseID") & """>" & rs("Name") & "</option>"
                                elseif trim(rs("ProductVersionID")) = trim(request("ProductID")) and rs("ReleaseID") > 0 and intReleaseID = 0 then
									Response.Write "<Option selected value=""" & rs("ID") & """ ReleaseID=""" & rs("ReleaseID") & """>" & rs("Name") & "</option>"
								elseif trim(rs("ProductVersionID")) = trim(request("ProductID")) and rs("ReleaseID") > 0 and intReleaseID > 0 and rs("ReleaseID") = intReleaseID then 
                                    Response.Write "<Option selected value=""" & rs("ID") & """ ReleaseID=""" & rs("ReleaseID") & """>" & rs("Name") & "</option>"
                                else
									Response.Write "<Option value=""" & rs("ID") & """ ReleaseID=""" & rs("ReleaseID") & """>" & rs("Name") & "</option>"
								end if
								strProdIDList = strProdIDList & "," & rs("ID")
								rs.MoveNext
							loop
							rs.Close
                            if LastDevCenter = "" then
                                response.write "</optgroup>"
                            end if

							if len(strProdIDList) <> 0 then
								strProdIDList = mid(strProdIDList,2)
							end if
							
							%>
							</Select>
							</TD>
							<TD valign=top>
                                <font size="2" face="Verdana"><b><u>Select Products by Cycle.</u></b><br>

                                    <%
        dim  strCycleProductLinks
		strCycleProductLinks = ""

		
        dim strLastProgramGroup
		strProductIDLinks = ""
        strLastProgramGroup = ""
        Lastprogram=""
		rs.Open "spGetProgramTree",cn,adOpenForwardOnly
		do while not rs.EOF
            if strLastProgramGroup ="" then
                strCycleProductLinks = strCycleProductLinks & "<a href=""javascript:GroupHeaderClicked(" & rs("programGroupID") & ");"">" & rs("ProgramGroup") & "</a><br><div style=""display:none"" id=""ProgramGroupList" & rs("ProgramGroupID") & """><ul style=""margin-top:0px;margin-bottom:0px;"">"
                strLastProgramGroup = rs("ProgramGroup")
            end if
			if LastProgram <> rs("Program") and LastProgram <> "" then
				if len(strProductIDLinks) > 1 then
					strProductIDLinks = mid(strProductIDLinks,2)
				end if
					strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a></li>"
				strProductIDLinks = ""
			end if
            if strLastProgramGroup <> rs("ProgramGroup") then
                strCycleProductLinks = strCycleProductLinks & "</ul></div><a href=""javascript:GroupHeaderClicked(" & rs("programGroupID") & ");"">" & rs("ProgramGroup") & "</a><br><div style=""display:none"" id=""ProgramGroupList" & rs("ProgramGroupID") & """><ul style=""margin-top:0px;margin-bottom:0px;"">"
                strLastProgramGroup = rs("ProgramGroup")
            end if
			strProductIDLinks = strProductIDLinks & "," & rs("ID")
			LastProgram = rs("Program") & ""
			rs.MoveNext
		loop
		if strLastProgramGroup <> "" and LastProgram <> "" then
			if len(strProductIDLinks) > 1 then
				strProductIDLinks = mid(strProductIDLinks,2)
			end if
			strCycleProductLinks = strCycleProductLinks & "<li><a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a></li>"
			strProductIDLinks = ""
            strCycleProductLinks = strCycleProductLinks & "</ul></div>"
        end if
		rs.Close		
		
        response.write      strCycleProductLinks                               
                                    
                                    
                                    
                                    %>

                                    </Font>

                
							
							</TD>
							</TR>
							</TABLE>
							
						</TD>
					</TR>
		
					</Table><BR>
				</TD></TR>

			<TR ID=Tab3 style="display:none"><TD>
				<font size=2 face=verdana>Step 3 of 3 - Notifications and Review<BR><BR></font>
				<table bgcolor=cornsilk border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
					<TR>
					<TD valign=top><b>Notify:</b></TD>
					<TD width="100%"><u>Notification of this change will be sent to:</u><BR>1. All selected deliverable developers and PM's.<BR>2. Qualification ODM PM's for the selected products.<BR>3. kidwell.proceng@hp.com<br>4. twnpdcnbcommoditytechnology@hp.com<BR>5. Submitter</TD>
					</TR>
					<TR>
					<TD valign=top><b>CC:</b></TD>
					<TD width="100%">
					<TABLE width=100% cellpadding=0 cellspacing=0 border=0><TR><TD width=100%><INPUT type="text" style="width:100%" id=txtCC name=txtCC value=""></TD><TD><INPUT type="button" value="Add" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()"></TD></TR></TABLE>
					</TD>
					</TR>
					<TR>
						<TD valign=top><b>Action:</b><INPUT type="hidden" id=txtActionText name=txtActionText></TD>
						<TD width="100%" ID=ReviewAction>Processing.  Please Wait...</TD>
					</TR>
					<TR>
						<TD valign=top><b>Versions:</b></TD>
						<TD width="100%" ID=VersionListDisplay>Processing.  Please Wait...
						</TD>
					</TR>
					<TR>
						<TD valign=top><b>Products:</b><INPUT type="hidden" id=txtProductNames name=txtProductNames></TD>
						<TD width="100%" ID=ReviewProducts>Processing.  Please Wait...</TD>
					</TR>
					</Table><BR>
				</TD></TR>

		</TABLE>
		
		<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=trim(Request("RootID"))%>">
		<INPUT type="hidden" id=txtCategoryID name=txtCategoryID value="<%=strCategoryID%>">
		<INPUT type="hidden" id=txtCategory name=txtCategory value="<%=strCategory%>">
        <input type="hidden" id=txtPulsarProductList name=txtPulsarProductList value="">
        <input type="hidden" id=txtLegacyProductList name=txtLegacyProductList value="">
        <input type="hidden" id=txtReleaseList name=txtReleaseList value="">
		</form>
<%
		end if


		
		set rs=nothing
		cn.Close
		set cn = nothing
	end if
%>
</BODY>
</HTML>