<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
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
	var ActionID;
	var RootID = frmUpdate.txtRootID.value;
	var CatID = frmUpdate.txtCategoryID.value;
	if (frmUpdate.optRestriction[0].checked)
		ActionID = "0"
	else
		ActionID = "1"
		
	jsrsExecute("RestrictVersionsRSget.asp", myCallback1, "VersionList", Array(CatID.toString(), ListValue(frmUpdate.lstProducts), frmUpdate.cboVendor.value, ActionID, RootID ));
}

function myCallback1( returnstring ){
	if (FiltersDirty)
		{
		VersionListDisplay.innerHTML = returnstring;
		}
	FiltersDirty = false;
	var strVendor = frmUpdate.cboVendor.options[frmUpdate.cboVendor.options.selectedIndex].text;
	var strAction;
	var strAction2;
	if (frmUpdate.optRestriction[0].checked)
		{
		strAction = "Restrict "
		strAction2 = "have been restricted"
		}
	else
		{
		strAction = "Unrestrict "
		strAction2 = "are no longer restricted"
		}
	if (frmUpdate.txtCategoryID.value=="2")
		strCategory = "hard drives";
	else if (frmUpdate.txtCategoryID.value=="3")
		strCategory = "memory modules";
	else if (frmUpdate.txtCategoryID.value=="15")
		strCategory = "optical drives";
	else if (frmUpdate.txtCategoryID.value=="71")
		strCategory = "batteries";
	else if (frmUpdate.txtCategoryID.value=="33")
		strCategory = "input devices";
	else if (frmUpdate.txtCategoryID.value=="36")
		strCategory = "AC adapters";
	else if (frmUpdate.txtCategoryID.value=="40")
		strCategory = "floppy drives";
	else if (frmUpdate.txtCategoryID.value=="131")
		strCategory = "panels";
	else
		strCategory = frmUpdate.txtCategory.value;
	
	ReviewAction.innerText = strAction + strVendor + " " + strCategory + " on the selected products."
	ReviewProducts.innerText = replace(ListText(frmUpdate.lstProducts),",",", ");
	frmUpdate.txtProductNames.value = replace(ListText(frmUpdate.lstProducts),",",", ");
	frmUpdate.txtActionText.value = "The following " + strVendor + " " + strCategory + " " + strAction2 + " on the selected products."
	if (VersionListDisplay.innerText == "No Versions Found.")
		window.parent.frames["LowerWindow"].cmdFinish.disabled=true;

}

function replace(sampleStr,replaceChars,replaceWith){

	var replaceRegExp = new RegExp(replaceChars,'g');

	return sampleStr.replace(replaceRegExp,replaceWith);
}


function ListValue ( myList ){
	var i;
	var strList;
	strList = "";
	for (i=0;i<myList.length;i++)
		if (myList.options[i].selected)
			if (strList=="")
				strList = myList.options[i].value;
			else
				strList = strList + "," + myList.options[i].value;
	return strList;
}

function ListText ( myList ){
	var i;
	var strList;
	strList = "";
	for (i=0;i<myList.length;i++)
		if (myList.options[i].selected)
			if (strList=="")
				strList = myList.options[i].text;
			else
				strList = strList + "," + myList.options[i].text;
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
		if (ListText(frmUpdate.lstProducts)=="")
			{
			alert("You must select at least one product.");
			frmUpdate.lstProducts.focus();
			strSuccess = false;
			}
		else if (frmUpdate.cboVendor.selectedIndex < 1)
			{
			alert("You must select a vendor.");
			frmUpdate.lstProducts.focus();
			strSuccess = false;
			}
		else
			PopulateVersions();
		}
	else if (TabID==2) //Saving
		{
		if (VersionListDisplay.innerText == "No Versions Found.")
			{
			alert("Your filters did not find any versions to update.");
			strSuccess =false;
			}	
		else
			{
			blnFound = false;
			if (typeof(frmUpdate.chkVersions.length) == "undefined")
				{
				if (frmUpdate.chkVersions.checked)
					blnFound = true;
				}
			else
				{
				for (i=0;i<frmUpdate.chkVersions.length;i++)
					if (frmUpdate.chkVersions[i].checked)
						blnFound = true;
				}		
			if (!blnFound)
				{
				alert("You must select at least one deliverable version.");
				strSuccess =false;
				}	
			else if (! VerifyEmail(frmUpdate.txtCC.value))
				{
				alert("The email address specified in the CC field does not appear to be valid.");
				frmUpdate.txtCC.focus();
				strSuccess =false;
				}
			else
				frmUpdate.submit();	
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
<font size=3 face=verdana><b>Set Deliverable Preference</b><BR><BR></font>

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
'		rs.Open "spGetProgramTree",cn,adOpenForwardOnly
'		Dim LastProgram
'		Dim LastProgramID
'		dim LastBusiness
'		do while not rs.EOF
'			if LastProgram <> rs("Program") and LastProgram <> "" then
'				if len(strProductIDLinks) > 1 then
'					strProductIDLinks = mid(strProductIDLinks,2)
'				end if
'				if LastBusiness = 2 then
''					strCycleProductLinksCons = strCycleProductLinksCons & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
	'			else
	'				strCycleProductLinksComm = strCycleProductLinksComm & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
	'			end if
	'			strProductIDLinks = ""
	'		end if
	'		strProductIDLinks = strProductIDLinks & "," & rs("ID")
	'		LastProgram = rs("Program") & ""
	'		lastProgramID = rs("ProgramID")
	'		LastBusiness = rs("BusinessID")
	'		rs.MoveNext
	'	loop
		
	'	if len(strProductIDLinks) > 1 then
	'		strProductIDLinks = mid(strProductIDLinks,2)
	'	end if
	'	if LastBusiness = 2 then
	'		strCycleProductLinksCons = strCycleProductLinksCons & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
	'	else
	'		strCycleProductLinksComm = strCycleProductLinksComm & "<a href=""javascript: SelectProducts('" & strProductIDLinks & "');"">" & LastProgram & "</a><BR>"
	'	end if	
	'	rs.Close			
	'	
		
		dim strBusinessProductLinks
		
		
%>		
		<form id=frmUpdate action=RestrictSave.asp method=post>
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
						<SELECT style="width:200" id=lstProducts name=lstProducts size=12 multiple LANGUAGE=javascript onchange="return lstProducts_onchange()">
							<%
							rs.open "Select v.ID, v.DOTSName as name, c.name as DevCenter from productversion v with (NOLOCK), devcenter c with (NOLOCK) where c.id = v.devcenter and v.oncommoditymatrix=1 order by c.name, v.dotsname;",cn,adOpenStatic
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
								if trim(rs("ID")) = trim(request("ProductID")) then
									Response.Write "<Option selected value=""" & rs("ID") & """>" & rs("Name") & "</option>"
								else
									Response.Write "<Option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
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
                                   <!-- <TABLE><TR><TD><u>Consumer</u>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD><u>Commercial</u></TD></TR>
                                    <TR>-->
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

										<!--<TD valign=top><%=strCycleProductLinksCons %></TD>
										<TD valign=top><%=strCycleProductLinksComm%></TD>
                                    </TR>
                                    </TABLE>-->
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
