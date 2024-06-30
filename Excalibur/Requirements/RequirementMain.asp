<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->
<DOCTYPE html>
<HTML>
<HEAD>
<title>Requirement</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../includes/client/json2.js"></script>
<script type="text/javascript" src="../includes/client/json_parse.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var CurrentState;
var States = new Array(2);
var FormLoading = true;
var DisplayedID;

function ProcessState() {
	var steptext;
	
	switch (CurrentState)
	{
		case "General":
			steptext = "";

			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Specification"; //"Specification"; 
			tabGeneral.style.display="";
			tabDeliverables.style.display="none";

			lblInstructions.innerText = "Enter the Specification for this Requirement.";
			//frmRequirement.txtSpecification.focus();
		
			window.scrollTo(0,0);		
		break;

		case "Deliverables":
			steptext = "";
	
			lblTitle.innerText = frmRequirement.txtRequirement.value + " Requirement - Deliverable List";
			tabGeneral.style.display="none";
			tabDeliverables.style.display="";

			lblInstructions.innerText = "Select the deliverable that will fulfill this requirement and choose the default distributions and image support";
		
			window.scrollTo(0,0);		
		break;

	}
}
function window_onload() {
	var i;
	var strID;
	var strName;

	frames.myEditor.document.body.contentEditable = "True";
	frames.myEditor.document.body.innerHTML = "<font face=verdana size=1>" + frmRequirement.txtSpecification.value + "</font>";
	DisplayedID = frmRequirement.txtDisplayedID.value;
    if (txtRequiresDeliverables.value=="False")
	    CurrentState =  "General"; 
    else
	    CurrentState =  "Deliverables"; //"General";
	ProcessState();
	FormLoading = false;
	self.focus();

    //Instantiate modalDialog load
	modalDialog.load();
}

function SelectTab(strStep) {
	var i;

	//Reset all tabs
	document.all("CellGeneralb").style.display="none";
	document.all("CellGeneral").style.display="";
	document.all("CellDeliverablesb").style.display="none";
	document.all("CellDeliverables").style.display="";

	//Highight the selected tab
	document.all("Cell"+strStep).style.display="none";
	document.all("Cell"+strStep+"b").style.display="";

	
	CurrentState = strStep;
	ProcessState();

}

function cboDCR_onchange() {
	var strShowValues = "";
	var strShowEditBoxes = "";
	
	DeleteImage.DelDCRID.value = frmRequirement.cboDCR.value;
	
	if (frmRequirement.cboDCR.value != "" && frmRequirement.txtDisplayedID.value != "")
		txtDeleteOK.value = "True";
	else
		txtDeleteOK.value = "False";

	if (frmRequirement.txtDisplayedID.value == "")
		return;
	
	if (frmRequirement.cboDCR.selectedIndex == 0)
		{
			strShowValues = "";
			strShowEditBoxes = "none";
		}
	else
		{
			strShowValues = "none";
			strShowEditBoxes = "";
		}
	
	frmRequirement.txtSpecification.style.display = strShowEditBoxes;
	lblSpecification.style.display = strShowValues;

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
		
		for (i=0;i<event.srcElement.length;i++)
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

function cboCategories_onclick() {
	var i;
	if (frmRequirement.cboCategories.selectedIndex == 0)
		for (i=2;i<tabDeliverables.rows.length;i++)
			tabDeliverables.rows(i).style.display="";
	else
	{
	for (i=2;i<tabDeliverables.rows.length;i++)
		if (tabDeliverables.rows(i).className == frmRequirement.cboCategories.options[frmRequirement.cboCategories.selectedIndex].value)
			tabDeliverables.rows(i).style.display="";
		else
			tabDeliverables.rows(i).style.display="none";
	}	
		
}

function UpdateDelList(ID,ProdID, prId){
	var strID = new Array;
	
	strID = window.showModalDialog("/Pulsar/product/AddRemoveRequirementDeliverables/?prid=" + prId, "", "dialogWidth:1024px;dialogHeight:950px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	window.location.href = window.location;
	/*var url = '/Pulsar/product/AddRemoveRequirementDeliverables/?prid=' + prId;
	modalDialog.open({ dialogTitle: 'Add/Remove Deliverables', dialogURL: '' + url + '', dialogHeight: 700, dialogWidth: 900, dialogResizable: true, dialogDraggable: true });*/
    
}


function ModDistribution(DelID,ProdID){
	var strID;
	var ResultArray;
	
    /*strID = window.showModalDialog("ChangeDistribution.asp?ProductID=" + ProdID + "&RootID=" + DelID ,"","dialogWidth:720px;dialogHeight:500px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");*/
	url = 'ChangeDistribution.asp?ProductID=' + ProdID + '&RootID=' + DelID;
	modalDialog.open({ dialogTitle: 'Modify Distribution', dialogURL: '' + url + '', dialogHeight: 550, dialogWidth: 650, dialogResizable: true, dialogDraggable: true });

    //save Version ID for results function: ---
	globalVariable.save(DelID, 'deliverable_id');
}

function ModDistResult(strID) {
    var ResultArray;
    var iDeliverableID = globalVariable.get('deliverable_id');
    if (typeof (strID) != "undefined") {
        ResultArray = strID.split("|")
        if (ResultArray[0] == "") {
            document.all("Dist" + iDeliverableID).innerText = "Not Specified";
        } else {
            document.all("Dist" + iDeliverableID).innerText = ResultArray[0];
        }
        if (ResultArray[1] == "") {
            document.all("Image" + iDeliverableID).innerText = "ALL";
        } else {
            document.all("Image" + iDeliverableID).innerText = ResultArray[1];
        }
    }
}

function ModImages(DelID, ProdID, Fusion, Pulsar){
    var url;
	
    if (Pulsar == 1) {
        url = '../Target/Pulsar/ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID;
        /*strID = window.showModalDialog("../Target/Pulsar/ChangeImages.asp?ProductID=" + ProdID + "&RootID=" + DelID, "", "dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
    } else if (Fusion == 1) {
        url = '../Target/Fusion/ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID;
        /*strID = window.showModalDialog("../Target/Fusion/ChangeImages.asp?ProductID=" + ProdID + "&RootID=" + DelID, "", "dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
    } else {
        url = '../Target/ChangeImages.asp?ProductID=' + ProdID + '&RootID=' + DelID;
        /*strID = window.showModalDialog("../Target/ChangeImages.asp?ProductID=" + ProdID + "&RootID=" + DelID, "", "dialogWidth:800px;dialogHeight:600px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");*/
    }

    modalDialog.open({ dialogTitle: 'Change Image', dialogURL: '' + url + '', dialogHeight: 650, dialogWidth: 850, dialogResizable: true, dialogDraggable: true });

    //save Version ID for results function: ---
    globalVariable.save(DelID, 'deliverable_id');
}

function ModImagesResult(strID) {
    if (typeof (strID) != "undefined") {
        if (strID == "") {
            document.all("Image" + globalVariable.get('deliverable_id')).innerText = "ALL";
        } else {
            document.all("Image" + globalVariable.get('deliverable_id')).innerText = strID;
        }
    }
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

.DelTable TBODY TD{
	BORDER-TOP: gray thin solid;
}


</STYLE>
<BODY bgcolor="ivory" LANGUAGE=javascript onload="return window_onload()">
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<%

	dim cn
	dim rs
	dim strBrand
	dim strBrandID
	dim strBrandList
	dim strOS
	dim strOSID
	dim strOSList
	dim strSW
	dim strSWID
	dim strSWList
	dim strType
	dim blnFound 
	dim strSKU
	dim strRegionMatrix
	dim strPriorityList
	dim i
	dim strLastGeo
	dim strAllRoots
	dim strImageIDList
	dim strImageNameList 
	dim strImageTag
	dim OSID
	dim SWID
	dim BrandID
	dim CurrentUser
	dim CurrentUserID
	dim CurrentUserSysAdmin
	dim CurrentWorkgroupID
	dim strPMID
	dim strSEPMID
	dim strShowValues
	dim strShowEditBoxes
	dim strShowDCR
	dim blnPOR
	dim strDCRs
	dim blnEditOK
	dim mstrCopyTag
	dim strRegionList
	dim PriorityArray
	dim blnDeleteOK
	dim strDisplayDelete
	dim strRequirement
	dim strCategories
	dim strDistributions
	dim cm
	dim p
    dim isFusion
    dim isPulsar

    isPulsar = 0

	CurrentUserSysAdmin = false
	
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
		CurrentUserSysAdmin = rs("SystemAdmin")
	end if
	rs.Close
	

	blnPOR = false
    isFusion = 0

	rs.Open "spGetProductVersion " & clng(request("ProdID")),cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		strSEPMID = rs("SEPMID") & ""
		strPMID = rs("PMID") & ""
        if rs("Fusion") then
            isFusion = 1
        else
            isFusion = 0
        end if
        if (rs("FusionRequirements")) then
            isPulsar = 1
        else
            isPulsar = 0  
        end if
	end if
	rs.Close

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	if (request("ID") <> "" and (CurrentUserSysAdmin or CurrentWorkgroupID = 15 or strSEPMID = CurrentUSerID or strPMID = CurrentUSerID)) or request("ID") = "" then
		blnEditOK = true
	else
		blnEditOK = false
	end if

	strShowDCR = "none"
	if (not (blnEditOK) or (blnPOR)) and request("ID") <> "" then
		strShowValues = ""
		strShowEditBoxes = "none"
		blnDeleteOK = false
	else
		strShowValues = "none"
		strShowEditBoxes = ""
		if request("CopyID") = "" and request("ID") <> "" then
			blnDeleteOK = true
		else
			blnDeleteOK = false	
		end if
	end if
	
	if blnDeleteOK then
		strDisplayDelete = ""
	else
		strDisplayDelete = "none"
	end if

	if blnPOR and blnEditOK then
		strShowDCR = "" 
		rs.Open "spListApprovedDCRs " & clng(request("ProdID")),cn,adOpenForwardOnly
		strDCRs = "<Option selected></option>"
		do while not rs.EOF
			strDCRs = strDCRs & "<option value=""" & rs("ID") & """>" & rs("ID") & " - " & rs("Summary") & "</option>"
			rs.MoveNext
		loop
		rs.Close
	else
		strShowDCR = "none" 
	end if

		
	'Load Current Values
	dim strSpecification
	dim PRID
	dim strImageSummary
	dim strSelectedRoots
    dim blnRequiresDeliverables
	strSpecification = ""
	strSelectedRoots = ""
	
	if Request("ID") <> ""then
		rs.Open "spGetRequirementDefinition " & clng(Request("ID")) & "," & clng(request("ProdID")) & ",0",cn,adOpenForwardOnly
		strSpecification = rs("Specification") & ""
		strRequirement = rs("Requirement") & ""
        blnRequiresDeliverables = rs("RequiresDeliverables")
		PRID = rs("ID")
		rs.Close
	end if
	
	strCategories = ""
	rs.Open "spGetDeliverableCategories",cn,adOpenForwardOnly
	do while not rs.EOF
		strCategories = strCategories & "<option value=" & rs("ID") & ">" & rs("Type") & " - " & rs("Name") & "</option>"
		rs.MoveNext
	loop
	rs.Close
	
	'Build Deliverable List
	strSelectedRoots = ""
	strAllRoots = ""
	rs.Open "spListDeliverablesByRequirement " & clng(Request("ID")) & "," & clng(request("ProdID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strAllRoots = strAllRoots & "<TR  style=""BORDER-TOP: gray thin solid""  valign=top bgcolor=cornsilk><TD colspan=3><font size=1 face=verdana><b>No Deliverable Selected</b></font></td></tr>"
	else
		set rs2 = server.CreateObject("ADODB.recordset")
	
		do while not rs.EOF
				strDistributions = ""
				rs2.Open "spGetDistributionRoot " & clng(request("ProdID")) & "," & clng(rs("DeliverableID")),cn,adOpenForwardOnly
				if rs2("typeid")= 1 then
					strImageSummary = "-"
					strDistributions = "-"
				else
					strImageSummary = rs2("ImageSummary") & ""
					if trim(strImageSummary) = "" then
						strImageSummary = "ALL"
					end if


					if rs2("Preinstall") then
						strDistributions = strDistributions & ",Preinstall"
					end if
				
					if rs2("Preload") then
						strDistributions = strDistributions & ",Preload"
					end if
	
					if rs2("Web") then
						strDistributions = strDistributions & ",Web"
					end if
	
					if rs2("DIB") then
						strDistributions = strDistributions & ",DIB"
					end if
	
					if rs2("SelectiveRestore") then
						strDistributions = strDistributions & ",Selective Restore"
					end if
	
					if rs2("ARCD") then
						strDistributions = strDistributions & ",DRCD"
					end if
					
					if rs2("DRDVD") then
						strDistributions = strDistributions & ",DRDVD"
					end if

					if rs2("RACD_Americas") then
						strDistributions = strDistributions & ",RACD-Americas"
					end if

					if rs2("RACD_EMEA") then
						strDistributions = strDistributions & ",RACD-EMEA"
					end if

					if rs2("RACD_APD") then
						strDistributions = strDistributions & ",RACD-APD"
					end if

					if rs2("OSCD") then
						strDistributions = strDistributions & ",OSCD"
					end if
					
					if rs2("DocCD") then
						strDistributions = strDistributions & ",Doc CD"
					end if
				
					if trim(rs2("Patch")&"") <> "0" then
						strDistributions = strDistributions & ",Patch"
					end if
					
					if strDistributions = "" then
									strDistributions = "Not Specified"
								else
						strDistributions = mid(strDistributions ,2)
					end if
				end if
				
				
				rs2.Close
				
				
				
				
				strSelectedRoots = strSelectedRoots & rs("name") & "<BR>"
		
				strAllRoots = strAllRoots & "<TR ID=""Row" & rs("DeliverableID") & """ style=""BORDER-TOP: gray thin solid""  valign=top bgcolor=cornsilk>"
				strAllRoots = strAllRoots & "<TD><font size=1 face=verdana>" & rs("Name") & "&nbsp;</font></TD>"
				if strDistributions = "-" then
					strAllRoots = strAllRoots & "<TD ><font size=1 face=verdana><b>" & strDistributions & "</b>&nbsp;</font></TD>"
				else
					strAllRoots = strAllRoots & "<TD ><font size=1 face=verdana><a hidefocus ID=""Dist" & rs("DeliverableID") & """ href=""javascript: ModDistribution(" & rs("DeliverableID") & "," & request("ProdID") & ");"">" & strDistributions & "</a>&nbsp;</font></TD>"
				end if
				if strImageSummary = "-" then
					strAllRoots = strAllRoots & "<TD ><font size=1 face=verdana><b>" & strImageSummary & "</b></font></TD>"
				else
					strAllRoots = strAllRoots & "<TD ><font size=1 face=verdana><a hidefocus ID=""Image" & rs("DeliverableID")  & """ href=""javascript: ModImages(" & rs("DeliverableID") & "," & request("ProdID") & "," & isFusion & "," & isPulsar & ");"">" & strImageSummary & "</a></font></TD>"
				end if
				strAllRoots = strAllRoots & "</tr>"
			rs.MoveNext
		loop
		
		set rs2 = nothing
	end if
	rs.Close
	
	
	if strSelectedRoots <> "" then
		strSelectedRoots = mid(strSelectedRoots,2)
	end if
	
	cn.Close
	set cn = nothing
	set rs = nothing
	
	
	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0
	
%>


<%if request("ID") <> "" and blnRequiresDeliverables then%>

<table Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0">
	<tr bgcolor="<%=strTitleColor%>">
		<td id="CellGeneral" style="Display:" width="10"><font size="2" color="black"><b>&nbsp;<a href="javascript:SelectTab('General')">Specification</a>&nbsp;</b></font></td>
		<td id="CellGeneralb" style="Display:none" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Specification&nbsp;</b></font></td>
		<td id="CellDeliverables" style="Display:none" width="10"><font size="2" color="white"><b>&nbsp;<a href="javascript:SelectTab('Deliverables')">Deliverables</a>&nbsp;</b></font></td>
		<td id="CellDeliverablesb" style="Display:" width="10" bgcolor="wheat"><font size="2" color="black"><b>&nbsp;Deliverables&nbsp;</b></font></td>
	</tr>
</table>
<hr color="Tan">
<%end if%>

<font face=verdana size=><b>
<label ID="lblTitle"></label></b></font>

<form id="frmRequirement" method="post" action="RequirementSave.asp">


<font size="2">
<label ID="lblInstructions"></label>
</font><BR><BR>

<table ID="tabGeneral" style="DISPLAY: none" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr style="Display:<%=strShowDCR%>">
		<td width=120 nowrap><b>Approved&nbsp;DCR:&nbsp;</b><font color="#ff0000" size="1">*</font></td>
		<td>
		<SELECT style="WIDTH:100%" id=cboDCR name=cboDCR LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()" onchange="return cboDCR_onchange()"><%=strDCRs%></SELECT>
		</td>
	</tr>
	<tr style="Display:none">
		<td valign=top width=120 nowrap><b>Specification:</b></td>
		<td>
			<LABEL ID=lblSpecification Style="Display:<%=strShowValues%>"><%=strSpecification%></LABEL>
			<TEXTAREA style="WIDTH:100%;HEIGHT:300px;Display:<%=strShowEditBoxes%>" id=txtSpecification name=txtSpecification><%=strSpecification%></TEXTAREA>
			<TEXTAREA style="WIDTH:200px;Display:none" id=tagSpecification name=tagSpecification><%=strSpecification%></TEXTAREA>

		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Requirement:</b></td>
		<td>
			<LABEL ID=lblRequirement><%=strRequirement%></LABEL>
		</td>
	</tr>
	<tr>
		<td valign=top width=120 nowrap><b>Specification:</b></td>
		<td>
			<font size=2 face=verdana>See PDD for official requirement specification.</font>
		</td>
	</tr>
	<TR style=display:>
		<td valign=top width=120 nowrap><b>Specification:</b></td>
		<TD>
			<a href="javascript:formatSelection('Bold');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Bold src="../images/BLD.BMP"></a><a href="javascript:formatSelection('Italic');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Italic src="../images/ITL.BMP"></a><a href="javascript:formatSelection('Underline');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Underline src="../images/UNDRLN.BMP"></a>
			
			<a href="javascript:InsertFormat('insertunorderedlist');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt="Unordered List" src="../images/uordList.BMP"></a><a href="javascript:InsertFormat('insertorderedlist');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt="Ordered List" src="../images/ordList.BMP"></a>
			
			<a href="javascript:InsertFormat('outdent');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Outdent src="../images/outdent.BMP"></a><a href="javascript:InsertFormat('indent');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Indent src="../images/indent.BMP"></a>
			
			<a href="javascript:InsertFormat('createLink');"><IMG style="BORDER-RIGHT: gainsboro thin outset; BORDER-TOP: white thin outset; BORDER-LEFT: white thin outset; BORDER-BOTTOM: gainsboro  thin outset" alt=Hyperlink src="../images/Hlink.BMP"></a>
			<BR>
				<IFRAME ID=myEditor style="WIDTH: 100%; HEIGHT: 350px" noResize></IFRAME>
				<SCRIPT>
					function formatSelection(strFormat) {
						// Get a text range for the selection
						frames.myEditor.focus();
						var tr = frames.myEditor.document.body.createTextRange();
    
						// Execute command
						tr.execCommand(strFormat);
    
						// Reselect and give the focus back to the editor
						tr.select()
						frames.myEditor.focus();
					}
					
					function SetFont(strFace,strSize){
						frames.myEditor.focus();
						var tr = frames.myEditor.document.body.createTextRange();
    
   
						// Reselect and give the focus back to the editor
						alert(frames.myEditor.document.body.innerHTML);
					//	frames.myEditor.focus();
					}


					function InsertFormat(strFormat) {
						// Get a text range for the selection
						frames.myEditor.focus();
						var tr = frames.myEditor.document.body.createTextRange();
						tr.execCommand(strFormat);
						
						if (strFormat == "createLink")
							{
								var oAnchors = frames.myEditor.document.all.tags("A");
								if (oAnchors != null) 
									{ 
										for(var i = oAnchors.length - 1; i >= 0; i--)
										  if (oAnchors[i].target != "_blank")
											oAnchors[i].target = "_blank"
									} 
							}

						// Reselect and give the focus back to the editor
						frames.myEditor.focus();
					}

				</SCRIPT>
	</TD></TR>
</table>

<div ID=DelList><table class="DelTable" ID="tabDeliverables" style="DISPLAY: none" WIDTH="100%"  BORDER=0 CELLSPACING="0" CELLPADDING="1">
<THEAD><TR><TD colspan=4 align=right><font size=1 face=verdana><a href="javascript: UpdateDelList(<%=request("ID")%>,<%=request("ProdID")%>,<%=PRID%>);">Add/Remove Deliverables</a></font><BR><BR></TD></tr>
<TR bgcolor=Wheat><TD><font size=1 face=verdana><b>Deliverable</b></font></TD><TD><font size=1 face=verdana><b>Distribution</b></font></TD><TD><font size=1 face=verdana><b>Images</b></font></TD></tr></THEAD>
	<TBODY>
	<%=strAllRoots%>
	</TBODY>
</table>
</div>
<input style="Display:none" type="text" id="ID" name="ID" value="<%=request("ID")%>">


<%
	if len(strImageTag) > 0 then
		strImageTag = mid(strImageTag,2)
	end if

	if len(strCopyTag) > 0 then
		strCopyTag = mid(strCopyTag,2)
	end if

	if len(strImageIDList) > 0 then
		strImageIDList = mid(strImageIDList,2)
	end if

	if len(strImageNameList) > 0 then
		strImageNameList = mid(strImageNameList,2)
	end if
%>

<%if request("CopyID") <> "" then%>
	<INPUT type="hidden" style="WIDTH:100%" id=txtTag name=txtTag value="<%=strCopyTag%>">
<%else%>
	<INPUT type="hidden" style="WIDTH:100%" id=txtTag name=txtTag value="<%=strImageTag%>">
<%end if%>
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtProdID name=txtProdID value="<%=request("ProdID")%>">
<INPUT type="hidden" id=txtDCRRequired name=txtDCRRequired value="<%=strShowDCR%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtRequirement name=txtRequirement value="<%=strRequirement%>">
<INPUT type="hidden" id=txtPRID name=txtPRID value=<%=PRID%>>
</form>

<form ID=DeleteReq method=post action=ReqDelete.asp>
	<INPUT type="hidden" id=Auth name=Auth value="">
	<INPUT type="hidden" id=DelReqeID name=DelReqID value="<%=request("ID")%>">
	<INPUT type="hidden" id=txtDelUserID name=txtDelUserID value="<%=CurrentUserID%>">
	<INPUT type="hidden" id=DelDCRID name=DelDCRID value="">
</form>
<INPUT type="hidden" id=txtDeleteOK name=txtDeleteOK value="<%=blnDeleteOK%>">
<INPUT type="hidden" id=txtRequiresDeliverables value="<%=trim(blnRequiresDeliverables)%>">

<TEXTAREA style="Display:none" rows=2 cols=20 id=txtDelList name=txtDelList><%=strSelectedRoots%></TEXTAREA>
</BODY>
</HTML>
