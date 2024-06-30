<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<DOCTYPE html>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<!-- #include file="../includes/bundleConfig.inc" -->
<script type="text/javascript" src="../includes/client/json2.js"></script>
<script type="text/javascript" src="../includes/client/json_parse.js"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {
	frmChange.submit();
}

function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else if (parent.parent.window.parent.loadDatatodiv != undefined) {
        parent.window.parent.closeExternalPopup();
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}

function ChangeThis_onclick() {
	frmChange.optThis.checked=true;
	frmChange.optFuture.checked=false;
}

function ChangeDefault_onclick() {
	frmChange.optThis.checked=false
	frmChange.optFuture.checked=true
}

function ChangeThis_onmouseover() {
	window.event.srcElement.style.cursor = "hand";
}

function ChangeDefault_onmouseover() {
	window.event.srcElement.style.cursor = "hand";

}

function ChangePI_onmouseover() {
    if (!frmChange.chkPreinstall.disabled)
	    window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}

function ChangePL_onmouseover() {
    if (!frmChange.chkPreLoad.disabled)
    	window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}

function ChangeDIB_onmouseover() {
    if (!frmChange.chkDropInBox.disabled)
    	window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}

function ChangeWeb_onmouseover() {
    if (!frmChange.chkWeb.disabled)
    	window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}
function ChangePatch_onmouseover() {
    if (!frmChange.chkPatch.disabled)
        window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}

function ChangeRO_onmouseover() {
    if (!frmChange.chkRCDOnly.disabled)
        window.event.srcElement.style.cursor = "hand";
    else
        window.event.srcElement.style.cursor = "default";
}

function ChangePI_onclick() {
	if (frmChange.chkPreinstall.checked)
		frmChange.chkPreinstall.checked = false;
	else if (!frmChange.chkPreinstall.disabled)
		frmChange.chkPreinstall.checked = true;
}

function ChangeDIB_onclick() {
	if (frmChange.chkDropInBox.checked)
		frmChange.chkDropInBox.checked = false;
	else if (!frmChange.chkDropInBox.disabled)
		frmChange.chkDropInBox.checked = true;
}

function ChangePL_onclick() {
	if (frmChange.chkPreLoad.checked)
		frmChange.chkPreLoad.checked = false;
	else if (!frmChange.chkPreLoad.disabled)
		frmChange.chkPreLoad.checked = true;
}

function ChangeWeb_onclick() {
	if (frmChange.chkWeb.checked)
		frmChange.chkWeb.checked = false;
	else if (!frmChange.chkWeb.disabled)
		frmChange.chkWeb.checked = true;
}

function ChangePatch_onclick() {
    if (frmChange.chkPatch.checked)
        frmChange.chkPatch.checked = false;
    else if (!frmChange.chkPatch.disabled)
        frmChange.chkPatch.checked = true;
}

function ChangeDRCD_onclick() {
	if (frmChange.chkDRCD.checked)
		frmChange.chkDRCD.checked = false;
	else if (!frmChange.chkDRCD.disabled)
		frmChange.chkDRCD.checked = true;
}

function ChangeDRDVD_onclick() {
	if (frmChange.chkDRDVD.checked)
		frmChange.chkDRDVD.checked = false;
	else if (!frmChange.chkDRDVD.disabled)
		frmChange.chkDRDVD.checked = true;
}

function ChangeOSCD_onclick() {
	if (frmChange.chkOSCD.checked)
		frmChange.chkOSCD.checked = false;
	else if (!frmChange.chkOSCD.disabled)
		frmChange.chkOSCD.checked = true;
}

function ChangeDocCD_onclick() {
	if (frmChange.chkDocCD.checked)
		frmChange.chkDocCD.checked = false;
	else if (!frmChange.chkDocCD.disabled)
		frmChange.chkDocCD.checked = true;
}

function ChangeRACDAmericas_onclick() {
	if (frmChange.chkRACDAmericas.checked)
		frmChange.chkRACDAmericas.checked = false;
	else if (!frmChange.chkRACDAmericas.disabled)
		frmChange.chkRACDAmericas.checked = true;
}

function ChangeRACDEMEA_onclick() {
	if (frmChange.chkRACDEMEA.checked)
		frmChange.chkRACDEMEA.checked = false;
	else if (!frmChange.chkRACDEMEA.disabled)
		frmChange.chkRACDEMEA.checked = true;
}

function ChangeRACDAPD_onclick() {
	if (frmChange.chkRACDAPD.checked)
		frmChange.chkRACDAPD.checked = false;
	else if (!frmChange.chkRACDAPD.disabled)
		frmChange.chkRACDAPD.checked = true;
}


function ChangeSR_onclick() {
	if (frmChange.chkSR.checked)
		frmChange.chkSR.checked = false;
	else if (!frmChange.chkSR.disabled)
		frmChange.chkSR.checked = true;
}

function chkDropInBox_onclick(){
	if (frmChange.chkDropInBox.checked)
		HWReqCell.style.display="";
	else if (!frmChange.chkDropInBox.disabled)
		HWReqCell.style.display="none";
}

function ChangePatch_onclick() {
    if (frmChange.chkPatch.checked)
        frmChange.chkPatch.checked = false;
    else if (!frmChange.chkPatch.disabled)
        frmChange.chkPatch.checked = true;

    PatchClicked();
}

function ChangeRO_onclick() {
    if (frmChange.chkRCDOnly.checked)
        frmChange.chkRCDOnly.checked = false;
    else if (!frmChange.chkRCDOnly.disabled)
        frmChange.chkRCDOnly.checked = true;
}

function PatchClicked() {
    if (frmChange.chkPatch.checked) {
        frmChange.chkPreinstall.checked = false;
        frmChange.chkPreinstall.disabled = true;
        frmChange.chkPreLoad.checked = false;
        frmChange.chkPreLoad.disabled = true;
        frmChange.chkRACDAPD.checked = false;
        frmChange.chkRACDAPD.disabled = true;
        frmChange.chkOSCD.checked = false;
        frmChange.chkOSCD.disabled = true;
        frmChange.chkRACDEMEA.checked = false;
        frmChange.chkRACDEMEA.disabled = true;
        frmChange.chkRACDAmericas.checked = false;
        frmChange.chkRACDAmericas.disabled = true;
        frmChange.chkDocCD.checked = false;
        frmChange.chkDocCD.disabled = true;
        frmChange.chkSR.checked = false;
        frmChange.chkSR.disabled = true;
        frmChange.chkDRCD.checked = false;
        frmChange.chkDRCD.disabled = true;
        frmChange.chkDRDVD.checked = false;
        frmChange.chkDRDVD.disabled = true;
        frmChange.chkWeb.checked = false;
        frmChange.chkWeb.disabled = true;
        frmChange.chkDropInBox.checked = false;
        frmChange.chkDropInBox.disabled = true;
        PickPreloadBrand.style.display = "none";
        PickPreinstallBrand.style.display = "none";
        HWReqCell.style.display = "none";
    }
    else {
        frmChange.chkPreinstall.disabled = false;
        frmChange.chkPreLoad.disabled = false;
        frmChange.chkRACDAPD.disabled = false;
        frmChange.chkRACDEMEA.disabled = false;
        frmChange.chkRACDAmericas.disabled = false;
        frmChange.chkDocCD.disabled = false;
        frmChange.chkSR.disabled = false;
        frmChange.chkDRCD.disabled = false;
        frmChange.chkDRDVD.disabled = false;
        frmChange.chkOSCD.disabled = false;
        frmChange.chkWeb.disabled = false;
        frmChange.chkDropInBox.disabled = false;
        PickPreloadBrand.style.display = "";
        PickPreinstallBrand.style.display = "";
    }

    //initialize modal dialog: --
    modalDialog.load();
}



function AddHWReq(){
	var NewValue="";
	var NewRow;
	var NewCell;
	var i;
	var blnFound;
	var NewReq= window.prompt("Enter the new DIB Optical Drive Requirement:","");
	if (!(NewReq == null || NewReq==""))
		{
		NewValue=NewReq.toUpperCase();
		blnFound = false
		for(i=0;i<frmChange.chkDIBHW.length;i++)
			if (NewValue==frmChange.chkDIBHW[i].value.toUpperCase())
				{
				frmChange.chkDIBHW[i].checked=true;
				blnFound = true
				}
		//Add to list
		if (!blnFound)
			{
		NewRow = ReqTable.insertRow(ReqTable.rows.length-1);
		//NewRow.name = "Row" + (ReqTable.rows.length-2);
		//NewRow.id = "Row" + (ApproverTable.rows.length-2);
		NewCell = NewRow.insertCell();
		NewCell.innerHTML = "<INPUT style=\"WIDTH:16;HEIGHT:16\" checked type=\"checkbox\" id=\"chkDIBHW \" name=\"chkDIBHW\" value=\"" + NewReq + "\">&nbsp;" + NewReq;
			
			
			}
		}
	}
	
	
	function PickBrand(strType){
	    var strID;
	    var url;
	
	    if (strType == 1) {
	        url = 'DistributionBrand.asp?AllBrands=' + frmChange.txtAllBrands.value + '&SelectedBrands=' + frmChange.txtPreinstallBrands.value;
	        //strID = window.showModalDialog("DistributionBrand.asp?AllBrands=" + frmChange.txtAllBrands.value + "&SelectedBrands=" + frmChange.txtPreinstallBrands.value, "", "dialogWidth:250px;dialogHeight:350px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	    } else {
	        url - 'DistributionBrand.asp?AllBrands=' + frmChange.txtAllBrands.value + '&SelectedBrands=' + frmChange.txtPreloadBrands.value;
	        //strID = window.showModalDialog("DistributionBrand.asp?AllBrands=" + frmChange.txtAllBrands.value + "&SelectedBrands=" + frmChange.txtPreloadBrands.value, "", "dialogWidth:250px;dialogHeight:350px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	    }
        
	    modalDialog.open({ dialogTitle: 'Distribution Brand', dialogURL: '' + url + '', dialogHeight: 400, dialogWidth: 300, dialogResizable: true, dialogDraggable: true });
	}

	function PickBrandResult(strID) {
	    if (typeof (strID) != "undefined") {
	        if (frmChange.txtAllBrands.value == strID)
	            strID = ""

	        if (strType == 1) {
	            frmChange.txtPreinstallBrands.value = strID;
	            if (strID == "")
	                lblPreinstall.innerText = "All Brands";
	            else
	                lblPreinstall.innerText = strID;
	        }
	        else {
	            frmChange.txtPreloadBrands.value = strID;
	            if (strID == "")
	                lblPreload.innerText = "All Brands";
	            else
	                lblPreload.innerText = strID;
	        }
	    }
    }

//-->
</SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<BODY  bgcolor=Ivory onload="javascript:PatchClicked();" style="padding:1px;"> 

<%
'onload="PatchClicked();"
	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strPreinstallActual
	dim strPreLoadActual
	dim strDropInBoxActual
	dim strWebActual
	dim blnPreinstallDev
	dim blnPreinstallROM
	dim blnCDImage
	dim blnISOImage
	dim blnFloppyDisk
	dim blnScriptpaq
	dim blnSoftpaq
	dim blnRompaq
	dim strSR
	dim strDRCD
	dim strDRDVD
	dim strRACDAmericas
	dim strRACDEMEA
	dim strRACDAPD
	dim strDOCCD
	dim strOSCD
	dim strHWReqs
	dim strAllHWReqs
	dim ShowDIBHW
	dim blnAR
	dim blnAdmin
	dim strPreinstallBrand
	dim strPreloadBrand
	dim strProdBrands
	dim BrandCount
    dim strPatch
    dim strCategoryID
    dim strShowPatch
    dim strRCDOnly
    dim blnRCDOnly
	
		blnAdmin = false

	strPreinstallActual = ""
	strPreLoadActual = ""
	strDropInBoxActual = ""
	strWebActual = ""
	strSR = ""
	strDRCD  = ""
	strDRDVD = ""
	strRACDAmericas = ""
	strRACDEMEA = ""
	strRACDAPD = ""
	strDOCCD = ""
	strOSCD = ""
	strHWReqs = ""
	strAllHWReqs = ""
    strPatch = ""
	blnAR = false
	strPreinstallBrand = ""
	strPreloadBrand = ""
	strProdBrands = ""
	strCategoryID=""
    strRCDOnly=""

	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	ShowDIBHW = "none"
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")


	if not blnLoadFailed then
		rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProdName = ""
			blnLoadFailed = true
		else
			strprodName = rs("name") & ""
		end if
		
		rs.Close
	end if

	BrandCount = 0
	if not blnLoadFailed then
		rs.Open "spListBrands4Product " & clng(request("ProductID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProdBrands = ""
			blnLoadFailed = true
		else
		
			do while not rs.EOF
				BrandCount = BrandCount + 1
				strProdBrands = strProdBrands & "," & rs("Abbreviation") & ""
				rs.MoveNext
			loop
			if strProdBrands <> "" then
				strProdBrands = mid(strProdBrands,2)
			end if
		end if
		
		rs.Close

	end if
	
	if not blnLoadFailed then
		rs.Open "spGetDeliverableVersionProperties " & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strDeliverable = ""
			blnLoadFailed = true
		else
			strDeliverable = rs("name") & "&nbsp;-&nbsp;" & rs("version")
			if rs("Revision") & "" <> "" then
				strDeliverable = strDeliverable & "," & rs("Revision") & ""
			end if
			if rs("Pass") & "" <> "" then
				strDeliverable = strDeliverable & "," & rs("Pass") & ""
			end if
			strDeliverable = strDeliverable & "&nbsp;"
	        strCategoryID = trim(rs("categoryid") & "")		
		end if
		
		rs.Close
	end if

	if not blnLoadFailed then
		'Get Distributions
		rs.Open "spGetDistributionALL " & clng(request("ProductID")) & "," & clng(request("VersionID")) & "," & clng(request("RootID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			blnLoadFailed = true
		else
			strPreinstallBrand = trim(rs("PreinstallBrand") & "")
			strPreloadBrand = trim(rs("PreloadBrand") & "")
			
			if  rs("PreinstallActual")  then
				strPreinstallActual = "checked"
			end if
			if rs("PreloadActual") then
				strPreloadActual = "checked"
			end if
			if rs("DropInBoxActual") then
				strDropInBoxActual = "checked"
				ShowDIBHW = ""
			end if
			if rs("WebActual") then
				strWebActual = "checked"
			end if
			if rs("SelectiveRestore") then
				strSR = "checked"
			end if
			if rs("ARCD") then
				strDRCD = "checked"
			end if

			if rs("DRDVD") then
				strDRDVD = "checked"
			end if

			if rs("RACD_Americas") then
				strRACDAmericas = "checked"
			end if

			if rs("RACD_APD") then
				strRACDAPD = "checked"
			end if

			if rs("RACD_EMEA") then
				strRACDEMEA = "checked"
			end if

			if rs("OSCD") then
				strOSCD = "checked"
			end if

			if rs("DOCCD") then
				strDOCCD = "checked"
			end if

			if rs("PatchActual") then
				strPatch = "checked"
			end if

            if rs("RCDOnly") then
				strRCDOnly = "checked"
			end if

			blnPreinstallDev = rs("PreinstallDev")
			blnPreinstallROM = rs("PreinstallROM")
			blnAR = rs("AR")
			blnCDImage = rs("CDIMage")
			blnISOImage = rs("ISOImage")
			blnFloppyDisk = rs("FloppyDisk")
			blnScriptpaq = rs("Scriptpaq")
			blnSoftpaq = rs("Softpaq")
			blnRompaq = rs("Rompaq")
            blnRCDOnly = rs("RCDOnly")

			strHWReqs = rs("DIBHWReq") & ""
		end if
		
		rs.Close
		
        if strCategoryID = "170" or strPatch = "checked" then
            strShowPatch=""
        else
            strShowPatch="none"
        end if

		'Load HW Requirements
		strAllHWReqs = ""
		if trim(strHWReqs) <> "" then
			dim HWArray
			HWArray=split(strHWReqs,",")
			for i = 0 to ubound(HWArray)
				if trim(HWArray(i))<> "" then
					strAllHWReqs = strAllHWReqs & "<tr><td><INPUT checked type=""checkbox"" style=""WIDTH:16;HEIGHT:16"" id=chkDIBHW name=chkDIBHW value=""" & trim(HWArray(i)) & """>&nbsp;" & trim(HWArray(i)) & "</td></tr>"
				end if
			next
		end if

		rs.Open "spListDIBHWRequirements",cn,adOpenForwardOnly
		do while not rs.EOF
			if instr(", " & strHWReqs & ",",", " & rs("Name") & ",") = 0 then
				strAllHWReqs = strAllHWReqs & "<tr><td><INPUT type=""checkbox"" style=""WIDTH:16;HEIGHT:16"" id=chkDIBHW name=chkDIBHW value=""" & rs("Name") & """>&nbsp;" & rs("Name") & "</td></tr>"
			end if
			rs.MoveNext
		loop
		rs.Close
		
		
		
	end if


%>



<h3>Edit Distribution<h3>
<h4><%=strDeliverable & " (" & strProdName & ")"%></h4>

<form ID=frmChange action="ChangeDistributionSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=txtPreinstall name=txtPreinstall value="<%=strPreinstallActual%>">
<INPUT type="hidden" id=txtPreLoad name=txtPreLoad value="<%=strPreloadActual%>">
<INPUT type="hidden" id=txtDropInBox name=txtDropInBox value="<%=strDropInBoxActual%>">
<INPUT type="hidden" id=txtWeb name=txtWeb value="<%=strWebActual%>">
<INPUT type="hidden" id=txtPatch name=txtPatch value="<%=strPatch%>">
<INPUT type="hidden" id=txtSR name=txtSR value="<%=strSR%>">
<INPUT type="hidden" id=txtAllBrands name=txtAllBrands value="<%=strProdBrands%>">
<INPUT type="hidden" id=txtPreinstallBrands name=txtPreinstallBrands value="<%=strPreinstallBrand%>">
<INPUT type="hidden" id=txtPreloadBrands name=txtPreloadBrands value="<%=strPreloadBrand%>">
<INPUT type="hidden" id=txtRCDOnly name=txtRCDOnly value="<%=strRCDOnly%>">
<%
			if strPreinstallBrand = "" then
				strPreinstallBrand = "All Brands"
			end if
			if strPreloadBrand = "" then
				strPreloadBrand = "All Brands"
			end if

%>

<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Distribution:&nbsp;&nbsp;</b></font></TD>
		<TD valign=top width=100%>
			<table width=100%><tr><td>
				<INPUT type="checkbox" <%=strPreinstallActual%> id=chkPreinstall name=chkPreinstall>
				&nbsp;<font size=2 face=verdana ID=ChangePI LANGUAGE=javascript onmouseover="return ChangePI_onmouseover()" onclick="return ChangePI_onclick()">Preinstall</font>
				<%if BrandCount > 1 then%>
					<span id=PickPreinstallBrand>&nbsp;-&nbsp;[&nbsp;<a href="javascript:PickBrand(1);"><span ID=lblPreinstall><%=strPreinstallBrand%></span></a>&nbsp;]&nbsp;</span>
				<%else%>
                    <span id=PickPreinstallBrand>&nbsp;</span>
				<%end if%>
				<BR>
				<INPUT type="checkbox" <%=strPreloadActual%> id=chkPreLoad name=chkPreLoad>
				&nbsp;<font size=2 face=verdana ID=ChangePL LANGUAGE=javascript onmouseover="return ChangePL_onmouseover()" onclick="return ChangePL_onclick()">Desktop Icon Launches Setup</font>
				<%if BrandCount > 1 then%>
					<span id=PickPreloadBrand>&nbsp;-&nbsp;[&nbsp;<a href="javascript:PickBrand(2);"><span ID=lblPreload><%=strPreloadBrand%></span></a>&nbsp;]</span>
				<%else%>
                    <span id=PickPreloadBrand>&nbsp;</span>
                <%end if%>
				<br>
				<INPUT type="checkbox" <%=strSR%> id=chkSR name=chkSR>
				&nbsp;<font size=2 face=verdana ID=ChangeSR LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeSR_onclick()">Software Setup (In App Recovery)</font>
                
                </td><td valign=top>

				<INPUT type="checkbox" <%=strDropInBoxActual%> id=chkDropInBox name=chkDropInBox LANGUAGE=javascript onclick="chkDropInBox_onclick();">
				&nbsp;<font size=2 face=verdana ID=ChangeDIB LANGUAGE=javascript onmouseover="return ChangeDIB_onmouseover()" onclick="return ChangeDIB_onclick()">Drop In Box</font>
				<br>
				<INPUT type="checkbox" <%=strWebActual%> id=chkWeb name=chkWeb>
				&nbsp;<font size=2 face=verdana ID=ChangeWeb LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeWeb_onclick()">Web Release</font>
                <br />
                <INPUT type="checkbox" <%=strRCDOnly%> id=chkRCDOnly name=chkRCDOnly>
				&nbsp;<font size=2 face=verdana ID=ChangeRO onmouseover="return ChangeRO_onmouseover()" onclick="return ChangeRO_onclick()">RCDOnly</font>
                <br />
                <span style="display:<%=strShowPatch%>">
				    <INPUT type="checkbox" <%=strPatch%> id=chkPatch name=chkPatch  onclick="javascript: PatchClicked();">
				    &nbsp;<font size=2 face=verdana ID=ChangePatch LANGUAGE=javascript onmouseover="return ChangePatch_onmouseover()" onclick="return ChangePatch_onclick()">Patch</font>
                </span>
                
                				
			</td></tr></table>
			

			<table width=100% border=0><TR><TD valign=top >
			<TAble cellspacing=0 border=1 bordercolor=navyblue width=100%><TR bgcolor=#ccccff><TD>
			<font size=2 face=verdana><b>Restore Media:</b></font><!--Installer Source-->
			</td></tr><tr bgcolor=Lavender><td>
			<table><TR><TD width=60>
			<b>Drivers:</td><td width=60 style="padding-top:0px !important"><INPUT type="checkbox" <%=strDRCD%> id=chkDRCD name=chkDRCD><font size=2 face=verdana ID=ChangeDRCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDRCD_onclick()">DRCD</font>
			</td><td>
			<INPUT type="checkbox" <%=strDRDVD%> id=chkDRDVD name=chkDRDVD><font size=2 face=verdana ID=ChangeDRDVD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDRDVD_onclick()">DRDVD</font>
			</td>

			<td width=60><INPUT type="hidden" id=txtDRCD name=txtDRCD value="<%=strDRCD%>">
			</td><td width=60>
			<INPUT type="hidden" id=txtDRDVD name=txtDRDVD value="<%=strDRDVD%>">
			</td>		
			</tr></table>
			
			<!--<font size=1 color=blue face=verdana>&nbsp;Coming soon!&nbsp;</font>-->
			<HR style="HEIGHT: 1px" color=LightSteelBlue>
			<table><TR><TD width=60>
			<b>RACD:</td><td width=60>
			<INPUT type="checkbox" <%=strRACDEMEA%> id=chkRACDEMEA name=chkRACDEMEA><font size=2 face=verdana ID=ChangeRACDEMEA LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDEMEA_onclick()">EMEA</font>
			</td><td>
			<INPUT type="checkbox" <%=strRACDAmericas%> id=chkRACDAmericas name=chkRACDAmericas><font size=2 face=verdana ID=ChangeRACDAmericas LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDAmericas_onclick()">Americas</font>
			</td><td>
			<INPUT type="checkbox" <%=strRACDAPD%> id=chkRACDAPD name=chkRACDAPD><font size=2 face=verdana ID=ChangeRACDAPD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDAPD_onclick()">APD</font>
			</td>

			<td width=60><INPUT type="hidden" id=txtRACDEMEA name=txtRACDEMEA value="<%=strRACDEMEA%>">
			</td><td>
			<INPUT type="hidden" id=txtRACDAmericas name=txtRACDAmericas value="<%=strRACDAmericas%>">
			</td><td>
			<INPUT type="hidden" id=txtRACDAPD name=txtRACDAPD value="<%=strRACDAPD%>">
			</td>
			</tr></table>
			
			<HR style="HEIGHT: 1px" color=LightSteelBlue>
			<table><TR><TD width=60>
			<b>Other:</TD><TD width=60><INPUT type="checkbox" <%=strOSCD%> id=chkOSCD name=chkOSCD><font size=2 face=verdana ID=ChangeOSCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeOSCD_onclick()">OSCD</font>
			</TD><TD>
			<INPUT type="checkbox" <%=strDocCD%> id=chkDocCD name=chkDocCD><font size=2 face=verdana ID=ChangeDocCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDocCD_onclick()">Doc CD</font>
			</td>

			<TD width=60><INPUT type="hidden" id=txtOSCD name=txtOSCD value="<%=strOSCD%>">
			</TD><TD>
			<INPUT type="hidden" id=txtDocCD name=txtDocCD value="<%=strDocCD%>">
			</td>
			
			</tr></table>
			</TD>
			<!--<TD valign=top>
			<font size=2 face=verdana><b>App Installer Category:<BR></b></font>
			<INPUT type="radio" id=radio1 name=radio1 checked> Hardware Device Driver<BR>
			<INPUT type="radio" id=radio1 name=radio1> Software Application
			</TD>-->
			</TR></table>
			</td><td  ID=HWReqCell style="Display:<%=ShowDIBHW%>" valign=top>
			<TAble cellspacing=0 border=1 bordercolor=navyblue width=100% bgcolor=Lavender><TR height=100% bgcolor=#ccccff><TD>
			<font size=2 face=verdana><b>DIB Optical Drive Req:</b></font>
			</td></tr><tr><td>
<!--				<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 140px; BACKGROUND-COLOR: white" id=DIV1>-->
				<div style="OVERFLOW-Y: auto; WIDTH: 100%; HEIGHT: 108px; BACKGROUND-COLOR: lavendar" id=DIV1>
					<TABLE ID="ReqTable" width=100%>
						<!--<THEAD><TR style="position:relative;top:expression(document.getElementById('DIV1').scrollTop-2);"><TD  bgcolor=gainsboro nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;</TD><TD  bgcolor=gainsboro style="width=100%; BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Hardware&nbsp;</TD></TR></THEAD>-->
						<%=strAllHWReqs%>
						<tr><td><font size=1><a href="javascript: AddHWReq();">Add New</a></td></tr>
					</TABLE>
            
				</div>			
			</TD>
			</TR></TABLE>			
			</td></tr></Table>
		</TD>
	</TR>
	
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Scope:</b></font></TD>
		<TD valign=top>
		<INPUT type="radio" id=optThis name=optScope value="1">&nbsp;<font size=2 face=verdana ID="ChangeThis" LANGUAGE=javascript onclick="return ChangeThis_onclick()" onmouseover="return ChangeThis_onmouseover()">Change this version only</font><BR>
		<INPUT type="radio" id=optFuture name=optScope value="2" checked>&nbsp;<font size=2 face=verdana Id="ChangeDefault" LANGUAGE=javascript onclick="return ChangeDefault_onclick()" onmouseover="return ChangeDefault_onmouseover()">Change this version and all future versions</font><BR>
		<!--<INPUT type="radio" id=optAll name=optScope value="3" >&nbsp;<font size=2 face=verdana>Change All Existing and Future Versions</font><BR>-->
		</TD>
	</TR>
</table>
</form>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript  onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></TD>
	</TR>
</TABLE>


<%
	set rs= nothing
	set cn=nothing
%>

</BODY>
</HTML>
				<!--<%if not (blnPreinstallDev or blnPreinstallROM) then%>
				<font size=2 face=verdana color=red> - Not Packaged for Preinstall</font><BR>
			<%else%>
				<BR>
			<%end if%>
			<%if not (blnCDImage or blnISOImage or blnAR) then%>
				<font size=2 color=red face=verdana> - Not Packaged for Drop in Box</font><BR>
			<%else%>
			
				<BR>
			<%if not(blnFloppyDisk or blnScriptpaq or blnSoftpaq or blnRompaq) then%>
				<font color=red size=2 face=verdana> - Not Packaged for Web Release</font><BR>
			<%else%>
				<BR>
			<%end if%>
				
			<%end if%>
			
			-->
