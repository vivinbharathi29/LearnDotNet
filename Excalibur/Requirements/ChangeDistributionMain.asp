<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {
	frmChange.submit();
}

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
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

function ChangeMinImgRec_onmouseover() {
    if (!frmChange.chkMinImgRec.disabled)
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

function ChangeSR_onclick() {
	if (frmChange.chkSR.checked)
		frmChange.chkSR.checked = false;
	else if (!frmChange.chkSR.disabled)
		frmChange.chkSR.checked = true;
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

function ChangeMinImgRec_onclick() {
	if (frmChange.chkMinImgRec.checked)
		frmChange.chkMinImgRec.checked = false;
	else if (!frmChange.chkMinImgRec.disabled)
		frmChange.chkMinImgRec.checked = true;
}

function ChangeRACDAPD_onclick() {
    if (frmChange.chkRACDAPD.checked)
        frmChange.chkRACDAPD.checked = false;
    else if (!frmChange.chkRACDAPD.disabled)
        frmChange.chkRACDAPD.checked = true;
}


function ChangePatch_onclick() {
	if (frmChange.chkPatch.checked)
		frmChange.chkPatch.checked = false;
	else if (!frmChange.chkPatch.disabled)
	    frmChange.chkPatch.checked = true;

    PatchClicked();
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
        frmChange.chkMinImgRec.checked = false;
        frmChange.chkMinImgRec.disabled = true;
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
        
    }
    else {
        frmChange.chkPreinstall.disabled = false;
        frmChange.chkPreLoad.disabled = false;
        frmChange.chkRACDAPD.disabled = false;
        frmChange.chkMinImgRec.disabled = false;
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
}


	function PickBrand(strType){
	var strID;
	

	
	if (strType==1)
		strID = window.showModalDialog("../Target/DistributionBrand.asp?AllBrands=" + frmChange.txtAllBrands.value + "&SelectedBrands=" + frmChange.txtPreinstallBrands.value ,"","dialogWidth:250px;dialogHeight:350px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	else
		strID = window.showModalDialog("../Target/DistributionBrand.asp?AllBrands=" + frmChange.txtAllBrands.value + "&SelectedBrands=" + frmChange.txtPreloadBrands.value ,"","dialogWidth:250px;dialogHeight:350px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
	if (typeof(strID) != "undefined")
		{
		if (frmChange.txtAllBrands.value == strID)
			strID = ""
			
		if (strType==1)
			{
			frmChange.txtPreinstallBrands.value = strID;
			if (strID == "")
				lblPreinstall.innerText = "All Brands";
			else
				lblPreinstall.innerText = strID;
			}
		else
			{
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
<BODY onload="PatchClicked();"  bgcolor=Ivory>

<%

	dim cn
	dim rs
	dim i
    dim strBrandDisplay
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strPreinstallActual
	dim strPatchActual
	dim strMinImgRecActual
	dim strPreLoadActual
	dim strDropInBoxActual
	dim strWebActual
	dim blnPreinstallDev
	dim blnPreinstallROM
	dim blnAR
	dim blnCDImage
	dim blnISOImage
	dim blnFloppyDisk
	dim blnScriptpaq
	dim blnSoftpaq
	dim blnRompaq
	dim strDRCD
	dim strDRDVD
	dim strRACDAmericas
	dim strRACDEMEA
	dim strRACDAPD
	dim strDOCCD
	dim strOSCD
	dim strSelectiveRestore
	dim strPreinstallBrand
	dim strPreloadBrand
	dim strProdBrands
	dim BrandCount
    dim strCategoryID
    dim strShowPatch


	strPreinstallActual = ""
	strPatchActual = ""
	strMinImgRecActual = ""
	strPreLoadActual = ""
	strDropInBoxActual = ""
	strWebActual = ""
	strSelectiveRestore = ""
	strDRCD  = ""
	strDRDVD = ""
	strRACDAmericas = ""
	strRACDEMEA = ""
	strRACDAPD = ""
	strDOCCD = ""
	strOSCD = ""
	strPreinstallBrand = ""
	strPreloadBrand = ""
	strProdBrands = ""
    strCategoryID=""
	
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	
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
		rs.Open "spGetRootProperties4Version " & clng(request("RootID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strDeliverable = ""
			blnLoadFailed = true
		else
			strDeliverable = rs("name") & ""
		end if
        strCategoryID = trim(rs("category") & "")		
		
		rs.Close
	end if
    strBrandDisplay = ""

	if not blnLoadFailed then
		rs.Open "spGetDistributionRoot " & clng(request("ProductID")) & ","  & clng(request("RootID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			blnLoadFailed = true
		else
		
			strPreinstallBrand = trim(rs("PreinstallBrand") & "")
			strPreloadBrand = trim(rs("PreloadBrand") & "")
		
			if  rs("Preinstall")  then
				strPreinstallActual = "checked"
			end if
			if  rs("Patch")  then
				strPatchActual = "checked"
	            strBrandDisplay = "none"
    		end if
			if  rs("MinImageRecovery")  then
				strMinImgRecActual = "checked"
			end if
			if rs("Preload") then
				strPreloadActual = "checked"
			end if
			if rs("DIB") then
				strDropInBoxActual = "checked"
			end if
			if rs("Web") then
				strWebActual = "checked"
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
			if rs("SelectiveRestore") then
				strSelectiveRestore = "checked"
			end if
			
			blnPreinstallDev = rs("PreinstallDev")
			blnPreinstallROM = rs("PreinstallROM")
			blnAR = rs("AR")
			blnCDImage = rs("CDIMage")
			blnISOImage = rs("ISOImage") & ""
			blnFloppyDisk = rs("FloppyDisk")
			blnScriptpaq = rs("Scriptpaq")
			blnSoftpaq = rs("Softpaq")
			blnRompaq = rs("Rompaq")			
		
			if blnISOImage="" then
				blnISOImage = 0
			end if
		end if
		
		
		
		
		rs.Close
	end if

    if strCategoryID = "170" or strPatch = "checked" then
        strShowPatch=""
    else
        strShowPatch="none"
    end if
%>



<h3>Edit Default Distribution<h3>
<h4><%=strDeliverable & " (" & strProdName & ")"%></h4>

<form ID=frmChange action="ChangeDistributionSave.asp" method=post>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
<INPUT type="hidden" id=txtAllBrands name=txtAllBrands value="<%=strProdBrands%>">
<INPUT type="hidden" id=txtPreinstallBrands name=txtPreinstallBrands value="<%=strPreinstallBrand%>">
<INPUT type="hidden" id=txtPreloadBrands name=txtPreloadBrands value="<%=strPreloadBrand%>">
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
		<TD valign=top><font size=2 face=verdana><b>Distribution:</b></font></TD>
		<TD valign=top width=100%>
		    <table width=100%><tr><td>
			<INPUT type="checkbox" <%=strPreinstallActual%> id=chkPreinstall name=chkPreinstall>&nbsp;<font size=2 face=verdana ID=ChangePI LANGUAGE=javascript onmouseover="return ChangePI_onmouseover()" onclick="return ChangePI_onclick()">Preinstall</font>
				<%if BrandCount > 1 then%>
					<span style="display:<%=strBrandDisplay%>" id=PickPreinstallBrand>&nbsp;-&nbsp;[&nbsp;<a href="javascript:PickBrand(1);"><span ID=lblPreinstall><%=strPreinstallBrand%></span></a>&nbsp;]&nbsp;</span>
                <%else%>
                    <span id=PickPreinstallBrand>&nbsp;</span>
               <%end if%>
			</td><td>
    			<INPUT type="checkbox" <%=strDropInBoxActual%> id=chkDropInBox name=chkDropInBox>&nbsp;<font size=2 face=verdana ID=ChangeDIB LANGUAGE=javascript onmouseover="return ChangeDIB_onmouseover()" onclick="return ChangeDIB_onclick()">Drop In Box</font>
			    </td></TR><tr><td>
			<INPUT type="checkbox" <%=strPreloadActual%> id=chkPreLoad name=chkPreLoad>&nbsp;<font size=2 face=verdana ID=ChangePL LANGUAGE=javascript onmouseover="return ChangePL_onmouseover()" onclick="return ChangePL_onclick()">Desktop Icon Launches Setup</font>
				<%if BrandCount > 1 then%>
					<span style="display:<%=strBrandDisplay%>" id=PickPreloadBrand>&nbsp;-&nbsp;[&nbsp;<a href="javascript:PickBrand(2);"><span ID=lblPreload><%=strPreloadBrand%></span></a>&nbsp;]</span>
                <%else%>
                    <span id=PickPreloadBrand>&nbsp;</span>
               <%end if%>
			</td><td>
			    <INPUT type="checkbox" <%=strWebActual%> id=chkWeb name=chkWeb>&nbsp;<font size=2 face=verdana ID=ChangeWeb LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeWeb_onclick()">Web Release</font>
			</td></TR><tr><td>
                <INPUT type="checkbox" <%=strSelectiveRestore%> id=chkSR name=chkSR>&nbsp;<font size=2 face=verdana ID=ChangeSR LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeSR_onclick()">Software Setup (Selective Restore)</font>
			</td><td>
                <span style="display:<%=strShowPatch%>">
			    <INPUT type="checkbox" <%=strPatchActual%> id=chkPatch name=chkPatch onclick="javascript: PatchClicked();">&nbsp;<font size=2 face=verdana ID=ChangePatch LANGUAGE=javascript onmouseover="return ChangePatch_onmouseover()" onclick="return ChangePatch_onclick()">Patch</font>
                </span>
			</td></TR><tr style="display:none"><td>
			    <INPUT type="checkbox" <%=strMinImgRecActual%> id=chkMinImgRec name=chkMinImgRec>&nbsp;<font size=2 face=verdana ID=ChangeMinImgRec LANGUAGE=javascript onmouseover="return ChangeMinImgRec_onmouseover()" onclick="return ChangeMinImgRec_onclick()">Minimum Image Recovery</font>
			</td></TR></table>
			<table width=100% border=0><TR><TD valign=top >
			<TAble cellspacing=0 border=1 bordercolor=navyblue width=100%><TR bgcolor=#ccccff><TD>
			<font size=2 face=verdana><b>Restore Media:</b></font><!--Installer Source-->
			</td></tr><tr bgcolor=Lavender><td>
			<table><TR><TD width=60>
			<b>Drivers:</td><td width=60><INPUT type="checkbox" <%=strDRCD%> id=chkDRCD name=chkDRCD><font size=2 face=verdana ID=ChangeDRCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDRCD_onclick()">DRCD</font>
			</td><td width=60>
			<INPUT type="checkbox" <%=strDRDVD%> id=chkDRDVD name=chkDRDVD><font size=2 face=verdana ID=ChangeDRDVD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDRDVD_onclick()">DRDVD</font>
			</td></tr></table>
			<!--<font size=1 color=blue face=verdana>&nbsp;Coming soon!&nbsp;</font>-->
			<HR style="HEIGHT: 1px" color=LightSteelBlue>
			<table><TR><TD width=60>
			<b>RACD:</td><td width=60>
			<INPUT type="checkbox" <%=strRACDEMEA%> id=chkRACDEMEA name=chkRACDEMEA><font size=2 face=verdana ID=ChangeRACDEMEA LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDEMEA_onclick()">EMEA</font>
			</td><td>
			<INPUT type="checkbox" <%=strRACDAmericas%> id=chkRACDAmericas name=chkRACDAmericas><font size=2 face=verdana ID=ChangeRACDAmericas LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDAmericas_onclick()">Americas</font>
			</td><td>
			<INPUT type="checkbox" <%=strRACDAPD%> id=chkRACDAPD name=chkRACDAPD><font size=2 face=verdana ID=ChangeRACDAPD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeRACDAPD_onclick()">APD</font>
			</td></tr></table>
			<HR style="HEIGHT: 1px" color=LightSteelBlue>
			<table><TR><TD width=60>
			<b>Other:</TD><TD width=60><INPUT type="checkbox" <%=strOSCD%> id=chkOSCD name=chkOSCD><font size=2 face=verdana ID=ChangeOSCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeOSCD_onclick()">OSCD</font>
			</TD><TD>
			<INPUT type="checkbox" <%=strDocCD%> id=chkDocCD name=chkDocCD><font size=2 face=verdana ID=ChangeDocCD LANGUAGE=javascript onmouseover="return ChangeWeb_onmouseover()" onclick="return ChangeDocCD_onclick()">Doc CD</font>
			</td></tr></table>
			</TD>
			<!--<TD valign=top>
			<font size=2 face=verdana><b>App Installer Category:<BR></b></font>
			<INPUT type="radio" id=radio1 name=radio1 checked> Hardware Device Driver<BR>
			<INPUT type="radio" id=radio1 name=radio1> Software Application
			</TD>-->
			</TR></table>
			</td><td  style="Display:none" valign=top>
			<TAble cellspacing=0 border=1 bordercolor=navyblue width=100% bgcolor=Lavender><TR height=100% bgcolor=#ccccff><TD>
			<font size=2 face=verdana><b>DIB Hardware Requirements:</b></font>
			</td></tr><tr><td>
<!--				<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 140px; BACKGROUND-COLOR: white" id=DIV1>-->
				<div style="OVERFLOW-Y: auto; WIDTH: 100%; HEIGHT: 108px; BACKGROUND-COLOR: lavendar" id=DIV1>
					<TABLE width=100%>
						<!--<THEAD><TR style="position:relative;top:expression(document.getElementById('DIV1').scrollTop-2);"><TD  bgcolor=gainsboro nowrap style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;</TD><TD  bgcolor=gainsboro style="width=100%; BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Hardware&nbsp;</TD></TR></THEAD>-->
						<tr><td><INPUT type="checkbox" style="WIDTH:16;HEIGHT:16" id=checkbox1 name=checkbox1>&nbsp;CD-RW</td></tr>
						<tr><td><INPUT type="checkbox" style="WIDTH:16;HEIGHT:16" id=checkbox1 name=checkbox1>&nbsp;DVD+RW</td></tr>
						<tr><td><font size=1><a href="">Add New</a></td></tr>
					</TABLE>
            
				</div>			
			</TD>
			</TR></TABLE>			
			</td></tr></Table>
			
		</TD>
</table>
</form>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript  onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD>
	</TR>
</TABLE>


<%
	set rs= nothing
	set cn=nothing
%>

</BODY>
</HTML>
