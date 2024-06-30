<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload(pulsarplusDivId) {
	if (typeof(txtSuccess) != "undefined"){
	    if (txtSuccess.value != "0") {
	        //close window
	        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
	            if (pulsarplusDivId == 'targetadvanced') {
	                parent.window.parent.ModDistResult(txtSuccess.value);
	                parent.window.parent.closeExternalPopup();
	            }
	            else {
	                parent.window.parent.closeExternalPopup();
	                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
	            }
	        }
	        else if (IsFromPulsarPlus()) {
	            window.parent.parent.parent.popupCallBack(1);
	            ClosePulsarPlusPopup();
	        }
	        else {
	            if (parent.window.parent.document.getElementById('modal_dialog')) {
	                //save value and return to parent page: ---
	                parent.window.parent.ModDistResult(txtSuccess.value);
	                parent.window.parent.modalDialog.cancel();
	            } else {
	                window.returnValue = txtSuccess.value;
	                window.close();
	            }
	        }
		}
		//else
		//	document.write ("Unable to update Distribution.  An unexpected error occurred.");	
	}
	//else
	//	{
	//	document.write ("Unable to update Distribution.  An unexpected error occurred.");
	//	}

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload('<%=Request("pulsarplusDivId")%>')">

<%

	dim i
	dim cn
	dim rs
	dim cm
	dim blnSuccess
	dim RowsEffected
	dim strPreinstall
	dim strPreinstallChange
	dim strPreload
	dim strPreloadChange	
	dim strDropInBox
	dim strWeb
	dim strMethods
	dim blnDRCD
	dim blnDRCDChange	
	dim blnDRDVD
	dim blnDRDVDChange	
	dim blnSR
	dim blnSRChange	
	dim blnRACDAmericas
	dim blnRACDAPD
	dim blnRACDEMEA
	dim blnOSCD
	dim blnDocCD
	dim strDistributionChange
	dim strPINDistributionChange
	dim strNewImageSummary
    dim blnPatch
    dim blnRCDOnly


	strNewImageSummary = "|"
	
	strPINDistributionChange = ""
	strDistributionChange = ""
	if request("chkPreinstall") = "on" then
		blnPreinstall = true
		if request("txtPreinstall") = "" then
			strPINDistributionChange = strPINDistributionChange & "From Not-Preinstall to Preinstall, "
			strDistributionChange = strDistributionChange & "From Not-Preinstall to Preinstall, "
		end if
	else
		blnPreinstall = false
		if request("txtPreinstall") <> "" then
			strPINDistributionChange = strPINDistributionChange & "From Preinstall to Not-Preinstall, "
			strDistributionChange = strDistributionChange & "From Preinstall to Not-Preinstall, "
		end if
	end if
	if request("chkPreload") = "on" then
		blnPreload = true
		if request("txtPreload") = "" then
			strPINDistributionChange = strPINDistributionChange & "From Not-Preload to Preload, "
			strDistributionChange = strDistributionChange & "From Not-Preload to Preload, "
		end if
	else
		blnPreload = false
		if request("txtPreload") <> "" then
			strPINDistributionChange = strPINDistributionChange & "From Preload to Not-Preload, "
			strDistributionChange = strDistributionChange & "From Preload to Not-Preload, "
		end if
	end if
	if request("chkDropInBox") = "on" then
		blnDropInBox = true
		if request("txtDropInBox") = "" then
			strDistributionChange = strDistributionChange & "From Not-DropInBox to DropInBox, "
		end if
	else
		blnDropInBox = false
		if request("txtDropInBox") <> "" then
			strDistributionChange = strDistributionChange & "From DropInBox to Not-DropInBox, "
		end if
	end if
	if request("chkWeb") = "on" then
		blnWeb = true
		if request("txtWeb") = "" then
			strDistributionChange = strDistributionChange & "From Not-Web to Web, "
		end if
	else
		blnWeb = false
		if request("txtWeb") <> "" then
			strDistributionChange = strDistributionChange & "From Web to Not-Web, "
		end if
	end if

	if request("chkPatch") = "on" then
		blnPatch = true
		if request("txtPatch") = "" then
			strDistributionChange = strDistributionChange & "From Not-Patch to Patch, "
		end if
	else
		blnPatch = false
		if request("txtPatch") <> "" then
			strDistributionChange = strDistributionChange & "From Patch to Not-Patch, "
		end if
	end if


	if request("chkSR") = "on" then
		blnSR = true
		if request("txtSR") = "" then
			strPINDistributionChange = strPINDistributionChange & "From Not-SoftwareSetup to SoftwareSetup, "
			strDistributionChange = strDistributionChange & "From Not-SoftwareSetup to SoftwareSetup, "
		end if
	else
		blnSR = false
		if request("txtSR") <>"" then
			strPINDistributionChange = strPINDistributionChange & "From SoftwareSetup to Not-SoftwareSetup, "
			strDistributionChange = strDistributionChange & "From SoftwareSetup to Not-SoftwareSetup, "
		end if
	end if
	if request("chkDRCD") = "on" then
		blnDRCD = true
		if request("txtDRCD") = "" then
			strPINDistributionChange = strPINDistributionChange & "From Not-DRCD to DRCD, "
			strDistributionChange = strDistributionChange & "From Not-DRCD to DRCD, "			
		end if
	else
		blnDRCD = false
		if request("txtDRCD") <> "" then
			strPINDistributionChange = strPINDistributionChange & "From DRCD to Not-DRCD, "
			strDistributionChange = strDistributionChange & "From DRCD to Not-DRCD, "
		end if
	end if
	if request("chkDRDVD") = "on" then
		blnDRDVD = true
		if request("txtDRDVD") = "" then
			strPINDistributionChange = strPINDistributionChange & "From Not-DRDVD to DRDVD, "
			strDistributionChange = strDistributionChange & "From Not-DRDVD to DRDVD, "
		end if
	else
		blnDRDVD = false
		if request("txtDRDVD") <> "" then
			strPINDistributionChange = strPINDistributionChange & "From DRDVD to Not-DRDVD, "
			strDistributionChange = strDistributionChange & "From DRDVD to Not-DRDVD, "
		end if
	end if
	if request("chkRACDAmericas") = "on" then
		blnRACDAmericas = true
		if request("txtRACDAmericas") = "" then
			strDistributionChange = strDistributionChange & "From Not-RACDAmericas to RACDAmericas, "
		end if
	else
		blnRACDAmericas = false
		if request("txtRACDAmericas") <> "" then
			strDistributionChange = strDistributionChange & "From RACDAmericas to Not-RACDAmericas, "
		end if
	end if
	if request("chkRACDAPD") = "on" then
		blnRACDAPD = true
		if request("txtRACDAPD") = "" then
			strDistributionChange = strDistributionChange & "From Not-RACDAPD to RACDAPD, "
		end if
	else
		blnRACDAPD = false
		if request("txtRACDAPD") <> "" then
			strDistributionChange = strDistributionChange & "From RACDAPD to Not-RACDAPD, "
		end if
	end if
	if request("chkRACDEMEA") = "on" then
		blnRACDEMEA = true
		if request("txtRACDEMEA") = "" then
			strDistributionChange = strDistributionChange & "From Not-RACDEMEA to RACDEMEA, "
		end if
	else
		blnRACDEMEA = false
		if request("txtRACDEMEA") <> "" then
			strDistributionChange = strDistributionChange & "From RACDEMEA to Not-RACDEMEA, "
		end if
	end if
	if request("chkDocCD") = "on" then
		blnDocCD = true
		if request("txtDocCD") = "" then
			strDistributionChange = strDistributionChange & "From Not-DocCD to DocCD, "
		end if
	else
		blnDocCD = false
		if request("txtDocCD") <> "" then
			strDistributionChange = strDistributionChange & "From DocCD to Not-DocCD, "
		end if
	end if
	if request("chkOSCD") = "on" then
		blnOSCD = true
		if request("txtOSCD") = "" then
			strDistributionChange = strDistributionChange & "From Not-OSCD to OSCD, "
		end if
	else
		blnOSCD = false
		if request("txtOSCD") <> "" then
			strDistributionChange = strDistributionChange & "From OSCD to Not-OSCD, "
		end if
	end if

    if request("chkRCDOnly") = "on" then
		blnRCDOnly = true
		if request("txtRCDOnly") = "" then
			strDistributionChange = strDistributionChange & "From Not-RCDOnly to RCDOnly, "
		end if
	else
		blnRCDOnly = false
		if request("txtRCDOnly") <> "" then
			strDistributionChange = strDistributionChange & "From RCDOnly to Not-RCDOnly, "
		end if
	end if
		
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	'Get User
	dim CurrentDomain
	dim Currentuser
	dim CurrentUserID
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
		CurrentUserID = rs("ID") 
	end if
	rs.Close	


	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")

    cm.ActiveConnection = cn
    cm.CommandText = "spUpdateDistributions"
    cm.CommandType = &H0004
       
	Set p = cm.CreateParameter("@ProductID",adInteger, &H0001)
	p.Value = request("txtProductID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID",adInteger, &H0001)
	p.Value = request("txtVersionID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Preinstall",adBoolean, &H0001)
	p.Value = blnPreinstall
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@PreinstallBrand",adVarChar, &H0001,50)
	p.Value = left(request("txtPreinstallBrands"),50)
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@Preload",adBoolean, &H0001)
	p.Value = blnPreload
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@PreloadBrand",adVarChar, &H0001,50)
	p.Value = left(request("txtPreloadBrands"),50)
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DropInBox",adBoolean, &H0001)
	p.Value = blnDropInBox
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DIBHWReq",adVarChar, &H0001,256)
	if blnDropInBox then
		p.Value = left(request("chkDIBHW"),256)
	else
		p.Value = ""
	end if
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@PINDistributionChange",adVarChar, &H0001,256)
		p.Value = strPINDistributionChange
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DistributionChange",adVarChar, &H0001,256)
		p.Value = left(strDistributionChange,256)
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
	p.Value = CurrentUserID
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@Web",adBoolean, &H0001)
	p.Value = blnWeb
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@SR",adBoolean, &H0001)
	p.Value = blnSR
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@DRCD",adBoolean, &H0001)
	p.Value = blnDRCD
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DRDVD",adBoolean, &H0001)
	p.Value = blnDRDVD
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@RACD_Americas",adBoolean, &H0001)
	p.Value = blnRACDAmericas
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@RACD_APD",adBoolean, &H0001)
	p.Value = blnRACDAPD
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@RACD_EMEA",adBoolean, &H0001)
	p.Value = blnRACDEMEA
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DocCD",adBoolean, &H0001)
	p.Value = blnDocCD
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@OSCD",adBoolean, &H0001)
	p.Value = blnOSCD
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@Patch",adTinyInt, &H0001)
	if blnPatch then
        p.Value = 1
    else
        p.Value = 0 
    end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@MinImgRec",adBoolean, &H0001)
	p.Value = 0 'blnMinImgRec
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@RCDOnly",adBoolean, &H0001)
	p.Value = blnRCDOnly
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@Type",adInteger, &H0001)
	p.Value = clng(request("optScope"))
	cm.Parameters.Append p

    cm.Execute RowsEffected
	Set cm = Nothing
	
	if RowsEffected <> 1 then
		blnSuccess = false
		cn.RollbackTrans
	else
		blnSuccess = true
		cn.CommitTrans
	end if	
	


	cn.execute "spCorrectImageSummary " & clng(request("txtProductID")) & "," &  clng(request("txtVersionID")) & "," & clng(request("optScope"))
	rs.Open "spGetImageSummary " & clng(request("txtProductID")) & "," &  clng(request("txtVersionID")) & "," & clng(request("optScope")),cn,adOpenStatic

	if not(rs.eof and rs.BOF) then
		strNewImageSummary = "|" & rs("ImageSummary") 
	end if
	
	rs.Close
	


	
	set rs=nothing
	set cn=nothing
	
	strMethods = ""
	if blnPreinstall then
		strMethods = strMethods & ", Preinstall"
	end if
	if blnPreload then
		strMethods = strMethods & ", Preload"
	end if
	if blnDropInBox then
		strMethods = strMethods & ", DIB"
	end if
	if blnWeb then
		strMethods = strMethods & ", Web"
	end if
	if blnSR then
		strMethods = strMethods & ", Selective Restore"
	end if
	if blnDRCD then
		strMethods = strMethods & ", DRCD"
	end if

	if blnDRDVD then
		strMethods = strMethods & ", DRDVD"
	end if

	if blnRACDAmericas then
		strMethods = strMethods & ", RACD-Americas"
	end if

	if blnRACDAPD then
		strMethods = strMethods & ", RACD-APD"
	end if

	if blnRACDEMEA then
		strMethods = strMethods & ", RACD-EMEA"
	end if
	if blnDocCD then
		strMethods = strMethods & ", Doc CD"
	end if
	if blnOSCD then
		strMethods = strMethods & ", OSCD"
	end if
	if blnPatch then
		strMethods = strMethods & ", Patch"
	end if
    if blnRCDOnly then
		strMethods = strMethods & ", RCDOnly"
	end if

	if len(strMethods) > 0 then
		strMethods = mid(strMethods,3)
	end if
%>
<%if blnSuccess then%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="<%=strMethods & strNewImageSummary%>">
<INPUT type="text" style="display:none" id=txtChange name=txtChange value="<%=strPINDistributionChange%>">
<%else%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="0">
<%end if%>
</BODY>
</HTML>
