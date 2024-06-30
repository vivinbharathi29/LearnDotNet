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
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value != "0")
		{
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.TargetQuickSaveCallBack(txtSuccess.value);
		        ClosePulsarPlusPopup();
		    } else {
		        if (parent.window.parent.document.getElementById('modal_dialog')) {
		            parent.window.parent.modalDialog.cancel(true);
		        } else {
		            window.returnValue = txtSuccess.value;
		            window.close();
		        }
		    }
		}
	//	else
	//		document.write ("Unable to update Distribution.  An unexpected error occurred.");	
	//	}
	//else
	//	{
	//	document.write ("Unable to update Distribution.  An unexpected error occurred.");
		}

}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

	dim i
	dim cn
	dim rs
	dim cm
	dim blnSuccess
	dim RowsEffected
	dim blnPreinstall
	dim blnPatch
	dim blnMinImgRec
	dim blnPreload
	dim blnDropInBox
	dim blnWeb
	dim blnSR
	dim blnDRCD
	dim blnDRDVD
	dim blnRACDAmericas
	dim blnRACDAPD
	dim blnRACDEMEA
	dim blnOSCD
	dim blnDocCD
	dim strMethods
	dim strNewImageSummary
	
	strNewImageSummary = "|"
	
	if request("chkPreinstall") = "on" then
		blnPreinstall = true
	else
		blnPreinstall = false
	end if

	if request("chkPatch") = "on" then
		blnPatch = true
	else
		blnPatch = false
	end if

	if request("chkMinImgRec") = "on" then
		blnMinImgRec = true
	else
		blnMinImgRec = false
	end if

	if request("chkPreload") = "on" then
		blnPreload = true
	else
		blnPreload = false
	end if
	if request("chkDropInBox") = "on" then
		blnDropInBox = true
	else
		blnDropInBox = false
	end if
	if request("chkWeb") = "on" then
		blnWeb = true
	else
		blnWeb = false
	end if

	if request("chkSR") = "on" then
		blnSR = true
	else
		blnSR = false
	end if

	if request("chkDRCD") = "on" then
		blnDRCD = true
	else
		blnDRCD = false
	end if

	if request("chkDRDVD") = "on" then
		blnDRDVD = true
	else
		blnDRDVD = false
	end if
	if request("chkRACDAmericas") = "on" then
		blnRACDAmericas = true
	else
		blnRACDAmericas = false
	end if
	if request("chkRACDAPD") = "on" then
		blnRACDAPD = true
	else
		blnRACDAPD = false
	end if
	if request("chkRACDEMEA") = "on" then
		blnRACDEMEA = true
	else
		blnRACDEMEA = false
	end if
	if request("chkDocCD") = "on" then
		blnDocCD = true
	else
		blnDocCD = false
	end if
	if request("chkOSCD") = "on" then
		blnOSCD = true
	else
		blnOSCD = false
	end if
	
	
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")

    cm.ActiveConnection = cn
    cm.CommandText = "spUpdateDistributions"
    cm.CommandType = &H0004
       
	Set p = cm.CreateParameter("@ProductID",adInteger, &H0001)
	p.Value = request("txtProductID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID",adInteger, &H0001)
	p.Value = request("txtRootID")
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

	Set p = cm.CreateParameter("@DIBHWReq",adVarChar, &H0001,30)
	p.Value = ""
    cm.Parameters.Append p
     
	Set p = cm.CreateParameter("@PINDistributionChange",adVarChar, &H0001,256)
		p.Value = ""
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@DistributionChange",adVarChar, &H0001,256)
		p.Value = ""
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@UserID",adInteger, &H0001)
	p.Value = 0
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

	Set p = cm.CreateParameter("@Patch", adTinyInt, &H0001)
    if blnPatch then
	    p.Value = 1
    else
	    p.Value = 0
    end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@MinImgRec",adBoolean, &H0001)
	p.Value = blnMinImgRec
    cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@Type",adInteger, &H0001)
	p.Value = 3
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
	cn.Execute "spCorrectImageSummary " & clng(request("txtProductID")) & "," &  clng(request("txtRootID")) & ",3"
	rs.Open "spGetImageSummary " & clng(request("txtProductID")) & "," &  clng(request("txtRootID")) & ",3",cn,adOpenStatic

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
	if blnPatch then
		strMethods = strMethods & ", Patch"
	end if
	if blnMinImgRec then
		strMethods = strMethods & ", MIR"
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

	
	if len(strMethods) > 0 then
		strMethods = mid(strMethods,3)
	end if
%>
<%if blnSuccess then%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="<%=strMethods & strNewImageSummary%>">
<%else%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="0">
<%end if%>
</BODY>
</HTML>
