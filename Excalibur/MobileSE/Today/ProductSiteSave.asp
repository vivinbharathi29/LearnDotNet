<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (txtSuccess.value!="0")
	{
	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        //save value and return to parent page: ---
	        parent.window.parent.SelectedSitesResult(txtSuccess.value);
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        window.returnValue = txtSuccess.value;
	        window.close();
	    }
	}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%

if trim(request("txtID")) = "0" then
    strSuccess = ""'request("txtSiteList") & "|" & replace(request("chkSite")," ","")
else
	dim strConnect
	dim cn
	dim cm
	dim rowseffected
    dim strSuccess
    
    strSuccess = request("txtSiteList")
	
'	strConnect = Session("PDPIMS_ConnectionString")
'	set cn = server.CreateObject("ADODB.Connection")
'	cn.ConnectionString = strConnect
'	cn.Open

 '   cn.begintrans

  ' 	set cm = server.CreateObject("ADODB.command")

'   	cm.ActiveConnection = cn
'   	cm.CommandText = "spUpdateRCTOSites"
'   	cm.CommandType =  &H0004

'   	Set p = cm.CreateParameter("@ProductID", 3, &H0001)
'   	p.Value = clng(request("txtID"))
'   	cm.Parameters.Append p

'   	Set p = cm.CreateParameter("@RCTOSites", 200, &H0001,50)
'   	p.Value = left(server.HTMLEncode(request("txtSiteList")),50)
'   	cm.Parameters.Append p

'   	cm.Execute RowsEffected
'   	if RowsEffected <> 1 then
'   	    strSuccess = "0"
'   	end if
'    set cm = nothing
    
    
'    if trim(strSuccess) = "0" then
'        cn.rollbacktrans
'    else
'        cn.committrans        
'    end if

'    cn.close
    
'    set cm = nothing
'	set cn = nothing

end if

%>
    <input id="txtSuccess" name="txtSuccess" type="text" value="<%=strSuccess%>">
</BODY>
</HTML>
