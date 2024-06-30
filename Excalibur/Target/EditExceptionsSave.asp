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
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value != "0") {
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                }
                else if (IsFromPulsarPlus()) {
                    window.parent.parent.parent.popupCallBack(1);
                    ClosePulsarPlusPopup();
                }
                else {
                    if (parent.window.parent.document.getElementById('modal_dialog')) {
                        //save value and return to parent page: ---
                        parent.window.parent.ChangeTargetNotesResult(txtSuccess.value);
                        parent.window.parent.modalDialog.cancel();
                    } else {
                        window.returnValue = txtSuccess.value;
                        window.close();
                    }
                }
            }
            else
                document.write("Unable to update Exceptions.  An unexpected error occurred.");
        }
        else {
            document.write("Unable to update Exceptions.  An unexpected error occurred.");
        }

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
	
	
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")

    cm.ActiveConnection = cn
    cm.CommandText = "spUpdateExceptions"
    cm.CommandType = &H0004
       
	Set p = cm.CreateParameter("@ProductID",adInteger, &H0001)
	p.Value = request("txtProductID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VersionID",adInteger, &H0001)
	p.Value = request("txtVersionID")
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@TargetNotes",adVarChar, &H0001,255)
	p.Value = left(request("txtExceptions"),255)
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@Cycle",adBoolean, &H0001)
	if request("chkOOC") = "on" then
		p.Value = true
	else
		p.Value = false
	end if
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
	set rs=nothing
	set cn=nothing
	
%>
<%if blnSuccess then%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="<%=request("txtExceptions")%>">
<%else%>
<INPUT type="text" style="display:none" id=txtSuccess name=txtSuccess value="0">
<%end if%>
</BODY>
</HTML>
