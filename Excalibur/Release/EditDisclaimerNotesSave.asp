<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    function window_onload() {
        if (typeof (txtSuccess) != "undefined") {
            if (txtSuccess.value == "1") {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
            }
            else {
                //document.write("<BR><font size=2 face=verdana>Unable to update program.</font>");
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel();
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
            }
        }
        else {
            //document.write("<BR><font size=2 face=verdana>Testing</font>");
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.returnValue = 1;
                window.parent.close();
            }
        }
    }
</SCRIPT>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%   
    dim cn
	dim cm
	dim strSuccess
	dim p
    
    dim ID
    ID = 0
    if Request.Form("txtID") <> "" then
        ID = Request.Form("txtID")
    end if

    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	cn.BeginTrans
	strSuccess = "1"
		
	set cm = server.CreateObject("ADODB.Command")
	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
		
	cm.CommandText = "usp_ProductVersion_UpdateDisclaimerNotes"	

    Set p = cm.CreateParameter("@ID", 3,  &H0001)
    p.Value = clng(ID)
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@ReleaseID", 3,  &H0001)
	p.Value = clng(Request.Form("txtReleaseID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@DisclaimerNotes", 200,  &H0001, 4000)
	p.Value = left(Request.Form("txtDisclaimerNotes"),4000)
	cm.Parameters.Append p
    
	Set p = cm.CreateParameter("@State", 3,  &H0001)
	p.Value = clng(Request.Form("optState"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@User", 200,  &H0001, 120)
	p.Value = left(Request.Form("txtUser"),120)
	cm.Parameters.Append p
					
	cm.Execute rowschanged
	set cm=nothing

	if cn.Errors.count > 0 then
		strSuccess = "0"
	end if	
		
	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

	cn.Close
	set cn = nothing
%>
<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">
</BODY>
</HTML>
