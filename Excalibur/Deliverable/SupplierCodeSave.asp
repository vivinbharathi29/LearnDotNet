<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
    if (typeof (txtSuccess) != "undefined") {
        if (txtSuccess.value == "1") {
            if (IsFromPulsarPlus()) {               
                    ClosePulsarPlusPopup();
                    window.parent.parent.parent.popupCallBack(1);
            }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.SupplierCodeResults();
                    parent.window.parent.modalDialog.cancel();
                } else {
                    window.returnValue = 1;
                    window.parent.close();
                }
            }
        } else {
            document.write("<BR><font size=2 face=verdana>Unable to update supplier code.</font>");
        }
    } else {
        document.write("<BR><font size=2 face=verdana>Unable to update supplier code.</font>");
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">

<%
	dim RowsChanged
	dim strSuccess
	dim cn
	dim cm
	
	strSuccess = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
		
	cn.BeginTrans

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spAddSupplierCode"
		
	Set p = cm.CreateParameter("@CategoryID", 3, &H0001)
	p.Value = clng(request("txtCategoryID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@VendorID", 3, &H0001)
	p.Value = clng(request("txtVendorID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@ID", 200, &H0001, 50)
	p.Value = left(request("txtCode"),50)
	cm.Parameters.Append p
	

	cm.Execute RowsChanged
	Set cm=nothing
	
	if RowsChanged <> 1 or cn.Errors.count > 0 then
		cn.RollbackTrans
	else
		cn.CommitTrans
		strSuccess = "1"
	end if
		
	cn.Close
	set cn = nothing
		

%>


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">


</BODY>
</HTML>
