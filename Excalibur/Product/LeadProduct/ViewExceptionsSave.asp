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
    <script src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value == "1")
		{
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.LeadProductSynchronizationCallback(txtSuccess.value);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        if (parent.window.parent.document.getElementById('modal_dialog')) {
		            parent.window.parent.modalDialog.cancel(true);
		        } else {
		            window.returnValue = txtSuccess.value;
		            window.close();
		        }
		    }
			}
		else if (txtSuccess.value == "2") {
		    if (IsFromPulsarPlus()) {
		        window.parent.parent.parent.LeadProductSynchronizationCallback(txtSuccess.value);
		        ClosePulsarPlusPopup();
		    }
		    else {
		        if (parent.window.parent.document.getElementById('modal_dialog')) {
		            parent.window.parent.modalDialog.cancel();
		        } else {
		            window.close();
		        }
		    }
		}
		}

}

//-->
</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="return window_onload()">


<%
	dim cn
	dim rs
	dim IDArray
	dim strID
	dim strSuccess
	dim blnFailed
	dim RowsUpdated
	
	strSuccess = 2

	if request("chkRemove") <> "" then

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") 
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")

		cn.BeginTrans
		blnFailed = false

		IDArray = split(request("chkRemove"),",")
		for each strItem in IDArray
			strType=ucase(left(trim(strItem),1))
			strID= split(mid(trim(strItem),2), ":")

			if strType = "R" then
				cn.Execute "spRemoveLeadProductRootException " & clng(strID[0]) & "," & clng(strID[1]), RowsUpdated
			else
				cn.Execute "spRemoveLeadProductVersionException " & clng(strID) & "," & clng(strID[1]),RowsUpdated
			end if
			if RowsUpdated <> 1 then
				blnFailed = true
				exit for
			end if	
		next		
		
		if blnFailed then
			cn.RollbackTrans
			strSuccess = 0
		else
			cn.CommitTrans
			strSuccess = 1
		end if


		set rs = nothing
		cn.Close
		set cn = nothing
	end if
%>

<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>
