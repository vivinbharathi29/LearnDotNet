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
	        //save array value and return to parent page: ---
	        parent.window.parent.modalDialog.passArgument(txtSuccess.value, 'brand_update_result');
	        parent.window.parent.ChoosenewBrandResult();
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
    <input id="txtSuccess" name="txtSuccess" type="hidden" value="<%=trim(request("cboNew"))%>">
</BODY>
</HTML>
