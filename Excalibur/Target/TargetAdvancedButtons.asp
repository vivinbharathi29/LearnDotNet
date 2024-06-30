<%@ Language=VBScript%>
<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
%>

<HTML>
<HEAD>
<META name="VI60_DefaultClientScript" content=JavaScript>
<script src="../Scripts/jquery-1.10.2.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdOK_onclick() {

	var ValidationFailed = false;
	var count =0;
	var pos;
	var strMessage = "";
/*	 if (window.parent.frames["MainWindow"].AddHFCN.txtTitle.value == "")
		{
		ValidationFailed = true;
		window.parent.frames["MainWindow"].AddHFCN.txtTitle.focus();
		window.alert("Title is required.");
		}
*/
	if (window.parent.frames["MainWindow"].frmTarget.IsPulsarProduct.value == 1) {
	    //check if atleast one release has been checked for the targeted versions
	    var tblVersion = window.parent.frames["MainWindow"].document.getElementById("tblTarget");
        var totalActiveRows = window.parent.frames["MainWindow"].document.getElementById("inpCount").value;
        var totalInActiveRows = window.parent.frames["MainWindow"].document.getElementById("inactiveCount").value;
        totalActiveRows = Number(totalInActiveRows) + Number(totalActiveRows);
	    for (var i = 0; i <= totalActiveRows ; i++) {
	        if (((window.parent.frames["MainWindow"].document.getElementById("select_" + i)) != null) && ((window.parent.frames["MainWindow"].document.getElementById("select_" + i)) != "undefined") && ((window.parent.frames["MainWindow"].document.getElementById("select_" + i).value) == "Targeted") &&
                   ((window.parent.frames["MainWindow"].document.getElementById("TargetedReleases_" + i)) != null) && ((window.parent.frames["MainWindow"].document.getElementById("TargetedReleases_" + i)) != "undefined") && ((window.parent.frames["MainWindow"].document.getElementById("TargetedReleases_" + i).value == "")))
	        {
	            strMessage = strMessage == "" ? window.parent.frames["MainWindow"].document.getElementById("aVersionID_" + i).innerHTML : strMessage + "," + window.parent.frames["MainWindow"].document.getElementById("aVersionID_" + i).innerHTML;
	        }
	    }
	   
	    if (strMessage.length > 0)
	    {
	        alert("Select at least 1 release to target for version listed: " + strMessage);
	        ValidationFailed = true;
	    }
	}

	if (typeof(window.parent.frames["MainWindow"].frmTarget) != "undefined"){
		if (! ValidationFailed){
			cmdOK.disabled = true;
			cmdCancel.disabled = true;	
			window.parent.frames["MainWindow"].frmTarget.submit();
		}
	}else{
	    if (parent.window.parent.document.getElementById('modal_dialog')) {
	        parent.window.parent.modalDialog.cancel();
	    } else {
	        if (IsFromPulsarPlus()) {
	            ClosePulsarPlusPopup();
	        }
	        else {
	            window.parent.close();
	        }
	    }
	}
}



function cmdCancel_onclick(pulsarplusDivId) {
    if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
        // For Closing current popup if Called from pulsarplus
        parent.window.parent.closeExternalPopup();
    }
    else if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        //if (window.confirm ("Are you sure you want to exit this screen without releasing this deliverable?") == true)
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }
}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<TABLE  BORDER=0 CELLSPACING=1 CELLPADDING=1 align=right>
	<TR>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></TD>
	</TR>
</TABLE>

</BODY>
</HTML>