<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<%
  Dim AppRoot
  AppRoot = Session("ApplicationRoot")
%>

<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
<script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

    function cmdCancel_onclick(pulsarplusDivId) {
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (window.parent.frames["UpperWindow"]) {
                parent.window.parent.modalDialog.cancel(false);
            } else {
                window.parent.close();
            }
        }
}

function cmdOK_onclick() {
    $("#divProg").show();
    //window.parent.frames["UpperWindow"].inlineProgressBarDiv.style.display = '';
    cmdCancel.disabled = true;
    cmdOK.disabled = true;
    //window.parent.frames["UpperWindow"].frmChange.hidFunction.value = "save";
    //window.parent.frames["UpperWindow"].frmChange.submit();	
    var sUserName = $("#txtUserName", window.parent.frames["UpperWindow"].document).val();
    var BID = $("#txtBID", window.parent.frames["UpperWindow"].document).val();
    var PVID = $("#txtPVID", window.parent.frames["UpperWindow"].document).val();
    var Releases = $("#txtReleases", window.parent.frames["UpperWindow"].document).val();
    var RTPDate = $("#txtRTPDate", window.parent.frames["UpperWindow"].document).val();
    var EMDate = $("#txtEMDate", window.parent.frames["UpperWindow"].document).val();
    
    $('input:checkbox[name="Base"]:checked', window.parent.frames["UpperWindow"].document).each(function () {
        if (!$(this).attr("disabled")) {
            var BaseID = ($(this).attr("id")).split('-');
                      
            var FeatureID = BaseID[0];
            var AvDetailID = BaseID[1];

            var sBaseGPG = $("#txtGPGDescription" + FeatureID + '-' + AvDetailID, window.parent.frames["UpperWindow"].document).val();
                    
            var iParentID = AvDetailID;

            $('input:checkbox[name="chkRow' + FeatureID + '-' + AvDetailID + '"]:checked', window.parent.frames["UpperWindow"].document).each(function () {
                if (!$(this).attr("disabled")) {                 
                    
                    var sValue = ($(this).val()).split('|');
                    var AvParentID = iParentID;
                    var shareAV = 0;
                    if (sValue[5] == "True")
                        shareAV = 1;

                    $("#txtHint").text("creating localized av for " + sBaseGPG + " - " + sValue[2] + " ...");

                    var url = "<%=AppRoot %>/SupplyChain/CreateLocalizedAV.asp?PVID=" + PVID + "&BID=" + BID + "&BaseGPG=" + sBaseGPG + "&FeatureID=" + FeatureID + "&UserName=" + sUserName + "&ConfigCode=" + sValue[2] + "&CountryCode=" + sValue[1] + "&AVParentID=" + AvParentID + "&GeoID=" + sValue[4] + "&ShareAV=" + shareAV + "&Releases=" + Releases + "&RTPDate=" + RTPDate + "&EMDate=" + EMDate + "&pulsarplusDivId=SupplyChain";
                    
                    var xmlhttp = new XMLHttpRequest();
                    
                    xmlhttp.open("POST", url, false);
                    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
                    xmlhttp.send("PVID=" + PVID + "&BID=" + BID + "&BaseGPG=" + sBaseGPG + "&FeatureID=" + FeatureID + "&UserName=" + sUserName + "&ConfigCode=" + sValue[2] + "&CountryCode=" + encodeURIComponent(sValue[1]) + "&AVParentID=" + AvParentID + "&GeoID=" + sValue[4] + "&ShareAV=" + shareAV + "&pulsarplusDivId=SupplyChain");
                    iParentID = xmlhttp.responseText;
                    
                }
            });                   
        }
    });
    
    setTimeout(function () {
        var pulsarplusDivId = $("#pulsarplusDivId").val();
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
    else 
    {
        if (window.parent.frames["UpperWindow"]) {
            parent.window.parent.ReloadLocalizeAVs("YES");
        } else {

            window.parent.close();
    }
    }
    }, 1000);   
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">
    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>
            <td>
                <input type="button" value="OK" id="cmdOK" name="cmdOK" language="javascript" onclick="return cmdOK_onclick()"></td>
            <td>
                <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" language="javascript" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></td>
        </tr>
    </table>
    <div id="divProg" style="display:none">  
        <p><span id="txtHint"></span></p> 
    </div> 
    <input id="pulsarplusDivId" type="hidden" value="<%=Request("pulsarplusDivId")%>" />
</body>
</html>