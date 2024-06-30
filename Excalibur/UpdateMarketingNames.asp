<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!-- #include file = "includes/Security.asp" -->
<!-- #include file="includes/DataWrapper.asp" -->
<!-- #include file="includes/no-cache.asp" -->
<!-- #include file="includes/lib_debug.inc" -->
<%

Dim regEx
Set regEx = New RegExp
regEx.Global = True

Dim AppRoot : AppRoot = Session("ApplicationRoot")

Dim rs, dw, cn, cmd

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Dim sPBID : sPBID = Request.QueryString("PBID")
Dim sBID : sBID = Request.QueryString("BID")
Dim sName : sName = Request.QueryString("Name")
Dim sTypeID : sTypeID = Request.QueryString("Type")
Dim sSeries : sSeries = Request.QueryString("Series")
Dim sPlatFormID : sPlatFormID = Request.QueryString("PlatFormID")

Dim sNameType
If sTypeID = 5 Then
   sNameType = "CTO Model Number"
Elseif sTypeID = 6 Then
   sNameType = "Short Name"
Elseif sTypeID = 7 Then
   sNameType = "HP Brand Name <br />(Service Tag up)"
Elseif sTypeID = 8 Then
   sNameType = "Model Number<br />(Service Tag down)"
Elseif sTypeID = 9 Then
   sNameType = "BIOS Branding"
End If
'PER EFREN'S REQUEST - DO NOT REMOVE
'Elseif sTypeID = 4 Then
'   sNameType = "BTO Service Tag Name"
sName = Request.QueryString("Name")

%>
<html>
<head>
<title></title>
<link rel="stylesheet" type="text/css" href="style/excalibur.css" />
<link href="style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
<script src="includes/client/jquery.min.js" type="text/javascript"></script>
<script src="includes/client/jquery-ui.min.js" type="text/javascript"></script>
<script type="text/javascript">

$(function () {
    $("input:button").button();
});

function Body_OnLoad()
{
    switch(frmMain.hdnTypeID.value) {
        case "5": //CTO Model Number
            $("#txtName").attr('maxlength', '50');
            break;
        case "6": //Short Name
            $("#txtName").attr('maxlength', '100');
            break;
        case "7": //HP Brand Name (Service Tag up) //Service Tag 
            $("#txtName").attr('maxlength', '50');
            break;
        case "8": //Model Number (Service Tag down) //Master Label 
            $("#txtName").attr('maxlength', '50');
            break;
        case "9": //BIOS Branding
            $("#txtName").attr('maxlength', '50');
            break;
        //PER EFREN'S REQUEST - DO NOT REMOVE
        //case "4": //BTO Service Tag Name
        //    $("#txtName").attr('maxlength', '30');
        //    break;
    } 
    
	if (frmMain.txtName.value == ""){
	    EditName();
	}
    else{
	    divName.style.display = "none";
	    divNameText.style.display = "";
    }
}

function EditName()
{
    divName.style.display = "";
    divNameText.style.display = "none";
	EditLink.style.display = "none";
}
function cmdOK_onclick() {
    //if (frmMain.txtName.value == "") {
    //    window.parent.CloseMarketingNameDialog();
    //    return;
    //}
    if (frmMain.txtName.value != "") {
        var regexp = /^[a-z\d\-_\s]+$/i;
        if (frmMain.txtName.value.search(regexp) == -1) {
            alert('Only Alpha Numeric Characters and Dashes Are Allowed');
            return;
        }
    }

    var parameters = "function=UpdateMarketingName&PBID=" + frmMain.hdnPBID.value + "&Name=" + encodeURIComponent(frmMain.txtName.value) + "&NameType=" + frmMain.hdnTypeID.value + "&BID=" + frmMain.hdnBID.value + "&Series=" + encodeURIComponent(frmMain.hdnSeries.value)+ "&PlatFormID=" + frmMain.hdnPlatFormID.value; 
    var request = null;
    //Initialize the AJAX variable.
    if (window.XMLHttpRequest) {        //Are we working with mozilla
        request = new XMLHttpRequest(); //Yes -- this is mozilla.
    } else { //Not Mozilla, must be IE
        request = new ActiveXObject("Microsoft.XMLHTTP");
    } //End setup Ajax.
    request.open("POST", "<%=AppRoot %>/UpdateMarketingNameSave.asp", false);
    request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    request.send(parameters);
    //alert(request.responseText);
    if (request.responseText == 'Success') {
        window.parent.CloseMarketingNameDialog();
        window.parent.location.reload(true);
    } else {
        alert("Update Failed");
    }
}
function cmdCancel_onclick() {
    window.parent.CloseMarketingNameDialog();
}
</script>
</head>
<body onload="Body_OnLoad()">
<form method="post" id="frmMain">
<table class="FormTable" style="background-color:cornsilk; width:100%; border-width:1px; border-spacing: 0px; padding:1px; border-color:tan;">
	<tr>
		<th style="width:30%; vertical-align:central">&nbsp;<%= sNameType%></th>
		<td><div id="divName" style="display:none">&nbsp;<input type="text" style="width:80%" id="txtName" name="txtName" value="<%= sName%>" /></div>
            <div id="divNameText"><table width="100%" border="0"><tr><td style="border:none"><%= sName%></td><td style="border:none; text-align:right"><a id="EditLink" href="javascript:EditName();">Edit</a></td></tr></table></div>
		</td>
	</tr>
</table>
<br /><br />
<div style="text-align:right; border-top:2px solid #b2b2b2"">
    <br />
    <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
    <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
</div>
<input type="hidden" id="hdnTypeID" name="hdnTypeID" value="<%=sTypeID%>">
<input type="hidden" id="hdnPBID" name="hdnPBID" value="<%=sPBID%>">
<input type="hidden" id="hdnBID" name="hdnBID" value="<%=sBID%>">
<input type="hidden" id="hdnSeries" name="hdnSeries" value="<%=sSeries%>">
    <input type="hidden" id="hdnPlatFormID" name="hdnPlatFormID" value="<%=sPlatFormID%>">
</form>
</body>
</html>
