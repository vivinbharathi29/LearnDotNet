<%@ Language="VBScript" %>
<%
  Dim AppRoot : AppRoot = Session("ApplicationRoot")
%>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript" />
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />

<link href="<%=AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
<script src="<%=AppRoot %>/includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
<script src="<%=AppRoot %>/includes/client/jquery-ui-1.10.4.min.js" type="text/javascript"></script>
    <script src="../Scripts/PulsarPlus.js"></script>
<script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--
function PreparePreview(){
    var strOutput="";
    var strType = "";
    var strAttachments = "";

    if (window.parent.frames["MainWindow"].frmMain.Radio1.checked)
        strType = "Ask a Question";
    else if (window.parent.frames["MainWindow"].frmMain.Radio2.checked)
        strType = "Report an Issue";
    else if (window.parent.frames["MainWindow"].frmMain.Radio3.checked)
        strType = "Make a Suggestion";
    else if (window.parent.frames["MainWindow"].frmMain.Radio4.checked)
        strType = "Request Admin Updates";

    if (window.parent.frames["MainWindow"].UploadPath1.innerText != "")
        strAttachments = strAttachments + "," + window.parent.frames["MainWindow"].UploadPath1.innerText;
    if (window.parent.frames["MainWindow"].UploadPath2.innerText != "")
        strAttachments = strAttachments + "," + window.parent.frames["MainWindow"].UploadPath2.innerText;
    if (window.parent.frames["MainWindow"].UploadPath3.innerText != "")
        strAttachments = strAttachments + "," + window.parent.frames["MainWindow"].UploadPath3.innerText;

    if (strAttachments !="")
        strAttachments = strAttachments.substr(1);

    strOutput = strOutput + window.parent.frames["MainWindow"].frmMain.txtSubject.value + "<BR><BR>";
    if (window.parent.frames["MainWindow"].frmMain.txtDetails.value != "")
        strOutput = strOutput + window.parent.frames["MainWindow"].frmMain.txtDetails.value + "<BR><BR>";
    strOutput = strOutput + "PROJECT: " + window.parent.frames["MainWindow"].frmMain.cboProject.options[window.parent.frames["MainWindow"].frmMain.cboProject.selectedIndex].text + "<BR>";
    strOutput = strOutput + "FEATURE/ISSUE: " + window.parent.frames["MainWindow"].frmMain.cboCategory.options[window.parent.frames["MainWindow"].frmMain.cboCategory.selectedIndex].text + "<BR>";
    strOutput = strOutput + "REQUEST TYPE: " + strType + "<BR><BR>";

    if (strAttachments != "")
        strOutput = strOutput + "ATTACHMENTS: " + strAttachments + "<BR><BR>";

    if (window.parent.frames["MainWindow"].frmMain.txtRequired.value != "")
        strOutput = strOutput + window.parent.frames["MainWindow"].frmMain.txtRequired.value.replace(/\r/gi, "<BR>") + "<BR><BR>";

    return strOutput;
}

function cmdNext_onclick() {

    if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value =="1")
        {
        if (window.parent.frames["MainWindow"].frmMain.txtSubject.value=="")
            {
            alert("Question/Request field is required.");
            window.parent.frames["MainWindow"].frmMain.txtSubject.focus();
            return;
            }
        if (window.parent.frames["MainWindow"].frmMain.cboProject.selectedIndex==0)
            {
            alert("Project field is required.");
            window.parent.frames["MainWindow"].frmMain.cboProject.focus();
            return;
            }
        if (window.parent.frames["MainWindow"].frmMain.cboCategory.selectedIndex==0)
            {
            alert("Feature/Issue field is required.");
            window.parent.frames["MainWindow"].frmMain.cboCategory.focus();
            return;
            }
        }
        
      if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value =="1" && (window.parent.frames["MainWindow"].frmMain.tagSearchedOn.value!=window.parent.frames["MainWindow"].frmMain.txtSubject.value) ||(window.parent.frames["MainWindow"].frmMain.CategoryChanged.value=="1") )
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = parseInt(window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value) + 1;
        window.parent.frames["MainWindow"].frmMain.submit();
        //$("#cmdPrevious").disabled=false;
        $("#cmdPrevious").removeAttr("disabled");
        }
    else if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="1" && window.parent.frames["MainWindow"].frmMain.txtResultsFound.value == "1")
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "2";
        window.parent.frames["MainWindow"].Tab1.style.display="none";
        window.parent.frames["MainWindow"].Tab2.style.display="";
        //$("#cmdPrevious").disabled=false;
        $("#cmdPrevious").removeAttr("disabled");
        }
    else if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="2" || (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="1" && window.parent.frames["MainWindow"].frmMain.txtResultsFound.value != "1"))
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "3";
        window.parent.frames["MainWindow"].Tab1.style.display="none";
        window.parent.frames["MainWindow"].Tab2.style.display="none";
        window.parent.frames["MainWindow"].Tab3.style.display="";

        window.parent.frames["MainWindow"].frmMain.txtDetails.focus();
        //$("#cmdPrevious").disabled=false;        
        $("#cmdPrevious").removeAttr("disabled");
        }
    else if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="3")
        {
        if (window.parent.frames["MainWindow"].frmMain.txtRequiredTemplate.value!=""  && (window.parent.frames["MainWindow"].frmMain.txtRequiredTemplate.value==window.parent.frames["MainWindow"].frmMain.txtRequired.value || window.parent.frames["MainWindow"].frmMain.txtRequired.value==""))
            {
            alert("You must supply the requested information in the Required Info field.");
            window.parent.frames["MainWindow"].frmMain.txtRequired.focus();
            return;
            }

        window.parent.frames["MainWindow"].EmailText.innerHTML = PreparePreview();
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "4";
        window.parent.frames["MainWindow"].Tab3.style.display="none";
        window.parent.frames["MainWindow"].Tab4.style.display="";
        //  $("#cmdFinish").attr("enabled", "enabled");     
        $("#cmdFinish").removeAttr("disabled");   
        $("#cmdNext").attr("disabled", "disabled");                
        }
}

function cmdPrevious_onclick(){
    if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value =="4")
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "3";
        window.parent.frames["MainWindow"].Tab4.style.display="none";
        window.parent.frames["MainWindow"].Tab3.style.display="";
        cmdFinish.disabled=true;
        cmdNext.disabled=false;
        }
    else if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="3" && window.parent.frames["MainWindow"].frmMain.txtResultsFound.value == "1")
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "2";
        window.parent.frames["MainWindow"].Tab3.style.display="none";
        window.parent.frames["MainWindow"].Tab2.style.display="";
        }
    else if (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="2" || (window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value=="3" && window.parent.frames["MainWindow"].frmMain.txtResultsFound.value != "1"))
        {
        window.parent.frames["MainWindow"].frmMain.txtCurrentStep.value = "1";
        window.parent.frames["MainWindow"].Tab3.style.display="none";
        window.parent.frames["MainWindow"].Tab2.style.display="none";
        window.parent.frames["MainWindow"].Tab1.style.display="";
        window.parent.frames["MainWindow"].frmMain.txtSubject.focus();
        cmdPrevious.disabled=true;
        }

}

function cmdFinish_onclick(){
    cmdCancel.disabled=true;
    cmdPrevious.disabled=true;
    cmdFinish.disabled=true;
    window.parent.frames["MainWindow"].frmMain.txtProjectName.value = window.parent.frames["MainWindow"].frmMain.cboProject.options[window.parent.frames["MainWindow"].frmMain.cboProject.selectedIndex].text;
    window.parent.frames["MainWindow"].frmMain.txtCategoryName.value = window.parent.frames["MainWindow"].frmMain.cboCategory.options[window.parent.frames["MainWindow"].frmMain.cboCategory.selectedIndex].text;
    window.parent.frames["MainWindow"].frmMain.action="SupportSave.asp";
    window.parent.frames["MainWindow"].frmMain.submit();
}

function cmdCancel_onclick() {
    cmdCancel.disabled = true;
    cmdPrevious.disabled = true;
    cmdFinish.disabled = true;
    if (window.parent.frames["MainWindow"].frmMain.txtSubject.value != "") {
        window.parent.frames["MainWindow"].frmCancel.txtCancelSummary.value = window.parent.frames["MainWindow"].frmMain.txtSubject.value;
        window.parent.frames["MainWindow"].frmCancel.txtCancelProject.value = window.parent.frames["MainWindow"].frmMain.cboProject.options[window.parent.frames["MainWindow"].frmMain.cboProject.selectedIndex].text;
        window.parent.frames["MainWindow"].frmCancel.txtCancelCategory.value = window.parent.frames["MainWindow"].frmMain.cboCategory.options[window.parent.frames["MainWindow"].frmMain.cboCategory.selectedIndex].text;
        window.parent.frames["MainWindow"].frmCancel.submit();
    } else {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        } else {
            if (CheckOpener() === false) {
                if (typeof parent.window.parent.ClosePropertiesDialog !== "undefined") {
                    parent.window.parent.ClosePropertiesDialog();
                } else if (typeof parent.window.parent.modalDialog.cancel !== "undefined") {
                    parent.window.parent.modalDialog.cancel(false);
                }
            } else {
                window.parent.close();
            }
        }
    }
}

function CheckOpener() {
    //If True, page opened with showModalDialog
    //if False, page opened with JQuery Modal Dialog
    var oWindow = window.dialogArguments;
    return (oWindow == null) ? false : true;
}
//-->
</script>
</head>
<body bgcolor="ivory">
<table border="0" cellspacing="1" cellpadding="1" align="right">
	<tr>
		<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" onclick="return cmdCancel_onclick()"></td>
		<td><input type="button" value="Previous" disabled id="cmdPrevious" name="cmdPrevious" onclick="return cmdPrevious_onclick()"></td>
		<td><input type="button" value="Next" id="cmdNext" name="cmdNext" onclick="return cmdNext_onclick()"></td>
		<td><input type="button" value="Finish" disabled id="cmdFinish" name="cmdFinish" onclick="return cmdFinish_onclick()"></td>
	</tr>
</table>
</body>
</html>