<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Common_FileUpload" Codebehind="FileUpload.aspx.vb" %>
<%@ Register TagPrefix="Upload" Namespace="Brettle.Web.NeatUpload" Assembly="Brettle.Web.NeatUpload" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>File Upload</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
	.ProgressBar {
		margin: 0px;
		border: 0px;
		padding: 0px;
		width: 100%;
		height: 3em;
	}
	</style>
    <script type="text/javascript" src="../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript">
    <!--
    function body_onLoad() {
        if (form1.hidReturnValue.value != "") {
            if (form1.hidAppName.value == "PulsarPlus")
            {
                parent.window.parent.UploadZipReturn(form1.hidReturnValue.value, form1.hidControlId.value);
                parent.window.parent.CloseDialog(form1.hidReturnValue.value);
            }
            else {
                if (CheckOpener() === false && parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.UploadZip_return(form1.hidReturnValue.value);
                    parent.window.parent.modalDialog.cancel(false);
                } else {
                    window.returnValue = form1.hidReturnValue.value;
                    window.parent.close();
                }
            }            
        }
    }

    function cancelButton_onClick() {
        if (form1.hidAppName.value == "PulsarPlus") {            
            parent.window.parent.CloseDialog(form1.hidReturnValue.value);
        }
        else {
            if (CheckOpener() === false && parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel(false);
            } else {
                window.parent.close();
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
<body onload="body_onLoad()" style="background-color: #FFFFF0;">
    <form id="form1" runat="server" submitdisabledcontrols="true">
        <h3>
            <asp:Label ID="lblTitle" runat="server" Text="Label"></asp:Label>
        </h3>
            <span id="errorText" runat="server" style="color: Red; font-size: small" visible="false">
                <asp:Label ID="lblErrUploading" runat="server" Text="An Error Occured While Uploading Your File."></asp:Label>
            </span>
            
        <p><span id="instructions" class="Note">Maximum Upload Size is 500MB</span>
            <Upload:InputFile ID="InputFile1" runat="server" Width="95%" /><br />
            <asp:RegularExpressionValidator ID="RegularExpressionValidator1" ControlToValidate="InputFile1" Display="Static" EnableClientScript="True" runat="server" />
            <Upload:ProgressBar ID="ProgressBar1" runat="server" Inline="True" Triggers="submitButton" Width="95%" />
            </p>
        <p>
            <asp:Button ID="submitButton" runat="server" Text="Submit" />
            <input id="cancelButton" type="button" value="Cancel" onclick="cancelButton_onClick()" />
            <asp:HiddenField ID="hidReturnValue" runat="server" />
            <asp:HiddenField ID="hidAppName" runat="server" />
            <asp:HiddenField ID="hidControlId" runat="server" />
        </p>
    </form>
</body>
</html>
