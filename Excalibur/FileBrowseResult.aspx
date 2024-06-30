<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.FileBrowseResult" Codebehind="FileBrowseResult.aspx.vb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>File Browse Results</title>

    <script type="text/javascript">
<!--

function LoadFileWindow_onload() {
	window.location.replace("FileBrowseError.asp");
}

function window_onload() {
	window.parent.location.replace(form1.hfPath.value);
}

function window_onload2() {
	window.location.replace(form1.hfPath.value);
}

//-->
    </script>


<link href="style/wizard%20style.css" rel="stylesheet" type="text/css" />
</head>
<body id="body" runat="server">
    <form id="form1" runat="server">
        <div>
            <p><span style="display:none;font-weight:bold; color:Red">Note: Make sure Excalibur is in your Popup Blockers Trusted Sites List</span></p>
            <asp:Panel ID="pnlFileInfo" runat="server" Visible="false" Width="100%">
                <asp:Label ID="lblFileName" runat="server" Text="File Name: " Font-Bold="True"></asp:Label>
                <asp:Label ID="lblFileNameText" runat="server" Text=""></asp:Label>
                <br />
                <asp:Label ID="lblFileSize" runat="server" Text="File Size: " Font-Bold="True"></asp:Label>
                <asp:Label ID="lblFileSizeText" runat="server" Text=""></asp:Label>
                <br />
                <asp:Label ID="lblDownload" runat="server" Text="Download: " Font-Bold="True"></asp:Label>
                <asp:LinkButton ID="lbDownload" runat="server">Click Here to Download</asp:LinkButton>
                <asp:Label ID="lblPopupWarning" runat="server" Text="<BR><BR>Note: Make sure Excalibur is in your Popup Blockers Trusted Sites List"  font-size=X-Small Font-Bold="True" ForeColor=Red></asp:Label>
            </asp:Panel>
            <asp:Panel ID="pnlFileMissing" runat="server" Width="100%" Visible="false">
                <asp:Label ID="lblFileError" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                
            </asp:Panel>
            <iframe runat="server" height="800" id="LoadFileWindow" style="width: 100%"></iframe>
        </div>
        <asp:HiddenField ID="hfPath" runat="server" />
        <asp:HiddenField ID="hfZipPath" runat="server" />
    </form>
</body>
</html>
