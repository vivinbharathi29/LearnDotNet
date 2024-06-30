<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.MobileSE_Today_AddAvNoToFeature" EnableEventValidation="false" ValidateRequest="false" Codebehind="AddAvNoToFeature.aspx.vb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Add AvNo To Feature</title>
    <link href="../../style/general.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
    <script type="text/javascript" language="javascript">
        function cmdCancel_onclick() {
            //window.parent.close();
            if (IsFromPulsarPlus()) {
                ClosePulsarPlusPopup();
            }
            else {
                parent.window.parent.CloseExistingAVNoToFeature();
            }
        }
    </script>
</head>
    
<body runat="server" id="thisBody">
    <form id="frmAddAvNoToFeature" runat="server">
        <div style="width: 392px; height: 140px;">
            <asp:Label ID="lblDelName" runat="server" Style="font-size: x-small; font-weight: bold;font-family: Verdana; text-align: center; position: absolute; top: 12px; left: 10px;width: 330px;"></asp:Label>
            <asp:Label ID="lblAV" runat="server" Text="AV No:" Style="font-size: x-small; font-weight: normal;font-family: Verdana; text-align: left; position: absolute; top: 76px; left: 16px;width: 140px; right: 1421px;"></asp:Label>
            <asp:Label runat="server" ID="lblNoData" Visible="false" ForeColor="Red" Style="position: absolute; top: 73px; left: 24px;" ></asp:Label>
            <asp:DropDownList runat="server" ID="cboAVDescription"  Style="position: absolute; top: 73px; left: 58px; width:330px; " ></asp:DropDownList>
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
            <hr />
            <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 179px;
                width: 35px; height: 24px; top: 128px;" />
            <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
                left: 220px; width: 61px; height: 24px; top: 128px; bottom: 661px;" OnClientClick="cmdCancel_onclick()" />
        </div>
    </form>
</body>
</html>