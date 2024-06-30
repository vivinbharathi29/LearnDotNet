<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.AddAvNoToDeliverable" EnableEventValidation="false" ValidateRequest="false" Codebehind="AddAvNoToDeliverable.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            window.parent.close();
        }
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 272px; height: 120px;">
        <asp:Label ID="lblDelName" runat="server" Style="font-size: x-small; font-weight: bold;
            font-family: Verdana; text-align: center; position: absolute; top: 12px; left: 10px;
            width: 274px;"></asp:Label>
        <asp:CheckBox ID="cbUpdateDesc" runat="server" Style="font-size: x-small; font-weight: normal;
            font-family: Verdana; text-align: left; position: absolute; top: 36px; left: 74px;
            width: 175px;" Text="Update AV Descriptions"/>
        <asp:Label ID="lblAV" runat="server" Text="AV No:" Style="font-size: x-small; font-weight: normal;
            font-family: Verdana; text-align: left; position: absolute; top: 76px; left: 46px;
            width: 140px; right: 1421px;"></asp:Label>
        <asp:TextBox ID="txtAV" runat="server" 
            Style="position: absolute; top: 73px; left: 98px;"></asp:TextBox>
        <br />
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
