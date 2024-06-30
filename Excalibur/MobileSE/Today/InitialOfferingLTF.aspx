<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.InitialOfferingLTF" EnableEventValidation="false" ValidateRequest="false" Codebehind="InitialOfferingLTF.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
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
            font-family: Verdana; text-align: left; position: absolute; top: 12px; left: 10px;
            width: 274px;"></asp:Label>
        <asp:Label ID="lblLTFAV" runat="server" Text="LTF AV Number:" Style="font-size: x-small;
            font-weight: normal; font-family: Verdana; text-align: left; position: absolute;
            top: 52px; left: 10px; width: 140px;"></asp:Label>
        <asp:Label ID="lblLTFSA" runat="server" Text="LTF SA Number:" Style="font-size: x-small;
            font-weight: normal; font-family: Verdana; text-align: left; position: absolute;
            top: 83px; left: 10px; width: 143px;"></asp:Label>
        <asp:TextBox ID="txtLTFAV" runat="server" Style="position: absolute; top: 48px; left: 121px;"></asp:TextBox>
        <asp:TextBox ID="txtLTFSA" runat="server" Style="position: absolute; top: 78px; left: 121px;"></asp:TextBox>
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
            width: 35px; height: 24px" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 220px; width: 61px; height: 24px" OnClientClick="cmdCancel_onclick()" />
    </div>
    </form>
</body>
</html>
