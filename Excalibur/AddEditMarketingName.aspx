<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.AddEditMarketingName" Codebehind="AddEditMarketingName.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        if (window.parent.frames["UpperWindow"]) {
            //save value and return to parent page: ---
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
        }
    }    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 277px; height: 167px;">
        <asp:TextBox ID="txtNewName" Visible="false" runat="server" MaxLength="50" Style="font-size: small;
            font-family: Verdana; text-align: center; position: absolute; top: 130px; left: 10px;
            width: 278px;"></asp:TextBox>
        <asp:TextBox ID="TextBox1" runat="server" MaxLength="50" Style="font-size: small;
            font-family: Verdana; text-align: center; position: absolute; top: 45px; left: 10px;
            width: 278px;"></asp:TextBox>
        <asp:TextBox ID="TextBox2" runat="server" MaxLength="50" Style="font-size: small;
            font-family: Verdana; text-align: center; position: absolute; top: 130px; left: 10px;
            width: 278px;"></asp:TextBox>
        <asp:Label ID="lblNewNameHeader" runat="server" Text="New Name" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 102px; left: 10px; width: 288px;"></asp:Label>
        <asp:Label ID="lblOldName" runat="server" Style="font-size: small; font-family: Verdana;
            text-align: center; position: absolute; top: 48px; left: 10px; width: 288px;"></asp:Label>
        <asp:Label ID="lblOldNameHeader" runat="server" Text="Old Name" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 20px; left: 10px; width: 288px;"></asp:Label>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <hr />
        <asp:Button ID="btnAutoGen" runat="server" Text="Auto-generate" Style="position: absolute;
            left: 15px; width: 102px; height: 24px; top: 185px;" />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 189px;
            width: 35px; height: 24px; top: 185px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 230px; width: 61px; height: 24px; top: 185px;" OnClientClick="cmdCancel_onclick()" />
    </div>
    </form>
</body>
</html>
