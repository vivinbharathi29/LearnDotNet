<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.InitialOfferingFilter" Codebehind="InitialOfferingFilter.aspx.vb" %>

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
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 396px; height: 120px;">
        <asp:Label ID="lblHeader" runat="server" Text="Please Select..." Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 4px; left: 10px; height: 19px; width: 395px;"></asp:Label>
        <asp:Label ID="lblBusUnit" runat="server" Text="Business Unit" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 35px; left: 19px; width: 174px;"></asp:Label>
        <asp:Label ID="lblCategory" runat="server" Text="Category" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 35px; left: 224px; width: 173px;"></asp:Label>
        <br />
        <asp:DropDownList ID="ddlCategory" runat="server" Style="position: absolute; left: 223px;
            height: 22px; width: 175px; top: 62px;">
        </asp:DropDownList>
        <asp:DropDownList ID="ddlBusUnit" runat="server" Style="position: absolute; left: 18px;
            height: 22px; width: 175px; top: 62px;">
            <asp:ListItem Value="0" Text=""></asp:ListItem>
            <asp:ListItem Value="1">Commercial</asp:ListItem>
            <asp:ListItem Value="2">Consumer</asp:ListItem>
        </asp:DropDownList>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <hr />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 297px;
            width: 35px; height: 24px; right: 1035px; top: 118px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 342px; width: 61px; height: 24px; top: 118px;" OnClientClick="cmdCancel_onclick()" />
    </div>
    </form>
</body>
</html>
