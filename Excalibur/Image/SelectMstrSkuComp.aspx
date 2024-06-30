<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.Image_SelectMstrSkuComp" Codebehind="SelectMstrSkuComp.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <title></title>
</head>
<body id="body" runat="server">
    <form id="form1" runat="server">
    <div style="width: 300px; text-align: right;">
        <asp:DropDownList ID="ddlMstrSkuComp" runat="server" Width="100%">
        <asp:ListItem Value="" Text="-- Use Default Value --"></asp:ListItem>
        </asp:DropDownList>
        <br />
        <br />
        <asp:Button ID="btnSave" runat="server" Text="Save" Width="75px" />&nbsp;
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Width="75px" />
    </div>
    </form>
</body>
</html>
