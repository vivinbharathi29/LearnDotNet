<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.InitialOfferingFilter2" Codebehind="InitialOfferingFilter2.aspx.vb" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxControlToolkit" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script type="text/javascript" src="Scripts/jquery-1.10.2.js"></script>
<script type="text/javascript" src="Scripts/Pulsar2.js"> </script>

<script type="text/javascript">
window.onload = function () {
     if (document.getElementById("preferredLayout").value == "" || document.getElementById("preferredLayout").value == "pulsar2") {
         document.getElementById("btnCancel").style.display = "none";
    }
}

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
    <div style="width: 396px; height: 276px;">
        <asp:Label ID="lblHeader1" runat="server" Text="Please Select..." Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 8px; left: 10px; height: 19px; width: 395px;"></asp:Label>
        <asp:Label ID="lblBusUnit" runat="server" Text="Business Unit" Style="font-size: x-small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 96px; left: 19px; width: 173px;"></asp:Label>
        <asp:Label ID="lblProductProgram" runat="server" Text="Product Group" Style="font-size: x-small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 96px; left: 19px; width: 173px;"></asp:Label>
        <asp:Label ID="lblReportType" runat="server" Text="Report Type:" Style="font-size: x-small;
            font-weight: bold; font-family: Verdana; text-align: left; position: absolute;
            top: 50px; left: 32px; width: 88px; right: 1247px;"></asp:Label>
        <asp:Label ID="lblCategory" runat="server" Text="Category" Style="font-size: x-small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 96px; left: 224px; width: 173px;"></asp:Label>
        <br />
        <asp:DropDownList ID="ddlProductProgram" runat="server" Style="position: absolute;
            left: 22px; height: 22px; width: 175px; top: 114px;">
        </asp:DropDownList>
        <asp:DropDownList ID="ddlCategory2" runat="server" Style="position: absolute; left: 223px;
            height: 22px; width: 175px; top: 114px;">
        </asp:DropDownList>
        <asp:DropDownList ID="ddlCategory" runat="server" Style="position: absolute; left: 223px;
            height: 22px; width: 175px; top: 114px;">
        </asp:DropDownList>
        <asp:DropDownList ID="ddlBusUnit" runat="server" Style="position: absolute; left: 22px;
            height: 22px; width: 175px; top: 114px;">
            <asp:ListItem Value="0" Text=""></asp:ListItem>
            <asp:ListItem Value="1">Commercial</asp:ListItem>
            <asp:ListItem Value="2">Consumer</asp:ListItem>
        </asp:DropDownList>
        <div style="position: absolute; top: 247px; left: 10px; width: 398px;">
            <hr />
        </div>
        <asp:RadioButtonList ID="rblReportType" runat="server" Style="position: absolute;
            top: 42px; left: 118px; width: 279px; height: 26px; margin-bottom: 1px;" RepeatDirection="Horizontal"
            AutoPostBack="true">
            <asp:ListItem Value="0" Selected="True">Initial Offering</asp:ListItem>
            <asp:ListItem Value="1">Commodity Guidance</asp:ListItem>
        </asp:RadioButtonList>
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 297px;
            width: 35px; height: 24px; right: 1035px; top: 263px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 342px; width: 61px; height: 24px; top: 263px;" OnClientClick="cmdCancel_onclick()" />
        <input type="hidden" id="preferredLayout" runat="server" value="" />
    </div>
    </form>
</body>
</html>