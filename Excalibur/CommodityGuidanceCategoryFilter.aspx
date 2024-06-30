<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.CommodityGuidanceCategoryFilter" Codebehind="CommodityGuidanceCategoryFilter.aspx.vb" %>

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
<body runat="server" id="Body1">
    <form id="form2" runat="server">
    <div style="width: 352px; height: 466px;">
        <asp:Label ID="lblHeader1" runat="server" 
            Text="Please Select AV Feature Category(s)" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 8px; left: 10px; height: 19px; width: 340px;"></asp:Label>
        <asp:Label ID="lblBusiness" runat="server" Text="Business:" Style="font-size: x-small;
            font-weight: bold; font-family: Verdana; text-align: left; position: absolute;
            top: 47px; left: 42px; width: 64px; right: 1261px;"></asp:Label>
        <br />
        <div style="position: absolute; top: 441px; left: 10px; width: 348px;">
            <hr />
        </div>
        <asp:RadioButtonList ID="rblBusiness" runat="server" Style="position: absolute; top: 39px;
            left: 122px; width: 230px; height: 26px; margin-bottom: 1px;" RepeatDirection="Horizontal"
            AutoPostBack="true">
            <asp:ListItem Value="1" Selected="True">Commercial</asp:ListItem>
            <asp:ListItem Value="2">Consumer</asp:ListItem>
        </asp:RadioButtonList>
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 254px;
            width: 35px; height: 24px; top: 459px;" />
        <asp:Button ID="Button2" runat="server" Text="Cancel" Style="position: absolute;
            left: 294px; width: 61px; height: 24px; top: 459px;" 
            OnClientClick="cmdCancel_onclick()" />
        <asp:ListBox ID="lbCommercial" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 19px; height: 359px;
            width: 330px; top: 74px"></asp:ListBox>
        <asp:ListBox ID="lbConsumer" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 19px; height: 359px;
            width: 330px; top: 74px"></asp:ListBox>
    </div>
    </form>
</body>
</html>
