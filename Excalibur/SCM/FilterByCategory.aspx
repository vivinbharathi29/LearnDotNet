<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_FilterByCategory" Codebehind="FilterByCategory.aspx.vb" %>

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
    <div style="width: 288px; height: 433px;">
        <asp:Label ID="lblHeader" runat="server" Text="Please Select Category(s) To Display"
            Style="font-size: small; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 15px; left: 10px; width: 288px;"></asp:Label>
        <asp:ListBox ID="lbCategories" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 22px; height: 344px;
            width: 261px; top: 49px"></asp:ListBox>
        <br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
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
        <br />
        <br />
        <br />
        <br />
        <hr />
        <asp:Button ID="btnDeselect" runat="server" Text="Show All" Style="position: absolute; left: 21px;
            width: 83px; height: 24px; top: 430px;" />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 189px;
            width: 35px; height: 24px; top: 430px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 230px; width: 61px; height: 24px; top: 430px;" 
            OnClientClick="cmdCancel_onclick()" />
    </div>
    </form>
</body>
</html>
