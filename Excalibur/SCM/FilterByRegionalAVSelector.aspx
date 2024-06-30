<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_FilterByRegionalAVSelector" Codebehind="FilterByRegionalAVSelector.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #form1
        {
            height: 500px;
        }
    </style>
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 420px; height: 541px; margin-left: 0px;">
        <asp:Label ID="lblHeader" runat="server" Text="Please Select Region and Product Brands To Display"
            Style="font-size: small; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 25px; left: 19px; width: 399px;"></asp:Label>
        <asp:ListBox ID="lbProductBands" runat="server" 
            DataTextField="ShortProdName" DataValueField="ProdID And Brand"
            SelectionMode="Multiple" Style="position: absolute; left: 26px; height: 382px;
            width: 394px; top: 118px; right: 657px;"></asp:ListBox>
        <br />
        <br />
        <div style="position: absolute; font-weight: bold; top: 52px; left: 44px;">Filter By 
            Product Type:</div>
        <br />
        <br />
        <asp:RadioButtonList ID="rblProductType" runat="server" 
            style="position: absolute; top: 50px; left: 203px;" 
            RepeatDirection="Horizontal" AutoPostBack="true">
            <asp:ListItem Value="1" Selected="True">Commercial</asp:ListItem>
            <asp:ListItem Value="2">Consumer</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <!--<asp:RadioButton ID="rbCons" 
            Style="position: absolute; top: 58px; left: 277px;" runat="server" Text="Consumer" AutoPostBack="true"/>
        <asp:RadioButton ID="rbAll" 
            style="position: absolute; top: 58px; left: 365px;" runat="server" Text="All" AutoPostBack="true" />
        <br />
        <asp:RadioButton ID="rbComm" 
            Style="position: absolute; top: 58px; left: 179px;" runat="server" Text="Commercial" AutoPostBack="true" /> -->
        <asp:RadioButtonList ID="rblProductStatus" runat="server" 
            style="position: absolute; top: 73px; left: 203px; width: 514px; height: 8px; z-index: 1;" 
            RepeatDirection="Horizontal" AutoPostBack="true">
            <asp:ListItem Value="5" Selected="True">All Active</asp:ListItem>
            <asp:ListItem Value="1">Definition</asp:ListItem>
            <asp:ListItem Value="2">Development</asp:ListItem>
            <asp:ListItem Value="3">Production</asp:ListItem>
            <asp:ListItem Value="4">Post-Production</asp:ListItem>
        </asp:RadioButtonList>
        <br />
        <div style="position: absolute; font-weight: bold; top: 100px; height: 12px; width: 170px; left: 30px;">
            Select a Product Brand:  <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <div style="font-weight: bold; position: absolute; top: 414px; left: 356px; width: 194px;">Select Your Region to Work With:</div>
                    <asp:RadioButtonList ID="rblSelectRegion" 
            style="position: absolute; top: 406px; left: 568px; font-weight: normal;" runat="server" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Value="1">Americas</asp:ListItem>
                        <asp:ListItem Value="2">EMEA</asp:ListItem>
                        <asp:ListItem Value="3">APJ</asp:ListItem>
                    </asp:RadioButtonList>
        <br />
        <hr style="position:absolute; top: 440px; left: -20px; width: 776px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 677px; width: 61px; height: 24px; top: 464px;" />
        <asp:Button ID="btnTest" runat="server" Text="Test" Style="position: absolute; left: 421px;
            width: 44px; height: 24px; top: -59px;" Visible="False" />
    </div>
<br />
<br />
<br />
        <asp:ListBox ID="lbCategories" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 450px; height: 382px;
            width: 327px; top: 118px"></asp:ListBox>
            <asp:Label ID="TestLabel" 
                style="position: absolute; top: 563px; left: 18px; height: 28px; width: 604px;" 
                runat="server" Text="Test" Visible="False" Font-Bold="True" 
                Font-Size="X-Large" ForeColor="Red"></asp:Label>
        <div style="position: absolute; font-weight: bold; top: 100px; height: 20px; width: 278px; left: 450px;">
            Please Select Category(s) To Display:</div>
    <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        </div>
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 647px;
            width: 35px; height: 24px; top: 564px; right: 196px;" />
    </form>
        <div style="position: absolute; font-weight: bold; top: 77px; left: 101px; margin-bottom: 4px;">
            Product Phase:</div>
        </body>
</html>
