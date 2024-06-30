<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_FilterByPlantsAVSelector" Codebehind="FilterByPlantsAVSelector.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Filter Regional AV Data by Plant</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #form1
        {
            height: 500px;
            width: 797px;
        }
    </style>
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 787px; height: 457px; margin-left: 0px;">
        <asp:Label ID="lblHeader" runat="server" Text="Plants and Categories Filter Page:"
            Style="font-size: medium; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 25px; left: 21px; width: 766px; height: 19px;"></asp:Label>
        <asp:ListBox ID="lbPlants" runat="server" 
            DataTextField="PlantName" DataValueField="RCTOPlantsID"
            SelectionMode="Multiple" Style="position: absolute; left: 30px; height: 382px;
            width: 394px; top: 75px; right: 653px;"></asp:ListBox>
        <br />
        <br />
        <br />
        <br />
        <br />
        <!--<asp:RadioButton ID="rbCons" 
            Style="position: absolute; top: 58px; left: 277px;" runat="server" Text="Consumer" AutoPostBack="true"/>
        <asp:RadioButton ID="rbAll" 
            style="position: absolute; top: 58px; left: 365px;" runat="server" Text="All" AutoPostBack="true" />
        <br />
        <asp:RadioButton ID="rbComm" 
            Style="position: absolute; top: 58px; left: 179px;" runat="server" Text="Commercial" AutoPostBack="true" /> -->
        <br />
        <div style="position: absolute; font-weight: bold; top: 54px; height: 16px; width: 202px; left: 30px; right: 845px;">
            Select a Plant to Filter by:  <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <hr style="position:absolute; top: 419px; left: -20px; width: 776px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 675px; width: 61px; height: 24px; top: 448px;" />
        <asp:Button ID="btnTest" runat="server" Text="Test" Style="position: absolute; left: 343px;
            width: 44px; height: 24px; top: -12px;" Visible="False" />
    </div>
<br />
<br />
<br />
        <asp:ListBox ID="lbCategories" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 458px; height: 382px;
            width: 327px; top: 76px"></asp:ListBox>
            <asp:Label ID="TestLabel" 
                style="position: absolute; top: 499px; left: 24px; height: 28px; width: 604px;" 
                runat="server" Text="Test" Visible="False" Font-Bold="True" 
                Font-Size="X-Large" ForeColor="Red"></asp:Label>
        <div style="position: absolute; font-weight: bold; top: 50px; height: 20px; width: 278px; left: 456px;">
            Please Select Category(s) To Display:</div>
    <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <br />
        </div>
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 640px;
            width: 35px; height: 25px; top: 502px; right: 402px;" />
    </form>
        </body>
</html>
