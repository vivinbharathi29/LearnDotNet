<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.FilterByMktCampaignsAVSelector" Codebehind="FilterByMktCampaignsAVSelector.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }

    function AvMktViewSelectionTool() {
        var querystring = window.location.search;
        var strID;

        //alert(querystring);
    }

    function TestFunction() {
        alter("Test");
    }
    function OpenEditScreen() {
        
    }

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Choose Marketing Campaign for this Region</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #form1
        {
            height: 590px;
            width: 574px;
        }
    </style>
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 573px; height: 529px; margin-left: 0px;">
        <asp:Label ID="lblHeader" runat="server" Text="Select Marketing Campaign to Display"
            Style="font-size: medium; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 25px; left: 21px; width: 760px; height: 19px;"></asp:Label>
    <asp:HyperLink ID="HyperLink1" runat="server" Target="_blank" NavigateUrl="MktCampaigns Edit Screen.aspx" Visible="False">HyperLink</asp:HyperLink>
        <br />
        <br />
        <asp:DropDownList ID="cboMktCamp" runat="server" DataTextField="CampaignName" DataValueField="MktCampaignsID"
            
            
            style="position: absolute; font-weight: bold; top: 89px; height: 26px; width: 236px; left: 324px; right: 622px;">
        </asp:DropDownList>
        <br />
        <asp:Button ID="btnAddCamps" runat="server" Text="Add Campaigns" 
            OnClientClick="AvMktViewSelectionTool()" Style="position: absolute; left: 336px;
            width: 99px; height: 26px; top: 54px; right: 747px;" />
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
    <br />
    <br />
    <br />
        <asp:ListBox ID="lbCategories" runat="server" DataTextField="AvFeatureCategory" DataValueField="AvFeatureCategoryID"
            SelectionMode="Multiple" Style="position: absolute; left: 232px; height: 382px;
            width: 327px; top: 151px"></asp:ListBox>
            <asp:Label ID="TestLabel" 
                style="position: absolute; top: 539px; left: 15px; height: 67px; width: 412px;" 
                runat="server" Text="Test" Visible="False" Font-Bold="True" 
                Font-Size="X-Large" ForeColor="Red"></asp:Label>
        <div style="position: absolute; font-weight: bold; top: 91px; height: 20px; width: 303px; left: 17px; ">
            Please Select a Go To Marketing Campaign:</div>
        <div style="position: absolute; font-weight: bold; top: 124px; height: 20px; width: 278px; left: 257px;">
            Please Select Category(s) To Display:</div>
    <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <asp:Label ID="Label4" 
            style="position: absolute; top: 53px; left: 630px; width: 99px;" runat="server" 
            Text="Mkt Camp Name:" Visible="false"></asp:Label>
        <asp:Label ID="Label3" 
            
            style="position: absolute; top: 26px; left: 649px; width: 80px; height: 12px; bottom: 190px;" runat="server" 
            Text="Mkt Camp ID:" Visible="false"></asp:Label>
        <asp:Label ID="Label1" 
            style="position: absolute; top: 90px; left: 668px; width: 59px;" runat="server" 
            Text="Plant ID:" Visible="false"></asp:Label>
        <asp:TextBox ID="txtMktCampName" style="position: absolute; top: 48px; left: 741px;" 
            runat="server" Visible="false"></asp:TextBox>
        <asp:TextBox ID="txtMktCampID" style="position: absolute; top: 20px; left: 740px;" 
            runat="server" Visible="false"></asp:TextBox>
        <asp:TextBox ID="txtPlantID" style="position: absolute; top: 88px; left: 742px;" 
            runat="server" Visible="false"></asp:TextBox>
        <asp:Label ID="Label2" 
            style="position: absolute; top: 124px; left: 647px; width: 76px; height: 19px;" 
            runat="server" Text="Plant Name:" Visible="False"></asp:Label>
        <asp:TextBox ID="txtPlantName" style="position: absolute; top: 123px; left: 742px; height: 22px;" 
            runat="server" Visible="false"></asp:TextBox>
        <br />
        </div>
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 442px;
            width: 35px; height: 25px; top: 552px; right: 705px;" />
    <p>
        <asp:Button ID="btnEditCamps" runat="server" Text="Edit Campaigns"
            
            Style="position: absolute; width: 99px; height: 26px; top: 54px; right: 625px; left: 458px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 490px; width: 62px; height: 24px; top: 553px;" />
        <asp:TextBox ID="txtGeoID" style="position: absolute; top: 238px; left: 740px;" 
            runat="server" Visible="false"></asp:TextBox>
    &nbsp;</p>
    <p>
        <asp:TextBox ID="txtMktEndDate" style="position: absolute; top: 198px; left: 738px; bottom: 127px;" 
            runat="server" Visible="false"></asp:TextBox>
        </p>

    <p>
        <asp:Label ID="lblGeoID" 
            style="position: absolute; top: 243px; left: 648px; width: 82px; height: 13px; bottom: 58px; " runat="server" 
            Text="Geo ID:" Visible="false"></asp:Label>
        <asp:Label ID="lblMktEndDate" 
            
            style="position: absolute; top: 205px; left: 648px; width: 82px; height: 13px; bottom: 129px; right: 452px;" runat="server" 
            Text="Mkt End Date:" Visible="False"></asp:Label>
        <asp:TextBox ID="txtMktStartDate" style="position: absolute; top: 170px; left: 740px; bottom: 166px;" 
            runat="server" visible="false"></asp:TextBox>
        </p>

    <p>
        <asp:Label ID="lblMktStartDate" 
            style="position: absolute; top: 175px; left: 642px; width: 86px; height: 13px; bottom: 203px; right: 454px;" runat="server" 
            Text="Mkt Start Date:" Visible="False"></asp:Label>
        </p>

    <asp:CheckBox ID="chkActive" 
        style="position:absolute; top: 271px; left: 743px; height: 20px; width: 73px; right: 366px;" 
        runat="server" Text="Active?" visible="false"/>

    </form>
        </body>
</html>
