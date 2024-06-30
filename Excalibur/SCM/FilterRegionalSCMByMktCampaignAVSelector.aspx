<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_FilterRegionalSCMByMktCamp" Codebehind="FilterRegionalSCMByMktCampaignAVSelector.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Filter Regional AV Data by Mkt Campaign</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        #form1
        {
            height: 500px;
            width: 441px;
        }
    </style>
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div style="width: 288px; height: 433px; margin-left: 0px;">
        <asp:Label ID="lblHeader" runat="server" Text="Filter Regional Data By Mkt Campaign:"
            Style="font-size: medium; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 15px; left: 10px; width: 288px; height: 19px;"></asp:Label>
        <asp:ListBox ID="lbMktCamp" runat="server" 
            DataTextField="CampaignName" DataValueField="MktCampaignsID" Style="position: absolute; left: 22px; height: 332px;
            width: 261px; top: 55px"></asp:ListBox>
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
<br />
<br />
<br />
            <asp:Label ID="TestLabel" 
                style="position: absolute; top: 454px; left: 8px; height: 59px; width: 284px; margin-right: 0px;" 
                runat="server" Text="Test" Visible="False" Font-Bold="True" 
                Font-Size="X-Large" ForeColor="Red"></asp:Label>
    <br /><br /><br /><br /><br /><br /><br /><br />
        <br />
        <br />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 168px;
            width: 35px; height: 24px; top: 422px;" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 212px; width: 61px; height: 24px; top: 422px;" 
            OnClientClick="cmdCancel_onclick()" />
        <hr Style="position: absolute; top: 395px; left: 10px;"/>
        </div>
    <p>
        <asp:Button ID="btnTest" runat="server" Text="Test" Style="position: absolute; left: 360px;
            width: 44px; height: 24px; top: 31px;" Visible="False" />
    </p>
    </form>
        </body>
</html>
