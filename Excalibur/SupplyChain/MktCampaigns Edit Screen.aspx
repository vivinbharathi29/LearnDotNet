<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SupMktCampaigns_Edit_Screen" Codebehind="MktCampaigns Edit Screen.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Add/Edit/Delete Marketing Campaigns:</title>
    <style type="text/css">
.FrozenHeader
        {
            /*background-color: #6B696B;*/
            font-family: Verdana;
            font-size: xx-small;
            position: relative;
            cursor: default;
            top: expression(document.getElementById("GridViewContainer").scrollTop-2);
            z-index: 10;
        }
        .BottomBorder
        {
            border-bottom-width: 2px;
            border-bottom-color: rgb(120,120,120);
            border-bottom-style: solid;
        }
        
        #form1
        {
            height: 464px;
            width: 690px;
        }
        
        </style>
        
        
<script type="text/javascript">
    function cmdCancel_onclick() {
        this.window.parent.close();
        return false;
    }
</script>

</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
        <br />
        <asp:Button ID="btnTest" runat="server" Text="Test" Style="position: absolute; left: 13px;
            width: 44px; height: 24px; top: 443px;" Visible="False" />
        <br />

        <!--<asp:RadioButton ID="rbCons" 
            Style="position: absolute; top: 58px; left: 277px;" runat="server" Text="Consumer" AutoPostBack="true"/>
        <asp:RadioButton ID="rbAll" 
            style="position: absolute; top: 58px; left: 365px;" runat="server" Text="All" AutoPostBack="true" />
        <br />
        <asp:RadioButton ID="rbComm" 
            Style="position: absolute; top: 58px; left: 179px;" runat="server" Text="Commercial" AutoPostBack="true" /> -->
        <asp:Label ID="lblHeader" runat="server" Text="Add/Edit/Delete Marketing Campaigns:"
            Style="font-size: medium; font-weight: bold; font-family: Verdana; text-align: center;
            position: absolute; top: 27px; left: 21px; width: 376px; height: 19px;"></asp:Label>

            <div style="position: absolute; font-weight: bold; top: 68px; height: 20px; width: 171px; left: 600px; right: 313px; text-align: right; visibility: hidden">
            Marketing Campaign ID:</div>

        <br />

            <div style="position: absolute; font-weight: bold; top: 69px; height: 20px; width: 194px; left: 21px; right: 869px; text-align: right;">
            Marketing Campaign Name:</div>

        <hr style="position:absolute; top: 421px; left: 13px; width: 619px; height: -12px;" />
        <br />
    <asp:TextBox style="position:absolute; top:69px; left: 777px; width:45px; height: 18px;" 
                ID="txtMktCampID" runat="server" MaxLength="100" Visible="False"></asp:TextBox>

    <br />
        <asp:Button ID="btnSaveNo" runat="server" Text="No" Style="position: absolute; left: 575px;
            width: 41px; height: 23px; top: 575px; " />
        <asp:Button ID="btnSaveYes" runat="server" Text="Yes" Style="position: absolute; left: 518px;
            width: 41px; height: 23px; top: 575px; " />
        <asp:Button ID="btnSave" runat="server" Text="Save" Style="position: absolute; left: 380px;
            width: 62px; height: 25px; top: 439px; " />
        <p>
    <asp:TextBox style="position:absolute; top:68px; left: 221px; width:346px; height: 21px;" 
                ID="txtMktCampName" runat="server" MaxLength="100"></asp:TextBox>

    <asp:TextBox style="position:absolute; top:105px; left: 223px; width:119px; height: 23px;" 
                ID="txtStartDate" runat="server"></asp:TextBox>

    <asp:TextBox style="position:absolute; top:104px; left: 448px; width:119px; height: 23px;" 
                ID="txtEndDate" runat="server"></asp:TextBox>

            <div style="position: absolute; font-weight: bold; top: 138px; height: 20px; width: 221px; left: 302px; right: 376px; text-align: right;">
                Active Marketing Campaign?</div>

            <div style="position: absolute; font-weight: bold; top: 107px; height: 20px; width: 80px; left: 362px; text-align: right;">
                End Date:</div>

            <div style="position: absolute; font-weight: bold; top: 107px; height: 20px; width: 194px; left: 23px; right: 692px; text-align: right;">
                Start Date:</div>

        <asp:CheckBox style="position: absolute; font-weight: bold; top: 137px; height: 20px; width: 108px; left: 524px; " 
            ID="chkActive" runat="server" />

        <asp:CheckBox style="position: absolute; font-weight: bold; top: 140px; height: 17px; width: 122px; left: 85px; " 
            ID="chkNewRec" runat="server" Visible="False" />

        <div style="position: absolute; font-weight: bold; top: 194px; height: 16px; width: 202px; left: 129px; right: 781px;">
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
    </div>

        </p>
        <p>
        </p>
        <p>
            &nbsp;</p>
        <p>
        <asp:Button ID="btnDelete" runat="server" Text="Delete" Style="position: absolute; left: 460px;
            width: 62px; height: 25px; top: 439px; right: 590px;" />
        </p>
        <p>
            <asp:Label ID="TestLabel" 
                style="position: absolute; top: 476px; left: 10px; height: 71px; width: 604px;" 
                runat="server" Text="Test" Visible="False" Font-Bold="True" 
                Font-Size="X-Large" ForeColor="Red"></asp:Label>
        </p>
        <p>
        <asp:ListBox ID="lbPlants" runat="server" 
            DataTextField="PlantName" DataValueField="RCTOPlantsID"
            SelectionMode="Multiple" Style="position: absolute; left: 126px; height: 194px;
            width: 394px; top: 216px; right: 592px;"></asp:ListBox>
        </p>
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 542px; width: 62px; height: 24px; top: 439px;" OnClientClick="cmdCancel_onclick();" />
    </form>
</body>
</html>
