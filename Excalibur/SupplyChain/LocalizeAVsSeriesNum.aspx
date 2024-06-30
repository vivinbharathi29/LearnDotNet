<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SupSCM_LocalizeAVsSeriesNum" Codebehind="LocalizeAVsSeriesNum.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        var pulsarplusDivId = $('#pulsarplusDivId').val();
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <textarea id="holdtext" style="display: none;" rows="0" cols="0"></textarea>
    <div style="width: 272px; height: 120px;">
        <asp:Label ID="lblHeader" runat="server" Text="Please Select..." Width="274px" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute"></asp:Label>
        <asp:Label ID="lblSeriesNum" runat="server" Text="Series Number" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 35px; left: 10px; width: 119px;"></asp:Label>
        <asp:Label ID="lblAvType" runat="server" Text="AV Type" Width="100px" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 35px; left: 172px;"></asp:Label>
        <br />
        <asp:DropDownList ID="ddlAvType" runat="server" Style="position: absolute; left: 174px;
            height: 23px; width: 100px; top: 62px; margin-bottom: 0px;">
            <asp:ListItem Value="0" Text=""></asp:ListItem>
            <asp:ListItem Value="2">HW Kits</asp:ListItem>
            <asp:ListItem Value="3">Keyboards</asp:ListItem>
            <asp:ListItem Value="1">OS Loc</asp:ListItem>
            <asp:ListItem Value="4">OS Restore</asp:ListItem>
        </asp:DropDownList>
        <asp:DropDownList ID="ddlSeriesNumbers" runat="server" Style="position: absolute;
            left: 18px; height: 23px; width: 100px; top: 62px;">
        </asp:DropDownList>
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <hr />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 179px;
            width: 35px; height: 24px" />
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 220px; width: 61px; height: 24px" OnClientClick="cmdCancel_onclick()" />
        <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
    </div>
    </form>
</body>
</html>
