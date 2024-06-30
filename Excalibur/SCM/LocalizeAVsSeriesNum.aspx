<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_LocalizeAVsSeriesNum" Codebehind="LocalizeAVsSeriesNum.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.parent.close();
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
    <table style="width: 280px; height: 150px;">
        <tr id="trKMAT" runat="server">
            <td colspan="2" style="font-size:9px; color:red;font-weight: bold; font-family: Verdana; text-align: left;">
                KMAT is not saved in Program Data. Cannot add localized AV.
            </td>
        </tr>
        <tr>
            <td colspan="2">
                <asp:Label ID="lblHeader" runat="server" Text="Please Select..." Width="270px" Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center;"></asp:Label>           
            </td>
        </tr>
        <tr>
            <td> <asp:Label ID="lblSeriesNum" runat="server" Text="Series Number" Style="font-size: small;
                font-weight: bold; font-family: Verdana; text-align: center; width: 119px;"></asp:Label>          
            </td>
            <td>
                <asp:Label ID="lblAvType" runat="server" Text="AV Type" Width="100px" Style="font-size: small;
                font-weight: bold; font-family: Verdana; text-align: center;"></asp:Label>
            </td>
        </tr>
        <tr>
            <td><asp:DropDownList ID="ddlSeriesNumbers" runat="server" Style="height: 23px; width: 100px; top: 62px;"></asp:DropDownList>              
            </td>
            <td>
                <asp:DropDownList ID="ddlAvType" runat="server" Style="height: 23px; width: 100px; top: 62px; margin-bottom: 0px;">
                    <asp:ListItem Value="0" Text=""></asp:ListItem>
                    <asp:ListItem Value="2">HW Kits</asp:ListItem>
                    <asp:ListItem Value="3">Keyboards</asp:ListItem>
                    <asp:ListItem Value="1">OS Loc</asp:ListItem>
                    <asp:ListItem Value="4">OS Restore</asp:ListItem>
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td colspan="2"><hr /></td>
        </tr>
        <tr>
            <td colspan="2" style="text-align:right">
                <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="width: 35px; height: 24px" />
                <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="width: 61px; height: 24px" OnClientClick="cmdCancel_onclick()" />
            </td>
        </tr>
        </table>      
    </form>
</body>
</html>
