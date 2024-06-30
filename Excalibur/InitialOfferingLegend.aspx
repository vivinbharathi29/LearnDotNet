<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.InitialOfferingLegend" EnableEventValidation="false" Codebehind="InitialOfferingLegend.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Initial Offering Legend</title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label runat="server" Text="Since last publish..." Width="100%" Style="text-align: center"></asp:Label><br />
    <br />
    <table style="width: 100%">
        <tr>
            <td>
                &nbsp; Program Checkbox Selected/Unselected
            </td>
            <td style="background-color: #FFE4E1">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp; Deliverable Added
            </td>
            <td style="background-color: #B0C4DE">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp; Deliverable Removed
            </td>
            <td>
                <asp:Label runat="server" ID="lblStrikethrough" Text="Strikethrough" Font-Strikeout="true" />
            </td>
        </tr>
        <tr>
            <td>
                &nbsp; New Deliverable
            </td>
            <td>
                <asp:Label runat="server" ID="Label1" Text="Bold" Font-Bold="true" />
            </td>
        </tr>
    </table>
    <hr />
    <table style="width: 100%">
        <tr>
            <td>
                &nbsp; Engineering Generated
            </td>
            <td>
                <asp:CheckBox ID="CheckBox1" runat="server" BackColor ="#6B696B" Checked="true" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
