<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SupSCM_DocKitData" Codebehind="DocKitData.aspx.vb" %>

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
    <asp:Label ID="lblHeader" runat="server" Style="font-size: small; font-weight: bold;
        font-family: Verdana; text-align: center; position: absolute; top: 12px; left: 12px;
        width: 508px; height: 17px;"></asp:Label>
    <div style="width: 508px; height: 600px; overflow:auto; position:absolute">
        <asp:GridView ID="gvDocKitData" runat="server" GridLines="vertical" AutoGenerateColumns="False"
            CellPadding="4" Style="position: absolute; top: 12px; left: 12px;" ForeColor="Black"
            BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid">
            <FooterStyle BackColor="#CCCC99" />
            <RowStyle BackColor="#F7F7DE" />
            <Columns>
                <asp:BoundField DataField="AssemblyNo_fk" HeaderText="Doc Kit Number" HeaderStyle-Wrap="false"
                    ItemStyle-Width="150px" HeaderStyle-Width="150px">
                    <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" />
                </asp:BoundField>
                <asp:BoundField DataField="AssyDash_fk" HeaderText="Doc Kit Dash" HeaderStyle-Wrap="false"
                    ItemStyle-Width="150px" HeaderStyle-Width="150px">
                    <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" />
                </asp:BoundField>
                <asp:BoundField DataField="HPDash" HeaderText="HP Dash" HeaderStyle-Wrap="false"
                    ItemStyle-Width="150px" HeaderStyle-Width="150px">
                    <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                    <ItemStyle Wrap="False" />
                </asp:BoundField>
            </Columns>
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#FFFFF0" />
        </asp:GridView>
    </div><hr style="margin-bottom: 3px; position: absolute; top: 635px; left: 10px; width: 520px;
        height: -11px;" />
    <asp:Button ID="btnCancel" runat="server" Text="Close" Style="position: absolute;
        left: 450px; width: 71px; height: 24px; top: 650px;" UseSubmitBehavior="true" />
    </form>
</body>
</html>
