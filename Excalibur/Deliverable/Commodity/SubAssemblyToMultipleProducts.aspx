<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SubAssemblyToMultipleProducts" EnableEventValidation="false" Codebehind="SubAssemblyToMultipleProducts.aspx.vb" %>

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
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label ID="lblHeader" runat="server"
        Style="font-size: x-small; font-weight: bold; font-family: Verdana; text-align: left;
        position: absolute; top: 15px; left: 13px; width: 292px;"></asp:Label>
    <br />
    <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 205px;
        width: 35px; height: 24px; top: 471px;" UseSubmitBehavior="true" />
    <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
        left: 250px; width: 61px; height: 24px; top: 471px;" 
        OnClientClick="cmdCancel_onclick()" />
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 302px" />
    <div runat="server" style="position: absolute; overflow: auto; top: 41px; left: 12px; width: 298px; height:410px">
        <asp:GridView ID="gvProducts" runat="server" GridLines="vertical" AutoGenerateColumns="False"
            CellPadding="4" ForeColor="Black" BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid"
            Style="width: 280px">
            <FooterStyle BackColor="#CCCC99" />
            <RowStyle BackColor="#F7F7DE" />
            <Columns>
                <asp:TemplateField HeaderText="Apply" HeaderStyle-HorizontalAlign="Center">
                    <HeaderTemplate>
                        <center>
                            <asp:CheckBox ID="cbxAll" runat="server" AutoPostBack="true" OnCheckedChanged="cbxAll_Checkedchanged" />
                        </center>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <center>
                            <asp:CheckBox ID="cbxSingle" runat="server" AutoPostBack="true" OnCheckedChanged="cbxSingle_Checkedchanged" />
                        </center>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="DOTSName" HeaderText="Product" HeaderStyle-HorizontalAlign="Center">
                </asp:BoundField>
                <asp:TemplateField HeaderText="Subassembly No." HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:Label ID="lblSubassembly" runat="server" Text='<%#Eval("Subassembly") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Subassembly No." HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:Label ID="lblServiceSubassembly" runat="server" Text='<%#Eval("ServiceSubassembly") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblDRID" runat="server" Text='<%#Eval("DRID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblPVID" runat="server" Text='<%#Eval("ProductVersionID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblProdDelRootID" runat="server" Text='<%#Eval("ProdDelRootID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#FFFFF0" />
        </asp:GridView>
    </div>
    </form>
</body>
</html>
