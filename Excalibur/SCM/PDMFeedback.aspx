<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.PDMFeedback"
    EnableEventValidation="false" Codebehind="PDMFeedback.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label ID="lblHeader" runat="server" Style="font-size: small; font-weight: bold;
        font-family: Verdana; text-align: center; position: absolute; top: 16px; left: 10px;
        width: 1068px; height: 17px;"></asp:Label>
    <br />
    <asp:Label ID="lblAvNo" runat="server" Text="AV Number:" Style="position: absolute;
        left: 17px; height: 23px; width: 88px; top: 40px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblNotActionable" runat="server" Style="position: absolute; left: 14px;
        text-align: center; height: 23px; width: 92px; top: 473px; background-color: #FFE4E1"></asp:Label>
    <asp:Label ID="lblNotActionableText" runat="server" Style="position: absolute; left: 110px;
        height: 23px; width: 182px; top: 477px; right: 1075px;" Text="= Action Item Not Actionable"></asp:Label>
    <asp:Label ID="lblAvNoText" runat="server" Style="position: absolute; left: 93px;
        height: 23px; width: 185px; top: 40px; right: 1089px;"></asp:Label>
    <asp:Button ID="btnClose" runat="server" Text="Close" Style="position: absolute;
        left: 1002px; width: 71px; height: 24px; top: 471px;" 
        UseSubmitBehavior="true" />
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 1068px;
        height: -11px;" />
    <asp:GridView ID="gvAvActionItems" Width="1068px" runat="server" GridLines="vertical" AutoGenerateColumns="False"
        CellPadding="4" Style="position: absolute; top: 73px; left: 12px;" ForeColor="Black"
        BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid" AllowSorting="true">
        <FooterStyle BackColor="#CCCC99" />
        <RowStyle BackColor="#F7F7DE" />
        <Columns>
            <asp:BoundField DataField="ActionName" HeaderText="Action Item" HeaderStyle-Wrap="false" SortExpression="ActionName">
                <HeaderStyle HorizontalAlign="Center" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="ProductName" HeaderText="Product/Brand" HeaderStyle-Wrap="false" SortExpression="ProductName">
                <HeaderStyle HorizontalAlign="Center" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="PCDate" HeaderText="PC Date" HeaderStyle-Wrap="false" DataFormatString="{0:MM/dd/yyyy hh:mm tt}" SortExpression="PCDate">
                <HeaderStyle HorizontalAlign="Center" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="MarketingDate" HeaderText="Marketing Date" HeaderStyle-Wrap="false" DataFormatString="{0:MM/dd/yyyy hh:mm tt}" SortExpression="MarketingDate">
                <HeaderStyle HorizontalAlign="Center" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="PhWebDate" HeaderText="PhWeb Date" HeaderStyle-Wrap="false" DataFormatString="{0:MM/dd/yyyy hh:mm tt}" SortExpression="PhWebDate">
                <HeaderStyle HorizontalAlign="Center" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:TemplateField HeaderText="Feedback" HeaderStyle-Wrap="true" ItemStyle-Wrap="true" SortExpression="Feedback">
                <ItemTemplate>
                    <asp:Label ID="lblFeedback" runat="server" Text='<%#Eval("PDMFeedback") %>' />
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField Visible="false">
                <ItemTemplate>
                    <asp:Label ID="lblNotActionable" runat="server" Text='<%#Eval("NotActionable") %>' />
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
        <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
        <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#FFFFF0" />
    </asp:GridView>
    </form>
</body>
</html>
