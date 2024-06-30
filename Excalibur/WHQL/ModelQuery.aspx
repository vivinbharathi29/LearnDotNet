<%@ Page Language="VB" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        SqlDataSource1.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings.Item("PRSConnectionString").ConnectionString
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SqlDataSource1.SelectParameters(0).DefaultValue = TextBox1.Text.Trim()
        GridView1.DataBind()
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SKU Model Search</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body id="ajaxProgress">
    <form id="form1" runat="server">
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <h2>
                SKU Model Search</h2>
            &nbsp;
            <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                <ContentTemplate>
                    <asp:Panel ID="Panel1" DefaultButton="Button1" runat="server">
                        <asp:Label ID="Label1" runat="server" Text="Model No:"></asp:Label>&nbsp;<asp:TextBox
                            ID="TextBox1" runat="server"></asp:TextBox>&nbsp;
                        <asp:Button ID="Button1" runat="server" Text="Search" OnClick="Button1_Click" /></asp:Panel>
                    <asp:UpdateProgress ID="UpdateProgress1" runat="server">
                        <ProgressTemplate>
                            <iframe frameborder="0" src="about:blank" style="border: 0px; position: absolute;
                                z-index: 9; left: 0px; top: 0px; width: expression(this.offsetParent.scrollWidth);
                                height: expression(this.offsetParent.scrollHeight); filter: progid:DXImageTransform.Microsoft.Alpha(Opacity=65, FinishOpacity=0, Style=0, StartX=0, FinishX=100, StartY=0, FinishY=100);">
                            </iframe>
                            <div style="font-family: Arial; font-size: 12px; position: absolute; z-index: 10;
                                left: expression((this.offsetParent.clientWidth/2)-(this.clientWidth/2)+this.offsetParent.scrollLeft);
                                top: expression((this.offsetParent.clientHeight/2)-(this.clientHeight/2)+this.offsetParent.scrollTop);">
                                <table align="center" cellpadding="0" cellspacing="0" border="0" width="150" height="50"
                                    style="background-color: #FCFCFC">
                                    <tr>
                                        <td width="13px">
                                            &nbsp;</td>
                                        <td valign="middle" align="center">
                                            <asp:Image ID="Image1" runat="server" ImageUrl="~/images/loading19.gif" /></td>
                                        <td valign="middle" style="font-family: Arial; font-size: 12px;">
                                            Processing...</td>
                                    </tr>
                                </table>
                            </div>
                        </ProgressTemplate>
                    </asp:UpdateProgress>
                    <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" CellPadding="4"
                        DataSourceID="SqlDataSource1" ForeColor="#333333" GridLines="None" AllowPaging="True"
                        AllowSorting="True" PageSize="25" Width="100%">
                        <EmptyDataTemplate>
                            Model Number Not Found.</EmptyDataTemplate>
                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                        <Columns>
                            <asp:BoundField DataField="Family" HeaderText="Family" SortExpression="Family" />
                            <asp:BoundField DataField="Version" HeaderText="Version" SortExpression="Version" />
                            <asp:BoundField DataField="SeriesSummary" HeaderText="SeriesSummary" SortExpression="SeriesSummary" />
                            <asp:BoundField DataField="OS" HeaderText="OS" SortExpression="OS" />
                            <asp:BoundField DataField="SkuDescription" HeaderText="SkuDescription" SortExpression="SkuDescription" />
                            <asp:BoundField DataField="SkuModel" HeaderText="SkuModel" SortExpression="SkuModel" />
                            <asp:BoundField DataField="SkuNo" HeaderText="SkuNo" SortExpression="SkuNo" />
                            <asp:BoundField DataField="CreatedDate" HeaderText="CreatedDate" SortExpression="CreatedDate" />
                        </Columns>
                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" HorizontalAlign="Left" />
                        <EditRowStyle BackColor="#999999" />
                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    </asp:GridView>
                    <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="Server=x;Initial Catalog=PRS;User ID=x;Password=x;"
                        ProviderName="System.Data.SqlClient" SelectCommand="SELECT DISTINCT &#13;&#10;                      ProductFamily.Name AS Family, ProductVersion.Version, Product_Brand.SeriesSummary, OSLookup.Name AS OS, ProductSKU.SkuDescription, &#13;&#10;                      ProductSKU.SkuModel, ProductSKU.SkuNo, ProductSKU.CreatedDate&#13;&#10;FROM         Product_Brand INNER JOIN&#13;&#10;                      ProductSKU INNER JOIN&#13;&#10;                      ProductSKUComponent ON ProductSKU.ID = ProductSKUComponent.SkuID ON Product_Brand.ID = ProductSKU.ProductBrandID INNER JOIN&#13;&#10;                      ProductVersion ON Product_Brand.ProductVersionID = ProductVersion.ID INNER JOIN&#13;&#10;                      ProductFamily ON ProductFamily.ID = ProductVersion.ProductFamilyID LEFT OUTER JOIN&#13;&#10;                      ImageDefinitions INNER JOIN&#13;&#10;                      Images ON ImageDefinitions.ID = Images.ImageDefinitionID INNER JOIN&#13;&#10;                      OSLookup ON ImageDefinitions.OSID = OSLookup.ID ON Images.ID = ProductSKUComponent.ImageID WHERE (ProductSKU.SkuModel LIKE '%' + @Param1 + '%') ORDER BY ProductSKU.SkuModel">
                        <SelectParameters>
                            <asp:Parameter Name="Param1" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                </ContentTemplate>
            </asp:UpdatePanel>
            &nbsp; &nbsp;&nbsp;
        </div>
    </form>
</body>
</html>
