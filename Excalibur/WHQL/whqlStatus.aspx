<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dl As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As System.Data.DataTable = dl.GetProductVersion(Request.QueryString("PVID"))
        If dt.Rows.Count > 0 Then
            Label6.Text = String.Format("{0} {1} WHQL Status", dt.Rows(0)("Name"), dt.Rows(0)("Version"))
        End If

        If DataList1.Items.Count > 0 Then
            Label5.Visible = False
        End If
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>WHQL Status</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <h2>
            <asp:Label ID="Label6" runat="server" Text="Platform WHQL Status"></asp:Label></h2>
        <p>
            <asp:Label ID="Label2" runat="server" Text="WHQL Submissions" CssClass="Heading"></asp:Label><br />
            <asp:DataList ID="DataList1" runat="server" DataSourceID="odsWhqlSubmissions" RepeatDirection="Horizontal"
                RepeatLayout="Flow" ShowFooter="False" ShowHeader="False">
                <ItemTemplate>
                    <strong><%# Container.DataItem( "SubmissionID" )%></strong>
                </ItemTemplate>
                <SeparatorTemplate>
                    ,
                </SeparatorTemplate>
            </asp:DataList>
            <asp:Label ID="Label5" runat="server" Text="There are no submissions to display" Font-Bold="True" ForeColor="Red"></asp:Label>
            <asp:ObjectDataSource ID="odsWhqlSubmissions" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectProductWhql" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:Parameter Name="ProductWhqlID" Type="String" />
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <p>
            <asp:Label ID="lblSkuWithoutWhql" runat="server" Text="SKUs Without WHQL Coverage"
                CssClass="Heading"></asp:Label><br />
            <asp:GridView ID="GridView1" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                DataSourceID="odsSkusWithoutWhql" AllowPaging="true">
                <Columns>
                    <asp:BoundField DataField="SkuNo" HeaderText="SKU" />
                    <asp:BoundField DataField="SkuDescription" HeaderText="SKU Desc." />
                    <asp:BoundField DataField="BaseUnitAvPN" HeaderText="Base Unit AV" />
                    <asp:BoundField DataField="CpuAvPN" HeaderText="CPU AV" />
                    <asp:BoundField DataField="ImagePN" HeaderText="Image ZWAR" ItemStyle-Wrap="false" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsSkusWithoutWhql" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectProductSkusWithoutWhql" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <p>
            <asp:Label ID="lblUnsignedDrivers" runat="server" Text="Images With Unsigned Drivers"
                CssClass="Heading"></asp:Label>
            <asp:GridView ID="GridView2" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                DataSourceID="odsUnsignedDrivers" AllowPaging="True">
                <Columns>
                    <asp:BoundField DataField="ZWAR" HeaderText="Image ZWAR" />
                    <asp:BoundField DataField="Brand" HeaderText="Brand" />
                    <asp:BoundField DataField="OS" HeaderText="OS" />
                    <asp:BoundField DataField="SWType" HeaderText="Software Type" />
                    <asp:BoundField DataField="DLLName" HeaderText="Driver File Name" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsUnsignedDrivers" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectImagesWithUnsignedDrivers" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <h2>
            <asp:Label ID="lblWindowsXPImages" runat="server" Text="Label">Windows XP Platform Requirements</asp:Label></h2>
        <p>
            <asp:Label ID="lblSigVerifyResults" runat="server" Text="SKUs Containing Images Without SigVerify Results"
                CssClass="Heading"></asp:Label>
            <asp:GridView ID="GridView3" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                DataSourceID="odsSigVerifyResults" AllowPaging="True">
                <Columns>
                    <asp:BoundField DataField="SkuNo" HeaderText="SKU Number" />
                    <asp:BoundField DataField="ImageZWAR" HeaderText="Image ZWAR" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsSigVerifyResults" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectSkuImageStatus" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                    <asp:Parameter DefaultValue="1" Name="OsFamilyID" Type="String" />
                    <asp:Parameter DefaultValue="0" Name="SigVerifyComplete" Type="String" />
                    <asp:Parameter DefaultValue="" Name="CheckLogo6Complete" Type="String" />
                    <asp:Parameter DefaultValue="" Name="WmiComplete" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <p>
            <asp:Label ID="Label1" runat="server" Text="WHQL Submissions Without BootVis Results"
                CssClass="Heading"></asp:Label>
            <asp:GridView ID="GridView4" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                DataSourceID="odsBootVisStatus" AllowPaging="true">
                <Columns>
                    <asp:BoundField DataField="SkuNo" HeaderText="SKU Number" />
                    <asp:BoundField DataField="ImageZWAR" HeaderText="Image ZWAR" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsBootVisStatus" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectProductWHQLWithoutBootVis" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <h2>
            <asp:Label ID="lblVistaImages" runat="server" Text="Label">Windows Vista Platform Requirements</asp:Label></h2>
        <p>
            <asp:Label ID="Label3" runat="server" Text="SKUs Containing Images Without CheckLogo6 Results"
                CssClass="Heading"></asp:Label>
            <asp:GridView ID="GridView5" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                AllowPaging="True" DataSourceID="odsCheckLogo6Results">
                <Columns>
                    <asp:BoundField DataField="SkuNo" HeaderText="SKU Number" />
                    <asp:BoundField DataField="ImageZWAR" HeaderText="Image ZWAR" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsCheckLogo6Results" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectSkuImageStatus" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                    <asp:Parameter DefaultValue="2" Name="OsFamilyID" Type="String" />
                    <asp:Parameter DefaultValue="" Name="SigVerifyComplete" Type="String" />
                    <asp:Parameter DefaultValue="0" Name="CheckLogo6Complete" Type="String" />
                    <asp:Parameter DefaultValue="" Name="WmiComplete" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
        <p>
            <asp:Label ID="Label4" runat="server" Text="SKUs Containing Images Without WMI Results"
                CssClass="Heading"></asp:Label>
            <asp:GridView ID="GridView6" runat="server" CssClass="FormTable" AutoGenerateColumns="False"
                AllowPaging="True" DataSourceID="odsWmiStatus">
                <Columns>
                    <asp:BoundField DataField="SkuNo" HeaderText="SKU Number" />
                    <asp:BoundField DataField="ImageZWAR" HeaderText="Image ZWAR" />
                </Columns>
                <EmptyDataTemplate>
                    <asp:Label ID="lblSkusWithoutWhqlNoRecordsFound" runat="server" Text="No Records Found"></asp:Label>
                </EmptyDataTemplate>
            </asp:GridView>
            <asp:ObjectDataSource ID="odsWmiStatus" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectSkuImageStatus" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                    <asp:Parameter DefaultValue="2" Name="OsFamilyID" Type="String" />
                    <asp:Parameter DefaultValue="" Name="SigVerifyComplete" Type="String" />
                    <asp:Parameter DefaultValue="" Name="CheckLogo6Complete" Type="String" />
                    <asp:Parameter DefaultValue="0" Name="WmiComplete" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </p>
    </form>
</body>
</html>
