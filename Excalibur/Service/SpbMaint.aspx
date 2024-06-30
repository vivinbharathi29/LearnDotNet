<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_SpbMaint" Codebehind="SpbMaint.aspx.vb" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>

    <script type="text/javascript">
function btnExportSpb_OnClick()
{
    __doPostBack('btnSpbExport','');
}
    </script>

    <form id="form1" runat="server">
        <div>
            <p>
                <asp:Label ID="lblProductVersion" runat="server" CssClass="Heading" Text="Label"
                    Font-Size="Large"></asp:Label>
                <br />
                <asp:Label ID="lblServiceProgramBom" runat="server" CssClass="Heading" Text="Service Program BOM (SPB) Maintenance"></asp:Label></p>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <asp:UpdatePanel ID="UpdatePanel1" runat="server" RenderMode="Inline">
                <ContentTemplate>
                    <p>
                        <asp:Label ID="lblServiceFamilyPn" runat="server" Text="Family SPS Pn"></asp:Label>
                        <asp:Label ID="lblFamilyPn" runat="server" Visible="False"></asp:Label>
                        <asp:TextBox ID="txtFamilyPn" runat="server"></asp:TextBox>
                        <asp:Button ID="btnSaveFamilyPn" runat="server" OnClick="btnSaveFamilyPn_Click" Text="Save" />
                        <asp:LinkButton ID="btnEditFamilyPn" runat="server" OnClick="btnEditFamilyPn_Click"
                            Visible="False">Edit</asp:LinkButton>
                    </p>
                    <p style="display:none;">
                        <asp:LinkButton ID="lbExportSpb" runat="server" Visible="False">Export Service Program BOM (SPB)</asp:LinkButton>
                    </p>
                    <p>
                    Export Using the Export Link on the previous screen.
                    </p>
                    <asp:Panel ID="pnlHidden" runat="server" Style="display: none;">
                        <asp:Button ID="btnSpbExport" runat="server" Text="Export" />
                    </asp:Panel>
                    <asp:Panel ID="pnlExportSpb" runat="server" Style="display: none;" CssClass="modalPopup"
                        Height="100px" Width="200px">
                        <p>
                            <asp:Label ID="lblCompareDt" runat="server" Text="Compare To: "></asp:Label>
                            <asp:DropDownList ID="ddlCompareDt" runat="server">
                            </asp:DropDownList><br />
                            <asp:CheckBox ID="cbPublishSpb" runat="server" Text="Publish" /><br />
                            <asp:CheckBox ID="cbNewSpb" runat="server" Text="New Spb" /><br />
                            <asp:Button ID="btnExportSpb" runat="server" Text="Export SPB" />&nbsp;
                            <asp:Button ID="btnCancelSpb" runat="server" Text="Cancel" />
                        </p>
                    </asp:Panel>
                    <cc1:ModalPopupExtender ID="ModalPopupExtender1" runat="server" TargetControlID="lbExportSpb"
                        OkControlID="btnExportSpb" OnOkScript="btnExportSpb_OnClick()" CancelControlID="btnCancelSpb"
                        PopupControlID="pnlExportSpb" BackgroundCssClass="modalBackground">
                    </cc1:ModalPopupExtender>
            <asp:DetailsView ID="DetailsView1" runat="server" CssClass="FormTable" Width="300px"
                AutoGenerateEditButton="True" HeaderText="Spare Kit Details" DefaultMode="Edit">
                <CommandRowStyle HorizontalAlign="Center" />
                <RowStyle HorizontalAlign="Left" Wrap="False" />
                <FieldHeaderStyle Font-Bold="True" Wrap="False" />
                <EditRowStyle Wrap="False" />
                <HeaderStyle Font-Bold="True" Font-Size="Larger" HorizontalAlign="Center" />
            </asp:DetailsView>
                    <asp:ObjectDataSource ID="odsSpbDetails" runat="server" OldValuesParameterFormatString="original_{0}"
                        SelectMethod="SelectSpbDetails" TypeName="HPQ.Excalibur.Data" UpdateMethod="UpdateServiceFamilyDetails"
                        InsertMethod="UpdateServiceFamilyDetails">
                        <InsertParameters>
                            <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text" />
                            <asp:Parameter Name="SpdmContactID" Type="String" />
                            <asp:Parameter Name="Active" Type="Boolean" />
                            <asp:Parameter Name="SharePointPath" Type="String" />
                            <asp:Parameter Name="SharedDrivePath" Type="String" />
                        </InsertParameters>
                        <UpdateParameters>
                            <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text" />
                            <asp:Parameter Name="SpdmContactID" Type="String" />
                            <asp:Parameter Name="Active" Type="Boolean" />
                            <asp:Parameter Name="SharePointPath" Type="String" />
                            <asp:Parameter Name="SharedDrivePath" Type="String" />
                        </UpdateParameters>
                        <SelectParameters>
                            <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text"
                                Type="String" />
                        </SelectParameters>
                    </asp:ObjectDataSource>
                    <asp:ObjectDataSource ID="odsSpdmUsers" runat="server" OldValuesParameterFormatString="original_{0}"
                        SelectMethod="ListSpdmUsers" TypeName="HPQ.Excalibur.Data"></asp:ObjectDataSource>
                    <asp:Panel ID="pnlDetails" runat="server" Style="display: none; text-align: center;"
                        CssClass="modalPopup" Width="300px">
                        <!-- Height="100px" Width="200px" -->
                        <asp:DetailsView ID="dvSpareKit" runat="server" CssClass="FormTable" Width="300px"
                            AutoGenerateEditButton="True" HeaderText="Spare Kit Details" DefaultMode="Edit">
                            <CommandRowStyle HorizontalAlign="Center" />
                            <RowStyle HorizontalAlign="Left" Wrap="False" />
                            <FieldHeaderStyle Font-Bold="True" Wrap="False" />
                            <EditRowStyle Wrap="False" />
                        </asp:DetailsView>
                    </asp:Panel>
                    <asp:Button runat="server" ID="hiddenTargetControlForModalPopup" Style="display: none" />
                    <cc1:ModalPopupExtender ID="mpeDetails" runat="server" TargetControlID="hiddenTargetControlForModalPopup"
                        OkControlID="" CancelControlID="" PopupControlID="pnlDetails" BackgroundCssClass="modalBackground">
                    </cc1:ModalPopupExtender>
                    <br />
                    <asp:GridView ID="gvSpareKits" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                        PageSize="20" CssClass="FormTable" DataSourceID="odsSpareKits">
                        <Columns>
                            <asp:CommandField ShowSelectButton="True">
                                <ItemStyle Width="100px" />
                            </asp:CommandField>
                            <asp:BoundField DataField="ChildPartNumber" HeaderText="Spare Kit" ReadOnly="True" />
                            <asp:BoundField DataField="RevisionLevel" HeaderText="Revision" ReadOnly="True" />
                            <asp:BoundField DataField="CrossPlantStatus" HeaderText="X Plant Status" ReadOnly="True" />
                            <asp:BoundField DataField="GPGDescription" HeaderText="Description" ReadOnly="True" />
                            <asp:BoundField DataField="ID" HeaderText="ID" Visible="False" />
                            <asp:BoundField DataField="HpPartNo" HeaderText="HpPartNo" Visible="False" />
                            <asp:BoundField DataField="Supplier" Visible="False" />
                            <asp:BoundField DataField="OsspPrderable" Visible="False" />
                            <asp:BoundField DataField="OdmPartNo" Visible="False" />
                            <asp:BoundField DataField="OdmPartDesc" Visible="False" />
                            <asp:BoundField DataField="OdmBulkPartNo" Visible="False" />
                            <asp:BoundField DataField="OdmProdMoq" Visible="False" />
                            <asp:BoundField DataField="OdmPostProdMoq" Visible="False" />
                            <asp:BoundField DataField="Model" Visible="False" />
                            <asp:BoundField DataField="SpareCategoryId" Visible="False" />
                            <asp:BoundField DataField="ReadOnly" Visible="False" />
                        </Columns>
                        <SelectedRowStyle BackColor="#8080FF" Width="50%" />
                    </asp:GridView>
                    <asp:ObjectDataSource ID="odsSpareKits" runat="server" OldValuesParameterFormatString="original_{0}"
                        SelectMethod="SelectSpareKits" TypeName="HPQ.Excalibur.Data" UpdateMethod="UpdateServiceSpareDetail">
                        <UpdateParameters>
                            <asp:Parameter Name="ServiceSpareDetailId" Type="String" />
                            <asp:Parameter Name="HpPartNo" Type="String" />
                            <asp:Parameter Name="Supplier" Type="String" />
                            <asp:Parameter Name="OsspOrderable" Type="Boolean" />
                            <asp:Parameter Name="OdmPartNo" Type="String" />
                            <asp:Parameter Name="OdmPartDesc" Type="String" />
                            <asp:Parameter Name="OdmBulkPartNo" Type="String" />
                            <asp:Parameter Name="OdmProdMoq" Type="String" />
                            <asp:Parameter Name="OdmPostProdMoq" Type="String" />
                            <asp:Parameter Name="Model" Type="String" />
                            <asp:Parameter Name="SpareCategoryId" Type="String" />
                            <asp:Parameter Name="ReadOnly" Type="Boolean" />
                        </UpdateParameters>
                        <SelectParameters>
                            <asp:ControlParameter ControlID="txtFamilyPn" Name="ServiceFamilyPn" PropertyName="Text"
                                Type="String" />
                        </SelectParameters>
                    </asp:ObjectDataSource>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>
    </form>
</body>
</html>
