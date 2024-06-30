<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.AvsMissingData" EnableEventValidation="false" Codebehind="AvsMissingData.aspx.vb" %>

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

    function CopyToClipboard(value) {
        window.clipboardData.setData("Text", value);
    }
 
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>AVs Missing Corporate Data</title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label ID="lblHeader" runat="server" Text="Please Select AVs To Hide (SCM / Program Matrix)"
        Style="font-size: small; font-weight: bold; font-family: Verdana; text-align: center;
        position: absolute; top: 15px; left: 10px; width: 1300px;"></asp:Label><!--<a id="lbExport" runat="server" href="#" onclick="ExportReport();">Export To Excel</a>-->
    <br />
    <div runat="server" style="position: absolute; top: 471px; align:left; width: 30%;">
		<div runat="server" style="width: 15%; float: left; display: inline-block; align:left;">
			 <asp:Button ID="btnSubmit" runat="server" Text="OK" Style=" 
				width: 35px; height: 24px;" UseSubmitBehavior="true" />
		</div>
		<div runat="server" style="width: 15%; display: inline-block; align:left;">
			 <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="
				width: 61px; height: 24px;" OnClientClick="cmdCancel_onclick()" />
		</div>
	</div>
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 1280px" />
    <div runat="server" style="position: absolute; overflow: auto; top: 41px; left: 12px;
        width: 1300px; height: 410px">
        <asp:GridView ID="gvAVsMissingData" runat="server" GridLines="vertical" AutoGenerateColumns="False"
            CellPadding="4" ForeColor="Black" BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid"
             HeaderStyle-Wrap="false">
            <FooterStyle BackColor="#CCCC99" />
            <RowStyle BackColor="#F7F7DE" />
            <Columns>
                <asp:TemplateField HeaderText="Hide" HeaderStyle-HorizontalAlign="Center">
                    <HeaderTemplate>
                        <center>
                            <asp:CheckBox ID="cbxAll" runat="server" AutoPostBack="true" OnCheckedChanged="cbxAll_Checkedchanged" />
                        </center>
                    </HeaderTemplate>
                    <ItemTemplate>
                        <center>
                            <asp:CheckBox ID="cbxAVsMissingData" runat="server" />
                        </center>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="AV Number" HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <left>
                             <asp:Label ID="lblAvNo" runat="server" Text='<%#Eval("AvNo") %>' Width="100px"/>
                        </left>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="GPG Description" HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <left>
                             <asp:Label ID="lblGPGDescription" runat="server" Text='<%#Eval("GPGDescription") %>' Width="205px"/>
                        </left>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Marketing Description (40 Char GPSy)" HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <left>
                             <asp:Label ID="lblMarketingDescription" runat="server" Text='<%#Eval("MarketingDescription") %>' Width="300px"/>
                        </left>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Marketing Description (100 Char PMG)" HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <left>
                             <asp:Label ID="lblMarketingDescriptionPMG" runat="server" Text='<%#Eval("MarketingDescriptionPMG") %>' Width="300px"/>
                        </left>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="CPLBlindDt" HeaderText="SA Date"
                    HeaderStyle-HorizontalAlign="Center" DataFormatString="{0:MM/dd/yyyy}" HtmlEncode="false">
                      <ItemStyle Width="100px"></ItemStyle>
                </asp:BoundField>
                 <asp:BoundField DataField="GeneralAvailDt" HeaderText="GA Date"
                    HeaderStyle-HorizontalAlign="Center" DataFormatString="{0:MM/dd/yyyy}" HtmlEncode="false">
                    <ItemStyle Width="100px"></ItemStyle>
                </asp:BoundField>
                <asp:BoundField DataField="RASDiscontinueDt" HeaderText="EM Date"
                    HeaderStyle-HorizontalAlign="Center" DataFormatString="{0:MM/dd/yyyy}" HtmlEncode="false">
                    <ItemStyle Width="100px"></ItemStyle>
                </asp:BoundField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblAvDetailID" runat="server" Text='<%#Eval("AvDetailID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblStatus" runat="server" Text='<%#Eval("Status") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblRasDiscoSysUpdate" runat="server" Text='<%#Eval("RasDiscoSysUpdate") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblMktDescSysUpdate" runat="server" Text='<%#Eval("MktDescSysUpdate") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblCplBlindSysUpdate" runat="server" Text='<%#Eval("CplBlindSysUpdate") %>' />
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
