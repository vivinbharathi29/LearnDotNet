<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="changelog.aspx.vb" Inherits="DummyVBApp.changelog" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SCM Change Log</title>
    <link href="../style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="../style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="style.css" />

    <script type="text/javascript">
        function BrandLink_onClick(ProductBrandID) {
            window.location.replace("changelog.aspx?ID=<%=Request("ID")%>&Class=<%=Request("Class")%>&BID=" + ProductBrandID);
        }

        function Row_OnMouseOver() {
            var node = window.event.srcElement;
            while (node.nodeName.toUpperCase() != "TR") {
                node = node.parentElement;
            }

            node.style.color = "red";
            node.style.cursor = "hand";
        }

        function Row_OnMouseOut() {
            var node = window.event.srcElement;
            while (node.nodeName.toUpperCase() != "TR") {
                node = node.parentElement;
            }

            node.style.color = "black";
        }

        function Row_OnClick() {
            var node = window.event.srcElement;

            if (node.type == "checkbox")
                return;

            while (node.nodeName.toUpperCase() != "TR") {
                node = node.parentElement;
            }

            var strID;
            strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=edit&PVID=" + node.pvid + "&CLID=" + node.clid, "", "dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
            document.location.reload();
        }

        function AddEntry(ProductVersionID, ProductBrandID) {
            var strID;
            strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=add&PVID=" + ProductVersionID + "&PBID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:275px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
            document.location.reload();
        }
    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <p>
                <span style="font-size: medium; font-weight: bold">
                    <asp:Label ID="lblProductName" runat="server" Text="Label"></asp:Label>
                    SCM Change Log</span>
            </p>
            <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
                <tr>
                    <td valign="top">
                        <table>
                            <tr>
                                <td valign="top" style="height: 14px; font-weight: bold; font-size: small; color: navy; font-family: Verdana;">Display:&nbsp;&nbsp;&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                    <td style="width: 100%">
                        <table>
                            <tr>
                                <td>
                                    <b>Brand:</b></td>
                                <td style="width: 100%">
                                    <asp:Repeater ID="rptrBrands" runat="server" OnItemDataBound="rptrBrands_ItemDataBound">
                                        <ItemTemplate>
                                            <a id="aBrand" runat="server"></a>
                                            <asp:LinkButton ID="lbBrand" runat="server">LinkButton</asp:LinkButton>
                                        </ItemTemplate>
                                        <SeparatorTemplate>
                                            &nbsp;|&nbsp;
                                        </SeparatorTemplate>
                                    </asp:Repeater>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <br />
            <br />
            <asp:Button ID="btnSaveChanges" runat="server" Text="Save Changes"
                OnClick="btnSaveChanges_Click" />
            <br />
            <br />
            <asp:LinkButton ID="lbAddItem" runat="server" Text="Add Item" />
            |
           
            <asp:LinkButton ID="lbShowAll" runat="server" Text="Show All"
                OnClick="lbShowAll_Click" />
            <br />
            <br />
            <div style="text-align: right">
                <asp:LinkButton ID="lbPagination" runat="server" Text="Un-Page"
                    OnClick="lbPagination_Click" />&nbsp;|&nbsp;
               
                <asp:Label ID="lblCurrentPage" runat="server" Text="Label" />
                <asp:LinkButton ID="lbPrev" runat="server" OnClick="lbPrev_Click" Text="<<" />
                <asp:LinkButton ID="lbNext" runat="server" OnClick="lbNext_Click" Text=">>" />
            </div>
            <span style="font-size: x-small; font-weight: bold">
                <asp:Label ID="lblBrand" runat="server" /></span> - <span style="color: Red">(Click
                    on the change row to view the details.)</span>
            <br />
            <br />
            <asp:Repeater ID="rptrChangeLog" runat="server" OnItemDataBound="rptrChangeLog_ItemDataBound">
                <HeaderTemplate>
                    <table id="TableSchedule" cellspacing="1" cellpadding="1" width="100%" border="1"
                        bordercolor="tan" bgcolor="ivory">
                        <col align="center" />
                        <col align="center" />
                        <col />
                        <col />
                        <col />
                        <col />
                        <col />
                        <col align="center" />
                        <col />
                        <col />
                        <col />
                        <tr>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Show<br />
                                On SCM</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Show<br />
                                On PM</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Change Date</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Changed By</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">AV No.</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">GPG Desc.</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Field</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Change<br />
                                Type</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Change From</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Change To</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">Comment / Reason</th>
                        </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr style="background-color: cornsilk" pvid="<%# Eval("ID")%>" clid="<%# Eval("ID")%>"
                        onmouseover="return Row_OnMouseOver()" onmouseout="return Row_OnMouseOut()" onclick="return Row_OnClick()">
                        <td class="cell" style="white-space: nowrap">
                            <asp:CheckBox ID="cbxShowOnScm" runat="server" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:CheckBox ID="cbxShowOnPm" runat="server" />
                        </td>
                        <td class="cell" style="white-space: nowrap; text-align: right">
                            <asp:Label ID="lblLastUpdDate" runat="server" Text="LastUpdDate" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblLastUpdUser" runat="server" Text="LastUpdUser" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblAvNo" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblGpgDescription" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblColumnChanged" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblChangeType" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblOldValue" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblNewValue" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblComments" runat="server" Text="Label"></asp:Label>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table>
                </FooterTemplate>
            </asp:Repeater>
        </div>
    </form>
</body>
</html>
