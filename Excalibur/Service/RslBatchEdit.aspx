<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_RslBatchEdit" Codebehind="RslBatchEdit.aspx.vb" %>

<%@ Register Assembly="ControlLibrary" Namespace="HPQ.CustomControls" TagPrefix="HPQ" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div style="width: 800px;">
        <table class="Table">
            <col width="50px" />
            <col width="100px" />
            <col width="300px" />
            <col width="150px" />
            <col width="50px" />
            <col width="50px" />
            <col width="50px" />
            <col width="50px" />
            <tr>
                <th>
                    <asp:CheckBox ID="cbCheckAll" runat="server" />
                </th>
                <th>
                    Spare Kit No
                </th>
                <th>
                    GPG Description
                </th>
                <th>
                    First Service Dt.
                </th>
                <th>
                    NA
                </th>
                <th>
                    LA
                </th>
                <th>
                    APJ
                </th>
                <th>
                    EMEA
                </th>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;
                </td>
                <td>
                    <hpq:datepicker id="DatePicker1" runat="server" />
                </td>
                <td>
                    <asp:CheckBox ID="cbNa" runat="server" />
                </td>
                <td>
                    <asp:CheckBox ID="cbLa" runat="server" />
                </td>
                <td>
                    <asp:CheckBox ID="cbApj" runat="server" />
                </td>
                <td>
                    <asp:CheckBox ID="cbEmea" runat="server" />
                </td>
            </tr>
        </table>
    </div>
    <div style="width: 817px; height: 600px; overflow: scroll; padding-right: 2px;">
        <asp:Repeater ID="Repeater1" runat="server">
            <HeaderTemplate>
                <table class="Table" style="width: 800px">
                    <col width="50px" />
                    <col width="100px" />
                    <col width="300px" />
                    <col width="150px" />
                    <col width="50px" />
                    <col width="50px" />
                    <col width="50px" />
                    <col width="50px" />
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td style="text-align: center;">
                        <asp:CheckBox ID="cb" runat="server" />
                    </td>
                    <td>
                        <%#DataBinder.Eval(Container.DataItem, "SpareKitNo")%>
                    </td>
                    <td>
                        <%#DataBinder.Eval(Container.DataItem, "Description")%>
                    </td>
                    <td>
                    <%#DataBinder.Eval(Container.DataItem, "FirstServiceDt")%>
                    </td>
                    <td>
                    <%#DataBinder.Eval(Container.DataItem, "GeoNa")%>
                    </td>
                    <td>
                    <%#DataBinder.Eval(Container.DataItem, "GeoLa")%>
                    </td>
                    <td>
                    <%#DataBinder.Eval(Container.DataItem, "GeoApj")%>
                    </td>
                    <td>
                    <%#DataBinder.Eval(Container.DataItem, "GeoEmea")%>
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
