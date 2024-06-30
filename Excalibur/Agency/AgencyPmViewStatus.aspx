<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.Agency_AgencyPmViewStatus" Codebehind="AgencyPmViewStatus.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <meta http-equiv="cache-control" content="no-cache" />
    <meta http-equiv="expires" content="0" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="pnlError" runat="server" CssClass="ui-state-error ui-corner-all" Visible="false">
            <p>
                <span class="ui-icon ui-icon-alert" style="float: left; margin-right: .3em;"></span>
                <strong>Alert:&nbsp;</strong><span id="errorText"><asp:Label ID="lblError" runat="server"
                    Text="Inavlid Request Provided"></asp:Label></span></p>
        </asp:Panel>
        <asp:Panel ID="pnlWarning" runat="server" CssClass="ui-state-highlight ui-corner-all"
            Visible="false">
            <p>
                <span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
                <strong>Notice:&nbsp;</strong><span id="warningText"><asp:Label ID="lblNorecords"
                    runat="server" Text="No Records to Display"></asp:Label></span></p>
        </asp:Panel>
        <asp:Panel ID="pnlItems" runat="server" ScrollBars="Auto" Height="250">
        
        <asp:Repeater ID="rptrItems" runat="server">
        <HeaderTemplate><ul></HeaderTemplate>
        <ItemTemplate><li><%# Container.DataItem("Name")%></li></ItemTemplate>
        <FooterTemplate></ul></FooterTemplate>
        </asp:Repeater>
       </asp:Panel>
    </div>
    </form>
</body>
</html>
