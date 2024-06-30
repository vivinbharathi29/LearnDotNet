<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Query_ScheduleChanges" EnableEventValidation="false" ViewStateEncryptionMode="never" Codebehind="ScheduleChanges.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
    <title>Schedule Change Report</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Repeater ID="rptrHistoryDetails" runat="server">
        <HeaderTemplate><table style="width:100%;border-collapse:collapse;"></HeaderTemplate>
        <FooterTemplate></table></FooterTemplate>
        </asp:Repeater>
        <br />
        <asp:Label ID="lblConfidential" runat="server" Text="HP - Confidential" CssClass="Confidential"></asp:Label>
        <br />
        <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
        <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>
    </div>
    </form>
</body>
</html>
