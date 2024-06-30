<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Query_ScheduleSummary" Codebehind="ScheduleSummary.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Schedule Summary Report</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1><asp:Label ID="lblTitle" runat="server" Text="Schedule Summary Report"></asp:Label></h1>
        
        
        <asp:GridView ID="gvScheduleSummary" runat="server" AllowSorting="true" CellPadding="5" BorderColor="tan" BorderWidth="2px">
        <HeaderStyle CssClass="TableHeader" />
        <RowStyle CssClass="Table" />
        <AlternatingRowStyle CssClass="Table" />
        </asp:GridView>
        <br />
        <asp:Label ID="lblConfidential" runat="server" Text="HP - Confidential" CssClass="Confidential"></asp:Label>
        <br />
        <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
        <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label></div>
    </form>
</body>
</html>
