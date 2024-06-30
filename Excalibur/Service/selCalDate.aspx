<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_selCalDate" Codebehind="selCalDate.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Calendar ID="theCalendar" runat="server" Height="116px" Width="232px" onselectionchanged="theCalendar_SelectionChanged">
        </asp:Calendar>
    </div>
    </form>
</body>
</html>
