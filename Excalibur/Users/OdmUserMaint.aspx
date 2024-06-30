<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Users_OdmUserMaint" Codebehind="OdmUserMaint.aspx.vb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin: 0 0 0 0;">
    <form id="form1" runat="server">
        <div style="background-color: Blue; width: 100%; padding: 10px;">
            <span style="color: White; font-family: Verdana; font-size: large; font-weight: bold;">
                Excalibur</span>
        </div>
        <div runat="server" id="divMain" style="margin: 5px 5px 5px 5px">
            <p>
                The user, <asp:Label ID="lblUserName" runat="server" Text="User Name"></asp:Label>, has not been active in the Excalibur system for 60 days. To keep this user active click
                the Re-Activate button. To remove this users access click the Remove Access button.</p>
            <div runat="server" id="divStatus">
                <asp:Label ID="lblStatus" runat="server" Text="Label"></asp:Label></div>
            <div runat="server" id="divButtons">
            <p>
                <asp:Button ID="btnReActivate" runat="server" Text="Re-Activate" Width="150px" /></p>
            <p>
                <asp:Button ID="btnRemoveAccess" runat="server" Text="Remove Access" Width="150px" /></p></div>
        </div>
    </form>
</body>
</html>
