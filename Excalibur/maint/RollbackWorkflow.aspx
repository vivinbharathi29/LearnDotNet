<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.maint_Default" Codebehind="RollbackWorkflow.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="frmMain" runat="server">
    <div>Enter the ID Number of the Deliverable to Rollback:<br />
        <asp:TextBox ID="txtRollback" runat="server"></asp:TextBox><br />
        <asp:Button ID="cmdRollback" runat="server" Text="Rollback" />
    </div>
    </form>
</body>
</html>
