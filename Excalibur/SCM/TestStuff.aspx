<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SCM_TestStuff" Codebehind="TestStuff.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <h1>TestStuff.Net Page</h1>
        <div>Server Variable: </div>
        <%: Response.Write(Request.ServerVariables("HTTP_SMUSERDN"))%>
        <br />
        <br />
        <div>Logon_User: </div>
        <%: Response.Write(Request.ServerVariables("LOGON_USER"))%>
        <br />
        <br />
        <div>LoggedInUser: </div>
        <%: Response.Write(Session("LoggedInUser"))%>
    </div>
    </form>
</body>
</html>
