<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SupSCM_PMGandGPSyExport" Codebehind="PMGandGPSyExport.aspx.vb" %>

<%@ Register Assembly="eWorld.UI, Version=2.0.6.2393, Culture=neutral, PublicKeyToken=24d65337282035f2"
    Namespace="eWorld.UI" TagPrefix="ew" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        //var bClose = document.getElementById("lblHidden");
        //if (bClose.value == "True") {
            window.parent.close();
        //}
    } 
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <input id="lblHidden" type="hidden" runat="server" />
    <asp:Label ID="lblErrorMessage" runat="server" Style="position: absolute; color: Red;
        top: 20px; left: 9px; width: 517px; height: 17px; font-size: x-small; font-weight: bold;
        text-align: right"></asp:Label>
    <div runat="server" id="divBody">
        <asp:RadioButtonList ID="rbStatus" runat="server" RepeatDirection="Horizontal" Style="position: absolute;
            top: 14px; left: 90px;">
            <asp:ListItem Value="0">PMG</asp:ListItem>
            <asp:ListItem Value="1">GPSy</asp:ListItem>
        </asp:RadioButtonList>
        <asp:ListBox ID="lbProducts" runat="server" DataTextField="FullName" DataValueField="ID"
            SelectionMode="Multiple" Style="position: absolute; left: 285px; height: 340px;
            width: 233px; top: 86px"></asp:ListBox>
        <asp:Label ID="lblHeader4" runat="server" Text="Desrciptions :" Style="position: absolute;
            top: 21px; left: 10px; width: 45px; height: 27px; right: 1312px; font-weight: bold"></asp:Label>
        <asp:Label ID="lblHeader2" runat="server" Text="Products" Style="position: absolute;
            top: 64px; left: 285px; width: 260px; height: 17px; font-weight: bold"></asp:Label>
        <hr style="width: 526px; margin-left: 0px; position: absolute; top: 48px; left: 4px;" />
        <br />
        <br />
        <br />
        <br />
        <asp:Label ID="lblHeader3" runat="server" Text="Product Groups" Style="width: 260px;
            height: 17px; font-weight: bold"></asp:Label>&nbsp;
        <asp:Label ID="Label1" runat="server" Text="Click to Add (Optional)" Style="width: 260px;
            height: 17px;"></asp:Label>
        <br />
        <br />
        <asp:ListBox ID="lbCycle" runat="server" DataTextField="FullName" DataValueField="ID"
            SelectionMode="Multiple" Style="height: 340px; width: 233px;" AutoPostBack="true">
        </asp:ListBox>
    </div>
    <hr style="width: 526px; margin-left: 0px; position: absolute; top: 422px; left: 4px;" />
    <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 429px;
        width: 35px; height: 24px; top: 437px;" />
    <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
        left: 467px; width: 61px; height: 24px; top: 437px;" OnClientClick="cmdCancel_onclick();" />
    
    <%--<iframe id="my_iframe" runat="server" onload="javascript:cmdCancel_onclick();"></iframe>--%>
    </form>
</body>
</html>
