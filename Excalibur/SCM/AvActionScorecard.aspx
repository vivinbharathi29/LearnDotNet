<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SCM_AvActionScorecard" Codebehind="AvActionScorecard.aspx.vb" %>

<%@ Register Assembly="eWorld.UI, Version=2.0.6.2393, Culture=neutral, PublicKeyToken=24d65337282035f2"
    Namespace="eWorld.UI" TagPrefix="ew" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
       <form id="form1" runat="server">
    <asp:RadioButtonList ID="rbStatus" runat="server" RepeatDirection="Horizontal" Style="position: absolute;
        top: 14px; left: 60px;">
        <asp:ListItem Value="0">Pending</asp:ListItem>
        <asp:ListItem Value="1">Completed</asp:ListItem>
        <asp:ListItem Value="2">All</asp:ListItem>
    </asp:RadioButtonList>
    <div id="divCalendar" runat="server" style="position: absolute; top: 65px; left: 100px;
        float: left">
        <ew:CalendarPopup PopupLocation="Bottom" ID="txtFromDt" Width="160px" runat="server"
            ImageUrl="../images/calendar.gif" Nullable="True" SelectedDate="" VisibleDate=""
            PostedDate="" UpperBoundDate="12/31/9999 23:59:59" DisableTextBoxEntry="false">
        </ew:CalendarPopup>
        <br />
        <ew:CalendarPopup PopupLocation="Bottom" Width="160px" ID="txtToDt" runat="server"
            ImageUrl="../images/calendar.gif" Nullable="True" SelectedDate="" VisibleDate=""
            PostedDate="" UpperBoundDate="12/31/9999 23:59:59" DisableTextBoxEntry="false">
        </ew:CalendarPopup>
    </div>
    <asp:ListBox ID="lbProducts" runat="server" DataTextField="FullName" DataValueField="ID"
        SelectionMode="Multiple" Style="position: absolute; left: 285px; height: 250px;
        width: 233px; top: 163px"></asp:ListBox>
    <asp:Label ID="lblHeader7" runat="server" Text="To :" Style="position: absolute;
        top: 99px; left: 10px; width: 77px; height: 17px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblHeader6" runat="server" Text="OR" Style="position: absolute; top: 80px;
        left: 313px; width: 36px; height: 17px; right: 1018px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblHeader8" runat="server" Text="days" 
           
           Style="position: absolute;
        top: 81px; left: 495px; width: 33px; height: 17px; margin-bottom: 0px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblHeader5" runat="server" Text="Created in the last" 
           
           Style="position: absolute;
        top: 80px; left: 343px; width: 118px; height: 17px; font-weight: bold; margin-bottom: 0px;"></asp:Label>
    <asp:Label ID="lblHeader9" runat="server" Text="Created From :" Style="position: absolute;
        top: 71px; left: 9px; width: 96px; height: 17px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblHeader4" runat="server" Text="Status :" Style="position: absolute;
        top: 21px; left: 10px; width: 45px; height: 27px; right: 1312px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblErrorMessage" runat="server" Style="position: absolute; color: Red;
        top: 20px; left: 285px; width: 241px; height: 17px; font-weight: bold; text-align:right"></asp:Label>
    <asp:Label ID="lblHeader2" runat="server" Text="Products" Style="position: absolute;
        top: 141px; left: 285px; width: 260px; height: 17px; font-weight: bold"></asp:Label>
    <asp:TextBox ID="txtDays" runat="server" Style="position: absolute; width: 29px;
        top: 75px; left: 454px; right: 884px;"></asp:TextBox>
    <hr style="width: 526px; margin-left: 0px; position: absolute; top: 48px; left: 4px;" />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <hr style="width: 526px;" />
    <asp:Label ID="lblHeader3" runat="server" Text="Product Groups" Style="width: 260px;
        height: 17px; font-weight: bold"></asp:Label>&nbsp;
    <asp:Label ID="Label1" runat="server" Text="Click to Add (Optional)" Style="width: 260px;
        height: 17px;"></asp:Label>
    <br />
    <br />
    <asp:ListBox ID="lbCycle" runat="server" DataTextField="FullName" DataValueField="ID"
        SelectionMode="Multiple" Style="height: 250px; width: 233px;" AutoPostBack="true">
    </asp:ListBox>
    <hr style="width: 526px; margin-left: 0px; position: absolute; top: 422px; left: 4px;" />
    <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 429px;
        width: 35px; height: 24px; top: 437px;" />
    <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
        left: 467px; width: 61px; height: 24px; top: 437px;" OnClientClick="cmdCancel_onclick()" />
    </form>
</body>
</html>

