<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.DCRWorkflowStatus" EnableEventValidation="false" Codebehind="DCRWorkflowStatus.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<script type="text/javascript">
    function btnTerminate_onclick() {
        if (window.confirm("Are you sure you want to terminate this workflow?")) {
            var DCRID = document.getElementById("vDCRID").value;
            var UserID = document.getElementById("vUserID").value;
            //var ApplicationRoot = document.getElementById("vApplicationRoot").value;
            //alert(ApplicationRoot);
            var strID;
            strID = window.parent.showModalDialog("DCRWorkflowTerminate.asp?DCRID=" + DCRID + "&UserID=" + UserID, "", "dialogTop:0;dialogLeft:0;dialogWidth:1px;dialogHeight:1px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
            window.close();
        }
        else {
            window.close();
        }
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <input id="vDCRID" type="hidden" runat="server" />
    <input id="vUserID" type="hidden" runat="server" />
    <input id="vApplicationRoot" type="hidden" runat="server" />
    <asp:Label ID="lblHeader" runat="server" Style="font-size: small; font-weight: bold;
        font-family: Verdana; text-align: center; position: absolute; top: 16px; left: 10px;
        width: 890px; height: 17px;"></asp:Label>
    <br />
    <asp:Label ID="lblAddComments" Text="Add Comments:" runat="server" Style="position: absolute;
        left: 17px; height: 23px; width: 113px; top: 475px; font-weight: bold; right: 1237px;"></asp:Label>
    <asp:TextBox ID="txtComments" runat="server" Style="position: absolute; left: 123px;
        height: 17px; width: 577px; top: 471px; right: 667px;" MaxLength="100"></asp:TextBox>
    <asp:Label ID="lblWorkflowType" Text="Workflow Type:" runat="server" Style="position: absolute;
        left: 17px; height: 23px; width: 117px; top: 97px; font-weight: bold; right: 1233px;"></asp:Label>
    <asp:Label ID="lblCreateDate" Text="Date Created:" runat="server" Style="position: absolute;
        left: 17px; height: 23px; width: 98px; top: 69px; font-weight: bold; right: 1252px;"></asp:Label>
    <asp:Label ID="lblCreatedBy" runat="server" Text="Created By:" Style="position: absolute;
        left: 17px; height: 23px; width: 88px; top: 40px; font-weight: bold"></asp:Label>
    <asp:Label ID="lblWorkflowTypeText" runat="server" Style="position: absolute; left: 136px;
        height: 23px; width: 659px; top: 97px;"></asp:Label>
    <asp:Label ID="lblCreateDateText" runat="server" Style="position: absolute; left: 136px;
        height: 23px; width: 141px; top: 68px;"></asp:Label>
    <asp:Label ID="lblCreatedByText" runat="server" Style="position: absolute; left: 136px;
        height: 23px; width: 185px; top: 40px; right: 1046px;"></asp:Label>
    <asp:Button ID="btnOK" runat="server" Text="OK" Style="position: absolute; left: 859px;
        width: 35px; height: 24px; top: 471px;" UseSubmitBehavior="true" />
    <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
        left: 824px; width: 71px; height: 24px; top: 471px;" UseSubmitBehavior="true" />
    <asp:Button ID="btnTerminate" runat="server" Text="Terminate" Style="position: absolute;
        left: 16px; width: 71px; height: 24px; top: 471px;" Visible="false" OnClientClick="btnTerminate_onclick();" />
    <asp:Button ID="btnSubmit" runat="server" Text="Complete" Style="position: absolute;
        left: 741px; width: 71px; height: 24px; top: 471px;" UseSubmitBehavior="true" />
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 890px;
        height: -11px;" />
    <asp:GridView ID="gvWorkflowStatus" runat="server" GridLines="vertical" AutoGenerateColumns="False"
        CellPadding="4" Style="position: absolute; top: 128px; left: 12px;" ForeColor="Black"
        BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid">
        <FooterStyle BackColor="#CCCC99" />
        <RowStyle BackColor="#F7F7DE" />
        <Columns>
            <asp:BoundField DataField="Milestone" HeaderText="Milestone" HeaderStyle-Wrap="false"
                ItemStyle-Width="300px" HeaderStyle-Width="300px">
                <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="Status" HeaderText="Status" HeaderStyle-Wrap="false" ItemStyle-Width="100px"
                HeaderStyle-Width="100px">
                <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="AssignedTo" HeaderText="Assigned To" HeaderStyle-Wrap="false"
                ItemStyle-Width="150px" HeaderStyle-Width="150px">
                <HeaderStyle HorizontalAlign="Left" Wrap="False"></HeaderStyle>
                <ItemStyle Wrap="False" />
            </asp:BoundField>
            <asp:BoundField DataField="Comments" HeaderText="Comments" HeaderStyle-Wrap="true"
                ItemStyle-Wrap="true" ItemStyle-Width="300px" HeaderStyle-Width="300px"></asp:BoundField>
        </Columns>
        <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
        <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#FFFFF0" />
    </asp:GridView>
    </form>
</body>
</html>
