<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.DCRWorkflow"
    EnableEventValidation="false" Codebehind="DCRWorkflow.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<script type="text/javascript">
    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            window.parent.close();
        }
    }
    function ValidateDates()
    {
        var ddlWorkflow = document.getElementById("ddlWorkflow");
        var workflowvalue = ddlWorkflow.options[ddlWorkflow.selectedIndex].value;
        if (document.getElementById("hdnIsPulsarProduct").value == "1") {
            if (workflowvalue == "7" && document.getElementById("hdnRTPDate").value == "" && document.getElementById("hdnEMDate").value == "") {
                alert("You must enter either the RTP or the EM date when adding the workflow of 'Change Product Life Cycle Dates'");
                return false;
            }
            if (document.getElementById("hdnRTPDate").value != "" && !isDate(document.getElementById("hdnRTPDate").value)) {
                alert("You must enter a valid date format in the RTP date field");
                return false;
            }
            if (document.getElementById("hdnEMDate").value != "" && !isDate(document.getElementById("hdnEMDate").value)) {
                alert("You must enter a valid date format in the EM date field");
                return false;
            }
        }
        return true;
    }
    function isDate(txtDate) {
        var currVal = txtDate;
        if (currVal == '')
            return false;

        var rxDatePattern = /^(\d{1,2})(\/|-)(\d{1,2})(\/|-)(\d{4})$/;
        var dtArray = currVal.match(rxDatePattern); // is format OK?

        if (dtArray == null)
            return false;

        dtMonth = dtArray[1];
        dtDay = dtArray[3];
        dtYear = dtArray[5];

        if (dtMonth < 1 || dtMonth > 12)
            return false;
        else if (dtDay < 1 || dtDay > 31)
            return false;
        else if ((dtMonth == 4 || dtMonth == 6 || dtMonth == 9 || dtMonth == 11) && dtDay == 31)
            return false;
        else if (dtMonth == 2) {
            var isleap = (dtYear % 4 == 0 && (dtYear % 100 != 0 || dtYear % 400 == 0));
            if (dtDay > 29 || (dtDay == 29 && !isleap))
                return false;
        }

        return true;
    }
    // function btnSubmit_Click() {
    //     var btnSubmit = document.getElementById('<%=btnSubmit.ClientID %>');
    //     var btnCancel = document.getElementById('<%=btnCancel.ClientID %>');
    //     btnSubmit.disabled = true;
    //     btnCancel.disabled = true;
    // }

</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>DCR Workflow</title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label ID="lblHeader" runat="server" Text="Please Select A Change Request Workflow" Style="font-size: small;
        font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
        top: 15px; left: 10px; width: 620px;"></asp:Label>
    <br />
    <asp:DropDownList ID="ddlWorkflow" runat="server" Style="position: absolute; left: 161px;
        height: 23px; width: 319px; top: 45px;" DataTextField="Name" DataValueField="ID"
        AutoPostBack="true">
    </asp:DropDownList>
    <asp:Label ID="lblDescription" runat="server" Style="position: absolute; left: 17px;
        height: 23px; width: 609px; top: 85px; text-align: center"></asp:Label>
    <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 527px;
        width: 35px; height: 24px; top: 471px;" UseSubmitBehavior="true" OnClientClick="return ValidateDates();" />
    <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
        left: 568px; width: 61px; height: 24px; top: 471px;" OnClientClick="cmdCancel_onclick()" />
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 623px" />
    <asp:GridView ID="gvWorkflowDefintions" runat="server" GridLines="vertical" AutoGenerateColumns="False"
        CellPadding="4" ForeColor="Black" BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid"
        Style="position: absolute; top: 129px; left: 12px;">
        <FooterStyle BackColor="#CCCC99" />
        <RowStyle BackColor="#F7F7DE" />
        <Columns>
            <asp:BoundField DataField="Milestone" HeaderText="Milestone" HeaderStyle-HorizontalAlign="Left"
                ItemStyle-Width="300px" HeaderStyle-Width="300px"></asp:BoundField>
            <asp:BoundField DataField="RoleName" HeaderText="Assign To" HeaderStyle-HorizontalAlign="Left"
                ItemStyle-Width="300px" HeaderStyle-Width="300px"></asp:BoundField>
        </Columns>
        <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
        <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="#FFFFF0" />
    </asp:GridView>
        <asp:HiddenField ID="hdnRTPDate" runat="server" />
        <asp:HiddenField ID="hdnEMDate" runat="server" />
        <asp:HiddenField ID="hdnIsPulsarProduct" runat="server" />
    </form>
</body>
</html>
