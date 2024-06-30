<%@ Page Language="VB" AutoEventWireup="true"
    Inherits="DummyVBApp.InitialOffering" EnableEventValidation="false" ValidateRequest="false" Codebehind="InitialOffering.aspx.vb" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="AjaxControlToolkit" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<script type="text/javascript">
    function window_onload() {
        if (document.getElementById("scmLoading") != null)
            document.getElementById("scmLoading").style.display = "none";
    }

    function ExportInitialOffering(Publish) {
        if (Publish == 1) {
            if (window.confirm("Are you sure you want to publish?")) {
                var querystring = window.location.search;
                var businessID = querystring.substring(10, 11);
                location.href = "/iPulsar/ExcelExport/InitialOffering.aspx?BusinessID=" + businessID + "&Publish=" + Publish;
            }
        } else {
            var querystring = window.location.search;
            var businessID = querystring.substring(10, 11);
            location.href = "/iPulsar/ExcelExport/InitialOffering.aspx?BusinessID=" + businessID + "&Publish=" + Publish;
        }
    }

    function ExportSubassemblyReport() {
        var querystring = window.location.search;
        var businessID = querystring.substring(10, 11);
        location.href = "/iPulsar/ExcelExport/InitialOfferingSubassemblyReport.aspx?BusinessID=" + businessID;
    }

    function Legend() {
        var strID;
        strID = window.parent.showModalDialog("InitialOfferingLegendFrame.asp", "", "dialogWidth:400px;dialogHeight:180px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
    }

	/*
    function CommodityGuidanceFilter() {
        var retValue;
        retValue = window.parent.showModalDialog("CommodityGuidanceCategoryFilterFrame.asp", "", "dialogWidth:375px;dialogHeight:500px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        if (retValue != undefined) {
            var querystring = window.location.search;
            var ProgramID = querystring.substring(16,querystring.indexOf("&Category"));
            var Program = querystring.substring(querystring.indexOf("&ProgramText")+13);
            location.href = "/iPulsar/ExcelExport/CommodityGuidance.aspx?ProgramID=" + ProgramID + "&Program=" + Program + "&FeatureCategoryIDs=" + retValue;
        }
    }
	*/

    function cbxSelect_onclick(DRID) {
        var querystring = window.location.search;
        var businessID = querystring.substring(10, 11);
        var delimeter = querystring.indexOf("CurrentUser=");
        var currentUser = querystring.substring(delimeter + 12);
        var chkClicked = document.getElementById(window.event.srcElement.name);
        var iChecked = 0;
        if (chkClicked.checked == true) {
            iChecked = 1
        }
        var parameters = "function=AddRemoveDeliverable&DRID=" + DRID + "&Selected=" + iChecked + "&CurrentUser=" + currentUser + "&BusinessID=" + businessID;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {        // Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else {                            //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "InitialOfferingUpdate.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
    }

    function cbxProduct_onclick(PVID, BID) {
        var querystring = window.location.search;
        var delimeter = querystring.indexOf("CurrentUser=");
        var currentUser = querystring.substring(delimeter + 12);
        var chkClicked = document.getElementById(window.event.srcElement.name);
        var businessID = querystring.substring(10, 11);
        var iChecked = 0;
        if (chkClicked.checked == true) {
            iChecked = 1
        }
        var rootid = chkClicked.className;
        var tr = window.event.srcElement.parentElement.parentElement;
        var tr2 = window.event.srcElement.parentElement.parentElement.parentElement;
        var stInnerHTML = tr.childNodes(0).innerHTML;
		if (querystring.indexOf("&ProgramText") != -1) {
			var end = tr.childNodes(0).innerHTML.indexOf(">");
			var start = tr.childNodes(0).innerHTML.indexOf("class=");
			var DRID = tr.childNodes(0).innerHTML.substring(start + 6, end);
		}else{
			var end = tr.childNodes(0).innerHTML.indexOf(">");
			var start = tr.childNodes(0).innerHTML.indexOf("class=");
			var DRID = tr.childNodes(0).innerHTML.substring(start + 6, end);
		}
		if (start == -1) {
			var end = tr2.childNodes(0).innerHTML.indexOf(">");
			var start = tr2.childNodes(0).innerHTML.indexOf("class=");
			var DRID = tr2.childNodes(0).innerHTML.substring(start + 6, end);
		}
		var parameters = "function=AddRemoveProduct&DRID=" + DRID + "&PVID=" + PVID + "&BID=" + BID + "&Selected=" + iChecked + "&CurrentUser=" + currentUser;
		var request = null;
		//Initialize the AJAX variable.
		if (window.XMLHttpRequest) {        // Are we working with mozilla
			request = new XMLHttpRequest(); //Yes -- this is mozilla.
		} else {                            //Not Mozilla, must be IE
			request = new ActiveXObject("Microsoft.XMLHTTP");
		} //End setup Ajax.
		request.open("POST", "InitialOfferingUpdate.asp", false);
		request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		request.send(parameters);

		parameters = "function=AddRemoveDeliverable&DRID=" + DRID + "&Selected=1&CurrentUser=" + currentUser + "&BusinessID=" + businessID;
		request = null;
		//Initialize the AJAX variable.
		if (window.XMLHttpRequest) {        // Are we working with mozilla
			request = new XMLHttpRequest(); //Yes -- this is mozilla.
		} else {                            //Not Mozilla, must be IE
			request = new ActiveXObject("Microsoft.XMLHTTP");
		} //End setup Ajax.
		request.open("POST", "InitialOfferingUpdate.asp", false);
		request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		request.send(parameters);
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../../style/general.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        /* A scrolable div */.GridViewContainer
        {
            overflow: auto;
        }
        /* to freeze column cells and its respecitve header*/.FrozenCell
        {
            background-color: #DDDDDD;
            font-family: Verdana;
            font-size: xx-small;
            color: Black;
            position: relative;
            cursor: default;
            left: expression(document.getElementById("GridViewContainer").scrollLeft-2);
        }
        /* for freezing column header*/.FrozenHeader
        {
            /*background-color: #6B696B;*/
            font-family: Verdana;
            font-size: xx-small;
            position: relative;
            cursor: default;
            top: expression(document.getElementById("GridViewContainer").scrollTop-2);
            z-index: 10;
        }
        /*for the locked columns header to stay on top*/.FrozenHeader.locked
        {
            z-index: 99;
        }
    </style>
</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" EnablePartialRendering="true" runat="server">
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
        <ContentTemplate>
            <table style="position: absolute; top: -3px; width: 100%">
                <tr>
                    <td>
                        <asp:RadioButtonList ID="rbStatus" runat="server" RepeatDirection="Horizontal" Style="font-family: Verdana;
                            font-size: xx-small" AutoPostBack="true">
                            <asp:ListItem Selected="True" Value="0">All</asp:ListItem>
                            <asp:ListItem Value="1">Selected</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                    <td>
                        <asp:LinkButton ID="lbSubmitAVChanges" runat="server" Text="Submit AV Changes"></asp:LinkButton>
                    </td>
                    <td>
                        <a id="lbSubassemblyReport" runat="server" href="#" onclick="ExportSubassemblyReport();">Export Subassembly Report</a>
                    </td>
                    <td>
                        <a id="lbExport0" runat="server" href="#" onclick="ExportInitialOffering(0);">Export Initial Offering</a>
                    </td>
                    <td>
                        <a href="#" runat="server" id="lbExport" onclick="ExportInitialOffering(1);">Publish
                            Initial Offering</a>
                    </td>
                    <td>
                        <!--<a href="#" runat="server" id="lbCommodityGuidance" onclick="CommodityGuidanceFilter();"></a>-->
                    </td>
                    <td>
                        <a href="#" onclick="Legend();">Legend</a>
                    </td>
                    <td>
                        <asp:TextBox runat="server" ID="lblHeader" Style="font-size: x-small; font-family: verdana;
                            text-align: right; width: 250px" Wrap="false" TextMode="SingleLine" BorderStyle="none"
                            BackColor="transparent" ReadOnly="true"></asp:TextBox>
                    </td>
                </tr>
            </table>
            <br />
            <div id="GridViewContainer" class="GridViewContainer" style="width: 100%; height: 665px;">
                <asp:GridView ID="gvIO" runat="server" GridLines="vertical" AutoGenerateColumns="false"
                    CellPadding="4" ForeColor="Black" BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid">
                    <FooterStyle BackColor="#CCCC99" />
                    <RowStyle BackColor="#F7F7DE" />
                    <Columns>
                        <asp:TemplateField HeaderText="Select" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center"
                            ItemStyle-CssClass="FrozenCell" HeaderStyle-CssClass="FrozenCell">
                            <ItemTemplate>
                                <asp:CheckBox ID="cbxSelect" runat="server" CssClass='<%#Eval("DRID")%>' onclick='<%# String.Format("cbxSelect_onclick({0});", Container.DataItem("DRID")) %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Deliverable Description" HeaderStyle-HorizontalAlign="Center"
                            HeaderStyle-Wrap="false" ItemStyle-CssClass="FrozenCell" HeaderStyle-CssClass="FrozenCell">
                            <ItemTemplate>
                                <asp:Label ID="lblDelDescr" runat="server" CssClass='<%#Eval("DRID")%>' Text='<%#Eval("Name")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="LTFAVNo" HeaderText="LTF AV No." HeaderStyle-HorizontalAlign="Center"
                            HeaderStyle-Wrap="false" ItemStyle-CssClass="FrozenCell" HeaderStyle-CssClass="FrozenCell">
                        </asp:BoundField>
                        <asp:BoundField DataField="LTFSubassemblyNo" HeaderText="LTF SA No." HeaderStyle-HorizontalAlign="Center"
                            HeaderStyle-Wrap="false" ItemStyle-CssClass="FrozenCell" HeaderStyle-CssClass="FrozenCell">
                        </asp:BoundField>
                        <asp:TemplateField Visible="false">
                            <ItemTemplate>
                                <asp:Label ID="lblDRID" runat="server" Text='<%#Eval("DRID")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField Visible="false">
                            <ItemTemplate>
                                <asp:Label ID="lblChangeStatus" runat="server" Text='<%#Eval("InitialOfferingChangeStatus")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField Visible="false">
                            <ItemTemplate>
                                <asp:Label ID="lblStatus" runat="server" Text='<%#Eval("InitialOfferingStatus")%>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                    <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
                    <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
                    <HeaderStyle CssClass="FrozenHeader" BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="#FFFFF0" />
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="100">
        <ProgressTemplate>
            <div style="text-align: center;">
                <span style="font: bold small verdana; color: #696969; vertical-align: top;">
                    <asp:Image ID="Image1" runat="server" ImageUrl="images/loading19.gif" />
                    Loading...</span>
            </div>
        </ProgressTemplate>
    </asp:UpdateProgress>
    </form>
</body>
</html>
