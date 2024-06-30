<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.Service_SpareKitMapMain2" Codebehind="SpareKitMapMain2.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script type="text/javascript" >
        var bRefreshCaller = false;
        var bPostBack = false;
        var bUnloadingVerified = false;
        var bButton = false;
        var bSubmit = false;

        function validateChange() {
            var oDirty = document.getElementById("hidDirtyFlag");
            var sDirty = oDirty.getAttribute("Value");

            if (sDirty == "true" && bButton == true) // when there is change and user clicks cancel - user is okay with canceling
            {            
                return true;
            }
            else if (sDirty == "false" && bButton == true) //when there is no change and user clicks cancel
            {
                return true;
            }
            else
            {
                if (sDirty == "false" && bButton == false)
                {
                    return true;
                }
                else
                { //when there is a change and User clicks 'x' icon in dialog
                    if (window.confirm("YOU HAVE NOT SAVED CHANGES TO THE AV MAPPING DATA.\nClick 'Cancel' and then the 'Submit' button on the page to save your changes prior to exiting."))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
        }

        function confirmExit(oSender) {
            var oDirty = document.getElementById("hidDirtyFlag");
            var sDirty = oDirty.getAttribute("Value");

            if ((!bPostBack) && (!bUnloadingVerified)) {
                var sSenderID;

                if (oSender == null) sSenderID = "body";
                else sSenderID = oSender.getAttribute("id");

                if ((sSenderID != null) && (sSenderID != undefined)) {
                    sSenderID = sSenderID.toString().toLowerCase();

                    switch (sSenderID) {

                        case "btncancel": // User is canceling the session
                            //if (sDirty == "false") {
                            //    window.returnValue = "cancel";
                            //    window.close();
                            //    return;
                            //}
                            //if (window.confirm("Exit without saving your changes?")) {
                            //    window.returnValue = "cancel";
                            //    window.close();
                            //    return;
                            //}

                            if (sDirty == "false")
                            {
                                bButton = true;
                                parent.window.parent.modalDialog.cancel(false); //process beforeClose validation
                                return;
                            }
                            else 
                            {
                                if (window.confirm("Exit without saving your changes?")) {
                                    bButton = true;
                                    parent.window.parent.modalDialog.cancel(false); //process beforeClose validation
                                } else {
                                    bButton = false;
                                }
                                return;
                            }

                            bUnloadingVerified = true; // Need to reset this flag
                            break;
                            //case "body":    // User is closing the window
                            //if (sDirty == "true") {
                            //    return "YOU HAVE NOT SAVED CHANGES TO THE AV MAPPING DATA.\nClick 'Cancel' and then the 'Submit' button on the page to save your changes prior to exiting.";
                            //} else {
                            //    window.returnValue = "cancel";
                            //}
                            //break;
                        default: // Ignore
                            break;
                    }
                }
            }
        }
     
        function window_onload() {            
            var sPageTitle = window.parent.document.title;            
            var title = parent.window.parent.modalDialog.customize('title', sPageTitle);
            //add beforeclose to modalparent 
            parent.window.parent.$("#modal_dialog").dialog({
                beforeClose: function (ev, ui) {
                    if (bSubmit == false) {
                        if (validateChange() !== true) {
                            return false;
                        } else {
                            parent.window.parent.ShowMapDetails_return('cancel');
                            return true;
                        }
                    }
                    else
                    {
                        parent.window.parent.ShowMapDetails_return('refresh');
                        return true;
                    }
                }
            });
        }
    </script>
</head>
<body id="body" runat="server" onbeforeunload="return confirmExit(null);">
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="1200" >
    </asp:ScriptManager>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:Panel ID="pnlMain" runat="server">
                <div runat="server" id="divErrors">
                    <asp:CustomValidator runat="server" ID="cValidator"></asp:CustomValidator></div>
                <div style="">
                    <asp:DataGrid ID="mapGrid" runat="server" ShowFooter="true" AutoGenerateColumns="false" CssClass="Table" >
                        <Columns>
                            <asp:TemplateColumn HeaderText="AV Category">
                                <FooterTemplate>
                                    <asp:DropDownList ID="add_ddlAvCategory" runat="server"  Width="95%" />
                                </FooterTemplate>
                                <ItemTemplate>
                                    <%#Container.DataItem("AvCategoryName")%>
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:DropDownList ID="ddlAvCategory" runat="server" />
                                </EditItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="AV Number">
                                <FooterTemplate>
                                    <asp:TextBox ID="add_AvNo" runat="server" />
                                    <asp:LinkButton runat="server" ID="lbFind" CommandName="Find" Text="Find" OnClientClick="bPostBack=true;" />
                                </FooterTemplate>
                                <ItemTemplate>
                                    <%#Container.DataItem("avno")%><asp:HiddenField ID="hidAvNo" runat="server" Value='<%#Container.DataItem("AvNo") %>' />
                                </ItemTemplate>
                                <EditItemTemplate>
                                    <asp:TextBox ID="txtAvNo" runat="server" Text='<%#Container.DataItem("AvNo") %>' />
                                    <asp:HiddenField ID="hidAvNo" runat="server" Value='<%#Container.DataItem("AvNo") %>' />
                                </EditItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn>
                                <FooterTemplate>
                                    <asp:LinkButton runat="server" CommandName="Insert" Text="Add" ID="lbAdd" OnClientClick="bPostBack=true;" />
                                </FooterTemplate>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lbDelete" runat="server" CommandName="Delete" Text="Delete" OnClientClick="bPostBack=true;" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <HeaderStyle CssClass="TableHeader" />
                    </asp:DataGrid>
                    <p style="text-align: right">
                        <asp:Button ID="btnSave" runat="server" Text="Submit" OnClientClick="bPostBack=true; bSubmit = true;" />
                        &nbsp;&nbsp;
                        <asp:Button ID="btnCancel" runat="server" Text="Cancel" OnClientClick="bPostBack=false; bUnloadingVerified = false; confirmExit(this);" />
                    </p>
                </div>
            </asp:Panel>
            <asp:Panel ID="pnlFind" runat="server" Visible="false">
                <asp:TextBox ID="txtAVFilter" Width="120px" runat="server" CssClass="TextBox" />
                <asp:Button ID="btnAVFilter" runat="server" Text="Filter" CssClass="Button" />
                <asp:Button ID="btnClear" runat="server" Text="Clear" CssClass="Button" />
                <br />
                <br />
                <asp:GridView ID="gvAVNumbers" runat="server" AutoGenerateColumns="False" CssClass="Table"
                    EmptyDataText="No AVs Found" AllowSorting="true" OnSorting="gvAVNumbers_Sorting">
                    <Columns>
                        <asp:TemplateField HeaderText="" HeaderStyle-Width="0px" HeaderStyle-Wrap="false">
                            <HeaderTemplate>
                                <center>
                                    <asp:CheckBox CssClass="CheckBox" ToolTip="Select All" ID="cbxAll" runat="server"
                                        AutoPostBack="true" OnCheckedChanged="cbxAll_Checkedchanged" />
                                </center>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <center>
                                    <asp:CheckBox CssClass="CheckBox" ID="cbxAVNumber" runat="server" />
                                    <%-- AutoPostBack="true" OnCheckedChanged="cbxAVNumber_Checkedchanged" --%>
                                </center>
                            </ItemTemplate>
                            <HeaderStyle Wrap="False" Width="0px"></HeaderStyle>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="AV Number" ItemStyle-Wrap="false" SortExpression="AvNo">
                            <ItemTemplate>
                                <asp:Label ID="lblAvNo" runat="server" Text='<%#Eval("AvNo") %>'>
                                </asp:Label>
                            </ItemTemplate>
                            <ItemStyle Wrap="False"></ItemStyle>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="GPG Description" ItemStyle-Wrap="false" SortExpression="GpgDescription">
                            <ItemTemplate>
                                <asp:Label ID="lblGpgDescription" runat="server" Text='<%#Eval("GpgDescription") %>'>
                                </asp:Label>
                            </ItemTemplate>
                            <ItemStyle Wrap="False"></ItemStyle>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="AV Feature Category" ItemStyle-Wrap="false" Visible="false">
                            <ItemTemplate>
                                <asp:Label ID="lblAvFeatureCategory" runat="server" Text='<%#Eval("AvFeatureCategory") %>'>
                                </asp:Label>
                            </ItemTemplate>
                            <ItemStyle Wrap="False"></ItemStyle>
                        </asp:TemplateField>
                    </Columns>
                    <HeaderStyle CssClass="TableHeader" />
                </asp:GridView>
                <br />
                <br />
                <asp:Label ID="lblAVs" runat="server" CssClass="Label"></asp:Label>
                <%-- AutoPostBack="true" OnCheckedChanged="cbxAVNumber_Checkedchanged" --%>
                <table border="0" cellspacing="1" cellpadding="1" align="right">
                    <tr>
                        <td>
                            <asp:Button ID="btnFindSubmit" runat="server" Text="Submit" CssClass="Button" />
                        </td>
                        <td>
                            <asp:Button ID="btnFindCancel" runat="server" Text="Cancel" CssClass="Button" />
                        </td>
                    </tr>
                </table>
            </asp:Panel>
            <asp:HiddenField ID="hidDirtyFlag" runat="server" Value="false" />
        </ContentTemplate>
    </asp:UpdatePanel>
    <asp:UpdateProgress ID="UpdateProgress1" runat="server" DisplayAfter="500">
        <ProgressTemplate>
            <div id="pageLoading">
                <div id="progressBackgroundFilter">
                </div>
                <div id="processMessage">
                    Request Processing<br />
                    <br />
                    <asp:Image ID="Image1" runat="server" ImageUrl="~/images/win8busy_light.gif" /></div>
            </div>
        </ProgressTemplate>
    </asp:UpdateProgress>
    </form>
</body>
</html>
