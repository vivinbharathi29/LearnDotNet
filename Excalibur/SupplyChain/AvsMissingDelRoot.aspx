<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SupAvsMissingDelRoot" EnableEventValidation="false" Codebehind="AvsMissingDelRoot.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>AVs Missing Deliverable Root</title>
    <link href="../../../style/general.css" rel="stylesheet" type="text/css" />
    <!-- #include file="../includes/bundleConfig.inc" -->
    <script runat="server">
    </script>
    <script type="text/javascript">
        function cmdCancel_onclick() {
            var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.closeExternalPopup();
            }
            else {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel();
                } else {
                    window.parent.close();
                }
            }
        }

        function ShowAvDetails(ProductVersionID, AvDetailID, ProductBrandID) {
            var strID;
            var url;
            url = "avFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID + "";
            //strID = window.parent.showModalDialog("avFrame.asp?Mode=edit&PVID=" + ProductVersionID + "&AVID=" + AvDetailID + "&BID=" + ProductBrandID, "", "dialogWidth:900px;dialogHeight:700px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No")
            modalDialog.open({ dialogTitle: 'AV Details', dialogURL: '' + url + '', dialogHeight: 700, dialogWidth: GetWindowSize('width'), dialogResizable: true, dialogDraggable: true });

            //window.location = window.location;
            //window.document.submit(); 
            //alert(window.parent.location);
            //alert(window.location);
            //window.location.reload();
            //window.opener.reload();
            //window.opener.refresh();
            //var avDetailId = window.document.getElementById(AvDetailID);
            //avDetailId.parentNode.parentNode.parentNode.style.display = "none";  
        }
    </script>
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <asp:Label ID="lblHeader" runat="server" Text="Please Add Deliverable Root Associations"
        Style="font-size: small; font-weight: bold; font-family: Verdana; text-align: center;
        position: absolute; top: 15px; left: 10px; width: 620px;"></asp:Label>
    <br />
    <asp:Button ID="btnCancel" runat="server" Text="Close" Style="position: absolute;
        left: 568px; width: 61px; height: 24px; top: 471px;" OnClientClick="cmdCancel_onclick()" />
    <hr style="margin-bottom: 3px; position: absolute; top: 461px; left: 10px; width: 623px" />
    <div runat="server" style="position: absolute; overflow: auto; top: 41px; left: 12px; width: 625px; height:410px">
        <asp:GridView ID="gvAVsMissingData" runat="server" GridLines="vertical" AutoGenerateColumns="False"
            CellPadding="4" ForeColor="Black" BackColor="White" BorderColor="#FFFFF0" BorderStyle="Solid"
            Style="width: 607px">
            <FooterStyle BackColor="#CCCC99" />
            <RowStyle BackColor="#F7F7DE" />
            <Columns>
                <asp:TemplateField HeaderStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <center>
                            <a href="#" id="<%#Eval("AvDetailID") %>" onclick="ShowAvDetails(<%#Eval("PVID")%>,<%#Eval("AvDetailID")%>,<%#Eval("BID")%>)">Add</a>
                        </center>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="AVNo" HeaderText="AV Number" HeaderStyle-HorizontalAlign="Center">
                </asp:BoundField>
                <asp:BoundField DataField="GPGDescription" HeaderText="GPG Description" HeaderStyle-HorizontalAlign="Center">
                </asp:BoundField>
                <asp:BoundField DataField="AvFeatureCategory" HeaderText="Feature Category" HeaderStyle-HorizontalAlign="Center">
                </asp:BoundField>
                <asp:TemplateField Visible="false">
                    <ItemTemplate>
                        <asp:Label ID="lblAvDetailID" runat="server" Text='<%#Eval("AvDetailID") %>' />
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
            <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
            <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
            <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
            <AlternatingRowStyle BackColor="#FFFFF0" />
        </asp:GridView>
        <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
    </div>
    </form>
</body>
</html>
