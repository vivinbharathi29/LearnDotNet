<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SCM_RegionalAVSelectorTool" Codebehind="RegionalAVSelectionTool.aspx.vb" %>
<%@ OutputCache Location="None" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Regional AV Selection Tool</title>
    <link href="../../style/general.css" rel="stylesheet" type="text/css" />
    <%--<link rel="stylesheet" type="text/css" href="<%= AppRoot %>/scm/style.css" />--%>
    <link rel="stylesheet" type="text/css" href="../../scm/style.css" />
    <%--<link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />--%>
    
    <style type="text/css">
        .BottomBorder
        {
            border-bottom-width: 2px;
            border-bottom-color: rgb(120,120,120);
            border-bottom-style: solid;
        }
        
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
    
<%--<script type="text/javascript" src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js"></script>

<script type="text/javascript" src="<%= AppRoot %>/includes/client/popup.js"></script>

<script type="text/javascript">
document.getElementById("Stuff").

</script>--%>

<%--<script type="text/javascript">
    function FilterAVRegionalProdBrandSelection() {
        var Categories = document.getElementById("hidSCMRegionFilter")
        //alert("Categories: " + Categories.value);
        var url = "FilterByRegionalAVSelectorFrame.asp" //?BusinessID=" + BusinessID + "&BID=" + BID + "&UserID=" + UserID + "&PVID=" + PVID + "&Categories=" + Categories.value;

        var retValue;
        retValue = window.parent.showModalDialog(url, "", "dialogWidth:315px;dialogHeight:475px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
        if (retValue != undefined)
        {
            //alert(PVID);
            window.location.replace("/scm/regionalAVSelectionTool.aspx" //?ID=" + PVID + "&BID=" + BID + "&SCMCategories=" + retValue);
        }
    }
</script>--%>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <div>
<%--    <p>&nbsp;<br />
    <br />
    Get with Kenneth to Add a trigger to the obseleting section of Global SCM to obselete the related
    regional dates records as well. And also a trigger to delete regional dates records when the user
    deletes data from the scm.</p>--%>
        <div style="border: thin solid rgb(128,128,128); background-color: rgb(220,220,220);width: 100%;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 20%; text-align: right; font-weight: bold; color: #000090;">
                        Current Region:&nbsp;
                    </td>
                    <td style="width: 80%">
                        <asp:Label ID="lblRegion" runat="server" Text="EMEA"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td style="width: 20%; text-align: right; font-weight: bold; color: #000090;">
                        Current Platforms:&nbsp;
                    </td>
                    <td style="width: 80%">
                        <asp:Label ID="CurrPlats" runat="server" Text="Outfield 1.0 > hp p | Clash 1.0 > hp p | Clash 1.0 > hp w | Aerosmith 1.0 > hp w | Foreigner 1.0 > hp m | Outfield 1.0 > hp p"></asp:Label>
                    </td>
                </tr>
                
                <!--
                <tr>
                    <td style="width: 20%; text-align: right; font-weight: bold; color: #000090;">
                        Current Platforms2:&nbsp;
                    </td>
                    <td style="width: 80%">
                        Outfield 1.0 > hp p <strong>|</strong> Clash 1.0 > hp p <strong>|</strong> Clash 1.0 > hp w <strong>|</strong> Aerosmith 1.0 > hp w <strong>|</strong> Foreigner 1.0 > hp m <strong>|</strong> Outfield 1.0 > hp p
                    </td>
                </tr>
                -->
            </table>
        </div>
        <br />
        <asp:ScriptManager ID="ScriptManager2" runat="server">
        </asp:ScriptManager>
             <asp:UpdatePanel ID="UpdatePanel2" runat="server" ChildrenAsTriggers="true"><ContentTemplate>
        <%--<asp:Button ID="btnSaveChanges" runat="server" Text="Save Changes" />--%>
        
                 <asp:LinkButton ID="lkbChangeSelection" visible="False" runat="server">Change Selection</asp:LinkButton>
        
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                 <br />
            <asp:CheckBox ID="chkAll" runat="server" Text="Select/UnSelect All" 
                TextAlign="Left" AutoPostBack="true" Visible="true" />
        <br />
        <!--<div style='position: absolute; text-align: center; vertical-align: middle; top: 500px; left: 500px; background-color: White; background-repeat: no-repeat; color: Red; font-weight: bold;'><br /> The Regional CPL Blind Date you entered does not fit within the range of the Global dates!<br /><br /></div>-->
        <br />
        <asp:Label ID="lblErrorMessage" Visible="false" style="font-weight: bolder; color: Red;" runat="server" Text=""></asp:Label>
        <asp:Label ID="MessageLabel" Visible="true" style="font-weight: bolder; color: Red;" runat="server">Loading Please Wait...</asp:Label>
            <br />
            <br />
            <asp:Label ID="lblRecCountMsg" runat="server" Font-Bold="True" 
                Text="Supported Configuration Matrix" Font-Size="X-Large"></asp:Label>
            &nbsp;
            <asp:Label ID="lblRecCountMsg1" runat="server" 
                Text="( 0 Avs Displayed )" Font-Size="X-Large"></asp:Label>
            <asp:Label ID="lblRecCountMsg2" runat="server" 
                style="font-weight: bolder; color: Red;" Font-Size="X-Large">Region</asp:Label><br />
          </ContentTemplate>        </asp:UpdatePanel>
            <div  runat="server" id="GridViewContainer" class="GridViewContainer" style="width: 97%; height: 640px;">
             <asp:UpdatePanel ID="UpdatePanel1" runat="server" ChildrenAsTriggers="true"><ContentTemplate>
              <asp:GridView ID="gvRegAVSelToolGrid" runat="server" BorderStyle="Solid" 
                AutoGenerateColumns="False" Font-Names="Verdana" Font-Size="8pt" 
                BorderWidth="2px" BorderColor="#787878">
                <Columns>
                    <asp:TemplateField HeaderText="Status" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblStatus" runat="server"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="AV Detail ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblAVDetailID" runat="server" Text='<%# Bind("AvDetailID") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="AV Regional Dates ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblAVRegionalDatesID" runat="server" Text='<%# Bind("AvRegionalDatesID") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>
                    <asp:TemplateField HeaderText="Feature Category">
                        <ItemTemplate>
                            <asp:Label ID="lblFeatCat" runat="server" Text='<%# Bind("AvFeatureCategory") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle BackColor="Yellow" CssClass="BottomBorder" />
                        <ItemStyle Width="95px" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="SELECT" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                         <ItemTemplate>
                                <asp:CheckBox ID="chkSelect" runat="server" Checked='<%# Bind("Select") %>'
                                    oncheckedchanged="chkSelect_CheckedChanged" AutoPostBack="True" />
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="70px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Cat Field" Visible="false" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                         <ItemTemplate>
                                <%--<asp:CheckBox ID="chkCatField" runat="server" Checked='<%# Bind("CatField") %>' />--%>
                                <asp:Label ID="chkCatField" Width="75px" runat="server" Visible="true" Text='<% Bind("CatField") %>' />
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="70px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Regional SA Date">
                         <ItemTemplate>
                                <asp:Label ID="lblRegCPLBlind" Width="75px" runat="server" Visible="false" Text='<%# Bind("RegCPLBlind") %>'></asp:Label>
                                <asp:TextBox ID="txtRegCPLBlind" Width="75px" runat="server" 
                                    text='<%# Bind("RegCPLBlind") %>' AutoPostBack="True" 
                                    ontextchanged="txtRegCPLBlind_TextChanged"></asp:TextBox>
                                    
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="75px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Regional EM Date">
                         <ItemTemplate>
                                <asp:Label ID="lblRegRASDics" Width="75px" runat="server" Visible="false" Text='<%# Bind("RegRASDics") %>'></asp:Label>
                                <asp:TextBox ID="txtRegRASDics" Width="75px" runat="server" 
                                    text='<%# Bind("RegRASDics") %>' 
                                    ontextchanged="txtRegRASDics_TextChanged" AutoPostBack="True"></asp:TextBox>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="75px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="AV#">
                         <ItemTemplate>
                                <asp:Label ID="lblAVNo" runat="server" Text='<%# Bind("AVNo") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Marketing Description (40 Char GPSy)">
                         <ItemTemplate>
                                <asp:Label ID="lblGPGDesc" runat="server" Text='<%# Bind("GPGDescription") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="300px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Product / Brand">
                         <ItemTemplate>
                                <asp:Label ID="lblProdName" runat="server" Text='<%# Bind("ShortProdName") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center" Width="125px"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Global SA Date">
                         <ItemTemplate>
                                <asp:Label ID="lblGlobalCPLBlind" runat="server" Text='<%# Bind("CBLBlindDate") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Global EM Date">
                         <ItemTemplate>
                                <asp:Label ID="lblGlobalRASDisc" runat="server" Text='<%# Bind("RASDiscontinueDate") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Global Series Config EOL">
                         <ItemTemplate>
                                <asp:Label ID="lblGlobalSeriesConfigEOL" runat="server" Text='<%# Bind("GSEndDate") %>'></asp:Label>
                         </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                         <ItemStyle HorizontalAlign="Center" Width="100px" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Configuration Rules">
                         <ItemTemplate>
                                <asp:Label ID="lblConfigRules" runat="server" Text='<%# Bind("ConfigRules") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle CssClass="BottomBorder" />
                         <ItemStyle Width="300px" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Category Rules" Visible="False">
                         <ItemTemplate>
                                <asp:Label ID="lblCategoryRules" runat="server" Text='<%# Bind("CategoryRules") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle CssClass="BottomBorder" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="IDS-SKUS">
                         <ItemTemplate>
                                <asp:Label ID="lblIDS_SKUS" runat="server" Text='<%# Bind("IdsSkus_YN") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="IDS-CTO">
                         <ItemTemplate>
                                <asp:Label ID="lblIDS_CTO" runat="server" Text='<%# Bind("IdsCto_YN") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="RCTO-SKUS">
                         <ItemTemplate>
                                <asp:Label ID="lblRCTO_SKUS" runat="server" Text='<%# Bind("RctoSkus_YN") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder"></HeaderStyle>
                         <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="RCTO-CTO">
                         <ItemTemplate>
                                <asp:Label ID="lblRCTO_CTO" runat="server" Text='<%# Bind("RctoCto_YN") %>'></asp:Label>
                         </ItemTemplate>
                         <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                         <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Product Brand ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblProductBrandID" runat="server" Text='<%# Bind("ProductBrandID") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Av Feature Category ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblAvFeatureCatID" runat="server" Text='<%# Bind("FeatureCategoryID") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Geo ID" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="lblGeoID" runat="server" Text='<%# Bind("GeoID") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Av Detail Product Brand ID" Visible="False">
                        <ItemTemplate>
                            <asp:Label ID="lblAvDetail_ProductBrandID" runat="server" Text='<%# Bind("MainAvDetProdBrandID") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Regional Selected" ControlStyle-Width="0px" Visible="true" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                        <ItemTemplate>
                            <asp:Label ID="chkCheckedRec" runat="server" Text='<%# Bind("CheckedRecFlag") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" Width="0px" />
                    </asp:TemplateField>
                    
                    <asp:TemplateField HeaderText="Regional Active" ControlStyle-Width="0px" Visible="true" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                        <ItemTemplate>
                                <asp:Label ID="chkRecSelected" runat="server" Text='<%# Bind("RecSelectedFlag") %>'></asp:Label>
                        </ItemTemplate>
                        <HeaderStyle HorizontalAlign="Center" BackColor="Ivory" CssClass="BottomBorder" />
                        <ItemStyle HorizontalAlign="Center" Width="0px" />
                    </asp:TemplateField>
                </Columns>
                <HeaderStyle BackColor="#FDE9D9" BorderStyle="Solid" BorderColor="#A0A0A0" 
                    BorderWidth="3px" CssClass="FrozenHeader" />
            </asp:GridView>
          </ContentTemplate>        </asp:UpdatePanel>
        </div>
    </div>
    </form>
</body>
</html>