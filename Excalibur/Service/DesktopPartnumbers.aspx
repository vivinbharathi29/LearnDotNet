<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_DesktopPartnumbers" Codebehind="DesktopPartnumbers.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Desktop PartNumbers</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />  
    <link href="../service/style.css" type="text/css" rel="stylesheet" />
    <link href="../service/sample.css" type="text/css" rel="stylesheet" />
    <script src="../_ScriptLibrary/jsrsClient.js" type="text/javascript"></script>
    <script src="../includes/client/popup.js" type="text/javascript"></script>    
</head>
<body>
    <form id="frmDesktopPartnumbers" runat="server">
    <div>
        <table  cellspacing="0" cellpadding="0" width="90%" align="center" border="0"  >
            <tr>
                <td style="padding-top: 10px; padding-bottom: 10; width:10%;" > 
                    <asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" />
                    <br />
                    <asp:DropDownList ID="ddlReportFormat" runat="server" >
                        <asp:ListItem Value="0" Text="HTML" />
                        <asp:ListItem Value="1" Text="Excel" />                            
                    </asp:DropDownList>
                    &nbsp;&nbsp;
                    <asp:Button ID="btnReport" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Submit"/>
                 </td>   
                <td>
                    <asp:HiddenField runat="server" id="txtComparisonDate" />
                     <input type="hidden" id="ProductVersionId" name="ProductVersionId" runat="server" />
                     <asp:Label ID="lblTitle" runat="server" Text="Desktop PartNumbers" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
                    &nbsp;&nbsp;&nbsp;<a href="/iPulsar/ExcelExport/ServiceRsl.aspx?ProductVersionId=<%=PVID%>">Export RSL</a>
                  
                </td>
                <td align="center">
                    <br />
                    <asp:Panel ID ="pnlUploadLinkAvSparekits" runat="server">
                        <asp:Label ID="lblFileToUpload" runat="server" CssClass="HeaderLabel" Text="Link Av to Sparekits:" />&nbsp;&nbsp;
                        <asp:FileUpload runat="server" ID="flServiceLinkAvSparekits" Width="200" ToolTip="Press to select a Link Av To Sparekits file to upload" BackColor="White"  />
                        <asp:Button ID="bntUploadFile" runat="server" BackColor="#333333" BorderColor="#333333" BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  Font-Size="X-Small" ForeColor="White" Height="18px" Text="Upload File" />
                        <br />
                        <asp:Label runat="server" ID="lblErrorMsg"  ForeColor="Red" Font-Size="Medium"></asp:Label>            
                    </asp:Panel> 
                </td>
            </tr>        
        </table>                              
         <br /><br />
         <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 90px" Runat="server" Width="99%">
            <table  cellspacing="0" cellpadding="0" width="80%" align="center" border="0"  >
                <tr>
                    <td>
                        <asp:GridView runat="server" ID="gvDesktopPartnumbers" Width="100%" AllowPaging="true" PageSize="50"
                           AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                             <asp:boundfield datafield="servicefamilypn" headertext="ServiceFamilypn" SortExpression="servicefamilypn" Visible="false" ItemStyle-Wrap="true"  />                                 
                             <asp:HyperLinkField DataTextField="spskitpn"  Target="_blank" HeaderText="Sps. PartNumber" SortExpression="spskitpn" ItemStyle-HorizontalAlign="Center"
                                                            DataNavigateUrlFormatString="../Service/DesktopPartnumbersDetails.aspx?servicefamilypn={0}&spskitpn={1}&spskitdescription={2}" DataNavigateUrlFields="ServiceFamilyPn,spskitpn,spskitdescription" />                            
                             <asp:boundfield datafield="spskitpn" headertext="spskitpn" Visible="false" />                                                                                           
							 <asp:boundfield datafield="SparekitCategory" headertext="Sps. Category" SortExpression="SparekitCategory"  ItemStyle-Wrap="true"  />                                 
							 <asp:boundfield datafield="spskitdescription" headertext="Sps. Description" SortExpression="spskitdescription" ItemStyle-Wrap="true"  />                                 
                             <asp:boundfield datafield="CsrLEvel" headertext="CSR Level" SortExpression="CsrLEvel"  ItemStyle-Wrap="true"  />                                 
                             <asp:boundfield datafield="disposition" headertext="Disposition" SortExpression="disposition"  ItemStyle-Wrap="true"  />                                 
                             <asp:boundfield datafield="WarrantyTier" headertext="Warranty Tier" SortExpression="WarrantyTier"  ItemStyle-Wrap="true"  />                                 
                             <asp:boundfield datafield="LocalStockAdvice" headertext="Local Stock Advice" SortExpression="LocalStockAdvice"  ItemStyle-Wrap="true"  />                                 
                             <asp:boundfield datafield="GeoNa" headertext="GeoNa" SortExpression="GeoNa"  ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  />  
                             <asp:boundfield datafield="GEoLA" headertext="GeoLA" SortExpression="GEoLA"  ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center" />  
                             <asp:boundfield datafield="GeoAPJ" headertext="GeoAPJ" SortExpression="GeoAPJ"  ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  />  
                             <asp:boundfield datafield="GeoEmea" headertext="GeoEmea" SortExpression="GeoEmea"  ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  />                               
                             <asp:boundfield datafield="Supplier" headertext="Supplier" SortExpression="Supplier"  ItemStyle-Wrap="true" />                              
                             <asp:HyperLinkField  Text="Map AV"  SortExpression="SparekitCreated" Target="_blank" HeaderText="Link to AV " ItemStyle-HorizontalAlign="Center"
                                                            DataNavigateUrlFormatString="../Service/SpareKitMapMain2.aspx?PVID={0}&SKID={1}&PBID={2}&MapId={3}&servicefamilypn={4}" 
                                                              DataNavigateUrlFields="ProductVersionID,SparekitID,ProductBrandID,skmapid,servicefamilypn" />                                         
                             <asp:boundfield datafield="AvNumber"  headertext="AvNumber" SortExpression="AvNumber" />
                             <asp:boundfield datafield="SparekitID" Visible="false" />
                             <asp:boundfield datafield="ProductVersionID" Visible="false" />
                             <asp:boundfield datafield="ProductBrandID" Visible="false" />                             
                          </Columns>
                       </asp:GridView>
                       <asp:GridView runat="server" ID="gvDesktopPartnumbersToExport"  Width="100%" AutoGenerateColumns="false"
                           GridLines="Both"  ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                              <asp:boundfield datafield="SparekitCategory" headertext="Sps. Category" ItemStyle-Wrap="true"  />   
                              <asp:boundfield datafield="spskitpn"  headertext="SparekitNumber" />
                              <asp:boundfield datafield="spskitdescription" headertext="spskitdescription" ItemStyle-Wrap="true"  />                                 
                              <asp:boundfield datafield="Supplier" headertext="Supplier" ItemStyle-Wrap="true" />                                                            
                              <asp:boundfield datafield="GeoNaNum" headertext="GeoNa" ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  />  
                              <asp:boundfield datafield="GEoLANum" headertext="GeoLA" ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center" />  
                              <asp:boundfield datafield="GeoAPJNum" headertext="GeoAPJ" ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  />  
                              <asp:boundfield datafield="GeoEmeaNum" headertext="GeoEmea" ItemStyle-Wrap="true" ItemStyle-HorizontalAlign="Center"  /> 
                               <asp:boundfield datafield="CsrLEvelID" headertext="CSRLevelID" ItemStyle-Wrap="true"  />                
                              <asp:boundfield datafield="CsrLEvel" headertext="CSRLevel" ItemStyle-Wrap="true"  />                
                              <asp:boundfield datafield="dispositionID" headertext="DispositionID" ItemStyle-Wrap="true"  />                           
                              <asp:boundfield datafield="disposition" headertext="Disposition" ItemStyle-Wrap="true"  />          
                              <asp:boundfield datafield="WarrantyTierID" headertext="WarrantyTierID" ItemStyle-Wrap="true"  />                               
                              <asp:boundfield datafield="WarrantyTier" headertext="WarrantyTier" ItemStyle-Wrap="true"  />        
                              <asp:boundfield datafield="LocalStockAdviceID" headertext="LocalStockAdviceID" ItemStyle-Wrap="true"  />                         
                              <asp:boundfield datafield="LocalStockAdvice" headertext="LocalStockAdvice" ItemStyle-Wrap="true"  />                              
                              <asp:boundfield datafield="AvNumber"  headertext="AvNumber" />
                          </Columns>
                       </asp:GridView>
                    </td>
                </tr>
            </table>
        </asp:panel>
        <!-- NO DATA PANEL -->
        <asp:panel id="pnlNoData" Runat="server">
             <table style="LEFT: 10px; POSITION: absolute; TOP: 200px; HEIGHT: 30px" cellspacing="0" cellpadding="0" width="98%" border="0">
                <tr style="width:100%;">
	                <td align="center" width="100%">
                        <asp:Label ID="msgSearchNoData" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>
        </asp:panel>
        <!-- END NO DATA PANEL -->	           
    </div>

    </form>
    
</body>
</html>
