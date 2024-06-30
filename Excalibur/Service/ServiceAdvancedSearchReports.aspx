<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_ServiceAdvancedSearchReports" Codebehind="ServiceAdvancedSearchReports.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Service Advanced Search Reports Page</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />    
   <!-- <script language="javascript" type="text/javascript">
        function validateFamilySkuToSparekits(sender, args) {
            var ProductSelected = document.getElementById("lstProducts");
            var SKUs=document.getElementById("txtSkuNumberProduct");
                alert(SKUs.value);
                args.IsValid = false;
                alert(ProductSelected);
                
               if ((document.all("lstProducts").value == "") && (document.all("txtSkuNumberProduct").value == "")) {
                  args.IsValid = false;
               } else {
                  args.IsValid = true;
               } 
            }

    </script>-->
</head>
<body>
    <form id="frmServiceAdvancedSearchReports" runat="server" >
    <div>
    <asp:Label ID="lblTitle" runat="server" Text="Service Advanced Search" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
            <br />
            <!-- FILTERS -->
            <table style="width: 100%;"  border="0">
              <!--<asp:ListItem Text="Sparekits to Products" Value="3"  Selected="True" ></asp:ListItem>-->                          
                <tr valign="bottom">
                    <td style=" padding-top: 10px; padding-bottom: 5px;" >
                        <asp:RadioButtonList runat="server" ID="rdReportType" RepeatDirection="Horizontal" AutoPostBack="true" >
                            <asp:ListItem Text="SpareKits Bom" Value="4" Selected="True"></asp:ListItem>
                            <asp:ListItem Text="Spare Kits by Category" Value="7" ></asp:ListItem>
                            <asp:ListItem Text="Av to SpareKits" Value="2" ></asp:ListItem>
                            <asp:ListItem Text="SKU to SpareKits" Value="5" ></asp:ListItem>
                            <asp:ListItem Text="Family to SKU and Sparekits" Value="6" ></asp:ListItem>                                                                                
                            <asp:ListItem Text="Used By" Value="8"></asp:ListItem>              
                            <asp:ListItem Text="RSL Change Log" Value="9"></asp:ListItem>              
                        </asp:RadioButtonList>
                    </td>      
                </tr>    
                <tr>
                    <td>
                        <asp:Button ID="btnSpareKitsPlus" runat="server" BackColor="#333333" BorderColor="#333333"
                                    BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                                    Font-Size="X-Small" ForeColor="White" Height="18px" Text="Open SKU to SpareKits Plus" 
                                    OnClientClick="window.open('../../Pulsar/Report/SKUtoSpareKits', '_blank')"/>
                    </td>
                </tr>
                <tr>
                    <td style="padding-top: 10px; padding-bottom: 10;" > 
                        <asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" />
                        <br />
                        <asp:DropDownList ID="ddlReportFormat" runat="server" >
                            <asp:ListItem Value="0" Text="HTML" />
                            <asp:ListItem Value="1" Text="Excel" />
                            <asp:ListItem Value="2" Text="Word" />
                            <asp:ListItem Value="3" Text="Csv" />
                        </asp:DropDownList>
                     </td>                     
                 </tr>
                  <tr>
                    <td style="padding-top: 10px; padding-bottom: 5;">
                        <asp:Button ID="btnReport" runat="server" BackColor="#333333" BorderColor="#333333"
                                BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                                Font-Size="X-Small" ForeColor="White" Height="18px" Text="Submit" 
                                OnClientClick="setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='ServiceAdvancedSearchReports.aspx';}, 5);"/>                         
                        <asp:Button ID="btnReset" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  CssClass="ReportButton"
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Reset" CausesValidation="false" />
                     </td>                    
                </tr>
            </table>
           <!-- END FILTERS -->
           <table>
            <tr><td></td><td></td></tr>
            <tr>
                <td colspan="2">
                    <br />
                    <asp:Label ID="lblError" runat="server" ForeColor="Red" ></asp:Label>
                </td>
            </tr>
           <asp:Panel id="pnlSpsToProducts" runat="server">
                <tr>

                    <td colspan="2">
                         <asp:RequiredFieldValidator runat="server" ID="rtxtSpsNumber" ControlToValidate="txtSpsNumber" 
                                         ErrorMessage="SpareKit Numbers: You have to write a SpareKit Number."></asp:RequiredFieldValidator>
                                         
                    </td>
                </tr>               
                <tr >
                     <td style="padding-top: 10px;">
                      <br />
                        
                       <asp:Label ID="lblSpspNumber" runat="server"  CssClass="HeaderLabel" Text="SpareKit Part Number(s)"  />
                        <br />
                        
                            <asp:TextBox ID="txtSpsNumber" runat="server" TextMode="MultiLine" Columns="50" Rows="10"></asp:TextBox>                            
                     </td>
                     <td valign="bottom"><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font></td>                     
                </tr>
                
           </asp:Panel>
           <asp:Panel id="pnlSpsBom" runat="server">               
                <tr>
                    <td colspan="2">
                       <asp:RequiredFieldValidator runat="server" ID="rtxtSPSBom" ControlToValidate="txtSPSBom" ErrorMessage="SpareKit Numbers: You have to write a SpareKit Number."></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr >
                     <td style="padding-top: 10px;">
                       <br />                        
                       <asp:Label ID="lblSpsBom" runat="server"  CssClass="HeaderLabel" Text="SpareKit Part Number(s)"  />
                        <br />
                        <asp:TextBox ID="txtSPSBom" runat="server" TextMode="MultiLine" Columns="71" Rows="15" MaxLength="1200"></asp:TextBox>                            
                     </td>
                     <td valign="bottom"><font size="1" face="verdana" color="green">&nbsp;(Max 100 Sparekits)</font></td>                     
                </tr>
           </asp:Panel>
           <asp:Panel id="pnlSpsByCategory" runat="server">
                <tr>
                    <td style="padding-top: 10px;">
                         <asp:RequiredFieldValidator runat="server" ID="rlstSpareCategory" ControlToValidate="lstSpareCategory" ErrorMessage="You have to select a Category Name."></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr >
                    <td>                   
                        <asp:Label ID="lblSPSCategory" runat="server"  CssClass="HeaderLabel" Text="Spare Kit Category" />
                        <br />
                        <asp:ListBox ID="lstSpareCategory" runat="server" Height="500" Width="150" SelectionMode="Multiple"></asp:ListBox>
                     </td>
                </tr>                           
           </asp:Panel>
           <asp:Panel id="pnlAvToSPS" runat="server">
                <tr>
                    <td colspan="2">
                        <asp:RequiredFieldValidator runat="server" ID="rdtxtAVNumber" ControlToValidate="txtAVNumber" ErrorMessage="You have to write an Av Number."></asp:RequiredFieldValidator>                        
                    </td>
                </tr>
                <tr >
                    <td style="padding-top: 10px;">
                        <br />
                        <asp:Label ID="lblAVNumber" runat="server"  CssClass="HeaderLabel" Text="AV Part Number(s)"  />
                        <br />
                        <asp:TextBox ID="txtAVNumber" runat="server"  TextMode="MultiLine" Columns="71" Rows="15" MaxLength="1200"></asp:TextBox>
                    </td>
                    <td valign="bottom"><font size="1" face="verdana" color="green">&nbsp;(Max 100 Avs)</font></td>
                </tr>                
           </asp:Panel>
           <asp:Panel id="pnlSkuToSpareKits" runat="server">
               <tr>
                 <td colspan="2">
                     <asp:RequiredFieldValidator runat="server" ID="rtxtSKUNumber" ControlToValidate="txtSKUNumber" ErrorMessage="Sku Number: You have to write a Sku Number."></asp:RequiredFieldValidator>
                </td>
               </tr>
               <tr >
                <td style="padding-top: 10px;">
                    <br />
                     <asp:Label ID="lblSKUNumebr" runat="server"  CssClass="HeaderLabel" Text="SKU Part Number(s)"  />
                    <br />
                    <asp:TextBox ID="txtSKUNumber" runat="server" TextMode="MultiLine" Columns="71" Rows="15" MaxLength="1250"></asp:TextBox>
                </td>       
                <td valign="bottom"><font size="1" face="verdana" color="green">&nbsp;(Max 100 SKUs)</font></td>             
              </tr>                
           </asp:Panel>
           <asp:Panel id="pnlFamilyToSkuToSparekits" runat="server">
               <tr>
                    <td>
                        <asp:CustomValidator ID="rCustomValtxtSkuNumberProduct" runat="server" 
                                    ControlToValidate="txtSkuNumberProduct" Display="Dynamic" 
                                    ErrorMessage="You must specify one the Product or the SKU" 
                                    OnServerValidate="cusSKUNumberProduct_ServerValidate"
                                    ValidateEmptyText="True"></asp:CustomValidator>                       
                    </td>
               </tr>
               <tr>
                     <td>
                        <asp:CheckBox runat="server" id="chkUnselectProducts" ToolTip="Unselect all Products" AutoPostBack="true" />&nbsp;
                        <asp:Label ID="lblProductos" runat="server"  CssClass="HeaderLabel" Text="Product Name" />
                        <br />
                        <asp:ListBox ID="lstProducts" runat="server" Height="500" Width="150" SelectionMode="Multiple"  ></asp:ListBox>
                    </td>   
                    <td style="padding-top: 10px;"><asp:Label ID="Label1" runat="server"  CssClass="HeaderLabel" Text="SKU Part Number(s)"  />
                        <br />
                        <asp:TextBox ID="txtSkuNumberProduct" runat="server" TextMode="MultiLine" Columns="71" Rows="15" MaxLength="1250"></asp:TextBox>
                        <font size="1" face="verdana" color="green">&nbsp;(Max 100 SKUs)</font>
                     </td>                     
                </tr>                 
            </asp:Panel>  
            <asp:Panel id="pnlUsedBy" runat="server">
                <tr>
                  <td style=" padding-top: 10px; padding-bottom: 5px;" >
                    <asp:RadioButtonList runat="server" ID="rdUsedBy" RepeatDirection="Horizontal" AutoPostBack="true">
                        <asp:ListItem Text="Family Name" Value="1"></asp:ListItem>
                        <asp:ListItem Text="Sparekits" Value="2" Selected="True" ></asp:ListItem>
                        <asp:ListItem Text="SubAssembly" Value="3"></asp:ListItem>
                        <asp:ListItem Text="Component" Value="4"></asp:ListItem>
                    </asp:RadioButtonList>                  
                  </td>                   
                </tr>
                <tr>
                     <td style="padding-top: 10px;">
                       <asp:Label ID="lblUsedBy" runat="server"  CssClass="HeaderLabel"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td >
                         <br />
                         <asp:RequiredFieldValidator runat="server" ID="rUsedByTxt" ControlToValidate="txtUsedBy" 
                                         ErrorMessage="Used By: You have to write a Number."></asp:RequiredFieldValidator>
                         <asp:RequiredFieldValidator runat="server" ID="rUsedByProduct" ControlToValidate="lstUsedByProducts" ErrorMessage="Product: You have to select a Product Name and a SKU number."></asp:RequiredFieldValidator>                         
                    </td>
                </tr>
                 <tr>
                    <td style="padding-top: 10px;">
                        <asp:TextBox ID="txtUsedBy" runat="server" TextMode="MultiLine" Columns="71" Rows="15" MaxLength="1200"></asp:TextBox>
                        <asp:ListBox ID="lstUsedByProducts" runat="server" Height="500" Width="150"></asp:ListBox>
                        <asp:Label runat="server" ID="lblMaxPartNumber"  Font-Names="Verdana" Font-Size="XX-Small" ForeColor="Green"></asp:Label>
                        <asp:Label ID="lblUsedBySku" runat="server" Text="SKU:   " CssClass="HeaderLabel" ></asp:Label>
                        <asp:TextBox ID="txtUsedBySKU" runat="server" Width="300"></asp:TextBox>                        
                    </td>
                </tr>            
            </asp:Panel>            
            <asp:Panel id="pnlRSLChangeLog" runat="server" Visible="false">                
               <tr>
                   <td>
                       <asp:RequiredFieldValidator runat="server" ID="RewRSLChangeLog" ControlToValidate="lstProductFamilies" ErrorMessage="Products: You have to select one." ></asp:RequiredFieldValidator>

                   </td>
               </tr>
               <tr>
                     <td>
                        <asp:CheckBox runat="server" id="chkProductsRSLChangeLog" ToolTip="Unselect all Products" AutoPostBack="true"  />&nbsp;

                        <asp:Label ID="lblProductFamilies" runat="server"  CssClass="HeaderLabel" Text="Product Name" />
                        <br />
                        <asp:ListBox ID="lstProductFamilies" runat="server" Height="500" Width="150" SelectionMode="Multiple"  ></asp:ListBox>
                    </td>                     
                </tr>                 
            </asp:Panel>
            <!-- Start- This report is not working -->
            <asp:Panel id="pnlAvBomReport" runat="server" Visible="false">                
                 <tr >
                     <td style="padding-top: 10px;">
                       <asp:Label ID="lblBomAvNumbers" runat="server"  CssClass="HeaderLabel" Text="AV Part Number(s)"  />
                        <br />
                       <asp:TextBox ID="txtBomAvNumbers" runat="server" TextMode="MultiLine" Columns="50" Rows="10"></asp:TextBox>
                    </td>
                    <td valign="bottom"><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font></td>
                 </tr>
                 <tr>
                    <td colspan="2">
                         <br />
                         <asp:RequiredFieldValidator runat="server" ID="rtxtBomAvNumbers" ControlToValidate="txtBomAvNumbers" ErrorMessage="Av Number: You have to write a Av Number."></asp:RequiredFieldValidator>
                    </td>
                </tr>
           </asp:Panel>
           <!-- End - This report is not working -->
                <tr>
                   <td align="right" style="padding-top: 20px;">
                      
                          
                            
                   </td>
                   <td>&nbsp;</td>
               </tr>     
          </table> 
    </div>
    </form>
</body>
</html>
