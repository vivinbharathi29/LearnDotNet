<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Search_SpareKit" Codebehind="SpareKit.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Service Advanced Search Page</title>    
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
    <link href="../includes/Css/jquery-ui-1.8.22.custom.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="../includes/Css/jquery.ui.datepicker.css" />
 
    <script src="../includes/Client/jquery-1.7.1.min.js" type="text/javascript"></script>
    <script src="../includes/Client/jquery-ui-1.8.18.custom.min.js" type="text/javascript"></script>
    <script type="text/javascript" src="../includes/Client/jquery.ui.datepicker.js"></script>
      

    <script type="text/javascript">
        function ActionCell_onmouseover() {
            window.event.srcElement.style.background = "gainsboro";
            window.event.srcElement.style.cursor = "hand";
            window.event.srcElement.style.color = "black";
        }

        function ActionCell_onmouseout() {
            window.event.srcElement.style.color = "white";
            window.event.srcElement.style.background = "#333333";
        }
      function getProfileName() {
            var strNewName = window.prompt("Enter a name for the profile.", "");
            if (strNewName != null) {
                frmSKUSearch.hidProfileName.value = strNewName;
                return true;
            }
            else {
                return false;
            }
        }

        function ShareProfile() {
            var strResult;
            strResult = window.showModalDialog("../Query/ProfileShare.asp?ID=" + frmSKUSearch.hidProfileId.value, "", "dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            return false;
        }
        
        $(function() {    
            $( "#DatepickerSKUStart" ).datepicker();  
            });  
        $(function() {    
            $( "#DatepickerSKUEnd" ).datepicker();  
            });  
            
        $(function() {    
            $( "#DatepickerSPSStart" ).datepicker();  
            });  
        $(function() {    
            $( "#DatepickerSPSEnd" ).datepicker();  
            });                      
        
   </script>
</head>
<body >
    <form id="frmSKUSearch" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>

            <!-- FILTERS -->
            <table style="width: 100%;">
                <tr>
                  <td colspan="3">
                        <asp:Label runat="server" ID="lblProfile" CssClass="HeaderLabel" Text="Report Profile:" />
                        <asp:DropDownList runat="server" ID="ddlReportProfiles" AutoPostBack="True" />
                        <asp:LinkButton ID="lbAddProfile" runat="server">Add</asp:LinkButton>
                        <asp:LinkButton ID="lbUpdateProfile" runat="server" Visible="False">Update</asp:LinkButton>
                        <asp:LinkButton ID="lbDeleteProfile" runat="server" Visible="False">Delete</asp:LinkButton>
                        <asp:LinkButton ID="lbRenameProfile" runat="server" Visible="False">Rename</asp:LinkButton>
                        <asp:LinkButton ID="lbShareProfile" runat="server" Visible="False" >Share</asp:LinkButton>
                        <asp:LinkButton ID="lbRemoveProfile" runat="server" Visible="False">Remove</asp:LinkButton>
                        <asp:Label ID="lblProfileOwnerHdr" runat="server" Visible="false" Text="Profile Owner:" CssClass="ProfileOwnerHdr" />
                        <asp:Label ID="lblProfileOwnerName" runat="server" Visible="false" CssClass="ProfileOwnerName" />
                        <asp:HiddenField ID="hidProfileName" Value="hidProfileName"  runat="server" />
                        <asp:HiddenField ID="hidProfileId" Value="hidProfileId" runat="server" />
                        <asp:HiddenField ID="hidCategories" runat="server" />
                        <asp:HiddenField ID="hidProductNames" runat="server" />
                        <asp:HiddenField ID="hidSkuNumber" runat="server" />
                        <asp:HiddenField ID="hidOSSP" runat="server" />
                        <asp:HiddenField ID="hidServiceFamPartNum" runat="server" />    
                             
                    </td>
                </tr>
                  <tr>
                    <td colspan="3" style="padding-top: 10px; padding-bottom: 5px">
                        <asp:Button ID="btnSummaryReport" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" CssClass="ReportButton"
                            Font-Size="X-Small" ForeColor="White" Height="18px" 
                            OnClientClick="window.document.forms[0].target='_blank'; setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='SpareKit.aspx';}, 5);"
                            PostBackUrl="SpareKitSummary.aspx" Text="Summary Report" />
                        <asp:Button ID="btnServiceBomReport" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  CssClass="ReportButton"
                            Font-Size="X-Small" ForeColor="White" Height="18px" 
                            OnClientClick="window.document.forms[0].target='_blank'; setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='SpareKit.aspx';}, 5);"
                            PostBackUrl="ServiceBomReport.aspx" Text="BOM Report" />
                         <asp:Button ID="btnReset" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  CssClass="ReportButton"
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Reset" />
                    </td>
                </tr>
                <tr>
                    <td colspan="3" style="padding-top: 5px; padding-bottom: 5px">
                        <asp:Label ID="lblHow" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="Use CTRL or SHIFT keys to select multiple items in lists"></asp:Label>
                    </td>
                </tr>
                <tr valign="middle">
                   <td colspan="4">
                        <asp:RadioButtonList runat="server" ID="rdProductType" AutoPostBack="true" RepeatDirection="Horizontal">
                            <asp:ListItem Text="All" Selected="true" Value="0"></asp:ListItem>
                            <asp:ListItem Text="Commercial" Value="1"></asp:ListItem>
                            <asp:ListItem Text="Consumer" Value="2" ></asp:ListItem>
                        </asp:RadioButtonList>                    
                   </td>
                </tr>
                <tr valign="middle">
                     <td style="width:15%;">
                       <asp:Label ID="lblProductName" runat="server"  CssClass="HeaderLabel" Text="Product Name" />
                       <br />
                       <asp:ListBox ID="lstProducts" runat="server" Height="250" Width="150" SelectionMode="Multiple" AutoPostBack="true"></asp:ListBox>
                       <br />
                       <asp:Label ID="lblProductOnlySummamry" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                   </td>
                    <td style="width:15%;   "  >
                       <asp:Label ID="lblSKCategory" runat="server"  CssClass="HeaderLabel" Text="Spare Kit Category" />
                       <br />
                       <asp:ListBox ID="lstSpareCategory" runat="server" Height="250" Width="150" SelectionMode="Multiple"></asp:ListBox>
                       <br />
                       <asp:Label ID="Label1" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                    </td>
                    <td style="width:15%;   "  >
                        <asp:Label ID="lblOSSP" runat="server"  CssClass="HeaderLabel" Text="OSSP" />
                        <br />
                        <asp:ListBox ID="lstOSSP" runat="server" Height="250" Width="150" SelectionMode="Multiple"></asp:ListBox>
                         <br />
                       <asp:Label ID="Label2" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                    </td>
                    <td style="width:55%;">
                        <table >
                          <tr align="right" >
                                <td><asp:Label ID="lblKMAT" runat="server"  CssClass="HeaderLabel" Text="KMAT Part Number(s)"  /></td>
                                <td><asp:TextBox ID="txtKMAT" runat="server" Width="350px"></asp:TextBox><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font></td>
                            </tr>  
                           
                            <tr align="right" >
                                <td><asp:Label ID="lblSKUNumebr" runat="server"  CssClass="HeaderLabel" Text="SKU Part Number(s)"  /></td>
                                <td>
                                    <asp:TextBox ID="txtSKUNumber" runat="server" Width="350px"></asp:TextBox><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font>
                                </td>
                            </tr>         
                          
                            <tr><td><br /></td></tr>
                            <tr>
                                <td  align="right"><asp:Label ID="lblMaxRows" runat="server" CssClass="HeaderLabel" Text="Max Rows:" /></td>
                                <td><asp:TextBox runat="server" ID="txtMaxRows" Text="2000" Width="45" ></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                                <td>
                                    <asp:CompareValidator runat="server" ControlToValidate="txtMaxRows" Operator="DataTypeCheck" Type="Integer"  ErrorMessage="Max Rows must be a number."></asp:CompareValidator>                                
                                </td>
                            </tr>
                           <tr >
                              <td  align="right"><asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" /></td>
                              <td>
                                <asp:DropDownList ID="ddlReportFormat" runat="server" >
                                    <asp:ListItem Value="0" Text="HTML" />
                                    <asp:ListItem Value="1" Text="Excel" />
                                    <asp:ListItem Value="2" Text="Word" />
                                </asp:DropDownList>
                             </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr><td><br /><br /></td></tr>             
                <tr >
                   <td>
                       <asp:Label ID="lblSKUGeo" runat="server"  CssClass="HeaderLabel" Text="SKU Geo" />                
                       <asp:CheckBoxList ID="chkSKUGeo" runat="server" AutoPostBack="true" 
                           RepeatDirection="Horizontal" Height="26px">
                           <asp:ListItem Selected="true" Text="All" Value="0"></asp:ListItem>
                           <asp:ListItem Text="NA" Value="1"></asp:ListItem>
                           <asp:ListItem Text="LA" Value="4"></asp:ListItem>
                           <asp:ListItem Text="APJ" Value="3"></asp:ListItem>
                           <asp:ListItem Text="EMEA" Value="2"></asp:ListItem>                           
                       </asp:CheckBoxList>
                       <asp:Label ID="Label5" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                   </td>
                   <td colspan="3">
                       <asp:Label ID="lblSpsGeo" runat="server"  CssClass="HeaderLabel" Text="Spare Kit Geo" />                
                       <asp:CheckBoxList ID="chkSpsGeo" runat="server" AutoPostBack="true" RepeatDirection="Horizontal">
                           <asp:ListItem Selected="true" Text="All" Value="0"></asp:ListItem>
                           <asp:ListItem Text="NA" Value="1"></asp:ListItem>
                           <asp:ListItem Text="LA" Value="4"></asp:ListItem>
                           <asp:ListItem Text="APJ" Value="3"></asp:ListItem>
                           <asp:ListItem Text="EMEA" Value="2"></asp:ListItem>                           
                       </asp:CheckBoxList>
                       <asp:Label ID="Label3" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>                  
                   </td>
                </tr>
                <tr>
                    <td valign="middle">
                        <!--
                        <asp:Label ID="lblSkuDate" runat="server"  CssClass="HeaderLabel" Text="SKU Date" /><br />                
                        <asp:TextBox id="DatepickerSKUStart" runat="server" MaxLength="8" Width="65"></asp:TextBox>&nbsp;&nbsp;
                        <asp:Label ID="lblSkuDateTo" runat="server"  Text="TO" />&nbsp;&nbsp;
                        <asp:TextBox id="DatepickerSKUEnd" runat="server" MaxLength="8" Width="65"></asp:TextBox>
                        -->
                    </td>
                    <td colspan="3" valign="middle">
                        <asp:Label ID="lblSpsDate" runat="server"  CssClass="HeaderLabel" Text="Spare Kit Update Date" /><br />                
                        <asp:TextBox id="DatepickerSPSStart" runat="server" MaxLength="8" Width="65"></asp:TextBox>&nbsp;&nbsp;
                        <asp:Label ID="lblSPSDateTo" runat="server"  Text="TO" />&nbsp;&nbsp;
                        <asp:TextBox id="DatepickerSPSEnd" runat="server" MaxLength="8" Width="65"></asp:TextBox>
                    </td>
                </tr>
                 <tr valign="middle">
                  <td> 
                       <asp:Label ID="lblServiceFamPartNum" runat="server"  CssClass="HeaderLabel" Text="Service Family Part Number(s)" />
                       <br />
                       <asp:Label ID="lblSRServiceFamPartNum" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                  </td> 
                  <td colspan="3">
                    <asp:TextBox ID="txtServiceFamPartNum" runat="server" Width="450px"></asp:TextBox><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font>                       
                  </td>
                </tr>
                <tr>
                    <td>
                        <asp:Label ID="lblSpsNumbers" runat="server"  CssClass="HeaderLabel" Text="Spare Kit Part Number(s)" />                      
                        <br />
                        <asp:Label ID="lblSpsNumbersSumReport" runat="server" ForeColor="Green" Font-Names="verdana" Font-Size="X-Small" Text="(applies to Summary Report)"></asp:Label>
                    </td>
                    <td colspan="3">
                      <asp:TextBox ID="txtSpsNumbers" runat="server" Width="450px"></asp:TextBox><font size="1" face="verdana" color="green">&nbsp;(comma&nbsp;seperated)</font>                        
                    </td>
                </tr>
                
                <!-- END FILTERS -->          
           </table>
           </ContentTemplate>
           <Triggers>
            <asp:AsyncPostBackTrigger ControlID="rdProductType" />
            <asp:AsyncPostBackTrigger ControlID="chkSKUGeo" />            
            <asp:AsyncPostBackTrigger ControlID="chkSpsGeo" />            
            <asp:AsyncPostBackTrigger ControlID="lstProducts" />
            <asp:AsyncPostBackTrigger ControlID="ddlReportProfiles"/>
            <asp:AsyncPostBackTrigger ControlID="btnReset"/>
            
           </Triggers>
       </asp:UpdatePanel>
     
      
    </div>    
    </form>
</body>
</html>
