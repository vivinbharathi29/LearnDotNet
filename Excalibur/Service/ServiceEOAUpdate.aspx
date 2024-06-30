<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_ServiceEOAUpdate"  Codebehind="ServiceEOAUpdate.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Service EOA Update</title>
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <link href="../style/shared.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="../Scripts/shared_functions.js" type="text/javascript"></script>

    <script type="text/javascript">

        function window_onload() {
            modalDialog.load(); 
        }

        function OpenURLDialog(intVersionId) { 
            var url =   "../Deliverable/EOLDate.asp?TypeID=2&ID=" + intVersionId; 
            modalDialog.open({ 
                dialogTitle: "Update Service EOA",   
                dialogURL: '' + url + '', 
                dialogHeight: 550, 
                dialogWidth: 440,   
                dialogResizable: true, 
                dialogDraggable: true 
            }); 
        }

        function EditSEOA(intVersionId) {
            OpenURLDialog(intVersionId);
        }

        function closeModalDialog(bReload) { 
            modalDialog.cancel(bReload); 
        }

        function ClosePopUp() {
            closeModalDialog(true);  
        }

        function CancelPopUp() {
            closeModalDialog(false);  
        }


/*
       
        function ClosePropertiesDialog(strID) {
            $("#iframeDialog").dialog("close");
            if (typeof (strID) != "undefined") window.location.reload(true);
        }

        function CloseIframeDialog() {
            $("#iframeDialog").dialog("close");
        }


        */
    </script>


</head>
<body onload="window_onload();">
    <form id="frmServiceEOAUpdate" runat="server">
      <div>
        <asp:Label ID="lblTitle" runat="server" Text="Service EOA Update" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
        <br />
        <table style="width: 100%;"  border="0">
            <tr>
                <td style="padding-top: 10px; padding-bottom: 10; width:15%;" > 
                    <asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" />
                    <br />
                    <asp:DropDownList ID="ddlReportFormat" runat="server" >
                        <asp:ListItem Value="0" Text="HTML" />
                        <asp:ListItem Value="1" Text="Excel" />                            
                    </asp:DropDownList>
                 </td>  
                 <td style="padding-top: 10px; padding-bottom: 5;" align="left" >
                    <asp:Button ID="btnReport" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Submit" 
                            
                    />
                     <asp:Button ID="btnReset" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Reset" />
                            
                    <br />
                 </td>   
                 <td ><asp:Label runat="server" ID="lblErrorMsg"  ForeColor="Red"></asp:Label></td>
            </tr>
            <tr> 
                <td  style=" width:150px;">
                    <asp:Label ID="lblCategories" runat="server" CssClass="HeaderLabel" Text="Categories:" /><br />
                    <asp:ListBox ID="lstCategories" runat="server" Height="150" Width="170" SelectionMode="Multiple" AutoPostBack="true"></asp:ListBox>
                </td>             
                <td  style=" width:150px;">
                    <asp:Label ID="lblSupplier" runat="server" CssClass="HeaderLabel" Text="Supplier:" /><br />
                    <asp:ListBox ID="lstSuppplier" runat="server" Height="150" Width="140" SelectionMode="Multiple" ></asp:ListBox><br />
                    <asp:LinkButton runat="server" ID="lnkAddSupplier" Text="Add Supplier"></asp:LinkButton>                    
                </td>
                <td   style=" width:250px;">
                      <asp:Label ID="lblHpPartNumner" runat="server" CssClass="HeaderLabel" Text="HP Part Number:" /><br />
                        <asp:TextBox runat="server" ID="txtHPPartNumber"></asp:TextBox>
                        <br /><br /><br />
                        <asp:Label ID="lblServioceEOA" runat="server" CssClass="HeaderLabel" Text="Service EOA:" /><br />
                        <asp:RadioButtonList runat="server"  ID="rdServiceEAO" RepeatDirection="Horizontal" AutoPostBack="false"  >
                            <asp:ListItem Value="0" Text="Unavailable" ></asp:ListItem>
                            <asp:ListItem Value="1" Text="Blank" ></asp:ListItem>
                            <asp:ListItem Value="2" Text="Middle" ></asp:ListItem>
                        </asp:RadioButtonList>
                </td>
                <td>
                    <asp:Panel ID ="pnlUploadEOAFile" runat="server">
                        <asp:Label ID="lblFileToUpload" runat="server" CssClass="HeaderLabel" Text="Service EOA File:" /><br />
                        <asp:FileUpload runat="server" ID="flServiceEOA" Width="200" ToolTip="Press to select a ServiceEOA file to upload" BackColor="White"  />
                        <asp:Button ID="bntUploadFile" runat="server" BackColor="#333333" BorderColor="#333333" BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  Font-Size="X-Small" ForeColor="White" Height="18px" Text="Upload File" />
                        <br />  <br /> 
                    </asp:Panel>                         
                </td>       
             </tr>
             <tr runat="server" id="trNewSupplier" visible="false">
                <td colspan="2" style="padding-left:10px;" >
                    <asp:Label runat="server" ID="lblAddSupplier" Text="Supplier Name"></asp:Label>
                    <asp:TextBox runat="server" ID="txtSupplierName" ></asp:TextBox>
                    <asp:Button runat="server" ID="btnAddSupplier" Text="OK"  />
                    <asp:Button runat="server" ID="btnCancelAddSupplier" Text="Cancel" />
                </td>
                <td colspan="2" style="padding-left:10px;" >
                    <asp:Label runat="server" ID="lblErrorMessage" ForeColor="Red" ></asp:Label>
                </td>
             </tr>       
         </table>    
         <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 280px" Runat="server" Width="99%">
                <br />    
              <asp:Label ID="lblConfidential" runat="server" Text="HP - Restricted" CssClass="Confidential"></asp:Label>
              <br />              <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
              <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>
              <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
                <tr>
                     <td>
                       <asp:GridView runat="server" ID="gvEOAUpdate" Width="90%" AllowPaging="true" PageSize="30"
                       AutoGenerateColumns="false" GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true" AllowSorting="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="ExcaliburID" headertext="ExcaliburID" SortExpression="ExcaliburId" />
                                <asp:boundfield datafield="Supplier" headertext="Supplier" SortExpression="Supplier"  />
                                <asp:boundfield datafield="Model" headertext="Model"  SortExpression="Model" />
                                <asp:boundfield datafield="HW" headertext="HW"  SortExpression="HW" />
                                <asp:boundfield datafield="FW" headertext="FW"  SortExpression="FW" />
                                <asp:boundfield datafield="HPPartNumber" headertext="HPPartNumber"  SortExpression="HPPartNumber" />
                                <asp:boundfield datafield="deliverablename" headertext="Description"  SortExpression="deliverablename" />            
                                <asp:boundfield datafield="samples" headertext="Samples"  DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center" SortExpression="samples" />            
                                <asp:boundfield datafield="ServiceEOA" headertext="Service EOA Date" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center" SortExpression="ServiceEOA" />
                                <asp:TemplateField HeaderText="Service Active Status" ItemStyle-HorizontalAlign="Center"> 
                                    <ItemTemplate>
                                        <asp:Label ID="lblActive" runat="server" Text='<%# Bind("ServiceActive") %>'></asp:Label>&nbsp;
                                        <asp:Label ID="Label1" runat="server" Text='<%# Eval("ServiceEOADate", "{0:d}") %>'></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Edit" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <input type="button" id="btnEdit" value="Edit" onclick="EditSEOA('<%#Eval("ExcaliburID")%>');">
                                    </ItemTemplate>
                                </asp:TemplateField>
                          </Columns>
                       </asp:GridView>
                    </td>  
                </tr>
              </table>
          </asp:panel>
           <!-- NO DATA PANEL -->
            <asp:panel id="pnlNoData" Runat="server">
                 <table style="LEFT: 10px; POSITION: absolute; TOP: 300px; HEIGHT: 30px" cellspacing="0" cellpadding="0" width="98%" border="0">
	                <tr style="width:100%;">
		                <td align="center" width="100%">
                            <asp:Label ID="msgSearchNoData" runat="server"></asp:Label>
                        </td>
	                </tr>
                </table>
            </asp:panel>
            <!-- END NO DATA PANEL -->

          <div style="display: none;">
              <div id="iframeDialog" title="ExtendTables">
                  <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
              </div>
          </div>
          <asp:HiddenField ID="hidEditPermission" runat="server" EnableViewState="true" Value="0" />
        </div>
    </form>
</body>
</html>
