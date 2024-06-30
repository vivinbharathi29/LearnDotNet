<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_ServiceBomReport" Codebehind="ServiceBomReport.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
     <title>Service Advanced Search - Bom Report</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmServiceBomReport" runat="server">
    <div>
        <asp:Label ID="lblTitle" runat="server" Text="Service Bom Report" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
         <br />
         <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 80px" Runat="server"
		    Width="98%">
	            <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
                    <tr>
		                <td colspan="2">
       	                     <asp:GridView runat="server" ID="gvData" Width="100%" AllowPaging="false" AllowSorting="true" AutoGenerateColumns="false"
                               GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                               AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" 
                              >
                               <Columns>
                                  <asp:boundfield datafield="Business" headertext="Business" SortExpression="Business" ItemStyle-Width="20%"/> 
                                   <asp:boundfield datafield="KMAT" headertext="KMAT" SortExpression="KMAT" ItemStyle-Width="20%"/> 
                                   <asp:boundfield datafield="SKU" headertext="SKU" SortExpression="SKU" ItemStyle-Width="20%"/> 
                                   <asp:boundfield datafield="SpareKitNo" headertext="SpareKitNo" SortExpression="SpareKitNo" ItemStyle-Width="20%"/> 
                                   <asp:boundfield datafield="Description" headertext="Description" SortExpression="Description" ItemStyle-Width="20%"/> 
                                </Columns>
                          </asp:GridView>
		                </td>
	                </tr>
	                </table>
                    <br />
                    <asp:Label ID="lblConfidential" runat="server" Text="HP - Confidential" CssClass="Confidential"></asp:Label>
                    <br />
                    <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
                    <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>

            </asp:panel>

            <!-- NO DATA PANEL -->
	        <asp:panel id="pnlNoData" Runat="server">
	             <table style="LEFT: 10px; POSITION: absolute; TOP: 80px; HEIGHT: 30px" cellspacing="0"
			                            cellpadding="0" width="98%" border="0">
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
