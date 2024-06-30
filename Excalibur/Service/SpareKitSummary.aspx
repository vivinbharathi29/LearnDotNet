<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Search_SpareKitSummary" Codebehind="SpareKitSummary.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Service Advanced Search - Summary</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmSpareKitDeatils" runat="server">
    <div>
     <asp:Label ID="lblTitle" runat="server" Text="Service Advanced Search" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
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
                               <asp:HyperLinkField DataNavigateUrlFields="SpareKitID,SpareKitNumber,ServiceFamilyPn" 
                                DataNavigateUrlFormatString="~/Service/SpareKitDetails.aspx?SpareKitID={0}&SpareKitNumber={1}&ServiceFamilyPn={2}"
                                 DataTextField="SpareKitNumber" SortExpression="SpareKitNumber" HeaderText="Spare Kit Part Number"  target = "_blank" />
                               <asp:boundfield datafield="CategoryName" headertext="Spare Kit Category" SortExpression="CategoryName" ItemStyle-Width="20%"/> 
                               <asp:boundfield datafield="SpareKitDescription" headertext="Spare Kit Desccription"  SortExpression="SpareKitDescription" ItemStyle-Width="20%"/> 
                               <asp:boundfield datafield="RevisionLevel" headertext="Revision Level" SortExpression="RevisionLevel" ItemStyle-Width="20%"/> 
                               <asp:boundfield datafield="CrossPlantStatus" headertext="Cross Plant Status" SortExpression="CrossPlantStatus" ItemStyle-Width="10%"  /> 
                               <asp:boundfield datafield="ServiceFamilyPn"  headertext="Service Family PatNumber" SortExpression="ServiceFamilyPn"  />                                
                               <asp:boundfield datafield="ProductName"  headertext="Product Name" SortExpression="ProductName"  />                                                               
                               <asp:boundfield datafield="LastUpdDate"  headertext="Last UpdDate" SortExpression="LastUpdDate"  />                                                               
                               
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
