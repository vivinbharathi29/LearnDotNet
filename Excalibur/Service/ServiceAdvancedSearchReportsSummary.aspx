<%@ Page Language="VB" AutoEventWireup="false" EnableEventValidation="false" Inherits="DummyVBApp.Service_ServiceAdvancedSearchReportsSummary" Codebehind="ServiceAdvancedSearchReportsSummary.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
   <title>Service Advanced Search Page</title>    
   <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Label ID="lblTitle" runat="server" Text="Service Advanced Search - " Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
        <br />
        <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 80px" Runat="server" Width="98%">
       
        <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
            <tr>
                <td>
                     <asp:GridView runat="server" ID="gvSpareKitsByCategory" Width="100%" AllowPaging="true" PageSize="30"
                       AllowSorting="true" AutoGenerateColumns="false"
                       GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>                                
                            <asp:boundfield datafield="SpareKitNumber" headertext="SpareKit Number" SortExpression="SpareKitNumber" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKit Description" SortExpression="SpareKitDescription" /> 	
                            <asp:boundfield datafield="CategoryName" headertext="Category Name" SortExpression="CategoryName" /> 
                        </Columns>
                    </asp:GridView>
                           
                     <asp:GridView runat="server" ID="gvSpareKitsBom" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>
                            <asp:boundfield datafield="SpareKitNumber" headertext="SpareKit Number" SortExpression="SpareKitNumber" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKit Description" SortExpression="SpareKitDescription" />                                 
                            <asp:boundfield datafield="Sa" headertext="SubAssembly" SortExpression="Sa" /> 
                            <asp:boundfield datafield="SaDescription" headertext="SubAssembly Description" SortExpression="SaDescription" />                                 
                            <asp:boundfield datafield="PartNumber" headertext="Component" SortExpression="PartNumber" /> 
                            <asp:boundfield datafield="ComponentDescription" headertext="Component Description" SortExpression="ComponentDescription" />                                 
                            <asp:boundfield datafield="SpareKitCategory" headertext="SpareKit Category" SortExpression="SpareKitCategory" />                   
                            <asp:boundfield datafield="OdmPartNo" headertext="ODM Part Number" SortExpression="OdmPartNo" />                   
                            <asp:boundfield datafield="OdmPartDesc" headertext="ODM Part Description" SortExpression="OdmPartDesc" />                   
                            <asp:boundfield datafield="OdmBulkPartNo" headertext="ODM Bulk Part Number" SortExpression="OdmBulkPartNo" />                   
                            <asp:boundfield datafield="OdmProdMoq" headertext="ODM Production MOQ" SortExpression="OdmProdMoq" />                   
                            <asp:boundfield datafield="OdmPostProdMoq" headertext="ODM Post-Production MOQ" SortExpression="OdmPostProdMoq" />                              
                        </Columns>
                    </asp:GridView>    
                     <asp:GridView runat="server" ID="gvReportAvNumberToSps" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>
                            <asp:boundfield datafield="AvNo" headertext="Av Number" SortExpression="AvNo" /> 
                            <asp:boundfield datafield="AvCategoryName" headertext="Av Category" SortExpression="AvCategoryName" /> 
                            <asp:boundfield datafield="SpareKitNumber" headertext="SpareKit Number" SortExpression="SpareKitNumber" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKitDescription" SortExpression="SpareKitDescription" />                                 
                            <asp:boundfield datafield="SkCategoryName" headertext="SpareKit Category" SortExpression="SkCategoryName" /> 
                        </Columns>
                    </asp:GridView> 
                    <asp:GridView runat="server" ID="gvReportSkuToSpareKits" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>
                            <asp:boundfield datafield="SKU" headertext="SKU" SortExpression="SKU" /> 
                            <asp:boundfield datafield="SparekitNo" headertext="Sparekit Number" SortExpression="SparekitNo" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKit Description" SortExpression="SpareKitDescription" />                                   
                            <asp:boundfield datafield="SpareKitCategory" headertext="SpareKit Category" SortExpression="SpareKitCategory" />                              
                        </Columns>
                    </asp:GridView>    
                    <asp:GridView runat="server" ID="gvReportProductToSkuToSparekits" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>                            
                            <asp:boundfield datafield="ServiceFamilyPn" headertext="Family Spares Part Number" SortExpression="ServiceFamilyPn" /> 
                            <asp:boundfield datafield="ProductName" headertext="Product Name" SortExpression="ProductName" /> 
                            <asp:boundfield datafield="SKU" headertext="SKU" SortExpression="SKU" /> 
                            <asp:boundfield datafield="skuDescription" headertext="SKU Description" SortExpression="skuDescription" />                             
                            <asp:boundfield datafield="AvNo" headertext="Av Number" SortExpression="AvNo" /> 
                            <asp:boundfield datafield="SparekitNo" headertext="Sparekit Number" SortExpression="SparekitNo" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKit Description" SortExpression="SpareKitDescription" />                                   
                            <asp:boundfield datafield="SparekitCategory" headertext="Sparekit Category" SortExpression="SparekitCategory" /> 
                        </Columns>
                    </asp:GridView>
                    <asp:GridView runat="server" ID="gvUsedBy" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>                            
                            <asp:boundfield datafield="SKU" headertext="SKU" SortExpression="SKU" /> 
                            <asp:boundfield datafield="skuDescription" headertext="SKU Desc." SortExpression="skuDescription" />                             
                            <asp:boundfield datafield="ProductName" headertext="Product Name" SortExpression="ProductName" /> 
                            <asp:boundfield datafield="ServiceFamilyPn" headertext="ServiceFamilyPn" SortExpression="ServiceFamilyPn" /> 
                            <asp:boundfield datafield="SparekitNo" headertext="Sparekit Number" SortExpression="SparekitNo" /> 
                            <asp:boundfield datafield="SpareKitDescription" headertext="SpareKit Desc." SortExpression="SpareKitDescription" />                                   
                            <asp:boundfield datafield="SubAsembly" headertext="SubAsembly" SortExpression="SubAsembly" />
                            <asp:boundfield datafield="SubAsemblyDescription" headertext="SubAsembly Desc." SortExpression="SubAsemblyDescription" />
                            <asp:boundfield datafield="Component" headertext="Component" SortExpression="Component" />
                            <asp:boundfield datafield="ComponentDescription" headertext="Component Desc." SortExpression="ComponentDescription" />
                            <asp:boundfield datafield="SparekitCategory" headertext="Sparekit Category" SortExpression="SparekitCategory" /> 
                        </Columns>
                    </asp:GridView>
                     <asp:GridView runat="server" ID="gvRSLChangeLog" Width="100%" AllowPaging="true" PageSize="30" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                        <Columns>                            
                            <asp:boundfield datafield="ServiceFamilyPn" headertext="Family Spares Part Number" SortExpression="ServiceFamilyPn" /> 
                            <asp:boundfield datafield="ChangeDt" headertext="Date" SortExpression="ChangeDt" /> 
                            <asp:boundfield datafield="SpareKitNo" headertext="Part Number" SortExpression="SpareKitNo" /> 
                            <asp:boundfield datafield="Description" headertext="Description" SortExpression="Description" /> 
                            <asp:boundfield datafield="ColumnChanged" headertext="Column" SortExpression="ColumnChanged" /> 
                            <asp:boundfield datafield="ChangeTypeDesc" headertext="Details" SortExpression="ChangeTypeDesc" /> 
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
