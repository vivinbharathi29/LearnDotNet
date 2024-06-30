<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Search_SpareKitDetails" Codebehind="SpareKitDetails.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript" >
        function row_onmouseover() {
	    window.event.srcElement.parentElement.style.cursor = "hand"
	    window.event.srcElement.parentElement.style.color = "red"

    }
    function row_onmouseout() {
	    window.event.srcElement.parentElement.style.color = "black"

}
    </script>
</head>
<body>
    <form id="frmAdvSearchDetail" runat="server">
    <asp:Label ID="lblTitle" runat="server" Text="Service Advanced Search - Detailed Report" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
   
    <div style=" width:95%;" >
    <br /> <br />
        <table cellspacing="0"  style=" width:100%;" cellpadding="2" bordercolor="tan" bgcolor="ivory"  border="1" >
            <tr bgcolor="cornsilk" >
                <td align="center"><asp:Label runat="server" ID="lblSpareKitNumber" Text="Spare Kit Part Num" CssClass="LabelHeader" ></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblSKCategory" Text="SpareKitCategory" CssClass="LabelHeader"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblSpareKitDesc" Text="Spare Kit Desc" CssClass="LabelHeader"></asp:Label></td>
            </tr>
             <tr>
                  <td align="center"><asp:Label runat="server" ID="lblSpareKitNumberValue" ></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblSKCategoryValue"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblSpareKitDescValue" ></asp:Label></td>
            </tr>
        </table>
        <table cellspacing="0"  style=" width:100%;" cellpadding="2" bordercolor="tan" bgcolor="ivory" border="1">
           <tr bgcolor="cornsilk">
                <td align="center"><asp:Label runat="server" ID="lblMaterialType" Text="Material Type" CssClass="LabelHeader"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblDivision" Text="Division" CssClass="LabelHeader"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblRevisionLevel" Text="Revision Level" CssClass="LabelHeader"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblCrossPlantStatus" Text="Cross Plant Status" CssClass="LabelHeader"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblOSSPOrderable" Text="OSSP Orderable" CssClass="LabelHeader"></asp:Label></td>
            </tr>
            <tr>
               <td align="center"><asp:Label runat="server" ID="lblMaterialTypeValue" ></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblDivisionValue" ></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblRevisionLevelValue"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblCrossPlantStatusValue"></asp:Label></td>
                <td align="center"><asp:Label runat="server" ID="lblOSSPOrderableValue" ></asp:Label></td>
            </tr>
       </table>
        <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
            <tr><td align="center"><br /><br /></td></tr>
             <tr><td align="center"><b>Bill of Material</b></td></tr>
            <tr style="width:100%;">
				    <td align="center" width="100%">
                        <asp:Label ID="msgSearchNoData" runat="server" Visible="false"></asp:Label>
                    </td>
			    </tr>
            <tr>
		        <td>
                    <asp:DataGrid ID="dgData" Runat="server" AutoGenerateColumns="False"
                            CellPadding="5" BorderWidth="2px" AllowPaging="false" 
                             CssClass="table" BorderColor="Tan">
                             <ItemStyle BackColor="Ivory" />
                             <HeaderStyle  BackColor="cornsilk" />
                            <Columns >
                               <asp:BoundColumn  HeaderText="Level" DataField="SaBomItemNo" ItemStyle-Width="3%" ></asp:BoundColumn>
			                   <asp:BoundColumn HeaderText="OSSP Orderable"  DataField="SpsKitOsspOrderable"  ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Line Item"  DataField="SaBomItemNo" ItemStyle-Width="4%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Spare Kit pn" DataField="spskitpn" ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Rev" DataField="SpsKitRev" ItemStyle-Width="3%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Cross Plant Status" DataField="SpsXplantStatus" ItemStyle-Width="7%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Qty" DataField="SpsQty" ItemStyle-Width="3%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Pri Alt Gen" DataField="PartPriAlt" ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="HP SA or Component" DataField="SaDescriptiON" ItemStyle-Width="8%" ItemStyle-Wrap="false"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="HP Part Desc" DataField="PartDescriptiON" ItemStyle-Width="8%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="ODM Part N." DataField="PartOdmPartNo" ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="ODM Part Desc" DataField="PartOdmPartDescriptiON" ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="ODM Bulk Part N." DataField="PartOdmBulkPartNo" ItemStyle-Width="5%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="ODM Production MOQ" DataField="PartOdmProductiONMoq" ItemStyle-Width="8%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="ODM Post-Production MOQ" DataField="PartOdmPostProductiONMoq" ItemStyle-Width="8%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Part Supplier (ODM/OEM)" DataField="PartSupplier" ItemStyle-Width="8%"></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Model / Mfg Pn" DataField="PartModel" ItemStyle-Width="6%" ></asp:BoundColumn>
                               <asp:BoundColumn HeaderText="Comments" DataField="SpsComments" ItemStyle-Width="4%"></asp:BoundColumn>
                           </Columns>
                     </asp:DataGrid>
                 </td>
            </tr>
         </table>
         <br />
         <asp:Label ID="lblConfidential" runat="server" Text="HP - Confidential" CssClass="Confidential"></asp:Label>
         <br />
         <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
         <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>
    </div>
    </form>
</body>
</html>
