<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.MobileSE_Today_TodayExtend" Codebehind="TodayExtend.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Extend Table Rows</title>
    <link href="../../Style/Excalibur.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="frmExtendTable" runat="server">        
    <div>
     <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
     <asp:TextBox runat="server" ID="txtHidUpdateSelectedAVs" Visible="false"></asp:TextBox>
     <asp:Label ID="lblTitle" runat="server" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
     <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 80px" Runat="server"
		Width="98%">
	        <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
                <tr style="display:">
                    <td><font face="verdana" size="1">
                        <a runat="server" id="lnkUpdateSelectedAVs" >Update Selected Avs</a><br/></font>                       
                       <br /> <br />
                    </td> 
                </tr>
	            <tr>
		            <td>
		           <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                   <ContentTemplate>
		              <asp:GridView runat="server" ID="gvAVsDeletedMappedSPS" Width="100%" AllowPaging="false" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" 
                          >
                           <Columns>
                               <asp:TemplateField>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkUpdateSelectedAVs" name="chkUpdateSelectedAVs" runat="server"/>             
                                    </ItemTemplate>
                               </asp:TemplateField>                             
                               <asp:boundfield datafield="AvDetailID" headertext="AV Detail" SortExpression="AvDetailID" ItemStyle-Width="10%"/> 
                               <asp:boundfield datafield="AvNo" headertext="AV Number"  SortExpression="AvNo" ItemStyle-Width="10%"/> 
                               <asp:boundfield datafield="category" headertext="Feature Category" SortExpression="category" ItemStyle-Width="20%"/> 
                               <asp:boundfield datafield="brand" headertext="Brand" SortExpression="brand" ItemStyle-Width="17%"  /> 
                               <asp:boundfield datafield="GPGDescription"  headertext="GPG Description" SortExpression="GPGDescription"   ItemStyle-Width="25%"  />                                
                               <asp:boundfield datafield="Deleted"  headertext="Deleted" SortExpression="Deleted"  ItemStyle-Width="18%"   />                                
                            </Columns>
                        </asp:GridView>
                        <asp:GridView runat="server" ID="gvAVsNotMappedToSPS" Width="100%" AllowPaging="false" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" 
                          >
                           <Columns>
                               <asp:boundfield datafield="ProductName" headertext="Program" SortExpression="ProductName"/>                         
                               <asp:boundfield datafield="BrandName" headertext="Brand" SortExpression="BrandName"/>                         
                               <asp:boundfield datafield="AvFeatureCategory" headertext="Feature Category" SortExpression="AvFeatureCategory"/>                         
                               <asp:boundfield datafield="GPGDescription" headertext="GPG Description" SortExpression="GPGDescription"/>                         
                               <asp:boundfield datafield="AvNo" headertext="AV" SortExpression="AvNo"/>                         
                            </Columns>
                        </asp:GridView>
                        <asp:GridView runat="server" ID="gvSPSNotMappedToAV" Width="100%" AllowPaging="false" AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" 
                          >
                           <Columns>
                               <asp:boundfield datafield="SpareKitNo" headertext="Spare Kit" SortExpression="SpareKitNo" />                
                               <asp:boundfield datafield="Description" headertext="Description" SortExpression="Description" />                
                               <asp:boundfield datafield="CategoryName" headertext="Feature Category" SortExpression="CategoryName" />                
                               <asp:boundfield datafield="DotsName" headertext="Products" SortExpression="DotsName" />                
                            </Columns>
                        </asp:GridView>
                        <!-- DotsName
                        LEFT(rs("DotsName"), LEN(rs("DotsName"))-2) & "&nbsp;&nbsp;-->
                    </ContentTemplate>
                <Triggers>
                    <asp:AsyncPostBackTrigger ControlID="gvAVsDeletedMappedSPS" />
                    <asp:AsyncPostBackTrigger ControlID="gvAVsNotMappedToSPS"/>
                    <asp:AsyncPostBackTrigger ControlID="gvSPSNotMappedToAV"/>
                </Triggers>
              </asp:UpdatePanel>
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
