<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_SparekitsEOSL" Codebehind="SparekitsEOSL.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Sparekits EOSL</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />       
</head>
<body>
    <form id="frmSparekitsEOSL" runat="server">
        <div>
        <asp:Label ID="lblTitle" runat="server" Text="Sparekits EOSL" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
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
                            OnClientClick="setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='SparekitsEOSL.aspx';}, 5);"
                    />
                 </td>   
            </tr>
          </table>
          <asp:panel id="pnlData" style="LEFT: 10px; POSITION: absolute; TOP: 90px" Runat="server" Width="99%">
              <asp:Label ID="lblConfidential" runat="server" Text="HP - Restricted" CssClass="Confidential"></asp:Label>
              <br />
              <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
              <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>
              <table  cellspacing="0" cellpadding="0" width="100%" align="center" border="0"  >
                <tr>
                    <td>
                       <asp:GridView runat="server" ID="gvSparekitsEOSL" Width="50%" AllowPaging="true" PageSize="30" AllowSorting="true"
                       AutoGenerateColumns="false" GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="SparekitNo" headertext="SparekitNo" ItemStyle-HorizontalAlign="Center" SortExpression="SparekitNo"  /> 
                                <asp:boundfield datafield="description" headertext="Description"  SortExpression="description"/> 
                                <asp:boundfield datafield="Business" headertext="Business"  SortExpression="Business" /> 
                                <asp:boundfield datafield="FamilyName" headertext="Family" SortExpression="FamilyName" /> 
                                 <asp:boundfield datafield="ServiceLifeDate" headertext="Max EOSL" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center" SortExpression="ServiceLifeDate"/>                                                         
                           </Columns>
                       </asp:GridView>
                    </td>           
                    
                    <td>
                       <asp:GridView runat="server" ID="gvSparekitsEOSL30" Width="20%" AllowPaging="true" PageSize="30"
                       AutoGenerateColumns="false" GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="SparekitNo" headertext="SparekitNo" ItemStyle-HorizontalAlign="Center"  /> 
                                <asp:boundfield datafield="serviceEOADate" headertext="Max EOA" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center"/> 
                           </Columns>
                       </asp:GridView>
                    </td>   
                    <td>
                       <asp:GridView runat="server" ID="gvSparekitsEOSL60" Width="20%" AllowPaging="true" PageSize="30"
                       AutoGenerateColumns="false" GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="SparekitNo" headertext="SparekitNo" ItemStyle-HorizontalAlign="Center"  /> 
                                <asp:boundfield datafield="serviceEOADate" headertext="Max EOA" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center"/> 
                           </Columns>
                       </asp:GridView>
                    </td>   
                    <td>
                       <asp:GridView runat="server" ID="gvSparekitsEOSL90" Width="20%" AllowPaging="true" PageSize="30"
                       AutoGenerateColumns="false" GridLines="Both"  ShowHeaderWhenEmpty="true" ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                       AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="SparekitNo" headertext="SparekitNo" ItemStyle-HorizontalAlign="Center"  /> 
                                <asp:boundfield datafield="serviceEOADate" headertext="Max EOA" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center"/> 
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
        </div>
    </form>
</body>
</html>
