<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_processSKUBOMAS" Codebehind="processSKUBOMAS.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>SKU BOM - Advanced Search Report</title>
    <script type="text/javascript" language="javascript">
    
    </script>
<style type="text/css">

TH
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana;
}

H3
{
    FONT-SIZE: small;
    FONT-FAMILY: Verdana;
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

.pinnedRow
{
	position: relative;
	
	
}

.pinnedCol
{
	position: relative;
	
}

.pinnedColRow
{
	position: relative;
	
}
.SKURow
{
    color:Black;
    background-color:Ivory;
    border-color:LightGray;
    border-width:1px;
    border-style:Solid;
    
    
}
.HeaderRow
{
    background-color:Ivory;
    border-color:LightGrey;
    border-width:1px;
    border-style:Solid;
    width:100%;
    
    
}
</style>
<link rel="stylesheet" type="text/css" href="../style/Excalibur.css" />
<link href="/style/wizard style.css" type="text/css" rel="stylesheet" />
</head>
<body>
<form id="theForm" name="theForm" runat="server" method="post" target="_self">
            <asp:Label ID="lblStatus" runat="server" Text="Generating Report, please wait..."></asp:Label><br />
            <asp:GridView ID="grdVwRpt" runat="server"  EnableViewState="false"
                BackColor="LightGray" Width="100%" >
                <HeaderStyle ForeColor="Black" BackColor="Beige" CssClass="pinnedRow" HorizontalAlign="Left" />
                <PagerSettings Position="TopAndBottom" Mode="NumericFirstLast" 
                    FirstPageText="&amp;lt;&amp;lt;First" LastPageText="Last&amp;gt;&amp;gt;" 
                    NextPageText="Next&amp;gt;" PreviousPageText="&amp;lt;Previous" />
                <RowStyle CssClass="SKURow" />
            </asp:GridView>
            <br />
            
            <asp:Label ID="lblHPC" runat="server" Text="HP Restricted" Font-Bold="True" Font-Size="X-Small" ForeColor="Red" Font-Names="verdana"></asp:Label>
</form>

</body>
</html>