<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_OARpt" Codebehind="OARpt.aspx.vb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

    <title></title>
<script type="text/javascript" language="javascript">

    function HeaderMouseOver() {
        window.event.srcElement.style.cursor = "hand";
        window.event.srcElement.style.color = "red";
    }

    function HeaderMouseOut() {
        window.event.srcElement.style.color = "black";
    }
        
    function window_onload() {
        if (window.frameElement == null) {
            window.DIV1.style.width = window.screen.availWidth - window.DIV1.style.left - 50;
            window.DIV1.style.height = window.screen.availHeight - window.DIV1.style.top - 320;
        }
        else {
            window.DIV1.style.width = window.frameElement.width - window.DIV1.style.left - 50;
            window.DIV1.style.height = window.frameElement.height - window.DIV1.style.top - 140;
        }
    }

</script>
<style>
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
	top: expression(document.getElementById('DIV1').scrollTop);
	
}

.pinnedCol
{
	position: relative;
	left: expression(document.getElementById('DIV1').scrollLeft);
}

.pinnedColRow
{
	position: relative;
	top: expression(document.getElementById('DIV1').scrollTop);
	left: expression(document.getElementById('DIV1').scrollLeft);
}

</style>
<link rel="stylesheet" type="text/css" href="../style/Excalibur.css" />
<link href="/style/wizard style.css" type="text/css" rel="stylesheet" />
</head>
<body >
<h3 style="font-family: Verdana; font-size: medium">Product OSSP Assignment Summary</h3>
<%
    If (Me.ReturnCode = 0) Then
        Response.Write(CreateDisplayBar())
    End If
    
%>
<br/>    
    <form id="form1" runat="server">
        <asp:LinkButton ID="lnkBtnExpExcel" runat="server">Export to Excel</asp:LinkButton>&nbsp;
        <asp:Label ID="lblStatus" runat="server"></asp:Label><br /><br />
        <div style="OVERFLOW-Y: scroll;OVERFLOW-X: scroll;width:800px;height:600px" id="DIV1">
            <asp:GridView ID="grdVwRpt" runat="server" AllowSorting="True" Width="100%"
                OnSorting="grdVwRpt_Sorting" cellpadding="2" cellspacing="1" 
                bgcolor="LightGrey" style="behavior:url(../includes/client/tablehl.htc);" 
                slcolor="#FFFFCC" hlcolor="#BEC5DE" BackColor="Ivory" BorderColor="LightGray" 
                BorderStyle="Solid" BorderWidth="1px" CaptionAlign="Left" >
                <HeaderStyle CssClass="pinnedRow" BackColor="Beige" BorderColor="LightGray"
                    BorderStyle="Solid" BorderWidth="1px" ForeColor="Black" />
            </asp:GridView>
        </div>
    </form>


    <script type="text/javascript" language="javascript">
        window_onload();
    </script>
</body>
</html>
