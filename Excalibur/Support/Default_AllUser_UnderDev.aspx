<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Support_Default" Codebehind="Default_AllUser_UnderDev.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excalibur Support</title>
    <link href="Style/Support.css" rel="stylesheet" type="text/css" />
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function window_onload(){
	var strPath;
	
	//strPath = window.showModalDialog("support.asp","","dialogWidth:600px;dialogHeight:420px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No") 
    }
	

//-->
</SCRIPT>
</head>
<body onload="window_onload();" style="margin-left:8px;">
    <form id="frmMain" runat="server">
    <h2>Excalibur Support</h2>

    <div>
        <br>
        <asp:Menu ID="mnuMain" runat="server" StaticSubMenuIndent="" Orientation="Horizontal">
            <Items>
                <asp:MenuItem Text="Requests" value="1" ></asp:MenuItem>
                <asp:MenuItem Text="Articles" value="2" ></asp:MenuItem>
            </Items>
        </asp:Menu>
    
    </div>
    </form>
</body>
</html>
