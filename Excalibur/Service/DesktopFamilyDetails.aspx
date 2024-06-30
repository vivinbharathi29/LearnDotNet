<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_DesktopFamilyDetails" Codebehind="DesktopFamilyDetails.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Desktop Family Details</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
    <base target="_self" />
    <script id="clientEventHandlersJS" type="text/javascript">
    
        function cmdDate_onclick(FieldID) {
            var oldValue;
            var strID;
            
            if (FieldID=='txtFCS')
		        oldValue = frmDesktopFamilyDetails.txtFCS.value;		                    	
	   
	        if (FieldID=='txtEndOfService')
		        oldValue = frmDesktopFamilyDetails.txtEndOfService.value;
	
	         strID = window.showModalDialog("../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		       	
	       	if (typeof(strID) == "undefined")
		        return;
         
            if (FieldID=='txtFCS')
		        frmDesktopFamilyDetails.txtFCS.value = strID;
		            
		    if (FieldID=='txtEndOfService')
		        frmDesktopFamilyDetails.txtEndOfService.value = strID;
	 
	    }    
</script>
    <style type="text/css">
        .style1
        {
            width: 112px;
        }
    </style>
</head>
<body>
    <form id="frmDesktopFamilyDetails" runat="server">
    <div>
        <p><b>Service Family Details</b></p><br /> 
        <asp:Label ID="lblServiceFamilyPn" runat="server" Text="Family SPS Pn"></asp:Label>
        <asp:Label ID="lblFamilyPn" runat="server"></asp:Label>
        <br />
        <table class="FormTable" width="60%" >
            <tr>
                <td>GPLM</td>
                <td>
                    <asp:DropDownList runat="server" id="ddlGPLM"></asp:DropDownList>
                </td>
            </tr>
             <tr>
                <td>SPB Auto Publish</td>
                <td>
                    <asp:CheckBox runat="server" ID="chkSPBAutoPub" />
                </td>
            </tr>
             <tr>
                <td>RSL Auto Publish</td>
                <td>
                    <asp:CheckBox runat="server" ID="chkRSLAutoPub"  />
                </td>
            </tr>
             <tr>
                <td>Series Name</td>
                <td>
                   <asp:Label runat="server" id="lblSeriesName"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>Business Unit</td>
                <td>
                    <asp:RadioButtonList runat="server" ID="rbBusiness"  RepeatDirection="Horizontal">
                        <asp:ListItem Value="0" Text="Commercial" Selected="True"></asp:ListItem>
                        <asp:ListItem Value="1" Text="Consumer"></asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
             <tr>
                <td>Project Code</td>
                <td>
                    <asp:Label runat="server" id="lblProjectCode"></asp:Label>
                </td>
            </tr>
             <tr>
                <td>Partner (ODM)</td>
                <td>
                    <asp:DropDownList runat="server" id="ddlPartner"></asp:DropDownList>
                </td>
            </tr>
        </table>
        
        <table border="0" >
            
            <tr valign="bottom"  >
                 <td >
                    <br /><br />
                    <asp:Label ID="lblPlatformName" runat="server" CssClass="HeaderLabel" Text="Platform Name:" /><br />
                    <asp:TextBox runat="server" ID="txtPlatformName" Width="200" MaxLength="30"></asp:TextBox>
                 </td>
                 <td style="width:5%;"></td>
                 <td style="width:20%;padding-left:10px;" align="left" class="style1" >
                    <asp:Label ID="lblFCS" runat="server"  CssClass="HeaderLabel" Text="FCS:" /><br />
                    <asp:TextBox runat="server" ID="txtFCS" MaxLength="10" Width="65" ></asp:TextBox>                                        
                </td>
                <td align="left">
                    <a href="javascript:cmdDate_onclick('txtFCS');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>
                </td>
            </tr>
            <tr valign="bottom" >
                <td></td>
                 <td style="width:5%;"></td>
                <td style="width:20%;padding-left:10px;" align="left" class="style1" >
                    <asp:Label ID="lblDateEndOfService" runat="server"  CssClass="HeaderLabel" Text="End of Service:"  /><br />
                    <asp:TextBox runat="server" ID="txtEndOfService" MaxLength="10" Width="65" ></asp:TextBox>
                </td>
                <td>
                    <a href="javascript:cmdDate_onclick('txtEndOfService');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>
                </td>
            </tr>
            <tr align="left">
                <td colspan="4" style="padding-left:10px;" >
                    <br /> <asp:Label runat="server" ID="lblErrorMessage" ForeColor="Red" ></asp:Label> <br />
                </td>
            </tr>     
        </table>
        
        
        <hr />       
     
        <div>
            <p style="text-align:left">
            <asp:Button id="btnSaveSFP" runat="server" Title="Save Changes" Text="Submit"   />
            &nbsp;&nbsp;<asp:Button id="bntClear" runat="server" ToolTip="Clear" Text="Clear Dates" />
            &nbsp;&nbsp;<asp:Button id="btnCancelSFP" runat="server" Title="Cancel Changes" Text="Cancel" OnClientClick="JavaScript:window.opener.location.reload(true);self.close();" />           
            </p>
        </div>
        

           
    </div>
    </form>
</body>
</html>
