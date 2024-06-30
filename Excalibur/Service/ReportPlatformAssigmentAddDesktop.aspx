<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_ReportPlatformAssigmentAddDesktop" Codebehind="ReportPlatformAssigmentAddDesktop.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>ReportPlatformAssigmentAddDesktop</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
    <base target="_self" />
    <script id="clientEventHandlersJS" type="text/javascript">
        function cmdDate_onclick(FieldID) {
            var oldValue;
            var strID;
            
            if (FieldID=='txtFCS')
		        oldValue = frmReportPlatformAssigmentAddDesktop.txtFCS.value;		                    	
	   
	        if (FieldID=='txtEndOfService')
		        oldValue = frmReportPlatformAssigmentAddDesktop.txtEndOfService.value;
	
	         strID = window.showModalDialog("../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		       	
	       	if (typeof(strID) == "undefined")
		        return;
         
            if (FieldID=='txtFCS')
		        frmReportPlatformAssigmentAddDesktop.txtFCS.value = strID;
		            
		    if (FieldID=='txtEndOfService')
		        frmReportPlatformAssigmentAddDesktop.txtEndOfService.value = strID;
	 
	    }
	    
	    
	  
          	
	</script>
</head>
<body>
    <form id="frmReportPlatformAssigmentAddDesktop" runat="server" name="frmReportPlatformAssigmentAddDesktop">
      <div>
        <p style="padding-left:10px;"><b>Add Desktop Platform</b></p>
        <table style="width: 100%;"  border="0">
            <tr>
                <td style=" width:25%; padding-left:10px;" ><asp:Label ID="lblProducFamily" runat="server" CssClass="HeaderLabel" Text="Product Family:" />
                    <br />
                    <asp:ListBox ID="lstProductFamily" runat="server" Height="200" Width="140" ></asp:ListBox>
                </td>
                <td style=" width:25%;" ><asp:Label ID="lblProductLine" runat="server" CssClass="HeaderLabel" Text="Product Line:" /><br />
                    <asp:ListBox ID="lstProductLine" runat="server" Height="200" Width="200" ></asp:ListBox>
                </td>
                <td style=" width:25%;" ><asp:Label ID="lblODM" runat="server" CssClass="HeaderLabel" Text="ODM:" /><br />
                    <asp:ListBox ID="lstODM" runat="server" Height="200" Width="140" ></asp:ListBox>
                </td>
                <td style=" width:25%;" ><asp:Label ID="lblGPLM" runat="server" CssClass="HeaderLabel" Text="GPLM:" /><br />
                    <asp:ListBox ID="lstGPLM" runat="server" Height="200" Width="140" ></asp:ListBox>
                </td>
            </tr>
            <tr>
                <td colspan="4" style="padding-left:10px;" >
                    <asp:LinkButton runat="server" ID="lnkBtAddProductFamily" Text="Add Family"></asp:LinkButton>
                </td>
            </tr>
            <tr runat="server" id="trNewProductFamily" visible="false">
                <td colspan="4" style="padding-left:10px;" >
                    <asp:Label runat="server" ID="lblProductFamilyName" Text="Product Family Name"></asp:Label>
                    <asp:TextBox runat="server" ID="txtProductFamilyName" ></asp:TextBox>
                    <asp:Button runat="server" ID="btnAddFamily" Text="OK"  />
                    <asp:Button runat="server" ID="btnCancelAddFamily" Text="Cancel" />
                </td>
            </tr>
            <tr>
                <td colspan="4" style="padding-left:10px;" >
                    <br /> <asp:Label runat="server" ID="lblErrorMessage" ForeColor="Red" ></asp:Label> <br />
                </td>
            </tr>       
            <tr>
                <td colspan="2" style="padding-left:10px;" >
                    <table>
                        <tr>
                            <td style=" width:25%;padding-left:10px;" >
                                <asp:Label ID="lblFamilyPn" runat="server" CssClass="HeaderLabel" Text="FamilyPn:" /><br />
                                <asp:TextBox runat="server" ID="txtServiceFamilyPn" Width="150" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td  style="padding-left:10px;">
                               <asp:Label ID="lblPlatformName" runat="server" CssClass="HeaderLabel" Text="Platform Name:" /><br />
                               <asp:TextBox runat="server" ID="txtPlatformName" Width="150" MaxLength="30"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="padding-left:10px;" colspan="2">
                                <asp:Label ID="lblPlatformDescription" runat="server" CssClass="HeaderLabel" Text="Platform Description:" /><br />
                                <asp:TextBox runat="server" ID="txtPlatformDescription" Width="300"></asp:TextBox>
                            </td>
                        </tr>
                                    
                        <tr>
                            <td style="padding-left:10px;" align="left">
                             <br /><br /> <br /><br />
                             <asp:Button runat="server" ID="bntAdd" Text="Submit" ToolTip="Click to Add a Desktop Platform" />
                             &nbsp;&nbsp;<asp:Button id="btnAddAnother" runat="server" ToolTip="Add Another" Text="Add Another" />
                             &nbsp;&nbsp;<asp:Button id="bntClear" runat="server" ToolTip="Clear" Text="Clear" />
                             &nbsp;&nbsp;<asp:Button id="btnCancel" runat="server" ToolTip="Cancel Changes" Text="Cancel" OnClientClick="JavaScript:window.close();" /> 
                            </td>
                        </tr>              
                    </table>
                </td>
                <td colspan="2" valign="top">
                    <table>
                        <tr >
                            <td style="padding-left:10px;" align="left" valign="middle">
                                <asp:Label ID="lblFCS" runat="server" CssClass="HeaderLabel" Text="FCS:" /><br />
                                <asp:TextBox runat="server" ID="txtFCS" Width="100" ></asp:TextBox>
                                <a href="javascript:cmdDate_onclick('txtFCS');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>
                            </td>
                        </tr>
                        <tr >
                            <td style="padding-left:10px;" align="left" valign="middle">
                                <asp:Label ID="lblDateEndOfService" runat="server" CssClass="HeaderLabel" Text="End of Service:" /><br />
                                <asp:TextBox runat="server" ID="txtEndOfService" Width="100" ></asp:TextBox>
                                <a href="javascript:cmdDate_onclick('txtEndOfService');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr> 
          
        </table>
    </div>
    </form>
</body>
</html>
