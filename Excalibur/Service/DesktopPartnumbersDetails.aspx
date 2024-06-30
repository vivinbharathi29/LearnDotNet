<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_DesktopPartnumbersDetails" Codebehind="DesktopPartnumbersDetails.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Desktop PartNumbers Details</title>
    <style type="text/css">
        body
        {
            font: xx-small verdana;
        }
        legend
        {
            font: bold x-small verdana;
            color: #004874;
        }
        .floatRight
        {
        	position: absolute;
            left: 250px;
            float: left;
            text-align: left;
        }
        .floatLeft
        {
            float: left;
            font: bold x-small verdana;
        }
        .floatText
        {
            position: absolute;
            left: 250px;
        }
        .display
        {
            text-align: left;
        }
        .link
        {
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .link:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .linkEditFloatRight
        {
            float: right;
            text-align: right;
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .linkEditFloatRight:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .inputBox
        {
            width: 350px;
        }
        .BomTable
        {
            width: 100%;
            border-bottom: solid 1px black;
            border-collapse: collapse;
        }
        .BomTable th
        {
            background: #004874;
            color: #ffffff;
        }
        .BomTable td
        {
            border-bottom: solid 1px black;
        }
    </style>
    
    <script id="clientEventHandlersJS" type="text/javascript">
        function cmdDate_onclick(FieldID) {
            var oldValue;
            var strID;
            
            if (FieldID=='txtFirstServiceDt')
		        oldValue = frmDesktopPartNumbersDetails.txtFirstServiceDt.value;		                    	
	   
	         strID = window.showModalDialog("../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		       	
	       	if (typeof(strID) == "undefined")
		        return;
         
            if (FieldID=='txtFirstServiceDt')
		        frmDesktopPartNumbersDetails.txtFirstServiceDt.value = strID;
		 
	    } 	
	</script>
</head>
<body>
    <form id="frmDesktopPartNumbersDetails" runat="server">
     <div style="width: 800px">
      <div id="KitDetails">
            <fieldset>
                  <div style="height: 25px">
                    <span id="Error" class="floatRight">
                        <asp:Label runat="server" ID="lblError" ForeColor="Red"></asp:Label>
                    </span>
                </div>
                <legend>Spare Kit -<asp:Label runat="server" id="lblUserName"></asp:Label></legend>
                <div style="height: 25px">
                    <span class="floatLeft">Part No:</span>
                    <span id="spsPartNoEdit" class="floatRight">
                        <asp:Label runat="server" ID="lblSpsPartNumber"></asp:Label>
                    </span>
                </div>
                <div style="height: 25px">
                    <span id="spsDescriptionLabel" class="floatLeft">Description:</span>
                    <span id="spsDescriptionDisplay" class="floatText"> <asp:Label runat="server" ID="lblSparekitDescription"></asp:Label></span>                    
                </div>
                <div style="height: 25px">
                    <span class="floatLeft">Part Type:</span>
                    <span id="spsCategory" class="floatRight">
                        <asp:DropDownList runat="server" ID="ddlCategories"></asp:DropDownList>                        
                    </span> 
                </div>
                <div style="height: 25px">
                    <span class="floatLeft">CSR Level:</span>
                    <span id="spsCsrLevelEdit" class="floatRight">
                         <asp:DropDownList runat="server" ID="ddlCustomerLevel"></asp:DropDownList>                                     
                    </span> 
                </div>
                <div style="height:25px">
                    <span class="floatLeft">Disposition:</span>
                    <span id="spsDispositionEdit" class="floatRight">
                       <asp:DropDownList runat="server" ID="ddlDisposition">
                        <asp:ListItem Text="-- Select Disposition --" Value="0" Selected="True"></asp:ListItem>
                        <asp:ListItem Text="1 - Disposable" Value="1"></asp:ListItem>
                        <asp:ListItem Text="2 - Repairable" Value="2"></asp:ListItem>
                        <asp:ListItem Text="3 - Return to Vendor" Value="3"></asp:ListItem>
                        <asp:ListItem Text="4 - Repair/Exhcange Only" Value="4"></asp:ListItem>
                       </asp:DropDownList>                                         
                    </span>                    
                </div>
                <div style="height:25px">
                    <span class="floatLeft">Warranty Labour Tier:</span>
                    <span id="spsWarrantyEdit" class="floatRight">
                         <asp:DropDownList runat="server" ID="ddlWarranty">
                             <asp:ListItem Text="-- Select Warraty Labour Tier --" Value="0" Selected="True"></asp:ListItem>
                             <asp:ListItem Text="A - 0 - 5 Minutes" Value="A"></asp:ListItem>
                             <asp:ListItem Text="B - 5 - 10 Minutes" Value="B"></asp:ListItem>
                             <asp:ListItem Text="C - 15 - 30 Minutes" Value="C"></asp:ListItem>
                             <asp:ListItem Text="D - No Reimbursment" Value="D"></asp:ListItem>
                         </asp:DropDownList>                                         
                    </span>                        
                </div>                
                <div style="height:25px">
                    <span class="floatLeft">Local Stock Advice:</span>
                    <span id="spsLocalStockAdviceEdit" class="floatRight">
                        <asp:DropDownList runat="server" ID="ddlLocalStockAdvice">
                              <asp:ListItem Text="-- Select Stock Advice --" Value="0" Selected="True"></asp:ListItem>
                              <asp:ListItem Text="1 - Don't stock local (non SPOF and not likely to fail)" Value="1"></asp:ListItem>
                              <asp:ListItem Text="2 - Stock Strategically (non SPOF and likely to fail)" Value="2"></asp:ListItem>
                              <asp:ListItem Text="3 - Stock Local (SPOF and not likely to fail)" Value="3"></asp:ListItem>
                              <asp:ListItem Text="4 - Stock Local Critical (SPOF and likely to fail)" Value="4"></asp:ListItem>
                        </asp:DropDownList>  
                    </span>
                </div>
                <div style="height:25px">
                    <span class="floatLeft">GEOS:</span>
                    <span id="spsGeosEdit" class="floatText">
                        <asp:CheckBox runat="server" ID="chkGeosNa" Text="NA" />
                        <asp:CheckBox runat="server" ID="chkGeosLa" Text="LA" />
                        <asp:CheckBox runat="server" ID="chkGeosApj" Text="APJ" />
                        <asp:CheckBox runat="server" ID="chkGeosEmea" Text="EMEA" />
                    </span>
		        </div>
		        <div style="height: 25px">
                    <span id="Span1" class="floatLeft">Supplier:</span>
                    <span id="Span2" class="floatText"><asp:TextBox runat="server" id="txtSupplier"></asp:TextBox></span>                    
                </div>
                <div style="height: 25px">
                    <span class="floatLeft">First Service Dt.:</span>
                    <span id="spsFirstServiceDtEdit" class="floatRight">
                        <asp:TextBox runat="server" ID="txtFirstServiceDt" CssClass="inputBox" Width="60px" ></asp:TextBox>        
                         <a href="javascript:cmdDate_onclick('txtFirstServiceDt');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>                        
			        </span>
			    </div>
			    <div style="height: 75px">
                    <span class="floatLeft">RSL Comments:</span>
                    <span class="floatRight">
                        <textarea runat="server" id="txtRslComment" rows="4" class="inputBox" cols="150"></textarea>
                    </span>
                </div>
            </fieldset>
      </div>
      <div id="BomInfo">
            <fieldset>
                <legend>BOM Info</legend><span id="BomDetails">&nbsp;</span>
                <asp:panel id="pnlBomData" Runat="server" >
                    <table  cellspacing="0" cellpadding="0" width="100%" border="0" >
                        <tr>
                            <td>
                             <asp:GridView runat="server" ID="gvBomData" Width="100%" AutoGenerateColumns="false"
                               GridLines="Both"  ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                               HeaderStyle-CssClass="BomTable"  >
                               <Columns>
                                    <asp:boundfield datafield="Level1" headertext="Kit" ItemStyle-Wrap="false"  /> 
                                    <asp:boundfield datafield="L1Description" headertext="Kit Desc." ItemStyle-Wrap="true"  /> 
                                    <asp:boundfield datafield="Level2" headertext="SA" ItemStyle-Wrap="false"  /> 
                                    <asp:boundfield datafield="L2Description" headertext="SA. Desc." ItemStyle-Wrap="true"  /> 
                                    <asp:boundfield datafield="Level3" headertext="Component" ItemStyle-Wrap="false"  /> 
                                    <asp:boundfield datafield="L3Description" headertext="Component Desc." ItemStyle-Wrap="true"  /> 
	                           </Columns>
                             </asp:GridView>
                           </td>
                      </tr>
                  </table>
                </asp:Panel>                
                <asp:panel id="pnlNoBomData" Runat="server">
                     <table style="LEFT: 10px; POSITION: absolute; TOP: 300px; HEIGHT: 30px" cellspacing="0" cellpadding="0" width="98%" border="0">
	                    <tr style="width:100%;">
		                    <td align="center" >
                                <asp:Label ID="msgSearchNoData" runat="server">Bom Not Found</asp:Label>
                            </td>
	                    </tr>
                    </table>
                </asp:panel>
            </fieldset>
        </div>
    </div>
    <div id="Buttons" >
        <table>
            <tr >
                <td style="padding-left:10px;" align="left" valign="middle">
                    <asp:Button runat="server" ID="bntAdd" Text="Submit" ToolTip="Submit" 
                        style="height: 26px" />
                    &nbsp;&nbsp;<asp:Button id="btnCancel" runat="server" ToolTip="Cancel" Text="Cancel" OnClientClick="JavaScript:window.close();" />   
                </td>
            </tr>
        </table>
    </div>
  </form>
</body>
</html>
