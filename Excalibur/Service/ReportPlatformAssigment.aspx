<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Service_ReportPlatformAssigment" Codebehind="ReportPlatformAssigment.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>ReportPlatformAssigment</title>
    <link href="../Style/Excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../Style/general.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        function AddPlatformDesktop()
        {
            var strResult;
	        strResult = window.showModalDialog("../Service/ReportPlatformAssigmentAddDesktop.aspx","","dialogWidth:700px;dialogHeight:550px;edge: Raised;maximize:no;center:Yes; help: No;resizable: Yes;status: No")  
              
        }
        
        function ActionCell_onmouseover() {
            window.event.srcElement.style.background = "gainsboro";
            window.event.srcElement.style.cursor = "hand";
            window.event.srcElement.style.color = "black";
        }

        function ActionCell_onmouseout() {
            window.event.srcElement.style.color = "white";
            window.event.srcElement.style.background = "#333333";
        }
    </script>
     <script id="clientEventHandlersJS" type="text/javascript">
        function cmdDate_onclick(FieldID) {
	            var strID;
	            var oldValue;
	            
	            if (FieldID=='txtStartDate')
		            oldValue = frmReportPlatformAssigment.txtStartDate.value;
	            else if (FieldID=='txtEndDate')
		            oldValue = frmReportPlatformAssigment.txtEndDate.value;
            		
	            strID = window.showModalDialog("../mobilese/today/caldraw1.asp",oldValue,"dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	            if (typeof(strID) == "undefined")
		            return;
		        
		        if (FieldID=='txtStartDate')
		            frmReportPlatformAssigment.txtStartDate.value = strID;
	            else if (FieldID=='txtEndDate')
		            frmReportPlatformAssigment.txtEndDate.value = strID;
	    }       	
		
	        
    </script>
</head>
<body>
    <form id="frmReportPlatformAssigment" runat="server">
    <div>
        <asp:Label ID="lblTitle" runat="server" Text="Report Platform Assignment Metrics" Font-Bold="true" Font-Names="verdana" Font-Size="Large"></asp:Label >
        <br />
        <!-- FILTERS -->
        <table style="width: 100%;"  border="0">
            <tr>
                <td style=" padding-top: 10px; padding-bottom: 10; width:15%" > 
                    <asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" />
                    <br />
                    <asp:DropDownList ID="ddlReportFormat" runat="server" >
                        <asp:ListItem Value="0" Text="HTML" />
                        <asp:ListItem Value="1" Text="Excel" />                            
                    </asp:DropDownList>
                </td>
                <td style="width:15%;">
                    <asp:RadioButtonList runat="server" ID="rdNotebookDesktop" RepeatDirection="Horizontal" AutoPostBack="true" >
                        <asp:ListItem Selected="True" Value="0" Text="Notebook"></asp:ListItem>
                        <asp:ListItem Value="1" Text="Desktop"></asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td style="padding-top: 10px; padding-bottom: 5;" align="left">
                    <asp:Button ID="btnReport" runat="server" BackColor="#333333" BorderColor="#333333"
                            BorderStyle="Solid" BorderWidth="1px" Font-Bold="True" 
                            Font-Size="X-Small" ForeColor="White" Height="18px" Text="Submit" 
                            OnClientClick="setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='ReportPlatformAssignmentMetrics.aspx';}, 5);"
                    />
                    <asp:Button ID="btnReset" runat="server" BackColor="#333333" BorderColor="#333333"
                        BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  Font-Size="X-Small" ForeColor="White" Height="18px" Text="Reset" />
                    <asp:Button ID="btnAddDesktop" runat="server" BackColor="#333333" BorderColor="#333333"
                        BorderStyle="Solid" BorderWidth="1px" Font-Bold="True"  Font-Size="X-Small" ForeColor="White" Height="18px" Text="Add Desktop"  />


                </td>        
               <td colspan="4">
                    <asp:Label runat="server" Visible="false"   id="lblDesktopAddCommunication" CssClass="Confidential">2015 Desktop Products are being added by the Desktop SEPM team.  For any Product prior to 2015, please click Add Product from the left menu and make sure you select the Legacy Product selection.</asp:Label>
				</td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="lblPlatform" runat="server" CssClass="HeaderLabel" Text="Platform:" /><br />
                    <asp:ListBox ID="lstPlatform" runat="server" Height="150" Width="140" SelectionMode="Multiple"></asp:ListBox>
                </td>
                <td style=" width:150px;"><asp:Label ID="lblODM" runat="server" CssClass="HeaderLabel" Text="ODM:" /><br />
                    <asp:ListBox ID="lstODM" runat="server" Height="150" Width="140" ></asp:ListBox>
                </td>
                <td style=" width:150px;"><asp:Label ID="lblGPLM" runat="server" CssClass="HeaderLabel" Text="GPLM:" /><br />
                    <asp:ListBox ID="lstGPLM" runat="server" Height="150" Width="140" ></asp:ListBox>
                </td>
                <td runat="server" id="tdSpdm" style=" width:150px;"><asp:Label ID="lblSpdn" runat="server" CssClass="HeaderLabel" Text="BOM Analyst:" /><br />
                    <asp:ListBox ID="lstSpdm" runat="server" Height="150" Width="140" ></asp:ListBox>
                </td>
                <td runat="server" id="tdPsm" style=" width:150px;"><asp:Label ID="lblPsm" runat="server" CssClass="HeaderLabel" Text="PSM:" /><br />
                    <asp:ListBox ID="lstPsm" runat="server" Height="150" Width="140" ></asp:ListBox>
                </td>
                <td runat="server" id="tdProductLine" style=" width:150px;">
                    <asp:Label ID="lblProductLine" runat="server" CssClass="HeaderLabel" Text="Product Line:" /><br />
                    <asp:ListBox ID="lstProductLine" runat="server" Height="150" Width="220" ></asp:ListBox>
                </td>
                <td valign="top">
                    <table>
                        <tr>
                            <td>
                                <asp:Label ID="lblFamilyPn" runat="server" CssClass="HeaderLabel" Text="FamilyPn:"  /><br />
                                <asp:TextBox runat="server" ID="txtServiceFamilyPn" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                         <tr>
                            <td>
                                <asp:Label ID="lbl" runat="server" CssClass="HeaderLabel" Text="Project Number:" /><br />
                                <asp:TextBox runat="server" ID="txtProjextNumber"></asp:TextBox>
                            </td>
                         </tr>
                         <tr><td><br /></td></tr>
                        <tr runat="server" id="trBusiness">
                            <td>
                                <asp:Label ID="lblBusiness" runat="server" CssClass="HeaderLabel" Text="Business:" /><br />
                                <asp:CheckBoxList runat="server" ID="ckBusiness" RepeatDirection="Horizontal" >
                                <asp:ListItem Value="1" Text="Commercial"></asp:ListItem>
                                <asp:ListItem Value="2" Text="Consumer"></asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                         </tr>
                        
                    </table>                   
                </td>
             </tr>
             <tr runat="server" id="trDates">
                <td colspan="5">
                    <table cellpadding="0" cellspacing="0">
                        <tr >
                            <td style="padding-top: 5px; width:120px; " ><asp:Label ID="lblDataRange" runat="server" CssClass="HeaderLabel" Text="Date Range:" /></td>
                            <td style="padding-top: 5px;"><asp:Label ID="lblFrom" runat="server" CssClass="HeaderLabel" Text="From:" /></td>
                            <td style="padding-top: 5px;">
                                <input style="width:80px" type="text" id="txtStartDate" runat="server" name="txtStartDate" />
                                <a href="javascript:cmdDate_onclick('txtStartDate');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose"  width="26" height="21" /></a>&nbsp;
                            </td>
                            <td style="padding-top: 5px;"><asp:Label ID="lblTo" runat="server" CssClass="HeaderLabel" Text="To:" /></td>
                            <td style="padding-top: 5px;">
                                <input style="width:80px" type="text" id="txtEndDate" runat="server" name="txtEndDate" />
                                <a href="javascript:cmdDate_onclick('txtEndDate');"><img src="../MobileSE/Today/images/calendar.gif" alt="Choose" width="26" height="21" /></a>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            </table>
            <!-- END FILTERS -->
  
            <asp:panel id="pnlData" Runat="server" >
                <br /><br /><br /><br />
                <asp:Label ID="lblConfidential" runat="server" Text="HP - Restricted" CssClass="Confidential"></asp:Label>
                <br />
                <asp:Label ID="lblLastRun" runat="server" Text="Report Generated "></asp:Label>
                <asp:Label ID="lblLastRunDate" runat="server" Text="Label"></asp:Label>
                </asp:panel>
                 <asp:Panel runat="server" ID="pnlDesktop" Visible="false" >
                  <table  cellspacing="0" cellpadding="0" width="100%" border="0"  >
                     <tr>
                        <td>
                             <asp:GridView runat="server" ID="gvPADesktop" Width="100%" AllowPaging="true" PageSize="100"
                               AllowSorting="true" AutoGenerateColumns="false"
                               GridLines="Both"  ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                               AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="" BorderColor="tan" >
                               <Columns>
                                    <asp:boundfield datafield="FamilyName" headertext="FamilyName" SortExpression="FamilyName" ItemStyle-Wrap="true"  /> 
                                    <asp:boundfield datafield="ServiceFamilyPn" headertext="ServiceFamilyPn" SortExpression="ServiceFamilyPn" ItemStyle-Wrap="true"  /> 
                                    <asp:boundfield datafield="Platform" headertext="Platform" SortExpression="Platform" ItemStyle-Wrap="true"  /> 
                                    <asp:boundfield datafield="ProjectNumber" headertext="ProjectNumber" SortExpression="ProjectNumber" ItemStyle-HorizontalAlign="Center" /> 
                                    <asp:boundfield datafield="GPLM" headertext="GPLM" SortExpression="GPLM" /> 
                                    <asp:boundfield datafield="ODM" headertext="ODM" SortExpression="ODM" /> 
                                    <asp:boundfield datafield="ProductLine" headertext="ProductLine" SortExpression="ProductLine" />                                     
                                    <asp:boundfield datafield="ProductLineDescription" headertext="ProductLine Description" SortExpression="ProductLineDescription" />            
                                    <asp:boundfield datafield="FCS" headertext="FCS" SortExpression="FCS" DataFormatString="{0:d}"/>  
                                    <asp:boundfield datafield="EndOfProduction" headertext="End of Manufacturing" SortExpression="EndOfProduction" DataFormatString="{0:d}"/>  
                                    <asp:boundfield datafield="ServiceLifeDate" headertext="End of Serv." SortExpression="ServiceLifeDate" DataFormatString="{0:d}"/>  
                                    <asp:boundfield datafield="ProductVersionID"/>  
                              </Columns>                              
                             </asp:GridView>      
                        </td>
                    </tr>
                  </table>
            </asp:Panel>        
            <asp:Panel runat="server" ID="pnlNotebook" Visible="true"  >
              <table  cellspacing="0" cellpadding="0" width="100%" border="0"  >
                <tr>
                    <td>
                         <asp:GridView runat="server" ID="gvPlatformAssignmentMetrics" Width="100%" AllowPaging="true" PageSize="100"
                           AllowSorting="true" AutoGenerateColumns="false"
                           GridLines="Both"  ShowFooter="false" ShowHeader="true"  CellPadding="5" BorderWidth="2px" 
                           AlternatingRowStyle-CssClass="Table" RowStyle-CssClass="Table"  HeaderStyle-CssClass="TableHeader" BorderColor="tan" >
                           <Columns>
                                <asp:boundfield datafield="Platform" headertext="Platform" SortExpression="Platform" ItemStyle-Wrap="true"  /> 
                                <asp:boundfield datafield="Business" headertext="Business" SortExpression="Business" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="FCS" headertext="FCS" SortExpression="FCS" DataFormatString="{0:d}"/>  
                                <asp:boundfield datafield="GPLM" headertext="GPLM" SortExpression="GPLM" /> 
                                <asp:boundfield datafield="SPDM" headertext="BOM Analysis" SortExpression="SPDM" /> 
                                <asp:boundfield datafield="ODM" headertext="ODM" SortExpression="ODM" /> 
                                <asp:boundfield datafield="PSM" headertext="PSM" SortExpression="PSM" /> 
                                <asp:boundfield datafield="NPI" headertext="NPI" SortExpression="NPI" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="servicefamilypn" headertext="ServiceFamilyPn" SortExpression="servicefamilypn" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="autopublishRsl" headertext="Automated RSL" SortExpression="autopublishRsl" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="active" headertext="Active SPB" SortExpression="active" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="ProjectNumber" headertext="ProjectNumber" SortExpression="ProjectNumber" ItemStyle-HorizontalAlign="Center" /> 
                                <asp:boundfield datafield="M1" headertext="M1" SortExpression="M1" DataFormatString="{0:d}"/> 
                                <asp:boundfield datafield="SpsM1Structured" headertext="Sps Structured" SortExpression="SpsM1Structured" DataFormatString="{0:p}"/> 
                                <asp:boundfield datafield="M2" headertext="M2" SortExpression="M2" DataFormatString="{0:d}"/> 
                                <asp:boundfield datafield="SasM2Released" headertext="Sa's Released" SortExpression="SasM2Released" DataFormatString="{0:p}"/> 
                                <asp:boundfield datafield="M3" headertext="M3" SortExpression="M3" DataFormatString="{0:d}"/> 
                                <asp:boundfield datafield="SpsM3Released" headertext="Sps Released" SortExpression="SpsM3Released" DataFormatString="{0:p}"/> 
                                <asp:BoundField datafield="SPSCount_SFPN" headertext="SPS by SfPn" SortExpression="SPSCount_SFPN" ItemStyle-HorizontalAlign="Center" />
                                <asp:BoundField datafield="SPSSCount_Rev" headertext="SPS by Rev" SortExpression="SPSSCount_Rev" ItemStyle-HorizontalAlign="Center" />                                                       
                           </Columns>                              
                         </asp:GridView>      
                    </td>
                </tr>
             </table>
           </asp:Panel>
        <asp:panel id="pnlNoData" Runat="server">
             <table style="LEFT: 10px; POSITION: absolute; TOP: 300px; HEIGHT: 30px" cellspacing="0" cellpadding="0" width="98%" border="0">
	            <tr style="width:100%;">
		            <td align="center" >
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