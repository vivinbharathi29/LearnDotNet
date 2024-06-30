<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Query_schedule" EnableEventValidation="false" ViewStateEncryptionMode="never" Codebehind="schedule.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Schedule Advanced Search Page</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
    <link href="../style/excalibur.css" rel="stylesheet" type="text/css" />
    <link href="../style/bubble.css" rel="stylesheet" type="text/css" />

    <script type="text/javascript">
<!--
        function getProfileName() {
            var strNewName = window.prompt("Enter a name for the profile.", "");
            if (strNewName != null) {
                form1.hidProfileName.value = strNewName;
                return true;
            }
            else {
                return false;
            }
        }

        function ShareProfile() {
            var strResult;
            strResult = window.showModalDialog("ProfileShare.asp?ID=" + form1.hidProfileId.value, "", "dialogWidth:700px;dialogHeight:400px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            return false;
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

//-->
    </script>
<style type="text/css">
.ProfileOwnerHdr{font: bold xx-small Verdana;}
.ProfileOwnerName{font: xx-small Verdana;}
</style>
</head>
<body id="body" runat="server">
    <form id="form1" runat="server">
    <div>
        <table style="width: 100%;">
            <tr>
                <td colspan="3">
                    <asp:Label runat="server" ID="lblProfile" CssClass="HeaderLabel" Text="Report Profile:" />
                    <asp:DropDownList runat="server" ID="ddlReportProfiles" OnSelectedIndexChanged="ddlReportProfiles_SelectedIndexChanged"
                        AutoPostBack="True" />
                    <asp:LinkButton ID="lbAddProfile" runat="server">Add</asp:LinkButton>
                    <asp:LinkButton ID="lbUpdateProfile" runat="server" Visible="False">Update</asp:LinkButton>
                    <asp:LinkButton ID="lbDeleteProfile" runat="server" Visible="False">Delete</asp:LinkButton>
                    <asp:LinkButton ID="lbRenameProfile" runat="server" Visible="False">Rename</asp:LinkButton>
                    <asp:LinkButton ID="lbShareProfile" runat="server" Visible="False" OnClientClick="ShareProfile();">Share</asp:LinkButton>
                    <asp:LinkButton ID="lbRemoveProfile" runat="server" Visible="False">Remove</asp:LinkButton>
                    <asp:Label ID="lblProfileOwnerHdr" runat="server" Visible="false" Text="Profile Owner:" CssClass="ProfileOwnerHdr" />
                    <asp:Label ID="lblProfileOwnerName" runat="server" Visible="false" CssClass="ProfileOwnerName" />
                    <asp:HiddenField ID="hidProducts" runat="server" />
                    <asp:HiddenField ID="hidGroups" runat="server" />
                    <asp:HiddenField ID="hidProfileName" runat="server" />
                    <asp:HiddenField ID="hidProfileId" runat="server" />
                    <div style="display: none">
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="3" style="padding-top: 10px; padding-bottom: 5px">
                    <asp:Button ID="btnSummaryReport" runat="server" BackColor="#333333" BorderColor="#333333"
                        BorderStyle="Solid" BorderWidth="1px" CssClass="ReportButton" Font-Bold="True"
                        Font-Size="X-Small" ForeColor="White" Height="18px" OnClientClick="window.document.forms[0].target='_blank'; setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='Schedule.aspx';}, 5);"
                        PostBackUrl="ScheduleSummary.aspx" Text="Summary Report" />
                    <asp:Button ID="btnHistoryReport" runat="server" BackColor="#333333" BorderColor="#333333"
                        BorderStyle="Solid" BorderWidth="1px" CssClass="ReportButton" Font-Bold="True"
                        Font-Size="X-Small" ForeColor="White" Height="18px" OnClientClick="window.document.forms[0].target='_blank'; setTimeout(function(){window.document.forms[0].target='';window.document.forms[0].action='Schedule.aspx';}, 5);"
                        PostBackUrl="ScheduleChanges.aspx" Text="History Report" />
                    <asp:Button ID="btnReset" runat="server" BackColor="#333333" BorderColor="#333333"
                        BorderStyle="Solid" BorderWidth="1px" CssClass="ReportButton" Font-Bold="True"
                        Font-Size="X-Small" ForeColor="White" Height="18px" Text="Reset" />
                </td>
            </tr>
            <tr>
                <td style="width: 150px;">
                    <asp:Label ID="lblProducts" runat="server" CssClass="HeaderLabel" Text="Products:"></asp:Label><br />
                    <asp:ListBox ID="lbProducts" runat="server" Height="150" Width="150" SelectionMode="Multiple">
                    </asp:ListBox>
                </td>
                <td style="width: 150px">
                    <asp:Label ID="lblProductGroups" runat="server" CssClass="HeaderLabel" Text="Product Groups:"></asp:Label><br />
                    <asp:ListBox ID="lbProductGroups" runat="server" Height="150" Width="150" SelectionMode="Multiple">
                    </asp:ListBox>
                </td>
                <td style="width: 100%" valign="top">
                    <table>
                        <tr>
                            <td>
                                <asp:Label ID="lblInlcudeSystemTeam" runat="server" CssClass="HeaderLabel" Text="Include System Team:" />
                            </td>
                            <td>
                                <asp:CheckBox ID="cbIncludeSystemTeam" runat="server" Text="" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="lblReportFormat" runat="server" CssClass="HeaderLabel" Text="Report Format:" />
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlReportFormat" runat="server">
                                    <asp:ListItem Value="0" Text="HTML" />
                                    <asp:ListItem Value="1" Text="Excel" />
                                    <asp:ListItem Value="2" Text="Word" />
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>                    
            <tr>
                <td colspan="3">
                    <asp:Button ID="btnRefreshMilestones" runat="server" Text="Refresh Milestones" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:Label ID="lblMilestones" runat="server" CssClass="HeaderLabel" Text="Milestones" />
                    <asp:Label ID="lblMilestoneNote" runat="server" Text="&nbsp;&nbsp;(<strong>Note:</strong> Default milestones are highlighted in green.)" />
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <table class="table" style="width: 100%;">
                        <asp:Repeater ID="rptMilestones" runat="server" OnItemDataBound="rptMilestones_ItemDataBound">
                            <HeaderTemplate>
                                <tr class="th">
                                    <th>
                                        Selected
                                    </th>
                                    <th>
                                        Phase
                                    </th>
                                    <th>
                                        Description
                                    </th>
                                    <th>
                                        Email
                                    </th>
                                    <th>
                                        Action Item
                                    </th>
                                    <th>
                                        Days of Notice
                                    </th>
                                    <th>
                                        Notes to Self
                                    </th>
                                </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr id="tblRow" runat="server" class="td">
                                    <td>
                                        <asp:CheckBox runat="server" ID="cbEnabled" />
                                    </td>
                                    <td>
                                        <%# Eval("phase_name") %>
                                    </td>
                                    <td>
                                        <a href="#" class="tt">
                                            <%# Eval("item_description") %>
                                            <span class="tooltip"><span class="top"></span><span class="middle">
                                                <%# Eval("item_definition") %>
                                            </span><span class="bottom"></span></span></a>
                                    </td>
                                    <td>
                                        <asp:CheckBox runat="server" ID="cbSendEmail" />
                                    </td>
                                    <td>
                                        <asp:CheckBox runat="server" ID="cbCreateAction" />
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ID="tbDaysDiff" Width="50px" AutoPostBack="false"></asp:TextBox>
                                    </td>
                                    <td>
                                        <asp:TextBox runat="server" ID="tbNoteToSelf" Rows="2" MaxLength="1000" Width="500px"
                                            AutoPostBack="false" TextMode="MultiLine"></asp:TextBox>
                                    </td>
                                </tr>
                            </ItemTemplate>
                        </asp:Repeater>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
