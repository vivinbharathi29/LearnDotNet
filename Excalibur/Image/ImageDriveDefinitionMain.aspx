<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.Image_ImageDriveDefinitionMain" Codebehind="ImageDriveDefinitionMain.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <p><asp:Label ID="Label1" runat="server" Text="Master Sku Components" CssClass="Title"></asp:Label></p>
        <asp:GridView ID="gvDrives" runat="server" AutoGenerateColumns="False" EmptyDataText="No Drives Defined"
            CssClass="FormTable" Width="400px" ShowFooter="true">
            <Columns>
                <asp:TemplateField>
                    <EditItemTemplate>
                        <asp:HiddenField runat="server" ID="hfId" Value='<%# Bind("ID") %>' />
                        <asp:LinkButton runat="server" CommandName="Update" Text="Save" ID="lbSave" CausesValidation="true" />
                        <asp:LinkButton runat="server" CommandName="Cancel" Text="Cancel" ID="lbCancel" CausesValidation="false" />
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:LinkButton runat="server" CommandName="Edit" Text="Edit" ID="lbEdit" CausesValidation="false" />
                    </ItemTemplate>
                    <FooterTemplate>
                        <asp:LinkButton runat="server" CommandName="Insert" Text="Add" ID="lbAdd" CausesValidation="true" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Division Cd.">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblDivCd" Text='<%# Eval("DivCd") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox runat="server" ID="tbDivCd" Text='<%# Bind("DivCd") %>' Width="50" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvDivCd" ControlToValidate="tbDivCd"
                            ErrorMessage="Division Cd." Display="None" />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:TextBox runat="server" ID="tbDivCd" Text="PORT" Width="50" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvDivCd" ControlToValidate="tbDivCd"
                            ErrorMessage="Division Cd." Display="None" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Site Cd.">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblSiteCd" Text='<%# Eval("SiteCd") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox runat="server" ID="tbSiteCd" Text='<%# Bind("SiteCd") %>' Width="50" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvSiteCd" ControlToValidate="tbSiteCd"
                            ErrorMessage="Site Cd." Display="None" />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:TextBox runat="server" ID="tbSiteCd" Text="HOU" Width="50" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvSiteCd" ControlToValidate="tbSiteCd"
                            ErrorMessage="Site Cd." Display="None" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Name">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblDriveName" Text='<%# Eval("DriveName") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox runat="server" ID="tbDriveName" Text='<%# Bind("DriveName") %>' />
                        <asp:RequiredFieldValidator runat="server" ID="rfvDriveName" ControlToValidate="tbDriveName"
                            ErrorMessage="Name" Display="None" />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:TextBox runat="server" ID="tbDriveName" Text="" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvDriveName" ControlToValidate="tbDriveName"
                            ErrorMessage="Name" Display="None" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Part No.">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblPartNo" Text='<%# Eval("PartNo") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox runat="server" ID="tbPartNo" Text='<%# Bind("PartNo") %>' Width="80" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvPartNo" ControlToValidate="tbPartNo"
                            ErrorMessage="Part No." Display="None" />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:TextBox runat="server" ID="tbPartNo" Text="" Width="80" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvPartNo" ControlToValidate="tbPartNo"
                            ErrorMessage="Part No." Display="None" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Part No. Rev.">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblPartNoRev" Text='<%# Eval("PartNoRev") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:TextBox runat="server" ID="tbPartNoRev" Text='<%# Bind("PartNoRev") %>' Width="30" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvPartNoRev" ControlToValidate="tbPartNoRev"
                            ErrorMessage="Part No. Rev." Display="None" />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:TextBox runat="server" ID="tbPartNoRev" Text="" Width="30" />
                        <asp:RequiredFieldValidator runat="server" ID="rfvPartNoRev" ControlToValidate="tbPartNoRev"
                            ErrorMessage="Part No. Rev." Display="None" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Is Assembly">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblIsAssembly" Text='<%# Eval("IsAssembly") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:CheckBox runat="server" ID="cbIsAssembly" Checked='<%# Bind("IsAssembly") %>' />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:CheckBox runat="server" ID="cbIsAssembly" />
                    </FooterTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Active">
                    <ItemTemplate>
                        <asp:Label runat="server" ID="lblActive" Text='<%# Eval("Active") %>' />
                    </ItemTemplate>
                    <EditItemTemplate>
                        <asp:CheckBox runat="server" ID="cbActive" Checked='<%# Bind("Active") %>' />
                    </EditItemTemplate>
                    <FooterTemplate>
                        <asp:CheckBox runat="server" ID="cbActive" Visible="false" />
                    </FooterTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>
    <asp:ValidationSummary ID="ValidationSummary1" runat="server" HeaderText="You must enter a value in the following fields:" />
    </form>
</body>
</html>
