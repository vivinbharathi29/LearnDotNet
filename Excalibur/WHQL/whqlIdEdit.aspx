<%@ Page Language="VB" %>

<%@ Register Assembly="eWorld.UI, Version=2.0.6.2393, Culture=neutral, PublicKeyToken=24d65337282035f2"
    Namespace="eWorld.UI" TagPrefix="ew" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Function NullTest(ByVal testval As Object) As Nullable(Of Date)
        
        If Not IsDBNull(testval) Then
            Return testval
        Else
            Return Nothing
        End If
        
    End Function
    
    Protected Function SetVisibleDate(ByVal TestVal As Object) As Date
        If Not IsDBNull(TestVal) Then
            Return TestVal
        Else
            Return Now()
        End If
    End Function

    Protected Sub dtvWhqlSubmission_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdateEventArgs)
        e.NewValues("SubmissionDt") = CType(dtvWhqlSubmission.FindControl("cpSubmissionDt"), eWorld.UI.CalendarPopup).SelectedValue
        e.NewValues("WhqlDt") = CType(dtvWhqlSubmission.FindControl("cpWhqlDt"), eWorld.UI.CalendarPopup).SelectedValue
        e.NewValues("DateReleased") = CType(dtvWhqlSubmission.FindControl("cpReleaseDt"), eWorld.UI.CalendarPopup).SelectedValue
    End Sub

    Protected Sub dtvWhqlSubmission_ModeChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link rel="stylesheet" type="text/css" href="/style/excalibur.css">
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <p>
                <asp:Label ID="lblDetails" runat="server" CssClass="Heading" Text="Submission Details"></asp:Label><br />
                <asp:DetailsView ID="dtvWhqlSubmission" runat="server" AutoGenerateRows="False" CssClass="FormTable"
                    Width="100%" DataSourceID="odsProductWhql" OnItemUpdating="dtvWhqlSubmission_ItemUpdating" OnModeChanged="dtvWhqlSubmission_ModeChanged">
                    <FieldHeaderStyle Width="125px" Font-Bold="True" />
                    <Fields>
                        <asp:BoundField DataField="SubmissionID" HeaderText="Submission ID:" />
                        <asp:TemplateField HeaderText="Submission Date:">
                            <EditItemTemplate>
                                <ew:CalendarPopup ID="cpSubmissionDt" runat="server" Nullable="True" ImageUrl="~/images/calendar.gif"
                                    ShowClearDate="true" ShowGoToToday="true"  
                                    SelectedValue='<%# NullTest(DataBinder.Eval(Container.DataItem, "SubmissionDt")) %>' 
                                    VisibleDate='<%# SetVisibleDate(DataBinder.Eval(Container.DataItem, "SubmissionDt")) %>'>
                                </ew:CalendarPopup>
                            </EditItemTemplate>
                            <InsertItemTemplate>
                                <asp:TextBox ID="txtSubmissionDt" runat="server" Text='<%# Bind("SubmissionDt") %>'></asp:TextBox>
                            </InsertItemTemplate>
                            <ItemTemplate>
                                <asp:Label ID="Label2" runat="server" Text='<%# Bind("SubmissionDt", "{0:d}") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="WHQL Approval Date:">
                            <EditItemTemplate>
                                <ew:CalendarPopup ID="cpWhqlDt" runat="server" Nullable="True" ImageUrl="~/images/calendar.gif"
                                    ShowClearDate="true" ShowGoToToday="true" 
                                    SelectedValue='<%# NullTest(DataBinder.Eval(Container.DataItem, "WhqlDt")) %>' 
                                    VisibleDate='<%# SetVisibleDate(DataBinder.Eval(Container.DataItem, "WhqlDt")) %>'>
                                </ew:CalendarPopup>
                            </EditItemTemplate>
                            <InsertItemTemplate>
                                <asp:TextBox ID="txtWhqlDt" runat="server" Text='<%# Bind("WhqlDt") %>'></asp:TextBox>
                            </InsertItemTemplate>
                            <ItemTemplate>
                                <asp:Label ID="Label3" runat="server" Text='<%# Bind("WhqlDt", "{0:d}") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Product Release Date:">
                            <EditItemTemplate>
                                <ew:CalendarPopup ID="cpReleaseDt" runat="server" Nullable="True" ImageUrl="~/images/calendar.gif"
                                    ShowClearDate="true" ShowGoToToday="true" 
                                    SelectedValue='<%# NullTest(DataBinder.Eval(Container.DataItem, "DateReleased")) %>' 
                                    VisibleDate='<%# SetVisibleDate(DataBinder.Eval(Container.DataItem, "DateReleased")) %>'>
                                </ew:CalendarPopup>
                            </EditItemTemplate>
                            <InsertItemTemplate>
                                <asp:TextBox ID="txtReleaseDt" runat="server" Text='<%# Bind("DateReleased") %>'></asp:TextBox>
                            </InsertItemTemplate>
                            <ItemTemplate>
                                <asp:Label ID="Label4" runat="server" Text='<%# Bind("DateReleased", "{0:d}") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Label Location:">
                            <EditItemTemplate>
                                <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Location") %>' Rows="2"
                                    Columns="40" TextMode="MultiLine"></asp:TextBox>
                            </EditItemTemplate>
                            <InsertItemTemplate>
                                <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Location") %>'></asp:TextBox>
                            </InsertItemTemplate>
                            <ItemTemplate>
                                <asp:Label ID="Label1" runat="server" Text='<%# Bind("Location") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:CheckBoxField DataField="LogoDisplayed" HeaderText="Logo Displayed:" />
                        <asp:CheckBoxField DataField="Milestone3" HeaderText="Milestone 3 Compliant:" />
                        <asp:CheckBoxField DataField="BootVisComplete" HeaderText="BootVis Complete:" />
                        <asp:CommandField ShowEditButton="True" />
                    </Fields>
                </asp:DetailsView>
                <asp:ObjectDataSource ID="odsProductWhql" runat="server" OldValuesParameterFormatString="original_{0}"
                    SelectMethod="SelectProductWhql" TypeName="HPQ.Excalibur.Data" UpdateMethod="UpdateProductWhql">
                    <UpdateParameters>
                        <asp:QueryStringParameter Name="ProductWhqlID" QueryStringField="WHQLID" Type="String" />
                        <asp:Parameter Name="SubmissionID" Type="String" />
                        <asp:Parameter Name="SubmissionDt" Type="String" />
                        <asp:Parameter Name="WhqlDt" Type="String" />
                        <asp:Parameter Name="DateReleased" Type="String" />
                        <asp:Parameter Name="Location" Type="String" />
                        <asp:Parameter Name="LogoDisplayed" Type="String" />
                        <asp:Parameter Name="Milestone3" Type="String" />
                        <asp:Parameter Name="BootVisComplete" Type="String" />
                    </UpdateParameters>
                    <SelectParameters>
                        <asp:QueryStringParameter Name="ProductWhqlID" QueryStringField="WHQLID" Type="String" />
                        <asp:Parameter Name="ProductVersionID" Type="String" />
                    </SelectParameters>
                </asp:ObjectDataSource>
            </p>
            <p>
                <asp:Label ID="lblSeries" runat="server" CssClass="Heading" Text="Series Included"></asp:Label>
                &nbsp;&nbsp;
                <asp:Label ID="lblSeriesText" runat="server" Text="Label"></asp:Label></p>
            <p>
                <asp:Label ID="lblCreate" runat="server" CssClass="Heading" Text="Base Unit, Processor, & OS Family List"></asp:Label>&nbsp;<asp:GridView
                    ID="GridView1" runat="server" AutoGenerateColumns="False" CssClass="FormTable"
                    DataSourceID="odsWhqlSubmissions" Width="100%">
                    <Columns>
                        <asp:BoundField HeaderText="OS Family" DataField="OS" />
                        <asp:BoundField HeaderText="Base Unit" DataField="BU" />
                        <asp:BoundField HeaderText="Processor" DataField="CPU" />
                    </Columns>
                </asp:GridView>
                <asp:ObjectDataSource ID="odsWhqlSubmissions" runat="server" OldValuesParameterFormatString="original_{0}"
                    SelectMethod="SelectWhqlSubmissions" TypeName="HPQ.Excalibur.Data">
                    <SelectParameters>
                        <asp:QueryStringParameter Name="ProductWhqlID" QueryStringField="WHQLID" Type="String" />
                    </SelectParameters>
                </asp:ObjectDataSource>
            </p>
            <p>
                &nbsp;<asp:Label ID="Label1" runat="server" CssClass="Heading" Text="Models Included"></asp:Label>
                <asp:DataList ID="DataList1" runat="server" CssClass="FormTable" DataSourceID="odsWhqlModelList"
                    RepeatColumns="3" ShowFooter="False" ShowHeader="False">
                    <ItemTemplate>
                        <%#DataBinder.Eval(Container.DataItem, "SkuModel")%>
                    </ItemTemplate>
                </asp:DataList>
                <asp:ObjectDataSource ID="odsWhqlModelList" runat="server" OldValuesParameterFormatString="original_{0}"
                    SelectMethod="ListWhqlModels" TypeName="HPQ.Excalibur.Data">
                    <SelectParameters>
                        <asp:QueryStringParameter DefaultValue="48" Name="ProductWhqlID" QueryStringField="WHQLID"
                            Type="String" />
                    </SelectParameters>
                </asp:ObjectDataSource>
            </p>
        </div>
    </form>
</body>
</html>
