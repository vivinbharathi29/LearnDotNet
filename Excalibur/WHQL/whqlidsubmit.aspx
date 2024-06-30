<%@ Page Language="VB" %>

<%@ Register Assembly="eWorld.UI, Version=2.0.6.2393, Culture=neutral, PublicKeyToken=24d65337282035f2"
    Namespace="eWorld.UI" TagPrefix="ew" %>

<%@ Register Assembly="skmValidators" Namespace="skmValidators" TagPrefix="cc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Dim sWhqlSubmission As String
    Dim sBaseUnitCpuCombo As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not Page.IsPostBack Then

            Wizard1.ActiveStepIndex = 0

            txtSubmissionDt.ShowClearDate = True
            txtSubmissionDt.ControlDisplay = DisplayType.TextBoxImage
            txtSubmissionDt.VisibleDate = Now()

            txtReleaseDt.ShowClearDate = True
            txtReleaseDt.ControlDisplay = DisplayType.TextBoxImage
            txtReleaseDt.VisibleDate = Now()

        End If

    End Sub

    Private Sub SetNoAvAlerts(ByVal AvType As String)
        lblErrorMessage.Text = String.Format("{0} AV Information is not present.  Unable to process Submission IDs at this time.", AvType)
        lblErrorMessage.Visible = True
        Wizard1.Enabled = False
    End Sub

    Protected Sub Wizard1_FinishButtonClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.WizardNavigationEventArgs)
        Dim iProductWhqlID As Integer
        Dim dl As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        iProductWhqlID = dl.InsertProductWhql(txtSubmissionID.Text.Trim(), txtSubmissionDt.SelectedValue.ToString(), _
            String.Empty, "1", Request.QueryString("PVID"), txtLabelLocation.Text.Trim(), _
            txtReleaseDt.SelectedValue.ToString, chkLogoDisplayed.Checked, chkMilestone3.Checked)

        If iProductWhqlID > 0 Then

            For iSeries As Integer = 0 To cblSeries.Items.Count - 1
                If cblSeries.Items(iSeries).Selected Then
                    dl.InsertProductWhqlSeries(iProductWhqlID, cblSeries.Items(iSeries).Value)
                End If
            Next

            For iOs As Integer = 0 To cblOsFamily.Items.Count - 1
                If cblOsFamily.Items(iOs).Selected Then
                    For iBu As Integer = 0 To cblBaseUnits.Items.Count - 1
                        If cblBaseUnits.Items(iBu).Selected Then
                            For iCpu As Integer = 0 To cblProcessors.Items.Count - 1
                                If cblProcessors.Items(iCpu).Selected Then
                                    dl.InsertWhqlAvDetail(iProductWhqlID, cblBaseUnits.Items(iBu).Value, cblProcessors.Items(iCpu).Value, cblOsFamily.Items(iOs).Value, rblBrand.SelectedValue)
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If

        Page.ClientScript.RegisterStartupScript(Me.GetType, "_CloseWindow", "function CloseWindow() { window.parent.close(); }", True)
        PageBody.Attributes("onload") = "CloseWindow()"

    End Sub

    Protected Sub Wizard1_NextButtonClick(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.WizardNavigationEventArgs)
        Select Case e.CurrentStepIndex
            Case 0  ' Submission Information
            Case 1  ' Brand
            Case 2  ' Series
            Case 3  ' OS
            Case 4  ' Base Units
            Case 5  ' CPU

                Dim tCell As TableCell
                Dim tRow As TableRow
                Dim dl As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
                Dim iWhqlSubmissionID As Integer = dl.GetProductWhqlID(txtSubmissionID.Text.Trim())
                Dim dtWhqlSubmission As System.Data.DataTable
                Dim dRow As System.Data.DataRow
                Dim sbSeriesList As StringBuilder = New StringBuilder()
                If iWhqlSubmissionID > 0 Then
                    dtWhqlSubmission = dl.SelectWhqlSubmissions(iWhqlSubmissionID)
                Else
                    dtWhqlSubmission = New System.Data.DataTable()
                    dtWhqlSubmission.Columns.Add("SubmissionID")
                    dtWhqlSubmission.Columns.Add("SubmissionDt")
                    dtWhqlSubmission.Columns.Add("DateReleased")
                    dtWhqlSubmission.Columns.Add("Location")
                    dtWhqlSubmission.Columns.Add("LogoDisplayed")
                    dtWhqlSubmission.Columns.Add("Milestone3")
                    dRow = dtWhqlSubmission.NewRow()
                    dRow("SubmissionID") = txtSubmissionID.Text.Trim()
                    dRow("SubmissionDt") = txtSubmissionDt.SelectedDate.ToShortDateString()
                    If Not IsNothing(txtReleaseDt.SelectedValue) Then
                        dRow("DateReleased") = txtReleaseDt.SelectedDate.ToShortDateString()
                    End If
                    dRow("Location") = txtLabelLocation.Text.Trim()
                    dRow("LogoDisplayed") = chkLogoDisplayed.Checked()
                    dRow("Milestone3") = chkMilestone3.Checked()
                    dtWhqlSubmission.Rows.Add(dRow)
                End If

                dtvWhqlSubmission.DataSource = dtWhqlSubmission
                dtvWhqlSubmission.DataBind()

                For iSeries As Integer = 0 To cblSeries.Items.Count - 1
                    If cblSeries.Items(iSeries).Selected Then
                        sbSeriesList.Append(cblSeries.Items(iSeries).Text)
                        sbSeriesList.Append(", ")
                    End If
                Next

                lblSeriesList.Text = sbSeriesList.Remove(sbSeriesList.Length - 2, 2).ToString()

                For iOs As Integer = 0 To cblOsFamily.Items.Count - 1
                    If cblOsFamily.Items(iOs).Selected Then
                        For iBu As Integer = 0 To cblBaseUnits.Items.Count - 1
                            If cblBaseUnits.Items(iBu).Selected Then
                                For iCpu As Integer = 0 To cblProcessors.Items.Count - 1
                                    If cblProcessors.Items(iCpu).Selected Then
                                        tRow = New TableRow
                                        tCell = New TableCell
                                        tCell.Text = txtSubmissionID.Text
                                        tRow.Cells.Add(tCell)
                                        tCell = New TableCell
                                        tCell.Text = cblOsFamily.Items(iOs).Text
                                        tRow.Cells.Add(tCell)
                                        tCell = New TableCell
                                        tCell.Text = cblBaseUnits.Items(iBu).Text
                                        tRow.Cells.Add(tCell)
                                        tCell = New TableCell
                                        tCell.Text = cblProcessors.Items(iCpu).Text
                                        tRow.Cells.Add(tCell)

                                        If dl.GetWhqlAvDetailCount(iWhqlSubmissionID, cblBaseUnits.Items(iBu).Value, cblProcessors.Items(iCpu).Value, cblOsFamily.Items(iOs).Value) = 0 Then
                                            tblSubmissionPreview.Rows.Add(tRow)
                                        Else
                                            tblSubmissionErrors.Rows.Add(tRow)
                                            pnlErrors.CssClass = ""
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                Next


            Case 6  ' Preview
            Case 7  ' Finished
        End Select
    End Sub

    Protected Sub Wizard1_CancelButtonClick(ByVal sender As Object, ByVal e As System.EventArgs)

        Page.ClientScript.RegisterStartupScript(Me.GetType, "_CloseWindow", "function CloseWindow() { window.parent.close(); }", True)
        PageBody.Attributes("onload") = "CloseWindow()"

    End Sub

    Protected Sub rblBrand_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        odsBaseUnits.Select()
        odsProcessors.Select()
        odsSeriesList.Select()
        cblBaseUnits.DataBind()
        cblProcessors.DataBind()
        cblSeries.DataBind()
    End Sub


    Protected Sub odsBaseUnits_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ObjectDataSourceStatusEventArgs)
        Dim dtCount As New System.Data.DataTable()
        dtCount = e.ReturnValue
        If dtCount.Rows.Count = 0 Then
            SetNoAvAlerts("Base Unit")
        End If
    End Sub

    Protected Sub odsProcessors_Selected(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.ObjectDataSourceStatusEventArgs)
        Dim dtcount As New System.Data.DataTable()
        dtcount = e.ReturnValue
        If dtcount.Rows.Count = 0 Then
            'SetNoAvAlerts("Processor")
            'CheckBoxListValidator2.MinimumNumberOfSelectedCheckBoxes = 0
        End If
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Untitled Page</title>
    <link rel="stylesheet" type="text/css" href="/style/excalibur.css">
</head>
<body id="PageBody" runat="server">
    <form id="form1" runat="server" language="javascript" ondrop="return form1_ondrop()">
        <div>
            <asp:Label ID="lblErrorMessage" runat="server" Text="Label" Visible="false" CssClass="ErrorText"></asp:Label>
            <asp:Wizard ID="Wizard1" runat="server" Width="100%" ActiveStepIndex="3" DisplaySideBar="False"
                OnFinishButtonClick="Wizard1_FinishButtonClick" OnNextButtonClick="Wizard1_NextButtonClick"
                Height="575px" DisplayCancelButton="True" OnCancelButtonClick="Wizard1_CancelButtonClick">
                <StepStyle VerticalAlign="Top" BorderStyle="None" />
                <WizardSteps>
                    <asp:WizardStep ID="WizardStep1" runat="server" Title="WHQL Submission">
                        <fieldset>
                            <legend>WHQL Submission</legend>
                            <br />
                            <table width="100%" class="FormTable">
                                <tr>
                                    <th>
                                        Submission ID:</th>
                                    <td>
                                        <asp:TextBox ID="txtSubmissionID" runat="server"></asp:TextBox>
                                        &nbsp;
                                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtSubmissionID"
                                            ErrorMessage="ID is required">ID is required</asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <th>
                                        Submission Date:</th>
                                    <td>
                                        <ew:CalendarPopup ID="txtSubmissionDt" runat="server" ImageUrl="../images/calendar.gif"
                                            Nullable="True" SelectedDate="" VisibleDate="" PostedDate="" UpperBoundDate="12/31/9999 23:59:59">
                                        </ew:CalendarPopup>
                                        &nbsp;
                                        <ew:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtSubmissionDt"
                                            ErrorMessage="Submission Date is Required">Submission Date is Required</ew:RequiredFieldValidator>
                                    </td>
                                </tr>
                                <tr>
                                    <th>
                                        Release Date:</th>
                                    <td>
                                        <ew:CalendarPopup ID="txtReleaseDt" runat="server" ImageUrl="../images/calendar.gif"
                                            Nullable="True" SelectedDate="" VisibleDate="" PostedDate="" UpperBoundDate="12/31/9999 23:59:59">
                                        </ew:CalendarPopup>
                                    </td>
                                </tr>
                                <tr>
                                    <th>
                                        Label Location:</th>
                                    <td>
                                        <asp:TextBox ID="txtLabelLocation" runat="server" Rows="2" TextMode="MultiLine" Text="Monitor &amp; Printed on Service Tag located on the bottom of unit." /></td>
                                </tr>
                                <tr>
                                    <th>
                                        Logo Displayed:</th>
                                    <td>
                                        <asp:CheckBox ID="chkLogoDisplayed" runat="server" /></td>
                                </tr>
                                <tr>
                                    <th>
                                        Milestone 3:</th>
                                    <td>
                                        <asp:CheckBox ID="chkMilestone3" runat="server" /></td>
                                </tr>
                            </table>
                            <br />
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep2" runat="server" Title="Product Brand">
                        <fieldset>
                            <legend>Product Brand</legend>
                            <br />
                            <table class="FormTable" cellpadding="1" cellspacing="0">
                                <tr>
                                    <th>
                                        Brand:</th>
                                    <td>
                                        <asp:RadioButtonList ID="rblBrand" runat="server" DataSourceID="odsBrands" DataTextField="BrandFullName"
                                            DataValueField="ProductBrandID" OnSelectedIndexChanged="rblBrand_SelectedIndexChanged">
                                        </asp:RadioButtonList><asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server"
                                            ErrorMessage="Brand selection is required." ControlToValidate="rblBrand">Brand selection is required.</asp:RequiredFieldValidator>
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:ObjectDataSource ID="odsBrands" runat="server" SelectMethod="ListProductBrands"
                                TypeName="HPQ.Excalibur.Data">
                                <SelectParameters>
                                    <asp:QueryStringParameter DefaultValue="100" Name="ProductVersionID" QueryStringField="PVID"
                                        Type="String" />
                                </SelectParameters>
                            </asp:ObjectDataSource>
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="wsSeries" runat="server" Title="Product Series">
                        <fieldset>
                            <legend>Product Series</legend>
                            <br />
                            <table class="FormTable" cellpadding="1" cellspacing="0">
                                <tr>
                                    <th>
                                        Series:</th>
                                    <td>
                                        <asp:CheckBoxList ID="cblSeries" runat="server" DataSourceID="odsSeriesList" DataTextField="Name"
                                            DataValueField="ID">
                                        </asp:CheckBoxList>
                                        <cc1:CheckBoxListValidator ID="CheckBoxListValidator3" runat="server" ControlToValidate="cblSeries"
                                            ErrorMessage="Series selection is required." MinimumNumberOfSelectedCheckBoxes="1">Series selection is required.</cc1:CheckBoxListValidator>
                                        <asp:ObjectDataSource ID="odsSeriesList" runat="server" OldValuesParameterFormatString="original_{0}"
                                            SelectMethod="ListBrandSeries" TypeName="HPQ.Excalibur.Data">
                                            <SelectParameters>
                                                <asp:ControlParameter ControlID="rblBrand" Name="ProductBrandID" PropertyName="SelectedValue"
                                                    Type="String" />
                                            </SelectParameters>
                                        </asp:ObjectDataSource>
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep3" runat="server" Title="OS Family">
                        <fieldset>
                            <legend>OS Family</legend>
                            <br />
                            <table class="FormTable" cellpadding="1" cellspacing="0">
                                <tr>
                                    <th>
                                        OS Family:</th>
                                    <td>
                                        <asp:CheckBoxList ID="cblOsFamily" runat="server" DataSourceID="odsOsFamilies" DataTextField="FamilyName"
                                            DataValueField="ID">
                                        </asp:CheckBoxList>
                                        <cc1:CheckBoxListValidator ID="cblvOsFamily" runat="server" ControlToValidate="cblOsFamily"
                                            ErrorMessage="OS Family Selection is Required" MinimumNumberOfSelectedCheckBoxes="1">OS Family Selection is Required</cc1:CheckBoxListValidator>
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:ObjectDataSource ID="odsOsFamilies" runat="server" SelectMethod="ListOsFamilies"
                                TypeName="HPQ.Excalibur.Data"></asp:ObjectDataSource>
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep4" runat="server" Title="Base Units Covered">
                        <fieldset>
                            <legend>Base Units</legend>
                            <br />
                            <table class="FormTable" cellpadding="1" cellspacing="0">
                                <tr>
                                    <th>
                                        Base Units:</th>
                                    <td>
                                        <asp:CheckBoxList ID="cblBaseUnits" runat="server" DataSourceID="odsBaseUnits" DataTextField="GPGDescription"
                                            DataValueField="AvDetailID">
                                        </asp:CheckBoxList>
                                        <cc1:CheckBoxListValidator ID="CheckBoxListValidator1" runat="server" ControlToValidate="cblBaseUnits"
                                            ErrorMessage="Base Unit(s) selection is required." MinimumNumberOfSelectedCheckBoxes="1">Base Unit(s) selection is required.</cc1:CheckBoxListValidator>
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:ObjectDataSource ID="odsBaseUnits" runat="server" OldValuesParameterFormatString="original_{0}"
                                SelectMethod="ListProductBrandBaseUnitsAvDetail" TypeName="HPQ.Excalibur.Data" OnSelected="odsBaseUnits_Selected">
                                <SelectParameters>
                                    <asp:ControlParameter ControlID="rblBrand" Name="ProductBrandID" PropertyName="SelectedValue"
                                        Type="String" />
                                </SelectParameters>
                            </asp:ObjectDataSource>
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep5" runat="server" Title="Processors Covered">
                        <fieldset>
                            <legend>Processors</legend>
                            <br />
                            <table class="FormTable" cellpadding="1" cellspacing="0">
                                <tr>
                                    <th>
                                        Processors:</th>
                                    <td>
                                        &nbsp;<asp:CheckBoxList ID="cblProcessors" runat="server" DataSourceID="odsProcessors"
                                            DataTextField="GPGDescription" DataValueField="AvDetailID">
                                        </asp:CheckBoxList>
                                        <cc1:CheckBoxListValidator ID="CheckBoxListValidator2" runat="server" ControlToValidate="cblProcessors"
                                            ErrorMessage="Processor Selection is Required" MinimumNumberOfSelectedCheckBoxes="1">Processor Selection is Required</cc1:CheckBoxListValidator>
                                    </td>
                                </tr>
                            </table>
                            <br />
                            <asp:ObjectDataSource ID="odsProcessors" runat="server" OldValuesParameterFormatString="original_{0}"
                                SelectMethod="ListProductBrandCpuAvDetail" TypeName="HPQ.Excalibur.Data" OnSelected="odsProcessors_Selected">
                                <SelectParameters>
                                    <asp:ControlParameter ControlID="rblBrand" Name="ProductBrandID" PropertyName="SelectedValue"
                                        Type="String" />
                                </SelectParameters>
                            </asp:ObjectDataSource>
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep6" runat="server" StepType="Finish" Title="Submission Preview">
                        <fieldset>
                            <legend>Submission Preview</legend>
                            <br />
                            <asp:DetailsView ID="dtvWhqlSubmission" runat="server" AutoGenerateRows="False" CssClass="FormTable"
                                Width="100%">
                                <FieldHeaderStyle Width="125px" Font-Bold="True" />
                                <Fields>
                                    <asp:BoundField DataField="SubmissionID" HeaderText="Submission ID:" />
                                    <asp:BoundField DataField="SubmissionDt" HeaderText="Submission Date:" />
                                    <asp:BoundField DataField="DateReleased" HeaderText="Release Date:" />
                                    <asp:BoundField DataField="Location" HeaderText="Label Location:" />
                                    <asp:CheckBoxField DataField="LogoDisplayed" HeaderText="Logo Displayed:" />
                                    <asp:CheckBoxField DataField="Milestone3" HeaderText="Milestone 3 Compliant:" />
                                </Fields>
                            </asp:DetailsView>
                            <br />
                            <asp:Label ID="lblSeries" runat="server" CssClass="Heading" Text="These Series will be linked"></asp:Label>
                            <br />
                            <asp:Label ID="lblSeriesList" runat="server" CssClass="FormTable" Width="100%"></asp:Label>
                            <br />
                            <asp:Label ID="lblCreate" runat="server" CssClass="Heading" Text="These records will be created."></asp:Label>
                            <asp:Table ID="tblSubmissionPreview" runat="server" CssClass="FormTable" Width="100%"
                                CellSpacing="0">
                                <asp:TableRow ID="TableRow1" runat="server" TableSection="TableHeader">
                                    <asp:TableCell ID="TableCell1" runat="server" Width="25%">Submission ID</asp:TableCell>
                                    <asp:TableCell ID="TableCell2" runat="server" Width="25%">OS Family</asp:TableCell>
                                    <asp:TableCell ID="TableCell3" runat="server" Width="25%">Base Unit</asp:TableCell>
                                    <asp:TableCell ID="TableCell4" runat="server" Width="25%">Processor</asp:TableCell>
                                </asp:TableRow>
                            </asp:Table>
                            <asp:Panel ID="pnlErrors" runat="server" CssClass="Hidden" Width="100%">
                                <br />
                                <asp:Label ID="lblErrors" runat="server" Text="These records are duplicates and will not be created."
                                    CssClass="ErrorHeading"></asp:Label>
                                <asp:Table ID="tblSubmissionErrors" runat="server" CssClass="FormTable" Width="100%"
                                    CellSpacing="0">
                                    <asp:TableRow ID="TableRow2" runat="server" TableSection="TableHeader">
                                        <asp:TableCell ID="TableCell5" runat="server" Width="25%">Submission ID</asp:TableCell>
                                        <asp:TableCell ID="TableCell6" runat="server" Width="25%">OS Family</asp:TableCell>
                                        <asp:TableCell ID="TableCell7" runat="server" Width="25%">Base Unit</asp:TableCell>
                                        <asp:TableCell ID="TableCell8" runat="server" Width="25%">Processor</asp:TableCell>
                                    </asp:TableRow>
                                </asp:Table>
                            </asp:Panel>
                            <br />
                        </fieldset>
                    </asp:WizardStep>
                    <asp:WizardStep ID="WizardStep7" runat="server" StepType="Complete" Title="Complete">
                    </asp:WizardStep>
                </WizardSteps>
                <NavigationStyle BackColor="Cornsilk" BorderColor="DarkGray" />
            </asp:Wizard>
        </div>
    </form>
</body>
</html>
