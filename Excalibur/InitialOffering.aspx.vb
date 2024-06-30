Imports System.Data
Imports System.Drawing
Partial Class InitialOffering
    Inherits System.Web.UI.Page

    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Private ReadOnly Property Business() As String
        Get
            Return Request.QueryString("Business")
        End Get
    End Property

    Private ReadOnly Property Category() As String
        Get
            Return Request.QueryString("Category")
        End Get
    End Property

    Private ReadOnly Property ProductProgram() As String
        Get
            Return Request.QueryString("ProductProgram")
        End Get
    End Property

    Private ReadOnly Property ProgramText() As String
        Get
            Return Request.QueryString("ProgramText")
        End Get
    End Property

    Private ReadOnly Property CurrentUser() As String
        Get
            Return GetUserId()
        End Get
    End Property
    Public Shared Function GetUserId() As String
        Dim name As String = HttpContext.Current.User.Identity.Name
        Dim userName As String() = name.Split("\")
        Dim dt1 As DataTable = HPQ.Excalibur.Employee.GetUserInfo(userName(1), userName(0))
        Return dt1.Rows(0).Item("ID")
    End Function

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public Shared Property dtData() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtData"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtData", Value)
        End Set
    End Property

    Public Shared Property dtProducts() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtProducts"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtProducts", Value)
        End Set
    End Property

    Public Shared Property dtDeliverables() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtDeliverables"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtDeliverables", Value)
        End Set
    End Property

    Public Shared Property bMarketingIOPlanner() As Boolean
        Get
            Return (GetSessionStateValue("bMarketingIOPlanner"))
        End Get
        Set(ByVal Value As Boolean)
            AddSessionStateValue("bMarketingIOPlanner", Value)
        End Set
    End Property

    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
        Try
            'lblHeader.Text = "text: " & ProductGroup
            If ProductProgram = "" Then
                'Initial Offering
                lbCommodityGuidance.Visible = False
                bMarketingIOPlanner = False
                Dim dt As DataTable = dw.SelectInitialOfferingRoleStatus(CurrentUser)
                If dt.Rows(0).Item("RoleStatus") = 1 Or CurrentUser = 5016 Or CurrentUser = 8 Then
                    bMarketingIOPlanner = True
                Else
                    lbSubmitAVChanges.Visible = False
                    lbExport.Visible = False
                End If

                dtData = dw.SelectInitialOfferingData(Business, Category)
                dtProducts = dw.SelectInitialOfferingProducts(Business)
            Else
                'Commodity Guidance
                bMarketingIOPlanner = False
                Dim dt As DataTable = dw.SelectInitialOfferingRoleStatus(CurrentUser)
                If dt.Rows(0).Item("RoleStatus") = 1 Or CurrentUser = 5016 Or CurrentUser = 8 Then
                    bMarketingIOPlanner = True
                End If
                dtData = dw.SelectCommodityGuidanceData(ProductProgram, Category)
                dtProducts = dw.SelectCommodityGuidanceProductsByProgram(ProductProgram)
                gvIO.Columns(0).Visible = False
                gvIO.Columns(2).Visible = False
                gvIO.Columns(3).Visible = False
                lbCommodityGuidance.InnerText = "Commodity Guidance Report (" & ProgramText & ")"
                lbExport0.Visible = False
                lbExport.Visible = False
                rbStatus.Visible = False
                lbSubmitAVChanges.Visible = False
                lbSubassemblyReport.Visible = False
            End If

            dtDeliverables = dw.SelectInitialOfferingDeliverables(Category)

            AddProducts()

            gvIO.DataSource = dtDeliverables
            gvIO.DataBind()

            PopulateSelectedProducts()

            If ProductProgram = "" Then
                SetDeliverableStatus()
            End If

            If Me.Page.IsPostBack = True Then
                lblHeader.Text = ""
            End If
        Catch ex As Exception
            lblHeader.Text = "Page_Init - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub PopulateSelectedProducts()
        Try
            For Each gvRow As GridViewRow In gvIO.Rows
                Dim lblDRID As System.Web.UI.WebControls.Label = gvRow.FindControl("lblDRID")
                For Each row As DataRow In dtProducts.Rows
                    Dim chk2 As System.Web.UI.WebControls.CheckBox = gvRow.FindControl(row("BID") & "_" & row("PVID"))
                    chk2.ToolTip = lblDRID.Text
                    chk2.CssClass = lblDRID.Text
                Next
            Next

            For Each gvRow As GridViewRow In gvIO.Rows
                Dim lblDRID As System.Web.UI.WebControls.Label = gvRow.FindControl("lblDRID")
                For Each row As DataRow In dtData.Rows
                    If row("DeliverableRootID") = lblDRID.Text Then
                        Dim chk As System.Web.UI.WebControls.CheckBox = gvRow.FindControl(row("BID") & "_" & row("PVID"))
                        If (Not chk Is Nothing) Then
                            If (row("IOGenerated") = False) Then
                                chk.BackColor = ColorTranslator.FromHtml("#6B696B")
                                chk.ToolTip = "Not IO Generated"
                            End If
                            chk.Checked = True
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            lblHeader.Text = "PopulateSelectedProducts - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

    Private Sub AddProducts()
        Try
            For Each row As DataRow In dtProducts.Rows

                Dim ckhColumn As New TemplateField()

                ckhColumn.HeaderTemplate = New InitialOfferingGridViewTemplate(ListItemType.Header, row("FullName"))
                ckhColumn.ItemTemplate = New InitialOfferingGridViewTemplate(ListItemType.Item, row("BID") & "_" & row("PVID"), row("PVID"), row("BID"), bMarketingIOPlanner)
                ckhColumn.HeaderStyle.Wrap = False

                gvIO.Columns.Add(ckhColumn)
            Next
        Catch ex As Exception
            lblHeader.Text = "AddProducts - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

    Private Sub SetDeliverableStatus()
        Try
            For Each gvRow As GridViewRow In gvIO.Rows
                Dim lblStatus As System.Web.UI.WebControls.Label = gvRow.FindControl("lblStatus")
                Dim lblChangeStatus As System.Web.UI.WebControls.Label = gvRow.FindControl("lblChangeStatus")
                Dim cbxSelect As System.Web.UI.WebControls.CheckBox = gvRow.FindControl("cbxSelect")
                Dim lblDelDescr As System.Web.UI.WebControls.Label = gvRow.FindControl("lblDelDescr")

                If bMarketingIOPlanner = False Then
                    cbxSelect.Enabled = False
                End If

                Select Case lblChangeStatus.Text
                    Case 1 'Add/Remove ProductBrand
                        gvRow.BackColor = Drawing.Color.MistyRose
                    Case 2 'Add Deliverable
                        gvRow.Cells(1).BackColor = Drawing.Color.LightSteelBlue
                    Case 3 'Remove Deliverable
                        lblDelDescr.Font.Strikeout = True
                        'System.Drawing.ColorTranslator.FromHtml("#DDDDDD")
                End Select

                Select Case lblStatus.Text
                    Case 0 'New
                        lblDelDescr.Font.Bold = True
                        cbxSelect.Checked = False
                    Case 1 'Selected
                        cbxSelect.Checked = True
                    Case 2 'Not Selected
                        cbxSelect.Checked = False
                End Select
            Next
        Catch ex As Exception
            lblHeader.Text = "SetDeliverableStatus - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

    Protected Sub rbStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbStatus.SelectedIndexChanged
        Try
            For Each gvRow As GridViewRow In gvIO.Rows
                Dim lblStatus As System.Web.UI.WebControls.Label = gvRow.FindControl("lblStatus")
                If rbStatus.SelectedItem.Value = 0 Then 'All
                    gvRow.Visible = True
                Else
                    If lblStatus.Text = 1 Then
                        gvRow.Visible = True
                    Else
                        gvRow.Visible = False
                    End If
                End If
            Next
        Catch ex As Exception
            lblHeader.Text = "rbStatus_SelectedIndexChanged - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

    Protected Sub lbSubmitAVChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbSubmitAVChanges.Click
        Try
            dw.UpdateInitialOfferingAVChanges(Business, CurrentUser)
            lblHeader.Text = "AV Changes Submitted Successfully"
        Catch ex As Exception
            lblHeader.Text = "lbPublish_Click - " & ex.Message.ToString
            lblHeader.ForeColor = Drawing.Color.Red
        End Try
    End Sub

End Class
