Public Class changelog
    Inherits System.Web.UI.Page

    Private _productVersionId As Integer = 0
    Private ReadOnly Property ProductVersionID() As Integer
        Get
            If _productVersionId = 0 Then
                _productVersionId = Convert.ToInt32(Request.QueryString("ID"))
            End If
            Return _productVersionId
        End Get
    End Property

    Private _productBrandId As Integer = 0
    Private ReadOnly Property ProductBrandId() As Integer
        Get
            If _productBrandId = 0 Then
                _productBrandId = Convert.ToInt32(Request.QueryString("BID"))
            End If
            Return _productBrandId
        End Get
    End Property

    Private _brandName = String.Empty
    Private Property BrandName() As String
        Get
            Return _brandName
        End Get
        Set(ByVal value As String)
            _brandName = value
        End Set
    End Property

    Private _isPc As Boolean = Nothing
    Private ReadOnly Property IsPC() As Boolean
        Get
            'If IsNothing(_isPc) Then
            Dim secObj As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(User.Identity.Name)
            '_isPc = (secObj.IsProgramCoordinator Or secObj.IsSysAdmin)
            'End If
            Return (secObj.IsProgramCoordinator Or secObj.IsSysAdmin)
        End Get
    End Property

    Private Property ShowAll() As Boolean
        Get
            If ViewState("ShowAll") Is Nothing Then
                Return False
            Else
                Return Convert.ToBoolean(ViewState("ShowAll").ToString())
            End If
        End Get
        Set(ByVal value As Boolean)
            ViewState("ShowAll") = value
        End Set
    End Property

    Private Property AllowPaging() As Boolean
        Get

            If ViewState("AllowPaging") Is Nothing Then
                Return True
            Else
                Return Convert.ToBoolean(ViewState("AllowPaging").ToString())
            End If
        End Get
        Set(ByVal value As Boolean)
            ViewState("AllowPaging") = value
        End Set
    End Property

    Private Property CurrentPage() As Integer
        Get
            Dim o As Object = ViewState("_CurrentPage")
            If o = Nothing Then
                Return 1
            Else
                Return Convert.ToInt32(o)
            End If
        End Get
        Set(ByVal value As Integer)
            ViewState("_CurrentPage") = value
        End Set
    End Property

    Private Sub LoadProductBrands()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable = dw.ListBrands4Product(ProductVersionID, 1)

        rptrBrands.DataSource = dt
        rptrBrands.DataBind()

        lblBrand.Text = String.Format("{0} SCM Change Log:", BrandName)

    End Sub

    Private Sub LoadChangeHistory()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable = dw.SelectAvHistory(ProductBrandId, String.Empty, String.Empty, String.Empty, ShowAll.ToString(), "90")
        Dim pds As PagedDataSource = New PagedDataSource()
        dt.DefaultView.Sort = "last_upd_date desc"
        pds.DataSource = dt.DefaultView()
        pds.AllowPaging = AllowPaging
        pds.PageSize = 50
        pds.CurrentPageIndex = CurrentPage - 1

        lblCurrentPage.Text = String.Format("Page: {0} of {1}", CurrentPage.ToString(), pds.PageCount.ToString())
        lbPrev.Enabled = Not pds.IsFirstPage
        lbNext.Enabled = Not pds.IsLastPage

        rptrChangeLog.DataSource = pds
        rptrChangeLog.DataBind()
    End Sub

    Private Sub GetProductInfo()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable = dw.GetProductVersion(ProductVersionID)
        If dt.Rows.Count = 1 Then
            lblProductName.Text = String.Format("{0} {1}", dt.Rows(0)("Name"), dt.Rows(0)("Version"))
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        lbAddItem.OnClientClick = String.Format("javascript:AddEntry({0}, {1}); return false;", ProductVersionID, ProductBrandId)
        If Not Page.IsPostBack Then
            GetProductInfo()
            LoadProductBrands()
            LoadChangeHistory()
        End If
    End Sub

    Protected Sub rptrBrands_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptrBrands.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim row As DataRowView = e.Item.DataItem
            Dim lbBrand As LinkButton = e.Item.FindControl("lbBrand")
            lbBrand.Text = row("name")
            lbBrand.OnClientClick = String.Format("javascript:BrandLink_onClick({0}); return false;", row("ProductBrandID"))
            If ProductBrandId = row("ProductBrandID") Then
                BrandName = row("name")
                'lbBrand.Enabled = False
                lbBrand.OnClientClick = "return false;"
                lbBrand.ForeColor = Drawing.Color.Black
                lbBrand.Font.Underline = False
            End If

        End If
    End Sub


    Protected Sub rptrChangeLog_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptrChangeLog.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim row As DataRowView = e.Item.DataItem

            Dim cbxShowOnScm As CheckBox = e.Item.FindControl("cbxShowOnScm")
            Dim cbxShowOnPm As CheckBox = e.Item.FindControl("cbxShowOnPm")
            Dim lblLastUpdDate As Label = e.Item.FindControl("lblLastUpdDate")
            Dim lblLastUpdUser As Label = e.Item.FindControl("lblLastUpduser")
            Dim lblAvNo As Label = e.Item.FindControl("lblAvNo")
            Dim lblGpgDescription As Label = e.Item.FindControl("lblGpgDescription")
            Dim lblColumnChanged As Label = e.Item.FindControl("lblColumnChanged")
            Dim lblChangeType As Label = e.Item.FindControl("lblChangeType")
            Dim lblOldValue As Label = e.Item.FindControl("lblOldValue")
            Dim lblNewValue As Label = e.Item.FindControl("lblNewValue")
            Dim lblComments As Label = e.Item.FindControl("lblComments")

            cbxShowOnScm.Attributes.Add("RecordID", row("ID").ToString())
            cbxShowOnScm.Attributes.Add("OldValue", row("ShowOnScm").ToString())
            cbxShowOnScm.Checked = row("ShowOnScm")


            cbxShowOnPm.Attributes.Add("RecordID", row("ID").ToString())
            cbxShowOnPm.Attributes.Add("OldValue", row("ShowOnPm").ToString())
            cbxShowOnPm.Checked = row("ShowOnPm")

            lblLastUpdDate.Text = Convert.ToDateTime(row("last_upd_date")).ToShortDateString() & "&nbsp;"
            lblLastUpdUser.Text = row("last_upd_user").ToString() & "&nbsp;"
            lblAvNo.Text = row("AvNo").ToString() & "&nbsp;"
            lblGpgDescription.Text = row("GPGDescription").ToString() & "&nbsp;"
            lblColumnChanged.Text = row("ColumnChanged").ToString() & "&nbsp;"
            lblChangeType.Text = row("AvChangeTypeCd").ToString() & "&nbsp;"

            lblOldValue.Text = row("OldValue").ToString() & "&nbsp;"


            lblNewValue.Text = row("NewValue").ToString() & "&nbsp;"


            lblComments.Text = row("Comments").ToString() & "&nbsp;"


            If Not IsPC Then
                cbxShowOnScm.Enabled = False
                cbxShowOnPm.Enabled = False
            End If
        End If
    End Sub

    Protected Sub lbNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbNext.Click
        CurrentPage += 1
        LoadChangeHistory()
    End Sub


    Protected Sub lbPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbPrev.Click
        If CurrentPage > 1 Then
            CurrentPage -= 1
            LoadChangeHistory()
        End If
    End Sub

    Protected Sub lbShowAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbShowAll.Click
        If ShowAll Then
            ShowAll = False
            lbShowAll.Text = "Show All"
        Else
            ShowAll = True
            lbShowAll.Text = "Filter List"
        End If

        CurrentPage = 1
        LoadChangeHistory()
    End Sub

    Protected Sub lbPagination_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbPagination.Click
        If AllowPaging Then
            AllowPaging = False
            lbPagination.Text = "Page"
        Else
            AllowPaging = True
            lbPagination.Text = "Un-Page"
        End If
        CurrentPage = 1
        LoadChangeHistory()
    End Sub

    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveChanges.Click
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim cbxShowOnScm As CheckBox
        Dim cbxShowOnPm As CheckBox
        Dim lRecordId As Long
        Dim oldShowOnScm As Boolean
        Dim oldShowOnPm As Boolean
        For Each item As RepeaterItem In rptrChangeLog.Items
            cbxShowOnScm = item.FindControl("cbxShowOnScm")
            cbxShowOnPm = item.FindControl("cbxShowOnPm")
            lRecordId = Long.Parse(cbxShowOnScm.Attributes("RecordID").ToString())
            oldShowOnScm = Boolean.Parse(cbxShowOnScm.Attributes("OldValue").ToString())
            oldShowOnPm = Boolean.Parse(cbxShowOnPm.Attributes("OldValue").ToString())

            If oldShowOnScm <> cbxShowOnScm.Checked Then
                dw.SetAvHistoryShowOnScmStatus(lRecordId.ToString(), cbxShowOnScm.Checked.ToString(), HPQ.Excalibur.Employee.GetUserName(User.Identity.Name))
            End If

            If oldShowOnPm <> cbxShowOnPm.Checked Then
                dw.SetAvHistoryShowOnPmStatus(lRecordId.ToString(), cbxShowOnPm.Checked.ToString(), HPQ.Excalibur.Employee.GetUserName(User.Identity.Name))
            End If
        Next
    End Sub

End Class