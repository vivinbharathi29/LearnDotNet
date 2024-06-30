Imports System.Data

Partial Class Service_SpbMaint
    Inherits System.Web.UI.Page
    ReadOnly Property ProductVersionId() As String
        Get
            Return Request.QueryString("ID")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetExpires(DateTime.Now())
        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        If Not Page.IsPostBack Then txtFamilyPn.Text = hpqData.GetServiceFamilyPn(ProductVersionId)

        If Not Page.IsPostBack And txtFamilyPn.Text.Trim <> String.Empty Then
            lblFamilyPn.Text = txtFamilyPn.Text
            txtFamilyPn.Visible = False
            lblFamilyPn.Visible = True
            btnSaveFamilyPn.Visible = False
            btnEditFamilyPn.Visible = True
            lbExportSpb.Visible = True
            dvSpareKit.Visible = True
        End If

        Dim dtSpbPublishDates As DataTable = hpqData.ListSpbPublishDates(txtFamilyPn.Text)
        If dtSpbPublishDates.Rows.Count > 0 Then
            ddlCompareDt.DataSource = hpqData.ListSpbPublishDates(txtFamilyPn.Text)
            ddlCompareDt.DataTextField = "ExportTime"
            ddlCompareDt.DataValueField = "ExportTime"
            ddlCompareDt.DataBind()
        Else
            ddlCompareDt.Visible = False
            lblCompareDt.Visible = False
        End If


        If Not Page.IsPostBack Then
            Dim drProductNfo As DataRow = hpqData.GetProductVersion(ProductVersionId).Rows(0)
            lblProductVersion.Text = drProductNfo("Name").ToString & " " & drProductNfo("Version").ToString
        End If
    End Sub

    Protected Sub btnSaveFamilyPn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        hpqData.SetServiceFamilyPn(ProductVersionId, txtFamilyPn.Text)
        If txtFamilyPn.Text.Trim() <> String.Empty Then
            lblFamilyPn.Text = txtFamilyPn.Text
            lblFamilyPn.Visible = True
            txtFamilyPn.Visible = False
            btnSaveFamilyPn.Visible = False
            btnEditFamilyPn.Visible = True
            lbExportSpb.Visible = True
            dvSpareKit.Visible = True
        Else
            lbExportSpb.Visible = False
            dvSpareKit.Visible = False
        End If
    End Sub

    Protected Sub btnExportSpb_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExportSpb.Click
        Response.Redirect(String.Format("/iPulsar/ExcelExport/SparesBom.aspx?ServiceFamilyPn={0}&chkPublish={1}&selCompareDt={2}", txtFamilyPn.Text, cbPublishSpb.Checked.ToString(), ddlCompareDt.SelectedValue.ToString()))
    End Sub

    Protected Sub ddlCompareDt_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCompareDt.PreRender
        'ddlCompareDt.Items.Insert(0, New ListItem("-- Select One --", Now.ToString()))
    End Sub

    Protected Sub gvSpareKits_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gvSpareKits.PageIndexChanging
        gvSpareKits.SelectedIndex = -1
        Bind_dvSpareKit()
    End Sub

    Protected Sub gvSpareKits_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvSpareKits.SelectedIndexChanged
        Bind_dvSpareKit()
        mpeDetails.Show()
    End Sub

    Protected Sub dvSpareKit_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdateEventArgs) Handles dvSpareKit.ItemUpdating

        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim HpPartNo As String = gvSpareKits.SelectedRow.Cells(1).Text
        Dim FamilyPn As String = txtFamilyPn.Text

        Dim notVersion As Boolean = False
        Dim i As Integer = 0
        Do Until notVersion Or i = dvSpareKit.Rows.Count - 2
            If dvSpareKit.Rows(i).Cells(0).Text.Contains("Category") Then
                notVersion = True
            Else
                i += 1
            End If
        Loop

        Dim dtls As SpareDetail = New SpareDetail()

        dtls.HpPartNo = HpPartNo
        dtls.ServiceFamilyPn = FamilyPn
        dtls.CategoryName = CType(dvSpareKit.Rows(i).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OsspOrderable = CType(dvSpareKit.Rows(i + 1).Cells(1).Controls(0), CheckBox).Checked
        dtls.OdmPartNo = CType(dvSpareKit.Rows(i + 2).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmPartDesc = CType(dvSpareKit.Rows(i + 3).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmBulkPartNo = CType(dvSpareKit.Rows(i + 4).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmProdMoq = CType(dvSpareKit.Rows(i + 5).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmPostProdMoq = CType(dvSpareKit.Rows(i + 6).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.Supplier = CType(dvSpareKit.Rows(i + 7).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.Comments = CType(dvSpareKit.Rows(i + 8).Cells(1).Controls(0), TextBox).Text.ToString()

        hpqData.UpdateServiceSpareDetail(dtls.ServiceFamilyPn, dtls.HpPartNo, dtls.CategoryName, dtls.OsspOrderable, dtls.OdmPartNo, dtls.OdmPartDesc, dtls.OdmBulkPartNo, dtls.OdmProdMoq, dtls.OdmPostProdMoq, dtls.Comments, dtls.Supplier)

        mpeDetails.Hide()
    End Sub

    Protected Sub dvSpareKit_ItemCanceling(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewCommandEventArgs) Handles dvSpareKit.ItemCommand
        mpeDetails.Hide()
    End Sub

    Protected Sub dvSpareKit_ModeChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewModeEventArgs) Handles dvSpareKit.ModeChanging
        dvSpareKit.ChangeMode(e.NewMode)
        If Not e.NewMode = DetailsViewMode.Insert Then
            Bind_dvSpareKit()
        End If
    End Sub

    Sub Bind_dvSpareKit()
        If gvSpareKits.SelectedIndex >= 0 Then
            Dim HpPartNo As String = gvSpareKits.SelectedRow.Cells(1).Text
            Dim ServiceFamilyPn As String = txtFamilyPn.Text

            Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim dtDetails As DataTable = hpqData.SelectServiceSpareDetails(ServiceFamilyPn, HpPartNo)
            If dtDetails.Rows.Count = 0 Then

                dvSpareKit.Visible = False
                Exit Sub
            End If
            Dim bSpareKit As Boolean = False
            If Not IsDBNull(dtDetails.Rows(0)("SpareKit")) Then
                bSpareKit = dtDetails.Rows(0)("SpareKit")
            End If
            Dim dt As DataTable = New DataTable()
            Dim sbVersionsSupported As StringBuilder = New StringBuilder()
            If bSpareKit Then
                dvSpareKit.HeaderText = "Spare Kit Details"
            Else
                dvSpareKit.HeaderText = "Part Details"
            End If
            dt.Columns.Add("Category")
            dt.Columns.Add("OSSP Orderable", GetType(Boolean))
            dt.Columns.Add("Odm Part No")
            dt.Columns.Add("Odm Part Description")
            dt.Columns.Add("Odm Bulk Part No")
            dt.Columns.Add("Odm Production MOQ")
            dt.Columns.Add("Odm Post Production MOQ")
            dt.Columns.Add("Supplier")
            dt.Columns.Add("Comments")

            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("CategoryName"))
            sbVersionsSupported.AppendFormat("|{0}", Convert.ToBoolean(dtDetails.Rows(0)("OsspOrderable")))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPartNo"))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPartDesc"))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmBulkPartNo"))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmProdMoq"))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPostProdMoq"))
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("Supplier").ToString())
            sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("Comments").ToString())

            dt.Rows.Add(sbVersionsSupported.Remove(0, 1).ToString().Split("|"))

            dvSpareKit.Visible = True
            dvSpareKit.DataSource = dt
            dvSpareKit.DataBind()
        Else
            dvSpareKit.Visible = False
        End If
    End Sub

    Protected Sub btnEditFamilyPn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        btnSaveFamilyPn.Visible = True
        btnEditFamilyPn.Visible = False
        lblFamilyPn.Visible = False
        txtFamilyPn.Visible = True

    End Sub

    Protected Sub btnSpbExport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSpbExport.Click
        Dim saArgs(4) As String
        saArgs(0) = txtFamilyPn.Text
        saArgs(1) = cbPublishSpb.Checked.ToString()
        saArgs(2) = ddlCompareDt.SelectedValue.ToString()
        saArgs(3) = cbNewSpb.Checked.ToString()

        Response.Redirect(String.Format("/iPulsar/ExcelExport/SparesBom.aspx?ServiceFamilyPn={0}&chkPublish={1}&chkNewMatrix={3}&selCompareDt={2}", saArgs))
    End Sub

    Protected Sub dvSpareKit_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles dvSpareKit.DataBound
        If dvSpareKit.DataItemCount = 0 Then
            dvSpareKit.AutoGenerateInsertButton = True
            dvSpareKit.ChangeMode(DetailsViewMode.Insert)
        End If
    End Sub

    Protected Sub dvSpareKit_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewInsertedEventArgs) Handles dvSpareKit.ItemInserted
        dvSpareKit.AutoGenerateInsertButton = False
    End Sub

End Class

#Region " SpareDetail Class "
Public Class SpareDetail

    Private _Categoryname As String
    Public Property CategoryName() As String
        Get
            Return _Categoryname
        End Get
        Set(ByVal value As String)
            _Categoryname = value
        End Set
    End Property

    Private _OdmPostProdMoq As String
    Public Property OdmPostProdMoq() As String
        Get
            Return _OdmPostProdMoq
        End Get
        Set(ByVal value As String)
            _OdmPostProdMoq = value
        End Set
    End Property


    Private _OdmProdMoq As String
    Public Property OdmProdMoq() As String
        Get
            Return _OdmProdMoq
        End Get
        Set(ByVal value As String)
            _OdmProdMoq = value
        End Set
    End Property

    Private _OdmBulkPartNo As String
    Public Property OdmBulkPartNo() As String
        Get
            Return _OdmBulkPartNo
        End Get
        Set(ByVal value As String)
            _OdmBulkPartNo = value
        End Set
    End Property


    Private _OdmPartDesc As String
    Public Property OdmPartDesc() As String
        Get
            Return _OdmPartDesc
        End Get
        Set(ByVal value As String)
            _OdmPartDesc = value
        End Set
    End Property

    Private _hpPartNo As String
    Public Property HpPartNo() As String
        Get
            Return _hpPartNo
        End Get
        Set(ByVal value As String)
            _hpPartNo = value
        End Set
    End Property


    Private _serviceFamilyPn As String
    Public Property ServiceFamilyPn() As String
        Get
            Return _serviceFamilyPn
        End Get
        Set(ByVal value As String)
            _serviceFamilyPn = value
        End Set
    End Property


    Private _OsspOrderable As Boolean
    Public Property OsspOrderable() As Boolean
        Get
            Return _OsspOrderable
        End Get
        Set(ByVal value As Boolean)
            _OsspOrderable = value
        End Set
    End Property


    Private _OdmPartNo As String
    Public Property OdmPartNo() As String
        Get
            Return _OdmPartNo
        End Get
        Set(ByVal value As String)
            _OdmPartNo = value
        End Set
    End Property

    Private _Comments As String
    Public Property Comments() As String
        Get
            Return _Comments
        End Get
        Set(ByVal value As String)
            _Comments = value
        End Set
    End Property

    Private _Supplier As String
    Public Property Supplier() As String
        Get
            Return _Supplier
        End Get
        Set(ByVal value As String)
            _Supplier = value
        End Set
    End Property
End Class
#End Region ' Spare Detail Class
