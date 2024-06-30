<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="HPQ.Excalibur" %>

<script runat="server">
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
    Private Property IsPC() As Boolean
        Get
            If ViewState("IsPC") Is Nothing Then
                Return False
            Else
                Return Convert.ToBoolean(ViewState("IsPC").ToString())
            End If
        End Get
        Set(ByVal value As Boolean)
            ViewState("IsPC") = value
        End Set
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
        Dim dt As DataTable = dw.ListBrands4Product(ProductVersionID, 2)

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
        If dt.Rows.Count > 0 Then
            lblProductName.Text = String.Format("{0} {1}", dt.Rows(0)("Name"), dt.Rows(0)("Version"))
        End If
    End Sub

    Private Sub GetUserInfo()
        Dim secObj As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(User.Identity.Name)
        If (secObj.IsProgramCoordinator Or secObj.IsSysAdmin) Then
            IsPC = True
        Else
            IsPC = False
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        lbAddItem.OnClientClick = String.Format("javascript:AddEntry({0}, {1}); return false;", ProductVersionID, ProductBrandId)
        lbExportToExcel.OnClientClick = String.Format("javascript:ExportToExcel({0}); return false;", ProductBrandId)
        If Not Page.IsPostBack Then
            GetUserInfo()
            GetProductInfo()
            LoadProductBrands()
            LoadChangeHistory()
        End If
        If Not IsPC Then
            btnSaveChanges.Enabled = False
            lbAddItem.Enabled = False
        End If
    End Sub

    Protected Sub rptrBrands_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs)
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim row As DataRowView = e.Item.DataItem
            Dim lbBrand As LinkButton = e.Item.FindControl("lbBrand")
            lbBrand.Text = row("name")
            lbBrand.OnClientClick = String.Format("javascript:({0}); return false;", row("ProductBrandID"))
            If ProductBrandId = row("ProductBrandID") Then
                BrandName = row("name")
                'lbBrand.Enabled = False
                lbBrand.OnClientClick = "return false;"
                lbBrand.ForeColor = Drawing.Color.Black
                lbBrand.Font.Underline = False
            End If

        End If
    End Sub


    Protected Sub rptrChangeLog_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs)
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim row As DataRowView = e.Item.DataItem

            Dim cbxShowOnScm As CheckBox = e.Item.FindControl("cbxShowOnScm")
            Dim cbxShowOnPm As CheckBox = e.Item.FindControl("cbxShowOnPm")
            Dim lblLastUpdDate As Label = e.Item.FindControl("lblLastUpdDate")
            Dim lblLastUpdUser As Label = e.Item.FindControl("lblLastUpduser")
            Dim lblFeatureID As Label = e.Item.FindControl("lblFeatureID")
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
            lblFeatureID.Text = row("FeatureID").ToString() & "&nbsp;"
            lblAvNo.Text = row("AvNo").ToString() & "&nbsp;"
            lblGpgDescription.Text = row("GPGDescription").ToString() & "&nbsp;"
            lblColumnChanged.Text = row("ColumnChanged").ToString() & "&nbsp;"
            lblChangeType.Text = row("AvChangeTypeCd").ToString() & "&nbsp;"

            'If row("OldValue").ToString.Length > 35 Then
            'lblOldValue.Text = String.Format("{0} ...", row("OldValue").ToString.Substring(0, 30))
            'Else
            lblOldValue.Text = row("OldValue").ToString() & "&nbsp;"
            'End If

            'If row("NewValue").ToString.Length > 35 Then
            'lblNewValue.Text = String.Format("{0} ...", row("newvalue").ToString.Substring(0, 30))
            'Else
            lblNewValue.Text = row("NewValue").ToString() & "&nbsp;"
            'End If

            'If row("Comments").ToString.Length > 35 Then
            'lblComments.Text = String.Format("{0} ...", row("Comments").ToString.Substring(0, 30))
            'Else
            lblComments.Text = row("Comments").ToString() & "&nbsp;"
            'End If

            If Not IsPC Then
                cbxShowOnScm.Enabled = False
                cbxShowOnPm.Enabled = False
            End If
        End If
    End Sub

    Protected Sub lbNext_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CurrentPage += 1
        LoadChangeHistory()
    End Sub


    Protected Sub lbPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If CurrentPage > 1 Then
            CurrentPage -= 1
            LoadChangeHistory()
        End If
    End Sub

    Protected Sub lbShowAll_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub lbPagination_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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

    Protected Sub btnSaveChanges_Click(ByVal sender As Object, ByVal e As System.EventArgs)
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
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>SCM Change Log</title>
    <link href="../style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="../style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="style.css" />

    <script type="text/javascript">
function BrandLink_onClick(ProductBrandID)
{
	window.location.replace("changelog.aspx?ID=<%=Request("ID")%>&Class=<%=Request("Class")%>&BID=" + ProductBrandID);
}

function Row_OnMouseOver()
{
	var node = window.event.srcElement;
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}
	
	node.style.color = "red";
	node.style.cursor = "hand";
}

function Row_OnMouseOut() {
	var node = window.event.srcElement;
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}

   	node.style.color = "black";
}

function Row_OnClick()
{
	var node = window.event.srcElement;
	
	if (node.type == "checkbox")
	    return;
	
	while (node.nodeName.toUpperCase() != "TR")
	{
		node = node.parentElement;
	}
	
	var strID;
	strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=edit&PVID=" + node.getAttribute("pvid") + "&CLID=" + node.getAttribute("clid"), "", "dialogWidth:500px;dialogHeight:400px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
	if (strID > 0)
	    document.location.reload();
}

function AddEntry( ProductVersionID, ProductBrandID )
{
	var strID;
	strID = window.parent.showModalDialog("editChngLogFrame.asp?Mode=add&PVID=" + ProductVersionID + "&PBID=" + ProductBrandID, "", "dialogWidth:500px;dialogHeight:275px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");
	if (strID > 0)
	    document.location.reload();
}

function ExportToExcel(ProductBrandID) {
    var url = "/IPulsar/SCM/SCMShowChangeLog.aspx?ProductBrandID=" + ProductBrandID + "&ShowAll=0";
    window.open(url, "_blank");
}
    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <p>
                <span style="font-size: medium; font-weight: bold">
                    <asp:Label ID="lblProductName" runat="server" Text="Label"></asp:Label>
                    SCM Change Log</span></p>
            <table class="DisplayBar" width="100%" cellspacing="0" cellpadding="2">
                <tr>
                    <td valign="top">
                        <table>
                            <tr>
                                <td valign="top" style="height: 14px; font-weight: bold; font-size: small; color: navy;
                                    font-family: Verdana;">
                                    Display:&nbsp;&nbsp;&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                    <td style="width: 100%">
                        <table>
                            <tr>
                                <td>
                                    <b>Brand:</b></td>
                                <td style="width: 100%">
                                    <asp:Repeater ID="rptrBrands" runat="server" OnItemDataBound="rptrBrands_ItemDataBound">
                                        <ItemTemplate>
                                            <a id="aBrand" runat="server"></a>
                                            <asp:LinkButton ID="lbBrand" runat="server">LinkButton</asp:LinkButton></ItemTemplate>
                                        <SeparatorTemplate>
                                            &nbsp;|&nbsp;</SeparatorTemplate>
                                    </asp:Repeater>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <br />
			<table style="width:100%">
				<tr>
					<td width="95%"><asp:Button ID="btnSaveChanges" runat="server" Text="Save Changes" onclick="btnSaveChanges_Click" /></td>
					<td width="5%"><asp:LinkButton ID="lbExportToExcel" runat="server" Text="Export_To_Excel"/></td>
				</tr>
            </table>
            <br />
            <asp:LinkButton ID="lbAddItem" runat="server" Text="Add Item" />
            |
            <asp:LinkButton ID="lbShowAll" runat="server" Text="Show All" 
                onclick="lbShowAll_Click" />
            <br />
            <br />
            <div style="text-align: right">
            <asp:LinkButton ID="lbPagination" runat="server" Text="Un-Page" 
                    onclick="lbPagination_Click" />&nbsp;|&nbsp;
                <asp:Label ID="lblCurrentPage" runat="server" Text="Label" />
                <asp:LinkButton ID="lbPrev" runat="server" OnClick="lbPrev_Click" Text="<<" />
                <asp:LinkButton ID="lbNext" runat="server" OnClick="lbNext_Click" Text=">>" /></div>
            <span style="font-size: x-small; font-weight: bold">
                <asp:Label ID="lblBrand" runat="server" /></span> - <span style="color: Red">(Click
                    on the change row to view the details.)</span>
            <br />
            <br />
            <asp:Repeater ID="rptrChangeLog" runat="server" OnItemDataBound="rptrChangeLog_ItemDataBound">
                <HeaderTemplate>
                    <table id="TableSchedule" cellspacing="1" cellpadding="1" width="100%" border="1"
                        bordercolor="tan" bgcolor="ivory">
                        <col align="center" />
                        <col align="center" />
                        <col />
                        <col />
                        <col align="center">
                        <col />
                        <col />
                        <col />
                        <col align="center" />
                        <col />
                        <col />
                        <col />
                        <tr>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Show<br />
                                On SCM</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Show<br />
                                On PM</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Change Date</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Changed By</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Feature ID.</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                AV No.</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                GPG Desc.</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Field</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Change<br />
                                Type</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Change From</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Change To</th>
                            <th style="white-space: nowrap; background-color: cornsilk; text-align: center; vertical-align: middle">
                                Comment / Reason</th>
                        </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr style="background-color: cornsilk" pvid="<%= ProductVersionID %>" clid="<%# Eval("ID")%>"
                        onmouseover="return Row_OnMouseOver()" onmouseout="return Row_OnMouseOut()" onclick="return Row_OnClick()">
                        <td class="cell" style="white-space: nowrap">
                            <asp:CheckBox ID="cbxShowOnScm" runat="server" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:CheckBox ID="cbxShowOnPm" runat="server" />
                        </td>
                        <td class="cell" style="white-space: nowrap; text-align: right">
                            <asp:Label ID="lblLastUpdDate" runat="server" Text="LastUpdDate" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblLastUpdUser" runat="server" Text="LastUpdUser" />
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblFeatureID" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblAvNo" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblGpgDescription" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblColumnChanged" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblChangeType" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblOldValue" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblNewValue" runat="server" Text="Label"></asp:Label>
                        </td>
                        <td class="cell" style="white-space: nowrap">
                            <asp:Label ID="lblComments" runat="server" Text="Label"></asp:Label>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table></FooterTemplate>
            </asp:Repeater>
        </div>
    </form>
</body>
</html>
