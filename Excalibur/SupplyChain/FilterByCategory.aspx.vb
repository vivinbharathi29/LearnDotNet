Imports System.Data
Partial Class SupChain_FilterByCategory
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property Categories() As String
        Get
            Return Request("Categories")
        End Get
    End Property

    Public ReadOnly Property GADateTo() As String
        Get
            Return Request("GADateTo")
        End Get
    End Property

    Public ReadOnly Property GADateFrom() As String
        Get
            Return Request("GADateFrom")
        End Get
    End Property

    Public ReadOnly Property SADateTo() As String
        Get
            Return Request("SADateTo")
        End Get
    End Property

    Public ReadOnly Property SADateFrom() As String
        Get
            Return Request("SADateFrom")
        End Get
    End Property

    Public ReadOnly Property EMDateTo() As String
        Get
            Return Request("EMDateTo")
        End Get
    End Property

    Public ReadOnly Property EMDateFrom() As String
        Get
            Return Request("EMDateFrom")
        End Get
    End Property

    Public ReadOnly Property NoLocalization() As String
        Get
            Return Request("NoLocalization")
        End Get
    End Property
    Public ReadOnly Property ReleaseIDs() As String
        Get
            Return Request("ReleaseIDs")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim dtCategories As DataTable
                dtCategories = dw.SelectSCMCategoriesFilter()
                lbCategories.DataSource = dtCategories
                lbCategories.DataBind()

                If Categories <> "" Then
                    'btnDeselect.Visible = True
                    Dim sCategories() As String = Categories.Split(",")
                    Dim item As ListItem
                    For Each item In lbCategories.Items
                        Dim i As Integer = 0
                        For i = 0 To sCategories.Length - 1
                            If sCategories(i) = item.Value Then
                                item.Selected = True
                            End If
                        Next
                    Next
                End If

                'If NoLocalization = "1" Then
                '    chkNoLocalizedAvs.Checked = True
                'End If
                Dim dtProductReleases As DataTable
                dtProductReleases = dw.Product_GetProductReleases(PVID)
                Dim strPassedInReleaseIDs As String
                strPassedInReleaseIDs = "," & ReleaseIDs & ","
                Dim cbx As System.Web.UI.WebControls.CheckBox
                For Each row As DataRow In dtProductReleases.Rows
                    cbx = New CheckBox()
                    cbx.Text = row.Item("ReleaseName")
                    cbx.ID = row.Item("ReleaseID")
                    If strPassedInReleaseIDs.IndexOf("," & row.Item("ReleaseID") & ",") > -1 Then
                        cbx.Checked = True
                    End If
                    cbx.EnableViewState = True
                    divReleases.Controls.Add(cbx)

                Next row


                txtGADateFrom.Text = GADateFrom
                txtGADateTo.Text = GADateTo
                txtSADateFrom.Text = SADateFrom
                txtSADateTo.Text = SADateTo
                txtEMDateFrom.Text = EMDateFrom
                txtEMDateTo.Text = EMDateTo

            End If
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim sNoLocalization As String = "0"
            Dim sCategories As String = ""
            Dim sGADateTo As String = ""
            Dim sGADateFrom As String = ""
            Dim sSADateTo As String = ""
            Dim sSADateFrom As String = ""
            Dim sEMDateTo As String = ""
            Dim sEMDateFrom As String = ""
            Dim sReleaseIDs As String = ""
            Dim item As ListItem
            Dim n As Integer = 0
            For Each item In lbCategories.Items
                If item.Selected Then
                    If sCategories = "" Then
                        sCategories = item.Value
                    Else
                        sCategories = sCategories & "," & item.Value
                    End If
                End If
            Next

            'If chkNoLocalizedAvs.Checked = True Then
            'sNoLocalization = "1"
            'End If
            sGADateTo = txtGADateTo.Text
            sGADateFrom = txtGADateFrom.Text

            sSADateTo = txtSADateTo.Text
            sSADateFrom = txtSADateFrom.Text

            sEMDateTo = txtEMDateTo.Text
            sEMDateFrom = txtEMDateFrom.Text
            sReleaseIDs = txtReleaseIDs.Value
            ProcessFilter(sCategories, sNoLocalization, sGADateTo, sGADateFrom, sSADateTo, sSADateFrom, sEMDateTo, sEMDateFrom, sReleaseIDs)
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Private Function ProcessFilter(ByVal sCategories As String, ByVal sNoLocalization As String,
                                   ByVal sGADateTo As String, sGADateFrom As String,
                                   ByVal sSADateTo As String, sSADateFrom As String,
                                   ByVal sEMDateTo As String, sEMDateFrom As String,
                                   ByVal sReleaseIDs As String) As Boolean
        Try
            Dim URL As String = Nothing
            'Dim applicationRoot As String = Session("ApplicationRoot")
            'If sCategories = "" Then
            'Response.Write("<script language='javascript'> {window.close();}</script>")
            'Else

            '  Me.thisBody.Attributes.Add("onload", String.Format("var vReturnValue = new Object(); vReturnValue.Categories = '{0}'; vReturnValue.NoLocalization = '{1}'; vReturnValue.GADateTo = '{2}'; vReturnValue.GADateFrom = '{3}'; vReturnValue.SADateTo = '{4}'; vReturnValue.SADateFrom = '{5}'; vReturnValue.EMDateTo = '{6}'; vReturnValue.EMDateFrom = '{7}'; window.returnValue = vReturnValue; window.close();", sCategories, sNoLocalization, sGADateTo, sGADateFrom, sSADateTo, sSADateFrom, sEMDateTo, sEMDateFrom))
            If Not String.IsNullOrEmpty(Request("pulsarplusDivId")) Then
                Me.thisBody.Attributes.Add("onload", String.Format("var vReturnValue = new Object(); vReturnValue.Categories = '{0}'; vReturnValue.NoLocalization = '{1}'; vReturnValue.GADateTo = '{2}'; vReturnValue.GADateFrom = '{3}'; vReturnValue.SADateTo = '{4}'; vReturnValue.SADateFrom = '{5}'; vReturnValue.EMDateTo = '{6}'; vReturnValue.EMDateFrom = '{7}'; vReturnValue.ReleaseIDs = '{8}'; parent.window.parent.CloseFilterDialog(vReturnValue);", sCategories, sNoLocalization, sGADateTo, sGADateFrom, sSADateTo, sSADateFrom, sEMDateTo, sEMDateFrom, sReleaseIDs))
            Else
                Me.thisBody.Attributes.Add("onload", String.Format("var vReturnValue = new Object(); vReturnValue.Categories = '{0}'; vReturnValue.NoLocalization = '{1}'; vReturnValue.GADateTo = '{2}'; vReturnValue.GADateFrom = '{3}'; vReturnValue.SADateTo = '{4}'; vReturnValue.SADateFrom = '{5}'; vReturnValue.EMDateTo = '{6}'; vReturnValue.EMDateFrom = '{7}'; vReturnValue.ReleaseIDs = '{8}'; window.parent.parent.CloseFilterDialog(vReturnValue);", sCategories, sNoLocalization, sGADateTo, sGADateFrom, sSADateTo, sSADateFrom, sEMDateTo, sEMDateFrom, sReleaseIDs))
            End If



            'End If
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try

        Return True

    End Function

    Protected Sub btnDeselect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeselect.Click
        Dim item As ListItem
        For Each item In lbCategories.Items
            If item.Selected Then
                item.Selected = False
            End If
        Next

        'chkNoLocalizedAvs.Checked = False
        txtGADateTo.Text = String.Empty
        txtGADateFrom.Text = String.Empty
        txtSADateFrom.Text = String.Empty
        txtSADateTo.Text = String.Empty
        txtEMDateFrom.Text = String.Empty
        txtEMDateTo.Text = String.Empty
        Dim dtProductReleases As DataTable
        dtProductReleases = dw.Product_GetProductReleases(PVID)
        Dim cbx As System.Web.UI.WebControls.CheckBox
        For Each row As DataRow In dtProductReleases.Rows
            cbx = New CheckBox()
            cbx.Text = row.Item("ReleaseName")
            cbx.ID = row.Item("ReleaseID")
            cbx.Checked = False
            divReleases.Controls.Add(cbx)

        Next row
        'btnDeselect.Visible = False
    End Sub

    'Protected Sub lbCategories_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbCategories.SelectedIndexChanged
    '    Dim item As ListItem
    '    For Each item In lbCategories.Items
    '        If item.Selected Then
    '            btnDeselect.Visible = True
    '            Exit Sub
    '        End If
    '    Next
    '    btnDeselect.Visible = False
    'End Sub
End Class
