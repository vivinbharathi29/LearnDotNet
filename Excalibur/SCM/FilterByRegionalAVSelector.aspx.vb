Imports System.Data
Partial Class SCM_FilterByRegionalAVSelector
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

    Public ReadOnly Property ProdBrands() As String
        Get
            Return Request("ProdBrands")
        End Get
    End Property

    Public ReadOnly Property Categories() As String
        Get
            Return Request("Categories")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                rbAll.Checked = True
                FilterByProductType("1", "5") 'Default is Commercial
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Private Sub FilterByProductType(ByVal strProdBrands As String, ByVal strProdStatusID As String)
        Try
            lbProductBands.BackColor = Drawing.Color.White
            rblSelectRegion.BackColor = Drawing.Color.FromArgb(255, 255, 240)
            TestLabel.Visible = False

            Dim dtProdBrands As DataTable
            dtProdBrands = HPQ.Excalibur.SupplyChain.SelectProductBrands(strProdBrands, strProdStatusID)
            lbProductBands.Items.Clear()
            lbProductBands.DataSource = dtProdBrands
            lbProductBands.DataBind()

            If ProdBrands <> "" Then
                'btnDeselect.Visible = True
                Dim sProdBrands() As String = ProdBrands.Split(",")
                Dim item As ListItem
                For Each item In lbProductBands.Items
                    Dim i As Integer = 0
                    For i = 0 To sProdBrands.Length - 1
                        If sProdBrands(i) = item.Value Then
                            item.Selected = True
                        End If
                    Next
                Next
            End If

            Dim dtCategories As DataTable
            dtCategories = dw.SelectAvFeatureCategoriesFilter(strProdBrands)
            lbCategories.Items.Clear()
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
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        'Try
        'Dim applicationRoot As String = Session("ApplicationRoot")

        Dim bolFound As Boolean = False
        Dim sProductBrands As String = ""
        Dim sProdVerID As String = ""
        Dim sProdBrandID As String = ""
        Dim sCategories As String = "-1"
        Dim strSelRegion As String = "-1"
        Dim item As ListItem
        Dim intPos As Integer = -1
        Dim intPos2 As Integer = -1

        For Each item In lbProductBands.Items
            If item.Selected Then
                If sProductBrands = "" Then
                    sProductBrands = item.Value
                Else
                    sProductBrands = sProductBrands & "," & item.Value
                End If
            End If
        Next

        'Separate out all of the product versions and product brands.
        If (sProductBrands <> "") Then
            Dim strSplitUpProdBrands() As String = sProductBrands.Split(",")
            For i As Integer = 0 To UBound(strSplitUpProdBrands)
                intPos = InStr(strSplitUpProdBrands(i), "-")

                If (sProdVerID = "") Then
                    sProdVerID = Left(strSplitUpProdBrands(i), (intPos - 1)) ' & ",Stuff"
                    sProdBrandID = Right(strSplitUpProdBrands(i), (Len(strSplitUpProdBrands(i)) - intPos))
                Else
                    ''Check to see if this Product Version ID has already been added to the sProdVerID variable.
                    bolFound = False
                    Dim strSplitUpProdVerIDs() As String = sProdVerID.Split(",")

                    For j As Integer = 0 To UBound(strSplitUpProdVerIDs)
                        If (strSplitUpProdVerIDs(j) = Left(strSplitUpProdBrands(i), (intPos - 1))) Then
                            bolFound = True
                        End If
                    Next j

                    If (bolFound = False) Then
                        sProdVerID = sProdVerID & "," & Left(strSplitUpProdBrands(i), (intPos - 1))
                    End If

                    sProdBrandID = sProdBrandID & "," & Right(strSplitUpProdBrands(i), (Len(strSplitUpProdBrands(i)) - intPos))
                    'sProdBrandID = UBound(strSplitUpProdVerIDs).ToString()
                End If
            Next i
        End If

        For Each item In lbCategories.Items
            If item.Selected Then
                If sCategories = "" Then
                    sCategories = item.Value
                Else
                    sCategories = sCategories & "," & item.Value
                End If
            End If
        Next

        strSelRegion = rblSelectRegion.SelectedValue.ToString()

        sProductBrands = sProdVerID & ":" & sProdBrandID

        If (sProductBrands = ":") Then
            lbProductBands.BackColor = Drawing.Color.Yellow
            If (strSelRegion = "") Then
                rblSelectRegion.BackColor = Drawing.Color.Yellow
                TestLabel.Text = "You must select at least one Product Brand and a Region!"
            Else
                rblSelectRegion.BackColor = Drawing.Color.FromArgb(255, 255, 240)
                TestLabel.Text = "You must select at least one Product Brand!"
            End If

            TestLabel.Visible = True

            Exit Sub
        Else
            lbProductBands.BackColor = Drawing.Color.White
            rblSelectRegion.BackColor = Drawing.Color.FromArgb(255, 255, 240)
            TestLabel.Visible = False
        End If

        If (strSelRegion = "") Then
            rblSelectRegion.BackColor = Drawing.Color.Yellow
            TestLabel.Visible = True
            TestLabel.Text = "You must select a Region!"

            Exit Sub
        Else
            rblSelectRegion.BackColor = Drawing.Color.FromArgb(255, 255, 240)
            TestLabel.Visible = False
        End If

        ProcessFilter(sProductBrands, sCategories, strSelRegion)

     
    End Sub

    Private Function ProcessFilter(ByVal sProdBrands As String, ByVal sCategory As String, ByVal strSelRegion As String) As Boolean
        Try
            Dim URL As String = Nothing
            'Dim applicationRoot As String = Session("ApplicationRoot")
            'If sCategories = "" Then
            'Response.Write("<script language='javascript'> {window.close();}</script>")
            'Else
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sProdBrands & ":" & sCategory & ":" & strSelRegion))
            'End If
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub rblProductType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblProductType.SelectedIndexChanged
        FilterByProductType(rblProductType.SelectedValue.ToString(), rblProductStatus.SelectedValue.ToString())
    End Sub

    Protected Sub cmdCancel_onclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'Response.Write("<script language='javascript'> {window.close();}</script>")

        Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))
    End Sub

    Protected Sub lbProductBands_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbProductBands.SelectedIndexChanged

    End Sub

    Protected Sub rblProductStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblProductStatus.SelectedIndexChanged
        FilterByProductType(rblProductType.SelectedValue.ToString(), rblProductStatus.SelectedValue.ToString())
    End Sub
End Class
