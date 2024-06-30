Imports System.Data

Partial Class Search_SpareKit
    Inherits BasePage

    Private Const PARTNER_OSSP_HP As Integer = 2
    Private Const REPORT_PROFILE_TYPE_ID As Integer = 9

    Private Const ALL As Integer = 0
    Private Const COMMERCIAL As Integer = 1
    Private Const CONSUMMER As Integer = 2

    Private Const AMERICAS As Integer = 1
    Private Const EMEA As Integer = 2
    Private Const ASIAPACIFIC As Integer = 3

    Private _employeeID As Integer = 0

    Private ReadOnly Property EmployeeID() As Integer
        Get
            If _employeeID = 0 Then
                _employeeID = HPQ.Excalibur.Employee.GetUserID(Session("LoggedInUser"))
            End If
            Return _employeeID
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Response.CacheControl = "No-cache"

            If Not Page.IsPostBack Then
                ' Dim currentUserName As String = Session("LoggedInUser").ToLower()

                lbAddProfile.Attributes.Add("onclick", "getProfileName();")
                lbRenameProfile.Attributes.Add("onclick", "getProfileName();")
                lbShareProfile.Attributes.Add("onclick", "ShareProfile();")

                btnSummaryReport.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
                btnSummaryReport.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
                btnServiceBomReport.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
                btnServiceBomReport.Attributes.Add("onmouseout", "ActionCell_onmouseout();")
                btnReset.Attributes.Add("onmouseover", "ActionCell_onmouseover();")
                btnReset.Attributes.Add("onmouseout", "ActionCell_onmouseout();")

                getSpareKitCategories()
                getProductNames() ' Family Products
                getOSSP()

                FillReportProfiles()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub rdProductType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdProductType.SelectedIndexChanged
        Try
            getProductNames()
            txtServiceFamPartNum.Text = ""
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lstProducts_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles lstProducts.SelectedIndexChanged
        Try
            Dim ProductSelected As String = String.Empty
            Dim lstProductFamilyPn As New SortedList
            lstProductFamilyPn = CType(Session("lstProductFamilyPn"), SortedList)
            For Each elem As ListItem In lstProducts.Items
                If elem.Selected = True Then
                    If ProductSelected = String.Empty Then
                        ProductSelected = lstProductFamilyPn(elem.Value).ToString
                    Else
                        If InStr(lstProductFamilyPn(elem.Value).ToString, ProductSelected.ToString, CompareMethod.Text) = 0 Then
                            ProductSelected = ProductSelected + "," + lstProductFamilyPn(elem.Value).ToString
                        End If
                    End If
                End If
            Next

            If ProductSelected <> String.Empty Then
                txtServiceFamPartNum.Text = ProductSelected
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub chkSKUGeo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSKUGeo.SelectedIndexChanged
        Try
            Dim iNumSelect As Integer = 0

            For Each elem As ListItem In chkSKUGeo.Items
                If elem.Selected = True And elem.Value <> ALL Then
                    iNumSelect = iNumSelect + 1
                    chkSKUGeo.Items(0).Selected = False
                End If
            Next

            If iNumSelect = 0 Then
                chkSKUGeo.Items(0).Selected = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub chkSpsGeo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkSpsGeo.SelectedIndexChanged
        Try
            Dim iNumSelect As Integer = 0

            For Each elem As ListItem In chkSpsGeo.Items
                If elem.Selected = True And elem.Value <> ALL Then
                    iNumSelect = iNumSelect + 1
                    chkSpsGeo.Items(0).Selected = False
                End If
            Next

            If iNumSelect = 0 Then
                chkSpsGeo.Items(0).Selected = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getOSSP()
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

            'ListPartners( string ReportType, string PartnerTypeId )
            Dim dtData As DataTable = dw.ListPartners(1, PARTNER_OSSP_HP)

            Dim dv As DataView = dtData.DefaultView
            dv.Sort = "name asc"

            lstOSSP.DataSource = dv
            lstOSSP.DataTextField = "name"
            lstOSSP.DataValueField = "ID"
            lstOSSP.DataBind()


            dtData = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getProductNames()
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim dtProducts As DataTable = dw.ListProducts("0", "5")

            Dim dv As DataView = dtProducts.DefaultView
            dv.Sort = "ProductVersionName asc"

            If dv.Table.Rows.Count > 0 Then

                Dim lstProductFamilyPn As New SortedList

                For Each Row As DataRow In dv.Table.Rows
                    lstProductFamilyPn.Add(Row("ProductVersionID").ToString, Row("ServiceFamilyPn").ToString)
                Next

                Session("lstProductFamilyPn") = lstProductFamilyPn
                If rdProductType.SelectedValue = ALL Then
                    lstProducts.DataSource = dv
                    lstProducts.DataTextField = "ProductVersionName"
                    lstProducts.DataValueField = "ProductVersionID" 'ServiceFamilyPn
                    lstProducts.DataBind()
                Else 'Consummer or Commercial
                    lstProducts.Items.Clear()
                    For Each Prod As DataRow In dv.Table.Rows
                        If Prod("BusinessId").ToString = rdProductType.SelectedValue Then
                            lstProducts.Items.Add(New ListItem(Prod("ProductVersionName").ToString, Prod("ProductVersionID").ToString))
                        End If
                    Next
                End If



            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSpareKitCategories()
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim dtData As DataTable = dw.ListServiceSpareCategories()

            lstSpareCategory.DataSource = dtData
            lstSpareCategory.DataTextField = "CategoryName"
            lstSpareCategory.DataValueField = "ID"
            lstSpareCategory.DataBind()

            dtData = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try
            txtKMAT.Text = ""
            txtSKUNumber.Text = ""
            txtMaxRows.Text = ""
            txtServiceFamPartNum.Text = ""
            ddlReportFormat.SelectedValue = 0
            lstProducts.SelectedIndex = -1
            lstOSSP.SelectedIndex = -1
            lstSpareCategory.SelectedIndex = -1
            rdProductType.SelectedValue = 0
            chkSKUGeo.SelectedValue = 0
            chkSpsGeo.SelectedValue = 0
            txtSpsNumbers.Text = String.Empty
            'DatepickerSKUStart.Text = String.Empty
            'DatepickerSKUEnd.Text = String.Empty
            DatepickerSPSStart.Text = String.Empty
            DatepickerSPSEnd.Text = String.Empty
           
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "PROFILES"

    Protected Sub lbAddProfile_Click(sender As Object, e As System.EventArgs) Handles lbAddProfile.Click
        Try
            GetSelectedItems()
            CreateProfile()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lbUpdateProfile_Click(sender As Object, e As System.EventArgs) Handles lbUpdateProfile.Click
        Try
            GetSelectedItems()
            UpdateProfile()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lbRenameProfile_Click(sender As Object, e As System.EventArgs) Handles lbRenameProfile.Click
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim profileSelectedValue As String = ddlReportProfiles.SelectedValue
            Dim profileID As String

            If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
                profileID = ddlReportProfiles.SelectedValue.Substring(2)
            Else
                profileID = ddlReportProfiles.SelectedValue.ToString()
            End If

            dw.RenameProfile(profileID, hidProfileName.Value)

            FillReportProfiles()

            ddlReportProfiles.SelectedValue = profileSelectedValue
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lbDeleteProfile_Click(sender As Object, e As System.EventArgs) Handles lbDeleteProfile.Click
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim profileID As String

            If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
                profileID = ddlReportProfiles.SelectedValue.Substring(2)
            Else
                profileID = ddlReportProfiles.SelectedValue.ToString()
                dw.DeleteReportProfile(profileID)
            End If
            Server.Transfer("SpareKit.aspx")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lbRemoveProfile_Click(sender As Object, e As System.EventArgs) Handles lbRemoveProfile.Click
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim profileID As String

            If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
                profileID = ddlReportProfiles.SelectedValue.Substring(2)
                dw.RemoveReportProfile(profileID, EmployeeID)
            End If

            FillReportProfiles()
            Server.Transfer("SpareKit.aspx")

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub CreateProfile()
        Try
            If hidProfileName.Value.Trim() <> String.Empty Then
                Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

                Dim newID As Long = dw.AddReportProfile(hidProfileName.Value, REPORT_PROFILE_TYPE_ID.ToString(), _
                                            EmployeeID, hidCategories.Value, hidProductNames.Value, hidSkuNumber.Value, hidOSSP.Value, hidServiceFamPartNum.Value)

                FillReportProfiles()

                ddlReportProfiles.SelectedValue = newID
                hidProfileId.Value = newID.ToString()

                lbUpdateProfile.Visible = True
                lbDeleteProfile.Visible = True
                lbRenameProfile.Visible = True
                lbShareProfile.Visible = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub UpdateProfile()
        Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim profileID As String

        If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then
            profileID = ddlReportProfiles.SelectedValue.Substring(2)
        Else
            profileID = ddlReportProfiles.SelectedValue.ToString()
        End If

        dw.UpdateProfile(profileID, hidCategories.Value, hidProductNames.Value, hidSkuNumber.Value, hidOSSP.Value, hidServiceFamPartNum.Value)

    End Sub

    Private Sub FillReportProfiles()
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim dtReportProfiles As DataTable = dw.ListReportProfiles(EmployeeID, REPORT_PROFILE_TYPE_ID)
            Dim dtReportProfilesShared As DataTable = dw.ListReportProfilesShared(EmployeeID, REPORT_PROFILE_TYPE_ID)
            Dim dtReportProfilesGroupShared As DataTable = dw.ListReportProfilesGroupShared(EmployeeID, REPORT_PROFILE_TYPE_ID)

            Dim dtReportProfileList As DataTable = New DataTable()
            dtReportProfileList.Columns.Add(New DataColumn("value"))
            dtReportProfileList.Columns.Add(New DataColumn("text"))

            Dim newRow As DataRow = dtReportProfileList.NewRow()
            newRow("value") = ""
            newRow("text") = "Use Options Selected Below"
            dtReportProfileList.Rows.Add(newRow)

            If dtReportProfiles.Rows.Count > 0 Then
                newRow = dtReportProfileList.NewRow()
                newRow("value") = ""
                newRow("text") = "----------------------------------------"
                dtReportProfileList.Rows.Add(newRow)

                For Each row As DataRow In dtReportProfiles.Rows()
                    newRow = dtReportProfileList.NewRow()
                    newRow("value") = row("ID")
                    newRow("text") = row("ProfileName")
                    dtReportProfileList.Rows.Add(newRow)
                Next
            End If

            If dtReportProfilesShared.Rows.Count > 0 Then
                newRow = dtReportProfileList.NewRow()
                newRow("value") = ""
                newRow("text") = "----------- Shared Profiles -----------"
                dtReportProfileList.Rows.Add(newRow)

                For Each row As DataRow In dtReportProfilesShared.Rows()
                    newRow = dtReportProfileList.NewRow()
                    newRow("value") = String.Format("S:{0}", row("ID"))
                    newRow("text") = row("ProfileName")
                    dtReportProfileList.Rows.Add(newRow)
                Next
            End If

            If dtReportProfilesGroupShared.Rows.Count > 0 Then
                newRow = dtReportProfileList.NewRow()
                newRow("value") = ""
                newRow("text") = "----------- Group Shared Profiles -----------"
                dtReportProfileList.Rows.Add(newRow)

                For Each row As DataRow In dtReportProfilesGroupShared.Rows()
                    newRow = dtReportProfileList.NewRow()
                    newRow("value") = String.Format("G:{0}", row("ID"))
                    newRow("text") = row("ProfileName")
                    dtReportProfileList.Rows.Add(newRow)
                Next
            End If

            ddlReportProfiles.DataSource = dtReportProfileList
            ddlReportProfiles.DataValueField = "value"
            ddlReportProfiles.DataTextField = "text"
            ddlReportProfiles.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub ddlReportProfiles_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ddlReportProfiles.SelectedIndexChanged
        Try
            lbUpdateProfile.Visible = False
            lbDeleteProfile.Visible = False
            lbRenameProfile.Visible = False
            lbRemoveProfile.Visible = False
            lbShareProfile.Visible = False
            lblProfileOwnerHdr.Visible = False
            lblProfileOwnerName.Visible = False


            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

            If ddlReportProfiles.SelectedValue <> String.Empty Then
                Dim dt As DataTable
                Dim dr As DataRow
                If ddlReportProfiles.SelectedValue.Substring(0, 1) = "S" Then ' SHARED
                    dt = dw.GetReportProfileShared(ddlReportProfiles.SelectedValue.Substring(2), EmployeeID)
                    dr = dt.Rows(0)

                    If dr("CanEdit") Then
                        lbUpdateProfile.Visible = True
                        lbRenameProfile.Visible = True
                    End If
                    If dr("CanDelete") Then
                        lbDeleteProfile.Visible = True
                    End If
                    lbRemoveProfile.Visible = True


                    lblProfileOwnerHdr.Visible = True
                    lblProfileOwnerName.Visible = True
                    lblProfileOwnerName.Text = dr("PrimaryOwner")

                ElseIf ddlReportProfiles.SelectedValue.Substring(0, 1) = "G" Then ' GROUP
                    dt = dw.GetReportProfileGroup(ddlReportProfiles.SelectedValue.Substring(2))
                    dr = dt.Rows(0)

                    If dr("CanEdit") Then
                        lbUpdateProfile.Visible = True
                        lbRenameProfile.Visible = True
                    End If
                    If dr("CanDelete") Then
                        lbDeleteProfile.Visible = True
                    End If
                    'lbRemoveProfile.Visible = True

                    lblProfileOwnerHdr.Visible = True
                    lblProfileOwnerName.Visible = True
                    lblProfileOwnerName.Text = dr("PrimaryOwner")

                    dt = dw.GetReportProfile(ddlReportProfiles.SelectedValue.Substring(2))
                    dr = dt.Rows(0)
                Else
                    dt = dw.GetReportProfile(ddlReportProfiles.SelectedValue.ToString())
                    dr = dt.Rows(0)

                    lbUpdateProfile.Visible = True
                    lbDeleteProfile.Visible = True
                    lbRenameProfile.Visible = True
                    lbShareProfile.Visible = True
                End If

                hidProfileName.Value = dr("ProfileName").ToString()
                hidProfileId.Value = dr("ID").ToString()
                ' hidCategories.Value, hidProductNames.Value, hidSkuNumber.Value, hidOSSP.Value, hidServiceFamPartNum.Value)
                'string Value15, string value45, string value46, string value47, string value52)

                hidCategories.Value = dr("Value15").ToString()
                hidProductNames.Value = dr("Value45").ToString()

                hidSkuNumber.Value = dr("Value46").ToString()
                hidOSSP.Value = dr("Value47").ToString()
                hidServiceFamPartNum.Value = dr("Value52").ToString()

                LoadProductNamesSavedValues()
                LoadCategoriesSavedValues()
                LoadOSSPSavedValues()

                txtServiceFamPartNum.Text = hidServiceFamPartNum.Value
                txtSKUNumber.Text = dr("Value46").ToString()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadProductNamesSavedValues()
        Try
            lstProducts.SelectedIndex = -1
            If hidProductNames.Value.Trim <> String.Empty Then
                Dim saSelectedItems As String() = hidProductNames.Value.Split("|")

                For Each value As String In saSelectedItems
                    lstProducts.Items.FindByValue(value).Selected = True
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadCategoriesSavedValues()
        Try
            lstSpareCategory.SelectedIndex = -1
            If hidCategories.Value.Trim <> String.Empty Then
                Dim saSelectedItems As String() = hidCategories.Value.Split("|")

                For Each value As String In saSelectedItems
                    lstSpareCategory.Items.FindByValue(value).Selected = True
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadOSSPSavedValues()
        Try
            lstOSSP.SelectedIndex = -1
            If hidOSSP.Value.Trim <> String.Empty Then
                Dim saSelectedItems As String() = hidOSSP.Value.Split("|")

                For Each value As String In saSelectedItems
                    lstOSSP.Items.FindByValue(value).Selected = True
                Next
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetSelectedItems()
        Try
            Dim sbSelectedCategories As Text.StringBuilder = New Text.StringBuilder
            For Each item As ListItem In lstSpareCategory.Items
                If item.Selected Then
                    sbSelectedCategories.Append(String.Format("{0}|", item.Value))
                End If
            Next
            If sbSelectedCategories.Length > 0 Then sbSelectedCategories.Remove(sbSelectedCategories.Length - 1, 1)
            hidCategories.Value = sbSelectedCategories.ToString()

            Dim sbSelectedProductNames As Text.StringBuilder = New Text.StringBuilder
            For Each item As ListItem In lstProducts.Items
                If item.Selected Then
                    sbSelectedProductNames.Append(String.Format("{0}|", item.Value))
                End If
            Next

            Dim sbSelectedOSSSP As Text.StringBuilder = New Text.StringBuilder
            For Each item As ListItem In lstOSSP.Items
                If item.Selected Then
                    sbSelectedOSSSP.Append(String.Format("{0}|", item.Value))
                End If
            Next

            If sbSelectedOSSSP.Length > 0 Then sbSelectedOSSSP.Remove(sbSelectedOSSSP.Length - 1, 1)
            hidOSSP.Value = sbSelectedOSSSP.ToString()

            hidSkuNumber.Value = txtSKUNumber.Text
            hidServiceFamPartNum.Value = txtServiceFamPartNum.Text

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


#End Region

  
  
   
    
   
End Class
