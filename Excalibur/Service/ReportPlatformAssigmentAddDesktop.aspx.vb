Imports System.Data

Partial Class Service_ReportPlatformAssigmentAddDesktop
    Inherits System.Web.UI.Page

    Private Const DIVISION_DESKTOP As String = "2"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                getProductFamily()
                getProductLines()
                getGPLM()
                getODM()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Try
            Dim OdmSelected As Boolean = False
            Dim GPLMSelected As Boolean = False
            Dim ProductFamilySelected As Boolean = False

            lblErrorMessage.Text = String.Empty

            'validate data imput
            For Each elem As ListItem In lstProductFamily.Items
                If elem.Selected = True Then
                    ProductFamilySelected = True
                End If
            Next

            If ProductFamilySelected = False Then
                lblErrorMessage.Text = "Product Family: You have to select one Product Family to the desktop platform. "
                Exit Sub
            End If

            'validate data imput
            For Each elem As ListItem In lstODM.Items
                If elem.Selected = True Then
                   OdmSelected = True
                End If
            Next

            If OdmSelected = False Then
                lblErrorMessage.Text = "ODM: You have to select one ODM to the desktop platform. "
                Exit Sub
            End If

            For Each elem As ListItem In lstGPLM.Items
                If elem.Selected = True Then
                    GPLMSelected = True
                End If
            Next

            If OdmSelected = False Then
                lblErrorMessage.Text = "GPLM: You have to select one GPLM to the desktop platform. "
                Exit Sub
            End If

            If txtServiceFamilyPn.Text = String.Empty Then
                lblErrorMessage.Text = "ServiceFamilyPn: You have to write a ServiceFamilyPn. "
                Exit Sub
            End If

            If txtPlatformName.Text = String.Empty Then
                lblErrorMessage.Text = "Platform Name: You have to write a Platform Name. "
                Exit Sub
            End If

            If txtPlatformDescription.Text = String.Empty Then
                lblErrorMessage.Text = "Platform Description: You have to write a Platform Description. "
                Exit Sub
            End If

            If txtFCS.Text <> String.Empty Then
                If Not IsDate(txtFCS.Text) Then
                    lblErrorMessage.Text = "FCS: You have to write a date FCS. "
                    Exit Sub
                End If
            End If

            If txtEndOfService.Text <> String.Empty Then
                If Not IsDate(txtEndOfService.Text) Then
                    lblErrorMessage.Text = "End of Service: You have to write an End of Service date. "
                    Exit Sub
                End If
            End If

            AddDesktopPlatform()
        Catch ex As Exception
            '     Violation of PRIMARY KEY constraint 'PK_ServiceFamilyDetails'. Cannot insert duplicate key in object 'dbo.ServiceFamilyDetails'. The 
            If InStr(ex.Message, "Violation of PRIMARY KEY constraint", CompareMethod.Text) = 1 Then
                lblErrorMessage.Text = "ServiceFamilyPn: There is another Product with the same value, you have to write another number. "
                Exit Sub
            Else
                Throw ex
            End If
        End Try
    End Sub

    Protected Sub btnAddAnother_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAnother.Click
        Try
            lblErrorMessage.Text = String.Empty

            lstGPLM.SelectedIndex = -1
            lstODM.SelectedIndex = -1
            lstProductFamily.SelectedIndex = -1
            lstProductLine.SelectedIndex = -1

            txtServiceFamilyPn.Text = String.Empty
            txtPlatformName.Text = String.Empty
            txtPlatformDescription.Text = String.Empty
            txtFCS.Text = String.Empty
            txtEndOfService.Text = String.Empty


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClear.Click
        Try
            For Each elem As ListItem In lstGPLM.Items
                elem.Selected = False
            Next

            For Each elem As ListItem In lstODM.Items
                elem.Selected = False
            Next

            For Each elem As ListItem In lstProductFamily.Items
                elem.Selected = False
            Next

            For Each elem As ListItem In lstProductLine.Items
                elem.Selected = False
            Next

            txtEndOfService.Text = String.Empty
            txtFCS.Text = String.Empty
            txtPlatformDescription.Text = String.Empty
            txtPlatformName.Text = String.Empty
            txtProductFamilyName.Text = String.Empty
            txtServiceFamilyPn.Text = String.Empty

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AddDesktopPlatform()
        Try
            Dim iRes As Integer = HPQ.Excalibur.Service.InsertDesktopPlatform(lstProductFamily.SelectedValue, lstProductLine.SelectedValue, txtServiceFamilyPn.Text, txtPlatformName.Text, txtPlatformDescription.Text, lstODM.SelectedValue, lstGPLM.SelectedValue, txtFCS.Text, txtEndOfService.Text)
   
            If iRes = -1 Then
                Dim script As String = "window.close();" 'window.opener.location.reload(true);
                System.Web.UI.ScriptManager.RegisterClientScriptBlock(Page, Me.[GetType](), "CloseWindow", script, True)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lnkBtAddProductFamily_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtAddProductFamily.Click
        Try
            trNewProductFamily.Visible = True
            lnkBtAddProductFamily.Enabled = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnAddFamily_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddFamily.Click
        Try
            'Insert Family Name in database
            If txtProductFamilyName.Text <> "" Then
                ' spAddNewProductFamily

                Dim iRes As Integer = HPQ.Excalibur.Service.InsertProductFamily(txtProductFamilyName.Text)

            Else
                lblErrorMessage.Text = "you have to write a Family Name"
                Exit Sub
            End If

            'Refresh Families
            trNewProductFamily.Visible = False
            getProductFamily()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnCancelAddFamily_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancelAddFamily.Click
        Try
            trNewProductFamily.Visible = False
            txtProductFamilyName.Text = ""
            lblErrorMessage.Text = ""
            lnkBtAddProductFamily.Enabled = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getProductFamily()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.GetProductFamilies()

            If dtData.Rows.Count > 0 Then
                lstProductFamily.DataSource = dtData
                lstProductFamily.DataTextField = "ProductFamilyName"
                lstProductFamily.DataValueField = "ProductFamilyID"
                lstProductFamily.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getProductLines()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.GetProductLines()

            If dtData.Rows.Count > 0 Then
                lstProductLine.DataSource = dtData
                lstProductLine.DataTextField = "Description"
                lstProductLine.DataValueField = "ID"
                lstProductLine.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getGPLM()
        Try
            Dim dtData As New DataTable
            Dim objData As New HPQ.Excalibur.Data

            dtData = objData.ListGplms()

            If dtData.Rows.Count > 0 Then
                lstGPLM.DataSource = dtData
                lstGPLM.DataTextField = "Name"
                lstGPLM.DataValueField = "ID"
                lstGPLM.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getODM()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.ListOsspPartners()

            If dtData.Rows.Count > 0 Then
                lstODM.DataSource = dtData
                lstODM.DataTextField = "Name"
                lstODM.DataValueField = "ID"
                lstODM.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub







End Class
