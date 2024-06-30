Imports System.Data

Partial Class Service_DesktopFamilyDetails
    Inherits System.Web.UI.Page

    Public ServiceFamilyPartNumner, PVID, PlatformName As String

    ' Flag for Edit Permission of Product Version's Service Family Part Number
    Private blnCanEditSFPN As Boolean = False

    Property CanEditSFPN() As Boolean
        Get
            Return blnCanEditSFPN
        End Get
        Set(ByVal value As Boolean)
            blnCanEditSFPN = value
        End Set
    End Property

    Private blnIsSpdmUser As Boolean = False

    Property IsSpdmUser() As Boolean
        Get
            Return blnIsSpdmUser
        End Get
        Set(ByVal value As Boolean)
            blnIsSpdmUser = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ServiceFamilyPartNumner = Request.QueryString("servicefamilypn")
            PVID = Request.QueryString("ProductVersionID")
            PlatformName = Request.QueryString("Platform")
            If Not Page.IsPostBack Then
                ddlPartner.Enabled = False
                'lblTitle.Text = lblTitle.Text + ": " + FamilyName + " - " + ServiceFamilyPartNumner
                lblFamilyPn.Text = ServiceFamilyPartNumner

                ' Initialize User Roles/Permissions
                Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(HttpContext.Current.User.Identity.Name.ToString())
                ' Set Edit Permissions
                IsSpdmUser = (objSec.IsSysAdmin Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.GPLM) Or objSec.UserInRole(HPQ.Excalibur.Security.ProgramRoles.ServiceBomAnalyst))
                CanEditSFPN = IsSpdmUser

                objSec = Nothing

                If Not CanEditSFPN Then
                    btnSaveSFP.Enabled = False
                Else
                    btnSaveSFP.Enabled = True
                End If

                getGPLM()

                'GetServiceManagers()
                getODM()
                'getSPDM()

                GetDesktopServicefamilyDetails()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetDesktopServicefamilyDetails()
        Try
            Dim dtData As New DataTable
            Dim ObjData As New HPQ.Excalibur.Data
            dtData = ObjData.SelectSpbDetails(ServiceFamilyPartNumner)

            ddlGPLM.SelectedValue = "-1"
       
            If dtData.Rows.Count > 0 Then
                'If Not IsDBNull(dtData.Rows(0)("SvcMgrContactID")) Then
                '    ddlServiceManager.Text = dtData.Rows(0)("SvcMgrContactID").ToString
                'Else
                '    ddlServiceManager.SelectedValue = "-1"
                'End If
                If Not IsDBNull(dtData.Rows(0)("GplmContactID")) Then
                    ddlGPLM.SelectedValue = dtData.Rows(0)("GplmContactID").ToString
                Else
                    ddlGPLM.SelectedValue = "-1"
                End If
                'If Not IsDBNull(dtData.Rows(0)("SpdmContactID")) Then
                '    ddlBomAnalyst.SelectedValue = dtData.Rows(0)("SpdmContactID").ToString
                'Else
                '    ddlBomAnalyst.SelectedValue = "-1"
                'End If

                If Not IsDBNull(dtData.Rows(0)("AutoPublishRsl")) Then
                    chkSPBAutoPub.Checked = CBool(dtData.Rows(0)("Active").ToString)
                Else
                    chkSPBAutoPub.Checked = False
                End If

                If Not IsDBNull(dtData.Rows(0)("AutoPublishRsl")) Then
                    chkRSLAutoPub.Checked = CBool(dtData.Rows(0)("AutoPublishRsl").ToString)
                Else
                    chkRSLAutoPub.Checked = False
                End If

                lblSeriesName.Text = dtData.Rows(0)("FamilyName").ToString
                rbBusiness.SelectedValue = dtData.Rows(0)("BusinessUnit").ToString

                lblProjectCode.Text = dtData.Rows(0)("ProjectCd").ToString

                If Not IsDBNull(dtData.Rows(0)("PartnerID")) Then
                    ddlPartner.SelectedValue = dtData.Rows(0)("PartnerID").ToString
                Else
                    ddlPartner.Items.Add(New ListItem("Select ...", ""))
                    ddlPartner.SelectedValue = "-1"
                End If

                If Not IsDBNull(dtData.Rows(0)("BusinessUnit")) Then
                    rbBusiness.SelectedItem.Value = dtData.Rows(0)("BusinessUnit")
                End If

                'Dates
                txtEndOfService.Text = dtData.Rows(0)("ServiceLifeDate").ToString

                Dim dtDesktopDates As New DataTable

                dtDesktopDates = HPQ.Excalibur.Service.GetServiceDates(PVID)
                If dtDesktopDates.Rows.Count > 0 Then
                    txtFCS.Text = dtDesktopDates.Rows(0)("actual_end_dt").ToString
                Else
                    txtFCS.Text = String.Empty
                End If

                Me.txtPlatformName.Text = PlatformName

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnSaveSFP_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveSFP.Click
        Try
            Dim ObjData As New HPQ.Excalibur.Data
            Dim GPLM As String = String.Empty


            If txtFCS.Text <> String.Empty Then
                If Not IsDate(txtFCS.Text) Then
                    lblErrorMessage.Text = "FCS Date: The data is not a date type.(mm/dd/yyy)"
                    Exit Sub
                End If
            End If

            If txtEndOfService.Text <> String.Empty Then
                If Not IsDate(txtEndOfService.Text) Then
                    lblErrorMessage.Text = "End Of Service Date: The data is not a date type.(mm/dd/yyy)"
                    Exit Sub
                End If
            End If

            If ddlGPLM.SelectedValue <> "-1" Then
                GPLM = ddlGPLM.SelectedValue
            End If


            Dim iRes As Integer = ObjData.UpdateServiceFamilyDetails(ServiceFamilyPartNumner, String.Empty, GPLM, IIf(chkSPBAutoPub.Checked, 1, 0), String.Empty, String.Empty, String.Empty, IIf(chkRSLAutoPub.Checked, 1, 0), rbBusiness.SelectedValue)

            If iRes = -1 Then
                iRes = HPQ.Excalibur.Service.UpdatePlatformName(PVID, txtPlatformName.Text)
                If iRes = -1 Then
                    iRes = HPQ.Excalibur.Service.UpdateDesktopDates(PVID, ServiceFamilyPartNumner, txtFCS.Text, txtEndOfService.Text)
                End If
            End If

            If iRes = -1 Then
                Dim script As String = "window.opener.location.reload(true);window.close();"
                System.Web.UI.ScriptManager.RegisterClientScriptBlock(Page, Me.[GetType](), "CloseWindow", script, True)


            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntClear.Click
        Try
            txtEndOfService.Text = String.Empty
            txtFCS.Text = String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getODM()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.ListOsspPartners()

            If dtData.Rows.Count > 0 Then
                ddlPartner.DataSource = dtData
                ddlPartner.DataTextField = "Name"
                ddlPartner.DataValueField = "ID"
                ddlPartner.DataBind()
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
                ddlGPLM.DataSource = dtData
                ddlGPLM.DataTextField = "Name"
                ddlGPLM.DataValueField = "ID"
                ddlGPLM.DataBind()
            End If

            ddlGPLM.Items.Add(New ListItem("Select ...", "-1"))
            ddlGPLM.SelectedValue = "-1"

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Private Sub getSPDM()
    '    Try
    '        Dim dtData As New DataTable
    '        Dim objData As New HPQ.Excalibur.Data

    '        dtData = objData.ListSpdms()

    '        If dtData.Rows.Count > 0 Then
    '            ddlBomAnalyst.DataSource = dtData
    '            ddlBomAnalyst.DataTextField = "Name"
    '            ddlBomAnalyst.DataValueField = "ID"
    '            ddlBomAnalyst.DataBind()
    '        End If

    '        ddlBomAnalyst.Items.Add(New ListItem("Select ...", "-1"))
    '        ddlBomAnalyst.SelectedItem.Value = "-1"

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Sub GetServiceManagers()
    '    Try
    '        Dim dtData As New DataTable
    '        Dim objData As New HPQ.Excalibur.Data

    '        dtData = objData.ListSvcManagers(String.Empty)

    '        If dtData.Rows.Count > 0 Then
    '            ddlServiceManager.DataSource = dtData
    '            ddlServiceManager.DataTextField = "Name"
    '            ddlServiceManager.DataValueField = "ID"
    '            ddlServiceManager.DataBind()
    '        End If

    '        ddlServiceManager.Items.Add(New ListItem("Select ...", "-1"))
    '        ddlServiceManager.SelectedItem.Value = "-1"

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

End Class
