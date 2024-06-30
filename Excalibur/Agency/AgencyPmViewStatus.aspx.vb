Imports HPQ.Excalibur
Imports System.Data

Partial Public Class Agency_AgencyPmViewStatus
    Inherits System.Web.UI.Page

    Private ReadOnly Property ReportTypeId As Integer
        Get
            Dim returnValue As Integer = 0
            Dim typeId As String = Request.QueryString("ReportTypeId")
            Integer.TryParse(typeId, returnValue)
            Return returnValue
        End Get
    End Property

    Private ReadOnly Property ProductVersionId As Integer
        Get
            Dim returnValue As Integer = 0
            Dim pvId As String = Request.QueryString("ProductVersionId")
            Integer.TryParse(pvId, returnValue)
            Return returnValue
        End Get
    End Property

    Private ReadOnly Property DeliverableVersionId As Integer
        Get
            Dim returnValue As Integer = 0
            Dim pvId As String = Request.QueryString("DeliverableVersionId")
            Integer.TryParse(pvId, returnValue)
            Return returnValue
        End Get
    End Property

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            Dim dt As DataTable = New DataTable
            Select Case ReportTypeId
                Case 1
                    dt = Agency.AgencyStatusSelectDocumentsForPmView(ProductVersionId, DeliverableVersionId, "1")
                Case 2
                    dt = Agency.AgencyStatusSelectDocumentsForPmView(ProductVersionId, DeliverableVersionId, "0")
                Case 3
                    dt = Agency.AgencyStatusSelectCompletedCountries(ProductVersionId, DeliverableVersionId)
                Case 4
                    dt = Agency.AgencyStatusSelectBlockedCountries(ProductVersionId, DeliverableVersionId)
                Case Else
                    pnlWarning.Visible = True
            End Select

            If dt.Rows.Count = 0 Then
                pnlWarning.Visible = True
            Else
                rptrItems.DataSource = dt
                rptrItems.DataBind()
            End If
        End If
    End Sub
End Class
