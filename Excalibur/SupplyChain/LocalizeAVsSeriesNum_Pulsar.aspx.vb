Imports System.Data

Partial Class LocalizeAVsSeriesNum_Pulsar
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property UserName() As String
        Get
            Return Request("UserName")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            hdnProductVersionId.Value = PVID
            If Not IsPostBack Then
                Dim dt As New DataTable
                dt = dw.ListFeatureCategoy_Localized()
                ddlAvType.DataSource = dt
                ddlAvType.DataBind()

                GetProductReleases(PVID)
                'GetProductDefaultDates(PVID)
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Select Case ddlAvType.SelectedItem.Value
                Case 0
                    'lblHeader.ForeColor = Drawing.Color.Red
                Case Else
                    '1=Images, 2=HWKits, 3=Keyboards, 4=OS Restore Media
                    Dim strPath As String
                    Dim intShowAllLocs As Integer = 0
                    If cbxShowAllLocs.Checked Then
                        intShowAllLocs = 1
                    End If
                    Dim strReleases As String
                    strReleases = hdnSelectedReleases.Value
                    Dim strRTPDate As String
                    Dim strEMDate As String
                    strRTPDate = hdnRTPDate.Value
                    strEMDate = hdnEMDate.Value
                    strPath = Me.ResolveUrl("~/SupplyChain/LocalizeAVs_Pulsar.asp")
                    If Not String.IsNullOrEmpty(Request("pulsarplusDivId")) Then
                        strPath = """" & strPath & "?Mode=add&pulsarplusDivId=" & Request("pulsarplusDivId") & "&PVID=" & PVID & "&BID=" & BID & "&AvType=" & ddlAvType.SelectedItem.Value & "&CategoryID=" & ddlAvType.SelectedItem.Value & "&UserName=" & UserName & "&ShowAllLocs=" & intShowAllLocs.ToString() & "&Releases=" & strReleases & "&RTPDate=" & strRTPDate & "&EMDate=" & strEMDate & """, """", ""dialogWidth:1200px;dialogHeight:1000px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No"""
                        thisBody.Attributes.Add("onload", "parent.window.parent.OpenLocalizeAVs(" & strPath & ");")
                    Else
                        strPath = """" & strPath & "?Mode=add&PVID=" & PVID & "&BID=" & BID & "&AvType=" & ddlAvType.SelectedItem.Value & "&CategoryID=" & ddlAvType.SelectedItem.Value & "&UserName=" & UserName & "&ShowAllLocs=" & intShowAllLocs.ToString() & "&Releases=" & strReleases & "&RTPDate=" & strRTPDate & "&EMDate=" & strEMDate & """, """", ""dialogWidth:1200px;dialogHeight:1000px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No"""
                        thisBody.Attributes.Add("onload", "window.parent.OpenLocalizeAVs(" & strPath & ");")
                    End If


            End Select
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub
    Private Sub GetProductReleases(ProductID As Integer)
        Dim dt As New DataTable()
        Dim dwData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        dt = dwData.Product_GetProductReleases(ProductID.ToString())
        Dim strProductRelease As String
        strProductRelease = ""
        For row As Integer = 0 To dt.Rows.Count - 1
            If strProductRelease = "" Then
                strProductRelease = Convert.ToString(dt.Rows(row)("ReleaseName")) + "," + Convert.ToString(dt.Rows(row)("ReleaseID")) + "," + Convert.ToString(dt.Rows(row)("RTPDate")) + "," + Convert.ToString(dt.Rows(row)("EMDate"))
            Else
                strProductRelease = strProductRelease + ";" + Convert.ToString(dt.Rows(row)("ReleaseName")) + "," + Convert.ToString(dt.Rows(row)("ReleaseID")) + "," + Convert.ToString(dt.Rows(row)("RTPDate")) + "," + Convert.ToString(dt.Rows(row)("EMDate"))
            End If
        Next
        hdnProductReleases.Value = strProductRelease
    End Sub

    Private Sub GetProductDefaultDates(ProductID As Integer)
        Dim dt As New DataTable()
        Dim dwData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        dt = dwData.Product_GetDefaultDates(ProductID.ToString())
        For row As Integer = 0 To dt.Rows.Count - 1
            hdnRTPDate.Value = dt.Rows(row)("RTPDate").ToString()
            hdnEMDate.Value = dt.Rows(row)("EMDate").ToString()
        Next
    End Sub

End Class
