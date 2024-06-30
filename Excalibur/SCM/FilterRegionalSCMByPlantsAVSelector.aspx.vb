Imports System.Data
Partial Class SCM_FilterRegionalSCMByPlant
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    Public _GeoID As String

    Public ReadOnly Property BusinessID() As String
        Get
            Return Request("BusinessID")
            'Return "1"
        End Get
    End Property

    Public ReadOnly Property GeoID2() As String
        Get
            Return Request("GeoID2")
            'Return "1"
        End Get
    End Property

    Public ReadOnly Property SCMPlants() As String
        Get
            Return Request("SCMPlants")
            'Return "1"
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
            'Return "1050"
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
            'Return "763"
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

    Public Property GeoID() As String
        Get
            Return _GeoID
            'Return "Americas"
        End Get
        Set(ByVal value As String)
            _GeoID = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim strGeoID As String

                'strBusinessID = Request.QueryString("BusinessID").ToString()

                Select Case GeoID2
                    Case "Americas"
                        strGeoID = "1"
                    Case "EMEA"
                        strGeoID = "2"
                    Case "APJ"
                        strGeoID = "3"
                End Select
                _GeoID = strGeoID

                FilterByProductType(strGeoID) 'Default is Commercial

                'lblHeader.Text = "BID: " & BID
                lblHeader.Text = "Please Select a plant to filter by!"

                TestLabel.Text = "strGeoID: " & strGeoID & " GeoID2: " & GeoID2
                'TestLabel.Visible = True
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Private Sub FilterByProductType(ByVal strGeoID As String)
        Try
            'Dim strSCMPlants As String = ""

            lbPlants.BackColor = Drawing.Color.White
            TestLabel.Visible = False

            'Dim strGeoID As String

            'strGeoID = "1"
            'strGeoID = Request.QueryString("GeoID").ToString()

            'Load the Plants listbox.
            Dim dtPlants As DataTable
            dtPlants = HPQ.Excalibur.SupplyChain.SelectRCTOPlants_ByGeoID(strGeoID)
            lbPlants.Items.Clear()
            lbPlants.DataSource = dtPlants
            lbPlants.DataBind()

            If SCMPlants <> "" Then
                'strSCMPlants = right(SCMPlants, (SCMPlants.Length - 2))
                'btnDeselect.Visible = True
                Dim sSCMPlants() As String = SCMPlants.Split(":")
                Dim item As ListItem
                For Each item In lbPlants.Items
                    Dim i As Integer = 0
                    For i = 0 To sSCMPlants.Length - 1
                        If sSCMPlants(i) = item.Value Then
                            item.Selected = True
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
            TestLabel.Visible = True
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        'Try
        Dim applicationRoot As String = Session("ApplicationRoot")

        Dim bolFound As Boolean = False
        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim sProdVerID As String = ""
        Dim sProdBrandID As String = ""
        Dim sCategories As String = ""
        Dim strSelRegion As String = "-1"
        Dim item As ListItem
        Dim intPos As Integer = -1
        Dim intPos2 As Integer = -1

        For Each item In lbPlants.Items
            If item.Selected Then
                If sPlantID = "" Then
                    sPlantID = item.Value
                Else
                    sPlantID = sPlantID & "," & item.Value
                End If
            End If
        Next

        For Each item In lbPlants.Items
            If item.Selected Then
                If sPlantName = "" Then
                    sPlantName = item.Text
                Else
                    sPlantName = sPlantName & "," & item.Text
                End If
            End If
        Next

        ProcessFilter(sPlantID, sPlantName)

        'Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    End Sub

    Private Function ProcessFilter(ByVal sPlantID As String, ByVal sPlantName As String) As Boolean
        Try
            Dim URL As String = Nothing
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sPlantID & ":" & sPlantName))
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub lbProductBands_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbPlants.SelectedIndexChanged

    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click

    End Sub
End Class
