Imports System.Data
Partial Class SCM_FilterRegionalSCMByMktCamp
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

    Public ReadOnly Property SCMMktCamp() As String
        Get
            Return Request("SCMMktCamp")
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
                lblHeader.Text = "Please Select a Marketing Campaign to filter by!"

                TestLabel.Text = "strGeoID: " & strGeoID & " GeoID: " & GeoID & " GeoID2: " & GeoID2
                'TestLabel.Text = "SCMMktCamp: " & SCMMktCamp
                'TestLabel.Visible = True
            End If
        Catch ex As Exception
            TestLabel.Text = ex.ToString
        End Try
    End Sub

    Private Sub FilterByProductType(ByVal strGeoID As String)
        Try
            'Dim strSCMMktCamps As String = ""

            lbMktCamp.BackColor = Drawing.Color.White
            TestLabel.Visible = False

            'Dim strGeoID As String

            'strGeoID = "1"
            'strGeoID = Request.QueryString("GeoID").ToString()

            'Load the Plants listbox.
            Dim dtMktCampList As DataTable
            dtMktCampList = HPQ.Excalibur.SupplyChain.ListAllMktCamps_ByRegion("1", strGeoID)
            lbMktCamp.Items.Clear()
            lbMktCamp.DataSource = dtMktCampList
            lbMktCamp.DataBind()

            If SCMMktCamp <> "" Then
                Dim sSCMMktCamps() As String = SCMMktCamp.Split(";")
                Dim item As ListItem
                For Each item In lbMktCamp.Items
                    Dim i As Integer = 0
                    For i = 0 To sSCMMktCamps.Length - 1
                        If sSCMMktCamps(i) = item.Value Then
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
        Dim sMktCampID As String = ""
        Dim sMktCampName As String = ""
        Dim sProdVerID As String = ""
        Dim sProdBrandID As String = ""
        Dim sCategories As String = ""
        Dim strSelRegion As String = "-1"
        Dim item As ListItem
        Dim intPos As Integer = -1
        Dim intPos2 As Integer = -1

        For Each item In lbMktCamp.Items
            If item.Selected Then
                If sMktCampID = "" Then
                    sMktCampID = item.Value
                Else
                    sMktCampID = sMktCampID & "," & item.Value
                End If
            End If
        Next

        For Each item In lbMktCamp.Items
            If item.Selected Then
                If sMktCampName = "" Then
                    sMktCampName = item.Text
                Else
                    sMktCampName = sMktCampName & "," & item.Text
                End If
            End If
        Next

        ProcessFilter(sMktCampID, sMktCampName)

        'Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    End Sub

    Private Function ProcessFilter(ByVal sMktCampID As String, ByVal sMktCampName As String) As Boolean
        Try
            Dim URL As String = Nothing
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sMktCampID & ":" & sMktCampName))
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

    Protected Sub lbProductBands_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbMktCamp.SelectedIndexChanged

    End Sub

    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click

    End Sub
End Class
