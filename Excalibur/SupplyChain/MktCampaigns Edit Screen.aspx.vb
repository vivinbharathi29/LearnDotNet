Imports System.Data
Partial Class SupMktCampaigns_Edit_Screen
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public strNewMktCampID As String = ""
    Public _bolNewRec As Boolean

    Public Property bolNewRec() As Boolean
        Get
            Return _bolNewRec
            'Return "Americas"
        End Get
        Set(ByVal value As Boolean)
            _bolNewRec = value
        End Set
    End Property

    Public ReadOnly Property bolInsertRec() As String
        Get
            '#If DEBUG Then
            '            Return "False"
            '#End If
            Return Request("bolInsertRec")
        End Get
    End Property

    Public ReadOnly Property CampaignID() As Integer
        Get
            '#If DEBUG Then
            '            Return "1"
            '#End If
            Return Request("CampaignID")
            'Return "1050"
        End Get
    End Property

    Public ReadOnly Property CampaignName() As String
        Get
            '#If DEBUG Then
            '            Return "Back To School"
            '#End If
            Return Request("CampaignName")
            'Return "763"
        End Get
    End Property

    Public ReadOnly Property StartDate() As String
        Get
            '#If DEBUG Then
            '            Return "2011-02-01"
            '#End If
            Return Request("StartDate")
        End Get
    End Property

    Public ReadOnly Property EndDate() As String
        Get
            '#If DEBUG Then
            '            Return "2012-10-31"
            '#End If
            Return Request("EndDate")
        End Get
    End Property

    Public ReadOnly Property GeoID() As String
        Get
            Return Request("GeoID")
        End Get
    End Property

    Public ReadOnly Property GeoID3() As String
        Get
            Return Request("GeoID3")
        End Get
    End Property

    Public ReadOnly Property ActiveCamp() As Boolean
        Get
            '#If DEBUG Then
            '            Return True
            '#End If
            Return Request("ActiveCamp")
        End Get
    End Property

    Public ReadOnly Property PlantID() As String
        Get
            '#If DEBUG Then
            '            Return "1,2"
            '#End If
            Return Request("PlantID")
        End Get
    End Property

    Public ReadOnly Property PlantName() As String
        Get
            '#If DEBUG Then
            '            Return "Houston Campus,Cupertino"
            '#End If
            Return Request("PlantName")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Me.Page.IsPostBack = False Then
            'bolNewRec = bolInsertRec 'The 'bolNewRec' variable will be used throughout the rest of the page.

            btnSaveYes.Visible = False
            btnSaveNo.Visible = False

            If (bolInsertRec = "True") Then
                chkNewRec.Checked = True
            Else
                chkNewRec.Checked = False
            End If

            If (chkNewRec.Checked = True) Then
                '_bolNewRec = True
                chkNewRec.Checked = True
                btnDelete.Visible = False
                chkActive.Checked = True
                btnCancel.Text = "Cancel"

                'TestLabel.Text = "GeoID: " & GeoID & " - GeoID3: " & GeoID3 & " - ActiveCamp: " & ActiveCamp & " - bolInsertRec: " & bolInsertRec
                'TestLabel.Visible = True
                'TestLabel.BackColor = Drawing.Color.White
                'TestLabel.ForeColor = Drawing.Color.Red
            Else
                '_bolNewRec = False
                chkNewRec.Checked = False
                btnDelete.Visible = True
                btnCancel.Text = "Close"

                txtMktCampID.Text = CampaignID
                txtMktCampName.Text = CampaignName
                txtStartDate.Text = Convert.ToDateTime(StartDate)
                txtEndDate.Text = Convert.ToDateTime(EndDate)
                chkActive.Checked = ActiveCamp

                'TestLabel.Text = "bolNewRec: " & bolNewRec
                'TestLabel.Visible = True
                'TestLabel.BackColor = Drawing.Color.White
                'TestLabel.ForeColor = Drawing.Color.Red
            End If

            Dim intPos As Integer = -1
            'Dim strGeoID As String = ""
            Dim strGeoID As String = GeoID
            Select Case GeoID
                Case "Americas"
                    strGeoID = "1"
                Case "EMEA"
                    strGeoID = "2"
                Case "APJ"
                    strGeoID = "3"
            End Select

            'TestLabel.Text = "GeoID Name: " & GeoID & " AND ID: " & strGeoID

            lbPlants.BackColor = Drawing.Color.White

            'Load the Plants listbox.
            Dim dtPlants As DataTable
            dtPlants = HPQ.Excalibur.SupplyChain.SelectRCTOPlants_ByGeoID(strGeoID)
            lbPlants.Items.Clear()
            lbPlants.DataSource = dtPlants
            lbPlants.DataBind()

            Dim item As ListItem
            If PlantID <> "" Then
                Dim sSCMPlants() As String = PlantID.ToString.Split(",")
                For Each item In lbPlants.Items
                    Dim i As Integer = 0
                    For i = 0 To sSCMPlants.Length - 1
                        If sSCMPlants(i) = item.Value Then
                            item.Selected = True
                        End If
                    Next
                Next
            End If

            'If (GeoID = "") Then
            '    If (GeoID3 = "") Then
            '        TestLabel.Text = "No GeoID is being passed to this page!"
            '    Else
            '        TestLabel.Text = "GeoID3 = " & GeoID3
            '    End If
            'Else
            '    TestLabel.Text = "GeoID = " & GeoID
            'End If
            'TestLabel.Visible = True
        End If

    End Sub

    'Protected Sub cmdCancel_onclick(ByVal sender As Object, ByVal e As System.EventArgs)
    '    'Response.Write("<script language='javascript'> {window.close();}</script>")

    '    Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", "-1"))

    'End Sub

    Protected Sub VerifySelectedPlants(ByVal strGeoIDValue As String)

        Dim bolcontinue As Boolean = True
        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim item As ListItem

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

        'Verify if the user changed the plants.
        Dim bolFound As Boolean = False
        If PlantID <> "" Then
            Dim sSCMPlants() As String = PlantID.ToString.Split(",")
            Dim i As Integer = 0
            For i = 0 To sSCMPlants.Length - 1
                bolFound = False
                Dim strSelPlantID() As String = sPlantID.Split(",")
                Dim j As Integer = 0
                For j = 0 To strSelPlantID.Length - 1
                    If (strSelPlantID(j) = sSCMPlants(i)) Then
                        bolFound = True

                        Exit For
                    End If
                Next

                If (bolFound = False) Then
                    btnDelete.Enabled = False
                    btnSave.Enabled = False
                    'lbPlants.BackColor = Drawing.Color.Yellow
                    btnSaveYes.Visible = True
                    btnSaveNo.Visible = True

                    TestLabel.Visible = True
                    TestLabel.Text = "You have changed the plants that belong to this Marketing Campaign.  This will cause all " _
                    & "data that has already been setup for this marketing campaign to be erased!  Are you sure you want to do this?"

                    bolcontinue = False
                    Exit For
                End If
            Next
        End If

        If (bolcontinue = True) Then
            'Now verify in the oposite direction.
            bolFound = False
            If PlantID <> "" Then
                Dim strSelPlantID() As String = sPlantID.Split(",")
                Dim j As Integer = 0
                For j = 0 To strSelPlantID.Length - 1
                    Dim sSCMPlants() As String = PlantID.ToString.Split(",")
                    Dim i As Integer = 0
                    For i = 0 To sSCMPlants.Length - 1
                        bolFound = False
                        If (sSCMPlants(i) = strSelPlantID(j)) Then
                            bolFound = True

                            Exit For
                        End If
                    Next

                    If (bolFound = False) Then
                        btnDelete.Enabled = False
                        btnSave.Enabled = False
                        'lbPlants.BackColor = Drawing.Color.Yellow
                        btnSaveYes.Visible = True
                        btnSaveNo.Visible = True

                        TestLabel.Visible = True
                        TestLabel.Text = "You have changed the plants that belong to this Marketing Campaign.  This will cause all " _
                        & "data that has already been setup for this marketing campaign to be erased!  Are you sure you want to do this?"

                        bolcontinue = False
                        Exit For
                    End If
                Next
            End If
        End If

        'This means the user didn't change any of the plants.  Continue on to save the record with the original plants.
        If (bolcontinue = True) Then
            Call SaveMktCampRec(PlantID, PlantName, strGeoIDValue)
        End If
    End Sub

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        'If this is a new record then there is no need to verify if the plant selections has changed.
        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim item As ListItem

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

        If (chkNewRec.Checked = True) Then
            Call SaveMktCampRec(sPlantID, sPlantName, GeoID)
        Else 'This is an update to the currently selected Marketing Campaign record and the plant selection verifer needs to check if the user changed any of the plants.
            If (PlantID = "") Then
                Call SaveMktCampRec(sPlantID, sPlantName, GeoID)
            Else 'Only verify if there was already plants associated.
                Call VerifySelectedPlants(GeoID)
            End If
        End If

    End Sub

    Private Sub SaveMktCampRec(ByVal strPlantID As String, ByVal strPlantName As String, ByVal strGeoIDValue As String)

        TestLabel.BackColor = Drawing.Color.Red

        btnSaveYes.Visible = False
        btnSaveNo.Visible = False
        TestLabel.Text = ""
        btnDelete.Enabled = True
        btnSave.Enabled = True

        'Validate required fields.
        If (txtMktCampName.Text = "") Then
            txtMktCampName.BackColor = Drawing.Color.Yellow
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            txtMktCampName.BackColor = Drawing.Color.White
        End If

        If (txtStartDate.Text = "") Then
            txtStartDate.BackColor = Drawing.Color.Yellow
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            txtStartDate.BackColor = Drawing.Color.White
        End If

        If (txtEndDate.Text = "") Then
            txtEndDate.BackColor = Drawing.Color.Yellow
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red
            Exit Sub
        Else
            txtEndDate.BackColor = Drawing.Color.White
        End If

        If (ValidateDate(txtStartDate.Text) = False) Then
            TestLabel.Text = "The Start Date text box is not a valid date, change it and try again!"
            TestLabel.Visible = True
            TestLabel.BackColor = Drawing.Color.Yellow
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red

            Exit Sub
        Else
            TestLabel.BackColor = Drawing.Color.White
        End If

        If (ValidateDate(txtEndDate.Text) = False) Then
            TestLabel.Text = "The End Date text box is not a valid date, change it and try again!"
            TestLabel.Visible = True
            TestLabel.BackColor = Drawing.Color.Yellow
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red

            Exit Sub
        Else
            TestLabel.BackColor = Drawing.Color.White
        End If

        If (CDate(txtStartDate.Text) > CDate(txtEndDate.Text)) Then
            TestLabel.Text = "The End Date text box cannot have a date that is prior to the start date, change it and try again!"
            TestLabel.Visible = True
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red
            txtEndDate.BackColor = Drawing.Color.Yellow

            Exit Sub
        Else
            txtEndDate.BackColor = Drawing.Color.White
        End If

        Dim strGeoID As String = ""
        Select Case strGeoIDValue
            Case "Americas"
                strGeoID = "1"
            Case "EMEA"
                strGeoID = "2"
            Case "APJ"
                strGeoID = "3"
        End Select

        'TestLabel.Text = strGeoID
        'testLabel.Visible = True
        'Exit Sub

        If (chkNewRec.Checked = True) Then 'Insert a new record.
            'First Check if the new record the user is entering already exists.
            If (HPQ.Excalibur.SupplyChain.DupRecCheck(strGeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text) = True) Then
                TestLabel.Visible = True
                TestLabel.Text = "A record with the same name, start date, end date and region already exists in Excalibur!" _
                & vbCrLf & vbCrLf & "Make a change to the new Marketing Campaign you are trying to create and try again."

                Exit Sub
            End If

            HPQ.Excalibur.SupplyChain.MktCampaigns_Insert(strGeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text, strPlantID, strPlantName)

            'Now get the ID of the Marketing Campaign that was just created.
            txtMktCampID.Text = HPQ.Excalibur.SupplyChain.MktCampaign_ReturnMktCampID(strGeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text)

            TestLabel.Text = "The new record has been added successfully!"
            TestLabel.Visible = True
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red

            btnDelete.Visible = True
            chkNewRec.Checked = False
            btnCancel.Text = "Close"
        Else 'Update the existing record.
            If (strGeoID = "") Then
                If (chkActive.Checked = True) Then
                    HPQ.Excalibur.SupplyChain.MktCampaigns_Update(txtMktCampID.Text, GeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text, "True", strPlantID, strPlantName)
                Else
                    HPQ.Excalibur.SupplyChain.MktCampaigns_Update(txtMktCampID.Text, GeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text, "False", strPlantID, strPlantName)
                End If
            Else
                If (chkActive.Checked = True) Then
                    HPQ.Excalibur.SupplyChain.MktCampaigns_Update(txtMktCampID.Text, strGeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text, "True", strPlantID, strPlantName)
                Else
                    HPQ.Excalibur.SupplyChain.MktCampaigns_Update(txtMktCampID.Text, strGeoID, txtMktCampName.Text, txtStartDate.Text, txtEndDate.Text, "False", strPlantID, strPlantName)
                End If
            End If

            'TestLabel.Text = "strGeoID: " & strGeoID & " GeoID: " & GeoID
            TestLabel.Text = "The record has been updated successfully!"
            TestLabel.Visible = True
            TestLabel.BackColor = Drawing.Color.White
            TestLabel.ForeColor = Drawing.Color.Red
        End If

    End Sub

    Private Function ValidateDate(ByVal strDateToValidate As String) As Boolean
        If IsDate(strDateToValidate) Then
            ValidateDate = True
        Else
            ValidateDate = False
        End If
    End Function

    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        HPQ.Excalibur.SupplyChain.MktCampaigns_Delete(txtMktCampID.Text)
        chkNewRec.Checked = True

        txtMktCampName.Text = ""
        txtStartDate.Text = ""
        txtEndDate.Text = ""
        chkActive.Checked = False

        TestLabel.Text = "This record has successfully been deleted."
        TestLabel.Visible = True
        TestLabel.BackColor = Drawing.Color.White
        TestLabel.ForeColor = Drawing.Color.Red

    End Sub

    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load

    End Sub

    Protected Sub btnSaveYes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveYes.Click

        'TestLabel.Text = GeoID
        'TestLabel.Visible = True
        'Exit Sub

        Dim sPlantID As String = ""
        Dim sPlantName As String = ""
        Dim item As ListItem

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

        'Add the code to delete all AV Marketing Campaign records for the currently selected Marketing Campaign.
        HPQ.Excalibur.SupplyChain.MktCampaigns_AVDetailProductBrandDeleteByMktCampOnly(txtMktCampID.Text)

        Call SaveMktCampRec(sPlantID, sPlantName, GeoID)

    End Sub

    Protected Sub btnSaveNo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSaveNo.Click

        Dim strGeoID As String = ""
        Select Case GeoID
            Case "Americas"
                strGeoID = "1"
            Case "EMEA"
                strGeoID = "2"
            Case "APJ"
                strGeoID = "3"
        End Select

        lbPlants.BackColor = Drawing.Color.White

        'Load the Plants listbox.
        Dim dtPlants As DataTable
        dtPlants = HPQ.Excalibur.SupplyChain.SelectRCTOPlants_ByGeoID(strGeoID)
        lbPlants.Items.Clear()
        lbPlants.DataSource = dtPlants
        lbPlants.DataBind()

        Dim item As ListItem
        If PlantID <> "" Then
            Dim sSCMPlants() As String = PlantID.ToString.Split(",")
            For Each item In lbPlants.Items
                Dim i As Integer = 0
                For i = 0 To sSCMPlants.Length - 1
                    If sSCMPlants(i) = item.Value Then
                        item.Selected = True
                    End If
                Next
            Next
        End If

        Call SaveMktCampRec(PlantID, PlantName, GeoID)

    End Sub
End Class
