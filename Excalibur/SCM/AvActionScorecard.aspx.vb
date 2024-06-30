Imports System.Data
Partial Class SCM_AvActionScorecard
    Inherits System.Web.UI.Page
    Dim DateRangeType As Integer '0=None, 1=Date Range, 2=Days
    Dim PVIDs As String = ""
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Private ReadOnly Property UserID() As String
        Get
            Return Request.QueryString("UserID")
        End Get
    End Property

    Public Shared Property dtProducts() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtProducts"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtProducts", Value)
        End Set
    End Property

    Public Shared Property dtPrograms() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtPrograms"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtPrograms", Value)
        End Set
    End Property

    Public Shared Property dtProductCycles() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtProductCycles"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtProductCycles", Value)
        End Set
    End Property

    Public Shared Property dtUserSettings() As System.Data.DataTable
        Get
            Return (GetSessionStateValue("dtUserSettings"))
        End Get
        Set(ByVal Value As System.Data.DataTable)
            AddSessionStateValue("dtUserSettings", Value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                dtUserSettings = dw.GetEmployeeUserSettings(UserID, 10)
                dtProducts = dw.SelectAvActionScorecardProducts
                dtPrograms = dw.ListPrograms(DBNull.Value.ToString)
                dtProductCycles = dw.SelectProductsByCycle()

                lbCycle.DataSource = dtPrograms
                lbCycle.DataBind()

                If dtUserSettings.Rows.Count > 0 Then
                    PopulateData()
                Else
                    lbProducts.DataSource = dtProducts
                    lbProducts.DataBind()
                End If
            Else
                lblErrorMessage.Text = ""
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Private Sub PopulateData()
        Try
            Dim sSettings() As String = dtUserSettings.Rows(0).Item("Setting").Split("|")

            rbStatus.SelectedValue = sSettings(0)

            Dim sPVIDs() As String = sSettings(1).Split(",")

            Dim colProductName As New DataColumn
            Dim colProductID As New DataColumn

            'Set Column DataTypes
            colProductName.DataType = System.Type.GetType("System.String")
            colProductID.DataType = System.Type.GetType("System.Int64")

            'Name Columns
            colProductName.ColumnName = "FullName"
            colProductID.ColumnName = "ID"

            'Add Columns
            Dim dtProductList As New DataTable
            dtProductList.Columns.Add(colProductName)
            dtProductList.Columns.Add(colProductID)

            'Load User Setting Values
            For Each row As DataRow In dtProducts.Rows
                Dim i As Integer = 0
                For i = 0 To sPVIDs.Length - 1
                    If sPVIDs(i) = row.Item("ID") Then
                        Dim row2 As DataRow
                        row2 = dtProductList.NewRow()
                        row2.Item("FullName") = row.Item("FullName")
                        row2.Item("ID") = row.Item("ID")
                        dtProductList.Rows.Add(row2)
                    End If
                Next
            Next

            'Load The Rest
            For Each row As DataRow In dtProducts.Rows
                Dim Exists As Boolean = False
                For Each row3 As DataRow In dtProductList.Rows
                    If row.Item("ID") = row3.Item("ID") Then
                        Exists = True
                    End If
                Next
                If Exists = False Then
                    Dim row4 As DataRow
                    row4 = dtProductList.NewRow()
                    row4.Item("FullName") = row.Item("FullName")
                    row4.Item("ID") = row.Item("ID")
                    dtProductList.Rows.Add(row4)
                End If
            Next

            lbProducts.DataSource = dtProductList
            lbProducts.DataBind()

            'Selected Chosen Products
            Dim item As ListItem
            Dim n As Integer = 0
            For n = 0 To sPVIDs.Length - 1
                For Each item In lbProducts.Items
                    If item.Value = sPVIDs(n) Then
                        item.Selected = True
                    End If
                Next
            Next

        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Protected Sub lbCycle_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbCycle.SelectedIndexChanged
        Try
            Dim iCycleID As Integer = lbCycle.SelectedItem.Value

            lbCycle.SelectedItem.Selected = False
            'lbCycle.SelectedItem.Enabled = False

            Dim colProductName As New DataColumn
            Dim colProductID As New DataColumn

            'Set Column DataTypes
            colProductName.DataType = System.Type.GetType("System.String")
            colProductID.DataType = System.Type.GetType("System.Int64")

            'Name Columns
            colProductName.ColumnName = "FullName"
            colProductID.ColumnName = "ID"

            'Add Columns
            Dim dtProductList As New DataTable
            dtProductList.Columns.Add(colProductName)
            dtProductList.Columns.Add(colProductID)

            Dim IDs As String = ""

            'Load Existing Selected Values
            Dim item As ListItem
            For Each item In lbProducts.Items
                Dim Exists As Boolean = False
                For Each row3 As DataRow In dtProductList.Rows
                    If item.Value = row3.Item("ID") Then
                        Exists = True
                    End If
                Next
                If Exists = False And item.Selected Then
                    Dim row As DataRow
                    row = dtProductList.NewRow()
                    row.Item("FullName") = item.Text
                    row.Item("ID") = item.Value
                    dtProductList.Rows.Add(row)
                    If IDs = "" Then
                        IDs = item.Value
                    Else
                        IDs = IDs & "," & item.Value
                    End If
                End If
            Next

            'Get Selected Cycle ProductIDs
            For Each row2 As DataRow In dtProductCycles.Rows
                If iCycleID = row2("CycleID") Then
                    'Load Selected Cycle Products
                    For Each item In lbProducts.Items
                        Dim Exists As Boolean = False
                        For Each row3 As DataRow In dtProductList.Rows
                            If item.Value = row3.Item("ID") Then
                                Exists = True
                            End If
                        Next
                        If Exists = False And item.Value = row2("ProductVersionID") Then
                            Dim row As DataRow
                            row = dtProductList.NewRow()
                            row.Item("FullName") = item.Text
                            row.Item("ID") = item.Value
                            dtProductList.Rows.Add(row)
                            If IDs = "" Then
                                IDs = item.Value
                            Else
                                IDs = IDs & "," & item.Value
                            End If
                        End If
                    Next
                End If
            Next

            'Load The Rest
            For Each item In lbProducts.Items
                Dim Exists As Boolean = False
                For Each row3 As DataRow In dtProductList.Rows
                    If item.Value = row3.Item("ID") Then
                        Exists = True
                    End If
                Next
                If Exists = False Then
                    Dim row As DataRow
                    row = dtProductList.NewRow()
                    row.Item("FullName") = item.Text
                    row.Item("ID") = item.Value
                    dtProductList.Rows.Add(row)
                End If
            Next

            lbProducts.DataSource = dtProductList
            lbProducts.DataBind()

            'Selected Chosen Products
            Dim sArray As String() = IDs.Split(",")
            Dim i As Integer = 0
            For i = 0 To sArray.Length - 1
                For Each item In lbProducts.Items
                    If item.Value = sArray(i) Then
                        item.Selected = True
                    End If
                Next
            Next

        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If rbStatus.SelectedIndex > -1 Then
                If ValidateDateRange() Then
                    If LoadProductList() Then
                        StoreEmployeeSettingsData()
                        ProcessReport()
                    Else
                        lblErrorMessage.Text = "Please Select Product(s) To Process"
                    End If
                Else
                    lblErrorMessage.Text = "Please Enter Date Range OR Day Span"
                End If
            Else
                lblErrorMessage.Text = "Please Select An AV Status"
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Private Sub StoreEmployeeSettingsData()
        Try
            Dim sSettings As String = rbStatus.SelectedItem.Value & "|" & PVIDs
            dw.UpdateEmployeeUserSetting(UserID, 10, sSettings)
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Private Function ProcessReport() As Boolean
        Try
            Dim URL As String = Nothing
            Dim applicationRoot As String = Session("ApplicationRoot")


            URL = "/iPulsar/ExcelExport/AvActionScorecard.aspx?FromDate=" & txtFromDt.SelectedValue.ToString & "&ToDate=" & _
txtToDt.SelectedValue.ToString & "&Days=" & txtDays.Text & "&DateRangeType=" & DateRangeType & "&PVIDs=" & PVIDs & _
"&Status=" & rbStatus.SelectedItem.Value

            Response.Write("<script language='javascript'> { window.parent.location = '" & URL & "';window.close();}</script>")
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Function

    Private Function ValidateDateRange() As Boolean
        Try
            Dim sTo As String = txtToDt.SelectedValue.ToString
            Dim sFrom As String = txtFromDt.SelectedValue.ToString

            If txtDays.Text.Trim = "" And txtFromDt.SelectedValue.ToString.Trim <> "" And txtToDt.SelectedValue.ToString.Trim <> "" Then
                DateRangeType = 1
                Return True
            ElseIf txtDays.Text.Trim <> "" And txtFromDt.SelectedValue.ToString.Trim = "" And txtToDt.SelectedValue.ToString.Trim = "" Then
                DateRangeType = 2
                Return True
            ElseIf txtDays.Text.Trim = "" And txtFromDt.SelectedValue.ToString.Trim = "" And txtToDt.SelectedValue.ToString.Trim = "" Then
                DateRangeType = 0
                Return True
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Function

    Private Function LoadProductList() As Boolean
        Try
            PVIDs = ""
            For Each item As ListItem In lbProducts.Items
                If item.Selected Then
                    If PVIDs = "" Then
                        PVIDs = item.Value
                    Else
                        PVIDs = PVIDs & "," & item.Value
                    End If
                End If
            Next
            If PVIDs <> "" Then
                Return True
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Function

End Class

