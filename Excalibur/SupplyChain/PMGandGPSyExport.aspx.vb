Imports System.Data

Partial Class SupSCM_PMGandGPSyExport
    Inherits System.Web.UI.Page
    Dim PVIDs As String = ""
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Private ReadOnly Property PVID() As String
        Get
            Return Request.QueryString("PVID")
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

    Public Shared Property bClosePage() As Boolean
        Get
            Return (GetSessionStateValue("bClosePage"))
        End Get
        Set(ByVal Value As Boolean)
            AddSessionStateValue("bClosePage", Value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                dtProducts = dw.SelectPMGandGPSyChangesProducts()
                dtPrograms = dw.ListPrograms(DBNull.Value.ToString)
                dtProductCycles = dw.SelectProductsByCycle()

                lbCycle.DataSource = dtPrograms
                lbCycle.DataBind()

                bClosePage = False
                lblHidden.Value = bClosePage

                PopulateData()


            Else
                lblErrorMessage.Text = ""
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Private Sub PopulateData()
        Try
            Dim sPVID As String = PVID

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
                If sPVID = row.Item("ID") Then
                    Dim row2 As DataRow
                    row2 = dtProductList.NewRow()
                    row2.Item("FullName") = row.Item("FullName")
                    row2.Item("ID") = row.Item("ID")
                    dtProductList.Rows.Add(row2)
                End If
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
            For Each item In lbProducts.Items
                If item.Value = sPVID Then
                    item.Selected = True
                End If
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
                If LoadProductList() Then
                    ProcessReport()
                Else
                    lblErrorMessage.Text = "Please Select Product(s) To Process"
                End If
            Else
                lblErrorMessage.Text = "Please Select A Description Type"
            End If
        Catch ex As Exception
            lblErrorMessage.Text = ex.Message
        End Try
    End Sub

    Private Function ProcessReport() As Boolean
        Try
            If rbStatus.SelectedValue = 0 Then
                Dim dtFileCount As DataTable = Nothing
                Dim i As Integer = 0
                Dim iFileCount As Integer = 0

                dtFileCount = dw.SelectPMG100CharChangesFileCount(PVIDs)
                iFileCount = dtFileCount.Rows(0).Item("FileCount")

                Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", PVIDs & ";" & iFileCount))
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
