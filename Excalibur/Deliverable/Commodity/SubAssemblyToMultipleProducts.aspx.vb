Imports System.Data

Partial Class SubAssemblyToMultipleProducts
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public ReadOnly Property DRID() As String
        Get
            Return Request("DRID")
        End Get
    End Property

    Public ReadOnly Property SA() As String
        Get
            Return Request("SA")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property

    Public ReadOnly Property SAType() As String '0 = Engineering Subassembly, 1 = Service Subassembly
        Get
            Return Request("SAType")
        End Get
    End Property

    Public ReadOnly Property SelectedProducts() As String
        Get
            Return Request("Selected")
        End Get
    End Property

    Public ReadOnly Property Assign() As String
        Get
            Return Request("Assign")
        End Get
    End Property

    Public Shared Property dt() As Data.DataTable
        Get
            Return (GetSessionStateValue("dt"))
        End Get
        Set(ByVal value As Data.DataTable)
            AddSessionStateValue("dt", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                'lblHeader.Text = DRID & "," & SAType & "," & Assign
                dt = dw.SelectProductsByDeliverable(DRID, SAType, Assign)
                If dt.Rows.Count = 0 Then
                    lblHeader.Text = "No Other Products Are Assigned This Deliverable"
                    btnSubmit.Visible = False
                Else
                    gvProducts.DataSource = dt
                    gvProducts.DataBind()

                    If SAType = 0 And Assign = 1 Then 'Subassemby and Assign to multiple products
                        lblHeader.Text = "Assign Engineering Subassembly No. " & SA
                        gvProducts.Columns(3).Visible = False
                    ElseIf SAType = 0 And Assign = 0 Then 'Subassembly and View products
                        lblHeader.Text = "List Other Products"
                        gvProducts.Columns(3).Visible = False
                        btnSubmit.Visible = False
                        btnCancel.Text = "Close"
                    ElseIf SAType = 1 And Assign = 1 Then 'Service Subassembly and Assign to multiple products
                        lblHeader.Text = "Assign Service Subassembly No. " & SA
                        gvProducts.Columns(2).Visible = False
                    ElseIf SAType = 1 And Assign = 0 Then 'Service Subassembly and View products
                        lblHeader.Text = "List Other Products"
                        gvProducts.Columns(2).Visible = False
                        btnSubmit.Visible = False
                        btnCancel.Text = "Close"
                    End If
                End If
            End If
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try
    End Sub

    Protected Sub cbxAll_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cbxAll As CheckBox = sender
        Dim row As GridViewRow
        For Each row In gvProducts.Rows
            Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxSingle")
            Dim lblServiceSubassembly As System.Web.UI.WebControls.Label = row.FindControl("lblServiceSubassembly")
            Dim lblSubassembly As System.Web.UI.WebControls.Label = row.FindControl("lblSubassembly")
            If cbxAll.Checked Then
                If cbx.Visible Then
                    cbx.Checked = True
                    If SAType = 0 Then
                        lblSubassembly.Text = SA
                        lblSubassembly.Font.Bold = True
                    ElseIf SAType = 1 Then
                        lblServiceSubassembly.Text = SA
                        lblServiceSubassembly.Font.Bold = True
                    End If
                End If
            Else
                If cbx.Visible Then
                    cbx.Checked = False
                    If SAType = 0 Then
                        lblSubassembly.Text = ""
                    ElseIf SAType = 1 Then
                        lblServiceSubassembly.Text = ""
                    End If
                End If
            End If
        Next
    End Sub

    Protected Sub cbxSingle_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cbxSingle As CheckBox = sender
        Dim lblServiceSubassembly As System.Web.UI.WebControls.Label = sender.parent.parent.FindControl("lblServiceSubassembly")
        Dim lblSubassembly As System.Web.UI.WebControls.Label = sender.parent.parent.FindControl("lblSubassembly")
        If cbxSingle.Checked Then
            If SAType = 0 Then
                lblSubassembly.Text = SA
                lblSubassembly.Font.Bold = True
            ElseIf SAType = 1 Then
                lblServiceSubassembly.Text = SA
                lblServiceSubassembly.Font.Bold = True
            End If
        Else
            If SAType = 0 Then
                lblSubassembly.Text = ""
            ElseIf SAType = 1 Then
                lblServiceSubassembly.Text = ""
            End If
        End If
    End Sub

    Protected Sub gvProducts_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvProducts.DataBound
        Dim sSelectedProducts() As String = SelectedProducts.Split(",")
        Dim row As GridViewRow
        For Each row In gvProducts.Rows
            Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxSingle")
            Dim lblServiceSubassembly As System.Web.UI.WebControls.Label = row.FindControl("lblServiceSubassembly")
            Dim lblSubassembly As System.Web.UI.WebControls.Label = row.FindControl("lblSubassembly")
            Dim lblPVID As System.Web.UI.WebControls.Label = row.FindControl("lblPVID")
            Dim lblProdDelRootID As System.Web.UI.WebControls.Label = row.FindControl("lblProdDelRootID")
            If lblPVID.Text = PVID Then
                cbx.Visible = False
            End If
            If Assign = 1 Then
                'If SAType = 0 And lblSubassembly.Text <> "" Then
                '    'cbx.Visible = False
                'ElseIf SAType = 1 And lblServiceSubassembly.Text <> "" Then
                '    'cbx.Visible = False
                'End If
                Dim i As Integer = 0
                For i = 0 To sSelectedProducts.Length - 1
                    If sSelectedProducts(i) = lblProdDelRootID.Text Then
                        cbx.Checked = True
                        If SAType = 0 Then
                            lblSubassembly.Text = SA
                            lblSubassembly.Font.Bold = True
                        Else
                            lblServiceSubassembly.Text = SA
                            lblServiceSubassembly.Font.Bold = True
                        End If
                    End If
                Next
            Else
                gvProducts.Columns(0).Visible = False
                'If SAType = 0 And lblSubassembly.Text = "" Then
                '    row.Visible = False
                'ElseIf SAType = 1 And lblServiceSubassembly.Text = "" Then
                '    row.Visible = False
                'End If
            End If
        Next
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim sSelectedProducts As String = ""
            Dim row As GridViewRow
            For Each row In gvProducts.Rows
                Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxSingle")
                Dim lblProdDelRootID As System.Web.UI.WebControls.Label = row.FindControl("lblProdDelRootID")
                If cbx.Checked Then
                    If sSelectedProducts = "" Then
                        sSelectedProducts = lblProdDelRootID.Text
                    Else
                        sSelectedProducts = sSelectedProducts & "," & lblProdDelRootID.Text
                    End If
                End If
            Next
            If sSelectedProducts = "" Then
                sSelectedProducts = "0"
            End If
            ProcessFilter(sSelectedProducts)
        Catch ex As Exception
            Response.Write(ex.InnerException.ToString)
        End Try
    End Sub

    Private Function ProcessFilter(ByVal sSelectedProducts As String) As Boolean
        Try
            Dim URL As String = Nothing
            Me.thisBody.Attributes.Add("onload", String.Format("window.returnValue = '{0}';window.close();", sSelectedProducts))
        Catch ex As Exception
            lblHeader.Text = ex.InnerException.ToString
        End Try
    End Function

End Class
