Imports System.Data

Partial Class DCRWorkflow
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public ReadOnly Property DCRID() As String
        Get
            Return Request("DCRID")
        End Get
    End Property

    Public ReadOnly Property UserID() As String
        Get
            Return Request("UserID")
        End Get
    End Property

    Public ReadOnly Property PVID() As String
        Get
            Return Request("PVID")
        End Get
    End Property
    Public ReadOnly Property RTPDate() As String
        Get
            Return Request("RTPDate")
        End Get
    End Property
    Public ReadOnly Property EMDate() As String
        Get
            Return Request("EMDate")
        End Get
    End Property

    Public Shared Property dtDefinitions() As Data.DataTable
        Get
            Return (GetSessionStateValue("dtDefinitions"))
        End Get
        Set(ByVal value As Data.DataTable)
            AddSessionStateValue("dtDefinitions", value)
        End Set
    End Property

    Public Shared Property dtWorkflows() As Data.DataTable
        Get
            Return (GetSessionStateValue("dtWorkflows"))
        End Get
        Set(ByVal value As Data.DataTable)
            AddSessionStateValue("dtWorkflows", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                'Dim dtDefinitionsFiltered As DataTable
                hdnRTPDate.Value = Convert.ToString(Request("RTPDate"))
                hdnEMDate.Value = Convert.ToString(Request("EMDate"))
                hdnIsPulsarProduct.Value = Convert.ToString(dw.IsPulsarProduct(PVID))
                dtWorkflows = dw.SelectDCRWorkflows()
                dtDefinitions = dw.SelectDCRWorkflowsDefinitions()

                'For Testing Purposes
                'If UserID <> 5016 Then
                '    dtWorkflows.Rows(7).Delete()
                'End If

                dtWorkflows.Rows.Add("", 0)
                dtWorkflows.DefaultView.Sort = String.Format("{0}", "Name")

                ddlWorkflow.DataSource = dtWorkflows.DefaultView.ToTable
                ddlWorkflow.DataBind()

                'Dim row1 As DataRow
                'For Each row1 In dtWorkflows.Rows
                lblDescription.Text = ""
                'Exit For
                'Next

                'dtDefinitionsFiltered = dtDefinitions
                'dtDefinitionsFiltered.DefaultView.RowFilter = "(WorkflowID = '1')"
                gvWorkflowDefintions.Visible = False
                'gvWorkflowDefintions.DataSource = dtDefinitionsFiltered
                'gvWorkflowDefintions.DataBind()
            End If
        Catch ex As Exception
            Response.Write("Error")
        End Try
    End Sub

    Protected Sub ddlWorkflow_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlWorkflow.SelectedIndexChanged
        If ddlWorkflow.SelectedItem.Value = 0 Then
            gvWorkflowDefintions.Visible = False
            lblDescription.Text = ""
            Exit Sub
        End If

        Dim row As DataRow
        For Each row In dtWorkflows.Rows
            If ddlWorkflow.SelectedItem.Value = row("ID") Then
                lblDescription.Text = row("Description")
                Exit For
            End If
        Next

        Dim dtDefinitionsFiltered As New DataTable
        dtDefinitionsFiltered = dtDefinitions
        dtDefinitionsFiltered.DefaultView.RowFilter = "(WorkflowID = '" & ddlWorkflow.SelectedItem.Value & "')"
        dtDefinitionsFiltered.DefaultView.Sort = "WorkflowID, MilestoneOrder"
        dtDefinitionsFiltered.DefaultView.ToTable()

        gvWorkflowDefintions.Visible = True
        gvWorkflowDefintions.DataSource = dtDefinitionsFiltered
        gvWorkflowDefintions.DataBind()
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            dw.InsertDCRWorkflowHistory(DCRID, ddlWorkflow.SelectedItem.Value, UserID, PVID, hdnRTPDate.Value, hdnEMDate.Value)
            Response.Write("<script language='javascript'> { window.close();}</script>")
        Catch ex As Exception
            Response.Write(ex.InnerException)
        End Try
    End Sub
End Class
