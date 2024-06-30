Imports System.Data

Partial Class AvsMissingData
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public Shared Function GetSessionStateValue(ByRef id As String) As Object
        Return System.Web.HttpContext.Current.Session(id)
    End Function

    Public Shared Sub AddSessionStateValue(ByRef id As String, ByRef obj As Object)
        System.Web.HttpContext.Current.Session.Add(id, obj)
    End Sub

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property IsPDM() As String
        Get
            Return Request("IsPDM")
        End Get
    End Property

    Public Shared Property dtAVs() As Data.DataTable
        Get
            Return (GetSessionStateValue("dtAVs"))
        End Get
        Set(ByVal value As Data.DataTable)
            AddSessionStateValue("dtAVs", value)
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Me.Page.IsPostBack Then
                dtAVs = dw.SelectAvWithMissingData(BID)
                If dtAVs.Rows.Count = 0 Then
                    lblHeader.Text = "There Are No AVs With Missing Corporate Data To Report"
                Else
                    gvAVsMissingData.DataSource = dtAVs
                    gvAVsMissingData.DataBind()
                End If
            End If
        Catch ex As Exception
            lblHeader.Visible = False
            gvAVsMissingData.Visible = False
            Response.Write(ex.ToString)
        End Try
    End Sub

    Protected Sub gvAVsMissingData_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles gvAVsMissingData.DataBound
        Dim row As GridViewRow
        For Each row In gvAVsMissingData.Rows
            Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxAVsMissingData")
            Dim lblStatus As System.Web.UI.WebControls.Label = row.FindControl("lblStatus")
            'Dim lblRasDiscoSysUpdate As System.Web.UI.WebControls.Label = row.FindControl("lblRasDiscoSysUpdate")
            'Dim lblMktDescSysUpdate As System.Web.UI.WebControls.Label = row.FindControl("lblMktDescSysUpdate")
            'Dim lblCplBlindSysUpdate As System.Web.UI.WebControls.Label = row.FindControl("lblCplBlindSysUpdate")

            If IsPDM = "True" Then
                cbx.Enabled = False
            ElseIf lblStatus.Text = "H" Then
                cbx.Visible = False
            End If

            'If lblRasDiscoSysUpdate.Text = 0 Then

            'End If
        Next

        If IsPDM = "True" Then
            Dim cbxAll As System.Web.UI.WebControls.CheckBox = gvAVsMissingData.HeaderRow.FindControl("cbxAll")
            cbxAll.Enabled = False
            btnCancel.Enabled = False
        End If
    End Sub

    Protected Sub cbxAll_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cbxAll As CheckBox = sender
        Dim row As GridViewRow
        For Each row In gvAVsMissingData.Rows
            Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxAVsMissingData")
            If cbxAll.Checked Then
                cbx.Checked = True
            Else
                cbx.Checked = False
            End If
        Next
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If IsPDM = "False" Then
                Dim row As GridViewRow
                For Each row In gvAVsMissingData.Rows
                    Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxAVsMissingData")
                    Dim lblAvDetailID As System.Web.UI.WebControls.Label = row.FindControl("lblAvDetailID")
                    If cbx.Checked Then
                        dw.UpdateAvStatus(lblAvDetailID.Text, 1) '0=hide
                    End If
                Next
            End If
            Response.Write("<script language='javascript'> {parent.window.parent.modalDialog.cancel(true);}</script>")
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try
    End Sub

End Class
