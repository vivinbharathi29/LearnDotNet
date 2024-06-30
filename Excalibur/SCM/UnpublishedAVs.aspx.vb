Imports System.Data

Partial Class UnpublishedAVs
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
                dtAVs = dw.SelectHiddenAvs(BID)
                If dtAVs.Rows.Count = 0 Then
                    lblHeader.Text = "There Are No Unpublished AVs To Report"
                Else
                    gvUnpublishedAVs.DataSource = dtAVs
                    gvUnpublishedAVs.DataBind()
                End If
            End If
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try
    End Sub

    Protected Sub cbxAll_Checkedchanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cbxAll As CheckBox = sender
        Dim row As GridViewRow
        For Each row In gvUnpublishedAVs.Rows
            Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxUnpublishedAVs")
            If cbxAll.Checked Then
                cbx.Checked = True
            Else
                cbx.Checked = False
            End If
        Next
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim row As GridViewRow
            For Each row In gvUnpublishedAVs.Rows
                Dim cbx As System.Web.UI.WebControls.CheckBox = row.FindControl("cbxUnpublishedAVs")
                Dim lblAvDetailID As System.Web.UI.WebControls.Label = row.FindControl("lblAvDetailID")
                If cbx.Checked Then
                    dw.UpdateAvStatus(lblAvDetailID.Text, 0) '0=unhide/activate
                End If
            Next
            Response.Write("<script language='javascript'> { window.close();}</script>")
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try
    End Sub
End Class
