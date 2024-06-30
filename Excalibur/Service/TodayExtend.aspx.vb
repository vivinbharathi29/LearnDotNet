Imports System.Data

Partial Class MobileSE_Today_TodayExtend
    Inherits System.Web.UI.Page

    Private Property CurrentUserID() As String
        Get
            Return ViewState("CurrentUserID")
        End Get
        Set(ByVal value As String)
            ViewState("CurrentUserID") = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                Dim FIELD_NAME_DATAGRID_ORDER As String = String.Empty

                pnlNoData.Visible = False
                pnlData.Visible = False

                lnkUpdateSelectedAVs.Attributes.Add("OnClick()", "bUpdateAVsDeletedMapped(" + CurrentUserID + ")")
                Dim TableName As String = Request.QueryString("TableName")
                CurrentUserID = Request.QueryString("CurrentUserID")

                '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
                'CurrentUserID = 4025

                gvAVsDeletedMappedSPS.Visible = False
                gvAVsNotMappedToSPS.Visible = False
                gvSPSNotMappedToAV.Visible = False

                Select Case TableName
                    Case "AVsDeletedMappedSPS"
                        lblTitle.Text = "AVs Deleted but Mapped to SpareKits"
                        FIELD_NAME_DATAGRID_ORDER = "CategoryName"
                        gvAVsDeletedMappedSPS.Visible = True
                        lnkUpdateSelectedAVs.Visible = True
                        GetListAvsDeleted()
                    Case "AVsNotMappedToSPS"
                        lblTitle.Text = "AVs Not Mapped to SpareKits"
                        FIELD_NAME_DATAGRID_ORDER = "DotsName"
                        gvAVsNotMappedToSPS.Visible = True
                        lnkUpdateSelectedAVs.Visible = False
                        GetListAVsNotMappedToSPS()
                    Case "SPSNotMappedToAV"
                        lblTitle.Text = "SpareKits Not Mapped to AVs"
                        FIELD_NAME_DATAGRID_ORDER = "DotsName"
                        lnkUpdateSelectedAVs.Visible = False
                        gvSPSNotMappedToAV.Visible = True
                        GetListSPSNotMappedToAV()
                    Case Else
                        Response.Write("Unknown Table")
                End Select


                ' Initial Order and field
                Session("OrderField") = FIELD_NAME_DATAGRID_ORDER


            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetSortDirection(ByVal column As String) As String
        Try
            ' By default, set the sort direction to ascending.
            Dim sortDirection = "ASC"

            ' Retrieve the last column that was sorted.
            Dim sortExpression = TryCast(ViewState("SortExpression"), String)

            If sortExpression IsNot Nothing Then
                ' Check if the same column is being sorted.
                ' Otherwise, the default value can be returned.
                If sortExpression = column Then
                    Dim lastDirection = TryCast(ViewState("SortDirection"), String)
                    If lastDirection IsNot Nothing _
                      AndAlso lastDirection = "ASC" Then

                        sortDirection = "DESC"

                    End If
                End If
            End If

            ' Save new values in ViewState.
            ViewState("SortDirection") = sortDirection
            ViewState("SortExpression") = column

            Return sortDirection
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Avs Deleted and Mapped To SPS"

    Protected Sub gvAVsDeletedMappedSPS_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvAVsDeletedMappedSPS.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvAVsDeletedMappedSPS.DataSource = dtData
                gvAVsDeletedMappedSPS.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetListAvsDeleted()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.getAvsDeletedMappedToSPS()

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                gvAVsDeletedMappedSPS.DataSource = dtData
                gvAVsDeletedMappedSPS.DataBind()
                Session("dtData") = dtData
            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "There are not AV's Deleted and Mapped to Spare Kits."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub lnkUpdateSelectedAVs_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkUpdateSelectedAVs.ServerClick
        Try
            Dim chkItem As CheckBox
            Dim SelectedItems As New StringBuilder
            For Each item As GridViewRow In gvAVsDeletedMappedSPS.Rows
                If item.RowType = DataControlRowType.DataRow Then
                    chkItem = CType((item.Cells(0).FindControl("chkUpdateSelectedAVs")), CheckBox)
                    If chkItem.Checked = True Then
                        If SelectedItems.Length = 0 Then
                            SelectedItems.Append(item.Cells(1).Text)
                        Else
                            SelectedItems.Append(",")
                            SelectedItems.Append(item.Cells(1).Text)
                        End If
                    End If
                End If
            Next

            'save the values in the DataBase
            ' Response.Write("SelectedItems : " + SelectedItems.ToString)
            If SelectedItems.Length > 0 Then
                'txtHidUpdateSelectedAVs.Text = SelectedItems.ToString
                If bUpdateDeletedAvMappedToSPS(SelectedItems.ToString) = True Then
                    Response.Write("Deleted Av Mapped To SPS- Status Changed To Deleted")
                End If
            Else
                Response.Write("You must select the Avs Deleted mapped to unmapped first.")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function bUpdateDeletedAvMappedToSPS(ByVal AVs As String) As Boolean

        Dim res As Integer

        res = HPQ.Excalibur.Service.UpdateDeletedAvMappedToSPS(AVs, CurrentUserID)

        GetListAvsDeleted()
        Return True
    End Function

#End Region

#Region "AVs Not Mapped to SpareKits"

    Protected Sub gvAVsNotMappedToSPS_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvAVsNotMappedToSPS.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvAVsNotMappedToSPS.DataSource = dtData
                gvAVsNotMappedToSPS.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetListAVsNotMappedToSPS()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.getAVsNotMappedToSPS(CurrentUserID)

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                gvAVsNotMappedToSPS.DataSource = dtData
                gvAVsNotMappedToSPS.DataBind()
                Session("dtData") = dtData
            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "There are not AV's not Mapped to SpareKit."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region "SpareKits Not Mapped to AVs"

    Protected Sub gvSPSNotMappedToAV_Sorting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewSortEventArgs) Handles gvSPSNotMappedToAV.Sorting
        Try
            'Retrieve the table from the session object.
            Dim dtData As DataTable = CType(Session("dtData"), DataTable)

            If Not dtData Is Nothing Then
                dtData.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)
                gvSPSNotMappedToAV.DataSource = dtData
                gvSPSNotMappedToAV.DataBind()
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetListSPSNotMappedToAV()
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.getSPSNotMappedToAV(CurrentUserID)

            Me.pnlData.Visible = True
            Me.pnlNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                gvSPSNotMappedToAV.DataSource = dtData
                gvSPSNotMappedToAV.DataBind()
                Session("dtData") = dtData
            Else
                Me.pnlData.Visible = False
                Me.pnlNoData.Visible = True
                msgSearchNoData.Text = "There are not AV's not Mapped to SpareKit."
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region







End Class
