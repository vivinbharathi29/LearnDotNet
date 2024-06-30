Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml

Partial Class service_DeleteSKAV
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Ajax.Utility.RegisterTypeForAjax(GetType(service_DeleteSKAV))
        displayresults.Visible = False
        If request.querystring("del") = "yes" Then
            delComfirmed()
        End If
      
    End Sub
    Public Sub popuLateGrid(ByVal pbid As String, ByVal avno As String)
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand

        ' Response.Write(pbid & "<br>")
        'Response.Write(avno)

        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("DeleteSkuSpareKits", Data.CommandType.StoredProcedure)
        objDW.CreateParameter(objComm, "@p_ProductBrandId", Data.SqlDbType.Int, pbid, 10)
        objDW.CreateParameter(objComm, "@avno", Data.SqlDbType.VarChar, avno, 50)
        objDT = objDW.ExecuteCommandTable(objComm)
        If (objDT.Rows.Count > 0) Then
            'response.write(objDT.Rows.Count)
            FileList.DataSource = objDT
            FileList.DataBind()
            Label1.Text = ""
            displayresults.Visible = True
            Summary.Text = ""
        Else
            displayresults.Visible = False
            Label1.Text = "AV not mapped for this product, please try again"
            Summary.Text = ""
        End If
    End Sub


    Protected Sub DeleteButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteButton.Click
        delComfirmed()
    End Sub
    Public Sub delComfirmed()
        Summary.Text = "The following Spare Kit / AV's mapping(s) have been removed: <br><br>"

        Dim currentRowsFilePath, skno As String

        'Enumerate the GridViewRows
        For index As Integer = 0 To FileList.Rows.Count - 1
            Dim cb As CheckBox = CType(FileList.Rows(index).FindControl("RowLevelCheckBox"), CheckBox)

            'If it's checked, delete it...
            If cb.Checked Then
                currentRowsFilePath = FileList.DataKeys(index).Value

                skno = getskNO(currentRowsFilePath)
                deleteAVMap(currentRowsFilePath)
                displayresults.Visible = True
                Summary.Text &= String.Concat("<li>", skno & "</li>")
            End If
        Next
        ' Page.ClientScript.RegisterStartupScript(Me.GetType(), "myCloseScript", "reloadParentAndClose()", True)
        Button1.visible = True

    End Sub


    Protected Sub CheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckAll.Click
        'Enumerate each GridViewRow
        displayresults.Visible = True
        For Each gvr As GridViewRow In FileList.Rows
            'Programmatically access the CheckBox from the TemplateField
            Dim cb As CheckBox = CType(gvr.FindControl("RowLevelCheckBox"), CheckBox)

            'Check it!
            cb.Checked = True
        Next
    End Sub

    Protected Sub UncheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles UncheckAll.Click
        'Enumerate each GridViewRow
        displayresults.Visible = True
        For Each gvr As GridViewRow In FileList.Rows
            'Programmatically access the CheckBox from the TemplateField
            Dim cb As CheckBox = CType(gvr.FindControl("RowLevelCheckBox"), CheckBox)

            'Uncheck it!
            cb.Checked = False
        Next
    End Sub

    Private Sub deleteAVMap(ByVal sskmapid As String)
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim sql As String = "update servicesparekitmapav SET Status = 'D' " & _
              " where Id = '" & sskmapid & "'"
        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand(sql, Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)
        ' Response.Write(sql)
    End Sub
    'Private Sub deletespareKit(ByVal skId As String, sfPn As String, userName As string)
    '    Dim objDT As DataTable = Nothing
    '    Dim objDW As HPQ.Data.DataWrapper = Nothing
    '    Dim objComm As Data.SqlClient.SqlCommand

    '    objDW = New HPQ.Data.DataWrapper()
    '    objComm = objDW.CreateCommand("usp_DeleteSpareKit_ServiceFamily", Data.CommandType.StoredProcedure)
    '    objDW.CreateParameter(objComm, "@p_SpareKitId", Data.SqlDbType.Int, skId, 10)
    '    objDW.CreateParameter(objComm, "@p_ServiceFamilyPn", Data.SqlDbType.VarChar, sfPn, 50)
    '    objDW.CreateParameter(objComm, "@p_UserName", Data.SqlDbType.VarChar, userName, 50)
    '    objDT = objDW.ExecuteCommandTable(objComm)

    'End Sub


    Private Function getskNO(ByVal skmapid As String)
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim objRow As DataRow
        Dim rr As String
        objDW = New HPQ.Data.DataWrapper()
        objComm = objDW.CreateCommand("select sparekitno from ServiceSpareKitMapav tb1, ServiceSpareKitMap tb2, ServiceSpareKit tb3 where tb1.ServiceSpareKitMapid = tb2.id and tb3.id = tb2.sparekitid and tb1.id ='" & skmapid & "'", Data.CommandType.Text)
        objDT = objDW.ExecuteCommandTable(objComm)


        For Each objRow In objDT.Rows

            rr = objRow.Item("sparekitno").ToString
        Next

        Return rr
    End Function


    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        'System.Web.UI.ScriptManager.RegisterClientScriptBlock(Me.Button1, Me.GetType(), "Close_Window", "window.returnValue='refresh';window.close();", True)
        Page.ClientScript.RegisterStartupScript(Me.GetType(), "myCloseScript", "reloadParentAndClose()", True)
    End Sub

    Protected Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Dim pbid As String = Request.QueryString("brandID")
        Dim avno As String = Request.Form("avtext")
        popuLateGrid(pbid, avno)     
    End Sub
End Class