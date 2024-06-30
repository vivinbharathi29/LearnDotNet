Partial Class InitialOfferingFilter2
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                Dim dtCategories As Data.DataTable = dw.SelectInitialOfferingCategories()
                Dim dtCategories2 As Data.DataTable = dw.SelectCommodityGuidanceCategories()
                Dim dtProductPrograms As Data.DataTable = dw.SelectCommodityGuidanceProductPrograms()

                If Request.Cookies("PreferredLayout2") Is Nothing Then
                    preferredLayout.Value = "pulsar2"
                Else
                    preferredLayout.Value = Request.Cookies("PreferredLayout2").Value
                End If

                lblProductProgram.Visible = False
                ddlProductProgram.Visible = False
                ddlCategory2.Visible = False

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                dtProductPrograms.Rows.Add(0, "", 0, "", "")
                dtProductPrograms.DefaultView.Sort = String.Format("ComboName", "{0}")
                dtProductPrograms = dtProductPrograms.DefaultView.ToTable

                ddlProductProgram.DataTextField = "ComboName"
                ddlProductProgram.DataValueField = "ProgramID"

                ddlProductProgram.DataSource = dtProductPrograms

                ddlProductProgram.DataBind()

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                dtCategories.Rows.Add("", 0)
                dtCategories.DefaultView.Sort = String.Format("Name", "{0}")
                dtCategories = dtCategories.DefaultView.ToTable

                ddlCategory.DataTextField = "Name"
                ddlCategory.DataValueField = "ID"

                ddlCategory.DataSource = dtCategories

                ddlCategory.DataBind()

                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                dtCategories2.Rows.Add("", 0)
                dtCategories2.DefaultView.Sort = String.Format("Name", "{0}")
                dtCategories2 = dtCategories2.DefaultView.ToTable

                ddlCategory2.DataTextField = "Name"
                ddlCategory2.DataValueField = "ID"

                ddlCategory2.DataSource = dtCategories

                ddlCategory2.DataBind()
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.InnerException.ToString
        End Try
    End Sub

    Protected Sub rblReportType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rblReportType.SelectedIndexChanged
        Try
            If rblReportType.SelectedItem.Value = 0 Then
                lblProductProgram.Visible = False
                ddlProductProgram.Visible = False
                ddlCategory2.Visible = False
                lblBusUnit.Visible = True
                ddlBusUnit.Visible = True
                ddlCategory.Visible = True
            ElseIf rblReportType.SelectedItem.Value = 1 Then
                lblProductProgram.Visible = True
                ddlProductProgram.Visible = True
                ddlCategory2.Visible = True
                lblBusUnit.Visible = False
                ddlBusUnit.Visible = False
                ddlCategory.Visible = False
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.InnerException.ToString
        End Try
    End Sub
    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim URL As String = Nothing
            Dim applicationRoot As String = Session("ApplicationRoot")
            Dim onLoadScript1 As String = " var retValue = '[retValue]'; if ($('#preferredLayout').val() == 'pulsar2') { intialOfferings(retValue);  return; } "
            Dim onLoadScript2 As String = " window.returnValue = retValue;window.close();"
            Dim onLoadScript As String
            Dim retValue As String

            If ddlBusUnit.SelectedItem.Value <> 0 And ddlCategory.SelectedItem.Value <> 0 And rblReportType.SelectedItem.Value = 0 Then
                'Initial Offering
                retValue = ddlBusUnit.SelectedItem.Value & "," & ddlCategory.SelectedItem.Value
                onLoadScript = (onLoadScript1 & onLoadScript2).Replace("[retValue]", retValue)
                Me.thisBody.Attributes.Add("onload", onLoadScript)
            ElseIf rblReportType.SelectedItem.Value = 1 And ddlCategory2.SelectedItem.Value <> 0 And ddlProductProgram.SelectedItem.Value <> 0 Then
                'Commodity Guidance
                retValue = 0 & "," & ddlProductProgram.SelectedItem.Value & "," & ddlCategory2.SelectedItem.Value & "," & ddlProductProgram.SelectedItem.Text
                onLoadScript = (onLoadScript1 & onLoadScript2).Replace("[retValue]", retValue)
                Me.thisBody.Attributes.Add("onload", onLoadScript)
            Else
                lblHeader1.ForeColor = Drawing.Color.Red
                lblHeader1.Text = "Please Select Values From Both Dropdown Lists"
                Exit Sub
            End If
        Catch ex As Exception
            lblHeader1.Text = ex.InnerException.ToString
        End Try
    End Sub

End Class