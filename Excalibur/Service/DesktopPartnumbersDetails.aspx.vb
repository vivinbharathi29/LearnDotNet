Imports System.Data

Partial Class Service_DesktopPartnumbersDetails
    Inherits System.Web.UI.Page

    Public ServiceFamilyPartNumner, SpsPartNumber, SpsDescription As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            ServiceFamilyPartNumner = Request.QueryString("servicefamilypn")
            SpsPartNumber = Request.QueryString("spskitpn")
            SpsDescription = Request.QueryString("spskitdescription")

            If Not Page.IsPostBack Then
                pnlBomData.Visible = True
                pnlNoBomData.Visible = False
                If ServiceFamilyPartNumner <> String.Empty Then
                    Dim l_User As String = LCase(Session("LoggedInUser"))
                    lblUserName.Text = GetUserName(l_User)

                    ' lblTitle.Text = lblTitle.Text + "- ServiceFamilyPn: " + ServiceFamilyPartNumner
                    lblSpsPartNumber.Text = SpsPartNumber
                    lblSparekitDescription.Text = SpsDescription

                    GetCategories()
                    GetCustomerLevel()
                   GetSparekitBom(SpsPartNumber)

                    ' see if the sparekit already exists.
                    GetKitDetails(SpsPartNumber)

                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub bntAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bntAdd.Click
        Try
            If bValidateFields() = True Then
                Dim iRes As Integer
                iRes = HPQ.Excalibur.Service.InsertDesktopSparekitAvNumber(ServiceFamilyPartNumner, lblSpsPartNumber.Text, SpsDescription, ddlCategories.SelectedValue, ddlCustomerLevel.SelectedValue, ddlDisposition.SelectedValue, ddlWarranty.SelectedValue, ddlLocalStockAdvice.SelectedValue, txtFirstServiceDt.Text, txtRslComment.InnerText, chkGeosNa.Checked, chkGeosLa.Checked, chkGeosApj.Checked, chkGeosEmea.Checked, txtSupplier.Text)
            
                If iRes = -1 Then
                    Dim script As String = "window.opener.location.reload(true);window.close();"
                    System.Web.UI.ScriptManager.RegisterClientScriptBlock(Page, Me.[GetType](), "CloseWindow", script, True)
                End If



            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function bValidateFields() As Boolean
        Try
            bValidateFields = True

               If txtFirstServiceDt.Text <> String.Empty Then
                If Not IsDate(txtFirstServiceDt.Text) Then
                    bValidateFields = False
                    lblError.Text = "First Service Date: This field have to be a date."
                    Exit Function
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub GetKitDetails(ByVal sparekitNumber As String)
        Try
            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetSparekitAV(sparekitNumber, ServiceFamilyPartNumner)

            If dtData.Rows.Count > 0 Then
                ' get the sparekit Data
                Dim r As DataRow = dtData.Rows(0)

                'If Not IsDBNull(r("ID").ToString()) AndAlso r("ID").ToString() <> String.Empty Then
                '    SparekitID = r("ID").ToString()
                'End If


                If Not IsDBNull(r("SpareCategoryId").ToString()) AndAlso r("SpareCategoryId").ToString() <> String.Empty Then
                    ddlCategories.SelectedValue = r("SpareCategoryId").ToString()
                End If

                 If Not IsDBNull(r("FirstServiceDt").ToString()) AndAlso r("FirstServiceDt").ToString() <> String.Empty Then
                    txtFirstServiceDt.Text = r("FirstServiceDt").ToString()
                End If

                If Not IsDBNull(r("Comments").ToString()) AndAlso r("Comments").ToString() <> String.Empty Then
                    txtRslComment.InnerText = r("Comments").ToString()
                End If

                If Not IsDBNull(r("CsrLevelId").ToString()) AndAlso r("CsrLevelId").ToString() <> String.Empty Then
                    ddlCustomerLevel.SelectedValue = r("CsrLevelId").ToString()
                End If

                If Not IsDBNull(r("Disposition").ToString()) AndAlso r("Disposition").ToString() <> String.Empty Then
                    ddlDisposition.SelectedValue = r("Disposition").ToString()
                End If

                If Not IsDBNull(r("WarrantyTier").ToString()) AndAlso r("WarrantyTier").ToString() <> String.Empty Then
                    ddlWarranty.SelectedValue = r("WarrantyTier").ToString()
                End If

                If Not IsDBNull(r("LocalStockAdvice").ToString()) AndAlso r("LocalStockAdvice").ToString() <> String.Empty Then
                    ddlLocalStockAdvice.SelectedValue = r("LocalStockAdvice").ToString()
                End If

                If Not IsDBNull(r("GeoNa").ToString()) AndAlso r("GeoNa").ToString() <> String.Empty Then
                    chkGeosNa.Checked = CBool(r("GeoNa").ToString())
                End If

                If Not IsDBNull(r("GeoLa").ToString()) AndAlso r("GeoLa").ToString() <> String.Empty Then
                    chkGeosLa.Checked = CBool(r("GeoLa").ToString())
                End If

                If Not IsDBNull(r("GeoApj").ToString()) AndAlso r("GeoApj").ToString() <> String.Empty Then
                    chkGeosApj.Checked = CBool(r("GeoApj").ToString())
                End If

                If Not IsDBNull(r("GeoEmea").ToString()) AndAlso r("GeoEmea").ToString() <> String.Empty Then
                    chkGeosEmea.Checked = CBool(r("GeoEmea").ToString())
                End If

                If Not IsDBNull(r("Supplier").ToString()) AndAlso r("Supplier").ToString() <> String.Empty Then
                    txtSupplier.Text = r("Supplier").ToString()
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetUserName(ByVal User As String) As String
        Try
            Dim Security As New HPQ.Excalibur.Security(User)

            GetUserName = Security.CurrentUserFullName()

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub GetSparekitBom(ByVal sparekitNumber As String)
        Try
            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetSparekitBom(sparekitNumber)

            If dtData.Rows.Count > 0 Then
                gvBomData.Visible = True
                gvBomData.DataSource = dtData
                gvBomData.DataBind()


                'Do Until rs.Eof
                '    If rs("Level1") & "" <> lastLevel1 And rs("Level1") <> "" Then
                '        returnValue = returnValue & "<tr><td>" & rs("Level1") & "" & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & rs("L1Description") & "" & "</td></tr>"
                '        lastLevel1 = rs("Level1")
                '    End If

                '    If rs("Level2") & "" <> lastLevel2 And rs("Level2") & "" <> "" Then
                '        returnValue = returnValue & "<tr><td>&nbsp;</td><td>" & rs("Level2") & "" & "</td><td>&nbsp;</td><td>" & rs("L2SortString") & "" & "</td><td>" & rs("L2Description") & "" & "</td></tr>"
                '        lastLevel2 = rs("Level2")
                '    End If

                '    If rs("Level3") & "" <> lastLevel3 And rs("Level3") & "" <> "" Then
                '        returnValue = returnValue & "<tr><td>&nbsp;</td><td>&nbsp;</td><td>" & rs("Level3") & "" & "</td><td>" & rs("L3SortString") & "" & "</td><td>" & rs("L3Description") & "" & "</td></tr>"
                '        lastLevel3 = rs("Level3")
                '    End If

                '    rs.MoveNext()
                'Loop

            Else
                pnlBomData.Visible = False
                pnlNoBomData.Visible = True
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetCustomerLevel()
        Try

            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetCustomerLevel()

            ddlCustomerLevel.Items.Add(New ListItem("-- Select Customer Level --", "0"))

            If dtData.Rows.Count > 0 Then
                For Each elem As DataRow In dtData.Rows
                    ddlCustomerLevel.Items.Add(New ListItem(elem("CsrDescription"), elem("ID")))
                Next

                'ddlCustomerLevel.DataSource = dtData

                'ddlCustomerLevel.DataValueField = "ID"
                'ddlCustomerLevel.DataTextField = "CsrDescription"
                'ddlCustomerLevel.DataBind()

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetCategories()
        Try

            Dim objData As New HPQ.Excalibur.Data

            Dim dtData As New DataTable
            dtData = objData.ListServiceSpareCategories()

            ddlCategories.Items.Add(New ListItem("-- Select Category --", "0"))

            If dtData.Rows.Count > 0 Then
                For Each elem As DataRow In dtData.Rows
                    ddlCategories.Items.Add(New ListItem(elem("CategoryName"), elem("ID")))
                Next

                'ddlCategories.DataSource = dtData
                'ddlCategories.DataTextField = "CategoryName"
                'ddlCategories.DataValueField = "ID"
                'ddlCategories.DataBind()
            End If

            '<option value="0">-- Select Category --</option>
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

   


End Class
