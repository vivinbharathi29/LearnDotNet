Imports System.Data

Partial Class Service_ServiceAdvancedSearchReports
    Inherits System.Web.UI.Page


    'Const REPORT_TYPE_SERVICE_BOM_AVNUMBERS As Integer = 1
    Const REPORT_TYPE_AV_TO_SPS As Integer = 2
    'Const REPORT_TYPE_SPS_TO_PRODUCTS As Integer = 3
    Const REPORT_TYPE_SERVICE_SPS_BOM As Integer = 4
    Const REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS As Integer = 5
    Const REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS As Integer = 6
    Const REPORT_TYPE_SPS_BY_CATEGORY As Integer = 7
    Const REPORT_TYPE_USED_BY As Integer = 8
    Const REPORT_TYPE_RSL_CHANGELOG As Integer = 9

    Private Const EXPORT_EXCEL As String = "1"
    Private Const EXPORT_WORD As String = "2"
    Private Const EXPORT_CSV As String = "3"

    Const USEDBY_FAMILYNAME As String = "1"
    Const USEDBY_SPAREKITS As String = "2"
    Const USEDBY_SUBASSEMBLY As String = "3"
    Const USEDBY_COMPONENT As String = "4"

    Public sMessage As String


    Private Property ReportType() As Integer
        Get
            Return ViewState("ReportType")
        End Get
        Set(ByVal value As Integer)
            ViewState("ReportType") = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                pnlAvBomReport.Visible = False 'No active option
                pnlSpsToProducts.Visible = False 'No active option

                pnlSpsBom.Visible = True
                pnlSpsByCategory.Visible = False
                pnlAvToSPS.Visible = False
                pnlSkuToSpareKits.Visible = False
                btnSpareKitsPlus.Visible = False
                pnlFamilyToSkuToSparekits.Visible = False
                pnlUsedBy.Visible = False
                pnlRSLChangeLog.Visible = False

                ReportType = REPORT_TYPE_SERVICE_SPS_BOM
                btnReport.OnClientClick = "frmServiceAdvancedSearchReports.onsubmit = function() {return true;}"
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReport.Click
        Try
            If Not Page.IsValid Then
                Return
            Else
                Dim Param As String = "ServiceAdvancedSearchReportsSummary.aspx?ReportType=" + rdReportType.SelectedValue + "&ExportType=" + ddlReportFormat.SelectedValue
                Dim sMultipleData As String = String.Empty
                lblError.Text = String.Empty

                Select Case ReportType
                    Case REPORT_TYPE_SERVICE_SPS_BOM
                        sMultipleData = GetMultipleDataCommaSeparated(txtSPSBom.Text)
                        If ValidateSpareKits(sMultipleData) = False Then
                            Exit Sub
                        End If
                        Param += "&txtSPSBom=" + sMultipleData
                    Case REPORT_TYPE_SPS_BY_CATEGORY
                        'Read the categories selected
                        Dim sbSpareKitCategory As StringBuilder = New StringBuilder()
                        For Each elem As ListItem In lstSpareCategory.Items
                            If elem.Selected Then
                                sbSpareKitCategory.Append(elem.Value & ",")
                            End If
                        Next
                        If sbSpareKitCategory.Length > 0 Then sbSpareKitCategory.Remove(sbSpareKitCategory.Length - 1, 1)
                        Param += "&txtSpsCategoryIDs=" + sbSpareKitCategory.ToString
                    Case REPORT_TYPE_AV_TO_SPS
                        sMultipleData = GetMultipleDataCommaSeparated(txtAVNumber.Text)
                        If ValidateAvNumbers(sMultipleData) = False Then
                            Exit Sub
                        End If
                        Param += "&txtAVNumber=" + sMultipleData.Replace("#", "%23")
                    Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                        sMultipleData = GetMultipleDataCommaSeparated(Trim(txtSKUNumber.Text))
                        If ValidateSkuNumbers(sMultipleData) = False Then
                            Exit Sub
                        End If
                        Param += "&txtSKUNumber=" + sMultipleData.Replace("#", "%23")
                    Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                        lblError.Text = String.Empty
                        'Read the Products selected
                        Dim sbProducts As StringBuilder = New StringBuilder()
                        For Each elem As ListItem In lstProducts.Items
                            If elem.Selected Then
                                sbProducts.Append(elem.Value & ",")
                            End If
                        Next
                        If sbProducts.Length > 0 Then sbProducts.Remove(sbProducts.Length - 1, 1)
                        Param += "&ProductVersionID=" + sbProducts.ToString

                        If txtSkuNumberProduct.Text <> String.Empty Then
                            sMultipleData = GetMultipleDataCommaSeparated(txtSkuNumberProduct.Text)
                            If ValidateSkuNumbers(sMultipleData) = False Then
                                Exit Sub
                            End If
                            Param += "&txtSkuNumberProduct=" + txtSkuNumberProduct.Text.Replace("#", "%23")
                        End If

                    Case REPORT_TYPE_USED_BY
                        Dim UsedBy As String = rdUsedBy.SelectedValue
                        Select Case UsedBy
                            Case USEDBY_FAMILYNAME
                                Param += "&UsedByType=" + USEDBY_FAMILYNAME
                                'Read the Products selected
                                Dim sbProducts As StringBuilder = New StringBuilder()
                                For Each elem As ListItem In lstUsedByProducts.Items
                                    If elem.Selected Then
                                        sbProducts.Append(elem.Value & ",")
                                    End If
                                Next
                                If sbProducts.Length > 0 Then sbProducts.Remove(sbProducts.Length - 1, 1)
                                Param += "&ProductVersionID=" + sbProducts.ToString
                                If txtUsedBySKU.Text <> String.Empty Then
                                    sMultipleData = GetMultipleDataCommaSeparated(txtUsedBySKU.Text)
                                    If ValidateSkuNumbers(sMultipleData) = False Then
                                        Exit Sub
                                    End If
                                    Param += "&txtUsedBySkuNumber=" + txtUsedBySKU.Text.Replace("#", "%23")
                                Else
                                    sMessage = "Sku Number: You have to write a SKU Number"
                                    lblError.Text = sMessage
                                    Exit Sub
                                End If
                            Case USEDBY_SPAREKITS
                                sMultipleData = GetMultipleDataCommaSeparated(txtUsedBy.Text)
                                If ValidateSpareKits(sMultipleData) = False Then
                                    Exit Sub
                                End If
                                Param += "&UsedByType=" + USEDBY_SPAREKITS + "&txtUsedByNumber=" + sMultipleData.Replace("#", "%23")
                            Case USEDBY_SUBASSEMBLY
                                sMultipleData = GetMultipleDataCommaSeparated(txtUsedBy.Text)
                                If ValidateSubAssemblies(sMultipleData) = False Then
                                    Exit Sub
                                End If
                                Param += "&UsedByType=" + USEDBY_SUBASSEMBLY + "&txtUsedByNumber=" + sMultipleData.Replace("#", "%23")
                            Case USEDBY_COMPONENT
                                sMultipleData = GetMultipleDataCommaSeparated(txtUsedBy.Text)
                                If ValidateComponents(sMultipleData) = False Then
                                    Exit Sub
                                End If
                                Param += "&UsedByType=" + USEDBY_COMPONENT + "&txtUsedByNumber=" + sMultipleData.Replace("#", "%23")
                        End Select
                    Case REPORT_TYPE_RSL_CHANGELOG
                        'Read the Product Family selected
                        Dim sbProducts As StringBuilder = New StringBuilder()
                        For Each elem As ListItem In lstProductFamilies.Items
                            If elem.Selected Then
                                sbProducts.Append(elem.Value & ",")
                            End If
                        Next
                        If sbProducts.Length > 0 Then sbProducts.Remove(sbProducts.Length - 1, 1)
                        Param += "&ProductVersionID=" + sbProducts.ToString

                End Select

                OpenWindow(Param)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Try
            lblError.Text = String.Empty

            Select Case ReportType
                Case REPORT_TYPE_SERVICE_SPS_BOM
                    txtSPSBom.Text = String.Empty
                Case REPORT_TYPE_SPS_BY_CATEGORY
                    'clear all Category Names
                    For Each elem As ListItem In lstSpareCategory.Items
                        elem.Selected = False
                    Next
                Case REPORT_TYPE_AV_TO_SPS
                    txtAVNumber.Text = String.Empty
                Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                    txtSKUNumber.Text = String.Empty
                Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                    'clear all Product Names
                    txtSkuNumberProduct.Text = String.Empty
                    For Each elem As ListItem In lstProducts.Items
                        elem.Selected = False
                    Next
                Case REPORT_TYPE_USED_BY
                    Select Case rdUsedBy.SelectedValue
                        Case USEDBY_FAMILYNAME
                            txtUsedBySKU.Text = String.Empty
                            For Each elem As ListItem In lstProducts.Items
                                elem.Selected = False
                            Next
                        Case USEDBY_SPAREKITS
                            txtUsedBy.Text = String.Empty
                        Case USEDBY_SUBASSEMBLY
                            txtUsedBy.Text = String.Empty
                        Case USEDBY_COMPONENT
                            txtUsedBy.Text = String.Empty
                    End Select
                Case REPORT_TYPE_RSL_CHANGELOG
                    For Each elem As ListItem In lstProductFamilies.Items
                        elem.Selected = False
                    Next
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub OpenWindow(ByVal QueryString As String)
        If ddlReportFormat.SelectedValue = EXPORT_EXCEL Or ddlReportFormat.SelectedValue = EXPORT_CSV Or ddlReportFormat.SelectedValue = EXPORT_WORD Then
            Server.Transfer(QueryString, True) ' Previouspage OK - Open in same window
        Else
            'HTML 
            ClientScript.RegisterStartupScript(Page.GetType(), "Open", "javascript:window.open('" + QueryString + "', '_blank')", True)
        End If
        'Server.Transfer(QueryString, True) ' Previouspage OK - Open in same window
        'window.open('WebForm2.aspx','_new')
        'ClientScript.RegisterStartupScript(Page.GetType(), "Open", "javascript:window.open('" + QueryString + "', '_blank')", True)
        '' jScript = String.Format("window.open('" + QueryString + "', '_blank');", sURL)
        'ScriptManager.RegisterClientScriptBlock(Me, GetType(Page), "Open", "javascript:window.open('" + QueryString + "')", True)
    End Sub

    Protected Sub rdReportType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdReportType.SelectedIndexChanged
        Try
            ReportType = rdReportType.SelectedValue
            rdUsedBy.SelectedValue = USEDBY_SPAREKITS
            pnlSpsToProducts.Visible = False 'No active option
            pnlAvBomReport.Visible = False 'No active option

            pnlSpsBom.Visible = False
            pnlSpsByCategory.Visible = False
            pnlAvToSPS.Visible = False
            pnlSkuToSpareKits.Visible = False
            pnlFamilyToSkuToSparekits.Visible = False
            pnlUsedBy.Visible = False
            pnlRSLChangeLog.Visible = False

            ClearFilters()

            Select Case ReportType
                Case REPORT_TYPE_SERVICE_SPS_BOM
                    pnlSpsBom.Visible = True
                    'txtSPSBom.Visible = True
                Case REPORT_TYPE_SPS_BY_CATEGORY
                    pnlSpsByCategory.Visible = True
                    getSpareKitCategories()
                Case REPORT_TYPE_AV_TO_SPS
                    pnlAvToSPS.Visible = True
                Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS
                    pnlSkuToSpareKits.Visible = True
                    btnSpareKitsPlus.Visible = True
                Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
                    pnlFamilyToSkuToSparekits.Visible = True
                    getProductNames(REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS) ' Family Products
                Case REPORT_TYPE_USED_BY
                    pnlUsedBy.Visible = True
                    lstUsedByProducts.Visible = False
                    txtUsedBySKU.Visible = False
                    txtUsedBy.Visible = True
                    lblUsedBySku.Visible = False
                    rUsedByTxt.Enabled = True
                    rUsedByProduct.Enabled = False

                    lblUsedBy.Text = "Used By Sparekits"
                    lblMaxPartNumber.Visible = True
                    lblMaxPartNumber.Text = "(Max 50 sparekits)"
                Case REPORT_TYPE_RSL_CHANGELOG
                    pnlRSLChangeLog.Visible = True
                    getProductNames(REPORT_TYPE_RSL_CHANGELOG) ' Family Products
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub rdUsedBy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdUsedBy.SelectedIndexChanged
        Try
            Dim UsedBy As String = rdUsedBy.SelectedValue

            txtUsedBy.Text = String.Empty
            lblError.Text = String.Empty

            lstUsedByProducts.Visible = False
            lblUsedBySku.Visible = False
            txtUsedBySKU.Visible = False

            rUsedByTxt.Enabled = False
            rUsedByProduct.Enabled = False

            txtUsedBy.Visible = False
            lblMaxPartNumber.Visible = False


            Select Case UsedBy
                Case USEDBY_FAMILYNAME
                    getUsedByProductNames() ' Used By Family Products
                    lblUsedBy.Text = "Used By Product Name"
                    lstUsedByProducts.Visible = True
                    txtUsedBySKU.Visible = True
                    lblUsedBySku.Visible = True
                    rUsedByProduct.Enabled = True
                    'lblMaxPartNumber.Visible = False
                Case USEDBY_SPAREKITS
                    lblUsedBy.Text = "Used By Sparekits"
                    lblMaxPartNumber.Visible = True
                    lblMaxPartNumber.Text = "(Max 50 Sparekits)"
                    txtUsedBy.Visible = True
                    rUsedByTxt.Enabled = True
                Case USEDBY_SUBASSEMBLY
                    lblUsedBy.Text = "Used By SubAssembly"
                    lblMaxPartNumber.Visible = True
                    lblMaxPartNumber.Text = "(Max 100 SubAssemblies)"
                    txtUsedBy.Visible = True
                    rUsedByTxt.Enabled = True
                Case USEDBY_COMPONENT
                    lblUsedBy.Text = "Used By Component"
                    lblMaxPartNumber.Visible = True
                    lblMaxPartNumber.Text = "(Max 100 Components)"
                    txtUsedBy.Visible = True
                    rUsedByTxt.Enabled = True
            End Select

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub chkUnselectProducts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUnselectProducts.CheckedChanged
        Try
            For Each elem As ListItem In lstProducts.Items
                elem.Selected = False
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub chkProductsRSLChangeLog_CheckedChanged(sender As Object, e As EventArgs) Handles chkProductsRSLChangeLog.CheckedChanged
        Try
            For Each elem As ListItem In lstProductFamilies.Items
                elem.Selected = False
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ClearFilters()
        Try
            lblError.Text = String.Empty
            txtAVNumber.Text = String.Empty
            txtSKUNumber.Text = String.Empty
            txtSpsNumber.Text = String.Empty
            txtSPSBom.Text = String.Empty
            txtUsedBySKU.Text = String.Empty
            txtBomAvNumbers.Text = String.Empty
            txtSkuNumberProduct.Text = ""

            'clear all Product Names
            For Each elem As ListItem In lstProducts.Items
                elem.Selected = False
            Next
            'clear all Category Names
            For Each elem As ListItem In lstSpareCategory.Items
                elem.Selected = False
            Next

            'clear all Product Names
            For Each elem As ListItem In lstProductFamilies.Items
                elem.Selected = False
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function GetMultipleDataCommaSeparated(ByVal sUserData As String) As String
        Try
            GetMultipleDataCommaSeparated = sUserData

            If sUserData <> String.Empty Then
                Dim sNewData As New StringBuilder

                For Each elem As String In sUserData.Split(vbCr)
                    If elem.Trim <> String.Empty Then
                        'if the element exists, we do not add to the stringbuilder
                        If InStr(sNewData.ToString, elem.Trim, CompareMethod.Text) = 0 Then
                            sNewData.Append(elem.Trim & ",")
                        End If
                    End If
                Next
                Return Left(sNewData.ToString, Len(sNewData.ToString) - 1)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub getProductNames(ByVal PanelName As Integer)
        Try
            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetServiceList_ProductsVersionNames()

            If dtData.Rows.Count > 0 Then
                If PanelName = REPORT_TYPE_RSL_CHANGELOG Then
                    lstProductFamilies.DataSource = dtData
                    lstProductFamilies.DataTextField = "dotsname"
                    lstProductFamilies.DataValueField = "ServiceFamilyPn" '"ProductVersionID" '
                    lstProductFamilies.DataBind()
                End If

                If PanelName = REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS Then
                    lstProducts.DataSource = dtData
                    lstProducts.DataTextField = "dotsname"
                    lstProducts.DataValueField = "ServiceFamilyPn" '"ProductVersionID" '
                    lstProducts.DataBind()
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getUsedByProductNames()
        Try
            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetServiceList_ProductsVersionNames()

            If dtData.Rows.Count > 0 Then
                lstUsedByProducts.DataSource = dtData
                lstUsedByProducts.DataTextField = "dotsname"
                lstUsedByProducts.DataValueField = "ProductVersionID" 'ServiceFamilyPn
                lstUsedByProducts.DataBind()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSpareKitCategories()
        Try
            Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
            Dim dtData As DataTable = dw.ListServiceSpareCategories()

            lstSpareCategory.DataSource = dtData
            lstSpareCategory.DataTextField = "CategoryName"
            lstSpareCategory.DataValueField = "ID"
            lstSpareCategory.DataBind()

            dtData = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "VALIDATIONS"

    Private Function ValidateBomAvNumbers(ByVal BomAvNumbers As String) As Boolean
        Try
            ValidateBomAvNumbers = True

            If ValidateAvNumbers(BomAvNumbers) = False Then
                ValidateBomAvNumbers = False
            Else
                If Me.txtBomAvNumbers.Text <> String.Empty Then
                    If ValidateAvNumberBaseUnit() = True Then
                        'If we have more than one AV number...
                        If txtBomAvNumbers.Text.Split(",").Length > 1 Then
                            'All av number myust have the same Kmat
                            If ValidateAvNumberKMAT() = True Then
                                Return True
                            Else
                                sMessage = "Invalid AvNumbers list.The Av Numbers do not share the same KMAT."
                                lblError.Text = sMessage
                                Return False
                            End If
                        Else
                            Return True
                        End If
                    Else
                        sMessage = "Invalid AvNumbers list. One or more of the Av Numbers in the list do not have a Base Unit Category."
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSpareKits(ByVal SPSNumbers As String) As Boolean
        Try
            ValidateSpareKits = True
            If SPSNumbers.Trim <> String.Empty Then
                If SPSNumbers.Trim.Length > 10 Then
                    'If we have more than one SPSNumber. We look for the "," char in the string
                    If InStr(SPSNumbers, ",", CompareMethod.Text) = 0 Then
                        sMessage = "SpareKit Numbers: You have to validate the SpareKit Numbers.  SpareKits contain between 6 and 10 characters, and are separated by comma."
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
                'More than one SPSNumbers
                If SPSNumbers.Split(",").Length > 1 Then
                    For Each elem As String In SPSNumbers.Split(",")
                        If ValidateSpareKit(elem.Trim) = False Then
                            lblError.Text = sMessage
                            Return False
                        End If
                    Next
                Else
                    Return ValidateSpareKit(SPSNumbers.Trim)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSpareKit(ByVal SpareKitNumber As String) As Boolean
        Try
            ValidateSpareKit = True
            'SPAREKIT NUMBER: They have between 6 and 10 characters
            'Validate Length
            If SpareKitNumber.Length > 10 Then
                sMessage = "SpareKit Number: You have to validate the SpareKit Numbers. SpareKits can not contain more than 10 characters."
                lblError.Text = sMessage
                Return False
            End If

            If SpareKitNumber.Length < 6 Then
                sMessage = "SpareKit Number: You have to validate the SpareKit Numbers. SpareKits can not contain less than 6 characters."
                lblError.Text = sMessage
                Return False
            End If

            If SpareKitNumber.Length > 7 Then
                If SpareKitNumber.Length < 10 Then
                    sMessage = "SpareKit Number: You have to validate the SpareKit Numbers. SpareKits must contain between 6 and 10 characters."
                    lblError.Text = sMessage
                    Return False
                End If
            End If

            If SpareKitNumber.Length > 6 Then
                If InStr(SpareKitNumber, "-", CompareMethod.Text) = 0 Then
                    sMessage = "SpareKit Number: You have to validate the SpareKit Numbers. All SpareKits with more than 6 characters must contains the character '-'."
                    lblError.Text = sMessage
                    Return False
                End If
                If Right(Left(SpareKitNumber, 7), 1) <> "-" Then ' '-' position 
                    sMessage = "SpareKit Number: You have to validate the SpareKit Numbers. All SpareKits with more than 6 characters must contains the Character '-' in the 7 Character."
                    lblError.Text = sMessage
                    Return False
                End If
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateComponents(ByVal ComponentNumbers As String) As Boolean
        Try
            ValidateComponents = True
            If ComponentNumbers.Trim <> String.Empty Then
                If ComponentNumbers.Trim.Length > 10 Then
                    'If we have more than one ComponentNumber. We look for the "," char in the string
                    If InStr(ComponentNumbers, ",", CompareMethod.Text) = 0 Then
                        sMessage = "Component Numbers: You have to validate the Component Numbers.  Components contain between 6 and 10 characters, and are separated by comma."
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
                'More than one ComponentNumbers
                If ComponentNumbers.Split(",").Length > 1 Then
                    For Each elem As String In ComponentNumbers.Split(",")
                        If ValidateComponent(elem.Trim) = False Then
                            lblError.Text = sMessage
                            Return False
                        End If
                    Next
                Else
                    Return ValidateComponent(ComponentNumbers.Trim)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateComponent(ByVal ComponentNumber As String) As Boolean
        Try
            ValidateComponent = True
            'Compoenent NUMBER: They have between 6 and 10 characters
            'Validate Length
            If ComponentNumber.Length > 10 Then
                sMessage = "Component Number: You have to validate the Component Numbers. Components can not contain more than 10 characters."
                lblError.Text = sMessage
                Return False
            End If

            If ComponentNumber.Length < 6 Then
                sMessage = "Component Number: You have to validate the Component Numbers. Components can not contain less than 6 characters."
                lblError.Text = sMessage
                Return False
            End If

            If ComponentNumber.Length > 7 Then
                If ComponentNumber.Length < 10 Then
                    sMessage = "Component Number: You have to validate the SpareKit Numbers. Components must contain between 6 and 10 characters."
                    lblError.Text = sMessage
                    Return False
                End If
            End If

            'first 6 characters in the string must be numbers
            If Not IsNumeric(Left(ComponentNumber, 6)) Then
                sMessage = "Component Number: You have to validate the Component Numbers. First 6 characters must be numeric."
                lblError.Text = sMessage
                Return False
            End If

            If ComponentNumber.Length > 6 Then
                If InStr(ComponentNumber, "-", CompareMethod.Text) = 0 Then
                    sMessage = "Component Number: You have to validate the Component Numbers. All Components with more than 6 characters must contains the character '-'."
                    lblError.Text = sMessage
                    Return False
                End If
                If Right(Left(ComponentNumber, 7), 1) <> "-" Then ' '-' position 
                    sMessage = "Component Number: You have to validate the Component Numbers. All Components with more than 6 characters must contains the Character '-' in the 7 Character."
                    lblError.Text = sMessage
                    Return False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSubAssemblies(ByVal SubAssemblyNumbers As String) As Boolean
        Try
            ValidateSubAssemblies = True
            If SubAssemblyNumbers.Trim <> String.Empty Then
                If SubAssemblyNumbers.Trim.Length > 10 Then
                    'If we have more than one ComponentNumber. We look for the "," char in the string
                    If InStr(SubAssemblyNumbers, ",", CompareMethod.Text) = 0 Then
                        sMessage = "SubAssembly Numbers: You have to validate the SubAssembly Numbers. SubAssemblies contain between 6 and 10 characters, and are separated by comma."
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
                'More than one SubAssembly numbers
                If SubAssemblyNumbers.Split(",").Length > 1 Then
                    For Each elem As String In SubAssemblyNumbers.Split(",")
                        If ValidateSubAssembly(elem.Trim) = False Then
                            lblError.Text = sMessage
                            Return False
                        End If
                    Next
                Else
                    Return ValidateSubAssembly(SubAssemblyNumbers.Trim)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSubAssembly(ByVal SubAssemblyNumber As String) As Boolean
        Try
            ValidateSubAssembly = True
            'SubAssembly NUMBER: They have between 6 and 10 characters
            'Validate Length
            If SubAssemblyNumber.Length > 10 Then
                sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. SubAssemblies can not contain more than 10 characters."
                lblError.Text = sMessage
                Return False
            End If

            If SubAssemblyNumber.Length < 6 Then
                sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. SubAssemblies can not contain less than 6 characters."
                lblError.Text = sMessage
                Return False
            End If

            If SubAssemblyNumber.Length > 7 Then
                If SubAssemblyNumber.Length < 10 Then
                    sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. SubAssemblies must contain between 6 and 10 characters."
                    lblError.Text = sMessage
                    Return False
                End If
            End If

            'first 6 characters in the string must be numbers
            If Not IsNumeric(Left(SubAssemblyNumber, 6)) Then
                sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. First 6 characters must be numeric."
                lblError.Text = sMessage
                Return False
            End If

            If SubAssemblyNumber.Length > 6 Then
                If InStr(SubAssemblyNumber, "-", CompareMethod.Text) = 0 Then
                    sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. All SubAssemblies with more than 6 characters must contains the character '-'."
                    lblError.Text = sMessage
                    Return False
                End If
                If Right(Left(SubAssemblyNumber, 7), 1) <> "-" Then ' '-' position 
                    sMessage = "SubAssembly Number: You have to validate the SubAssembly Numbers. All SubAssemblies with more than 6 characters must contains the Character '-' in the 7 Character."
                    lblError.Text = sMessage
                    Return False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateAvNumbers(ByVal AvNumbers As String) As Boolean
        Try
            ValidateAvNumbers = True
            If AvNumbers.Trim <> String.Empty Then
                If AvNumbers.Trim.Length > 11 Then
                    'If we have more than one AvNumber. We look for the "," char in the string
                    If InStr(AvNumbers, ",", CompareMethod.Text) = 0 Then
                        sMessage = "Av Numbers: You have to validate the AvNumbers. Av Numbers contain between 7 and 11 characters, and are separated by comma"
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
                'More than one AvNumber
                If AvNumbers.Split(",").Length > 1 Then
                    For Each elem As String In AvNumbers.Split(",")
                        If ValidateAVNumber(elem.Trim) = False Then
                            Return False
                        End If
                    Next
                Else
                    Return ValidateAVNumber(AvNumbers.Trim)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateAVNumber(ByVal AVNumber As String) As Boolean
        Try
            ValidateAVNumber = True
            'AV NUMBER: It is a mask similar to -----AV, or -----AV#---
            'They have between 7 and 11 characters

            'Validate Length
            If AVNumber.Length > 11 Then
                sMessage = "Av Numbers: You have to validate the AV Numbers. Avs can not contain more than 11 characters."
                lblError.Text = sMessage
                Return False
            End If

            If AVNumber.Length < 7 Then
                sMessage = "Av Numbers: You have to validate the AV Numbers. Avs can not contain less than 7 characters."
                lblError.Text = sMessage
                Return False
            End If

            If AVNumber.Length > 7 Then
                If AVNumber.Length < 11 Then
                    sMessage = "Av Numbers: You have to validate the AV Numbers. Avs must be between 7 and 11 characters."
                    lblError.Text = sMessage
                    Return False
                End If
            End If

            If InStr(AVNumber.ToUpper, "AV", CompareMethod.Text) = 0 Then 'AVCHAR
                sMessage = "Av Numbers: You have to validate the AV Numbers. All Avs must contains the character 'AV'."
                lblError.Text = sMessage
                Return False
            End If

            If Right(Left(AVNumber, 7), 2).ToUpper <> "AV" Then 'AV position 
                sMessage = "Av Numbers: You have to validate the AV Numbers. All Avs must contains the Character AV in the 6 and 7 Characters."
                lblError.Text = sMessage
                Return False
            End If

            If AVNumber.Length > 7 Then 'the Avnumber must have the # char in the 8 char position and followed by three Letters
                If Right(Left(AVNumber, 8), 1) <> "#" Then
                    sMessage = "Av Numbers: You have to validate the AV Numbers. All Avs must contains the Character # in the 8 Character."
                    lblError.Text = sMessage
                    Return False
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSkuNumber(ByVal SKUNumber As String) As Boolean
        Try
            ValidateSkuNumber = True
            'SKU NUMBER: They have between 11 characters. In the 8 Character they must have the '#' character.

            'Validate Length
            If SKUNumber.Length > 11 Or SKUNumber.Length < 11 Then
                sMessage = "SKU Number: You have to validate the SKU Numbers. SKU Number have 11 characters."
                lblError.Text = sMessage
                Return False
            End If

            If InStr(SKUNumber.ToUpper, "#", CompareMethod.Text) = 0 Then
                sMessage = "SKU Number: You have to validate the SKU Numbers. All SKU Numbers must contains the character '#'."
                lblError.Text = sMessage
                Return False
            End If

            If Right(Left(SKUNumber, 8), 1) <> "#" Then
                sMessage = "SKU Number: You have to validate the SKU Numbers. All SKU Numbers must contains the Character # in the 8 Character."
                lblError.Text = sMessage
                Return False
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateSkuNumbers(ByVal SkuNumbers As String) As Boolean
        Try
            ValidateSkuNumbers = True
            If SkuNumbers.Trim <> String.Empty Then
                If SkuNumbers.Trim.Length > 11 Then
                    'If we have more than one SkuNumbers. We look for the "," char in the string
                    If InStr(SkuNumbers, ",", CompareMethod.Text) = 0 Then
                        sMessage = "Sku Numbers: You have to validate the Sku Numbers. SKU contain 11 characters, and are separated by comma"
                        lblError.Text = sMessage
                        Return False
                    End If
                End If
                'More than one SPSNumbers
                If SkuNumbers.Split(",").Length > 1 Then
                    For Each elem As String In SkuNumbers.Split(",")
                        If ValidateSkuNumber(elem.Trim) = False Then
                            Return False
                        End If
                    Next
                Else
                    Return ValidateSkuNumber(SkuNumbers.Trim)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateAvNumberKMAT() As Boolean
        Try
            Dim dtData As New DataTable
            dtData = HPQ.Excalibur.Service.GetServiceAVNumbersKMAT(txtBomAvNumbers.Text)

            If dtData.Rows.Count > 0 Then
                'the number of rows OF Avs with same KMAT must 1
                Dim aAvNumbers As Array
                aAvNumbers = txtBomAvNumbers.Text.Split(",")
                If aAvNumbers.Length > 1 Then
                    If dtData.Rows.Count > 1 Then
                        Return False
                    Else
                        Return True 'All AV numbers share the same KMAT
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If

            dtData = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Function ValidateAvNumberBaseUnit() As Boolean
        Try
            Dim dtData As New DataTable

            dtData = HPQ.Excalibur.Service.GetServiceAVNumbersBaseUnitCategory(txtBomAvNumbers.Text)

            If dtData.Rows.Count > 0 Then
                'all the AV Numbers must be BAse Unit Category
                Dim aAvNumbers As Array
                aAvNumbers = txtBomAvNumbers.Text.Split(",")

                If dtData.Rows.Count = aAvNumbers.Length Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If

            dtData = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Protected Sub cusSKUNumberProduct_ServerValidate(ByVal sender As Object, ByVal e As ServerValidateEventArgs)
        Try
            lblError.Text = String.Empty

            Dim ProductSelected As Boolean = False

            For Each elem As ListItem In lstProducts.Items
                If elem.Selected = True Then
                    ProductSelected = True
                    Exit For
                End If
            Next

            If (ProductSelected = False And txtSkuNumberProduct.Text = String.Empty) OrElse ProductSelected = True And txtSkuNumberProduct.Text <> String.Empty Then
                e.IsValid = False
            Else
                e.IsValid = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    'Protected Sub TextValidateSKUNumberProduct(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtSkuNumberProduct.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '            rCustomValtxtSkuNumberProduct.ErrorMessage = sMessage
    '        Else
    '            args.IsValid = True
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Function ValidateFieldsReports() As Boolean
    '    Try
    '        Select Case ReportType
    '            Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS 'Validate Sku Numbers
    '                Return ValidateSkuNumbers(txtSKUNumber.Text)
    '        End Select
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    'Private Function ValidateFieldsReports() As Boolean
    '    Try
    '        Select Case ReportType
    '            Case REPORT_TYPE_SERVICE_SPS_BOM 'Validate Sparekit Numbers
    '                Return ValidateSpareKits(txtSPSBom.Text)
    '                'Case REPORT_TYPE_SPS_TO_PRODUCTS 'Validate Sparekit Numbers
    '                Return ValidateSpareKits(txtSpsNumber.Text)
    '            Case REPORT_TYPE_AV_TO_SPS 'Validate AV Numbers
    '                Return ValidateAvNumbers(txtAVNumber.Text)
    '            Case REPORT_TYPE_SERVICE_BOM_AVNUMBERS
    '                Return ValidateAvNumbers(txtBomAvNumbers.Text)
    '            Case REPORT_TYPE_SERVICE_SKU_TO_SPAREKITS 'Validate Sku Numbers
    '                Return ValidateSkuNumbers(txtSKUNumber.Text)
    '            Case REPORT_TYPE_SERVICE_PRODUCTVERSION_TO_SKU_TO_SPAREKITS
    '                Return ValidateSkuNumbers(txtSkuNumberProduct.Text)
    '        End Select
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    'Protected Sub TextValidateSpsNumber(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtSpsNumber.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '            rCustomValtxtSpsNumber.ErrorMessage = sMessage
    '        Else
    '            args.IsValid = True
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Protected Sub TextValidateSPSBom(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtSPSBom.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '            rCustomValtxtSPSBom.ErrorMessage = sMessage
    '        Else
    '            args.IsValid = True
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub


    'Protected Sub TextValidateAvNumber(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtAVNumber.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '            rCustomValtxtAVNumber.ErrorMessage = sMessage
    '        Else
    '            args.IsValid = True
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Protected Sub TextValidateSKUNumber(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtSKUNumber.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '            rCustomValtxtSKUNumber.ErrorMessage = sMessage
    '        Else
    '            args.IsValid = True
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Protected Sub TextValidateBomAvNumbers(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles rCustomValtxtBomAvNumbers.ServerValidate
    '    Try
    '        If ValidateFieldsReports() = False Then
    '            args.IsValid = False
    '        Else
    '            If Me.txtBomAvNumbers.Text <> String.Empty Then
    '                If ValidateAvNumberBaseUnit() = True Then
    '                    'If we have more than one AV number...
    '                    If txtBomAvNumbers.Text.Split(",").Length > 1 Then
    '                        'All av number myust have the same Kmat
    '                        If ValidateAvNumberKMAT() = True Then
    '                            args.IsValid = True
    '                        Else
    '                            sMessage = "Invalid AvNumbers list.The Av Numbers do not share the same KMAT."
    '                            args.IsValid = False
    '                        End If
    '                    Else
    '                        args.IsValid = True
    '                    End If
    '                Else
    '                    sMessage = "Invalid AvNumbers list. One or more of the Av Numbers in the list do not have a Base Unit Category."
    '                    args.IsValid = False
    '                End If
    '            End If
    '        End If

    '        rCustomValtxtBomAvNumbers.ErrorMessage = sMessage
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub




#End Region





End Class
