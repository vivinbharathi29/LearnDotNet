Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading

Partial Class Service_processSKUBOMAS
    Inherits System.Web.UI.Page

    Enum ReportType
        HTML = 0
        EXCEL = 1
        TEXT = 2
    End Enum

#Region " Properties "
    Private strRequestMethod As String = "A"
    Private Property RequestMethod As String
        Get
            Return strRequestMethod
        End Get
        Set(ByVal value As String)
            strRequestMethod = value
        End Set
    End Property


    Private enCurrentReportType As ReportType
    Private Property CurrentReportType As ReportType
        Get
            Return enCurrentReportType
        End Get
        Set(ByVal value As ReportType)
            enCurrentReportType = value
        End Set
    End Property

    Private strFileName As String = ""
    Private Property FileName As String
        Get
            Return strFileName
        End Get
        Set(ByVal value As String)
            strFileName = value
        End Set
    End Property

    Private strColumnHeaders As String = ""
    Private Property ColumnHeaders As String
        Get
            Return strColumnHeaders
        End Get
        Set(ByVal value As String)
            strColumnHeaders = value
        End Set
    End Property


    Private intUserId As Integer = GetUserID()
    Protected Property UserId As Integer
        Get
            Return intUserId
        End Get
        Set(ByVal value As Integer)
            intUserId = value
        End Set
    End Property

    Private strRequestGUID As String
    Protected Property RequestGUID As String
        Get
            Return strRequestGUID
        End Get
        Set(ByVal value As String)
            strRequestGUID = value
        End Set
    End Property

    Private intReturnCode As Integer = 0
    Protected Property ReturnCode As Integer
        Get
            Return intReturnCode
        End Get
        Set(ByVal value As Integer)
            intReturnCode = value
        End Set
    End Property

    Private strReturnDesc As String
    Protected Property ReturnDesc As String
        Get
            Return strReturnDesc
        End Get
        Set(ByVal value As String)
            strReturnDesc = value
        End Set
    End Property

    Private strProfileID As String
    Protected Property ProfileID As String
        Get
            Return strProfileID
        End Get
        Set(ByVal value As String)
            strProfileID = value
        End Set
    End Property
#End Region

    Private Function GetUserID() As Integer
        Dim intUserID As Integer = 0
        Dim objSec As HPQ.Excalibur.Security
        Dim strLOGON_USER As String = ""
        Dim strHTTPCTXT_USER As String = ""

        Try
            strLOGON_USER = Session("LoggedInUser")
            strHTTPCTXT_USER = Session("LoggedInUser")

            If (strHTTPCTXT_USER.Trim.Length > 0) Then
                objSec = New HPQ.Excalibur.Security(strHTTPCTXT_USER.Trim)
                intUserID = objSec.CurrentUserID
            ElseIf (strLOGON_USER.Trim.Length > 0) Then
                objSec = New HPQ.Excalibur.Security(strLOGON_USER.Trim)
                intUserID = objSec.CurrentUserID
            Else
                Me.ReturnCode = -1
                Me.ReturnDesc = "<b>ERROR DETERMINING USER VERIFICATION PARAMETERS</b>"
                intUserID = 0
            End If

        Catch ex As Exception
            Me.ReturnCode = -1
            Me.ReturnDesc = "<b>USER VERIFICATION ERROR</b> - " & ex.Message
            intUserID = 0
        Finally
            objSec = Nothing
        End Try

        GetUserID = intUserID

    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoServerCaching()

        ' (Request.Headers("PARAMETERS")) ' SAVE WITH IDENTIFIERS (USER_ID, PARAMETERS, SESSION, TIMESTAMP)...RETURN, THEN LAUNCH WINDOW FROM CLIENT SIDE HANDLER
        ' MAY WANT TO CONVERT THIS TO A SET OF Public Web Methods for a Web Service in the future
        ' PROCESS ACCORDING TO MODE PASSED AS QUERYSTRING VARIABLE (i.e. - 0=CreateUserRequest, 1=ProcessUserRequest, etc.)

        Dim strRM As String = Request.QueryString("rm")
        Dim strMode As String = Request.QueryString("m")
        Dim strParameters As String = ""
        Dim strProfileID As String = "0"

        If (Not strRM Is Nothing) Then
            Me.RequestMethod = strRM.Trim.ToUpper
        Else
            strRM = ""
        End If

        If (strMode.Trim().Length > 0) Then
            If (strMode = "0") Then
                '*****************************************************************************************************************
                ' CREATE REPORT REQUEST
                '*****************************************************************************************************************
                If (Me.RequestMethod = "A") Then
                    strParameters = Request.Headers("PARAMETERS").Trim
                    strParameters = Replace(strParameters, "'", "")
                    strProfileID = UCase(Request.Headers("PID").Trim)

                    If (strProfileID.IndexOf("S") > -1 Or strProfileID.IndexOf("G") > -1) Then
                        strProfileID = Replace(strProfileID, "S", "")
                        strProfileID = Replace(strProfileID, "G", "")
                    Else
                        strProfileID = "0"
                    End If

                    Me.ProfileID = strProfileID

                ElseIf (Me.RequestMethod = "F") Then
                    strParameters = Request.Form("PARAMETERS").Trim
                    strParameters = Replace(strParameters, "'", "")
                    strProfileID = UCase(Request.Form("PID").Trim)

                    If (strProfileID.IndexOf("S") > -1 Or strProfileID.IndexOf("G") > -1) Then
                        strProfileID = Replace(strProfileID, "S", "")
                        strProfileID = Replace(strProfileID, "G", "")
                    Else
                        strProfileID = "0"
                    End If

                    Me.ProfileID = strProfileID

                Else
                    Response.Write("<script type='text/javascript' language='javascript'>window.parent.processAlternateResponse(null, true, 'ERROR - Unknown request Method.');</script>")
                    Response.End()
                End If

                    If (strParameters.Length = 0) Then
                        If (Me.RequestMethod = "F") Then
                            Response.Write("<script type='text/javascript' language='javascript'>window.parent.processAlternateResponse(null, true, 'ERROR - INVALID CREDENTIALS');</script>")
                        Else
                            Response.Write("<b>ERROR - INVALID CREDENTIALS</b>")
                            Response.End()
                        End If
                    End If

                    If (CreateUserRequest(strParameters)) Then
                        If (Me.RequestMethod = "A") Then
                            Response.Write(Me.RequestGUID)
                            Response.End()
                        ElseIf (Me.RequestMethod = "F") Then
                            If (Me.ReturnCode = 0) Then
                                Response.Write("<script type='text/javascript' language='javascript'>window.parent.processAlternateResponse('" & Me.RequestGUID & "', false, null);</script>")
                                Response.End()
                            Else
                                Response.Write("<script type='text/javascript' language='javascript'>window.parent.processAlternateResponse(null, true, 'ERROR - " & Me.ReturnDesc & "');</script>")
                                Response.End()
                            End If
                        End If
                    Else
                        If (Me.RequestMethod = "F") Then
                            Response.Write("<script type='text/javascript' language='javascript'>window.parent.processAlternateResponse(null, true, 'ERROR - " & Me.ReturnDesc & "');</script>")
                        Else
                            Response.Write("ERROR - " & Me.ReturnDesc)
                            Response.End()
                        End If
                    End If
                    '*****************************************************************************************************************

            ElseIf (strMode = "1") Then
                '*****************************************************************************************************************
                ' GENERATE REPORT
                '*****************************************************************************************************************
                Dim strRG As String = Request.QueryString("rg")
                Dim strParams As String
                Dim strFN As String = Request.QueryString("fn")

                strProfileID = UCase(Request.QueryString("PID").Trim)

                If (strProfileID.IndexOf("S") > -1 Or strProfileID.IndexOf("G") > -1) Then
                    strProfileID = Replace(strProfileID, "S", "")
                    strProfileID = Replace(strProfileID, "G", "")
                Else
                    strProfileID = "0"
                End If

                Me.ProfileID = strProfileID

                If (strRG.Trim.Length > 0) Then

                    ' DETERMINE THE ReportType and Process accordingly ---> Put this into property when calling GetParameters
                    Me.RequestGUID = strRG.Trim.ToUpper
                    strParams = GetParameters(strRG)

                    If (Me.ReturnCode <> 0) Then
                        DisplayMessage("<b>ERROR</b> - " & Me.ReturnDesc, "lblStatus", Drawing.Color.Red)
                        UpdateUserRequest(Me.RequestGUID)
                    Else

                        If (Not strFN Is Nothing) Then
                            Me.FileName = strFN.Trim ' NEED TO ADD VALIDATION OF FILENAME (invalid chars etc.)
                        Else
                            Me.FileName = ""
                        End If

                        If Not GenerateReport(strParams.Trim) Then
                            If (Me.ReturnCode <> 0) Then
                                DisplayMessage("<b>ERROR</b> - " & Me.ReturnDesc, "lblStatus", Drawing.Color.Red)
                            Else
                                DisplayMessage(Me.ReturnDesc, "lblStatus", Drawing.Color.Black)
                            End If
                        Else
                            ' Update the User Request Record, consider making a function...test for thread term errors
                            'Me.ReturnCode = 0
                            'Me.ReturnDesc = "Successfully retrieved records using the specified filter."
                            'UpdateUserRequest(Me.RequestGUID)
                        End If


                    End If
                Else
                    Response.Write("ERROR - NO REQUEST IDENTIFIER SPECIFIED")
                    Response.End()
                End If
                '*****************************************************************************************************************

            ElseIf (strMode = "2") Then
                '*****************************************************************************************************************
                ' 
                '*****************************************************************************************************************

                '*****************************************************************************************************************
            Else
                Response.Write("ERROR - Unknown calling Mode.")
                Response.End()
            End If

        Else
            Response.Write("ERROR - Calling Mode not specified.")
            Response.End()
        End If

    End Sub

    Protected Function CreateUserRequest(ByVal strParameters As String) As Boolean

        ' Require the user to specify at least 1 filter parameter (may require more or specific ones in the future, e.g. - dates)
        If (NumParameters(strParameters) = 0) Then

            If (Me.ReturnCode = 0) Then
                Me.ReturnCode = -1
                Me.ReturnDesc = "Not enough filter parameters specified.  Please specify as many filter parameters as possible in order to limit the scope of the results."
            End If

            CreateUserRequest = False
            Exit Function
        End If

        Dim blnResult As Boolean = False
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim aryParameters() As String = strParameters.Split("|")
        Dim strProfileID As String = "0"



        Try

            ' Validate user requesting this action, disregard...this will occur later

            objDW = New HPQ.Data.DataWrapper()

            objComm = objDW.CreateCommand("usp_InsertBTOSSUSERREQUEST", Data.CommandType.StoredProcedure)

            objDW.CreateParameter(objComm, "@USER_ID", Data.SqlDbType.Int, GetUserID().ToString, 8, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@SESSION_ID", Data.SqlDbType.VarChar, Session.SessionID, Session.SessionID.Length)
            objDW.CreateParameter(objComm, "@PARAMETERS", Data.SqlDbType.VarChar, strParameters, strParameters.Length)
            objDW.CreateParameter(objComm, "@PROFILE_ID", Data.SqlDbType.Int, Me.ProfileID, 8, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@REQUEST_GUID", Data.SqlDbType.UniqueIdentifier, "", 36, ParameterDirection.InputOutput)
            objDW.CreateParameter(objComm, "@RETURN_CODE", Data.SqlDbType.Int, "0", 8, ParameterDirection.InputOutput)
            objDW.CreateParameter(objComm, "@RETURN_DESC", Data.SqlDbType.VarChar, "", 255, ParameterDirection.InputOutput)

            objDW.ExecuteCommandNonQuery(objComm)

            Dim intRC As Integer

            If Integer.TryParse(objComm.Parameters("@RETURN_CODE").Value.ToString(), intRC) Then
                Me.ReturnCode = intRC
                Me.ReturnDesc = objComm.Parameters("@RETURN_DESC").Value.ToString()

                If (intRC = 0) Then
                    blnResult = True
                    Me.RequestGUID = objComm.Parameters("@REQUEST_GUID").Value.ToString()
                End If
            Else
                Me.ReturnDesc = "Unknown failure."
            End If

        Catch ex As Exception
            Me.ReturnCode = -1
            Me.ReturnDesc = "<b>CreateUserRequest</b> - " & ex.Message

        Finally
            objComm = Nothing
            objDW = Nothing
        End Try

        CreateUserRequest = blnResult

    End Function


    Protected Sub UpdateUserRequest(ByVal strRG As String)
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand

        Try
            objDW = New HPQ.Data.DataWrapper()

            objComm = objDW.CreateCommand("usp_UpdateBTOSSUSERREQUEST", Data.CommandType.StoredProcedure)

            objDW.CreateParameter(objComm, "@USER_ID", Data.SqlDbType.Int, Me.UserId.ToString, 8, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@SESSION_ID", Data.SqlDbType.VarChar, Session.SessionID, Session.SessionID.Length)
            objDW.CreateParameter(objComm, "@REQUEST_GUID", Data.SqlDbType.UniqueIdentifier, strRG, 36, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@RETURN_CODE", Data.SqlDbType.VarChar, Me.ReturnCode.ToString, 10, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@RETURN_DESC", Data.SqlDbType.VarChar, Me.ReturnDesc, 255, ParameterDirection.Input)

            objDW.ExecuteCommandNonQuery(objComm)

        Catch ex As Exception
            ' Ignore
        Finally
            objComm = Nothing
            objDW = Nothing
        End Try

    End Sub


    Protected Function GetParameters(ByVal strRG As String) As String
        Dim strParameters As String = ""

        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim objRow As DataRow

        Try

            ' Validate user requesting this action, disregard...this will occur later

            objDW = New HPQ.Data.DataWrapper()

            objComm = objDW.CreateCommand("usp_SelectBTOSSUSERREQUEST", Data.CommandType.StoredProcedure)

            objDW.CreateParameter(objComm, "@USER_ID", Data.SqlDbType.Int, GetUserID().ToString, 8, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@SESSION_ID", Data.SqlDbType.VarChar, Session.SessionID, Session.SessionID.Length)
            objDW.CreateParameter(objComm, "@REQUEST_GUID", Data.SqlDbType.UniqueIdentifier, strRG, 36, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@RETURN_CODE", Data.SqlDbType.Int, "0", 8, ParameterDirection.InputOutput)
            objDW.CreateParameter(objComm, "@RETURN_DESC", Data.SqlDbType.VarChar, "", 255, ParameterDirection.InputOutput)

            objDT = objDW.ExecuteCommandTable(objComm)

            Dim intRC As Integer

            If Integer.TryParse(objComm.Parameters("@RETURN_CODE").Value.ToString(), intRC) Then
                Me.ReturnCode = intRC
                Me.ReturnDesc = objComm.Parameters("@RETURN_DESC").Value.ToString()

                If (intRC = 0) Then
                    If (objDT.Rows.Count > 0) Then
                        objRow = objDT.Rows(0)
                        strParameters = objRow(0).ToString()
                    End If

                End If
            Else
                Me.ReturnDesc = "Unknown failure."
            End If

        Catch ex As Exception
            Me.ReturnCode = -1
            Me.ReturnDesc = "<b>GetParameters</b> - " & ex.Message

        Finally
            objDT = Nothing
            objComm = Nothing
            objDW = Nothing
        End Try

        GetParameters = strParameters

    End Function

    Private Function GenerateReport(ByVal strParameters As String) As Boolean
        Dim blnResult As Boolean = False
        Dim objDT As DataTable
        Dim strFileName As String = ""

        ' Check for the CurrentReportType and process accordingly, may modulate into 3 or more called routines
        Me.UserId = GetUserID()

        Try

            objDT = New DataTable ' may remove this
            objDT = GetDataTable(strParameters)

            If ((Not objDT Is Nothing) And (Me.ReturnCode = 0)) Then
                Me.ReturnCode = 0
                Me.ReturnDesc = "Successfully retrieved records using the specified filter."

                Select Case Me.CurrentReportType

                    Case ReportType.HTML
                        '***********************************************************************************************************
                        ' HTML
                        '***********************************************************************************************************
                        Response.Clear()
                        Response.Charset = "" ' ISO or the like
                        Response.Cache.SetCacheability(HttpCacheability.NoCache)
                        Response.ContentType = "text/html"

                        Dim objGV As GridView = Me.FindControl("grdVwRpt")

                        objGV.DataSource = objDT.DefaultView.ToTable
                        objGV.DataBind()

                        DisplayMessage("<b>Total Record(s):</b> " & objDT.Rows.Count.ToString(), "lblStatus", Drawing.Color.Black)

                        UpdateUserRequest(Me.RequestGUID)
                        '***********************************************************************************************************

                    Case ReportType.EXCEL
                        '***********************************************************************************************************
                        ' EXCEL
                        '***********************************************************************************************************

                        If (Me.FileName.Trim.Length = 0) Then
                            strFileName = "FileName.xls"
                        Else
                            strFileName = Me.FileName
                        End If

                        Response.Clear()
                        Response.Buffer = True
                        Response.ContentType = "application/vnd.ms-excel"
                        Response.AddHeader("content-disposition", "attachment;filename=" & strFileName)
                        Response.Charset = "" ' ISO or the like
                        'Response.Cache.SetCacheability(HttpCacheability.Private)

                        Dim objStrWriter As System.IO.StringWriter = New System.IO.StringWriter
                        Dim objHTMLWriter As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(objStrWriter)

                        Dim objForm As HtmlForm = New HtmlForm

                        Dim objGV As GridView = New GridView
                        objForm.Controls.Add(objGV)


                        objGV.DataSource = objDT.DefaultView.ToTable
                        objGV.DataBind()

                        objGV.RenderControl(objHTMLWriter)
                        Response.Write(objStrWriter.ToString)
                        Response.Write("<span id='lblHPC' style='color:Red;font-family:verdana;font-size:X-Small;font-weight:bold;'><br>HP Restricted</span><br />")

                        objStrWriter = Nothing
                        objHTMLWriter = Nothing
                        objGV = Nothing
                        objForm = Nothing

                        UpdateUserRequest(Me.RequestGUID)
                        Response.Flush()
                        Response.End()
                        '***********************************************************************************************************

                    Case ReportType.TEXT
                        '***********************************************************************************************************
                        ' TEXT (Pipe Delimited)
                        '***********************************************************************************************************
                        Dim objRow As DataRow

                        If (Me.FileName.Trim.Length = 0) Then
                            strFileName = "FileName.txt"
                        Else
                            strFileName = Me.FileName
                        End If

                        Response.Clear()
                        Response.ContentType = "text/plain"
                        Response.AddHeader("content-disposition", "attachment;filename=" & strFileName)
                        Response.Charset = "" ' ISO or the like
                        'Response.Cache.SetCacheability(HttpCacheability.NoCache)

                        ' Eventually allow User Defined or more than one delimiter
                        ' May also allow qualifiers
                        ' May also make Header Row optional


                        Response.Write(Me.ColumnHeaders & vbCrLf)

                        For Each objRow In objDT.Rows
                            Response.Write(objRow(0) & vbCrLf)
                        Next

                        Response.Write("HP RESTRICTED" & vbCrLf)

                        UpdateUserRequest(Me.RequestGUID)

                        Response.End()

                        '***********************************************************************************************************
                End Select

                blnResult = True
            Else
                If ((objDT Is Nothing) And (Me.ReturnCode = 0)) Then
                    UpdateUserRequest(Me.RequestGUID)
                ElseIf (Me.ReturnCode <> 0) Then
                    UpdateUserRequest(Me.RequestGUID)
                End If
            End If

        Catch ex As Exception

            If TypeOf ex Is StackOverflowException Or TypeOf ex Is OutOfMemoryException Then
                Me.ReturnCode = -1

                If Me.CurrentReportType = ReportType.HTML Or Me.CurrentReportType = ReportType.TEXT Then
                    Me.ReturnDesc = "<b>Server Memory Error.  Please try again OR adjust the filter options in order to narrow the scope of the results.</b>"
                ElseIf Me.CurrentReportType = ReportType.EXCEL Then
                    Me.ReturnDesc = "<b>Server Memory Error.  Please try again OR adjust the filter options in order to narrow the scope of the results.</b><br/><b>You may want to consider exporting this data using the TEXT Report Type.</b>"
                End If

            ElseIf Not TypeOf ex Is ThreadAbortException Then

                Me.ReturnCode = -1
                Me.ReturnDesc = "<b>" & ex.Message & "</b>"
                'Thread.ResetAbort()

            End If

            UpdateUserRequest(Me.RequestGUID)

        Finally
            objDT = Nothing
        End Try

        GenerateReport = blnResult

    End Function


    Function GetDataTable(ByVal strParameters As String) As DataTable

        ' Require the user to specify at least 1 filter parameter (may require more or specific ones in the future, e.g. - dates)
        If (strParameters.Length = 0) Then
            Me.ReturnCode = -1
            Me.ReturnDesc = "<b>Not enough filter parameters specified.</b>"
            GetDataTable = Nothing
            Exit Function
        End If

        ' Retrieve the results
        Dim objDT As DataTable = Nothing
        Dim objDW As HPQ.Data.DataWrapper = Nothing
        Dim objComm As Data.SqlClient.SqlCommand
        Dim aryParameters() As String = strParameters.Split("|")
        Dim blnHasRows As Boolean = True
        Dim blnReturnPSV As Boolean = False
        Dim blnSPBLogFilter As Boolean = False
        Dim blnSKUAVLogFilter As Boolean = False
        Dim strProfileID As String = "0"


        Try

            If (aryParameters.Length >= 27) Then

                Select Case aryParameters(26)
                    Case "HTML" : Me.CurrentReportType = ReportType.HTML ' FORGOT WAY TO DO THIS
                    Case "EXCEL" : Me.CurrentReportType = ReportType.EXCEL ' FORGOT WAY TO DO THIS
                    Case "TEXT" : Me.CurrentReportType = ReportType.TEXT ' FORGOT WAY TO DO THIS
                        blnReturnPSV = True
                    Case Else : Me.CurrentReportType = ReportType.HTML
                End Select

                If (aryParameters.Length >= 31) Then
                    If (Boolean.TryParse(aryParameters(31), blnSPBLogFilter)) Then

                    End If
                End If

                If (aryParameters.Length >= 32) Then
                    If (Boolean.TryParse(aryParameters(32), blnSKUAVLogFilter)) Then

                    End If
                End If
            End If

            ' Validate user requesting this action

            objDW = New HPQ.Data.DataWrapper()

            objComm = objDW.CreateCommand("usp_BTOSSAdvancedSearch", Data.CommandType.StoredProcedure)

            objComm.CommandTimeout = 120

            objDW.CreateParameter(objComm, "@ColumnIDs", Data.SqlDbType.VarChar, aryParameters(0), 500)
            objDW.CreateParameter(objComm, "@ColumnOrderByIDs", Data.SqlDbType.VarChar, aryParameters(1), 500)
            objDW.CreateParameter(objComm, "@ColumnOrderAscDesc", Data.SqlDbType.VarChar, aryParameters(2), 500)

            objDW.CreateParameter(objComm, "@ServiceGeoNA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(3)))
            objDW.CreateParameter(objComm, "@ServiceGeoLA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(4)))
            objDW.CreateParameter(objComm, "@ServiceGeoAPJ", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(5)))
            objDW.CreateParameter(objComm, "@ServiceGeoEMEA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(6)))

            objDW.CreateParameter(objComm, "@SKGeoNA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(7)))
            objDW.CreateParameter(objComm, "@SKGeoLA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(8)))
            objDW.CreateParameter(objComm, "@SKGeoAPJ", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(9)))
            objDW.CreateParameter(objComm, "@SKGeoEMEA", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(10)))

            objDW.CreateParameter(objComm, "@ProductBrandIDs", Data.SqlDbType.VarChar, aryParameters(11), 8000)
            objDW.CreateParameter(objComm, "@ServiceCategoryIDs", Data.SqlDbType.VarChar, aryParameters(12), 8000)
            objDW.CreateParameter(objComm, "@OSSPIDs", Data.SqlDbType.VarChar, aryParameters(13), 8000)
            objDW.CreateParameter(objComm, "@KMATs", Data.SqlDbType.VarChar, aryParameters(14), 8000)
            objDW.CreateParameter(objComm, "@SKUs", Data.SqlDbType.VarChar, aryParameters(15), 8000)
            objDW.CreateParameter(objComm, "@AVs", Data.SqlDbType.VarChar, aryParameters(16), 4000)
            objDW.CreateParameter(objComm, "@RequireAllAVs", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(17)))
            objDW.CreateParameter(objComm, "@SKs", Data.SqlDbType.VarChar, aryParameters(18), 8000)
            objDW.CreateParameter(objComm, "@SFPNS", Data.SqlDbType.VarChar, aryParameters(19), 8000)

            objDW.CreateParameter(objComm, "@LastAction", Data.SqlDbType.VarChar, aryParameters(20), 5)
            objDW.CreateParameter(objComm, "@ActionDateFrom", Data.SqlDbType.VarChar, aryParameters(21), 20)
            objDW.CreateParameter(objComm, "@ActionDateTo", Data.SqlDbType.VarChar, aryParameters(22), 20)
            objDW.CreateParameter(objComm, "@ActionDateColumnID", Data.SqlDbType.VarChar, aryParameters(27), 5)

            objDW.CreateParameter(objComm, "@SKUAVs", Data.SqlDbType.VarChar, aryParameters(24), 4000)
            objDW.CreateParameter(objComm, "@RequireAllSKUAVs", Data.SqlDbType.Bit, Boolean.Parse(aryParameters(25)))

            objDW.CreateParameter(objComm, "@SAs", Data.SqlDbType.VarChar, aryParameters(28), 8000)
            objDW.CreateParameter(objComm, "@COMPs", Data.SqlDbType.VarChar, aryParameters(29), 8000)

            objDW.CreateParameter(objComm, "@PROD_DIV", Data.SqlDbType.TinyInt, aryParameters(30), 8)

            objDW.CreateParameter(objComm, "@ReturnXML", Data.SqlDbType.Bit, False)
            objDW.CreateParameter(objComm, "@ReturnPSV", Data.SqlDbType.Bit, blnReturnPSV)

            objDW.CreateParameter(objComm, "@USER_ID", Data.SqlDbType.Int, GetUserID().ToString, 8, ParameterDirection.Input)
            objDW.CreateParameter(objComm, "@SESSION_ID", Data.SqlDbType.VarChar, Session.SessionID, 100)
            objDW.CreateParameter(objComm, "@REQUEST_GUID", Data.SqlDbType.UniqueIdentifier, Me.RequestGUID.ToUpper, 36)

            objDW.CreateParameter(objComm, "@PROFILE_ID", Data.SqlDbType.Int, Me.ProfileID, 8, ParameterDirection.Input)

            objDW.CreateParameter(objComm, "@SPBLogFilter", Data.SqlDbType.Bit, blnSPBLogFilter)

            objDW.CreateParameter(objComm, "@SKUAVLogFilter", Data.SqlDbType.Bit, blnSKUAVLogFilter)

            objDW.CreateParameter(objComm, "@COLUMN_HEADERS", Data.SqlDbType.VarChar, "", 8000, ParameterDirection.InputOutput)

            objDW.CreateParameter(objComm, "@RETURN_CODE", Data.SqlDbType.Int, "0", 8, ParameterDirection.InputOutput)
            objDW.CreateParameter(objComm, "@RETURN_DESC", Data.SqlDbType.VarChar, "", 255, ParameterDirection.InputOutput)

            objDT = objDW.ExecuteCommandTable(objComm)

            Dim intRC As Integer

            If Integer.TryParse(objComm.Parameters("@RETURN_CODE").Value.ToString(), intRC) Then

                Me.ReturnCode = intRC

                If (objDT.Rows.Count > 0) Then
                    Me.ReturnDesc = objComm.Parameters("@RETURN_DESC").Value.ToString()

                    If Me.CurrentReportType = ReportType.TEXT Then
                        Me.ColumnHeaders = objComm.Parameters("@COLUMN_HEADERS").Value.ToString()
                    End If

                ElseIf (Me.ReturnCode = 0) Then
                    blnHasRows = False
                    Me.ReturnDesc = "<b>No Records returned.</b>"
                Else
                    Me.ReturnDesc = objComm.Parameters("@RETURN_DESC").Value.ToString()
                End If

            Else
                Me.ReturnCode = -1
                Me.ReturnDesc = objComm.Parameters("@RETURN_DESC").Value.ToString()
            End If

        Catch ex As Exception

            Me.ReturnCode = -1

            Try ' SqlException Cast
                Dim objSQLExc As System.Data.SqlClient.SqlException

                objSQLExc = TryCast(ex, System.Data.SqlClient.SqlException)

                ' Trap for specific error(s) and create user friendly version of message
                If objSQLExc.Number = -2 Then
                    ' Connection Timeout expired
                    Me.ReturnCode = -2
                    Me.ReturnDesc = "<b>Query timed out.  Please try again OR consider adjusting the filter options in order to narrow the scope of the results.</b>"
                Else
                    Me.ReturnDesc = "<b>GetDataTable</b> - " & ex.Message
                End If

                UpdateUserRequest(Me.RequestGUID)

            Catch ex2 As Exception
                Me.ReturnDesc = "<b>GetDataTable</b> - " & ex.Message

                UpdateUserRequest(Me.RequestGUID)
            End Try

        Finally
            If Not blnHasRows Then
                objDT = Nothing
            End If
            objComm = Nothing
            objDW = Nothing
        End Try

        GetDataTable = objDT

    End Function


    Function NumParameters(ByVal strParameters As String) As Integer
        Dim aryParameters() As String = strParameters.Split("|")
        Dim intTotalParameters As Integer = 0
        Dim blnChecked As Boolean = False

        Try

            If (aryParameters.Length >= 27) Then

                'GeoNA
                If (Boolean.TryParse(aryParameters(3), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'GeoLA
                If (Boolean.TryParse(aryParameters(4), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'GeoAPJ
                If (Boolean.TryParse(aryParameters(5), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'GeoEMEA
                If (Boolean.TryParse(aryParameters(6), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'SKGeoNA
                If (Boolean.TryParse(aryParameters(7), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'SKGeoLA
                If (Boolean.TryParse(aryParameters(8), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'SKGeoAPJ
                If (Boolean.TryParse(aryParameters(9), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'SKGeoEMEA
                If (Boolean.TryParse(aryParameters(10), blnChecked)) Then
                    If blnChecked Then intTotalParameters = intTotalParameters + 1
                End If

                'Product Brands
                If (aryParameters(11).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'Service Categories
                If (aryParameters(12).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'OSSPs
                If (aryParameters(13).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'KMATS
                If (aryParameters(14).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'SKUS
                If (aryParameters(15).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'AV Qualifiers
                If (aryParameters(16).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'SKs
                If (aryParameters(18).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'SFPNs
                If (aryParameters(19).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'Action Date Range
                If (aryParameters(21).Length > 0) And (aryParameters(22).Length > 0) Then intTotalParameters = intTotalParameters + 1

                'SKU AVs
                If (aryParameters(24).Length > 0) Then intTotalParameters = intTotalParameters + 1

                If (aryParameters(28).Length > 0) Then intTotalParameters = intTotalParameters + 1

                If (aryParameters(29).Length > 0) Then intTotalParameters = intTotalParameters + 1
                ' Ignore all other filter options as they are not granular enough

            ElseIf (aryParameters.Length > 0) And (aryParameters.Length < 27) Then
                Me.ReturnCode = -1
                Me.ReturnDesc = "Invalid/Corrupt Parameters specified."
            End If

        Catch ex As Exception
            Me.ReturnCode = -1
            Me.ReturnDesc = ex.Message
        End Try

        NumParameters = intTotalParameters

    End Function

    Private Sub DisplayMessage(ByVal strMsg As String, ByVal strControlName As String, ByVal objColor As System.Drawing.Color)
        Dim objStatusLabel As Label = DirectCast(Me.FindControl(strControlName), Label)

        Try
            objStatusLabel.ForeColor = objColor
            objStatusLabel.Text = strMsg
        Catch ex As Exception
            ' Do Nothing
        End Try

    End Sub

End Class
