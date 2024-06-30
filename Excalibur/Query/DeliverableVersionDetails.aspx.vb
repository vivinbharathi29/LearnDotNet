Imports System.Data
Imports System.Data.SqlClient

Partial Class Query_DeliverableVersionDetails
    Inherits System.Web.UI.Page

    Const PARTNERID_HP As Integer = 1
    Const OSID_INDEPENDENT As Integer = 16
    Const LANGID_INDEPENDENT As Integer = 58
    Const ADVANCED_SEARCH_FUNCTION As Integer = 8  ' This is our function number (see cmdReport_onclick function in Deliverables.asp)

    Public Const MAXCOUNT_DELIVERABLES_REPORT = 100

    Public Enum DeliverableTypes
        Hardware = 1
        Software = 2
        Firmware = 3
        Documentation = 4
    End Enum

    Public Enum ContentFormats
        HTML = 0
        Excel = 1
        Word = 2
    End Enum

    Dim _Conn As SqlClient.SqlConnection
    Dim _RequestedVersionIDs(0) As Long
    Dim _RequestedDOTSNames(0) As String
    Dim _ContentFormat As Integer

    Public Structure CurrentUserInfo
        Public ID As Long
        Public PartnerID As Long
        Public Name As String
        Public Email As String
        Public Username As String
        Public Domain As String

        Public IsAdmin As Boolean
    End Structure

    Public Structure MilestoneRecord
        Public Name As String
        Public Status As String
        Public PlannedDate As String
        Public ActualDate As String
    End Structure

    Public Structure OTSRecord
        Public Number As String
        Public ShortDescription As String
        Public HTMLLink As String
    End Structure

    Public Structure DeliverableVersionRecord
        Public ID As Long
        Public Name As String
        Public Version As String
        Public Revision As String
        Public Pass As String
    End Structure

    Public Structure ProductRecord
        Public ID As Long
        Public ProductDeliverableReleaseID As Long
        Public Name As String
        Public ProjectManager As String
        Public DeveloperStatus As String
        Public Status As String
        Public TestingSummary As String
        Public PINOrTestNotes As String
        Public PilotStatus As String
        Public Restrictions As String
    End Structure

    Public Structure VersionProperties4Web
        Public RootID As Long
        Public VersionID As Long
        Public ErrMsg As String
        Public VersionIDHTML As String

        Public TypeID As Long
        Public LevelID As Long
        Public FullName As String           ' Constructed from Name, Version, Revision, and Pass
        Public Name As String
        Public Version As String
        Public Revision As String
        Public Pass As String
        Public Filename As String
        Public CodeName As String
        Public DeveloperID As Long
        Public Developer As String
        Public DevManager As String
        Public BuildLevel As String
        Public HFCN As Boolean
        Public CategoryID As Long

        Public Vendor As String
        Public VendorID As Long
        Public VendorVersion As String      ' The vendor version of this version of the deliverable
        Public VersionVendor As String      ' Vendor name of this version of the deliverable
        Public VersionVendorID As Long

        Public Supplier As String
        Public SupplierID As String

        Public Multilangauge As Boolean
        Public DeliverableSpec As String
        Public ImagePath As String
        Public Comments As String
        Public Changes As String
        Public DOTSName As String

        Public MilestoneList() As MilestoneRecord

        ' Samples Available
        Public SampleDate As String
        Public SampleConfidence As Long

        ' Intro Date/Mass Production
        Public IntroDate As String
        Public IntroConfidence As Long

        ' Available Until
        Public EOLDate As String
        Public Active As Boolean

        ' Special Notes
        Public InstallableUpdate As Boolean
        Public PackageForWeb As Boolean
        Public SpecialNotes As String

        ' Icons Installed
        Public IconDesktop As Boolean
        Public IconMenu As Boolean
        Public IconTray As Boolean
        Public IconPanel As Boolean
        Public IconsInstalled As String

        ' Property Tabs
        Public PropertyTabs As String

        ' Packaging
        Public Preinstall As Boolean
        Public FloppyDisk As Boolean
        Public Scriptpaq As Boolean
        Public CDImage As Boolean
        Public ISOImage As Boolean
        Public AR As Boolean
        Public Packaging As String

        ' ROM Components
        Public Binary As Boolean
        Public Rompaq As Boolean
        Public PreinstallROM As Boolean
        Public CAB As Boolean
        Public ROMComponents As String

        ' Replicated By
        Public Replicater As String

        ' CD Part Number / Kit Number
        Public CDPartNumber As String
        Public CDKitNumber As String

        ' Selected Languages
        Public SelectedLangIDs As String
        Public SelectedLangs As String
        Public PartNumbers As String
        Public CDKitNumbers As String

        ' Selected OSes
        Public SelectedOSIDs As String
        Public SelectedOSes As String

        ' RoHS/Green Spec
        Public RohsID As Long
        Public GreenSpecID As Long
        Public RoHSGreenSpec As String

        ' OTS
        Public OTSList() As OTSRecord

        ' PNP Devices
        Public PNPDevices() As String

        ' Deliverables Dependencies
        Public SelectedDelDependencies() As DeliverableVersionRecord
        Public RootDelDependencies() As DeliverableVersionRecord
        Public DeliverableDependencies As String
        Public SWDependencies As String


        ' Products
        Public ProductList() As ProductRecord
        Public ProductListHasTestingSummary As Boolean

        Public WorkflowCompleted As Boolean
        Public WorkflowLocation As String

        Public Certification As String
        Public CertificationStatus As String
        Public CertificationDate As String
        Public CertificationID As Long
        Public CertificationComments As String

        Public CATStatus As String
        Public CategoryName As String
        Public Spec As String
        Public ModelNumber As String
        Public PartNumber As String
    End Structure

    Dim _CurrentUser As CurrentUserInfo
    Dim _VP4Web(0) As VersionProperties4Web
    Dim _Idx As Integer = 0 ' Active index to array of VersionProperties4Web (_VP4Web)
    Dim _AdvancedSearch As Boolean
    Dim _ReportQueryCount As Integer

    Sub Page_Load(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.CacheControl = "No-cache"

        Title = "Deliverable Version Details"
        If Request.HttpMethod = "POST" And Request("txtFunction") = ADVANCED_SEARCH_FUNCTION Then
            _AdvancedSearch = True
            _ContentFormat = Request("cboFormat")
        End If

        If Not _AdvancedSearch Then
            _RequestedVersionIDs(0) = 0
            If IsNumeric(Request("ID")) Then
                _RequestedVersionIDs(0) = Request("ID")
                If (_RequestedVersionIDs(0) = 0) Then
                    _VP4Web(0).ErrMsg = "No deliverable version is specified"
                    Exit Sub
                End If
            Else
                _VP4Web(0).ErrMsg = "ERROR: Invalid deliverable version specified; deliverable version must be numeric."
            End If

            ' For a single deliverable request, go ahead and get the data
            GetDeliverables()
        Else
            If ContentFormat = ContentFormats.Word Then
                Response.ContentType = "application/msword"
            ElseIf ContentFormat = ContentFormats.Excel Then
                Response.ContentType = "application/vnd.ms-excel"
            End If

            Title = Request("txtTitle")
        End If
    End Sub

    Private Sub UpdateLoadingMessage(msg As String, msgId As Integer)
        If ContentFormat = ContentFormats.HTML Then
            Response.Write("<script type=""text/javascript"">" & vbCrLf)
            Response.Write("<!--script" & vbCrLf)
            Response.Write("updateLoadingMsg(""" & msg & """, " & msgId & ")" & vbCrLf)
            Response.Write("-->" & vbCrLf)
            Response.Write("</script>" & vbCrLf)
            Response.Flush()
        End If
    End Sub

    Protected Sub GetDeliverables()
        If OpenConnection() Then
            GetCurrentUserInfo()
            If _AdvancedSearch Then GetVerionIDsForReport()

            If ReportAvailableDeliverablesCount > MAXCOUNT_DELIVERABLES_REPORT Then
                UpdateLoadingMessage("NOTE: Found " & ReportAvailableDeliverablesCount & " deliverables but only displaying " & _
                                     "maximum of " & MAXCOUNT_DELIVERABLES_REPORT & " deliverables.<br /><br />", 2)
            End If

            Dim i As Integer = 0
            For Each verID As Long In _RequestedVersionIDs
                If verID <> 0 Then
                    ReDim Preserve _VP4Web(i)
                    _VP4Web(i).VersionID = verID
                    _VP4Web(i).VersionIDHTML = verID
                    _VP4Web(i).RootID = RootIDFromVersionID(_VP4Web(i).VersionID)
                    For Each DOTSName As String In _RequestedDOTSNames
                        _VP4Web(i).DOTSName = DOTSName
                    Next
                    If _AdvancedSearch Then UpdateLoadingMessage("Loading " & (i + 1) & " of " & _
                        _RequestedVersionIDs.Length & " deliverable(s)...", 1)
                    If Not GetVersionProperties4Web() Then
                        _VP4Web(i).ErrMsg = "Unable to find the requested deliverable version ID = " & _VP4Web(i).VersionID
                    End If
                    i = i + 1
                End If
            Next

            If _AdvancedSearch Then
                UpdateLoadingMessage("Displaying deliverables...", 1)
                If i = 0 And _VP4Web(0).ErrMsg = "" Then
                    _VP4Web(0).ErrMsg = "No deliverables found matching the specified criteria!"
                End If
            End If

            CloseConnection()
        End If

        If _VP4Web.Length = 1 And Not _AdvancedSearch And _VP4Web(0).FullName <> "" Then _
            Title = Title & ": " & _VP4Web(0).FullName
        _Idx = 0    ' Reset to 1st deliverable
    End Sub

    Private Function OpenConnection() As Boolean
        OpenConnection = False
        Try
            _Conn = New SqlConnection
            _Conn.ConnectionString = ConfigurationManager.ConnectionStrings("PRSConnectionString").ConnectionString
            _Conn.Open()

            If _Conn.State = ConnectionState.Open Then OpenConnection = True

        Catch ex As Exception
            'Maybe for debugging only: ex.Message 'DP: maybe no, since the error message was useless
            _VP4Web(0).ErrMsg = ex.Message
        End Try
    End Function

    Private Sub CloseConnection()
        If Not _Conn Is Nothing Then
            _Conn.Close()
            _Conn.Dispose()
            _Conn = Nothing
        End If
    End Sub

    Private Function GetReportProductList() As String
        Dim strProducts As String = Request("lstProducts"), strProductsPulsar As String = Request("lstProductsPulsar"), strGroupSQl As String = ""
        Dim productReleases As String()
        Dim productReleaseId As String()
        Dim productPulsarIds As String = ""
        Dim productReleaseIds As String = ""


        If strProductsPulsar IsNot Nothing And strProductsPulsar IsNot "" Then
            productReleases = Request("lstProductsPulsar").Split(New Char() {","})
            For Each productRelease As String In productReleases
                If InStr(productRelease, ":") > 0 Then
                    productReleaseId = productRelease.Split(New Char() {":"})
                    productPulsarIds = productPulsarIds + productReleaseId(0) + ","
                    productReleaseIds = productReleaseIds + productReleaseId(1) + ","
                End If
            Next
        End If

        If strProducts <> "" Then
            strProducts = strProducts.TrimEnd(",", "")
        End If
        If productPulsarIds <> "" Then
            productPulsarIds = productPulsarIds.TrimEnd(",", "")
        End If
        If (strProducts <> "" And productPulsarIds <> "") Then
            strProducts = strProducts + "," + productPulsarIds
        End If

        If (strProducts = "" And productPulsarIds <> "") Then
            strProducts = productPulsarIds
        End If

        If (strProducts <> "" And productPulsarIds = "") Then
            strProducts = strProducts
        End If

        If Left(strProducts, 1) = "," Then
            strProducts = Mid(strProducts, 2)
        End If


        If Request("lstProductGroups") <> "" Then
            Dim ProductGroupsArray() As String
            Dim ProductGroupArray() As String
            Dim strProductGroup As String
            Dim lastProductGroup As String
            Dim strProductGroupFilter As String
            Dim strCycleList As String

            ProductGroupsArray = Split(Request("lstProductGroups"), ",")
            lastProductGroup = 0
            strProductGroupFilter = ""
            strCycleList = ""

            For Each strProductGroup In ProductGroupsArray
                If InStr(strProductGroup, ":") > 0 Then
                    ProductGroupArray = Split(strProductGroup, ":")
                    If Trim(lastProductGroup) <> "0" And Trim(ProductGroupArray(0)) <> "2" And lastProductGroup <> Trim(ProductGroupArray(0)) Then
                        strProductGroupFilter = strProductGroupFilter & " ) and  "
                    End If
                    If Trim(lastProductGroup) <> Trim(ProductGroupArray(0)) Then
                        If Trim(ProductGroupArray(0)) = "1" Then
                            strProductGroupFilter = strProductGroupFilter & " ( partnerid = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        ElseIf Trim(ProductGroupArray(0)) = "2" Then
                            strCycleList = strCycleList & "," & CLng(ProductGroupArray(1))
                        ElseIf Trim(ProductGroupArray(0)) = "3" Then
                            strProductGroupFilter = strProductGroupFilter & " ( devcenter = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        ElseIf Trim(ProductGroupArray(0)) = "4" Then
                            strProductGroupFilter = strProductGroupFilter & " ( productstatusid = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        End If
                    Else
                        If Trim(ProductGroupArray(0)) = "1" Then
                            strProductGroupFilter = strProductGroupFilter & " or partnerid = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        ElseIf Trim(ProductGroupArray(0)) = "2" Then
                            strCycleList = strCycleList & "," & CLng(ProductGroupArray(1))
                        ElseIf Trim(ProductGroupArray(0)) = "3" Then
                            strProductGroupFilter = strProductGroupFilter & " or devcenter = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        ElseIf Trim(ProductGroupArray(0)) = "4" Then
                            strProductGroupFilter = strProductGroupFilter & " or productstatusid = " & Trim(ProductGroupArray(1))
                            lastProductGroup = Trim(ProductGroupArray(0))
                        End If
                    End If
                End If
            Next
            If strProductGroupFilter <> "" Then
                strGroupSQl = strGroupSQl & " and ( " & ScrubSQL(strProductGroupFilter) & ") ) "
            End If
            If strCycleList <> "" Then
                strGroupSQl = strGroupSQl & " and id in (Select ProductVersionid from product_program with (NOLOCK) where programid in (" & Mid(strCycleList, 2) & ")) "
            End If
            If strGroupSQl <> "" Then
                strGroupSQl = Mid(strGroupSQl, 5)

                Dim cmd As SqlCommand = New SqlCommand("Select ID from productversion with (NOLOCK) where " & strGroupSQl, _Conn)
                cmd.CommandType = CommandType.Text
                Dim dr As SqlDataReader = cmd.ExecuteReader()

                Do While dr.Read()
                    strProducts = strProducts & ", " & dr("ID")
                Loop
                dr.Close()
            End If

            If strProducts = "" Then
                strProducts = "0"
            ElseIf Left(strProducts, 2) = ", " Then
                strProducts = Mid(strProducts, 3)
            End If
        End If

        GetReportProductList = strProducts
    End Function

    Private Function GetVerionIDsForReport() As Boolean
        Dim strSQL As String = "", strSQLBase As String, criteria As String
        GetVerionIDsForReport = False

        strSQLBase = "FROM ProductVersion AS pv WITH (NOLOCK)  " & _
                        "INNER JOIN ProductFamily AS f WITH (NOLOCK) ON pv.ProductFamilyID = f.ID  " & _
                        "INNER JOIN DeliverableRoot AS r WITH (NOLOCK)  " & _
                        "INNER JOIN DeliverableVersion AS v WITH (NOLOCK) ON r.ID = v.DeliverableRootID  " & _
                        "INNER JOIN Product_Deliverable AS pd WITH (NOLOCK) ON v.ID = pd.DeliverableVersionID  " & _
                        "INNER JOIN DeliverableCoreTeam AS ct WITH (NOLOCK) ON r.CoreTeamID = ct.ID  " & _
                        "INNER JOIN DeliverableCategory AS c WITH (NOLOCK) ON r.CategoryID = c.ID  " & _
                        "INNER JOIN Vendor AS vd WITH (NOLOCK) ON v.VendorID = vd.ID  " & _
                        "INNER JOIN Employee AS e2 WITH (NOLOCK) ON r.DevManagerID = e2.ID  " & _
                        "INNER JOIN Employee AS e WITH (NOLOCK) ON v.DeveloperID = e.ID ON pv.ID = pd.ProductVersionID " & _
                        "LEFT OUTER JOIN PilotStatus AS ps WITH (NOLOCK) ON pd.PilotStatusID = ps.ID  " & _
                        "LEFT OUTER JOIN TestStatus AS ts WITH (NOLOCK) ON pd.TestStatusID = ts.ID " & _
                        "LEFT OUTER JOIN Product_Deliverable_Release AS pdr WITH (NOLOCK) ON pdr.ProductDeliverableID = pd.ID " & _
                        "LEFT OUTER JOIN ProductVersionRelease AS pvr WITH (NOLOCK) ON pvr.ID = pdr.ReleaseID " & _
                        "WHERE 1=1 "

        criteria = GetReportProductList()
        If criteria <> "" Then strSQL = strSQL & " and pd.ProductVersionID in ( " & ScrubSQL(criteria) & " )"

        criteria = Request("lstOS")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and v.ID in (Select DeliverableVersionID from os_delver with (NOLOCK) where osid in ( " & _
            ScrubSQL(criteria) & ") union select 0)"

        criteria = Request("lstLanguage")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and v.ID in (Select DeliverableVersionID from language_delver with (NOLOCK) where languageid in ( " & _
            ScrubSQL(criteria) & ")  union select 0 )"

        criteria = Request("lstVendor")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and v.VendorID in (" & ScrubSQL(criteria) & ")"

        criteria = Request("lstCategory")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and r.CategoryID in (" & ScrubSQL(criteria) & ")"

        criteria = Request("lstCoreTeam")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and r.CoreTeamID in (" & ScrubSQL(criteria) & ")"

        criteria = Request("lstDeveloper")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and v.DeveloperID in (" & ScrubSQL(criteria) & ")"

        criteria = Request("lstDevManager")
        If Left(criteria, 1) = "," Then criteria = Mid(criteria, 2)
        If criteria <> "" Then strSQL = strSQL & " and r.DevManagerID in (" & ScrubSQL(criteria) & ")"

        If Request("Type") <> "" Then
            If IsNumeric(Request("Type")) Then
                strSQL = strSQL & " and r.typeid = " & CLng(Request("Type"))
            End If
        End If

        If Request("chkSCRestricted") <> "" Then strSQL = strSQL & " and case when pv.fusionrequirements = 1 then pdr.supplychainrestriction else pd.supplychainrestriction end = 1 "

        criteria = Request("lstRoot")
        If criteria <> "" Then strSQL = strSQL & " and r.ID in (" & ScrubSQL(criteria) & ")"

        If Request("chkTarget") = "on" Then strSQL = strSQL & " and case when pv.fusionrequirements = 1 then pdr.Targeted else pd.Targeted end = 1 "
        If Request("chkInImage") = "on" Then strSQL = strSQL & " and pd.InImage = 1"

        If Request("txtSearch") <> "" Then
            Dim strSearch As String, strSQLSearch As String = ""

            strSearch = Replace(Replace(Replace(Replace(Request("txtSearch"), """", ""), "'", ""), "%", ""), "*", "")
            strSearch = ScrubSQL(strSearch)

            If Request("chkNameSearch") <> "" Then strSQLSearch = strSQLSearch & " or r.Name like '%" & strSearch & "%' or  v.DeliverableName like '%" & strSearch & "%'"
            If Request("chkChangesSearch") <> "" Then strSQLSearch = strSQLSearch & " or v.Changes like '%" & strSearch & "%'"
            If Request("chkDescriptionSearch") <> "" Then strSQLSearch = strSQLSearch & " or r.Description like '%" & strSearch & "%'"
            If Request("chkCommentsSearch") <> "" Then strSQLSearch = strSQLSearch & " or v.Comments like '%" & strSearch & "%'"
            If strSQLSearch <> "" Then
                strSQLSearch = Mid(strSQLSearch, 4)
                strSQL = strSQL & " and ( " & strSQLSearch & " )"
            End If
        End If

        If Request("chkDevelopment") <> "" Or Request("chkTest") <> "" Or Request("chkRelease") <> "" Or Request("chkComplete") <> "" Then
            Dim strSQL2 As String = ""

            If Request("chkDevelopment") <> "" Then strSQL2 = strSQL2 & " or v.location like '%Development%'"
            If Request("chkTest") <> "" Then strSQL2 = strSQL2 & " or v.location like '%Test%'"
            If Request("chkRelease") <> "" Then strSQL2 = strSQL2 & " or v.location like '%Release%'"
            If Request("chkComplete") <> "" Then strSQL2 = strSQL2 & " or v.location like '%Complete%' or  v.location like '%PM%'"

            If strSQL2 <> "" Then
                strSQL2 = Mid(strSQL2, 4)
                strSQL = strSQL & " and ( " & strSQL2 & " )"
            End If
        End If

        If Request("chkFailed") <> "" Then strSQL = strSQL & " and  v.location like '%Failed%'"

        If Trim(Request("txtNumbers")) <> "" Then strSQL = strSQL & " and v.id in ( " & ScrubSQL(Request("txtNumbers")) & " )"

        If Trim(Request("txtAdvanced")) <> "" Then strSQL = strSQL & " and ( " & ScrubSQL(Trim(Request("txtAdvanced"))) & " )"

        If strSQL <> "" Then
            ' wgomero: PBI 18749 get Product Names in addition to the ID
            strSQL = "SELECT distinct v.ID,  DOTSName = PV.DOTSName + CASE WHEN PV.FusionRequirements = 1 THEN ' (' + pvr.Name + ')' ELSE ' (Legacy)'	END " & strSQLBase & strSQL
            Dim cmd As SqlCommand = New SqlCommand(strSQL, _Conn)
            cmd.CommandType = CommandType.Text

            Try
                Dim i As Integer = 0
                Dim dr As SqlDataReader = cmd.ExecuteReader()

                _ReportQueryCount = 0
                Do While dr.Read()
                    _ReportQueryCount = _ReportQueryCount + 1
                    If i < MAXCOUNT_DELIVERABLES_REPORT Then
                        ReDim Preserve _RequestedVersionIDs(i)
                        ReDim Preserve _RequestedDOTSNames(i)
                        _RequestedVersionIDs(i) = dr("ID")
                        _RequestedDOTSNames(i) = dr("DOTSName") ' wgomero: PBI 18749 put product names into a string
                        i = i + 1
                    End If
                Loop
                dr.Close()
            Catch ex As Exception
                _VP4Web(0).ErrMsg = "FAILED: Error occurred during database query - " & ex.Message
            End Try
        Else
            _VP4Web(0).ErrMsg = "No report criteria selected. Please select the appropriate criteria and try again."
        End If
    End Function

    Private Function ScrubSQL(strWords As String) As String
        Dim badChars() As String = New String() {"select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update"}
        Dim newChars As String
        Dim i As Integer

        newChars = strWords
        For i = 0 To UBound(badChars)
            newChars = Replace(newChars, badChars(i), "")
        Next
        ScrubSQL = newChars
    End Function

    Private Sub AddToList(ByRef strList As String, strToAdd As String, seperator As String)
        strToAdd = Trim(strToAdd)
        If strToAdd <> "" Then
            If strList <> "" Then strList = strList & seperator & " "
            strList = strList & strToAdd
        End If
    End Sub

    Private Function FullName(name As String, version As String, revision As String, pass As String) As String
        Dim fn As String

        ' Create full name from Name, Version, Revision and Pass
        fn = name
        fn = fn & " [" & version
        If revision <> "" Then fn = fn & "," & revision
        If pass <> "" Then fn = fn & "," & pass
        fn = fn & "]"

        FullName = fn
    End Function

    Private Function RootIDFromVersionID(versionID As Long) As Long
        RootIDFromVersionID = 0
        On Error GoTo dbError

        Dim cmd As SqlCommand = New SqlCommand("spGetRootID", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", versionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        If (dr.Read()) Then RootIDFromVersionID = dr("ID")

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function RemoveTimeFromDateTime(dateTime As String) As String
        RemoveTimeFromDateTime = dateTime
        If InStr(dateTime, " ") > 0 Then RemoveTimeFromDateTime = Left(dateTime, InStr(dateTime, " ") - 1)
    End Function

    Private Function GetUserName(userID As Long) As String
        On Error GoTo dbError
        'Issue fix WorkItem#41131 *** START ***
        'Storeprocedure spGetEmployees take more time to get single valuse it will return more than 20000 records and loop and validate to get single value
        Dim strSQL As String
        strSQL = "SELECT Name FROM Employee with (NOLOCK) where id=" & userID
        Dim cmd As SqlCommand = New SqlCommand(strSQL, _Conn)
        cmd.CommandType = CommandType.Text
        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Do While dr.Read()
            GetUserName = dr("Name").ToString().Trim()
        Loop
        'Issue fix WorkItem#41131 *** END ***
        'Dim cmd As SqlCommand = New SqlCommand("spGetEmployees", _Conn)
        'cmd.CommandType = CommandType.StoredProcedure
        'Dim dr As SqlDataReader = cmd.ExecuteReader()

        'Do While dr.Read()
        '    If userID = dr("ID") Then
        '        GetUserName = dr("Name").ToString().Trim()
        '        Exit Do
        '    End If
        'Loop

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Protected Function SMRLink(SMRID As String) As String
        Dim IDs() As String
        Dim stringSeparators() As String = {"[-]"}
        If Not SMRID Is Nothing Then
            IDs = SMRID.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries)
            If (IDs.Length = 3) Then
                If IDs(0).Trim() <> "" And IDs(1).Trim() <> "" And IDs(2).Trim() <> "" Then
                    If _CurrentUser.PartnerID <> PARTNERID_HP Then
                        SMRLink = SMRID
                    Else
                        SMRLink = Application("Release_Houston_ServerName") & "/softpaq/SPApproval.aspx?DBID=" & IDs(1) & "&Src=" & IDs(0) & "&SCID=" & IDs(2) & "&EID=" & _CurrentUser.ID
                    End If
                End If
            End If
        End If
    End Function

    Private Function GetCurrentUserInfo() As Boolean
        GetCurrentUserInfo = False
        On Error GoTo dbError

        Dim pos As Integer
        _CurrentUser.Username = LCase(Session("LoggedInUser"))
        pos = InStr(_CurrentUser.Username, "\")
        If pos > 0 Then
            _CurrentUser.Domain = Left(_CurrentUser.Username, pos - 1)
            _CurrentUser.Username = Mid(_CurrentUser.Username, pos + 1)
        End If

        Dim cmd As SqlCommand = New SqlCommand("spGetUserInfo", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@userName", _CurrentUser.Username))
        cmd.Parameters.Add(New SqlParameter("@Domain", _CurrentUser.Domain))

        _CurrentUser.ID = 0
        Dim dr As SqlDataReader = cmd.ExecuteReader()
        If (dr.Read()) Then
            _CurrentUser.ID = dr("ID")
            _CurrentUser.IsAdmin = dr("SystemAdmin")
            _CurrentUser.PartnerID = dr("PartnerID")
            _CurrentUser.Email = dr("Email").ToString().Trim()
            _CurrentUser.Name = dr("Name").ToString().Trim()

            GetCurrentUserInfo = True
        End If

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetBuildLevel(devTypeID As Long, levelID As Long) As String
        On Error GoTo AdoError

        Dim levelType As Long
        levelType = 1
        If devTypeID <> DeliverableTypes.Hardware Then levelType = 2

        Dim cmd As SqlCommand = New SqlCommand("spListDeliverableLevels", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@TypeID", levelType))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Do While (dr.Read())
            If levelID = dr("ID") Then
                GetBuildLevel = dr("name").ToString.Trim()
                Exit Do
            End If
        Loop

AdoError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetMilestoneList() As Boolean
        GetMilestoneList = False
        On Error GoTo dbError

        Dim cmd As SqlCommand = New SqlCommand("spGetDelMilestoneList", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@DeliverableRootID", RootID))
        cmd.Parameters.Add(New SqlParameter("@DeliverableVersionID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        ReDim Preserve _VP4Web(_Idx).MilestoneList(0)
        Dim count As Integer

        count = 1
        Do While dr.Read()
            ReDim Preserve _VP4Web(_Idx).MilestoneList(count)

            _VP4Web(_Idx).MilestoneList(count).ActualDate = dr("Actual").ToString().Trim()
            If _VP4Web(_Idx).MilestoneList(count).ActualDate = "" Then _VP4Web(_Idx).MilestoneList(count).ActualDate = "&nbsp;"
            _VP4Web(_Idx).MilestoneList(count).ActualDate = RemoveTimeFromDateTime(_VP4Web(_Idx).MilestoneList(count).ActualDate)

            _VP4Web(_Idx).MilestoneList(count).Name = dr("Milestone").ToString().Trim()
            _VP4Web(_Idx).MilestoneList(count).Status = dr("Status").ToString().Trim()
            _VP4Web(_Idx).MilestoneList(count).PlannedDate = dr("Planned").ToString().Trim()
            If _VP4Web(_Idx).MilestoneList(count).PlannedDate = "" Then _VP4Web(_Idx).MilestoneList(count).PlannedDate = "&nbsp;"
            _VP4Web(_Idx).MilestoneList(count).PlannedDate = RemoveTimeFromDateTime(_VP4Web(_Idx).MilestoneList(count).PlannedDate)

            count = count + 1
        Loop
        GetMilestoneList = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Sub GetVersionDetails(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).TypeID = dr("TypeID")
        If Not IsDBNull(dr("LevelID")) Then _VP4Web(_Idx).LevelID = dr("LevelID")

        _VP4Web(_Idx).Name = dr("DeliverableName").ToString().Trim()
        If (_VP4Web(_Idx).Name = "") Then _VP4Web(_Idx).Name = dr("Name").ToString().Trim()

        _VP4Web(_Idx).Version = dr("Version").ToString().Trim()
        _VP4Web(_Idx).Revision = dr("Revision").ToString().Trim()
        If (_VP4Web(_Idx).Revision = "") Then _VP4Web(_Idx).Revision = "&nbsp;"
        _VP4Web(_Idx).Pass = dr("Pass").ToString().Trim()
        If (_VP4Web(_Idx).Pass = "") Then _VP4Web(_Idx).Pass = "&nbsp;"
        _VP4Web(_Idx).Filename = dr("Filename").ToString().Trim()
        _VP4Web(_Idx).CodeName = dr("CodeName").ToString().Trim()
        If (_VP4Web(_Idx).CodeName = "") Then _VP4Web(_Idx).CodeName = "&nbsp;"
        _VP4Web(_Idx).DeveloperID = dr("DeveloperID")
        _VP4Web(_Idx).DevManager = dr("DevManager").ToString().Trim()
        _VP4Web(_Idx).ModelNumber = dr("ModelNumber").ToString().Trim()
        If (_VP4Web(_Idx).ModelNumber = "") Then _VP4Web(_Idx).ModelNumber = "&nbsp;"
        _VP4Web(_Idx).PartNumber = dr("PartNumber").ToString().Trim()
        If (_VP4Web(_Idx).PartNumber = "") Then _VP4Web(_Idx).PartNumber = "&nbsp;"
        _VP4Web(_Idx).HFCN = dr("HFCN")
        _VP4Web(_Idx).CategoryID = dr("CategoryID")

        _VP4Web(_Idx).FullName = FullName(_VP4Web(_Idx).Name, _VP4Web(_Idx).Version, _VP4Web(_Idx).Revision, _VP4Web(_Idx).Pass)

        _VP4Web(_Idx).Vendor = dr("Vendor").ToString().Trim()
        _VP4Web(_Idx).VendorID = dr("Vendorid")
        _VP4Web(_Idx).VersionVendor = dr("VersionVendor").ToString().Trim()
        _VP4Web(_Idx).VersionVendorID = dr("VersionVendorid")
        _VP4Web(_Idx).VendorVersion = dr("VendorVersion").ToString().Trim()

        _VP4Web(_Idx).Supplier = dr("Supplier").ToString().Trim()
        _VP4Web(_Idx).SupplierID = dr("SupplierID").ToString().Trim()

        If (_VP4Web(_Idx).VendorVersion = "") Then _VP4Web(_Idx).VendorVersion = "&nbsp;"

        _VP4Web(_Idx).Multilangauge = dr("MultiLanguage")
        _VP4Web(_Idx).DeliverableSpec = dr("DeliverableSpec").ToString().Trim()
        If Left(_VP4Web(_Idx).DeliverableSpec, 2) = "\\" Then
            _VP4Web(_Idx).DeliverableSpec = "<a href=""file://" & _VP4Web(_Idx).DeliverableSpec & """>" & _VP4Web(_Idx).DeliverableSpec & "</a>"
        End If
        _VP4Web(_Idx).ImagePath = dr("ImagePath").ToString().Trim()
        If Left(_VP4Web(_Idx).ImagePath, 2) = "\\" Then
            _VP4Web(_Idx).ImagePath = "<a href=""file://" & _VP4Web(_Idx).ImagePath & """>" & _VP4Web(_Idx).ImagePath & "</a>"
        End If
        _VP4Web(_Idx).Comments = dr("Comments").ToString().Trim()
        _VP4Web(_Idx).Changes = dr("Changes").ToString().Trim()
        _VP4Web(_Idx).Changes = Replace(_VP4Web(_Idx).Changes, vbCrLf, "<br />")
    End Sub

    Private Sub GetSampleDate(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).SampleDate = dr("SampleDate").ToString().Trim()
        If _VP4Web(_Idx).SampleDate = "" Then _VP4Web(_Idx).SampleDate = "Unknown"
        _VP4Web(_Idx).SampleDate = RemoveTimeFromDateTime(_VP4Web(_Idx).SampleDate)
        _VP4Web(_Idx).SampleConfidence = dr("SamplesConfidence")
    End Sub

    Private Sub GetIntroDate(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).IntroDate = RemoveTimeFromDateTime(dr("IntroDate").ToString().Trim())
        _VP4Web(_Idx).IntroConfidence = dr("IntroConfidence")
    End Sub

    Private Sub GetAvailability(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).EOLDate = RemoveTimeFromDateTime(dr("EOLDate").ToString().Trim())
        _VP4Web(_Idx).Active = dr("ActiveVersion")
    End Sub

    Private Sub GetSpecialNotes(ByRef dr As SqlDataReader)
        If Not IsDBNull(dr("InstallableUpdate")) Then _VP4Web(_Idx).InstallableUpdate = dr("InstallableUpdate")
        If Not IsDBNull(dr("PackageForWeb")) Then _VP4Web(_Idx).PackageForWeb = dr("PackageForWeb")
        If _VP4Web(_Idx).InstallableUpdate Then _VP4Web(_Idx).SpecialNotes = "Installable&nbsp;Update"
        If _VP4Web(_Idx).PackageForWeb Then
            If _VP4Web(_Idx).SpecialNotes <> "" Then _VP4Web(_Idx).SpecialNotes = _VP4Web(_Idx).SpecialNotes & ",&nbsp;"
            _VP4Web(_Idx).SpecialNotes = _VP4Web(_Idx).SpecialNotes & "Package For Web"
        End If
    End Sub

    Private Sub GetIconsInstalled(ByRef dr As SqlDataReader)
        ' N/A on hardware
        If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then Exit Sub

        _VP4Web(_Idx).IconDesktop = dr("IconDesktop")
        _VP4Web(_Idx).IconMenu = dr("IconMenu")
        _VP4Web(_Idx).IconTray = dr("IconTray")
        _VP4Web(_Idx).IconPanel = dr("IconPanel")
        If _VP4Web(_Idx).IconDesktop Then _VP4Web(_Idx).IconsInstalled = "Desktop"
        If _VP4Web(_Idx).IconMenu Then
            If _VP4Web(_Idx).IconsInstalled <> "" Then _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & ",&nbsp;"
            _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & "Start&nbsp;Menu"
        End If
        If _VP4Web(_Idx).IconTray Then
            If _VP4Web(_Idx).IconsInstalled <> "" Then _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & ",&nbsp;"
            _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & "System&nbsp;Tray"
        End If
        If _VP4Web(_Idx).IconPanel Then
            If _VP4Web(_Idx).IconsInstalled <> "" Then _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & ",&nbsp;"
            _VP4Web(_Idx).IconsInstalled = _VP4Web(_Idx).IconsInstalled & "Control&nbsp;Panel"
        End If
    End Sub

    Private Sub GetPropertyTabs(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).PropertyTabs = dr("PropertyTabs").ToString().Trim()
        _VP4Web(_Idx).PropertyTabs = Replace(_VP4Web(_Idx).PropertyTabs, """", "&quot;")
    End Sub

    Private Sub GetPackaging(ByRef dr As SqlDataReader)
        ' N/A on hardware
        If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then Exit Sub

        _VP4Web(_Idx).Preinstall = dr("Preinstall")
        _VP4Web(_Idx).FloppyDisk = dr("FloppyDisk")
        _VP4Web(_Idx).Scriptpaq = dr("Scriptpaq")
        _VP4Web(_Idx).CDImage = dr("CDImage")
        _VP4Web(_Idx).ISOImage = dr("ISOImage")
        _VP4Web(_Idx).AR = dr("AR")

        If _VP4Web(_Idx).Preinstall Then _VP4Web(_Idx).Packaging = "Preinstall"
        If _VP4Web(_Idx).FloppyDisk Then
            If _VP4Web(_Idx).Packaging <> "" Then _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & ",&nbsp;"
            _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & "Diskette"
        End If
        If _VP4Web(_Idx).Scriptpaq Then
            If _VP4Web(_Idx).Packaging <> "" Then _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & ",&nbsp;"
            _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & "Scriptpaq"
        End If
        If _VP4Web(_Idx).CDImage Then
            If _VP4Web(_Idx).Packaging <> "" Then _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & ",&nbsp;"
            _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & "CD&nbsp;Files"
        End If
        If _VP4Web(_Idx).ISOImage Then
            If _VP4Web(_Idx).Packaging <> "" Then _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & ",&nbsp;"
            _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & "ISO&nbsp;Image"
        End If
        If _VP4Web(_Idx).AR Then
            If _VP4Web(_Idx).Packaging <> "" Then _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & ",&nbsp;"
            _VP4Web(_Idx).Packaging = _VP4Web(_Idx).Packaging & "Replicater&nbsp;Only"
        End If
    End Sub

    Private Sub GetROMComponents(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).Binary = dr("Binary")
        _VP4Web(_Idx).Rompaq = dr("Rompaq")
        _VP4Web(_Idx).PreinstallROM = dr("PreinstallROM")
        _VP4Web(_Idx).CAB = dr("CAB")

        If _VP4Web(_Idx).Binary Then _VP4Web(_Idx).ROMComponents = "Binary"
        If _VP4Web(_Idx).Rompaq Then
            If _VP4Web(_Idx).ROMComponents <> "" Then _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & ",&nbsp;"
            _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & "Rompaq"
        End If
        If _VP4Web(_Idx).PreinstallROM Then
            If _VP4Web(_Idx).ROMComponents <> "" Then _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & ",&nbsp;"
            _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & "Preinstall"
        End If
        If _VP4Web(_Idx).CAB Then
            If _VP4Web(_Idx).ROMComponents <> "" Then _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & ",&nbsp;"
            _VP4Web(_Idx).ROMComponents = _VP4Web(_Idx).ROMComponents & "CAB"
        End If
    End Sub

    Private Sub GetReplicatedBy(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).Replicater = dr("Replicater").ToString().Trim()
    End Sub

    Private Sub GetCDPartNumber(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).CDPartNumber = dr("CDPartNumber").ToString().Trim()
        _VP4Web(_Idx).CDKitNumber = dr("CDKitNumber").ToString().Trim()

        If (_VP4Web(_Idx).CDKitNumber = "N/A") Then _VP4Web(_Idx).CDKitNumber = ""
    End Sub

    Private Sub GetRoHSGreenSpec(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).RohsID = dr("RohsID")
        _VP4Web(_Idx).GreenSpecID = dr("GreenSpecID")
    End Sub

    Private Function GetRoHSGreenSpec() As Boolean
        GetRoHSGreenSpec = False
        On Error GoTo dbError

        Dim greenSpec As String, rohs As String
        Dim cmd As SqlCommand = New SqlCommand("spGetRoHSGreenDisplayName", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@RohsID", _VP4Web(_Idx).RohsID))
        cmd.Parameters.Add(New SqlParameter("@GreenSpecID", _VP4Web(_Idx).GreenSpecID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        _VP4Web(_Idx).RoHSGreenSpec = "Unknown"

        If (dr.Read()) Then
            rohs = dr("Rohs").ToString().Trim()
            greenSpec = dr("GreenSpec").ToString().Trim()

            If rohs <> "" And greenSpec <> "" Then
                _VP4Web(_Idx).RoHSGreenSpec = rohs & "_" & greenSpec
            ElseIf rohs <> "" Then
                _VP4Web(_Idx).RoHSGreenSpec = rohs
            ElseIf greenSpec <> "" Then
                _VP4Web(_Idx).RoHSGreenSpec = greenSpec
            End If
        End If
        GetRoHSGreenSpec = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetOTSList() As Boolean
        GetOTSList = False
        On Error GoTo dbError

        ReDim Preserve _VP4Web(_Idx).OTSList(0)
        Dim cmd As SqlCommand = New SqlCommand("spGetOTSByDelVersion", _Conn)
        Dim count As Integer

        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@DelVerID", VersionID))
        Dim dr As SqlDataReader = cmd.ExecuteReader()

        count = 1
        Do While dr.Read()
            ReDim Preserve _VP4Web(_Idx).OTSList(count)

            _VP4Web(_Idx).OTSList(count).Number = dr("OTSNumber").ToString().Trim()
            _VP4Web(_Idx).OTSList(count).ShortDescription = dr("shortdescription").ToString().Trim()
            _VP4Web(_Idx).OTSList(count).HTMLLink = "<a target=_blank href=""http://16.81.19.70/search/ots/Report.asp?txtReportSections=1&txtObservationID=" & _
            _VP4Web(_Idx).OTSList(count).Number & "&Sort1Column=o.observationid&Sort1Direction=asc"">" & _VP4Web(_Idx).OTSList(count).Number & "</a> - " & _VP4Web(_Idx).OTSList(count).ShortDescription

            count = count + 1
        Loop
        GetOTSList = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetSelectedOSes() As Boolean
        GetSelectedOSes = False
        On Error GoTo dbError

        Dim cmd As SqlCommand = New SqlCommand("spGetSelectedOS", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Dim count As Integer

        Do While dr.Read()
            If dr("ID") <> OSID_INDEPENDENT Then
                AddToList(_VP4Web(_Idx).SelectedOSIDs, dr("ID").ToString(), ",")
                AddToList(_VP4Web(_Idx).SelectedOSes, dr("Name").ToString(), ";")
            End If
        Loop
        GetSelectedOSes = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetSelectedLanguages() As Boolean
        GetSelectedLanguages = False
        On Error GoTo dbError

        Dim cmd As SqlCommand = New SqlCommand("spGetSelectedLanguages", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()

        Do While dr.Read()
            If dr("ID") <> LANGID_INDEPENDENT Then
                AddToList(_VP4Web(_Idx).SelectedLangIDs, dr("ID").ToString(), ",")

                If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                    AddToList(_VP4Web(_Idx).SelectedLangs, dr("Name").ToString(), ";")
                Else
                    AddToList(_VP4Web(_Idx).SelectedLangs, dr("Abbreviation").ToString() & " - " & dr("Name").ToString(), ";")

                    If dr("PartNumber").ToString().Trim() <> "" Then
                        AddToList(_VP4Web(_Idx).PartNumbers, dr("Abbreviation").ToString() & " - " & dr("Name").ToString() & ": " & dr("PartNumber").ToString(), ";")
                    End If

                    If dr("CDKitNumber").ToString().Trim() <> "" Then
                        AddToList(_VP4Web(_Idx).CDKitNumbers, dr("Abbreviation").ToString() & " - " & dr("Name").ToString() & ": " & dr("CDKitNumber").ToString(), ";")
                    End If
                End If
            End If
        Loop
        GetSelectedLanguages = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Sub GetPNPDevices(ByRef dr As SqlDataReader)
        ReDim Preserve _VP4Web(_Idx).PNPDevices(0)
        Dim devices As String
        Dim count As Integer

        count = 1
        devices = dr("PNPDevices").ToString().Trim()

        Do While InStr(devices, vbCrLf) > 0
            ReDim Preserve _VP4Web(_Idx).PNPDevices(count)
            _VP4Web(_Idx).PNPDevices(count) = Left(devices, InStr(devices, vbCrLf) - 1)
            devices = Mid(devices, InStr(devices, vbCrLf) + 2)
            count = count + 1
        Loop
        If devices <> "" Then
            ReDim Preserve _VP4Web(_Idx).PNPDevices(count)
            _VP4Web(_Idx).PNPDevices(count) = devices
        End If
    End Sub

    Private Sub GetDeliverablesDependencies(ByRef dr As SqlDataReader)
        _VP4Web(_Idx).SWDependencies = dr("SWDependencies").ToString().Trim()
    End Sub

    Private Function GetDeliverablesDependencies() As Boolean
        GetDeliverablesDependencies = False
        On Error GoTo dbError

        ReDim Preserve _VP4Web(_Idx).SelectedDelDependencies(0)
        ReDim Preserve _VP4Web(_Idx).RootDelDependencies(0)

        Dim cmd As SqlCommand = New SqlCommand("spGetSelectedDepends", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Dim count As Integer

        count = 1
        Do While dr.Read()
            ReDim Preserve _VP4Web(_Idx).SelectedDelDependencies(count)

            _VP4Web(_Idx).SelectedDelDependencies(count).ID = dr("ID")
            _VP4Web(_Idx).SelectedDelDependencies(count).Name = dr("name").ToString().Trim()
            _VP4Web(_Idx).SelectedDelDependencies(count).Version = dr("version").ToString().Trim()
            _VP4Web(_Idx).SelectedDelDependencies(count).Revision = dr("revision").ToString().Trim()
            _VP4Web(_Idx).SelectedDelDependencies(count).Pass = dr("pass").ToString().Trim()

            count = count + 1
        Loop
        dr.Close()

        cmd = New SqlCommand("spGetDepends4Version", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", RootID))

        dr = cmd.ExecuteReader()
        count = 1

        Do While dr.Read()
            ReDim Preserve _VP4Web(_Idx).RootDelDependencies(count)

            _VP4Web(_Idx).RootDelDependencies(count).ID = dr("ID")
            _VP4Web(_Idx).RootDelDependencies(count).Name = dr("name").ToString().Trim()
            _VP4Web(_Idx).RootDelDependencies(count).Version = dr("version").ToString().Trim()
            _VP4Web(_Idx).RootDelDependencies(count).Revision = dr("revision").ToString().Trim()
            _VP4Web(_Idx).RootDelDependencies(count).Pass = dr("pass").ToString().Trim()

            count = count + 1
        Loop
        dr.Close()

        For Each d As DeliverableVersionRecord In _VP4Web(_Idx).SelectedDelDependencies
            If Trim(d.Name) <> "" Then _
                AddToList(_VP4Web(_Idx).DeliverableDependencies, FullName(d.Name, d.Version, d.Revision, d.Pass), ",")
        Next

        GetDeliverablesDependencies = True

dbError:
        dr = Nothing
    End Function

     Private Function GetProductList() As Boolean
        GetProductList = False
        On Error GoTo dbError

        ReDim Preserve _VP4Web(_Idx).ProductList(0)
        Dim spName As String

        If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
            spName = "spGetproductStatus4Commodity"
        Else
            spName = "spGetproductStatus4Deliverable"
        End If

        Dim cmd As SqlCommand = New SqlCommand(spName, _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@VersionID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Dim count As Integer

        count = 1
        Do While dr.Read()
            If _CurrentUser.PartnerID = PARTNERID_HP Or _CurrentUser.PartnerID = dr("PartnerID") Then
                ReDim Preserve _VP4Web(_Idx).ProductList(count)

                _VP4Web(_Idx).ProductList(count).ID = dr("pdid")
                _VP4Web(_Idx).ProductList(count).ID = dr("ProductDeliverableReleaseID")
                _VP4Web(_Idx).ProductList(count).Name = dr("product").ToString().Trim()
                If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                    _VP4Web(_Idx).ProductList(count).ProjectManager = "Not Assigned"
                    If Not IsDBNull(dr("Commoditypm")) Then
                        _VP4Web(_Idx).ProductList(count).ProjectManager = dr("Commoditypm").ToString().Trim()
                        If _VP4Web(_Idx).ProductList(count).ProjectManager = "" Then _VP4Web(_Idx).ProductList(count).ProjectManager = "Not Assigned"
                    End If
                Else
                    _VP4Web(_Idx).ProductList(count).ProjectManager = dr("sepm").ToString().Trim()
                End If

                If dr("DeveloperNotificationStatus") = 1 Then
                    If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                        _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Approved For Testing"
                    Else
                        _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Approved"
                    End If
                ElseIf dr("DeveloperNotificationStatus") = 2 Then
                    _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Rejected"
                ElseIf dr("DeveloperNotificationStatus") = 0 Then
                    _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Under Review"
                Else
                    _VP4Web(_Idx).ProductList(count).DeveloperStatus = "TBD"
                End If

                If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                    If dr("DeveloperTestStatus") = 1 Then
                        _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Approved For Production"
                    ElseIf dr("DeveloperTestStatus") = 2 Then
                        _VP4Web(_Idx).ProductList(count).DeveloperStatus = "Not Approved For Production"
                    End If
                End If

                If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                    If dr("TestStatus").ToString().Trim() = "Date" Then
                        _VP4Web(_Idx).ProductList(count).Status = RemoveTimeFromDateTime(dr("TestDate").ToString().Trim())
                    Else
                        _VP4Web(_Idx).ProductList(count).Status = dr("TestStatus").ToString().Trim()
                        If _VP4Web(_Idx).ProductList(count).Status = "" Then
                            _VP4Web(_Idx).ProductList(count).Status = "Not Used"
                        End If
                    End If
                Else
                    If dr("Targeted") Then
                        _VP4Web(_Idx).ProductList(count).Status = "Targeted"
                    ElseIf Not dr("Prereleased") Then
                        _VP4Web(_Idx).ProductList(count).Status = "Pending"
                    Else
                        Dim pmAlert As String
                        pmAlert = dr("PMAlert").ToString().Trim()

                        _VP4Web(_Idx).ProductList(count).Status = "Available"
                        If pmAlert = "1" Or pmAlert = "True" Then _VP4Web(_Idx).ProductList(count).Status = "In Progress"
                    End If
                End If
                If _VP4Web(_Idx).TypeID = DeliverableTypes.Hardware Then
                    _VP4Web(_Idx).ProductList(count).PINOrTestNotes = dr("TargetNotes").ToString().Trim()

                    _VP4Web(_Idx).ProductList(count).PilotStatus = dr("PilotStatus").ToString().Trim()
                    If _VP4Web(_Idx).ProductList(count).PilotStatus = "P_Scheduled" Then
                        _VP4Web(_Idx).ProductList(count).PilotStatus = RemoveTimeFromDateTime(dr("PilotDate").ToString().Trim())
                    Else
                        If _VP4Web(_Idx).ProductList(count).PilotStatus = "" Then _VP4Web(_Idx).ProductList(count).PilotStatus = "N/A"

                        Dim nSupplyChainRestriction As Integer = 0, nConfigurationRestriction As Integer = 0
                        If Not IsDBNull(dr("SupplyChainRestriction")) Then nSupplyChainRestriction = dr("SupplyChainRestriction")
                        If Not IsDBNull(dr("ConfigurationRestriction")) Then nConfigurationRestriction = dr("ConfigurationRestriction")

                        If nSupplyChainRestriction <> 0 And nConfigurationRestriction <> 0 Then
                            _VP4Web(_Idx).ProductList(count).Restrictions = "Supply, Config"
                        ElseIf nSupplyChainRestriction <> 0 Then
                            _VP4Web(_Idx).ProductList(count).Restrictions = "Supply"
                        ElseIf nConfigurationRestriction <> 0 Then
                            _VP4Web(_Idx).ProductList(count).Restrictions = "Config"
                        Else
                            _VP4Web(_Idx).ProductList(count).Restrictions = "&nbsp;"
                        End If
                    End If
                Else
                    ' Defaulting bPreinstall and bInImage seem to match old code behaviour
                    Dim bPreinstall As Boolean = True, bInImage As Boolean = True
                    If Not IsDBNull(dr("Preinstall")) Then bPreinstall = dr("Preinstall")
                    If Not IsDBNull(dr("InImage")) Then bInImage = dr("InImage")

                    If (Not bPreinstall) And dr("Prereleased") Then
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "N/A"
                    ElseIf Not bPreinstall Then
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "TBD"
                    ElseIf (Not bInImage) And dr("targeted") Then
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "In Progress"
                    ElseIf dr("InImage") Then
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "In Image"
                    ElseIf _VP4Web(_Idx).ProductList(count).Status = "Available" Then
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "Available"
                    Else
                        _VP4Web(_Idx).ProductList(count).PINOrTestNotes = "Pending"
                    End If
                End If

                count = count + 1
            End If
        Loop
        GetProductList = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Function GetProductListTestingSummary() As Boolean
        GetProductListTestingSummary = False
        On Error GoTo dbError

        Dim pdids As String, pdrids As String, strSQL As String

        pdids = ""
        pdrids = ""
        For Each pr As ProductRecord In _VP4Web(_Idx).ProductList
            If pr.ID <> 0 And pr.ProductDeliverableReleaseID = 0 Then
                AddToList(pdids, pr.ID, ",")
            Else
                AddToList(pdrids, pr.ProductDeliverableReleaseID, ",")
            End If
        Next

        strSQL = "Select pd.id as pdid, pd.ConfigurationRestriction, pd.SupplyChainRestriction, v.TTS, pv.wwanproduct, pd.DeveloperTestStatus, pd.DeveloperTestNotes, pd.IntegrationTestStatus, pd.IntegrationTestNotes, pd.ODMTestStatus, pd.ODMTestnotes, c.requiresTTS, c.requiresodmtestfinalapproval, c.requiresWWANtestfinalapproval, c.requiresMITtestfinalapproval, c.requiresdeveloperfinalapproval ,pd.WWANTestStatus, pd.WWANTestNotes, pd.TestStatusID, pd.targetNotes, pd.TestConfidence, pv.devcenter, pd.TestDate, pd.DCRID,pd.developernotificationstatus, c.name as category, c.id as CategoryID, r.name as DeliverableName, v.version, v.revision, v.pass, v.vendorid, v.deliverablerootid, v.modelnumber, v.partnumber, vd.name as vendor, f.name +  ' ' + pv.version as Product, pv.PartnerID, pd.RiskRelease, ProductDeliverableReleaseID = 0" & _
            " from product_deliverable pd with (NOLOCK), deliverableversion v with (NOLOCK), productversion pv with (NOLOCK), productfamily f with (NOLOCK), vendor vd with (NOLOCK), deliverableroot r with (NOLOCK), deliverablecategory c with (NOLOCK) " & _
            " where(PV.id = pd.productversionid)" & _
            " and pd.deliverableversionid = v.id" & _
            " and f.id=pv.productfamilyid" & _
            " and c.id = r.categoryid" & _
            " and vd.id=v.vendorid" & _
            " and r.id = v.deliverablerootid" & _
            " and pd.deliverableversionid = " & VersionID & _
            " and pd.id in (" & pdids & ") Union " & _
            "Select pd.id as pdid, pdr.ConfigurationRestriction, pdr.SupplyChainRestriction, v.TTS, pv.wwanproduct, pdr.DeveloperTestStatus, pdr.DeveloperTestNotes, IntegrationTestStatus = isnull(pdr.IntegrationTestStatus,0), IntegrationTestNotes = isnull(pdr.IntegrationTestNotes,''), ODMTestStatus = isnull(pdr.ODMTestStatus,0), ODMTestnotes = isnull(pdr.ODMTestnotes,''), c.requiresTTS, c.requiresodmtestfinalapproval, c.requiresWWANtestfinalapproval, c.requiresMITtestfinalapproval, c.requiresdeveloperfinalapproval , WWANTestStatus = isnull(pdr.WWANTestStatus,0), WWANTestNotes = isnull(pdr.WWANTestNotes,''), pdr.TestStatusID, targetNotes = isnull(pdr.targetNotes,''), pdr.TestConfidence, pv.devcenter, pdr.TestDate, pdr.DCRID, pdr.developernotificationstatus, c.name as category, c.id as CategoryID, r.name as DeliverableName, v.version, v.revision, v.pass, v.vendorid, v.deliverablerootid, v.modelnumber, v.partnumber, vd.name as vendor, f.name +  ' ' + pv.version as Product, pv.PartnerID, pdr.RiskRelease, pdr.id as ProductDeliverableReleaseID" & _
            " from product_deliverable pd with (NOLOCK), product_deliverable_release pdr with (NOLOCK), deliverableversion v with (NOLOCK), productversion pv with (NOLOCK), productfamily f with (NOLOCK), vendor vd with (NOLOCK), deliverableroot r with (NOLOCK), deliverablecategory c with (NOLOCK) " & _
            " where(PV.id = pd.productversionid)" & _
            " and pd.deliverableversionid = v.id" & _
            " and pd.id = pdr.ProductDeliverableID" & _
            " and f.id=pv.productfamilyid" & _
            " and c.id = r.categoryid" & _
            " and vd.id=v.vendorid" & _
            " and r.id = v.deliverablerootid" & _
            " and pd.deliverableversionid = " & VersionID & _
            " and pdr.id in (" & pdrids & ")"

        Dim cmd As SqlCommand = New SqlCommand(strSQL, _Conn)
        cmd.CommandType = CommandType.Text

        Dim blnRequiresODMSignoff As Boolean, blnRequiresMITSignoff As Boolean, blnRequiresDeveloperSignoff As Boolean
        Dim blnRequiresWWANSignoff As Boolean
        Dim TestStatusArray() As String, testStatus As String
        Dim dr As SqlDataReader = cmd.ExecuteReader()

        _VP4Web(_Idx).ProductListHasTestingSummary = False
        TestStatusArray = Split("TBD," & _
            "<span class=""testStatusPassed"">Passed</span>," & _
            "<span class=""testStatusFailed"">Failed</span>," & _
            "<span class=""testStatusBlocked"">Blocked</span>,Watch,N/A", ",")

        Do While dr.Read()
            For i As Integer = 1 To UBound(_VP4Web(_Idx).ProductList)
                If _VP4Web(_Idx).ProductList(i).ID = dr("pdid") And _VP4Web(_Idx).ProductList(i).ProductDeliverableReleaseID = dr("ProductDeliverableReleaseID") Then
                    'blnRequiresDeveloperSignoff = dr("requiresdeveloperfinalapproval")
                    blnRequiresODMSignoff = dr("requiresodmtestfinalapproval")
                    blnRequiresMITSignoff = dr("requiresMITtestfinalapproval")
                    blnRequiresWWANSignoff = False
                    If Not IsDBNull(dr("WWANProduct")) Then blnRequiresWWANSignoff = dr("WWANProduct") And dr("requiresWWANtestfinalapproval")

                    If blnRequiresMITSignoff And dr("IntegrationTestStatus") >= 0 And dr("IntegrationTestStatus") <= UBound(TestStatusArray) Then
                        AddToList(_VP4Web(_Idx).ProductList(i).TestingSummary, "MIT(" & TestStatusArray(dr("IntegrationTestStatus")) & ")", ",")
                    End If

                    If blnRequiresODMSignoff And dr("ODMTestStatus") >= 0 And dr("ODMTestStatus") <= UBound(TestStatusArray) Then
                        testStatus = TestStatusArray(dr("ODMTestStatus"))
                        AddToList(_VP4Web(_Idx).ProductList(i).TestingSummary, "ODM(" & TestStatusArray(dr("ODMTestStatus")) & ")", ",")
                    End If

                    If blnRequiresWWANSignoff And dr("WWANTestStatus") >= 0 And dr("WWANTestStatus") <= UBound(TestStatusArray) Then
                        testStatus = TestStatusArray(dr("WWANTestStatus"))
                        AddToList(_VP4Web(_Idx).ProductList(i).TestingSummary, "COMM(" & testStatus & ")", ",")
                        AddToList(_VP4Web(_Idx).ProductList(i).TestingSummary, "TTS(" & dr("TTS") & ")", ",")
                    End If

                    If _VP4Web(_Idx).ProductList(i).TestingSummary <> "" Then _VP4Web(_Idx).ProductListHasTestingSummary = True
                End If
            Next
        Loop
        GetProductListTestingSummary = True

dbError:
        If Not dr Is Nothing Then dr.Close()
        dr = Nothing
    End Function

    Private Function GetVersionProperties4Web() As Boolean
        GetVersionProperties4Web = False
        On Error GoTo dbError

        ReDim Preserve _VP4Web(_Idx)
        Dim cmd As SqlCommand = New SqlCommand("spGetVersionProperties4Web", _Conn)
        cmd.CommandType = CommandType.StoredProcedure
        cmd.Parameters.Add(New SqlParameter("@ID", VersionID))

        Dim dr As SqlDataReader = cmd.ExecuteReader()
        Dim propExist As Boolean = False

        If (dr.Read()) Then
            propExist = True

            GetVersionDetails(dr)
            GetSampleDate(dr)
            GetIntroDate(dr)
            GetAvailability(dr)
            GetSpecialNotes(dr)
            GetIconsInstalled(dr)
            GetPropertyTabs(dr)
            GetPackaging(dr)
            GetROMComponents(dr)
            GetReplicatedBy(dr)
            GetCDPartNumber(dr)
            GetRoHSGreenSpec(dr)
            GetPNPDevices(dr)
            GetDeliverablesDependencies(dr)
        End If
        dr.Close()

        If propExist Then
            GetMilestoneList()
            GetRoHSGreenSpec()
            GetOTSList()
            GetSelectedOSes()
            GetSelectedLanguages()
            GetDeliverablesDependencies()
            GetProductList()
            GetProductListTestingSummary()

            _VP4Web(_Idx).Developer = GetUserName(_VP4Web(_Idx).DeveloperID)
            _VP4Web(_Idx).BuildLevel = GetBuildLevel(_VP4Web(_Idx).TypeID, _VP4Web(_Idx).LevelID)

            If CurrentUser.Name = _VP4Web(_Idx).DevManager Or CurrentUser.ID = _VP4Web(_Idx).DeveloperID Or _
                CurrentUser.IsAdmin Then _
                _VP4Web(_Idx).VersionIDHTML = "<a target=""_blank"" href=""../WizardFrames.asp?Type=1&ID=" & _
                    _VP4Web(_Idx).VersionID & """>" & _VP4Web(_Idx).VersionID & "</a>"

            GetVersionProperties4Web = True
        End If
dbError:
        _Idx = _Idx + 1 ' Always advance to the next index
    End Function

    Protected Function GetNextDeliverable() As Boolean
        GetNextDeliverable = False
        _Idx = _Idx + 1

        If UBound(_VP4Web) >= _Idx Then GetNextDeliverable = True Else _Idx = 0
    End Function

    Protected Function IsValidQuery() As Boolean
        IsValidQuery = VP4Web.RootID <> 0 And VP4Web.VersionID <> 0 And VP4Web.Name <> ""
    End Function

    Public ReadOnly Property RootID() As Long
        Get
            Return _VP4Web(_Idx).RootID
        End Get
    End Property

    Public ReadOnly Property VersionID() As Long
        Get
            Return _VP4Web(_Idx).VersionID
        End Get
    End Property

    Public ReadOnly Property DOTSName() As String
        Get
            Return _VP4Web(_Idx).DOTSName
        End Get
    End Property

    Public ReadOnly Property ErrorMsg() As String
        Get
            Return _VP4Web(_Idx).ErrMsg
        End Get
    End Property

    Public ReadOnly Property CurrentUser() As CurrentUserInfo
        Get
            Return _CurrentUser
        End Get
    End Property

    Public ReadOnly Property AdvancedSearch() As Boolean
        Get
            Return _AdvancedSearch
        End Get
    End Property

    Public ReadOnly Property ContentFormat As Long
        Get
            Return _ContentFormat
        End Get
    End Property

    Public ReadOnly Property VP4Web() As VersionProperties4Web
        Get
            Return _VP4Web(_Idx)
        End Get
    End Property

    Public ReadOnly Property ReportDeliverablesCount As Integer
        Get
            Return _VP4Web.Length
        End Get
    End Property

    Public ReadOnly Property ReportAvailableDeliverablesCount As Integer
        Get
            Return _ReportQueryCount
        End Get
    End Property
End Class
