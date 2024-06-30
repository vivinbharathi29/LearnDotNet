Imports System.Data
Imports System.Data.SqlClient

Partial Class Service_OARpt
    Inherits System.Web.UI.Page

    Private intUserId As Integer
    Protected Property UserId As Integer
        Get
            Return intUserId
        End Get
        Set(ByVal value As Integer)
            intUserId = value
        End Set
    End Property

    Private intPartnerId As Integer
    Protected Property PartnerId As Integer
        Get
            Return intPartnerId
        End Get
        Set(ByVal value As Integer)
            intPartnerId = value
        End Set
    End Property


    Private intReturnCode As Integer
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

    Private intShowAll As Integer = 0
    Protected Property ShowAll as integer
        Get
            If Integer.TryParse(Request("ShowAll"), intShowAll) Then
                Return intShowAll
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intShowAll = value
        End Set
    End Property

    Private intBusiness As Integer = 0
    Protected Property Business As Integer
        Get
            If Integer.TryParse(Request("Business"), intBusiness) Then
                Return intBusiness
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intBusiness = value
        End Set
    End Property

    Private intDevCenter As Integer = 0
    Protected Property DevCenter As Integer
        Get
            If Integer.TryParse(Request("DevCenter"), intDevCenter) Then
                Return intDevCenter
            Else
                Return 0
            End If

        End Get
        Set(ByVal value As Integer)
            intDevCenter = value
        End Set
    End Property

    Private intOSSP As Integer = 0
    Protected Property OSSP As Integer
        Get
            If Integer.TryParse(Request("OSSP"), intOSSP) Then
                Return intOSSP
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Integer)
            intOSSP = value
        End Set
    End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.Cache.SetAllowResponseInBrowserHistory(False)
        Response.Cache.SetNoServerCaching()

        Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(Session("LoggedInUser"))
        Me.UserId = objSec.CurrentUserID
        Me.PartnerId = objSec.CurrentPartnerID
        objSec = Nothing

        Dim strRC As String = "0"
        Dim objDT As DataTable = HPQ.Excalibur.Service.SelectServiceFamilyOsspAssignments(Me.UserId.ToString, Me.ShowAll.ToString, Me.DevCenter.ToString, Me.Business.ToString, Me.OSSP.ToString, strRC, strReturnDesc)

        Dim intNumRecords As Integer = objDT.Rows.Count

        Me.ReturnCode = Integer.Parse(strRC)
        Me.ReturnDesc = strReturnDesc ' Should not need this bloack

        ViewState("RptData") = objDT
        grdVwRpt.DataSource = objDT
        grdVwRpt.DataBind()

        objDT = Nothing

        If Not Me.Page.IsPostBack Then ' Initial load
            ViewState("SortDirection") = "ASC"
            ViewState("SortExpression") = "Product Version"
            ViewState("UserID") = Me.UserId
            ViewState("PartnerID") = Me.PartnerId
        End If

        If (Me.ReturnCode <> 0) Or (intNumRecords = 0) Then
            ' Hide the Export to Excel Link
            Dim objLink As LinkButton = DirectCast(Me.FindControl("lnkBtnExpExcel"), LinkButton)
            objLink.Visible = False

            If (Me.ReturnCode <> 0) Then
                ' Return error message in bold red font
                DisplayMessage("<b>" & Me.ReturnDesc & "</b>", "lblStatus", Drawing.Color.Red)
            Else
                DisplayMessage(Me.ReturnDesc, "lblStatus", Drawing.Color.Black)
            End If

        Else
            DisplayMessage(Me.ReturnDesc, "lblStatus", Drawing.Color.Black)
        End If



    End Sub

    Private Sub DisplayMessage(ByVal strMsg As String, ByVal strControlName As String, ByVal objColor As System.Drawing.Color)
        Dim objStatusLabel As Label = DirectCast(Me.FindControl(strControlName), Label)
        objStatusLabel.ForeColor = objColor
        objStatusLabel.Text = strMsg
    End Sub


    Private Function GetSortDirection(ByVal column As String) As String

        ' By default, set the sort direction to ascending.
        Dim sortDirection As String = "ASC"

        ' Retrieve the last column that was sorted.
        Dim sortExpression As String = TryCast(ViewState("SortExpression"), String)

        If sortExpression IsNot Nothing Then
            ' Check if the same column is being sorted.
            ' Otherwise, the default value can be returned.
            If sortExpression = column Then

                Dim lastDirection As String = TryCast(ViewState("SortDirection"), String)

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

    End Function

    Protected Sub grdVwRpt_Sorting(ByVal sender As Object, ByVal e As GridViewSortEventArgs)
        Dim objDT As New DataTable

        objDT = ViewState("RptData")
        objDT.DefaultView.Sort = e.SortExpression & " " & GetSortDirection(e.SortExpression)

        ViewState("RptData") = objDT.DefaultView.ToTable
        grdVwRpt.DataSource = ViewState("RptData")
        grdVwRpt.DataBind()

        objDT = Nothing

        DisplayMessage("Sorted records by [" & ViewState("SortExpression").ToString & "] in " & ViewState("SortDirection") & "ENDING order.", "lblStatus", Drawing.Color.Black)
    End Sub


    Protected Sub lnkBtnExpExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBtnExpExcel.Click

        ' Validate user requesting this download
        Dim objSec As HPQ.Excalibur.Security = New HPQ.Excalibur.Security(Session("LoggedInUser"))
        Dim intUserId As Integer = objSec.CurrentUserID
        Dim intPartnerId As Integer = objSec.CurrentPartnerID
        objSec = Nothing

        If ((ViewState("UserID") = intUserId) And (ViewState("PartnerID") = intPartnerId)) Then

            ' MAY WANT TO EXPAND IN THE FUTURE TO ACCOMMODATE LATEST AND/OR MULTIPLE VERSIONS OF MS-EXCEL
            Response.Clear()
            Response.AddHeader("content-disposition", "attachment;filename=FileName.xls")
            Response.Charset = "" ' ISO or the like
            'Response.Cache.SetCacheability(HttpCacheability.NoCache)
            Response.ContentType = "application/vnd.xls"

            Dim objStrWriter As System.IO.StringWriter = New System.IO.StringWriter
            Dim objHTMLWriter As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(objStrWriter)

            Dim objForm As HtmlForm = New HtmlForm

            Dim objGV As GridView = New GridView
            objForm.Controls.Add(objGV)

            Dim objDT As New DataTable

            objDT = ViewState("RptData")
            objDT.DefaultView.Sort = ViewState("SortExpression") & " " & ViewState("SortDirection")

            objGV.DataSource = objDT.DefaultView.ToTable
            objGV.DataBind()

            objGV.RenderControl(objHTMLWriter)

            Response.Write(objStrWriter.ToString)

            objDT = Nothing
            objStrWriter = Nothing
            objHTMLWriter = Nothing

            Response.End()

        Else
            DisplayMessage("<b>User credentials not verified for this action.</b>", "lblStatus", Drawing.Color.Red)
        End If

    End Sub

    '##############################################################################################################################################
    ' CREATE THE APPROPRIATE VERSION OF THE DISPLAY BAR
    '##############################################################################################################################################
    Public Function CreateDisplayBar() As String

        Dim strDisplayBar As String = "<table class=""DisplayBar"" Width=100% CellSpacing=0 CellPadding=2 >"

        Try

            strDisplayBar += "<TR><TD valign=top><table><tr><td valign=top><font color=navy face=verdana size=2><b>Display:&nbsp;&nbsp;&nbsp;</b></font></td></tr></table><TD width=100%><table>"

            ' Get the list of Product Statuses
            strDisplayBar += "<tr><td nowrap><b>Show:</b></td><td width='100%'>" & GetProductStatuses(Me.ShowAll.ToString) & "</td></tr>"

            ' Get the list of Dev Centers
            strDisplayBar += "<tr><td nowrap><b>Dev Center:</b></td><td width='100%'>" & GetDevCenters(Me.DevCenter.ToString) & "</td></tr>"

            ' Get the list of Business Links
            strDisplayBar += "<tr><td nowrap><b>Business:</b></td><td width='100%'>" & GetBusinessLinks(Me.Business) & "</td></tr>"

            If Me.PartnerId = 1 Then ' Get a list of all active OSSPs
                strDisplayBar += "<tr><td nowrap><b>OSSP:</b></td><td width='100%' >" & GetOSSPs(Me.OSSP.ToString) & "</td></tr>"
            End If

            strDisplayBar += "</table></td></tr></table>"

        Catch ex As Exception
            DisplayMessage(Me.ReturnDesc, "lblStatus", Drawing.Color.Red)

        End Try

        CreateDisplayBar = strDisplayBar

    End Function

    '##############################################################################################################################################
    ' GET A LIST OF ALL PRODUCT STATUSES CURRENTLY IN THE SYSTEM  ---> EVENTUALLY MAKE A METHOD CALL
    '##############################################################################################################################################
    Private Function GetProductStatuses(ByVal strSelection As String) As String

        Dim strResult As String = ""
        Dim strID As String
        Dim strDesc As String
        Dim objDT As DataTable

        Try

            strID = "0"
            strDesc = "All Active"

            If (Trim(strID) = Trim(strSelection)) Then
                strResult = strDesc
            Else
                strResult = "<a href='OARpt.aspx?ShowAll=" & strID & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
            End If

            objDT = HPQ.Excalibur.Product.ListProductStatuses()

            If (objDT.Rows.Count > 0) Then

                objDT = objDT.DefaultView.ToTable

                For Each objRow As DataRow In objDT.Rows
                    strID = objRow("ID").ToString
                    strDesc = objRow("Name").ToString

                    If (Trim(strID) = Trim(strSelection)) Then
                        If (Len(strResult) > 0) Then
                            strResult += "&nbsp;|&nbsp;" & strDesc
                        Else
                            strResult = strDesc
                        End If
                    Else
                        If (Len(strResult) > 0) Then
                            strResult += "&nbsp;|&nbsp;<a href='OARpt.aspx?ShowAll=" & strID & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
                        Else
                            strResult = "<a href='OARpt.aspx?ShowAll=" & strID & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
                        End If
                    End If

                Next

            End If


        Catch ex As Exception
            Me.ReturnDesc = ex.Message

        Finally

        End Try

        GetProductStatuses = strResult

    End Function
    '##############################################################################################################################################

    '##############################################################################################################################################
    ' GET A LIST OF ALL DEV CENTERS CURRENTLY IN THE SYSTEM 
    '##############################################################################################################################################
    Private Function GetDevCenters(ByVal strSelection As String) As String
        Dim strResult As String = ""
        Dim strID As String
        Dim strDesc As String
        Dim objDT As DataTable

        Try

            strID = "0"
            strDesc = "All"

            If (Trim(strID) = Trim(strSelection)) Then
                strResult = strDesc
            Else
                strResult = "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & strID & "'>" & strDesc & "</a>"
            End If

            objDT = HPQ.Excalibur.Product.ListDevCenters()

            If (objDT.Rows.Count > 0) Then

                objDT = objDT.DefaultView.ToTable

                For Each objRow As DataRow In objDT.Rows
                    strID = objRow("ID").ToString
                    strDesc = objRow("Name").ToString

                    If (Trim(strID) = Trim(strSelection)) Then
                        If (Len(strResult) > 0) Then
                            strResult += "&nbsp;|&nbsp;" & strDesc
                        Else
                            strResult = strDesc
                        End If
                    Else
                        If (Len(strResult) > 0) Then
                            strResult += "&nbsp;|&nbsp;<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & strID & "'>" & strDesc & "</a>"
                        Else
                            strResult += "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&Business=" & Me.Business.ToString & "&DevCenter=" & strID & "'>" & strDesc & "</a>"
                        End If
                    End If

                Next

            End If

        Catch ex As Exception
            Me.ReturnDesc = ex.Message
        Finally

        End Try


        GetDevCenters = strResult

    End Function
    '##############################################################################################################################################

    '##############################################################################################################################################
    ' GET A LIST OF ALL ACTIVE OSSPS CURRENTLY IN THE SYSTEM 
    '##############################################################################################################################################
    Private Function GetOSSPs(ByVal strSelection As String) As String
        Dim strResult As String = ""
        Dim strID As String
        Dim strDesc As String
        Dim blnActive As Boolean
        Dim objData As HPQ.Excalibur.Data
        Dim objDT As DataTable

        Try
            strID = "0"
            strDesc = "All Active"

            If (Trim(strID) = Trim(strSelection)) Then
                strResult = strDesc
            Else
                strResult = "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & strID & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
            End If

            objData = New HPQ.Excalibur.Data
            objDT = objData.ListPartners("1", "2")

            If (objDT.Rows.Count > 0) Then

                For Each objRow As DataRow In objDT.Rows

                    strID = objRow("ID").ToString
                    strDesc = objRow("name").ToString
                    blnActive = objRow("active")

                    If (blnActive) Then
                        If (Trim(strID) = Trim(strSelection)) Then
                            If (Len(strResult) > 0) Then
                                strResult += "&nbsp;|&nbsp;" & strDesc
                            Else
                                strResult = strDesc
                            End If
                        Else
                            If (Len(strResult) > 0) Then
                                strResult += "&nbsp;|&nbsp;<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & strID & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
                            Else
                                strResult += "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & strID & "&Business=" & Me.Business.ToString & "&DevCenter=" & Me.DevCenter.ToString & "'>" & strDesc & "</a>"
                            End If
                        End If
                    End If
                Next

            End If

        Catch ex As Exception
            Me.ReturnDesc = ex.Message
        Finally
            objData = Nothing
        End Try

        GetOSSPs = strResult

    End Function
    '##############################################################################################################################################


    '##############################################################################################################################################
    ' CREATE THE BUSINESS LINKS
    '##############################################################################################################################################
    Private Function GetBusinessLinks(ByVal intSelection As Integer) As String
        Dim strResult As String = ""

        Select Case intSelection
            Case 1
                strResult = _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
                 "Commercial&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=2'>Consumer</a></font>"
            Case 2
                strResult = _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
                 "Consumer</font>"
            Case 0
                strResult = _
                 "All&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=2'>Consumer</a></font>"
            Case Else
                strResult = _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=0'>All</a>&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=1'>Commercial</a>&nbsp;|&nbsp;" & _
                 "<a href='OARpt.aspx?ShowAll=" & Me.ShowAll.ToString & "&OSSP=" & Me.OSSP.ToString & "&DevCenter=" & Me.DevCenter.ToString & "&Business=2'>Consumer</a></font>"
        End Select

        GetBusinessLinks = strResult

    End Function
    '##############################################################################################################################################

End Class
