Imports HPQ.Excalibur
Imports System.Data
Imports System.Configuration
Imports System.IO
Imports System.Net.Mail
Imports System.Security.Principal

Partial Class Users_OdmUserMaint
    Inherits System.Web.UI.Page

#Region " Properties "
    Private _EmployeeId As Integer = 0
    Public ReadOnly Property EmployeeID() As Integer
        Get
            If _EmployeeId = 0 Then
                Integer.TryParse(Request.QueryString("EID"), _EmployeeId)
            End If
            Return _EmployeeId
        End Get
    End Property

    Private _currentDomain As String
    Public Property CurrentDomain() As String
        Get
            If _currentDomain = String.Empty Then
                _currentDomain = ViewState.Item("currentDomain").ToString()
            End If
            Return _currentDomain
        End Get
        Set(ByVal value As String)
            _currentDomain = value
            ViewState.Item("currentDomain") = _currentDomain
        End Set
    End Property

    Private _currentUser As String
    Public Property CurrentUser() As String
        Get
            If _currentUser = String.Empty Then
                _currentUser = ViewState.Item("currentUser").ToString()
            End If
            Return _currentUser
        End Get
        Set(ByVal value As String)
            _currentUser = value
            ViewState.Item("currentUser") = _currentUser
        End Set
    End Property

    Private _currentUserId As Integer
    Public Property CurrentUserID() As Integer
        Get
            If _currentUserId = 0 Then
                Integer.TryParse(ViewState.Item("currentUserID").ToString, _currentUserId)
            End If
            Return _currentUserId
        End Get
        Set(ByVal value As Integer)
            _currentUserId = value
            ViewState.Item("currentUserID") = _currentUserId
        End Set
    End Property

    Private _currentUserPartnerID As Integer
    Public Property CurrentUserPartnerID() As Integer
        Get
            If _currentUserPartnerID = 0 Then
                Integer.TryParse(ViewState.Item("currentUserPartnerID").ToString(), _currentUserPartnerID)
            End If
            Return _currentUserPartnerID
        End Get
        Set(ByVal value As Integer)
            _currentUserPartnerID = value
            ViewState.Item("currentUserPartnerID") = _currentUserPartnerID
        End Set
    End Property

    Private _currentUserIsBpiaApprover As Boolean
    Public Property CurrentUserIsBpiaApprover() As Boolean
        Get
            Boolean.TryParse(ViewState.Item("currentUserIsBpiaApprover").ToString(), _currentUserIsBpiaApprover)
            Return _currentUserIsBpiaApprover
        End Get
        Set(ByVal value As Boolean)
            _currentUserIsBpiaApprover = value
            ViewState.Item("currentUserIsBpiaApprover") = _currentUserIsBpiaApprover
        End Set
    End Property

    Private Const ACTIVE As String = "1"
    Private Const DELETE As String = "3"

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dt As DataTable

        If Not Page.IsPostBack Then
            Dim secObj As HPQ.Excalibur.Security = New Security(Session("LoggedInUser"))
            CurrentUser = secObj.CurrentUser
            CurrentDomain = secObj.CurrentUserDomain
            CurrentUserID = secObj.CurrentUserID
            CurrentUserPartnerID = secObj.CurrentPartnerID
            CurrentUserIsBpiaApprover = secObj.BpiaApprover
        End If

        If CurrentUserPartnerID <> 1 And Not CurrentUserIsBpiaApprover Then
            divMain.Visible = False
            Response.Write("You are not authorized to view this page.")
            Response.End()
        End If

        If Not Page.IsPostBack Then
            dt = hpqData.SelectEmployees(EmployeeID, String.Empty, String.Empty, String.Empty, String.Empty)
            If dt.Rows.Count = 0 Then
                divMain.Visible = False
                Response.Write("User not found")
                Response.End()
            End If
            divStatus.Visible = False
            lblUserName.Text = dt.Rows(0)("Name")
        End If
    End Sub

    Protected Sub btnReActivate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReActivate.Click
        HPQ.Excalibur.Employee.UpdateEmployeeOdmLoginStatus(EmployeeID.ToString(), ACTIVE)
        lblStatus.Text = "User has been reactivated."
        divStatus.Visible = True
        divButtons.Visible = False
    End Sub

    Protected Sub btnRemoveAccess_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRemoveAccess.Click
        HPQ.Excalibur.Employee.UpdateEmployeeOdmLoginStatus(EmployeeID.ToString(), DELETE)
        lblStatus.Text = "User has been scheduled for deactivation."
        divStatus.Visible = True
        divButtons.Visible = False
    End Sub
End Class
