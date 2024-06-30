Imports System.Data

Partial Class Search_SpareKitDetails
    Inherits System.Web.UI.Page

    Private Property SpareKitID() As String
        Get
            Return ViewState("SpareKitID")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitID") = value
        End Set
    End Property

    Private Property SpareKitNumber() As String
        Get
            Return ViewState("SpareKitNumber")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitNumber") = value
        End Set
    End Property

    Private Property SpareKitServiceFamilyPn() As String
        Get
            Return ViewState("SpareKitServiceFamilyPn")
        End Get
        Set(ByVal value As String)
            ViewState("SpareKitServiceFamilyPn") = value
        End Set
    End Property

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
              
                Dim sPrevPage As String = Request.QueryString("PageName")

                ' we can get this page through the Summary or the main one and there are not parameters for the query.
                If sPrevPage = "SpareKit.aspx" Then

                Else
                    If Not Request.QueryString("SpareKitID") Is Nothing Then SpareKitID = Request.QueryString("SpareKitID")
                    If Not Request.QueryString("SpareKitNumber") Is Nothing Then SpareKitNumber = Request.QueryString("SpareKitNumber")
                    If Not Request.QueryString("ServiceFamilyPn") Is Nothing Then SpareKitServiceFamilyPn = Request.QueryString("ServiceFamilyPn")

                    If SpareKitNumber = String.Empty Then
                        msgSearchNoData.Text = "You need to select a spare kit number to get the deatails."
                        Exit Sub
                    End If
                    getSpareKit()
                    GetSpareKitFamilyPartNumberDetails()

                    lblLastRunDate.Text = Date.Now.ToLongDateString()
                End If

            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getSpareKit()
        Try
            Dim dtData As New DataTable

            'spGetServiceSpareKit
            dtData = HPQ.Excalibur.Service.getServiceSpareKit(SpareKitNumber)

            If dtData.Rows.Count > 0 Then
                Dim Row As DataRow = dtData.Rows(0)

                lblSpareKitNumberValue.Text = Row("SpareKitNo").ToString
                lblSKCategoryValue.Text = Row("CategoryName").ToString
                lblSpareKitDescValue.Text = Row("SpareKitDesc").ToString

                lblMaterialTypeValue.Text = Row("MaterialType")
                lblDivisionValue.Text = Row("Division")
                lblRevisionLevelValue.Text = Row("RevisionLevel").ToString
                lblCrossPlantStatusValue.Text = Row("CrossPlantStatus").ToString
                lblOSSPOrderableValue.Text = Row("OsspOrderable").ToString
            Else
                msgSearchNoData.Text = "There are not Spare Kits for the filters selected."
            End If
           
          
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub GetSpareKitFamilyPartNumberDetails()
        Try
            Dim dtData As New DataTable

            '"usp_SelectServiceProgramBomSS" p_ServiceFamilyPn,p_SpareKitNumber
            dtData = HPQ.Excalibur.Service.GetServiceSpareKitBomDetails(SpareKitServiceFamilyPn, SpareKitNumber)
            msgSearchNoData.Visible = False
            If dtData.Rows.Count > 0 Then
                dgData.DataSource = dtData
            Else
                msgSearchNoData.Visible = True
                msgSearchNoData.Text = "There are not Spare Kits for the filters selected."
            End If

            dgData.DataBind()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

      
End Class
