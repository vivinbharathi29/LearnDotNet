<%@ Page Language="VB" %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    
    Private _HpPartNo As String = Nothing
    ReadOnly Property HpPartNo() As String
        Get
            If IsNothing(_HpPartNo) Then
                _HpPartNo = Request.QueryString("HPPN")
            End If
            Return _HpPartNo
        End Get
    End Property

    Private _ServiceFamilyPn As String = Nothing
    ReadOnly Property ServiceFamilyPn() As String
        Get
            If IsNothing(_ServiceFamilyPn) Then
                _ServiceFamilyPn = Request.QueryString("SFPN")
            End If
            Return _ServiceFamilyPn
        End Get
    End Property
    
    Private blnOSSPUser As Boolean = False
    Public Property OSSPUser As Boolean
        Get
            Return blnOSSPUser
        End Get
        Set(ByVal value As Boolean)
            blnOSSPUser = value
        End Set
    End Property
    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        
        
        '******************************************************************************************************
        '
        ' DETERMINE IF THE USER IS AN OSSP PARTNER (TYPE=2)
        '
        '******************************************************************************************************
        Me.OSSPUser = IsOSSPUser()
        '******************************************************************************************************           
        
        If Not Page.IsPostBack Then
            If HpPartNo = String.Empty Or ServiceFamilyPn = String.Empty Then
                lblError.Text = "Required Parameters are missing.<br />Page Load Aborted."
                lblError.Visible = True
                
                If (Me.OSSPUser) Then
                    dvSpareKitRO.Visible = False
                Else
                    dvSpareKit.Visible = False
                End If
                
            Else
                lblError.Visible = False
                
                If (Me.OSSPUser) Then
                    ' dvSpareKit.Visible = False
                    dvSpareKitRO.Visible = True
                Else
                    dvSpareKit.Visible = True
                    ' dvSpareKitRO.Visible = False
                End If
                
                Bind_dvSpareKit()
            End If
        End If
        
    End Sub

    Private Function IsOSSPUser() As Boolean
        Dim blnResult As Boolean = False
        Dim strCurrentUser As String = LCase(Trim(Session("LoggedInUser")))
        Dim strCurrentDomain As String = ""
        Dim intCurrentPartnerTypeID As Integer
        Dim objRow As DataRow

        If InStr(strCurrentUser, "\") > 0 Then
            strCurrentDomain = Left(strCurrentUser, InStr(strCurrentUser, "\") - 1)
            strCurrentUser = Mid(strCurrentUser, InStr(strCurrentUser, "\") + 1)
        End If

        Dim objDW As HPQ.Data.DataWrapper = New HPQ.Data.DataWrapper()
        Dim objComm As SqlCommand

        objComm = objDW.CreateCommand("usp_GetUserType", CommandType.StoredProcedure)

        objDW.CreateParameter(objComm, "@UserName", SqlDbType.VarChar, strCurrentUser, 30)
        objDW.CreateParameter(objComm, "@Domain", SqlDbType.VarChar, strCurrentDomain, 30)

        Dim objDT As DataTable = objDW.ExecuteCommandTable(objComm)

        Try
            If (Not objDT Is Nothing) Then
                objRow = objDT.Rows(0)

                If (IsNumeric(objRow("PartnerTypeID"))) Then
                    intCurrentPartnerTypeID = CInt(objRow("PartnerTypeID"))

                    If (intCurrentPartnerTypeID = 2) Then
                        blnResult = True
                    End If
                End If
            End If

        Catch ex As Exception
        Finally
            If (Not objDT Is Nothing) Then
                objDT.Dispose()
                objDT = Nothing
            End If
            objComm = Nothing
            objDW = Nothing
        End Try


        IsOSSPUser = blnResult
    End Function

    Protected Sub dvSpareKit_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewUpdateEventArgs) Handles dvSpareKit.ItemUpdating

        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim intResult As Integer = 0
        Dim strErrorMsg As String = ""
        
        lblError.Visible = False
        
 

        Dim cbx As CheckBox
        Dim notVersion As Boolean = False
        Dim i As Integer = 0
        Do Until notVersion Or i = dvSpareKit.Rows.Count - 2
            If dvSpareKit.Rows(i).Cells(0).Text.Contains("Category") Then
                notVersion = True
            Else

                i += 1
            End If
        Loop

        Dim dtls As SpareDetail = New SpareDetail()

        dtls.HpPartNo = HpPartNo
        dtls.ServiceFamilyPn = ServiceFamilyPn
        dtls.CategoryName = CType(dvSpareKit.Rows(i).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OsspOrderable = CType(dvSpareKit.Rows(i + 1).Cells(1).Controls(0), CheckBox).Checked
        dtls.OdmPartNo = CType(dvSpareKit.Rows(i + 2).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmPartDesc = CType(dvSpareKit.Rows(i + 3).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmBulkPartNo = CType(dvSpareKit.Rows(i + 4).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmProdMoq = CType(dvSpareKit.Rows(i + 5).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.OdmPostProdMoq = CType(dvSpareKit.Rows(i + 6).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.Supplier = CType(dvSpareKit.Rows(i + 7).Cells(1).Controls(0), TextBox).Text.ToString()
        dtls.Comments = CType(dvSpareKit.Rows(i + 8).Cells(1).Controls(0), TextBox).Text.ToString()

        Try ' Call Update Method
            ' Could add output parameter(s) to the stored procedure called by this method in the future
            intResult = hpqData.UpdateServiceSpareDetail(dtls.ServiceFamilyPn, dtls.HpPartNo, dtls.CategoryName, dtls.OsspOrderable, dtls.OdmPartNo, dtls.OdmPartDesc, dtls.OdmBulkPartNo, dtls.OdmProdMoq, dtls.OdmPostProdMoq, dtls.Comments, dtls.Supplier)
                                       
        Catch ex As Exception
            
            ' Set default "unfriendly" message
            strErrorMsg = ex.Message
            
            Try ' SqlException Cast
                Dim oSQLExc As System.Data.SqlClient.SqlException
               
                oSQLExc = TryCast(ex, System.Data.SqlClient.SqlException)
                
                ' Trap for specific error(s) and create user friendly version of message(s)
                Select Case oSQLExc.Number
                    Case 520
                        strErrorMsg = "No Partner ID Found." 
                    Case 547
                        strErrorMsg = "The ODM Part Number is already in use." 
                End Select
                             
            Catch ex2 As Exception
                ' Resort to possibly inconsistent and/or erroneous approach
                If (strErrorMsg.ToUpper.IndexOf("CC_SERVICESPAREDETAIL_ODMPARTNO") > -1) Or strErrorMsg.ToUpper.IndexOf("CHECK CONSTRAINT") > -1 Then
                    ' Create user friendly version of message
                    strErrorMsg = "The ODM Part Number is already in use." 
                End If
            End Try ' SqlException Cast
            
            lblError.Text = strErrorMsg
            lblError.Visible = True
            
        End Try ' Call Update Method
        
        closeWindow.Value = (lblError.Visible = False).ToString().ToLower()
        Body.Attributes.Add("onLoad", "body_onLoad()")
    End Sub

    Protected Sub dvSpareKit_ItemCanceling(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewCommandEventArgs) Handles dvSpareKit.ItemCommand
        
        closeWindow.Value = "true"
        Body.Attributes.Add("onLoad", "body_onLoad()")
        
    End Sub

    Protected Sub dvSpareKit_ModeChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DetailsViewModeEventArgs) Handles dvSpareKit.ModeChanging
        dvSpareKit.ChangeMode(e.NewMode)
        If Not e.NewMode = DetailsViewMode.Insert Then
            Bind_dvSpareKit()
        End If
    End Sub

    Sub Bind_dvSpareKit()
        Dim hpqData As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()
        Dim dtDetails As DataTable = hpqData.SelectServiceSpareDetails(ServiceFamilyPn, HpPartNo)
        If dtDetails.Rows.Count = 0 Then
            lblError.Text = String.Format("Part Number {0} not found.", HpPartNo)
            lblError.Visible = True
            dvSpareKit.Visible = False
            Exit Sub
        End If
        Dim bSpareKit As Boolean = False
        If Not IsDBNull(dtDetails.Rows(0)("SpareKit")) Then
            bSpareKit = dtDetails.Rows(0)("SpareKit")
        End If
        Dim dt As DataTable = New DataTable()
        Dim sbVersionsSupported As StringBuilder = New StringBuilder()
        If bSpareKit Then
        Else
            dvSpareKit.HeaderText = "Part Details"
        End If
        dt.Columns.Add("Category")
        dt.Columns.Add("OSSP Orderable", GetType(Boolean))
        dt.Columns.Add("Odm Part No")
        dt.Columns.Add("Odm Part Description")
        dt.Columns.Add("Odm Bulk Part No")
        dt.Columns.Add("Odm Production MOQ")
        dt.Columns.Add("Odm Post Production MOQ")
        dt.Columns.Add("Supplier")
        dt.Columns.Add("Comments")

        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("CategoryName"))
        sbVersionsSupported.AppendFormat("|{0}", Convert.ToBoolean(dtDetails.Rows(0)("OsspOrderable")))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPartNo"))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPartDesc"))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmBulkPartNo"))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmProdMoq"))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("OdmPostProdMoq"))
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("Supplier").ToString())
        sbVersionsSupported.AppendFormat("|{0}", dtDetails.Rows(0)("Comments").ToString())
        
        dt.Rows.Add(sbVersionsSupported.Remove(0, 1).ToString().Split("|"))

        If (Me.OSSPUser) Then
            dvSpareKitRO.Visible = True
            dvSpareKitRO.DataSource = dt
            dvSpareKitRO.DataBind()
        Else
            dvSpareKit.Visible = True
            dvSpareKit.DataSource = dt
            dvSpareKit.DataBind()
        End If

    End Sub

#Region " SpareDetail Class "
    Public Class SpareDetail

        Private _Categoryname As String
        Public Property CategoryName() As String
            Get
                Return _Categoryname
            End Get
            Set(ByVal value As String)
                _Categoryname = value
            End Set
        End Property

        Private _OdmPostProdMoq As String
        Public Property OdmPostProdMoq() As String
            Get
                Return _OdmPostProdMoq
            End Get
            Set(ByVal value As String)
                _OdmPostProdMoq = value
            End Set
        End Property


        Private _OdmProdMoq As String
        Public Property OdmProdMoq() As String
            Get
                Return _OdmProdMoq
            End Get
            Set(ByVal value As String)
                _OdmProdMoq = value
            End Set
        End Property

        Private _OdmBulkPartNo As String
        Public Property OdmBulkPartNo() As String
            Get
                Return _OdmBulkPartNo
            End Get
            Set(ByVal value As String)
                _OdmBulkPartNo = value
            End Set
        End Property


        Private _OdmPartDesc As String
        Public Property OdmPartDesc() As String
            Get
                Return _OdmPartDesc
            End Get
            Set(ByVal value As String)
                _OdmPartDesc = value
            End Set
        End Property

        Private _hpPartNo As String
        Public Property HpPartNo() As String
            Get
                Return _hpPartNo
            End Get
            Set(ByVal value As String)
                _hpPartNo = value
            End Set
        End Property


        Private _serviceFamilyPn As String
        Public Property ServiceFamilyPn() As String
            Get
                Return _serviceFamilyPn
            End Get
            Set(ByVal value As String)
                _serviceFamilyPn = value
            End Set
        End Property


        Private _OsspOrderable As Boolean
        Public Property OsspOrderable() As Boolean
            Get
                Return _OsspOrderable
            End Get
            Set(ByVal value As Boolean)
                _OsspOrderable = value
            End Set
        End Property


        Private _OdmPartNo As String
        Public Property OdmPartNo() As String
            Get
                Return _OdmPartNo
            End Get
            Set(ByVal value As String)
                _OdmPartNo = value
            End Set
        End Property
        
        Private _Comments As String
        Public Property Comments() As String
            Get
                Return _Comments
            End Get
            Set(ByVal value As String)
                _Comments = value
            End Set
        End Property
        
        Private _Supplier As String
        Public Property Supplier() As String
            Get
                Return _Supplier
            End Get
            Set(ByVal value As String)
                _Supplier = value
            End Set
        End Property
    End Class
#End Region ' Spare Detail Class


</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        function body_onLoad() {

            if (document.getElementById("closeWindow").value == "true")
                window.parent.close();
        }
    </script>
</head>
<body id="Body" runat="server">
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="lblError" runat="server" Text="Label" Visible="False" ForeColor="Red"
                Font-Bold="True" Font-Overline="False" Font-Size="Medium"></asp:Label>
            <% If Not Me.OSSPUser Then%>
            <asp:DetailsView ID="dvSpareKit" runat="server" CssClass="FormTable" Width="300px"
                AutoGenerateEditButton="True" HeaderText="Spare Kit Details" DefaultMode="Edit">
                <CommandRowStyle HorizontalAlign="Center" />
                <RowStyle HorizontalAlign="Left" Wrap="False" />
                <FieldHeaderStyle Font-Bold="True" Wrap="False" />
                <EditRowStyle Wrap="False" />
                <HeaderStyle Font-Bold="True" Font-Size="Larger" HorizontalAlign="Center" />
            </asp:DetailsView>
            <% Else%>
            <asp:DetailsView ID="dvSpareKitRO" runat="server" CssClass="FormTable" Width="300px"
                AutoGenerateEditButton="False" HeaderText="Spare Kit Details" DefaultMode="ReadOnly">
                <CommandRowStyle HorizontalAlign="Center" />
                <RowStyle HorizontalAlign="Left" Wrap="False" />
                <FieldHeaderStyle Font-Bold="True" Wrap="False" />
                <EditRowStyle Wrap="False" />
                <HeaderStyle Font-Bold="True" Font-Size="Larger" HorizontalAlign="Center" />
            </asp:DetailsView>
            <% End If %>
        <asp:HiddenField ID="closeWindow" runat="server" value="false" />
        </div>
    </form>
</body>
</html>
