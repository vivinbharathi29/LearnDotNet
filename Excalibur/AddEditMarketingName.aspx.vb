Imports System.Data
Partial Class AddEditMarketingName
    Inherits System.Web.UI.Page
    Dim dw As HPQ.Excalibur.Data = New HPQ.Excalibur.Data()

    Public ReadOnly Property BID() As String
        Get
            Return Request("BID")
        End Get
    End Property

    Public ReadOnly Property PBID() As String
        Get
            Return Request("PBID")
        End Get
    End Property

    Public ReadOnly Property Name() As String
        Get
            Return Request("Name")
        End Get
    End Property

    Public ReadOnly Property NameType() As String
        Get
            Return Request("NameType")
        End Get
    End Property

    Public ReadOnly Property GeneratedName() As String
        Get
            Return Request("GeneratedName")
        End Get
    End Property

    Public ReadOnly Property Series() As String
        Get
            Return Request("Series")
        End Get
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Me.Page.IsPostBack = False Then
                'NameType                               Name  
                '0 = System Tag, 1 = BIOS Branding      0 = New Name

                lblOldNameHeader.Text = ""
                lblOldNameHeader.Visible = False

                lblOldName.Text = ""
                lblOldName.Visible = False

                TextBox1.Text = ""
                TextBox1.Visible = False

                lblNewNameHeader.Text = ""
                lblNewNameHeader.Visible = False

                TextBox2.Text = ""
                TextBox2.Visible = False

                If NameType <> "2" Then
                    btnAutoGen.Visible = False
                End If
                'lblOldNameHeader.Text = BID & Name & NameType
                'lblOldNameHeader.Visible = True

                If NameType = "0" And Name = "0" Then
                    lblOldNameHeader.Text = "New System Tag Name"
                    lblOldNameHeader.Visible = True
                    TextBox1.Visible = True
                    TextBox1.Focus()
                ElseIf NameType = "0" And Name <> "0" And Name.Trim <> "" Then
                    lblOldNameHeader.Text = "Old System Tag Name"
                    lblOldNameHeader.Visible = True
                    lblOldName.Visible = True
                    lblOldName.Text = Name
                    lblNewNameHeader.Visible = True
                    lblNewNameHeader.Text = "New System Tag Name"
                    TextBox2.Visible = True
                    TextBox2.Focus()
                ElseIf NameType = "1" And Name = "0" Then
                    lblOldNameHeader.Text = "New BIOS Branding Name"
                    lblOldNameHeader.Visible = True
                    TextBox1.Visible = True
                    TextBox1.Focus()
                ElseIf NameType = "1" And Name <> "0" And Name.Trim <> "" Then
                    lblOldNameHeader.Text = "Old BIOS Branding Name"
                    lblOldNameHeader.Visible = True
                    lblOldName.Visible = True
                    lblOldName.Text = Name
                    lblNewNameHeader.Visible = True
                    lblNewNameHeader.Text = "New BIOS Branding Name"
                    TextBox2.Visible = True
                    TextBox2.Focus()
                ElseIf NameType = "2" And Name = "0" Then
                    lblOldNameHeader.Text = "New Logo Badge C Cover Name"
                    lblOldNameHeader.Visible = True
                    TextBox1.Visible = True
                    TextBox1.Focus()
                ElseIf NameType = "2" And Name <> "0" And Name.Trim <> "" Then
                    lblOldNameHeader.Text = "Old Logo Badge C Cover Name"
                    lblOldNameHeader.Visible = True
                    lblOldName.Visible = True
                    lblOldName.Text = Name
                    lblNewNameHeader.Visible = True
                    lblNewNameHeader.Text = "New Logo Badge C Cover Name"
                    TextBox2.Visible = True
                    TextBox2.Focus()
                End If
            End If
        Catch ex As Exception
            lblOldNameHeader.Text = ex.ToString
            lblOldNameHeader.Visible = True
        End Try
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            Dim sNewName As String = ""
            If Name = "0" Then
                'If txtName.Text.Trim = "" Then
                'If NameType <> 2 Then
                '    lblOldNameHeader.ForeColor = Drawing.Color.Red
                '    Exit Sub
                'End If
                'End If
                sNewName = TextBox1.Text.Trim
            ElseIf Name <> "0" Then
                'If txtNewName.Text.Trim = "" Then
                '    If NameType <> 2 Then
                '        lblNewNameHeader.ForeColor = Drawing.Color.Red
                '        Exit Sub
                '    End If
                'End If
                sNewName = TextBox2.Text.Trim
            End If

            If UpdateNaming(sNewName) Then
                Response.Write("<script language='javascript'> {if (window.parent.frames['UpperWindow']) {parent.window.parent.modalDialog.cancel(true);} else {window.close();}}</script>")
            End If
        Catch ex As Exception
            lblOldNameHeader.Text = ex.ToString
            lblOldNameHeader.Visible = True
        End Try
    End Sub

    Protected Sub btnAutoGen_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAutoGen.Click
        Try
            'If UpdateNaming("") Then
            '    Response.Write("<script language='javascript'> { window.close();}</script>")
            'End If
            TextBox2.Text = GeneratedName
        Catch ex As Exception
            lblOldNameHeader.Text = ex.ToString
            lblOldNameHeader.Visible = True
        End Try
    End Sub

    Private Function UpdateNaming(ByVal sNewName As String) As Boolean
        Try
            dw.UpdateAddEditMarketingName(BID, sNewName, NameType, PBID, Series)
            Return True
        Catch ex As Exception
            lblOldNameHeader.Text = ex.ToString
            lblOldNameHeader.Visible = True
        End Try
    End Function
End Class
