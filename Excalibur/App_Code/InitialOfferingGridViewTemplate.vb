Imports Microsoft.VisualBasic

Public Class InitialOfferingGridViewTemplate
    Inherits System.Web.UI.Page
    Implements ITemplate
    Private templateType As ListItemType
    Private columnName As String
    Private iPVID As Integer
    Private iBID As Integer
    Private bMarketingIOPlanner As Boolean

    Public Sub New(ByVal type As ListItemType, ByVal colname As String)
        templateType = type
        columnName = colname
    End Sub

    Public Sub New(ByVal type As ListItemType, ByVal colname As String, ByVal PVID As Integer, ByVal BID As Integer, ByVal MarketingIOPlanner As Boolean)
        templateType = type
        columnName = colname
        iPVID = PVID
        iBID = BID
        bMarketingIOPlanner = MarketingIOPlanner
    End Sub

    Public Sub InstantiateIn(ByVal container As System.Web.UI.Control) Implements ITemplate.InstantiateIn
        Dim lc As New Literal()
        Dim lb As New LinkButton()
        Dim ckh As New CheckBox()
        Dim tb1 As New TextBox()

        Select Case templateType
            Case ListItemType.Header
                lc.Text = columnName
                container.Controls.Add(lc)
                Exit Select
            Case ListItemType.Item
                ckh.ID = columnName
                If bMarketingIOPlanner Then
                    ckh.Attributes.Add("onclick", "cbxProduct_onclick(""" & iPVID & """,""" & iBID & """);")
                Else
                    ckh.Enabled = False
                End If
                container.Controls.Add(ckh)
                Exit Select
            Case ListItemType.EditItem
                Dim tb As New TextBox()
                tb.Text = ""
                container.Controls.Add(tb)
                Exit Select
            Case ListItemType.Footer
                lc.Text = "<I>" & columnName & "</I>"
                container.Controls.Add(lc)
                Exit Select
        End Select
    End Sub
End Class


