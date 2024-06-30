<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dl As HPQ.Excalibur.Data = New HPQ.Excalibur.Data
        
        For Each item As ListItem In cblImagesWithOut.Items
            If item.Selected Then
                dl.LeverageImageWhqlStatus(item.Value, ddlImagesWith.SelectedValue)
            End If
        Next
        
        cblImagesWithOut.ClearSelection()
        cblImagesWithOut.DataBind()
        ddlImagesWith.DataBind()

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Page.ClientScript.RegisterStartupScript(Me.GetType, "_CloseWindow", "function CloseWindow() { window.parent.close(); }", True)
        PageBody.Attributes("onload") = "CloseWindow()"

    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Leverage Whql Image Status</title>
    <link href="../style/Excalibur.css" rel="stylesheet" type="text/css" />
    <style>
        fieldset
        {
            display: block;
            width: 400px;
            padding: 10px;
        }
    </style>
    <meta http-equiv="CACHE-CONTROL" content="NO-CACHE">
    <meta http-equiv="PRAGMA" content="NO-CACHE">
</head>
<body id="PageBody" runat="server">
    <form id="form1" runat="server">
        <div>
            <fieldset>
                <legend>Images With Logo Results</legend>
                <asp:DropDownList ID="ddlImagesWith" runat="server" AutoPostBack="True" DataSourceID="odsImagesWithWhql"
                    DataTextField="ListItems" DataValueField="SkuNumber">
                </asp:DropDownList><asp:ObjectDataSource ID="odsImagesWithWhql" runat="server" OldValuesParameterFormatString="original_{0}"
                    SelectMethod="ListImagesWithWhqlStatus" TypeName="HPQ.Excalibur.Data">
                    <SelectParameters>
                        <asp:QueryStringParameter Name="ProductVersionID" QueryStringField="PVID" Type="String" />
                    </SelectParameters>
                </asp:ObjectDataSource>
            </fieldset>
            <fieldset>
                <legend>Images Needing Logo Results</legend>
                <div style="height: 400px; width: 400px; overflow-y: scroll;">
                    <asp:CheckBoxList ID="cblImagesWithOut" runat="server" DataSourceID="odsImagesWithoutWhql"
                        DataTextField="ListItems" DataValueField="SkuNumber">
                    </asp:CheckBoxList><asp:ObjectDataSource ID="odsImagesWithoutWhql" runat="server"
                        OldValuesParameterFormatString="original_{0}" SelectMethod="ListImagesWithOutWhqlStatus"
                        TypeName="HPQ.Excalibur.Data">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="ddlImagesWith" Name="ImageSkuNumber" PropertyName="SelectedValue"
                                Type="String" />
                        </SelectParameters>
                    </asp:ObjectDataSource>
                </div>
            </fieldset>
            <br />
            <div style="width: 420px; text-align: right">
                <asp:Button ID="Button1" runat="server" Text="Save" OnClick="Button1_Click" />&nbsp;
                <asp:Button ID="Button2" runat="server" Text="Close" OnClick="Button2_Click" />
            </div>
        </div>
    </form>
</body>
</html>
