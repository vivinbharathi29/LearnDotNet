<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>

<script type="text/javascript">
function copyTableToClipboard()
{
    var tableDiv = document.getElementById('tableData');
    var holdText = document.getElementById('holdtext');
    holdText.innerText = '<html><body>' + tableDiv.innerHTML + '</body></html>';
    var Copied = holdText.createTextRange();
    Copied.execCommand("RemoveFormat");
    Copied.execCommand("Copy");
    //alert(tableDiv.innerHTML);
}
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Untitled Page</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
        <textarea id="holdtext" style="display: none;" rows="0" cols="0"></textarea>
        <div>
            <p>
                <asp:Label ID="Label1" runat="server" Text="Orphaned AVs" Style="font-size: large;
                    font-weight: bold;"></asp:Label></p>
<!--            <p>
                <span style="color: blue; font: verdana bold xx-small; text-decoration: underline;
                    cursor: pointer;" onclick="copyTableToClipboard()">Copy to Clipboard</span></p> -->
            <div id="tableData">
                <asp:GridView ID="GridView1" runat="server" DataSourceID="ObjectDataSource1" CellPadding="4"
                    ForeColor="#333333" GridLines="None" EmptyDataText="No Orphans Found">
                </asp:GridView>
            </div>
            <asp:ObjectDataSource ID="ObjectDataSource1" runat="server" OldValuesParameterFormatString="original_{0}"
                SelectMethod="SelectAvsNotInKmatBom" TypeName="HPQ.Excalibur.Data">
                <SelectParameters>
                    <asp:QueryStringParameter Name="ProductBrandId" QueryStringField="BID" Type="String" />
                </SelectParameters>
            </asp:ObjectDataSource>
        </div>
    </form>
</body>
</html>
