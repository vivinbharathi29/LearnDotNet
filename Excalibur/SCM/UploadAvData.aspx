<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SCM_UploadAvData" Codebehind="UploadAvData.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        window.parent.close();
    }    
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div style="width: 440px; top: 15px; left: 15px; position: absolute; height: 158px;">
        <asp:Label ID="Label1" runat="server" Text="Select the file to upload and click the Upload button."
            Style="font-family: Verdana; width: 440px; font-size: medium"></asp:Label>
        <br />
        <asp:Label ID="Label2" runat="server" Text="Only include the tabled section including column headers."
            Style="position: absolute; font-family: Verdana; width: 440px; font-size: small;
            text-align: center"></asp:Label>
        <br />
        <br />
        <asp:FileUpload ID="FileUpload1" runat="server" Width="440px" Height="26px" />
        <br />
        <asp:RegularExpressionValidator ID="RegularExpressionValidator1" ControlToValidate="FileUpload1"
            ValidationExpression="([\s\S]+(?=\.(xls$|xlsx$))\.\2)" Display="Dynamic" ErrorMessage="Only Excel XLS or XLSX files allowed."
            EnableClientScript="True" runat="server" Style="font-family: Verdana; font-size: small" />
        <hr />
        <asp:Button ID="submitButton" runat="server" Text="Upload" OnClientClick="inlineProgressBarDiv.style.display='';"
            Style="position: absolute; left: 310px; width: 60px; height: 26px;" />
        <asp:Button ID="cancelButton" runat="server" Text="Cancel" CausesValidation="False"
            OnClientClick="cmdCancel_onclick()" Style="position: absolute; left: 379px; height: 26px;
            width: 60px;" />
        <br />
        <br />
        <pre id="bodyPre" runat="server" style="font-family: Verdana; font-size: small" />
        <div id="inlineProgressBarDiv" style="display: none;">
            <div style="font-family: Verdana; font-size: 12px; position: absolute; z-index: 10;
                position: absolute; left: 101px">
                <table align="center" cellpadding="0" cellspacing="0" style="background-color: Ivory;
                    width: 300px; height: 50px;">
                    <tr>
                        <td width="13px">
                            &nbsp;
                        </td>
                        <td valign="middle" align="center">
                            <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/loading.gif" />
                        </td>
                        <td valign="middle" style="font-family: Verdana; font-size: 12px;">
                            Processing, Please Wait...
                        </td>
                    </tr>
                </table>
            </div>
        </div>
    </div>
    </form>
</body>
</html>
