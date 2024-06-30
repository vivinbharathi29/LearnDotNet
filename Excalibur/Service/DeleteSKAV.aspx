<%@ Page Language="vb" AutoEventWireup="false" Inherits="DummyVBApp.service_DeleteSKAV" EnableEventValidation="False" EnableViewState="true" Codebehind="DeleteSKAV.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <title></title>
 <script language="javascript" type="text/javascript">
     function confirmExit(oSender) {
       //  window.returnValue = "cancel";
         //  window.opener.location.refresh();
        // window.opener.atest()
     }

     function confirm_delete() {       
            if (confirm("Are you sure you want to delete these items(s)?") == true) {
                var getHtml = service_DeleteSKAV.delComfirmed(one, two);
                return true;
              
            } else {           
               return false;
           }
       }
       function reloadParentAndClose() {
           var f = window.opener.top.frames;
           for (var i = f.length - 1; i > -1; --i)
               f[i].location.reload();
           window.parent.close();     
       }
 </script>
</head>
<body id="body" runat="server" onbeforeunload="return confirmExit(null);" >

 <form id="form1" runat="server">
 Search AV#: <asp:TextBox ID="avtext" runat="server"></asp:TextBox>
 <asp:Button ID="Button2" runat="server" Text="Submit" />

 <div id = "displayresults" runat="server">
  <table width = "850" border="0">
  <tr>
        <td width = "50%" align = "center">
         <p>
   <asp:Button ID="CheckAll" runat="server" Text="Check All" />
   &nbsp;
   <asp:Button ID="UncheckAll" runat="server" Text="Uncheck All" />
   &nbsp;
  
   <asp:Button AccessKey="s" ID="DeleteButton" Text="Remove selected" runat="server" OnClick="DeleteButton_Click" OnClientClick="confirm_delete(); return false" />
</p>
  <div>
    <asp:GridView ID="FileList" runat="server"
    AutoGenerateColumns="False" DataKeyNames="sskmav_id" BorderWidth="0" Width="600" HorizontalAlign="Left"  HeaderStyle-BackColor="Khaki" AlternatingRowStyle-BackColor="Honeydew" >
   
    <Columns>
    
        <asp:TemplateField ItemStyle-HorizontalAlign="Left">
            <ItemTemplate>
                <asp:CheckBox runat="server" ID="RowLevelCheckBox" />
            </ItemTemplate>
   
        </asp:TemplateField >
        <asp:BoundField DataField="sparekitno" HeaderText="Spare Kit Number"  ItemStyle-HorizontalAlign="Left"  />
        <asp:BoundField DataField="Description" HeaderText="GPG Description" ItemStyle-HorizontalAlign="left"/>
        <asp:BoundField DataField="CategoryName" HeaderText="Description" ItemStyle-HorizontalAlign="left"> 
            <ItemStyle HorizontalAlign="left" />
        </asp:BoundField>      
         <asp:BoundField DataField="AVno" HeaderText="AV" ItemStyle-HorizontalAlign="left"/>
    </Columns>
</asp:GridView>
      
    </div>

    
        
        </td>
        <td valign="top">
          <br /> <br />
          <center>
            <asp:Button ID="Button1" runat="server" Text="Close Window"  Visible="false" 
                  onclick="Button1_Click"/></center>
          <br />
    <asp:label ID="Summary" runat="server" text=""></asp:label>

        </td> 
  </tr>

  </table>
 </div><br /><br />
 &nbsp   &nbsp   &nbsp  <asp:label ID="Label1" runat="server" text=""></asp:label>
    </form>
  
</body>

</html>
