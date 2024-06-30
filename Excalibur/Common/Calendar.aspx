<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.Common_Calendar" Codebehind="Calendar.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Calendar</title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" /> 
    <script type="text/javascript" language="javascript">
        function SetDate(dateValue) {
            // retrieve from the querystring the value of the Ctl param,
            // that is the name of the input control on the parent form
            // that the user want to set with the clicked date
            //ctl = window.location.search.substr(1).substring(4);
            ctl = document.getElementById("ctl").value;
            // set the value of that control with the passed date
            thisForm = window.opener.document.forms[0].elements[ctl].value =
               dateValue;
            // close this popup
            self.close();
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Calendar ID="Calendar1" runat="server"  OnDayRender="Calendar1_DayRender">
        </asp:Calendar>
        <asp:HiddenField ID="ctl" runat="server" />
    </form>
</body>
</html>
