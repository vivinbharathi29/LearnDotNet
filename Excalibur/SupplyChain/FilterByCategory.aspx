<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.SupChain_FilterByCategory" Codebehind="FilterByCategory.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        //window.parent.close();
        var pulsarplusDivId = document.getElementById('hdnTabName');
        if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
            // For Closing current popup
            parent.window.parent.closeExternalPopup();
            return false;
        }
        else {
            window.parent.parent.CloseFilterDialog();
            return false;
        }
    }

    function PopupPicker(ctl)
    {
        //var mainEvent = window.event;       
        MyPopUpWin(ctl);
    }

    function MyPopUpWin(ctl) {
        
        var strID;
        
        strID = window.showModalDialog("../mobilese/today/caldraw1.asp", ctl, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) == "undefined")
            return

        document.getElementById(ctl).value = strID;
    }

    var dtCh = "/";
    var minYear = 1900;
    var maxYear = 2100;

    function isInteger(s) {
        var i;
        for (i = 0; i < s.length; i++) {
            // Check that current character is number.
            var c = s.charAt(i);
            if (((c < "0") || (c > "9"))) return false;
        }
        // All characters are numbers.
        return true;
    }

    function stripCharsInBag(s, bag) {
        var i;
        var returnString = "";
        // Search through string's characters one by one.
        // If character is not in bag, append to returnString.
        for (i = 0; i < s.length; i++) {
            var c = s.charAt(i);
            if (bag.indexOf(c) == -1) returnString += c;
        }
        return returnString;
    }

    function daysInFebruary(year) {
        // February has 29 days in any year evenly divisible by four,
        // EXCEPT for centurial years which are not also divisible by 400.
        return (((year % 4 == 0) && ((!(year % 100 == 0)) || (year % 400 == 0))) ? 29 : 28);
    }
    function DaysArray(n) {
        for (var i = 1; i <= n; i++) {
            this[i] = 31
            if (i == 4 || i == 6 || i == 9 || i == 11) { this[i] = 30 }
            if (i == 2) { this[i] = 29 }
        }
        return this
    }

    function isDate(dtStr) {
        var daysInMonth = DaysArray(12)
        var pos1 = dtStr.indexOf(dtCh)
        var pos2 = dtStr.indexOf(dtCh, pos1 + 1)
        var strMonth = dtStr.substring(0, pos1)
        var strDay = dtStr.substring(pos1 + 1, pos2)
        var strYear = dtStr.substring(pos2 + 1)
        strYr = strYear
        if (strDay.charAt(0) == "0" && strDay.length > 1) strDay = strDay.substring(1)
        if (strMonth.charAt(0) == "0" && strMonth.length > 1) strMonth = strMonth.substring(1)
        for (var i = 1; i <= 3; i++) {
            if (strYr.charAt(0) == "0" && strYr.length > 1) strYr = strYr.substring(1)
        }
        month = parseInt(strMonth)
        day = parseInt(strDay)
        year = parseInt(strYr)
        if (pos1 == -1 || pos2 == -1) {
            alert("The date format should be : mm/dd/yyyy")
            return false
        }
        if (strMonth.length < 1 || month < 1 || month > 12) {
            alert("Please enter a valid month")
            return false
        }
        if (strDay.length < 1 || day < 1 || day > 31 || (month == 2 && day > daysInFebruary(year)) || day > daysInMonth[month]) {
            alert("Please enter a valid day")
            return false
        }
        if (strYear.length != 4 || year == 0 || year < minYear || year > maxYear) {
            alert("Please enter a valid 4 digit year between " + minYear + " and " + maxYear)
            return false
        }
        if (dtStr.indexOf(dtCh, pos2 + 1) != -1 || isInteger(stripCharsInBag(dtStr, dtCh)) == false) {
            alert("Please enter a valid date")
            return false
        }
        return true
    }

    function ValidateForm() {
       // alert($("input[type='checkbox']:checked").attr('id'));
       // var n = $('input:checkbox[id^="Release_"]:checked').length;
       // alert(n); // count of checked checkboxes
        var sReleaseIDs = "";
        $('input:checkbox:checked').each(function () {
            if (sReleaseIDs == "")
                sReleaseIDs = $(this).attr("id");
            else
                sReleaseIDs = sReleaseIDs + "," + $(this).attr("id");
        });
        $('#txtReleaseIDs').val(sReleaseIDs);
        var dt = document.getElementById("txtGADateFrom");
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus();
            return false
        }
        dt = document.getElementById("txtGADateTo"); 
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus()
            return false
        }
        dt = document.getElementById("txtSADateFrom"); 
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus()
            return false
        }
        dt = document.getElementById("txtSADateTo"); 
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus()
            return false
        }
        dt = document.getElementById("txtEMDateFrom"); 
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus()
            return false
        }
        dt = document.getElementById("txtEMDateTo"); 
        if (dt.value.length > 0 && isDate(dt.value) == false) {
            dt.focus()
            return false
        }
        return true
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
     <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="form1" name="form1" runat="server" onsubmit="return ValidateForm()">
    <div style="width: 1095px; height: 250px;">
        <asp:Label ID="lblHeader" runat="server" Text="Please Select filters To Display"
                Style="font-size: small; font-weight: bold; font-family: Verdana"></asp:Label>
            <hr />
            <br />
        <table style="width:100%; height:100%"><tr><td style="vertical-align:top">            
            <asp:Label ID="lblCategories" runat="server" Text="Category(s)" Style="font-weight: bold; font-family: Verdana"></asp:Label><br />
            <asp:ListBox ID="lbCategories" runat="server" DataTextField="Name" DataValueField="SCMCategoryID"
                SelectionMode="Multiple" Height="200px" Width="785px" ></asp:ListBox><br /><br />
            <%--<asp:CheckBox ID="chkNoLocalizedAvs" runat="server" Text="No Localizations" TextAlign="Left" Font-Bold="true" />--%>
        </td><td style="vertical-align:top; text-align:left">
            <table border="0">
                <tr><td style="font-weight: bold; font-family: Verdana">General Availability (GA) Date</td></tr>
                <tr><td style="vertical-align:middle">From <asp:TextBox ID="txtGADateFrom" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtGADateFrom');" style="vertical-align:top"/> To <asp:TextBox ID="txtGADateTo" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtGADateTo');" style="vertical-align:top" />
                </td></tr>

                <tr><td style="font-weight: bold; font-family: Verdana; height: 30px; vertical-align:bottom">Select Availability (SA) Date</td></tr>
                <tr><td>From <asp:TextBox ID="txtSADateFrom" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtSADateFrom');" style="vertical-align:top" /> To <asp:TextBox ID="txtSADateTo" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtSADateTo');" style="vertical-align:top" />
                </td></tr>

                <tr><td style="font-weight: bold; font-family: Verdana; height: 30px; vertical-align:bottom">End of Manufacturing (EM) Date</td></tr>
                <tr><td>From <asp:TextBox ID="txtEMDateFrom" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtEMDateFrom');" style="vertical-align:top" /> To <asp:TextBox ID="txtEMDateTo" MaxLength="10" Width="75px" runat="server"></asp:TextBox><img src="../mobilese/today/images/calendar.gif" alt="Choose Date" border="0" width="26" height="23" onclick="javascript:PopupPicker('txtEMDateTo');" style="vertical-align:top" />
                </td></tr>
                <tr><td style="font-family: Verdana; height: 30px; vertical-align:bottom"><b>Releases:</b> <div runat="server" id="divReleases"> </div></td></tr>
            </table>
             </td></tr></table>
        <hr /><br /><br />
        <asp:Button ID="btnDeselect" runat="server" Text="Remove Filters" Style="left: 21px;
            width: 100px; height: 24px; top: 430px;" />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="left: 189px;
            width: 35px; height: 24px; top: 430px;" />
        
        <input type="button" id="btnCancel" value="Cancel" onclick="cmdCancel_onclick();" class="ui-button ui-state-default" />
        <input type="hidden" runat="server" id="txtReleaseIDs" name="txtReleaseIDs" value=""/> 
         <input type="hidden" id="hdnTabName" name="hdnTabName" value="<%= Request("pulsarplusDivId")%>" />
    </div>
    </form>
</body>
</html>
