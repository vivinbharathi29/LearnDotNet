<%@ Page Language="VB" AutoEventWireup="false" Inherits="DummyVBApp.SupSCM_StructureBOM"
    EnableEventValidation="false" Codebehind="StructureBOM.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
</script>

<script type="text/javascript">
    function cmdCancel_onclick() {
        var pulsarplusDivId = document.getElementById('pulsarplusDivId').value;
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }

    //    function setOptions(chosen) {
    //        var selbox = document.myform.opttwo;

    //        selbox.options.length = 0;
    //        if (chosen == " ") {
    //            selbox.options[selbox.options.length] = new Option(' ', ' ');
    //        }
    //        if (chosen == "1") {
    //            selbox.options[selbox.options.length] = new Option('Images');
    //        }
    //        if (chosen == "2") {
    //            selbox.options[selbox.options.length] = new Option('second choice - option one', 'twoone');
    //            selbox.options[selbox.options.length] = new Option('second choice - option two', 'twotwo');
    //        }
    //        if (chosen == "3") {
    //            selbox.options[selbox.options.length] = new Option('third choice - option one', 'threeone');
    //            selbox.options[selbox.options.length] = new Option('third choice - option two', 'threetwo');
    //        }
    //    }

    function gup(name) {
        name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
        var regexS = "[\\?&]" + name + "=([^&#]*)";
        var regex = new RegExp(regexS);
        var results = regex.exec(window.location.href);
        if (results == null) 
            return "";
        else 
            return results[1]; 
    }

    function setOptions() {
        var ddlSub = document.getElementById('<%=ddlSub.ClientID %>');
        var ddlMain = document.getElementById('<%=ddlMain.ClientID %>');
        var ddlMainValue = ddlMain.options[ddlMain.selectedIndex].value;
        var opt1 = document.createElement("option");
        var opt2 = document.createElement("option");
        var opt3 = document.createElement("option");
        //ddlMain.options.style.backgroundColor = 'white'; 'backgroundColor not compatible in Chrome and not needed
        ddlSub.options.length = 0;

        if (ddlMainValue == 0) {
            opt1.text = "";
            opt1.value = "0";
            ddlSub.options.add(opt1);

            var txtItemNumber = document.getElementById('<%=txtItemNumber.ClientID %>');
            txtItemNumber.value = "";
        }
        else if (ddlMainValue == 1) {
            //03-28-16 Malichi (PBI 18679)(Remove Image SA drop down value and change COA SA to DPK SA)
            opt1.text = "Doc Kit SA";
            opt1.value = "1";
            opt2.text = "DPK SA";
            opt2.value = "3";

            //Order Below Alphabetically
            ddlSub.options.add(opt1);
            ddlSub.options.add(opt2);

            var txtItemNumber = document.getElementById('<%=txtItemNumber.ClientID %>');
            txtItemNumber.value = "2";
        }
        else if (ddlMainValue == 2) {
            opt1.text = "AC Adapter SA";
            opt1.value = "4";
            //Order Below Alphabetically
            ddlSub.options.add(opt1);

            var BusinessID = gup('BusinessID');
            //alert(BusinessID);

            if (BusinessID == 1) {
                opt2.text = "Power Cord SA";
                opt2.value = "5";

                //Order Below Alphabetically 
                ddlSub.options.add(opt2);
            }

            var txtItemNumber = document.getElementById('<%=txtItemNumber.ClientID %>');
            txtItemNumber.value = "4";

        }
        else if (ddlMainValue == 3) {
            var BusinessID = gup('BusinessID');
            //alert(BusinessID);

            if (BusinessID == 2) {
                opt1.text = "Keyboard SA";
                opt1.value = "6";
                opt2.text = "Power Cord SA";
                opt2.value = "5";

                //Order Below Alphabetically 
                ddlSub.options.add(opt1);
                ddlSub.options.add(opt2);
            }
            else {
                opt1.text = "Keyboard SA";
                opt1.value = "6";

                ddlSub.options.add(opt1);
            }

            var txtItemNumber = document.getElementById('<%=txtItemNumber.ClientID %>');
            txtItemNumber.value = "6";
        }
        //setDefaultItemNumber(ddlSubValue);
    }

    function setDefaultItemNumber() {
        var ddlSub = document.getElementById('<%=ddlSub.ClientID %>');
        var ddlSubValue = ddlSub.options[ddlSub.selectedIndex].value;
        var txtItemNumber = document.getElementById('<%=txtItemNumber.ClientID %>');
        if (ddlSubValue == 0) {
            txtItemNumber.value = "";
        }
        if (ddlSubValue == 1) { //Doc Kit SA
            txtItemNumber.value = "2";
        }
        if (ddlSubValue == 2) { //Image SA
            txtItemNumber.value = "1"; 
        }
        if (ddlSubValue == 3) { //DPK SA
            txtItemNumber.value = "3";
        }
        if (ddlSubValue == 4) { //AC Adapter SA
            txtItemNumber.value = "4";
        }
        if (ddlSubValue == 5) { //Power Cord SA
            txtItemNumber.value = "5";
        }
        if (ddlSubValue == 6) { //Keyboard SA
            txtItemNumber.value = "6";
        }
    }
    
    function btnSubmit_Click() {
        var btnSubmit = document.getElementById('<%=btnSubmit.ClientID %>');
        var btnCancel = document.getElementById('<%=btnCancel.ClientID %>');
        btnSubmit.disabled = true;
        btnCancel.disabled = true;
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />
</head>
<body runat="server" id="thisBody">
    <form id="myform" runat="server">
    <textarea id="holdtext" style="display: none;" rows="0" cols="0"></textarea>
    <div style="width: 333px; height: 120px;">
        <asp:Label ID="lblHeader2" runat="server" Text="Item No." Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute;
            top: 111px; left: 213px; height: 19px; width: 65px;"></asp:Label>
        <asp:Label ID="lblHeader3" runat="server" Text="Subassembly" Style="font-size: small; font-weight: bold;
            font-family: Verdana; text-align: center; position: absolute; top: 43px; left: 222px;
            height: 17px; width: 102px;"></asp:Label>
        <asp:Label ID="lblHeader1" runat="server" Text="AV" Style="font-size: small; font-weight: bold;
            font-family: Verdana; text-align: center; position: absolute; top: 45px; left: 53px;
            height: 17px; width: 53px;"></asp:Label>
        <asp:Label ID="lblHeader" runat="server" Text="Structure..." 
            Style="font-size: small;
            font-weight: bold; font-family: Verdana; text-align: center; position: absolute; top: 15px; left: 10px; width: 335px;"></asp:Label>
        <br />
        <asp:DropDownList ID="ddlMain" runat="server" Style="position: absolute; left: 17px;
            height: 23px; width: 123px; top: 67px; right: 1227px;" 
            AutoPostBack="false" OnChange="setOptions()">
              <asp:ListItem Value="0" Text=""></asp:ListItem>
            <asp:ListItem Value="1">Image AV</asp:ListItem>
            <asp:ListItem Value="2">HW Kit AV</asp:ListItem>
            <asp:ListItem Value="3">Keyboard AV</asp:ListItem>
        </asp:DropDownList>
        <asp:TextBox ID="txtItemNumber" runat="server" Style="position: absolute; left: 308px;
            height: 16px; width: 20px; top: 112px; right: 1039px; text-align: center" 
            MaxLength="1">
        </asp:TextBox>
        <asp:DropDownList ID="ddlSub" runat="server" Style="position: absolute; left: 212px;
            height: 23px; width: 123px; top: 66px; margin-bottom: 0px;" 
            AutoPostBack="false" OnChange="setDefaultItemNumber()">
        </asp:DropDownList>
        <!-- <select name="optone" size="1" style="position: absolute; left: 17px; height: 23px;
            width: 100px; top: 45px; right: 959px;" onchange="setOptions(document.myform.optone.options[document.myform.optone.selectedIndex].value);">
            <option value=" " selected="selected"></option>
            <option value="1">Doc Kit</option>
        </select>
        <select name="opttwo" size="1" style="position: absolute; left: 174px; height: 23px;
            width: 100px; top: 45px; margin-bottom: 0px;">
            <option value=" " selected="selected"></option>
        </select> -->
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <br />
        <hr />
        <asp:Button ID="btnSubmit" runat="server" Text="OK" Style="position: absolute; left: 229px;
            width: 35px; height: 24px; top: 161px;"/>
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 274px; width: 61px; height: 24px; top: 161px;" 
            OnClientClick="cmdCancel_onclick()" />
        <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
    </div>
    </form>
</body>
</html>
