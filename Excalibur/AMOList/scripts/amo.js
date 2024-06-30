function row_onmouseover(evt) {
    var objUnknown, id;
    if (!evt) evt = window.event;
    objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    id = objUnknown.id;

    if ((objUnknown.tagName.toUpperCase() == "TD") && (id != 'sd')) {
        //textdecoration = objUnknown.style.textDecoration;
        objUnknown.style.cursor = 'pointer';	// hand pointer for both Firefox and IE
        //objUnknown.style.textDecoration = objUnknown.style.textDecoration + ' underline';
    }
    return true;
}

function row_onmouseout(evt) {
    var objUnknown, id;
    if (!evt) evt = window.event;
    objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    id = objUnknown.id;

    if ((objUnknown.tagName.toUpperCase() == "TD") && (id != 'sd')) {
        objUnknown.style.cursor = 'auto';
        //objUnknown.style.textDecoration = textdecoration;
    }
    return true;
}

function btnAdd_onClick() {
    OpenCreateAMOFeatureProperties();
}

function checkEnter(theitem, e) {
    if (!e) e = window.event;
    var charCode = e.keyCode ? e.keyCode : e.which;
    if (charCode == 13) {
        theitem.blur()
        return false;
    }
}

function editText(evtobj, ModuleID, Field, intMaxlength, repository, fullname) {
    var sHTML
    if (evtobj.innerHTML.indexOf("editcell" + ModuleID + Field) < 0) {
        evtobj.style.textDecoration = "none"; // get rid of underline

        // have to replace quotation marks too and escape it and unescape it in the getText function
        // so single and double quotes work		

        var sTempValue = evtobj.innerHTML.replace(/&nbsp;/g, "").replace(/"/g, "&quot;")

        sHTML = "<input onKeyPress='return checkEnter(this, event)' "
        sHTML += "onBlur='javascript:getText(event," + ModuleID + ",\"" + Field + "\", \"" + escape(sTempValue) + "\", \"" + repository + "\",\"" + fullname + "\");' "
        sHTML += "type=text maxlength=" + intMaxlength + " size=" + intMaxlength + " value=\"" + sTempValue + "\" "
        sHTML += "id='editcell" + ModuleID + Field + "' NAME='editcell" + ModuleID + Field + "'>";

        evtobj.innerHTML = sHTML

        document.getElementById("editcell" + ModuleID + Field).focus()
    }
}

function editOption(evtobj, ModuleID, Field, intMaxlength, repository, fullname) {
    var sHTML
    if (evtobj.innerHTML.indexOf("optioncell" + ModuleID + Field) < 0) {
        evtobj.style.textDecoration = "none"; // get rid of underline

        // have to replace quotation marks too and escape it and unescape it in the getOption function
        // so single and double quotes work		
        var sTempValue = evtobj.innerHTML.replace(/&nbsp;/g, "").replace(/"/g, "&quot;")

        sHTML = "<select id='optioncell" + ModuleID + Field + "' name='optioncell" + ModuleID + Field + "' ";
        sHTML += "onBlur='javascript:getOption(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")' ";
        sHTML += "onchange='javascript:getOption(event," + ModuleID + ",\"" + Field + "\", \"" + sTempValue + "\", \"" + repository + "\",\"" + fullname + "\")'>";
        if (sTempValue == "Global") {
            sHTML += "<option value=0> </option>"
            sHTML += "<option value=1 selected>Global</option>"
            sHTML += "<option value=2>Blind</option>"
            sHTML += "<option value=3>Full</option></select>"
        } else if (sTempValue == "Blind") {
            sHTML += "<option value=0> </option>"
            sHTML += "<option value=1 >Global</option>"
            sHTML += "<option value=2 selected>Blind</option>"
            sHTML += "<option value=3>Full</option></select>"
        } else if (sTempValue == "Full") {
            sHTML += "<option value=0> </option>"
            sHTML += "<option value=1>Global</option>"
            sHTML += "<option value=2>Blind</option>"
            sHTML += "<option value=3 selected>Full</option></select>"
        } else {
            sHTML += "<option value=0 selected> </option>"
            sHTML += "<option value=1 >Global</option>"
            sHTML += "<option value=2>Blind</option>"
            sHTML += "<option value=3>Full</option></select>"
        }

        evtobj.innerHTML = sHTML
        document.getElementById("optioncell" + ModuleID + Field).focus()
    }
}

function editDate(evtobj, ModuleID, RegionID, Field, repository, fullname) {
    var sHTML
    if (evtobj.innerHTML.indexOf("editcell" + ModuleID + Field) < 0) {
        evtobj.style.textDecoration = "none"; // get rid of underline

        sHTML = "<input onKeyPress='return checkEnter(this, event)' "
        sHTML += "onBlur='javascript:getDateValue(event," + ModuleID + "," + RegionID + ",\"" + Field + "\", \"" + evtobj.innerHTML.replace(/&nbsp;/g, "") + "\", \"" + repository + "\",\"" + fullname + "\")' "
        sHTML += "type=text maxlength=10 size=10 value=\"" + evtobj.innerHTML.replace(/&nbsp;/g, "") + "\" "
        sHTML += "id='editcell" + ModuleID + Field + "' NAME='editcell" + ModuleID + Field + "'>";

        evtobj.innerHTML = sHTML
        document.getElementById("editcell" + ModuleID + Field).focus()
    }
}

function editCurrency(evtobj, ModuleID, Field, intMaxlength, repository, fullname) {
    var sHTML
    if (evtobj.innerHTML.indexOf("editcell" + ModuleID + Field) < 0) {
        evtobj.style.textDecoration = "none"; // get rid of underline

        sHTML = "<input onKeyPress=\"return(currencyFormat(this, ',', '.', event, 20))\" "
        sHTML += "onBlur='javascript:getCurrency(event," + ModuleID + ",\"" + Field + "\", \"" + evtobj.innerHTML.replace(/&nbsp;/g, "") + "\", \"" + repository + "\",\"" + fullname + "\")' "
        sHTML += "type=text maxlength=" + intMaxlength + " size=" + intMaxlength + " value=\"" + evtobj.innerHTML.replace(/&nbsp;/g, "") + "\" "
        sHTML += "id='editcell" + ModuleID + Field + "' NAME='editcell" + ModuleID + Field + "'>";

        evtobj.innerHTML = sHTML
        document.getElementById("editcell" + ModuleID + Field).focus()
    }
}

// From http://javascript.internet.com/forms/auto-currency.html?start_comment=0
// Auto Currency. Inserts the proper seperators to automatically format any currency field while typing.
// Added maxlen to limit the length of the characters typed.
// Call like: <input type=text name=test length=15 onKeyPress="return(currencyFormat(this, ',', '.', event, 20))">
function currencyFormat(fld, milSep, decSep, e, maxlen) {
    var sep = 0;
    var key = '';
    var i = j = 0;
    var len = len2 = 0;
    var strCheck = '0123456789';
    var aux = aux2 = '';
    if (!e) e = window.event;
    var whichCode = e.which ? e.which : e.keyCode;

    if (whichCode == 13) { // Enter
        fld.blur()
        return false;
    }
    if (whichCode == 8) return true;  // Delete
    key = String.fromCharCode(whichCode);  // Get key value from key code
    if (strCheck.indexOf(key) == -1) return false;  // Not a valid key
    len = fld.value.length;
    if (len >= maxlen)
        return false
    for (i = 0; i < len; i++)
        if ((fld.value.charAt(i) != '0') && (fld.value.charAt(i) != decSep)) break;

    aux = '';
    for (; i < len; i++)
        if (strCheck.indexOf(fld.value.charAt(i)) != -1) aux += fld.value.charAt(i);

    aux += key;
    len = aux.length;
    if (len == 0) fld.value = '';
    if (len == 1) fld.value = '0' + decSep + '0' + aux;
    if (len == 2) fld.value = '0' + decSep + aux;
    if (len > 2) {
        aux2 = '';
        for (j = 0, i = len - 3; i >= 0; i--) {
            if (j == 3) {
                aux2 += milSep;
                j = 0;
            }
            aux2 += aux.charAt(i);
            j++;
        }
        fld.value = '';
        len2 = aux2.length;
        for (i = len2 - 1; i >= 0; i--)
            fld.value += aux2.charAt(i);
        fld.value += decSep + aux.substr(len - 2, len);
    }
    return false;
}

// From: http://javascript.internet.com/forms/currency-format.html
// This script accepts a number or string and formats it like U.S. currency. 
// Adds the dollar sign, rounds to two places past the decimal, adds place holding zeros, 
// and commas where appropriate. Occurs when the user clicks the button or when they finish 
// entering the money amount (and click into the next field).
function formatCurrency(num) {
    num = num.toString().replace(/\$|\,/g, '');
    if (isNaN(num))
        num = "0";
    sign = (num == (num = Math.abs(num)));
    num = Math.floor(num * 100 + 0.50000000001);
    cents = num % 100;
    num = Math.floor(num / 100).toString();
    if (cents < 10)
        cents = "0" + cents;
    for (var i = 0; i < Math.floor((num.length - (1 + i)) / 3) ; i++)
        num = num.substring(0, num.length - (4 * i + 3)) + ',' +
	num.substring(num.length - (4 * i + 3));
    return (((sign) ? '' : '-') + num + '.' + cents);
}

/*****************************************************************
//Function:     ChangeRAS_GPSy();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getCurrency(evt, ModuleID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement
    Parent.style.textDecoration = "underline"; // add underline back

    var ajaxurl = "";
    var errormsg = "";

    if (objUnknown.id == "editcell" + ModuleID + Field) {
        var NewValue = objUnknown.value;
        if (OldValue == NewValue) { // nothing changed
            if (NewValue == '') {
                Parent.innerHTML = '&nbsp;';
            } else {
                Parent.innerHTML = objUnknown.value;
            }
        } else { // something changed, save the data
            //var objRS = RSGetASPObject("AMO_RS.asp");

            ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=" + Field + "&Value=" + formatCurrency(NewValue) + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
            //var objResult = objRS.setFieldValue(repository, ModuleID, Field, escape(formatCurrency(NewValue)), fullname);

            $.ajax({
                url: ajaxurl,
                type: "GET",
                async: false,
                success: function (data) {
                    errormsg = data;
                },
                error: function (xhr, status, error) {
                    errormsg = error;
                    erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                }
            })

            if (errormsg == "success") {
                // highlight the field
                Parent.className = "clsAMO_ChangedCell";
                if (NewValue == '') {
                    Parent.innerHTML = '&nbsp;';
                } else {
                    Parent.innerHTML = formatCurrency(NewValue);
                }

                if (Field == 'amocost' && NewValue != '') {
                    // need to update AMO Price field now too by multiplying by 2
                    var wwpObject = document.getElementById("wwp" + ModuleID);
                    var testvalue = formatCurrency(stripCharsInBag(NewValue, ',') * 2);
                    if (testvalue.length <= 20) {
                        var wwpPrice = testvalue;
                        //objRS = RSGetASPObject("AMO_RS.asp");

                        ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=amowwprice&Value=" + wwpPrice + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                        //objResult = objRS.setFieldValue(repository, ModuleID, 'amowwprice', escape(wwpPrice), fullname);

                        $.ajax({
                            url: ajaxurl,
                            type: "GET",
                            async: false,
                            success: function (data) {
                                errormsg = data;
                            },
                            error: function (xhr, status, error) {
                                errormsg = error;
                                erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                            }
                        })


                        if (errormsg == "success") {
                            // highlight the field
                            if (wwpObject != null) {
                                wwpObject.className = "clsAMO_ChangedCell";
                                if (wwpPrice == '') {
                                    wwpObject.innerHTML = "&nbsp;";
                                } else {
                                    wwpObject.innerHTML = formatCurrency(wwpPrice);
                                }
                            }
                        }
                    }
                }
                // update status to 'In Process'
                var asObject = document.getElementById("as" + ModuleID);
                if ((asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled")
                        && (Field == "amocost" || Field == "amowwprice")) {
                    asObject.innerHTML = "In Process";
                }
            }
        }
    }
}

/*****************************************************************
//Function:     getOption();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getOption(evt, ModuleID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement;
    var intOldValue, NewValue;
    var asObject;
    var objRS;
    var objResult;
    var ajaxurl = "";
    var errormsg = "";

    var intNewValue = objUnknown.value;

    Parent.style.textDecoration = "underline"; // add underline back

    if (OldValue == "Global") intOldValue = 1;
    else if (OldValue == "Blind") intOldValue = 2;
    else if (OldValue == "Full") intOldValue = 3;
    else intOldValue = 0;

    if (intNewValue == 1) NewValue = "Global";
    else if (intNewValue == 2) NewValue = "Blind";
    else if (intNewValue == 3) NewValue = "Full";
    else NewValue = "&nbsp;";

    if (intOldValue == intNewValue) { // nothing changed
        if (OldValue == "")
            OldValue = "&nbsp;";
        Parent.innerHTML = OldValue;
    } else {
        // something changed, save the data
        //objRS = RSGetASPObject("AMO_RS.asp");	

        ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=" + Field + "&Value=" + intNewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
        //objResult = objRS.setFieldValue(repository, ModuleID, Field, intNewValue, fullname);

        $.ajax({
            url: ajaxurl,
            type: "GET",
            async: false,
            success: function (data) {
                errormsg = data;
            },
            error: function (xhr, status, error) {
                errormsg = error;
                erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
            }
        })

        if (errormsg == "success") {
            // highlight the field
            Parent.className = "clsAMO_ChangedCell";
            Parent.innerHTML = NewValue;
        }

        if (intNewValue == 1) {
            if (Field != "Visibility_NA") {
                asObject = document.getElementById("vna" + ModuleID);
                if (asObject.innerHTML != "Global") {
                    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=Visibility_NA&Value=1" + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                    //objResult = objRS.setFieldValue(repository, ModuleID, "Visibility_NA", 1, fullname);

                    $.ajax({
                        url: ajaxurl,
                        type: "GET",
                        async: false,
                        success: function (data) {
                            errormsg = data;
                        },
                        error: function (xhr, status, error) {
                            errormsg = error;
                            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                        }
                    })

                    asObject.innerHTML = "Global";
                    asObject.className = "clsAMO_ChangedCell";
                }
            }
            if (Field != "Visibility_AP") {
                asObject = document.getElementById("vap" + ModuleID);
                if (asObject.innerHTML != "Global") {
                    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=Visibility_AP&Value=1" + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                    //objResult = objRS.setFieldValue(repository, ModuleID, "Visibility_AP", 1, fullname);

                    $.ajax({
                        url: ajaxurl,
                        type: "GET",
                        async: false,
                        success: function (data) {
                            errormsg = data;
                        },
                        error: function (xhr, status, error) {
                            errormsg = error;
                            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                        }
                    })

                    asObject.innerHTML = "Global";
                    asObject.className = "clsAMO_ChangedCell";
                }
            }
            if (Field != "Visibility_EM") {
                asObject = document.getElementById("vem" + ModuleID);
                if (asObject.innerHTML != "Global") {
                    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=Visibility_EM&Value=1" + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                    //objResult = objRS.setFieldValue(repository, ModuleID, "Visibility_EM", 1, fullname);

                    $.ajax({
                        url: ajaxurl,
                        type: "GET",
                        async: false,
                        success: function (data) {
                            errormsg = data;
                        },
                        error: function (xhr, status, error) {
                            errormsg = error;
                            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                        }
                    })

                    asObject.innerHTML = "Global";
                    asObject.className = "clsAMO_ChangedCell";
                }
            }
            if (Field != "Visibility_LA") {
                asObject = document.getElementById("vla" + ModuleID);
                if (asObject.innerHTML != "Global") {
                    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=Visibility_LA&Value=1" + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                    //objResult = objRS.setFieldValue(repository, ModuleID, "Visibility_LA", 1, fullname);

                    $.ajax({
                        url: ajaxurl,
                        type: "GET",
                        async: false,
                        success: function (data) {
                            errormsg = data;
                        },
                        error: function (xhr, status, error) {
                            errormsg = error;
                            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                        }
                    })

                    asObject.innerHTML = "Global";
                    asObject.className = "clsAMO_ChangedCell";
                }
            }
        }

        // update status to 'In Process'
        asObject = document.getElementById("as" + ModuleID);
        if (asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled") {
            //asObject.innerHTML = "In Process";
        }
    }
}

/*****************************************************************
//Function:     getProductOption();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getProductOption(evt, ModuleID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement;
    var NewValue;
    var asObject;
    var objRS;
    var objResult;
    var ajaxurl = "";
    var errormsg = "";

    //var cell = Parent.closest('td');
    //var iCellIndex = cell.cellIndex;
    //var iRowIndex = cell.parentNode.rowIndex;

    NewValue = objUnknown.value;

    //alert(NewValue);

    Parent.style.textDecoration = "underline"; // add underline back

    if (OldValue == NewValue) { // nothing changed
        Parent.innerHTML = OldValue;
    } else {  //something changed
        if (NewValue == "") {
            NewValue = "&nbsp;";
            Parent.innerHTML = NewValue;
        } else {
            // something changed, save the data
            //objRS = RSGetASPObject("AMO_RS.asp");	

            ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=" + Field + "&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
            //objResult = objRS.setFieldValue(repository, ModuleID, Field, NewValue, fullname);	

            $.ajax({
                url: ajaxurl,
                type: "GET",
                async: false,
                success: function (data) {
                    errormsg = data;
                },
                error: function (xhr, status, error) {
                    errormsg = error;
                    erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                }
            })

            if (errormsg == "success") {
                //highlight the field
                Parent.className = "clsAMO_ChangedCell";
                if (NewValue == "") {
                    NewValue = "&nbsp;";
                }
                Parent.innerHTML = NewValue;
            }

            // update status to 'In Process'
            asObject = document.getElementById("as" + ModuleID);
            if (asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled" && asObject.innerHTML != "Complete") {
                asObject.innerHTML = "In Process";
            }
        }
    }
}

/*****************************************************************
//Function:     getSerialFlag();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getSerialFlag(evt, ModuleID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement;
    var NewValue;
    var asObject;
    var objRS;
    var objResult;
    var ajaxurl = "";
    var errormsg = "";

    NewValue = objUnknown.value;

    Parent.style.textDecoration = "underline";
    OldValue = unescape(OldValue)
    if (OldValue == NewValue) {
        if (OldValue == "")
            OldValue = "&nbsp;";
        Parent.innerHTML = OldValue;
    }
    else {
        // something changed, save the data
        //objRS = RSGetASPObject("AMO_RS.asp");        

        ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=" + Field + "&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
        //objResult = objRS.setFieldValue(repository, ModuleID, Field, NewValue, fullname);

        $.ajax({
            url: ajaxurl,
            type: "GET",
            async: false,
            success: function (data) {
                errormsg = data;
            },
            error: function (xhr, status, error) {
                errormsg = error;
                erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
            }
        })

        if (errormsg == "success") {
            //highlight the field
            Parent.className = "clsAMO_ChangedCell";
            if (NewValue == "") {
                NewValue = "&nbsp;";
            }
            Parent.innerHTML = NewValue;
        }
    }

}

/*****************************************************************
//Function:     getText();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getText(evt, ModuleID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement
    var bchangestatus
    Parent.style.textDecoration = "underline"; // add underline back

    var ajaxurl = "";
    var errormsg = "";

    OldValue = unescape(OldValue)
   
    if (objUnknown.id == "editcell" + ModuleID + Field) {
        var NewValue = objUnknown.value

        if (NewValue == '' || NewValue == '&nbsp;'|| NewValue == '-') {
            Parent.innerHTML = OldValue;
            //warnInvalid(objUnknown, "Please enter a value in the field, no changes were saved");
            return false;    
        }

        if (Field == 'netweight' || Field == 'exportweight' || Field == 'airpackedweight' || Field == 'airpackedcubic' || Field == 'exportcubic') {
            if (!isInteger(NewValue, true)) {
                warnInvalid(objUnknown, "Please enter only whole numbers in the field");
                return false;
            }
        }
        if (Field == 'shortdescription') {
            if (isWhitespace(NewValue)) {
                warnInvalid(objUnknown, "The Short Description is a required field");
                return false;
            }
        }

        if (OldValue == NewValue) { // nothing changed
            if (NewValue == '') {
                Parent.innerHTML = "&nbsp;"
            } else {
                NewValue = NewValue.replace(/</g, "< ");
                NewValue = NewValue.replace(/>/g, " >");
                Parent.innerHTML = NewValue;
            }
        } else { // something changed, save the data
            //var objRS = RSGetASPObject("AMO_RS.asp");

            NewValue = NewValue.replace(/< /g, "<");
            NewValue = NewValue.replace(/ >/g, ">");

            OldValue = "";

            for (var i = 0; i < NewValue.length; i++) {
                if (NewValue.substring(i, i + 1) == '+') {
                    OldValue += "!%%%!"
                } else {
                    OldValue += NewValue.substring(i, i + 1)
                }
            }

            NewValue = OldValue;

            ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=" + Field + "&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
            //var objResult = objRS.setFieldValue(repository, ModuleID, Field, escape(NewValue), fullname);

            NewValue = NewValue.replace(/!%%%!/g, "+");

            $.ajax({
                url: ajaxurl,
                type: "GET",
                async: false,
                success: function (data) {
                    errormsg = data;
                },
                error: function (xhr, status, error) {
                    errormsg = error;
                    erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                }
            })

            if (errormsg == "success") {
                // highlight the field
                Parent.className = "clsAMO_ChangedCell";
                if (NewValue == '') {
                    Parent.innerHTML = "&nbsp;"
                } else {
                    NewValue = NewValue.replace(/</g, "< ");
                    NewValue = NewValue.replace(/>/g, " >");
                    Parent.innerHTML = NewValue;
                }
                // update status to 'In Process'
                //only do this for certain fields, not all of them, Ywang 6/11/2004
                var asObject = document.getElementById("as" + ModuleID);

                if (Field == "infocomment" || Field == "shortdescription" || Field == "bluepartno" || Field == "redpartno" || Field == "replacement" || Field == "manufacturecountry" || Field == "warrantycode") {
                    bchangestatus = true;
                }

                if ((asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled") && bchangestatus) {
                    asObject.innerHTML = "In Process";
                }
            }
        }
    }
}

function calculateCPLBlindDate(RASDate) {
    var somedate = new Date(RASDate);
    var themonth = somedate.getMonth();
    var theday = somedate.getDate();
    var theyear = somedate.getFullYear();
    //1 day prior to GA date : SUG 9763,Vinutha
    somedate = new Date(theyear, themonth, theday - 1);
    return (somedate.getMonth() + 1) + '/' + somedate.getDate() + '/' + somedate.getFullYear();
}


function calculateObsoleteDate(RASDate) {
    var somedate = new Date(RASDate);
    var themonth = somedate.getMonth();
    var theday = somedate.getDate();
    var theyear = somedate.getFullYear();

    // add 3 month to date	
    themonth = themonth + 4;
    var timeA = new Date(theyear, themonth, 1)
    var timeB = new Date(timeA - (60 * 60 * 24 * 1000)); // subtract 1 day
    return (timeB.getMonth() + 1) + '/' + timeB.getDate() + '/' + timeB.getFullYear();
}

/*****************************************************************
//Function:     getDateValue();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************/
function getDateValue(evt, ModuleID, RegionID, Field, OldValue, repository, fullname) {
    if (!evt) evt = window.event;
    var objUnknown = evt.srcElement ? evt.srcElement : evt.target;
    var Parent = objUnknown.parentNode; //parentElement;
    var asObject, rasdiscObject, revaObject, gbsdObject;
    var discontinueDate, mrAvailableDate, globalSeriesDate;
    var ajaxurl = "";
    var errormsg = "";

    rasdiscObject = document.getElementById("rasdisc" + ModuleID + RegionID)
    gbsdObject = document.getElementById("gbsd" + ModuleID + RegionID)
    revaObject = document.getElementById("reva" + ModuleID + RegionID)
    asObject = document.getElementById("as" + ModuleID);

    if (objUnknown.id == "editcell" + ModuleID + Field) {
        if (!checkDate(objUnknown, "", true)){
            return false;
        }
        var NewValue = objUnknown.value

        if (OldValue == NewValue) { // nothing changed
            if (NewValue == '') {
                Parent.innerHTML = "&nbsp;"
            } else {
                Parent.innerHTML = NewValue;
            }
        } else { // something changed, save the data
            if (gbsdObject != null) {
                if ((Field == 'bomrevadate' || Field == 'rasdiscontinuedate') && RegionID == 334 && gbsdObject.innerHTML != '&nbsp;') {

                    if (!confirm("Making a change to RAS Availablity Date, RAS Discontinue Date, or Global Series Config EOL may break a business rule (Global Series Config EOL has to be in the range of RAS Availablity Date and RAS Discontinue Date).\nAre you sure you want proceed ?")) {
                        if (OldValue == ''){
                            OldValue = "&nbsp;"
                        }
                        Parent.innerHTML = OldValue;
                        Parent.style.textDecoration = "underline"; // add underline back
                        return false;
                    }
                }
            }

            if (Field == 'globalseriesdate') {
                if (OldValue == ''){
                    OldValue = "&nbsp;"
                }
                if (revaObject.innerHTML != '&nbsp;'){
                    mrAvailableDate = new Date(revaObject.innerHTML)
                } else {
                    alert("Please enter RAS Availability Date before proceeding");
                    Parent.innerHTML = OldValue;
                    Parent.style.textDecoration = "underline"; // add underline back
                    return false;
                }
                if (rasdiscObject.innerHTML != '&nbsp;'){
                    discontinueDate = new Date(rasdiscObject.innerHTML)
                } else {
                    alert("Please enter RAS Discontinue Date before proceeding");
                    Parent.innerHTML = OldValue;
                    Parent.style.textDecoration = "underline"; // add underline back
                    return false;
                }

                globalSeriesDate = new Date(NewValue)

                if (globalSeriesDate < mrAvailableDate || globalSeriesDate > discontinueDate) {
                    alert("Global Series Config EOL must fall between the RAS Availability Date and the RAS Discontinue Date");
                    Parent.innerHTML = OldValue;
                    Parent.style.textDecoration = "underline"; // add underline back
                    return false;
                }
            }

            //Update Date field 
            ajaxurl = "AMO_SetDateFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&RegionID=" + RegionID + "&Field=" + Field + "&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
            //var objResult = objRS.SaveDateFieldValue(repository, ModuleID, RegionID, Field, escape(NewValue), fullname);

            $.ajax({
                url: ajaxurl,
                type: "GET",
                async: false,
                success: function (data) {
                    errormsg = data;
                },
                error: function (xhr, status, error) {
                    errormsg = error;
                    erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                }
            })

            if (errormsg == "success") {
                // highlight the field
                Parent.className = "clsAMO_ChangedCell";
                if (NewValue == '') {
                    Parent.innerHTML = "&nbsp;"
                } else {
                    Parent.innerHTML = NewValue; //objUnknown.value;																	
                }

                //update Obsolete date to be the last day of the month on the 3rd month

                // update status to 'In Process'
                if ((asObject != null && asObject.innerHTML != "In Process" && asObject.innerHTML != "New" && asObject.innerHTML != "Reject" && asObject.innerHTML != "Re-enabled")
                        && (Field == "bomrevadate" || Field == "rasdiscontinuedate" || Field == "cplblinddate" || Field == "obsoletedate" || Field == "globalseriesdate")) {
                    asObject.innerHTML = "In Process";
                }
            }

            // see if the RAS Availability Date was changed
            if (Field == 'bomrevadate' && NewValue != '') {
                // calculate CPL Blind Date
                var cplObject = document.getElementById("cpl" + ModuleID + RegionID)
                NewValue = calculateCPLBlindDate(NewValue)

                //objRS = RSGetASPObject("AMO_RS.asp");

                ajaxurl = "AMO_SetDateFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&RegionID=" + RegionID + "&Field=cplblinddate&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //objResult = objRS.setDateFieldValue(repository, ModuleID, RegionID, 'cplblinddate', escape(NewValue), fullname);

                $.ajax({
                    url: ajaxurl,
                    type: "GET",
                    async: false,
                    success: function (data) {
                        errormsg = data;
                    },
                    error: function (xhr, status, error) {
                        errormsg = error;
                        erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                    }
                })

                if (errormsg == "success") {
                    // highlight the field
                    if (cplObject != null) {
                        cplObject.className = "clsAMO_ChangedCell";
                        if (NewValue == '') {
                            cplObject.innerHTML = "&nbsp;"
                        } else {
                            cplObject.innerHTML = NewValue;
                        }
                    }
                }
            }
            if (Field == 'rasdiscontinuedate' && NewValue != '') {
                // calculate Obsolete Date
                var obdObject = document.getElementById("obd" + ModuleID + RegionID)
                NewValue = calculateObsoleteDate(NewValue);

                //objRS = RSGetASPObject("AMO_RS.asp");

                ajaxurl = "AMO_SetDateFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&RegionID=" + RegionID + "&Field=obsoletedate&Value=" + NewValue + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
                //objResult = objRS.setDateFieldValue(repository, ModuleID, RegionID, 'obsoletedate', escape(NewValue), fullname);

                $.ajax({
                    url: ajaxurl,
                    type: "GET",
                    async: false,
                    success: function (data) {
                        errormsg = data;
                    },
                    error: function (xhr, status, error) {
                        errormsg = error;
                        erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
                    }
                })

                if (errormsg == "success") {
                    // highlight the field
                    if (obdObject != null) {
                        obdObject.className = "clsAMO_ChangedCell";
                        if (NewValue == '') {
                            obdObject.innerHTML = "&nbsp;"
                        } else {
                            obdObject.innerHTML = NewValue;
                        }
                    }
                }
            }          
            Parent.style.textDecoration = "underline"; // add underline back
        }
    }
}

//4-29-2008 Added by Dien
function cM_ChangeIDP(evtobj, ModuleID, SetStatus, objUnknown, repository, fullname) {
    var menu1 = new Array();
    var wide = 100;
    var myclass = objUnknown.id + ModuleID; // getAttribute("class") for some reason doesn't want to work in IE

    if (SetStatus == 1) {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeIPD(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Ignore SCL Deployment Plan</a>");
    } else {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeIPD(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">SCL Deployment Plan</a>");
    }

    showmenu(evtobj, menu1.join(""), wide + 'px');
}

function cM_ChangeMOLHide(evtobj, ModuleID, SetStatus, objUnknown, repository, fullname) {
    var menu1 = new Array();
    var wide = 100;
    var myclass = objUnknown.id + ModuleID; // getAttribute("class") for some reason doesn't want to work in IE

    if (SetStatus == 1) {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeMOLHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Hide from PRL</a>");
    } else {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeMOLHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Show in PRL</a>");
    }

    showmenu(evtobj, menu1.join(""), wide + 'px');
}


function cM_ChangeSCMHide(evtobj, ModuleID, SetStatus, objUnknown, repository, fullname) {
    var menu1 = new Array()
    var wide = 100;
    var myclass = objUnknown.id + ModuleID; // getAttribute("class") for some reason doesn't want to work in IE

    if (SetStatus == 1) {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeSCMHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Hide from SCM</a>");
    } else {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeSCMHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Show in SCM</a>");
    }

    showmenu(evtobj, menu1.join(""), wide + 'px');
}

function cM_ChangeSCLHide(evtobj, ModuleID, SetStatus, objUnknown, repository, fullname) {
    var menu1 = new Array()
    var wide = 100;
    var myclass = objUnknown.id + ModuleID; // getAttribute("class") for some reason doesn't want to work in IE

    if (SetStatus == 1) {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeSCLHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Hide from SCL</a>");
    } else {
        menu1.push("<a href='../nj.asp' onClick=\"ChangeSCLHideStatus(" + ModuleID + "," + SetStatus + ",'" + myclass + "','" + repository + "','" + fullname + "'); return false;\">Show in SCL</a>");
    }

    showmenu(evtobj, menu1.join(""), wide + 'px');
}

//*****************************************************************
//Function:     ChangeIPD();
//Modified:     4-29-2008 Added by Dien 
//              Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeIPD(ModuleID, SetStatus, myclass, repository, fullname) {
    var objUnknown = getElementsByClassName(document, "td", myclass);
    var strCheckmark
    var ajaxurl = "";
    var errormsg = "";

    //var objRS = RSGetASPObject("AMO_RS.asp");
    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=NoSCLDeploy&Value=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
    //var objResult = objRS.setFieldValue(repository, ModuleID, "NoSCLDeploy", SetStatus, fullname);

    $.ajax({
        url: ajaxurl,
        type: "GET",
        async: false,
        success: function (data) {
            errormsg = data;
        },
        error: function (xhr, status, error) {
            errormsg = error;
            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
        }
    })

    if (errormsg == "success") {
        if (SetStatus == 1) {
            // add the checkmark
            strCheckmark = "<img onclick='javascript:ClickEvent(event);return true;' "
            strCheckmark += "id='idp' title='" + objUnknown[0].title + "' "
            strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>"
            objUnknown[0].innerHTML = strCheckmark
        } else {
            // clear the checkmark
            objUnknown[0].innerHTML = "&nbsp;"
        }
    }
    hidemenu();
}

//*****************************************************************
//Function:     ChangeIPD();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeMOLHideStatus(ModuleID, SetStatus, myclass, repository, fullname) {
    var objUnknown = getElementsByClassName(document, "td", myclass);
    var strCheckmark
    var ajaxurl = "";
    var errormsg = "";

    //var objRS = RSGetASPObject("AMO_RS.asp");
    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=molhide&Value=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
    //var objResult = objRS.setFieldValue(repository, ModuleID, "molhide", SetStatus, fullname);

    $.ajax({
        url: ajaxurl,
        type: "GET",
        async: false,
        success: function (data) {
            errormsg = data;
        },
        error: function (xhr, status, error) {
            errormsg = error;
            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
        }
    })

    if (errormsg == "success") {
        if (SetStatus == 1) {
            // add the checkmark
            strCheckmark = "<img onclick='javascript:ClickEvent(event);return true;' "
            strCheckmark += "id='molhide' title='" + objUnknown[0].title + "' "
            strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>"
            objUnknown[0].innerHTML = strCheckmark
        } else {
            // clear the checkmark
            objUnknown[0].innerHTML = "&nbsp;"
        }
    }
    hidemenu();
}

//*****************************************************************
//Function:     ChangeSCMHideStatus();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeSCMHideStatus(ModuleID, SetStatus, myclass, repository, fullname) {
    var objUnknown = getElementsByClassName(document, "td", myclass);
    var strCheckmark
    var ajaxurl = "";
    var errormsg = "";

    //var objRS = RSGetASPObject("AMO_RS.asp");
    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=scmhide&Value=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
    //var objResult = objRS.setFieldValue(repository, ModuleID, "scmhide", SetStatus, fullname);

    $.ajax({
        url: ajaxurl,
        type: "GET",
        async: false,
        success: function (data) {
            errormsg = data;
        },
        error: function (xhr, status, error) {
            errormsg = error;
            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
        }
    })

    if (errormsg == "success") {
        if (SetStatus == 1) {
            // add the checkmark
            strCheckmark = "<img onclick='javascript:ClickEvent(event);return true;' "
            strCheckmark += "id='scmhide' title='" + objUnknown[0].title + "' "
            strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>"
            objUnknown[0].innerHTML = strCheckmark
        } else {
            // clear the checkmark
            objUnknown[0].innerHTML = "&nbsp;"
        }
    }
    hidemenu();
}

//*****************************************************************
//Function:     ChangeSCLHideStatus();
//Modified:     Harris, Valerie (5/17/2016) - PBI 17492/ Task 20251
//*****************************************************************
function ChangeSCLHideStatus(ModuleID, SetStatus, myclass, repository, fullname) {
    var objUnknown = getElementsByClassName(document, "td", myclass);
    var strCheckmark
    var ajaxurl = "";
    var errormsg = "";

    //var objRS = RSGetASPObject("AMO_RS.asp");
    ajaxurl = "AMO_SetFieldValue.asp?RGS=1&ModuleID=" + ModuleID + "&Field=sclhide&Value=" + SetStatus + "&FullName=" + fullname + "&UserID=" + $("#inpUserID").val();
    //var objResult = objRS.setFieldValue(repository, ModuleID, "sclhide", SetStatus, fullname);

    $.ajax({
        url: ajaxurl,
        type: "GET",
        async: false,
        success: function (data) {
            errormsg = data;
        },
        error: function (xhr, status, error) {
            errormsg = error;
            erroutputArea.innerHTML = "<p><font color=red>" + error + "</font></p>";
        }
    })

    if (errormsg == "success") {
        if (SetStatus == 1) {
            // add the checkmark
            strCheckmark = "<img onclick='javascript:ClickEvent(event);return true;' "
            strCheckmark += "id='sclhide' title='" + objUnknown[0].title + "' "
            strCheckmark += "src=\"../library/Images/gifs/bluechecktrans.gif\" alt=\"\" width=17 height=15 border=0>"
            objUnknown[0].innerHTML = strCheckmark
        } else {
            // clear the checkmark
            objUnknown[0].innerHTML = "&nbsp;"
        }
    }
    hidemenu();
}


//*****************************************************************
//Description:  Open popup window in compatible mode for Infragistics JS control
//Function:     ShowModuleProperties();
//*****************************************************************
function ShowModuleProperties(ModuleID, IRSLink) {
    var iScreenWidth = window.screen.width;
    var iScreenHeight = window.screen.height;
    var w = null;

    //Get Width of Window based on User screen Resolution
    if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
        w = (iScreenWidth - 280);
    } else if (iScreenWidth >= 1600) {
        w = (iScreenWidth - 280);
    } else if (iScreenWidth <= 1024) {
        w = (iScreenWidth - 200);
    }

    //Get Height of Window based on User screen Resolution
    if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
        h = (iScreenHeight - 250);
    } else if (iScreenWidth >= 1600) {
        h = (iScreenHeight - 250);
    } else if (iScreenWidth <= 1024) {
        h = (600);
    }

    var left = (screen.width / 2) - (w / 2);
    var top = (screen.height / 2) - (h / 2);

    //Open popup using MSEdge/Chrome compatible page so that latest version on Infragistics will work
    window.open("../popup/amoproperties.aspx?FeatureID=" + ModuleID + "", "_blank", "resizable=yes,menubar=no,scrollbars=no,toolbar=no,top=" + top + ",left=" + left + ",width=" + w + ",height=" + h + "");
}

function btnDeselectAll_Click() {
    var i;
    var coll = document.getElementsByName("chkBlkStatus");

    if (coll == null) {
        alert("There are no options with RAS Review or RAS Update status available.");
        return false;
    }

    for (i = 0; i < coll.length; i++) {
        coll[i].checked = false
    }
}

function btnSelectAll_Click() {
    var i;
    var coll = document.getElementsByName("chkBlkStatus");

    if (coll == null) {
        alert("There are no options with RAS Review or RAS Update status available.");
        return false;
    }

    for (i = 0; i < coll.length; i++) {
        coll[i].checked = true
    }
}

function btnBulkDateChange_onClick() {
    var strModuleIDs = "";
    var strAllModuleIDs = "";
    var msg = "";
    var arrModId_Status;
    var strModId_Status;
    var bRight;

    var collStatus = document.getElementsByName("chkBlkStatus"); //thisform["chkBlkStatus"];

    if (collStatus == null) {
        alert("There are no options to select.");
        return false;
    }

    for (var i = 0; i < collStatus.length; i++) {
        if (collStatus[i].checked) {
            strModId_Status = collStatus[i].value;
            arrModId_Status = strModId_Status.split("|");
            strAllModuleIDs = strAllModuleIDs + "," + arrModId_Status[0];
            if (arrModId_Status[1] != "RAS Review" && arrModId_Status[1] != "RAS Update" && arrModId_Status[1] != "Disabled" && arrModId_Status[2] != 0)
                strModuleIDs = strModuleIDs + "," + arrModId_Status[0] + "|" + arrModId_Status[3];
        }
    }
    if (strModuleIDs != "")
        strModuleIDs = strModuleIDs.slice(1);

    if (strAllModuleIDs != "" && strModuleIDs == "") {
        alert("Please select different options, you don't have the right to edit the date for selected options.");
        return false;
    } else if (strModuleIDs == "") {
        alert("Please select at least one option for dates change.");
        return false;
    }

    window.open("AMO_AddBulksDate_popup.asp?nModuleIdRegionIds=" + strModuleIDs, "", "width=450,height=450,resizable=no,menubar=no,scrollbars=no,toolbar=no")
}

function btnExpand_onclick(ModuleID) {
    var i;
    var curCtrl = document.getElementsByName("AMO" + ModuleID);

    var IE = document.all ? true : false	// IE uses document.all and Firefox doesn't

    if (curCtrl != null) {
        if (document.getElementById("btnExpand" + ModuleID).value == "+") {
            for (i = 0; i < curCtrl.length; i++) {
                if (IE == true) {
                    curCtrl[i].style.display = "block";
                } else {
                    curCtrl[i].style.display = "table-row";
                }
            }
            document.getElementById("btnExpand" + ModuleID).value = "-";

        } else {
            for (i = 0; i < curCtrl.length; i++) {
                curCtrl[i].style.display = "none";
            }
            document.getElementById("btnExpand" + ModuleID).value = "+";
        }
    }
    return true;
}

//*****************************************************************
//Description:  Load Jquery UI datepicker for AMO date fields
//Function:     load_datePicker();
//*****************************************************************
function load_datePicker() {
    var $browser = get_browser();

    if ($(".filter-dateselection").length > 0) {
        $(".filter-dateselection").datepicker({
            showOn: 'button',
            buttonText: 'From',
            buttonImageOnly: true,
            buttonImage: '../images/calendarButton.gif',
            dateFormat: "mm/dd/yy",
            changeMonth: true,
            changeYear: true,
            firstDay: 7
        });
    }

    if ($(".localization-dateselection").length > 0) {
        $("input").filter('.localization-dateselection').datepicker({
            showOn: 'button',
            buttonText: 'From',
            buttonImageOnly: true,
            buttonImage: '../images/calendarButton.gif',
            dateFormat: "mm/dd/yy",
            firstDay: 7
        });
    }

    $(".ui-datepicker-trigger").css('cursor', 'pointer');
    if ($browser === 'Chrome' || $browser === 'Netscape') {
        $(".ui-datepicker-trigger").css("margin-bottom", "-6px");
    } else {
        $(".ui-datepicker-trigger").css("margin-bottom", "-4px");
    }
    $(".ui-datepicker-trigger").css("margin-left", "3px");

    $(".ui-datepicker").css('margin-left', '-50px', 'margin-top', '-50px');

}


//*****************************************************************
//Description:  Detect Browser Type
//Function:     get_browser();
//*****************************************************************
function get_browser() {
    var N = navigator.appName, ua = navigator.userAgent, tem;
    var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
    if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
    M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
    return M[0];
}

//*****************************************************************
//Description:  Open popup window in compatible mode for Infragistics JS control
//Function:     OpenAMOFeatureCreatePopUp();
//*****************************************************************
function OpenCreateAMOFeatureProperties(url) {
    var iScreenWidth = window.screen.width;
    var iScreenHeight = window.screen.height;
    var w = null;

    //Get Width of Window based on User screen Resolution
    if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
        w = (iScreenWidth - 280);
    } else if (iScreenWidth >= 1600) {
        w = (iScreenWidth - 280);
    } else if (iScreenWidth <= 1024) {
        w = (iScreenWidth - 200);
    }

    //Get Height of Window based on User screen Resolution
    if (iScreenWidth >= 1280 && iScreenWidth < 1600) {
        h = (iScreenHeight - 250);
    } else if (iScreenWidth >= 1600) {
        h = (iScreenHeight - 250);
    } else if (iScreenWidth <= 1024) {
        h = (600);
    }

    var left = (screen.width / 2) - (w / 2);
    var top = (screen.height / 2) - (h / 2);

   //Open popup using MSEdge/Chrome compatible page so that latest version on Infragistics will work
   window.open("../popup/amoproperties.aspx?FeatureID=0", "_blank", "resizable=yes,menubar=no,scrollbars=no,toolbar=no,top=" + top + ",left=" + left + ",width=" + w + ",height=" + h + "");
}

//*****************************************************************
//Description:  Opeb AMO Feature Properties page; called from popup
//Function:     ViewAMOFeature();
//*****************************************************************
function ViewAMOFeature() {
    var iFeatureID = getParameterByName('FeatureID');
    var sAMOPropertiesURL = "";
    var sHTML = "";

    iFeatureID = parseFloat(iFeatureID);

    if (iFeatureID == 0) {
        sAMOPropertiesURL = "../../../../IPulsar/Features/AMOFeatureProperties.aspx?FromASP=1&FeatureID=0";
    } else {
        sAMOPropertiesURL = "../../../../IPulsar/Features/AMOFeatureProperties.aspx?FromModule=1&FromASP=1&FeatureID=" + iFeatureID + "";
    }

    //Change page location to AMO Features Properties: ---
    window.location.href = sAMOPropertiesURL;
}

/*******************************************************/
//Description: get parameter values
/*******************************************************/
function getParameterByName(name) {
    name = name.replace(/[\[]/, "\\[").replace(/[\]]/, "\\]");
    var regex = new RegExp("[\\?&]" + name + "=([^&#]*)"),
        results = regex.exec(location.search);
    return results == null ? "" : decodeURIComponent(results[1].replace(/\+/g, " "));
}

