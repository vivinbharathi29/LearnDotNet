<%@  language="VBScript" %>
<!-- #include file = "../../includes/noaccess.inc" -->
<html>
<head>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<script id="clientEventHandlersJS" language="javascript">
<!--
    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            window.parent.Cancel();
        }
    }

function AddBase(strBase) {
    document.getElementById("txtSub").value = strBase;
}

function AddServiceBase(strBase) {
    document.getElementById("txtServiceSub").value = strBase;
}

function alphanumeric(inputtxt)  
{  
    var letterNumber = /^[0-9a-zA-Z]+$/;  
    if (inputtxt.match(letterNumber))
    {  
        return true;  
    }  
    else  
    {   
       return false;   
    }  
}

function cmdOK_onclick() {    
    if (document.getElementById("txtSub").value.length != 6 && document.getElementById("txtSub").value != "") {
        alert("The Engineering Subassembly Number must be 6 characters long.");
        document.getElementById("txtSub").focus();
        return;
    }
    else if (document.getElementById("txtServiceSub").value.length != 6 && document.getElementById("txtServiceSub").value != "") {
        alert("The Service Subassembly Number must be 6 characters long.")
        document.getElementById("txtServiceSub").focus();
        return;
    }
    else if (document.getElementById("txtDash").value.length != 3 && document.getElementById("txtDash").value != "") {
        alert("The Engineering Dash Number must be 3 characters long.")
        document.getElementById("txtDash").focus();
        return;
    }
    else if (document.getElementById("txtServiceDash").value.length != 3 && document.getElementById("txtServiceDash").value != "") {
        alert("The Service Dash Number must be 3 characters long.")
        document.getElementById("txtServiceDash").focus();
        return;
    }
    else if (!(alphanumeric(document.getElementById("txtSub").value)) && document.getElementById("txtSub").value != "") {
        alert("The Engineering Subassembly Number must be alpha-numeric.")
        document.getElementById("txtSub").focus();
        return;
    }
    else if (!(alphanumeric(document.getElementById("txtServiceSub").value)) && document.getElementById("txtServiceSub").value != "") {
        alert("The Service Subassembly Number must be alpha-numeric.")
        document.getElementById("txtServiceSub").focus();
        return;
    }
    else if (!(alphanumeric(document.getElementById("txtDash").value)) && document.getElementById("txtDash").value != "") {
        alert("The Dash Number must be alpha-numeric.")
        document.getElementById("txtDash").focus();
        return;
    }
    else if (!(alphanumeric(document.getElementById("txtServiceDash").value)) && document.getElementById("txtServiceDash").value != "") {
        alert("The Service Dash Number must be alpha-numeric.")
        document.getElementById("txtDash").focus();
        return;
    }
    if (document.getElementById("txtSelectedProducts").value != "") {
        document.getElementById("txtID").value = document.getElementById("txtID").value + "," + document.getElementById("txtSelectedProducts").value;
        }            
        //alert(frmMain.txtID.value);  
        document.getElementById("frmMain").submit();
}

function cboDeliverable_onchange() {
    RootIDText.innerText = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].value;
    if (document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].Bridged == "1")
        BridgeText.innerHTML = "&nbsp;-&nbsp;<font color=red>Bridged</font>";
    else
        BridgeText.innerHTML = "";

    document.getElementById("txtRootID").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].value;
    document.getElementById("txtID").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].PDRID;
    document.getElementById("txtSub").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].Subassembly;
    document.getElementById("txtDash").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].Spin;
    document.getElementById("txtServiceSub").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].ServiceSubassembly;
    document.getElementById("txtServiceDash").value = document.getElementById("cboDeliverable").options[document.getElementById("cboDeliverable").selectedIndex].ServiceSpin;
}
function window_onload() {
    if (document.getElementById("blnFound").value == "True") {
        LoadExistingDDLValues(document.getElementById("ElementValues").value);
    }
}

function LoadExistingDDLValues(strElementValues) {
    var i;
    var j;
    var TypeID;
    var strNameFormat = "";
    var strNewRow = "";
    var RequiresFormattedName = document.getElementById("RequiresFormattedName").value;

    if (RequiresFormattedName == "False") {
        strNewRow = "<table id=tbName>"
        strNewRow = strNewRow + "<tr><td><label ID=lblFinishedName><font color=black>" + document.getElementById("txtDelName").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "</table>"
        NameRowUpdate.innerHTML = strNewRow;
        NameRowUpdate.style.display = "";
        document.getElementById("txtDelName").style.display = "none";
        return;
    }

    
    for (i = 0; i < document.getElementById("cboAllCategory").length; i++) {
        if (document.getElementById("cboAllCategory").options[i].value == document.getElementById("cboCategory").value) {
            
            if (document.getElementById("cboNameFormat").options[i].text != "") {
                strNameFormat = document.getElementById("cboNameFormat").options[i].text;
            }
            
        }
    }

    if (strNameFormat != "") {
        var NewRows = strNameFormat.split(";");
        var FormatParts;
        var j;

        strElementValues = strElementValues.replace("\'", "");
        strElementValues = strElementValues.replace("'", "");

        var Elements = strElementValues.split(";");
        //alert(Elements);
        var ElementValues;
        var k;
        var cboValues;
        cboValues = "";

        strNewRow = "<table id=tbName oncontextmenu=displayMenu() ondblclick=doSelection() onmousedown=doSelection() onmousemove=doSelection() onmouseup=doSelection()>"
        var ExistingValues = document.getElementById("ExistingNameElements").value;
        //if (ExistingValues != "") {
        ExistingValues = ExistingValues.split("|");
        strNewRow = strNewRow + "<tr><td><b>Engineering:</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName><font color=black>" + document.getElementById("txtDelName").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>GPG (40-char SA):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName2><font color=black>" + document.getElementById("txtDelName2").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>GPG-PhWeb (40-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName3><font color=black>" + document.getElementById("txtDelName3").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>ZSRP (29-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName4><font color=black>" + document.getElementById("txtDelName4").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>GPSy (40-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName5><font color=black>" + document.getElementById("txtDelName5").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>GPSy (200-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName6><font color=black>" + document.getElementById("txtDelName6").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>PMG (100-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName7><font color=black>" + document.getElementById("txtDelName7").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "<tr><td><b>PMG (250-char AV):</b>&nbsp;&nbsp;</td><td><label ID=lblFinishedName8><font color=black>" + document.getElementById("txtDelName8").value + "</font>&nbsp;&nbsp;</label> <a href=# id=linkEdit style=display:none onclick=EditName();>Edit</a></td></tr>"
        strNewRow = strNewRow + "</table><table id=tbEdit>"

        for (i = 0; i < NewRows.length; i++)
            if (NewRows[i] != "") {
            FormatParts = NewRows[i].split("|");
            if (FormatParts.length == 7) {
                if (FormatParts[6] == 0) {
                    //alert(ExistingValues[i]);
                    if (typeof (ExistingValues[i]) != "undefined") {
                        strNewRow = strNewRow + "<tr style=Display:none><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onblur=\"LookupRootNameCount()\" onkeyup=\"return cboElement_onchange()\" value=" + ExistingValues[i] + "><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                    } else {
                    strNewRow = strNewRow + "<tr style=Display:none><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onblur=\"LookupRootNameCount()\" onkeyup=\"return cboElement_onchange()\"><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                    }
                }
                else if (FormatParts[6] == 1) {
                    cboValues = strNewRow + "<tr style=Display:none><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><select id=cboElement name=cboElement class=cbo style=WIDTH:150px onchange=\"return cboElement_onchange();\"><option></option>";
                    for (k = 0; k < Elements.length; k++) {
                        ElementValues = Elements[k].split("|");
                        if (ElementValues[1] == FormatParts[5]) {
                            if (ElementValues[0] == ExistingValues[i]) {
                                cboValues = cboValues + "<Option Selected Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                            } else {
                                cboValues = cboValues + "<Option Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                            }
                        }
                    }
                    if (cboValues != "") {
                        strNewRow = cboValues + "</select><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>";
                        cboValues = "";
                    }
                }
                else if (FormatParts[6] == 2) {
                    //alert(ExistingValues[i]);
                    if (typeof (ExistingValues[i]) != "undefined") {
                        strNewRow = strNewRow + "<tr style=Display:none><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onblur=\"LookupRootNameCount()\" onkeyup=\"return cboElement_onchange()\" value=" + ExistingValues[i] + "><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                    } else {
                    strNewRow = strNewRow + "<tr style=Display:none><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><INPUT type=\"text\" id=cboElement class=" + FormatParts[5] + " name=cboElement style=WIDTH:150px onblur=\"LookupRootNameCount()\" onkeyup=\"return cboElement_onchange()\"><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>"
                    }
                }
                else if (FormatParts[6] == 3) {
                    cboValues = strNewRow + "<tr><td><label ID=lblPreNamePart style=Display:none>" + FormatParts[0] + "</label><b>" + FormatParts[1] + ":</b></td><td><select id=cboElement name=cboElement class=cbo style=WIDTH:150px onchange=\"return cboElement_onchange();\"><option></option>";
                    for (k = 0; k < Elements.length; k++) {
                        ElementValues = Elements[k].split("|");
                        if (ElementValues[1] == FormatParts[5]) {
                            if (ElementValues[0] == ExistingValues[i]) {
                                cboValues = cboValues + "<Option Selected Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                            } else {
                                cboValues = cboValues + "<Option Value=" + ElementValues[0] + ">" + ElementValues[2] + "</OPTION>";
                            }
                        }
                    }
                    if (cboValues != "") {
                        strNewRow = cboValues + "</select><label ID=lblPostNamePart>" + FormatParts[2] + "</label><label ID=lblNameFieldDiv style=\"Display:none\">" + FormatParts[3] + "</label><label><font color=green>&nbsp;&nbsp;" + FormatParts[4] + "</font></label><label ID=lblElementID style=Display:none>" + FormatParts[5] + "</label></td></tr>";
                        cboValues = "";
                    }
                }
            }
        }
        strNewRow = strNewRow + "</table>";
        NameRowUpdate.innerHTML = strNewRow;
        NameRowUpdate.style.display = "";
        document.getElementById("txtDelName").style.display = "none";
    } 
}

function cboElement_onchange() {
    var strBuild = "";
    var strName2 = "";
    var strName3 = "";
    var strName4 = "";
    var strName5 = "";
    var strName6 = "";
    var strName7 = "";
    var strName8 = "";

    var strElementValues = document.getElementById("ElementValues").value;
    var strAvPrefixValues = document.getElementById("AvPrefixValues").value;

    strElementValues = strElementValues.replace("\'", "");
    strElementValues = strElementValues.replace("'", "");

    strAvPrefixValues = strAvPrefixValues.replace("\'", "");
    strAvPrefixValues = strAvPrefixValues.replace("'", "");

    var Elements = strElementValues.split(";");
    var Prefixes = strAvPrefixValues.split(";");

    var ElementValues;
    var PrefixValues;

    var strComments = "";
    var i;
    var Element;

    if (typeof (document.getElementById("cboElement").length) == "undefined") {
        strBuild = strBuild + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName2 = strName2 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName3 = strName3 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName4 = strName4 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName5 = strName5 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName6 = strName6 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName7 = strName7 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
        strName8 = strName8 + document.getElementById("lblPreNamePart").innerText + document.getElementById("cboElement").options(document.getElementById("cboElement").selectedIndex).text + document.getElementById("lblPostNamePart").innerText + document.getElementById("lblNameFieldDiv").innerText;
    }
    else {
        for (i = 0; i < document.getElementById("cboElement").length; i++) {
            //alert(frmMain.cboElement(i).tagName);
            if (document.getElementById("cboElement" + i).tagName == "INPUT") {
                //alert(Elements);
                //alert(frmMain.cboElement(i).className);
                if (document.getElementById("cboElement" + i).value != "" && document.getElementById("cboElement" + i).className == "name=cboElement") {
                    strBuild = strBuild + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName2 = strName2 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName3 = strName3 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName4 = strName4 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName5 = strName5 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName6 = strName6 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName7 = strName7 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                    strName8 = strName8 + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                } else if (document.getElementById("cboElement" + i).value != "") {
                    for (k = 0; k < Elements.length; k++) {
                        ElementValues = Elements[k].split("|");
                        if (ElementValues[1] == document.getElementById("cboElement" + i).className) {
                            //alert(frmMain.cboElement(i).className);
                            if (trim(ElementValues[3]) == "[text]") {
                                //alert(frmMain.cboElement(i).value);
                                strBuild = trim(strBuild) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[4]) == "[text]") {
                                strName2 = trim(strName2) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[5]) == "[text]") {
                                strName3 = trim(strName3) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[6]) == "[text]") {
                                strName4 = trim(strName4) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[7]) == "[text]") {
                                strName5 = trim(strName5) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[8]) == "[text]") {
                                strName6 = trim(strName6) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[9]) == "[text]") {
                                strName7 = trim(strName7) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[10]) == "[text]") {
                                strName8 = trim(strName8) + " " + document.getElementById("lblPreNamePart" + i).innerText + document.getElementById("cboElement" + i).value + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                        }
                    }
                }
            }
            else {

                if (document.getElementById("cboElement" + i).options(document.getElementById("cboElement" + i).selectedIndex).value != "") {
                    for (k = 0; k < Elements.length; k++) {
                        ElementValues = Elements[k].split("|");
                        if (ElementValues[0] == document.getElementById("cboElement" + i).options(document.getElementById("cboElement" + i).selectedIndex).value) {
                            //alert(ElementValues);
                            //alert(ElementValues[10]);
                            if (trim(ElementValues[3]) != "") {
                                strBuild = trim(strBuild) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[3] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[4]) != "") {
                                strName2 = trim(strName2) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[4] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[5]) != "") {
                                strName3 = trim(strName3) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[5] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[6]) != "") {
                                strName4 = trim(strName4) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[6] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[7]) != "") {
                                strName5 = trim(strName5) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[7] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[8]) != "") {
                                strName6 = trim(strName6) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[8] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[9]) != "") {
                                strName7 = trim(strName7) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[9] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                            if (trim(ElementValues[10]) != "") {
                                strName8 = trim(strName8) + " " + document.getElementById("lblPreNamePart" + i).innerText + ElementValues[10] + document.getElementById("lblPostNamePart" + i).innerText + document.getElementById("lblNameFieldDiv" + i).innerText;
                            }
                        }
                    }
                }
            }
        }
    }

    document.getElementById("txtDelName").value = strBuild;
    document.getElementById("lblFinishedName").innerText = strBuild;

    document.getElementById("txtDelName2").value = strName2;
    document.getElementById("lblFinishedName2").innerText = strName2;

    document.getElementById("txtDelName3").value = strName3;
    document.getElementById("lblFinishedName3").innerText = strName3;

    document.getElementById("txtDelName4").value = strName4;
    document.getElementById("lblFinishedName4").innerText = strName4;

    document.getElementById("txtDelName5").value = strName5;
    document.getElementById("lblFinishedName5").innerText = strName5;

    document.getElementById("txtDelName6").value = strName6;
    document.getElementById("lblFinishedName6").innerText = strName6;

    document.getElementById("txtDelName7").value = strName7;
    document.getElementById("lblFinishedName7").innerText = strName7;

    document.getElementById("txtDelName8").value = strName8;
    document.getElementById("lblFinishedName8").innerText = strName8;
    
    //alert(Prefixes.length);
    //alert(Prefixes);
    for (l = 0; l < Prefixes.length; l++) {
        PrefixValues = Prefixes[l].split("|");
        //alert(PrefixValues[0]);
        if (PrefixValues[0] == document.getElementById("tagCategory").value) {
            //alert(PrefixValues[1]);
            if (trim(PrefixValues[2]) != "") {
                document.getElementById("txtDelName2").value = trim(PrefixValues[2]) + " " + strName2;
                document.getElementById("lblFinishedName2").innerText = trim(PrefixValues[2]) + " " + strName2;
            }
            if (trim(PrefixValues[3]) != "") {
                document.getElementById("txtDelName3").value = trim(PrefixValues[3]) + " " + strName3;
                document.getElementById("lblFinishedName3").innerText = trim(PrefixValues[3]) + " " + strName3;
            }
            if (trim(PrefixValues[4]) != "") {
                document.getElementById("txtDelName4").value = trim(PrefixValues[4]) + " " + strName4;
                document.getElementById("lblFinishedName4").innerText = trim(PrefixValues[4]) + " " + strName4;
            }
            if (trim(PrefixValues[5]) != "") {
                document.getElementById("txtDelName5").value = trim(PrefixValues[5]) + " " + strName5;
                document.getElementById("lblFinishedName5").innerText = trim(PrefixValues[5]) + " " + strName5;
            }
            if (trim(PrefixValues[6]) != "") {
                document.getElementById("txtDelName6").value = trim(PrefixValues[6]) + " " + strName6;
                document.getElementById("lblFinishedName6").innerText = trim(PrefixValues[6]) + " " + strName6;
            }
            if (trim(PrefixValues[7]) != "") {
                document.getElementById("txtDelName7").value = trim(PrefixValues[7]) + " " + strName7;
                document.getElementById("lblFinishedName7").innerText = trim(PrefixValues[7]) + " " + strName7;
            }
            if (trim(PrefixValues[8]) != "") {
                document.getElementById("txtDelName8").value = trim(PrefixValues[8]) + " " + strName8;
                document.getElementById("lblFinishedName8").innerText = trim(PrefixValues[8]) + " " + strName8;
            }
        }
    }   
}

function trim(stringToTrim) {
    return stringToTrim.replace(/^\s+|\s+$/g, "");
}

function LookupRootNameCount() {
    document.getElementById("txtDelName").value = document.getElementById("txtDelName").value.replace(String.fromCharCode(8211), "-").replace(/ +/g, " ");

    var strName = document.getElementById("txtDelName").value;
    var strID = document.getElementById("ID").value;
}

function EditName() {
    var table = document.getElementById("tbEdit");
    table.style.display = "";

    var table2 = document.getElementById("linkEdit");
    linkEdit.style.display = "none";
}

function AddSAToProduct(RootID, SAType, PVID, Assign, ReleaseID) {
    var strID
    var sub
    var spin
    if (SAType == 0) {
        sub = document.getElementById("txtSub");
        spin = document.getElementById("txtDash");
    } else {
        sub = document.getElementById("txtServiceSub");
        spin = document.getElementById("txtServiceDash");
    }

    var SelectedProducts = document.getElementById("txtSelectedProducts").value;

    if (sub.value == "" && spin.value == "" && Assign == 1) {
        alert("Please Enter A Subassembly Number.");
        return;
    }
    
    var retValue;
    retValue = window.parent.showModalDialog("<%=AppRoot %>/Excalibur/Deliverable/Commodity/SubAssemblyToMultipleProductsFramePulsar.asp?Mode=add&DRID=" + RootID + "&SA=" + sub.value + "-" + spin.value + "&SAType=" + SAType + "&PVID=" + PVID + "&Selected=" + SelectedProducts + "&Assign=" + Assign + "&ReleaseID=" + ReleaseID, "", "dialogWidth:335px;dialogHeight:516px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");


    if (retValue != undefined && retValue != "-1") {
        
        document.getElementById("txtSelectedProducts").value = retValue;
        var count;

        if (SAType == 0) {
            if (retValue == "0") {
                document.getElementById("lnkAddSAToProduct").innerHTML = "0&nbsp;Additional&nbsp;Product(s)&nbsp;Selected";
                document.getElementById("txtSelectedProducts.value").value = "";
            } else {                
                count = retValue.split(",");
                document.getElementById("lnkAddSAToProduct").innerHTML = count.length + "&nbsp;Additional&nbsp;Product(s)&nbsp;Selected";
                document.getElementById("txtSelectedProducts").value = retValue;
            }
        } else {
            if (retValue == "0") {
                document.getElementById("lnkServiceAddSAToProduct").innerHTML = "0&nbsp;Additional&nbsp;Product(s)&nbsp;Selected";
                document.getElementById("txtSelectedProducts").value = "";
            } else {
                count = retValue.split(",");
                document.getElementById("lnkServiceAddSAToProduct").innerHTML = count.length + "&nbsp;Additional&nbsp;Product(s)&nbsp;Selected";
                document.getElementById("txtSelectedProducts").value = retValue;
            } 
        }
    }
}

function copyText(asHTML) {
    var r3 = document.selection.createRange();
    if (r3.boundingWidth > 0) {
        var str = asHTML ? r3.htmlText : r3.text;
        window.clipboardData.setData('Text', str);
        //document.getElementById('output').innerText = str;
    }
}

var g_rngAnchor = null;
var g_wordSelect = false;
var g_inMenu = false;
function doSelection() {
    try {
        var e = window.event;
        e.cancelBubble = true;
        e.returnValue = false;
        // events.innerHTML = e.type + "-" + e.srcElement.tagName + " (" + e.clientX + ":" + e.clientY + ")<br/>" + events.innerHTML;
        switch (e.type) {
            case 'dblclick':
                // Select a whole word.
                g_rngAnchor = document.selection.createRange();
                if (g_rngAnchor.boundingWidth == 0) {
                    g_rngAnchor.moveToPoint(e.clientX, e.clientY);
                    g_wordSelect = true;
                }
                g_rngAnchor.moveStart("word", -1);
                g_rngAnchor.moveEnd("word");
                g_rngAnchor.select();
                break;

            case 'mousedown':
                // Set a selection point.
                if (e.button == 1/*left*/ && !g_inMenu) {
                    g_wordSelect = false;
                    document.selection.empty();
                    g_rngAnchor = document.selection.createRange();
                    g_rngAnchor.moveToPoint(e.clientX, e.clientY);
                    if (e.shiftKey) {
                        g_rngAnchor.moveStart("word", -1);
                        g_rngAnchor.moveEnd("word");
                        g_wordSelect = true;
                    }
                }
                break;

            case 'mousemove':
                // Drag the selection. (SHIFT+drag to select words.)
                if (e.button == 1 || g_wordSelect) {
                    var r2 = document.selection.createRange();
                    r2.moveToPoint(e.clientX, e.clientY);
                    if (r2.compareEndPoints('StartToStart', g_rngAnchor) == -1) {
                        r2.setEndPoint('StartToEnd', g_rngAnchor);
                        if (e.shiftKey || g_wordSelect) r2.moveStart("word", -1);
                        r2.select();
                    }
                    else {
                        g_rngAnchor.setEndPoint('EndToEnd', r2);
                        if (e.shiftKey || g_wordSelect) g_rngAnchor.moveEnd("word");
                        g_rngAnchor.select();
                    }
                }
                break;

            case 'mouseup':
                g_wordSelect = false;
        }
        // output.innerHTML = g_rngAnchor ? g_rngAnchor.boundingLeft + ":" + g_rngAnchor.boundingWidth : "null";
    }
    catch (e) { }
}

function displayMenu() {
    g_inMenu = true;
    g_wordSelect = false;
    menu1.setCapture(true);
    menu1.style.display = "";
    menu1.style.posLeft = Math.min(event.clientX, document.body.offsetWidth - (menu1.clientWidth + 5));
    menu1.style.posTop = Math.min(event.clientY, document.body.offsetHeight - (menu1.clientHeight + 5));
    event.returnValue = false;
}

function switchMenu() {
    event.cancelBubble = true;
    var el = event.srcElement;
    if (el.className.indexOf("menuItem") == -1) return;
    if (el.className.indexOf("menuItemSelected") != -1) {
        el.className = el.className.replace(/ menuItemSelected/, "");
    }
    else {
        el.className += " menuItemSelected";
    }
}

function clickMenu() {
    var el = event.srcElement;
    switch (el.id) {
        case "mnuCopyAsText":
            copyText(false); break;
        case "mnuCopyAsHTML":
            copyText(true); break;
        default:
            // Not defined
    }
    menu1.style.posLeft = -1000;
    menu1.releaseCapture();
    g_inMenu = false;
}

function SwitchRelease(ProdID, VersionID, RootID, ReleaseID) {
    document.location = "SubAssemblyMainPulsar.asp?ProductID=" + ProdID + "&VersionID=" + VersionID + "&RootID=" + RootID + "&ReleaseID=" + ReleaseID + "&ShowOnlyTargetedRelease=" + document.getElementById("txtShowOnlyTargetedRelease").value + "&RowID=" + document.getElementById("RowID").value;
}

//-->
</script>

<style type="text/css">
#debug {
	background-color: #ddd;
	font-size: x-small;
}
#menu1 {
	position: absolute;
	padding: 3px 5px;
	font-family: "Segoe UI", "MS Sans Serif", sans-serif;
	font-size: 9pt;
	color: MenuText;
	background-color: Menu;
	border: outset 2px ThreeDHighlight;
	border-bottom-color: ThreeDDarkShadow;
	border-right-color: ThreeDLightShadow;
}
.menuItem {
	margin: 2px 0;
	padding: 2px 3px;
}
.menuItemDisabled {
	color: GrayText;
}
.menuItemSelected {
	background-color: Highlight;
}
</style>

</head>
<link rel="stylesheet" type="text/css" href="../../style/programoffice.css">
<body LANGUAGE="javascript" bgcolor="Ivory" onload="return window_onload()">
    <%
	dim cn
	dim rs
	dim blnFound
	dim strID
	dim strProduct
	dim strProductID 
	dim strRootID 
	dim strDeliverable
	dim strSubassembly
	dim strSpin
	dim strType
	dim BaseSubassembliesAvailable
	dim strBaseSubassemblyList
	dim strServiceSubassembly
	dim strServiceSpin
	dim strServiceType
	dim BaseServiceSubassembliesAvailable
	dim strBaseServiceSubassemblyList
	dim strRootOptions
	dim strRootCount
	dim strBridged
	dim strTDCBaseList
	dim strHDCBaseList
	dim strServiceBaseList
	dim strDelName
    dim strDelName2
    dim strDelName3
    dim strDelName4
    dim strDelName5
    dim strDelName6
    dim strDelName7
    dim strDelName8
    dim strElementValues
    dim strExistingNameElements
	dim	strRequiresFormattedName
	dim strTypeID
	dim strCatID
	dim strAvPrefixValues
    dim intReleaseID

    intProductDelRootReleaseID = 0
    strAvPrefixValues = ""
    strDelName = ""
	strDelName2 = ""
	strDelName3 = ""
	strDelName4 = ""
	strDelName5 = ""
	strDelName6 = ""
	strDelName7 = ""
	strDelName8 = ""		
    strElementValues = ""
    strExistingNameElements = ""
	strRequiresFormattedName = ""
	strTypeID = ""
	strCatID = ""
	strTDCBaseList = ""
	strHDCBaseList = ""
	strServiceBaseList = ""
	strRootOptions = ""
	strRootCount = 0
	BaseSubassembliesAvailable = 0
	BaseServiceSubassembliesAvailable = 0
	strBaseSubassemblyList = ""
	strBaseServiceSubassemblyList = ""
	strBridged = ""
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	if request("ProductID") = "" or request("RootID") = "" then
		Response.Write "Not enough information supplied to process your request."
	else
	    'Get User
	    dim CurrentDomain
	    dim CurrentUser
	    dim CurrentUserID
	    dim blnEngCoordinator
	    dim blnSvcCoordinator
	    dim blnOkEnable
	    dim strShowEng
	    dim strShowSvc
	    dim strSQL
	    	    
	    blnEngCoordinator = false
	    blnSvcCoordinator = false
        blnOkEnable=false

	    strShowEng = "none"
	    strShowSvc = "none"

	    CurrentUser = lcase(Session("LoggedInUser"))

	    if instr(currentuser,"\") > 0 then
		    CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		    Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	    end if

	    set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
	    cm.CommandType = 4
	    cm.CommandText = "spGetUserInfo"
	
	    Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	    p.Value = Currentuser
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	    p.Value = CurrentDomain
	    cm.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 
	
	    if not (rs.EOF and rs.BOF) then
    		CurrentUserID = rs("ID") & ""
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    		if rs("EngCoordinator") > 0 or rs("SAAdmin") > 0 then
    		    blnEngCoordinator = true
	            strShowEng = ""
    		end if
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    		if rs("ServiceCoordinator") then
	            blnSvcCoordinator = true
        	    strShowSvc = ""
	        end if
            '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
	        if rs("ServiceCoordinator") or rs("EngCoordinator") > 0 or rs("SAAdmin") > 0 then
	            blnOkEnable = true        	    
	        end if
    	else
	    	CurrentUserID = 0
	    end if
	    rs.Close
	   
        dim strReleaseLink
        dim intDefaultReleaseID
        if Request("ReleaseID") <> "" then
            intDefaultReleaseID = clng(Request("ReleaseID"))	    
        end if
        
        if (intDefaultReleaseID = 0) then
		    strSql = "select top 1 pr.ID, pr.Name from ProductVersion_Release pvr join ProductVersionRelease pr on pr.ID = pvr.ReleaseID where pvr.ProductVersionID= " & clng(Request("ProductID")) & " order by pr.ReleaseYear desc, pr.ReleaseMonth desc;"
		    rs.open strSql,cn
		    if not (rs.EOF and rs.BOF) then
                intDefaultReleaseID = rs("ID")
		    end if	
		    rs.close
	    end if

        if Request("RootID") > 0 then
            strSql = "select Distinct pvr.Name, ReleaseID = pvr.ID, pvr.ReleaseYear, pvr.ReleaseMonth " &_
                     "from Product_Delroot pd with (NOLOCK) " &_
                     "inner join Product_Delroot_Release pdr with (NOLOCK) on pdr.ProductDelRootID = pd.ID " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & request("ProductID") & " and pd.DeliverableRootID= " & Request("RootID") & " order by pvr.Name, pvr.ID, pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
        else 
            strSql = "select pvr.Name, ReleaseID = pvr.ID, pvr.ReleaseYear, pvr.ReleaseMonth " &_
                     "from Product_Deliverable pd " &_
                     "inner join Product_Deliverable_Release pdr on pdr.ProductDeliverableID = pd.ID " &_
                     "inner join ProductVersionRelease pvr on pvr.ID = pdr.ReleaseId " &_
                     "where pd.ProductVersionID= " & request("ProductID") & " and pd.DeliverableVersionID= " & Request("VersionID")
            
            if Request("ShowOnlyTargetedRelease") = 1 then
                strSql = strSql + " and pdr.targeted = 1"
            end if
            
            strSql = strSql + " order by pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
         
        end if

        rs.open strSql, cn
        strReleaseLink = "Switch Releases:&nbsp;"
        
        Do until rs.EOF            
        
            if strReleaseLink <> "Switch Releases:&nbsp;" then
                strReleaseLink = strReleaseLink & " | " 
            end if

            if rs("ReleaseID") = intDefaultReleaseID then
                strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                intReleaseID = rs("ReleaseID")
            else
                strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProductID") & "," & Request("VersionID") & "," & Request("RootID") & "," & rs("ReleaseID") & ");"">" & rs("Name") & "</a>"
            end if
                
            rs.MoveNext
        Loop

        if intReleaseID = 0 then
            strReleaseLink = "Switch Releases:&nbsp;"
            dim count
            count = 0
            rs.MoveFirst
            Do until rs.EOF            
        
                if strReleaseLink <> "Switch Releases:&nbsp;" then
                    strReleaseLink = strReleaseLink & " | " 
                end if

                if  count = 0 then
                    strReleaseLink = strReleaseLink & "<b>" & rs("Name") & "</b>"
                    intReleaseID = rs("ReleaseID")
                else
                    strReleaseLink = strReleaseLink & "<a href=""#"" onclick=""SwitchRelease(" & request("ProductID") & "," & Request("VersionID") & "," & Request("RootID") & "," & rs("ReleaseID") & ");"">" & rs("Name") & "</a>"
                end if
                
                count = count + 1
                rs.MoveNext
            Loop
        end if
        rs.Close

		if request("VersionID") = "" or request("VersionID") = 0 then
			rs.Open "spGetCommoditySubAssemblyNumberRelease " & clng(request("ProductID")) & "," & clng(request("RootID")) & "," & intReleaseID,cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				strSubassembly = ""
				strServiceSubassembly = ""
				blnFound = false
			else
				blnFound = true
				
				strID = rs("ID") & ""
				strProduct = rs("Product") & ""
				strProductID = rs("ProductID") & ""
				strRootID = rs("RootID") & ""
				strDeliverable = rs("Deliverable") & ""
				strSubassembly = trim(rs("SubAssembly") & "")
				strSpin = trim(rs("Spin") & "")
				strServiceSubassembly = trim(rs("ServiceSubAssembly") & "")
				strServiceSpin = trim(rs("ServiceSpin") & "")
				BaseSubassembliesAvailable = 1
                BaseServiceSubassembliesAvailable = 1
				strRootCount = 1
			end if
			rs.Close
		else
			rs.Open "spListSubassembliesBridgedForRootRelease " & clng(request("ProductID")) & "," & clng(request("RootID")) & "," & intReleaseID,cn,adOpenForwardOnly
			if rs.EOF and rs.BOF then
				strSubassembly = ""
				strServiceSubassembly = ""
				blnFound = false
			else
				blnFound = true
				
				strID = rs("ID") & ""
				strProduct = rs("Product") & ""
				strProductID = rs("ProductID") & ""
				strRootID = rs("RootID") & ""
				strDeliverable = rs("Deliverable") & ""
				strSubassembly = trim(rs("SubAssembly") & "")
				strSpin = trim(rs("Spin") & "")
				strServiceSubassembly = trim(rs("ServiceSubAssembly") & "")
				strServiceSpin = trim(rs("ServiceSpin") & "")
				BaseSubassembliesAvailable = 1
                BaseServiceSubassembliesAvailable = 1
				if rs("Bridged") then
					strBridged = "&nbsp;-&nbsp;<font color=red>Bridged</font>"
				else
					strBridged = ""
				end if
			
				do while not rs.EOF
					if strRootCount = 0 then
						strRootOptions = strRootOptions & "<option Bridged=""" & trim(rs("Bridged") & "") & """ Subassembly=""" & trim(rs("Subassembly") & "") & """ ServiceSubassembly=""" & trim(rs("ServiceSubassembly") & "") & """ Spin=""" & trim(rs("Spin") & "") & """ ServiceSpin=""" & trim(rs("ServiceSpin") & "") & """ PDRID=" & rs("ID") & " selected value=""" & trim(rs("RootID")) & """>" & rs("Deliverable") & "</option>"
					else
						strRootOptions = strRootOptions & "<option Bridged=""" & trim(rs("Bridged") & "") & """ Subassembly=""" & trim(rs("Subassembly") & "") & """ ServiceSubassembly=""" & trim(rs("ServiceSubassembly") & "") & """ Spin=""" & trim(rs("Spin") & "") & """ ServiceSpin=""" & trim(rs("ServiceSpin") & "") & """ PDRID=" & rs("ID") & " value=""" & trim(rs("RootID")) & """>" & rs("Deliverable") & "</option>"
					end if
					strRootCount = strRootCount + 1
					rs.MoveNext				
				loop
			end if
			rs.Close		
		end if
		'response.Write(request("VersionID"))
        'response.End()
		'Lookup the base number used on other roots for this dev center
		if trim(strSubassembly) = "" and isnumeric(strProductID) and isnumeric(strRootID) then
			rs.open "spGetSubassembyDefaultBaseRelease " & clng(strProductID) & "," & clng(strRootID),cn,adOpenForwardOnly
			do while not rs.eof
				BaseSubassembliesAvailable = BaseSubassembliesAvailable +1
				strSubassembly = rs("BaseNumber") & ""
				strBaseSubassemblyList = strBaseSubassemblyList & ",<a href=""javascript:  AddBase('" & rs("BaseNumber") & "');"">" & rs("BaseNumber") & "</a>"
				rs.MoveNext
			loop
			if BaseSubassembliesAvailable > 1 then
				strSubassembly = ""
				strBaseSubassemblyList = "Available: " & mid(strBaseSubassemblyList,2)
			else
				strBaseSubassemblyList = ""
			end if
			rs.close
		end if
	
		'Lookup the service base number used on other roots (TDC, HDC, and Service)
		if trim(strServiceSubassembly) = "" and isnumeric(strProductID) and isnumeric(strRootID) then
			rs.open "spGetServiceSubassembyDefaultBaseRelease " & clng(strProductID) & "," & clng(strRootID),cn,adOpenForwardOnly
			do while not rs.eof
                if rs("TypeID") = 2 and instr(strBaseServiceSubassemblyList,rs("BaseNumber")) = 0 then			
				    BaseServiceSubassembliesAvailable = BaseServiceSubassembliesAvailable +1
    				strServiceSubassembly = rs("BaseNumber") & ""
				    strBaseServiceSubassemblyList = strBaseServiceSubassemblyList & ",<a href=""javascript:  AddServiceBase('" & rs("BaseNumber") & "');"">" & rs("BaseNumber") & "</a>"
                end if
				rs.MoveNext
			loop
			if BaseServiceSubassembliesAvailable > 1 then
				strServiceSubassembly = ""
    		    strBaseServiceSubassemblyList = "Available: " & mid(strBaseServiceSubassemblyList,2)
			else
				strBaseServiceSubassemblyList = ""
			end if
			rs.close
		end if			
	
	rs.Open "spGetRootProperties " & clng(request("RootID")),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<INPUT style=""Display:"" type=""text"" id=txtID name=ID value=0>"
	else
		strDisplayedID = rs("ID") & ""
		strDelName = replace(rs("Name") & "","""","&quot;")
		strDelName2 = replace(rs("Name2") & "","""","&quot;")
		strDelName3 = replace(rs("Name3") & "","""","&quot;")
		strDelName4 = replace(rs("Name4") & "","""","&quot;")
		strDelName5 = replace(rs("Name5") & "","""","&quot;")
		strDelName6 = replace(rs("Name6") & "","""","&quot;")
		strDelName7 = replace(rs("Name7") & "","""","&quot;")
		strDelName8 = replace(rs("Name8") & "","""","&quot;")
		strExistingNameElements = rs("NameElements")
		strRequiresFormattedName = rs("RequiresFormattedName")
		strTypeID = rs("TypeID") & ""
		strCatID = rs("CategoryID") & ""
    end if
	rs.close
	end if
	
    strSQL = "usp_SelectDeliverableElementDDLValues"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
	    if strElementValues = "" then
	        strElementValues = "'" & rs("ID") & "|" & rs("ElementID") & "|" & rs("ElementValue") & "|" & rs("Value1") & "|" & rs("Value2") & "|" & rs("Value3") & "|" & rs("Value4") & "|" & rs("Value5") & "|" & rs("Value6") & "|" & rs("Value7") & "|" & rs("Value8")
	    else
	        strElementValues =  strElementValues & ";" & rs("ID") & "|" & rs("ElementID") & "|" & rs("ElementValue") & "|" & rs("Value1") & "|" & rs("Value2") & "|" & rs("Value3") & "|" & rs("Value4") & "|" & rs("Value5") & "|" & rs("Value6") & "|" & rs("Value7") & "|" & rs("Value8")
	    end if
		rs.MoveNext
	loop
	rs.Close
	strElementValues =  strElementValues & "'"
	
	strSQL = "usp_SelectDeliverableCategoryAvPrefixValues"
	rs.Open strSQL,cn,adOpenForwardOnly
	do while not rs.EOF
	    if strAvPrefixValues = "" then
	        strAvPrefixValues = "'" & rs("ID") & "|" & rs("AvPrefix")
	    else
	        strAvPrefixValues =  strAvPrefixValues & ";" & rs("ID") & "|" & rs("AvPrefix")
	    end if
		rs.MoveNext
	loop
	rs.Close
	strAvPrefixValues =  strAvPrefixValues & "'"
    Response.Write "<form ID=frmMain action=""SubassemblySavePulsar.asp"" method=post>"

    %>
	    <input style="display: none" type="text" id="ID" name="ID" value="<%=request("RootID")%>">
        <input id="strNameElements" name="strNameElements" type="hidden">
        <input style="display: none" type="text" id="ExistingNameElements" name="ExistingNameElements" value="<%=strExistingNameElements%>">
        <input style="display: none" type="text" id="RequiresFormattedName" name="RequiresFormattedName" value="<%=strRequiresFormattedName%>">
        <input style="display: none" type="text" id="ElementValues" name="ElementValues" value="<%=strElementValues%>">
        <input style="display: none" type="text" id="txtDelName2" name="txtDelName2" value="<%=strDelName2%>">
        <input style="display: none" type="text" id="txtDelName3" name="txtDelName3" value="<%=strDelName3%>">
        <input style="display: none" type="text" id="txtDelName4" name="txtDelName4" value="<%=strDelName4%>">
        <input style="display: none" type="text" id="txtDelName5" name="txtDelName5" value="<%=strDelName5%>">
        <input style="display: none" type="text" id="txtDelName6" name="txtDelName6" value="<%=strDelName6%>">
        <input style="display: none" type="text" id="txtDelName7" name="txtDelName7" value="<%=strDelName7%>">
        <input style="display: none" type="text" id="txtDelName8" name="txtDelName8" value="<%=strDelName8%>">
        <input style="display: none" type="text" id="AvPrefixValues" name="AvPrefixValues" value="<%=strAvPrefixValues%>">
        <input style="display: none" type="text" id="txtSelectedProducts" name="txtSelectedProducts">  
        <input style="display: none" type="text" id="txtReleaseID" name="txtReleaseID" value="<%=intReleaseID %>"> 
        <input type="hidden" id="txtShowOnlyTargetedRelease" name="txtShowOnlyTargetedRelease" value="<%=Request("ShowOnlyTargetedRelease")%>" />
        
        <input style="display: none" type="text" id="RowID" name="RowID" value="<%=Request("RowID")%>" />
	<%        

	if blnFound then   
	    response.Write "<font size=2 face=verdana><b>Subassembly Numbers</b></font>"
        response.Write "<p>" & strReleaseLink & "</p>"
		
		Response.Write "<table bgcolor=cornsilk border=1 bordercolor=tan cellpadding=2 cellspacing=0 width=""100%"">"
		Response.Write "<TR><TD><b>Product: </b></TD><TD width=""100%"">" & strProduct & "</td></tr>"
		Response.Write "<TR><TD><b>Root ID: </b></TD><TD width=""100%""><span ID=RootIDText>" & request("RootID") & "</span><span ID=BridgeText>" & strBridged & "</span></td></tr>"
		%>
		<tr>
		  <td width="150" nowrap><b>Deliverable&nbsp;Name:</b>&nbsp;&nbsp;</td>
		  <td colspan="10" width="100%">
		    <input oncontextmenu="displayMenu()" ondblclick="doSelection()" onmousedown="doSelection()" onmousemove="doSelection()" onmouseup="doSelection()" LANGUAGE="javascript" onblur="LookupRootNameCount();" id="txtDelName" name="txtDelName" style="WIDTH: 100%; HEIGHT: 22px" size="27" maxlength="120" value="<%=strDelName%>">
		    <div ID=NameRowUpdate style="Display:none"></div>
		  </td>
	    </tr>
	    <tr> 
		  <select style="Display:none" id="cboAllCategory" name="cboAllCategory" style="WIDTH: 250px" >			
		  <%
          dim strSelectedCategories
          dim strNameFormats
          strSelectedCategories = ""
          strNameFormats = ""
    	  strSQL = "spListDeliverableCategoriesByType"
          rs.Open strSQL,cn,adOpenForwardOnly
          do while not rs.EOF
	          Response.Write "<Option Value=" & rs("ID") & ">" & rs("DeliverableTypeID") & ":" & rs("Category") & "</OPTION>"
	          strNameFormats=	strNameFormats & "<option value=""" & rs("ID") &   """>" & rs("NameFormat") & "</option>"
              if trim(strTypeID) = trim(rs("DeliverableTypeID") & "") then
		          if strCatID = rs("ID") & "" then
			          strSelectedCategories = strSelectedCategories & "<Option selected Value=" & rs("ID") & ">" & rs("category") & "</OPTION>"
		          else
			          strSelectedCategories = strSelectedCategories & "<Option Value=" & rs("ID") & ">" & rs("Category") & "</OPTION>"
		          end if
	          end if
	          rs.MoveNext
           loop
           rs.Close
	       %>
		   </select>			
		   <SELECT style="Display:none" id=cboNameFormat name=cboNameFormat><%=strNameFormats%></SELECT>	
           <input id="tagCategory" name="tagCategory" type="hidden" value="<%=trim(strCatID)%>">
           <select id="cboCategory" name="cboCategory" style="WIDTH: 250px;display:none" LANGUAGE="javascript">			
		      <option selected></option>
			  <%=strSelectedCategories%>
		   </select>
        </tr>
	    <%  
		if request("ProductID") = "0" then
			Response.Write "<TR style=""display:" & strShowEng & """><TD><b>Number:&nbsp;</b></TD><TD><INPUT type=""text"" id=txtSub name=txtSub maxlength=6 value=""" & strSubassembly & """>" & "-XXX&nbsp;&nbsp;"
			%>
            <a href="#" style="font-size:x-small" id="lnkAddSAToProduct" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,1,<%=intReleaseID%>);">Assign To Multiple Products</a> | <a href="#" style="font-size:x-small" id="lnkViewProducts" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,0,<%=intReleaseID%>);">View Other Products</a></td></tr>
            <%
			strType=1
		elseif trim(strSubassembly) <> "" then
			Response.Write "<TR style=""display:" & strShowEng & """><TD nowrap valign=top><b>Eng.&nbsp;Subassembly:&nbsp;</b></TD><TD><INPUT type=""text"" style=""width:90"" id=txtSub name=txtSub maxlength=6 value=""" & strSubassembly & """> <b>-</b> <INPUT type=""text"" id=txtDash name=txtDash style=""width:50"" maxlength=3 value=""" & strSpin & """>&nbsp;&nbsp;"
			'Response.Write "<TR><TD><b>Dash:&nbsp;</b></TD><TD><INPUT type=""text"" id=txtDash name=txtDash style=""width:70"" maxlength=3 value=""" & strSpin & """></td></tr>" 
			%>
            <a href="#" style="font-size:x-small" id="lnkAddSAToProduct" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,1,<%=intReleaseID%>);">Assign To Multiple Products</a> | <a href="#" style="font-size:x-small" id="lnkViewProducts" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,0,<%=intReleaseID%>);">View Other Products</a></td></tr>
            <%			
			strType=2
		else
			Response.Write "<TR style=""display:" & strShowEng & """><TD valign=top nowrap><b>Eng.&nbsp;Subassembly:&nbsp;</b></TD><TD><INPUT type=""text"" style=""width:90"" id=txtSub name=txtSub maxlength=6> <b>-</b> <INPUT type=""text"" id=txtDash name=txtDash style=""width:50"" maxlength=3 value=""" & strSpin & """>&nbsp;&nbsp;"
			'Response.Write "<TR><TD><b>Dash:&nbsp;</b></TD><TD><INPUT type=""text"" id=txtDash name=txtDash style=""width:70"" maxlength=3 value=""" & strSpin & """></td></tr>" 
			%>
            <a href="#" style="font-size:x-small" id="lnkAddSAToProduct" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,1,<%=intReleaseID%>);">Assign To Multiple Products</a> | <a href="#" style="font-size:x-small" id="lnkViewProducts" onclick="AddSAToProduct(<%=request("RootID")%>,0,<%=strProductID%>,0,<%=intReleaseID%>);">View Other Products</a><BR><%=strBaseSubassemblyList%></td></tr>
            <%
			strType=3
		end if

		if trim(strServiceSubassembly) <> "" then
			Response.Write "<TR style=""display:" & strShowSvc & """><TD nowrap valign=top><b>Service&nbsp;Subassembly:&nbsp;</b></TD><TD><INPUT type=""text"" style=""width:90"" id=""txtServiceSub"" name=""txtServiceSub"" maxlength=6 value=""" & strServiceSubassembly & """> <b>-</b> <INPUT type=""text"" id=txtServiceDash name=txtServiceDash style=""width:50"" maxlength=3 value=""" & strServiceSpin & """>&nbsp;&nbsp;"
			 %>
             <a href="#" style="font-size:x-small" id="lnkServiceAddSAToProduct" onclick="AddSAToProduct(<%=request("RootID")%>,1,<%=strProductID%>,1,<%=intReleaseID%>);">Assign To Multiple Products</a> | <a href="#" style="font-size:x-small" id="lnkViewProducts" onclick="AddSAToProduct(<%=request("RootID")%>,1,<%=strProductID%>,0,<%=intReleaseID%>);">View Other Products</a></td></tr>
             <%
			strServiceType=2
		else
			Response.Write "<TR style=""display:" & strShowSvc & """><TD valign=top nowrap><b>Service&nbsp;Subassembly:&nbsp;</b></TD><TD><INPUT type=""text"" style=""width:90"" id=""txtServiceSub"" name=""txtServiceSub"" maxlength=6> <b>-</b> <INPUT type=""text"" id=txtServiceDash name=txtServiceDash style=""width:50"" maxlength=3 value=""" & strServiceSpin & """>&nbsp;&nbsp;"
			 %>
             <a href="#" style="font-size:x-small" id="lnkServiceAddSAToProduct" onclick="AddSAToProduct(<%=request("RootID")%>,1,<%=strProductID%>,1,<%=intReleaseID%>);">Assign To Multiple Products</a> | <a href="#" style="font-size:x-small" id="lnkViewProducts" onclick="AddSAToProduct(<%=request("RootID")%>,1,<%=strProductID%>,0,<%=intReleaseID%>);">View Other Products</a><BR><%=strBaseServiceSubassemblyList%></td></tr>
             <%
			strServiceType=3
		end if
		
		if trim(request("IDList") & "") <> "" then
			Response.Write "<TR><TD valign=top nowrap><b>Also&nbsp;Update:&nbsp;</b></TD><TD>"
			response.Write "<div style=""BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll;OVERFLOW-X: scroll; BORDER-LEFT: steelblue 1px solid; BORDER-BOTTOM: steelblue 1px solid; WIDTH: 280px; HEIGHT: 100px; BACKGROUND-COLOR: white"" id=DIV1>"
			response.Write "<TABLE width=100% ID=TableUpdate>"
				response.Write "<THEAD><tr style=""position:relative;top:expression(document.getElementById('DIV1').scrollTop-2);"">"
					response.Write "<TD width=10 style=""BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset""  bgcolor=#c9ddff>&nbsp;ID&nbsp;</TD>"
					response.Write "<TD width=70 style=""BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset""  bgcolor=#c9ddff>&nbsp;Product&nbsp;</TD>"
					response.Write "<TD width=100 style=""BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"" bgcolor=#c9ddff>&nbsp;Deliverable&nbsp;</TD>"
				response.Write "</tr></THEAD>"
			strSQl = "Select r.id as RootID, r.name as Deliverable, p.dotsname as Product " & _
			         "from product_delroot pd, deliverableroot r, productversion p " & _
			         "where pd.productversionid = p.id " & _
			         "and pd.deliverablerootid = r.id " & _
			         "and pd.id in (" & scrubsql(request("IDList")) & ")"
			rs.open strSQL,cn
			do while not rs.eof
			    response.Write "<TR><TD nowrap>&nbsp;" & rs("RootID") & "&nbsp;</td><td nowrap>&nbsp;" & rs("Product") & "&nbsp;</td><td nowrap>&nbsp;" & rs("Deliverable") & "</td></tr>"
                rs.movenext
            loop
            rs.close    
			Response.Write"</Table></div></td></tr>"
		end if
		
		Response.Write "</table>"
'		if BaseSubassembliesAvailable = 0 then
'			Response.Write "<INPUT type=""hidden"" id=txtCopyToAll name=txtCopyToAll value=""1"">"
'		else
'			Response.Write "<INPUT type=""hidden"" id=txtCopyToAll name=txtCopyToAll value=""0"">"
'		end if
        if blnEngCoordinator then
            response.Write "<input style=""display:none"" id=""txtEngCoordinator"" name=""txtEngCoordinator"" type=""text"" value=""1"">"
        end if
        if blnSvcCoordinator then
            response.Write "<input style=""display:none"" id=""txtSvcCoordinator"" name=""txtSvcCoordinator"" type=""text"" value=""1"">"
        end if
        if request("IDList") <> "" then
		    Response.Write "<INPUT type=""hidden"" id=txtID name=txtID value=""" & strID & "_" & intReleaseID & "," & server.HTMLEncode(request("IDList")) & """>"
		else
		    Response.Write "<INPUT type=""hidden"" id=txtID name=txtID value=""" & strID & "_" & intReleaseID & """>"
        end if
		Response.Write "<INPUT type=""hidden"" id=txtType name=txtType value=""" & strType & """>"
		Response.Write "<INPUT type=""hidden"" id=txtRootID name=txtRootID value=""" & request("RootID") & """>"
	
		Response.write "<hr>"
		if blnOkEnable=true then
		    Response.Write "<table border=0 width=""100%""><tr><td style=""display:none"" valign=top><font color=green size=1 family=verdana>Note: The Root Subassembly will be copied to all products from you lab Office using this deliverable.</a></td><td align=right><INPUT type=""Button"" value=""OK"" id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cmdOK_onclick()"">&nbsp;<INPUT type=""Button"" value=""Cancel"" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick=""return cmdCancel_onclick()""></td></tr></table>"
		else
		    Response.Write "<table border=0 width=""100%""><tr><td style=""display:none"" valign=top><font color=green size=1 family=verdana>Note: The Root Subassembly will be copied to all products from you lab Office using this deliverable.</a></td><td align=right><INPUT type=""Button"" value=""OK""  disabled id=cmdOK name=cmdOK LANGUAGE=javascript onclick=""return cmdOK_onclick()"">&nbsp;<INPUT type=""Button"" value=""Cancel"" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick=""return cmdCancel_onclick()""></td></tr></table>"
		end if
		
	else
		Response.Write "<BR>Unable to find some information needed to display this dialog."
	end if

    
    Response.Write "</form>"    
    
	cn.Close
	set rs = nothing
	set cn=nothing

	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
    %>


    <div id="debug"></div>
    <div id="events"></div>

    <!-- Simple context menu -->
    <div id="menu1" onclick="clickMenu()" onmousemove="event.cancelBubble = true" 
	    onmouseout="switchMenu()" onmouseover="switchMenu()" style="display: none">
	    <div id="mnuCopyAsText" class="menuItem">Copy</div>
    </div>

    <input type="hidden" id="blnFound" name="blnFound" value="<%=blnFound%>" />
</body>
</html>