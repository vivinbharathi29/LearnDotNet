<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    var KeyString = "";

    function combo_onkeypress() {
        if (event.keyCode == 13) {
            KeyString = "";
            frmEmployee.submit();
        }
        else {
            KeyString = KeyString + String.fromCharCode(event.keyCode);
            event.keyCode = 0;
            var i;
            var regularexpression;

            for (i = 0; i < event.srcElement.length; i++) {
                regularexpression = new RegExp("^" + KeyString, "i")
                if (regularexpression.exec(event.srcElement.options[i].text) != null) {
                    event.srcElement.selectedIndex = i;
                };

            }
            return false;
        }
    }

    function combo_onfocus() {
        KeyString = "";
    }

    function combo_onclick() {
        KeyString = "";
    }

    function combo_onkeydown() {
        if (event.keyCode == 8) {
            if (String(KeyString).length > 0)
                KeyString = Left(KeyString, String(KeyString).length - 1);
            return false;
        }
    }

    function Left(str, n) {
        if (n <= 0)     // Invalid bound, return blank string
            return "";
        else if (n > String(str).length)   // Invalid bound, return
            return str;                // entire string
        else // Valid bound, return appropriate substring
            return String(str).substring(0, n);
    }


    function cmdCancel_onclick() {
        if (IsFromPulsarPlus()) {
            ClosePulsarPlusPopup();
        }
        else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.close();
            }
        }
    }

    function cmdReset_onclick() {
        frmEmployee.cboEmployee.selectedIndex = 0;
        frmEmployee.submit();
    }

    function cmdOK_onclick() {
        frmEmployee.submit();
    }

    function window_onload() {
        frmEmployee.cboEmployee.focus();
    }

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	FONT-FAMILY=Verdana;
	FONT-SIZE=x-small;
}
</STYLE>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%
		dim CurrentDomain
		dim CurrentUser
		dim strEmployees
		dim CurrentUserID
'		CurrentUser = lcase(Session("LoggedInUser"))
	
'		if instr(currentuser,"\") > 0 then
'			CurrentDomain = lcase(left(currentuser, instr(currentuser,"\") - 1))
'			Currentuser = lcase(mid(currentuser,instr(currentuser,"\") + 1))
'		end if
	
	
%>
<form ID=frmEmployee action="ChangePhWebSave.asp" method=post>
<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<TR>
	<TD><b>Choose&nbsp;Marketing:</b>&nbsp;&nbsp;</TD>
	<TD width=100%><SELECT style="width:100%" id=cboEmployee name=cboEmployee LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
		<OPTION value="0"></OPTION>
		<optgroup label="Commercial">

	<%
	dim strImpersonateID
	dim strEmployeeID
	dim sort
	sort = 1
		
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	
	strEmployeeID=request("EmployeeID")
    strImpersonateID=request("MarketingImpersonateID")

	rs.Open "usp_SelectMarketingList " ,cn,adOpenForwardOnly
	do while not rs.EOF
		if rs("Sort") <> sort then
		    Response.Write "</optgroup><optgroup label=Consumer>"
		    sort = 2
		end if
		if rs("ID") <> clng(strEmployeeID) then
			if trim(rs("ID")) = strImpersonateID then
				Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>" & vbcrlf
			else
				Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>" & vbcrlf
			end if
		end if
		rs.MoveNext
	loop
	rs.Close
	Response.Write "</optgroup>"
	set rs = nothing	
	cn.Close
	set cn=nothing

%>
		</SELECT>
	</TD>
</TR>
</TABLE>

<TABLE width=100%><TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Reset" id=cmdReset name=cmdReset LANGUAGE=javascript onclick="return cmdReset_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr></table>
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=strEmployeeID%>">
</form>
</BODY>
</HTML>
