<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    var KeyString = "";

    function combo_onkeypress() {
        if (event.keyCode == 13) {
            KeyString = "";
            cmdOK_onclick();
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
        if (parent.window.parent.document.getElementById('modal_dialog')) {
            parent.window.parent.modalDialog.cancel();
        } else {
            window.close();
        }
    }


    function cmdOK_onclick() {
        if (frmMain.cboEmployee.selectedIndex > 0) {
            frmMain.submit();
        }
        else {
            alert("Please select a person from the list.");
        }
    }

    function window_onload() {
        frmMain.cboEmployee.focus();
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
		CurrentUser = lcase(Session("LoggedInUser"))
	
		if instr(currentuser,"\") > 0 then
			CurrentDomain = lcase(left(currentuser, instr(currentuser,"\") - 1))
			Currentuser = lcase(mid(currentuser,instr(currentuser,"\") + 1))
		end if
	
	
%>
<font size=2 face=verdana><b>
<%if request("RoleID") ="1" then
	Response.write "Choose PM"
  else
	Response.write "Choose Developer"
  end if
%>
 for this Component</b></font>
 <form id=frmMain method=post action="ChooseComponentOwnerSave.asp">
<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<TR>
	<TD><b>Name:</b>&nbsp;&nbsp;</TD>
	<TD width=100%><SELECT style="width:100%" id=cboEmployee name=cboEmployee LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION selected value="0"></OPTION>

	<%
	dim strImpersonateID
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")
	
	rs.Open "spGetEmployees",cn,adOpenForwardOnly
	strEmployees = ""
	do while not rs.EOF
		if rs("Active") = 1 then
			if rs("ID") = clng(request("OwnerID")) then
				Response.Write "<OPTION selected value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>" & vbcrlf			
			else
				Response.Write "<OPTION value=""" & rs("ID") & """>" & rs("Name") & "</OPTION>" & vbcrlf
			end if
		end if
		rs.MoveNext
	loop
	
	set rs = nothing	
	cn.Close
	set cn=nothing

%>



		</SELECT>
	</TD>
</TR>
</TABLE>

<TABLE width=100%><TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="    return cmdCancel_onclick()"></td></tr></table>
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtRoleID name=txtRoleID value="<%=request("RoleID")%>">

</form>

</BODY>
</HTML>