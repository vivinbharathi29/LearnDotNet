<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	var OutArray = new Array();
	if (typeof(txtSuccess) != "undefined")
		{
		if (txtSuccess.value != "0")
			{
    		OutArray[0] = txtID.value;
    		OutArray[1] = txtOSCode.value;
    		OutArray[2] = txtPartNumber.value;
    		OutArray[3] = txtMultiLanguage.value;
    		OutArray[4] = txtRev.value;
    		OutArray[5] = txtPNRev.value;
    		OutArray[6] = txtStatus.value;
    		OutArray[7] = txtSomeLangPN.value;
    		window.returnValue = OutArray;
			window.parent.close();
			}
        }

}


//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%
    dim LangArray
    dim strLang

    LangArray = split(request("txtLangIDList"),",")
%>

<%
	dim strSuccess
	dim cn
	dim cm
	dim RowsUpdated
    dim FoundSomeByLangPN
	
    strSuccess = "1"
    FoundSomeByLangPN = ""

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open

	cn.BeginTrans
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spUpdateDeliverablePreinstallProperties"
	
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(Request("txtID"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Rev", 3, &H0001)
	if trim(Request("txtRev")) = "" then
		p.Value = null
	else
		p.Value = clng(Request("txtRev"))
	end if
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Team", 3, &H0001)
	if Request("txtTeam") = "" then
		p.Value = 0
	else
		p.Value = clng(Request("txtTeam"))
	end if
	cm.Parameters.Append p
	
	Set p = cm.CreateParameter("@PreinstallPrepStatus", 3, &H0001)
	p.Value = clng(Request("cboPrepStatus"))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@OSCode", 200, &H0001,5)
	p.Value = ucase(left(request("txtOSCode"),5))
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@PartNumber", 200, &H0001,100)
	if trim(request("txtMultiLanguage")) = "0" then
        p.value = ""
    else
        p.Value = ucase(left(request("txtPartNumber"),100))
	end if
    cm.Parameters.Append p

	Set p = cm.CreateParameter("@PartNumberRev", 3, &H0001)
	if Request("txtPNRev") = "" then
		p.Value = null
	else
		p.Value = clng(Request("txtPNRev"))
	end if
	cm.Parameters.Append p

	cm.Execute RowsUpdated
	
	if RowsUpdated <> 1 then
		strSuccess = "0"
	end if
     set cm = nothing

    'Update the language records.
    if strSuccess = "1" then
        'Update each language with a specific part number
    	if trim(request("txtMultiLanguage")) = "0" then
            for each strLang in LangArray
                

        	    set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
	            cm.CommandType = 4
	            cm.CommandText = "spUpdateLanguagePartNumbers"
                response.write 	"<BR>spUpdateLanguagePartNumbers "
	            Set p = cm.CreateParameter("@DeliverableVersionID", 3, &H0001)
	            p.Value = clng(Request("txtID"))
	            cm.Parameters.Append p
                response.write 	clng(Request("txtID")) & ", "

	            Set p = cm.CreateParameter("@PartNumber", 200, &H0001,100)
	            p.Value = ucase(left(trim(request("txtPN" & trim(strLang))),100))
	            cm.Parameters.Append p
                response.write 	ucase(left(trim(request("txtPN" & trim(strLang))),100)) & ", "
                if ucase(left(trim(request("txtPN" & trim(strLang))),100)) <> "" then
                    FoundSomeByLangPN="1"
                end if
	            Set p = cm.CreateParameter("@LangID", 3, &H0001)
	            p.Value = clng(strLang)
	            cm.Parameters.Append p
                response.write 	clng(strLang)

                cm.execute

                if cn.Errors.count > 0 then
                    strSuccess = "0" 
                    exit for
                end if

                set cm = nothing
            next
        else
            'Update all languages with the part number that was entered
        	set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
	        cm.CommandType = 4
	        cm.CommandText = "spUpdateLanguagePartNumbers"
            response.write 	"<BR>spUpdateLanguagePartNumbers "
	
	        Set p = cm.CreateParameter("@DeliverableVersionID", 3, &H0001)
	        p.Value = clng(Request("txtID"))
	        cm.Parameters.Append p
            response.write 	clng(Request("txtID")) & ", "

	        Set p = cm.CreateParameter("@PartNumber", 200, &H0001,100)
	        p.Value = ucase(left(trim(request("txtPartNumber")),100))
	        cm.Parameters.Append p
            response.write 	ucase(left(trim(request("txtPartNumber")),100))

            cm.execute

            if cn.Errors.count > 0 then
                strSuccess = "0" 
            end if

            set cm = nothing

        end if 
    end if






	if strSuccess = "0" then
		cn.RollbackTrans
	else
		cn.CommitTrans
	end if

    'Create part number display for "By Language" deliverables
	if trim(request("txtMultiLanguage")) = "0" then
    	dim strSQL, rs, strPartNumber
        set rs = server.CreateObject("ADODB.recordset")
        strSQL = "spGetSelectedLanguages " & clng(Request("txtID"))
        rs.open strSQL,cn
        strPartNumber = ""
        do while not rs.eof
            if trim(rs("partnumber") & "") = "" then
                strPartNumber = strPartNumber  & rs("CVRCode") & ": MISSING<BR>"
            else
                strPartNumber = strPartNumber  & rs("CVRCode") & ": " & rs("partnumber") & "<BR>"
            end if
            rs.movenext
        loop
        set rs = nothing
    else
        strPartNumber = server.htmlencode(ucase(left(request("txtPartNumber"),100)))
    end if		
	cn.Close
	set cn = nothing
		


%>

<br><br>
<INPUT type="hidden" id=txtID name=txtID value="<%=clng(Request("txtID"))%>">
<INPUT type="hidden" id=txtPartNumber name=txtPartNumber value="<%=strpartNumber%>">
<INPUT type="hidden" id=txtOSCode name=txtOSCode value="<%=server.htmlencode(ucase(left(request("txtOSCode"),5)))%>">
<INPUT type="hidden" id=txtMultiLanguage name=txtMultiLanguage value="<%=server.htmlencode(trim(request("txtMultiLanguage")))%>">
<INPUT type="hidden" id=txtRev name=txtRev value="<%=clng(Request("txtRev"))%>">
<INPUT type="hidden" id=txtPNRev name=txtPNRev value="<%=clng(Request("txtPNRev"))%>">
<INPUT type="hidden" id=txtStatus name=txtStatus value="<%=clng(Request("cboPrepStatus"))%>">
<INPUT type="hidden" id=txtSomeLangPN name=txtSomeLangPN value="<%=FoundSomeByLangPN%>">


<INPUT type="text" id=txtSuccess name=txtSuccess value="<%=strSuccess%>">

</BODY>
</HTML>
