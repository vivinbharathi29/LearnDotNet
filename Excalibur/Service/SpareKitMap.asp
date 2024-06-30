<%@  language="VBScript" %>
<!--#include file="../includes/DataWrapper.asp"-->
<!--#include file="../includes/no-cache.asp"-->
<% 
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True

    regEx.Pattern = "[^0-9]"
    Dim DRID : DRID = regEx.Replace(Request.QueryString("DRID"), "")
    Dim SKID : SKID = regEx.Replace(Request.QueryString("SKID"), "")
    regEx.Pattern = "[^0-9-]"
    Dim SFPN : SFPN = trim(Request.QueryString("SFPN") )
    
	dim ConfigErrors
	
	Set MyBrow=Server.CreateObject("MSWC.BrowserType")
		
	ConfigErrors = ""
	if lcase(MyBrow.browser) <> "ie" or clng(left(MyBrow.version,1)) < 4 then
		ConfigErrors = ConfigErrors & "Internet Explorer 4.0 or greater is required<br>"		
	end if
	
	if not MyBrow.frames then
		ConfigErrors = ConfigErrors & "Frames must be enabled<br>"
	end if
	if not MyBrow.tables then
		ConfigErrors = ConfigErrors & "Table support is required.<br>"
	end if
	if not MyBrow.cookies then
		ConfigErrors = ConfigErrors & "Cookie support is required.<br>"
	end if
	if not MyBrow.javascript then
		ConfigErrors = ConfigErrors & "Javascript support is required.<br>"
	end if

	Dim pageTitle
    Dim rs, dw, cn, cmd
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    If SKID <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "usp_SelectSpareKitDetails")
        dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 0, CLng(SKID)
        dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, ""
        dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, SFPN
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        If Not rs.Eof Then
            pageTitle = rs("Description") & ""
        End If
        rs.close
    End If

    If pageTitle <> "" Then 
        pageTitle = "Spare Kit AV Map - " & pageTitle
    Else
        pageTitle = "Spare Kit AV Map"
    End If
%>
<html>
<head>
    <title><%=pageTitle %></title>
    <!-- #include file="../includes/bundleConfig.inc" -->
</head>
<frameset rows="*" id="TopWindow">
<frame id="UpperWindow" name="UpperWindow" src="SpareKitMapMain2.aspx?<%= Request.QueryString %>">
</frameset>
<input type="hidden" id="txtTitle" name="txtTitle" value="<%=pageTitle%>" />
</html>
<script type="text/javascript">
    $(window).load(function () {
        var sPageTitle = $("#txtTitle").val();
        globalVariable.save(sPageTitle, 'spare_kit_title');
    });
</script>


