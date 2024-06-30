<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
    frmPDMFeedback.submit();
}

function window_onload() {
    frmPDMFeedback.txtPDMFeedback.focus();
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
<BODY bgcolor=Ivory LANGUAGE=javascript>

<%
    dim CurrentUser
	dim AvActionItemID
	CurrentUser=request("CurrentUser")
    AvActionItemID=request("AvActionItemID")
%>
<form ID=frmPDMFeedback action="PDMFeedbackSave2.asp" method=post>
<TABLE width=100%><TR><TD align=center style="font-family:Verdana"><b>Please Enter PDM Feedback Below - <%=AvNo%></b></td></tr></table>
<br />
<input type="text" id=txtPDMFeedback name=txtPDMFeedback maxlength="100" style="position:absolute; width:560px" />
<br />
<br />
<TABLE width=100%><TR><TD align=right><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">&nbsp;<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></td></tr></table>
<INPUT type="hidden" id=txtCurrentUser name=txtCurrentUser value="<%=CurrentUser%>">
<INPUT type="hidden" id=txtAvActionItemID name=txtAvActionItemID value="<%=AvActionItemID%>">
</form>
</BODY>
</HTML>
