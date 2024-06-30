<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<%

If Request("AVID") <> Request("hidAVID") And Request("hidAVID") <> "" Then
	Response.Redirect "mrDatesButtons.asp?PVID=" & Request("PVID") & "&BID=" & Request("BID") & "&AVID=" & Request("hidAVID")
End If


Dim rs, dw, cn, cmd
Dim iPrevAv, iCurAv, iNextAv

	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set cn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "usp_SelectAvDetail_Pulsar")
	dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Trim(Request("PVID"))
	dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, Trim(Request("BID"))
	dw.CreateParameter cmd, "@p_AvDetailID", adInteger, adParamInput, 8, ""
	dw.CreateParameter cmd, "@p_AvNo", adVarchar, adParamInput, 18, ""
	dw.CreateParameter cmd, "@p_GpgDescription", adVarchar, adParamInput, 50, ""
	dw.CreateParameter cmd, "@p_UPC", adChar, adParamInput, 12, ""
	dw.CreateParameter cmd, "@p_Status", adChar, adParamInput, 1, "A"
	dw.CreateParameter cmd, "@p_KMAT", adChar, adParamInput, 6, ""
	Set rs = dw.ExecuteCommandReturnRS(cmd)

	Do Until rs.EOF Or iCurAv = CLng(Trim(Request("AVID")))
		iPrevAv = iCurAv
		iCurAv = rs("AvDetailID")
		rs.MoveNext
	Loop

	If Not rs.EOF Then
		iNextAv = rs("AvDetailID")
	End If

	rs.Close


%>
<html>
<head>
<script language="JavaScript" src="../includes/client/Common.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() 
{
	window.parent.close();
}

function document_OnLoad()
{
	if (typeof(window.frmButtons.cmdNext) == 'object')
	{
		if (window.frmButtons.hidNext.value == '')
			window.frmButtons.cmdNext.disabled = true;
	}	
	if (typeof(window.frmButtons.cmdPrev) == 'object')
	{
		if (window.frmButtons.hidPrev.value == '')
			window.frmButtons.cmdPrev.disabled = true;
	}
}

function cmdPrev_onclick()
{
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidPrev.value;
	frmButtons.hidAVID.value = frmButtons.hidPrev.value;
	window.parent.frames["UpperWindow"].navigate("mrDatesDetail.asp?AVID=" + frmButtons.hidAVID.value + "&BID=<%=Request("BID")%>&PVID=<%=Request("PVID")%>");
	frmButtons.submit();
}

function cmdNext_onclick()
{
	window.parent.frames["UpperWindow"].frmMain.hidAVID.value = frmButtons.hidNext.value;
	frmButtons.hidAVID.value = frmButtons.hidNext.value;
	window.parent.frames["UpperWindow"].navigate("mrDatesDetail.asp?AVID=" + frmButtons.hidAVID.value + "&BID=<%=Request("BID")%>&PVID=<%=Request("PVID")%>");
	frmButtons.submit();
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<FORM id="frmButtons"  method=post>
<INPUT type="hidden" id=hidPrev name=hidPrev value="<%= iPrevAv%>">
<INPUT type="hidden" id=hidNext name=hidNext value="<%= iNextAv%>">
<INPUT type="hidden" id=hidAVID name=hidAVID value="<%= iCurAv%>">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" WIDTH=100%>
	<tr><TD width=33% align=left>
			&nbsp;</TD>
		<TD width=33% align=center>
			<INPUT type="button" value="Prev" id=cmdPrev name=cmdPrev onclick="return cmdPrev_onclick()">
			<INPUT type="button" value="Next" id=cmdNext name=cmdNext onclick="return cmdNext_onclick()">
		</TD>
		<TD width=33% align=right><INPUT type="button" value="Close" id=cmdCancel name=cmdCancel  onclick="return cmdCancel_onclick()"  ></TD>
	</tr>
</table>

</FORM>
</body>
</html>