<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file = "../includes/DataWrapper.asp" --> 
<!-- #include file = "../includes/Security.asp" --> 
<%
Dim m_ProductID
Dim m_FormSave

m_ProductID = Request("ID")
m_FormSave = True

Call Main()

Sub Main()
	If cbool(Request.Form("hidSave")) then
		Call SaveForm()
	Else
		Call DisplayForm()
	End If
End Sub

Sub SaveForm()

	Dim dw, cn, cmd
	Dim RecordCount
	
	If Len(Trim(Request.Form("cboDrc"))) > 0 Then
		
		Set dw = New DataWrapper
		Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
		Set cmd = dw.CreateCommandSP(cn, "usp_InsertAgencyStatusDCR")
		dw.CreateParameter cmd, "@p_AgencyStatusID", adInteger, adParamInput, 8, Request.Form("hidStatusID")
		dw.CreateParameter cmd, "@p_DcrID", adInteger, adParamInput, 8, Trim(Request.Form("cboDcr"))

		RecordCount = dw.ExecuteNonQuery(cmd)
	
	
		If RecordCount = 0 Then
			Response.Write "<H3>Error Saving Record</H3>"
			Response.end
			Exit Sub
		End If
	End If

	set dw = nothing
	set cn = nothing
	set cmd = nothing

%>
<HTML>
<HEAD>
</HEAD>
<BODY bgcolor="ivory" onload="window_onload();">
</BODY>
</HTML>
<%
End Sub

Sub FillDcrStatus(ProductID)
	Dim dw, cn, cmd, rs
	
	Set dw = New DataWrapper
	Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
	Set cmd = dw.CreateCommandSP(cn, "spListApprovedDCRs")
	dw.CreateParameter cmd, "@ProdID", adInteger, adParamInput, 8, Trim(m_ProductID)
	Set rs = dw.ExecuteCommandReturnRS(cmd)
	
	Do until rs.eof
		Response.Write "<option value=""" & rs("ID") & """>" & rs("ID") & ":" & rs("Summary") & "</option>"					
		rs.movenext
	Loop

	rs.close

	set dw = nothing
	set cn = nothing
	set cmd = nothing
	set rs = nothing
End Sub

Sub DisplayForm()
%>

<HTML>
<HEAD>
</HEAD>
<BODY bgcolor="ivory" >
<FORM id=frmMain method="post" action=ChooseDCRMain.asp>
<label ID="lblTitle"></label>&nbsp;If this program is within 8 weeks of FCS choose the appropriate DCR:</b></font>
<table ID="tabGeneral" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr ID=DcrSelectRow >
		<td valign=top width=140 nowrap><b>DCR:</b>&nbsp;<font color="red" size="1">*</font>&nbsp;</td>
		<td>
			<SELECT id=cboDcr name=cboDcr>
			<option value="">-- Select A DCR --</option>
			<% FillDcrStatus(m_ProductID) %>
			</SELECT>
		</td>
	</tr>
</table>
<input type="hidden" id="hidSave" name="hidSave" value="true">
<input type="hidden" id="hidStatusID" name="hidStatusID" value="<%= Request("StatusID")%>">
</FORM>
</BODY>
</HTML>
<%
End Sub
%>
