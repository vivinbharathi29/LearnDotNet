<%@ Language=VBScript %>
<!-- #include file = "../includes/DataWrapper.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
TABLE TR TD
{
	vertical-align: text-top;
}
</STYLE>
<LINK rel="stylesheet" type="text/css" href="../style/Excalibur.css">
</HEAD>
<BODY>
<FORM action="rptDiff.asp" method=POST id=form1 name=form1><h2>Localization Comparison Reports</h2>
<P>Select the brands you would like to compare and the report you would like to see.</P>
<P><TABLE BORDER=0>
	<TR>
		<TD ALIGN=top><SELECT size=15 id=cboProdBrand name=cboProdBrand multiple>
<%	set cn = server.CreateObject("ADODB.Connection")
	Set dw = New DataWrapper
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set cm = dw.CreateCommandSP(cn, "usp_ListProductBrands")
	Set rs = dw.ExecuteCommandReturnRS(cm)

	If rs.EOF And rs.BOF Then
		Response.Write "<OPTION>Error Retrieving Data</OPTION>"
		Response.End
	End If
	
	Do Until rs.EOF
		Response.Write "<OPTION Value=" & rs("ProductBrandID") & ">" & rs("BrandFullName") & "</OPTION>"
		rs.MoveNext
	Loop
%>
</SELECT></TD>
		<TD ALIGN=texttop>
<INPUT type="radio" id=ReportNo name=ReportNo value=1 title="Country Delta Report" CHECKED>Country Delta Report<br>
<INPUT type="radio" id=ReportNo name=ReportNo value=2 title="Localization Assignment Delta Report">Localization Assignment Delta Report<br>
<INPUT type="radio" id=ReportNo name=ReportNo value=3 title="Localization Definition Delta Report">Localization Definition Delta Report
<BR><BR><INPUT type="submit" value="Create Report" id=submit1 name=submit1>
</TD>
	</TR>
</TABLE>


</P>
</FORM>
</BODY>
</HTML>
