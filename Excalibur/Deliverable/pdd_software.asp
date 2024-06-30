<%@ Language=VBScript %>

<%
Option Explicit
Response.Buffer = True
%>
<!-- #include file="../includes/ExcelExport.asp" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file="pdd_software_xlstyle.asp" -->
<%
Dim cn
Dim cmd
Dim dw
Dim rs
Dim xlDoc
Dim fileName

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "spGetProductVersionName")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("PVID")
set rs = dw.ExecuteCommandReturnRS(cmd)

fileName = rs("name")
fileName = Replace(fileName, " ", "_") & "_Software_PDD.xls"

Set cmd = dw.CreateCommandSP(cn, "rpt_pdd_Software")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)	
		
xlDoc = OpenXlDoc(excelStyle)
xlDoc = xlDoc & OpenXlWorksheet("Software")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(200)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & DrawXlHeaderRow(rs)

Do Until rs.EOF
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(0), "CellLeft")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(1), "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(2), "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields(3), "CellRight")
	xlDoc = xlDoc & CloseXlRow()
	rs.MoveNext
Loop
	
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & AddXlCell("", "Bot")
xlDoc = xlDoc & CloseXlRow()

xlDoc = xlDoc & CloseXlTable()
xlDoc = xlDoc & CloseXlWorksheet()
xlDoc = xlDoc & CloseXlDoc()

rs.Close
	
Response.Clear
Response.AddHeader "Content-Disposition", "attachment; filename=" & fileName
Response.ContentType = "application/vnd.ms-excel"
Response.Write xlDoc
				
Set rs = nothing
Set cn = nothing
Set cmd = nothing
Set dw = nothing
%>
