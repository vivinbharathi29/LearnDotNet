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
Dim Category
Dim RootName

Set rs = Server.CreateObject("ADODB.RecordSet")
Set dw = New DataWrapper
Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

Set cmd = dw.CreateCommandSP(cn, "spGetProductVersionName")
dw.CreateParameter cmd, "@ID", adInteger, adParamInput, 8, Request("PVID")
set rs = dw.ExecuteCommandReturnRS(cmd)

fileName = rs("name")
fileName = Replace(fileName, " ", "_") & "_Commodity_PDD.xls"

Set cmd = dw.CreateCommandSP(cn, "usp_pdd_SelectCommodities")
dw.CreateParameter cmd, "@p_ProductVersionID", adInteger, adParamInput, 8, Request("PVID")
Set rs = dw.ExecuteCommandReturnRS(cmd)	

xlDoc = OpenXlDoc(excelStyle)
xlDoc = xlDoc & OpenXlWorksheet("Commodities")
xlDoc = xlDoc & OpenXlTable()
xlDoc = xlDoc & AddColumn(70)
xlDoc = xlDoc & AddColumn(150)
xlDoc = xlDoc & AddColumn(150)
xlDoc = xlDoc & AddColumn(70)

xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCell("ID", "HdrLeft")
xlDoc = xlDoc & AddXlCell("Supplier", "Hdr")
xlDoc = xlDoc & AddXlCell("Model No.", "Hdr")
xlDoc = xlDoc & AddXlCell("POR/DCR", "HdrRight")
xlDoc = xlDoc & CloseXlRow()

Do Until rs.EOF
	If Category <> rs("Category") Then
		Category = rs("Category")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(Category, "CellCategory", 4)
		xlDoc = xlDoc & CloseXlRow()
	End If
	If RootName <> rs("RootName") Then
		RootName = rs("RootName")
		xlDoc = xlDoc & OpenXlRow("")
		xlDoc = xlDoc & AddXlCellSpan(RootName, "CellRoot", 4)
		xlDoc = xlDoc & CloseXlRow()
	End If
	xlDoc = xlDoc & OpenXlRow("")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("VersionID"), "CellLeft")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("Vendor"), "Cell")
	xlDoc = xlDoc & DrawXlCell(rs.Fields("ModelNumber"), "Cell")
	Select Case rs("DCRID")
		Case 0,1
			xlDoc = xlDoc & AddXlCell("POR", "CellRight")
		Case 2
			xlDoc = xlDoc & AddXlCell("HFCN", "CellRight")
		Case Else
			xlDoc = xlDoc & AddXlCell("DCR: " & rs("DCRID"), "CellRight")
	End Select
	xlDoc = xlDoc & CloseXlRow()
	rs.MoveNext
Loop
	
xlDoc = xlDoc & OpenXlRow("")
xlDoc = xlDoc & AddXlCellSpan("", "Bot", 4)
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
