<%
Function ExportToExcel(Source, SheetName)
    Dim excelStyle
    Dim excelDoc
	Dim startExcelXML
	Dim endExcelXML
	Dim rowCount
	Dim sheetCount
	Dim field
	Dim x
	Dim rowType
	Dim xmlString
	Dim adBool
	Dim adByte
	
	'Set source = Server.CreateObject("ADODB.RecordSet")
    startExcelXML = "<?xml version=""1.0""?>" & _
		"<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" " & _ 
		"xmlns:o=""urn:schemas-microsoft-com:office:office"" " & _ 
		"xmlns:x=""urn:schemas-microsoft-com:office:excel"" " & _
		"xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" " & _
		"xmlns:html=""http://www.w3.org/TR/REC-html40"" " & _
		"xmlns:u1=""urn:schemas-microsoft-com:office: excel"">"

     endExcelXML = "</Workbook>"

     rowCount = 0
     sheetCount = 1
     
'    <xml version>
'    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
'    xmlns:o="urn:schemas-microsoft-com:office:office"
'    xmlns:x="urn:schemas-microsoft-com:office:excel"
'    xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
'    <Styles>
'    <Style ss:ID="Default" ss:Name="Normal">
'      <Alignment ss:Vertical="Bottom"/>
'      <Borders/>
'      <Font/>
'      <Interior/>
'      <NumberFormat/>
'      <Protection/>
'    </Style>
'    <Style ss:ID="BoldColumn">
'      <Font x:Family="Swiss" ss:Bold="1"/>
'    </Style>
'    <Style ss:ID="StringLiteral">
'      <NumberFormat ss:Format="@"/>
'    </Style>
'    <Style ss:ID="Decimal">
'      <NumberFormat ss:Format="0.0000"/>
'    </Style>
'    <Style ss:ID="Integer">
'      <NumberFormat ss:Format="0"/>
'    </Style>
'    <Style ss:ID="DateLiteral">
'      <NumberFormat ss:Format="mm/dd/yyyy;@"/>
'    </Style>
'    </Styles>
'    <Worksheet ss:Name="Sheet1">
'    </Worksheet>
'    </Workbook>

    excelDoc = startExcelXML
    excelDoc = excelDoc & "<Worksheet ss:Name=""" & SheetName & """>"
    excelDoc = excelDoc & "<Table>"
    excelDoc = excelDoc & "<Column ss:AutoFitWidth=""1""/>"
    excelDoc = excelDoc & "<Column ss:AutoFitWidth=""0"" ss:Width=""66.75"" ss:Span=""2""/>"
    excelDoc = excelDoc & "<Row>"
    for each field in source.Fields
      excelDoc = excelDoc & "<Cell ss:StyleID=""BoldColumn""><Data ss:Type=""String"">"
      excelDoc = excelDoc & field.name
      excelDoc = excelDoc & "</Data></Cell>"
      x = x + 1
	next
	
    excelDoc = excelDoc & "</Row>"

	do until source.EOF
    
      rowCount = rowCount + 1
      
      'if the number of rows is > 64000 create a new page to continue output
      If rowCount=64000 Then
        rowCount = 0
        sheetCount = sheetCount + 1
        excelDoc = excelDoc & "</Table>"
        excelDoc = excelDoc & " </Worksheet>"
        excelDoc = excelDoc & "<Worksheet ss:Name=""Sheet" + sheetCount + """>"
        excelDoc = excelDoc & "<Table>"
	  End If
      
      excelDoc = excelDoc & "<Row>" 'ID=" + rowCount + "
      for each field in source.Fields
      
        'System.Type rowType
        rowType = field.Type
        
        select case rowType
          case adBSTR, adChar, adVarChar, adWChar, _
               adVarWChar, adLongVarChar, adLongVarWChar
             
             xmlString = field.value
             xmlString = Trim(xmlString)
             xmlString = Replace(xmlString,"&","&amp;")
             xmlString = Replace(xmlString,">","&gt;")
             xmlString = Replace(xmlString,"<","&lt;")
             excelDoc = excelDoc & "<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">"
             excelDoc = excelDoc & xmlString
             excelDoc = excelDoc & "</Data></Cell>"

           case adDate
             'Excel has a specific Date Format of YYYY-MM-DD followed by  
             'the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
             'The Following Code puts the date stored in XMLDate 
             'to the format above
             XMLDate = field.value
             XMLDatetoString = "" 'Excel Converted Date
             XMLDatetoString = DatePart("yyyy", XMLDate) & "-"
             If DatePart("m", XMLDate) < 10 Then 
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("m", XMLDate) & "-"
			 Else
				XMLDatetoString = XMLDatetoSTring & DatePart("m", XMLDate) & "-"
			 End If

             If DatePart("d", XMLDate) < 10 Then 
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("d", XMLDate) & "T"
			 Else
				XMLDatetoString = XMLDatetoSTring & DatePart("d", XMLDate) & "T"
			 End If
			 
             If DatePart("h", XMLDate) < 10 Then
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("h", XMLDate) & ":"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("h", XMLDate) & ":"
             End If

             If DatePart("m", XMLDate) < 10 Then
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("m", XMLDate) & ":"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("m", XMLDate) & ":"
             End If

             If DatePart("s", XMLDate) < 10 Then
   				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("s", XMLDate) & ".000"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("s", XMLDate) & ".000"
			 End If
			 
             excelDoc = excelDoc & "<Cell ss:StyleID=""DateLiteral""><Data ss:Type=""DateTime"">"
             excelDoc = excelDoc & XMLDatetoString
             excelDoc = excelDoc & "</Data></Cell>"

			case adBool
                excelDoc = excelDoc & "<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adByte
                excelDoc = excelDoc & "<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adInteger
                excelDoc = excelDoc & "<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adDouble
                excelDoc = excelDoc & "<Cell ss:StyleID=""Decimal""><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case NULL
                excelDoc = excelDoc & "<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String""></Data></Cell>"

            end select
		Next
          
          excelDoc = excelDoc & "</Row>"
          
		  source.MoveNext
        loop

        excelDoc = excelDoc & "</Table>"
        excelDoc = excelDoc & " </Worksheet>"
        excelDoc = excelDoc & endExcelXML
		ExportToExcel =  excelDoc
End Function

Function OpenXlDoc(XlStyle)
	Dim startExcelXML
	
    startExcelXML = "<?xml version=""1.0""?>" & _
		"<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet"" " & _ 
		"xmlns:o=""urn:schemas-microsoft-com:office:office"" " & _ 
		"xmlns:x=""urn:schemas-microsoft-com:office:excel"" " & _
		"xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"" " & _
		"xmlns:html=""http://www.w3.org/TR/REC-html40"" " & _
		"xmlns:u1=""urn:schemas-microsoft-com:office: excel"">"
	
	OpenXlDoc = startExcelXML & XlStyle
End Function

Function CloseXlDoc()
	CloseXlDoc = "</Workbook>"
End Function

Function OpenXlWorksheet(WorksheetName)
	OpenXlWorksheet = "<Worksheet ss:Name=""" & WorksheetName & """>"
End Function

Function CloseXlWorksheet()
	CloseXlWorksheet = "</Worksheet>"
End Function

Function OpenXlTable()
	OpenXlTable = "<Table>"
End Function

Function CloseXlTable()
	CloseXlTable = "</Table>"
End Function

Function AddColumn(Width)
	AddColumn = "<Column ss:Width=""" & Width & """/>"
End Function

Function DrawXlHeaderRow(RecordSet)
	Dim xlDoc
	Dim field
	
    xlDoc = "<Row>"
    for each field in RecordSet.Fields
	  If field.name = RecordSet.Fields(0).name Then
		xlDoc = xlDoc & "<Cell ss:StyleID=""HdrLeft""><Data ss:Type=""String"">"
	  ElseIf field.name = RecordSet.Fields(RecordSet.Fields.Count - 1).name Then
		xlDoc = xlDoc & "<Cell ss:StyleID=""HdrRight""><Data ss:Type=""String"">"
	  Else
		xlDoc = xlDoc & "<Cell ss:StyleID=""Hdr""><Data ss:Type=""String"">"
	  End If
      xlDoc = xlDoc & field.name
      xlDoc = xlDoc & "</Data></Cell>"
	next
	
    DrawXlHeaderRow = xlDoc & "</Row>"

End Function

Function OpenXlRow(Height)
	If Trim(Height) = "0" Then
		OpenXlRow = "<Row ss:AutoFitHeight=""1"" ss:AutoHeight=""1"">"
	ElseIf Len(Trim(Height)) > 0 Then
		OpenXlRow = "<Row ss:AutoFitHeight=""0"" ss:AutoHeight=""0"" ss:Height=""" & Height & """>"
	Else
		OpenXlRow = "<Row>"
	End If
End Function

Function CloseXlRow()
	CloseXlRow = "</Row>"
End Function

Function DrawXlCell(field, style)
	Dim rowType
	Dim xmlString
	Dim excelDoc
	Dim XMLDate
	Dim XMLDatetoString
	
    rowType = field.Type
    
    'Response.Write "RowType:" & rowType
    'Response.End
    
    select case rowType
         case adBSTR, adChar, adVarChar, adWChar, _
               adVarWChar, adLongVarChar, adLongVarWChar
             
             If IsNull(field.value) Then
				xmlString = ""
			 Else
				xmlString = field.value
				xmlString = Trim(xmlString)
				xmlString = Replace(xmlString,"&","&amp;")
				xmlString = Replace(xmlString,">","&gt;")
				xmlString = Replace(xmlString,"<","&lt;")
			 End If
             excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""String"">"
             excelDoc = excelDoc & xmlString
             excelDoc = excelDoc & "</Data></Cell>"

           case adDate, adDBDate, adDBTime, adDBTimeStamp
             'Excel has a specific Date Format of YYYY-MM-DD followed by  
             'the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
             'The Following Code puts the date stored in XMLDate 
             'to the format above
             XMLDate = field.value
             XMLDatetoString = "" 'Excel Converted Date
             XMLDatetoString = DatePart("yyyy", XMLDate) & "-"
             If DatePart("m", XMLDate) < 10 Then 
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("m", XMLDate) & "-"
			 Else
				XMLDatetoString = XMLDatetoSTring & DatePart("m", XMLDate) & "-"
			 End If

             If DatePart("d", XMLDate) < 10 Then 
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("d", XMLDate) & "T"
			 Else
				XMLDatetoString = XMLDatetoSTring & DatePart("d", XMLDate) & "T"
			 End If
			 
             If DatePart("h", XMLDate) < 10 Then
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("h", XMLDate) & ":"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("h", XMLDate) & ":"
             End If

             If DatePart("m", XMLDate) < 10 Then
				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("m", XMLDate) & ":"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("m", XMLDate) & ":"
             End If

             If DatePart("s", XMLDate) < 10 Then
   				XMLDatetoString = XMLDatetoSTring & "0" & DatePart("s", XMLDate) & ".000"
             Else
				XMLDatetoString = XMLDatetoSTring & DatePart("s", XMLDate) & ".000"
			 End If
			 
             excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & "_Date""><Data ss:Type=""DateTime"">"
             excelDoc = excelDoc & XMLDatetoString
             excelDoc = excelDoc & "</Data></Cell>"

			case adBoolean
                excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""String"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adSmallInt
                excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adInteger
                excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case adDouble
                excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""Number"">"
                excelDoc = excelDoc & field.value
                excelDoc = excelDoc & "</Data></Cell>"

            case NULL
                excelDoc = excelDoc & "<Cell ss:StyleID=""" & style & """><Data ss:Type=""String""></Data></Cell>"

            end select
            
	DrawXlCell = excelDoc

End Function

Function AddXlCell(Value, Style)
	Dim xmlString
	
	xmlString = Trim(Value&"")
	xmlString = Replace(xmlString,"&","&amp;")
	xmlString = Replace(xmlString,">","&gt;")
	xmlString = Replace(xmlString,"<","&lt;")
	If isNumeric(xmlString) Then
		AddXlCell = "<Cell ss:StyleID=""" & style & """><Data ss:Type=""Number"">" & xmlString & "</Data></Cell>"
	Else
		AddXlCell = "<Cell ss:StyleID=""" & style & """><Data ss:Type=""String"">" & xmlString & "</Data></Cell>"
	End If
End Function

Function AddXlCellSpan(Value, Style, CellSpan)
	Dim xmlString
	
	xmlString = Trim(Value&"")
	xmlString = Replace(xmlString,"&","&amp;")
	xmlString = Replace(xmlString,">","&gt;")
	xmlString = Replace(xmlString,"<","&lt;")
	AddXlCellSpan = "<Cell ss:MergeAcross=""" & CellSpan - 1 & """ ss:StyleID=""" & style & """><Data ss:Type=""String"">" & xmlString & "</Data></Cell>"
End Function

Function AddXlCellSpanBoth(Value, Style, CellSpan, ColSpan)
	Dim xmlString
	
	xmlString = Trim(Value&"")
	xmlString = Replace(xmlString,"&","&amp;")
	xmlString = Replace(xmlString,">","&gt;")
	xmlString = Replace(xmlString,"<","&lt;")
	AddXlCellSpanBoth = "<Cell ss:MergeAcross=""" & CellSpan - 1 & """ ss:MergeDown=""" & ColSpan - 1 & """ ss:StyleID=""" & style & """><Data ss:Type=""String"">" & xmlString & "</Data></Cell>"
End Function

Function AddXlHtmlCell(Value, Style, CellSpan)
	AddXlHtmlCell = "<Cell ss:MergeAcross=""" & CellSpan - 1 & """ ss:StyleID=""" & style & """><ss:Data ss:Type=""String"" xmlns=""http://www.w3.org/TR/REC-html40"">" & Value&"" & "</ss:Data></Cell>"
End Function
%>