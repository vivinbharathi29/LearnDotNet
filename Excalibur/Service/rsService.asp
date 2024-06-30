<!-- #include file="../_ScriptLibrary/jsrsServer.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<% jsrsDispatch( "GetBomDetails GetKitDetails GetKitsForProductCategory DeleteSpareKit GetRslPublishDates GetSpbPublishDates SetServiceReportSKU" ) %>
<script language="vbscript" runat="server">
'
' Set Session Variable
'
Function SetServiceReportSKU( strVariableValue )
	
	Session("SKU")=strVariableValue

	SetServiceReportSKU=strVariableValue

End Function


'
' Get Bom Details
'    
Function GetBomDetails( PartNumber )
    on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd
    Dim lastLevel1, lastLevel2, lastLevel3

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_SelectPartBom")
    dw.CreateParameter cmd, "@p_PartNumber", adVarchar, adParamInput, 18, PartNumber

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = ""
    returnValue = "<table class=""BomTable""><tr><th>Kit</th><th>SA</th><th>Component</th><th>&nbsp;</th><th>Description</th></tr>"
    
    If rs.Eof Then
        returnValue = returnValue & "<tr><td colspan=""4"">Bom Not Found</td></tr>"
    End If
    
    lastLevel1 = ""
    lastLevel2 = ""
    lastLevel3 = ""   
    
    Do Until rs.Eof
        If rs("Level1") & "" <> lastLevel1 AND rs("Level1") <> "" Then
            returnValue = returnValue & "<tr><td>" & rs("Level1")&"" & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>" & rs("L1Description")&"" & "</td></tr>"
            lastLevel1 = rs("Level1")
        End If

        If rs("Level2") & "" <> lastLevel2 AND rs("Level2") & "" <> "" Then
            returnValue = returnValue & "<tr><td>&nbsp;</td><td>" & rs("Level2")&"" & "</td><td>&nbsp;</td><td>" & rs("L2SortString")&"" & "</td><td>" & rs("L2Description")&"" & "</td></tr>"
            lastLevel2 = rs("Level2")
        End If

        If rs("Level3") & "" <> lastLevel3 AND rs("Level3") & "" <> "" Then
            returnValue = returnValue & "<tr><td>&nbsp;</td><td>&nbsp;</td><td>" & rs("Level3")&"" & "</td><td>" & rs("L3SortString")&"" & "</td><td>" & rs("L3Description")&"" & "</td></tr>"
            lastLevel3 = rs("Level3")
        End If

        rs.MoveNext
    Loop

    rs.Close
    
    returnValue = returnValue & "</table>"
    GetBomDetails = returnValue
End Function  

Function GetKitDetails( SpareKitNo, SpareKitId, ServiceFamilyPn )  
    on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_SelectSpareKitDetails")
    If SpareKitId = "" Then
        dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 1, ""
        dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, SpareKitNo
    Else
        dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 1, SpareKitId
        dw.CreateParameter cmd, "@p_SpareKitNo", adChar, adParamInput, 10, ""
    End If
    dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, ServiceFamilyPn

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = "No Rows"

    Dim strDescription


    If NOT rs.Eof Then
	
	strDescription=rs("Description")
	If(Len(Trim(strDescription))>40) Then
		strDescription=Mid(Trim(strDescription),1,40)
	End If

        returnValue = rs("ID") & "|" & rs("SpareKitNo") & "|" & strDescription & "|" & rs("SpareCategoryId") & "|" & _
            rs("CsrLevelId") & "|" & rs("Disposition") & "|" & rs("WarrantyTier") & "|" & rs("Notes") & "|" & rs("Comments") & "|" & _
            rs("FirstServiceDt") & "|" & rs("GeoNa") & "|" & rs("GeoLa") & "|" & rs("GeoApj") & "|" & rs("GeoEmea") & "|" & _
            rs("LocalStockAdvice") & "|" & rs("MfgSubAssembly") & "|" & rs("SvcSubAssembly")
    End If
    
    rs.Close
    
    GetKitDetails = returnValue
End Function  


Function GetKitsForProductCategory( CategoryId, ProductVersionId )
    on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd
    Dim lastLevel1, lastLevel2, lastLevel3

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_ListServiceSpareKits")
    dw.CreateParameter cmd, "@p_CategoryId", adInteger, adParamInput, 1, CategoryId
    dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 1, ProductVersionId

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = "<table id=""tblSpareKits"" width=""100%""><tr style=""position: relative; top: expression(document.getElementById('divSpareKits').scrollTop-2)"">" & _
                        "<td class=""header"" onclick=""SortTable( 'tblSpareKits', 0 ,0,1);"">&nbsp;Spare Kit No&nbsp;</td>" & _
                        "<td class=""header"" onclick=""SortTable( 'tblSpareKits', 1 ,0,1);"">&nbsp;Description&nbsp;</td></tr>"

	do while not rs.EOF
	    returnValue = returnValue & "<tr class=""Row"" SKNO=""" & rs("SpareKitNo") & """ id=""" & Trim(rs("Description")&"") & """ ondblclick=""ondblclick_Row();"" onmousedown=""onmousedown_Row();"" onmouseup=""onmouseup_Row();"">"
	    returnValue = returnValue & "<td nowrap>" & rs("SpareKitNo") & "</td>"
	    returnValue = returnValue & "<td nowrap>" & rs("Description") & "</td>"
	    returnValue = returnValue & "</tr>"
		rs.MoveNext
	loop
	rs.Close
        returnValue = returnValue & "</table>"
    GetKitsForProductCategory = returnValue
End Function

Function DeleteSpareKit( ServiceFamilyPn, SpareKitId, UserName )
    on error resume next

    Dim returnValue
    Dim dw, cn, cmd


    Dim strHTTPRef


    strHTTPRef=TRIM(UCASE(Request.ServerVariables("HTTP_REFERER")))

    IF(InStr(strHTTPRef, "/PMVIEW.ASP")>0) THEN

	    Set dw = New DataWrapper
	    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

	    Set cmd = dw.CreateCommAndSP(cn, "usp_DeleteSpareKit_ServiceFamily")
	    dw.CreateParameter cmd, "@p_SpareKitId", adInteger, adParamInput, 1, SpareKitId
	    dw.CreateParameter cmd, "@p_ServiceFamilyPn", adChar, adParamInput, 10, ServiceFamilyPn
	    dw.CreateParameter cmd, "@p_UserName", adVarchar, adParamInput, 50, UserName

	    returnValue = dw.ExecuteNonQuery(cmd)
    
    	If err.number = 0 Then
            DeleteSpareKit = SpareKitId
	Else
            DeleteSpareKit = err.Description
	End If
    ELSE
	DeleteSpareKit = "ERROR - INVALID DELETION ATTEMPT."
    END IF

End Function

Function GetRslPublishDates(ProductVersionId)
    'on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommandSP(cn, "usp_ListRslPublishDates")
    dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 10, ProductVersionId

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = ""

    If Not rs.EOF Then

        returnValue = "<select class=""form"" name=""selCompareDt"" id=""selCompareDt"">"
        returnValue = returnValue & "<option selected value=""" & rs("ChangeDt") & """>" & DayOfWeek(rs("ChangeDt")) & " " & rs("ChangeDt") & "</option>"
        rs.MoveNext

        Do Until rs.EOF
            returnValue = returnValue & "<option value=""" & rs("ChangeDt") & """>" & DayOfWeek(rs("ChangeDt")) & " " & rs("ChangeDt") & "</option>"
            rs.MoveNext
        Loop

        returnValue = returnValue & "</select>"

    End If

    GetRslPublishDates = returnValue

    rs.Close
    set rs = nothing
    cn.Close
    set cn=nothing
End Function

function GetSpbPublishDates(strItem)
    on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_ListSpbPublishDates")
    dw.CreateParameter cmd, "@p_ServiceFamilyPn", adVarchar, adParamInput, 10, strItem

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = ""

    If Not rs.EOF Then

        returnValue = "<select class=""form"" name=""selCompareDt"" id=""selCompareDt"">"
        returnValue = returnValue & "<option selected value=""" & rs("ExportTime") & """>" & DayOfWeek(rs("ExportTime")) & " " & rs("ExportTime") & "</option>"
        rs.MoveNext

        Do Until rs.EOF
            returnValue = returnValue & "<option value=""" & rs("ExportTime") & """>" & DayOfWeek(rs("ExportTime")) & " " & rs("ExportTime") & "</option>"
            rs.MoveNext
        Loop

        returnValue = returnValue & "</select>"

    End If

    GetSpbPublishDates = returnValue

    rs.Close
    set rs = nothing
    cn.Close
    set cn=nothing
end function

Function DayOfWeek(InputDate)

    DIM iDay, strDayName
    iDay = DatePart("w", InputDate)

    SELECT CASE iDay
    Case "1" strDayName = "Sun"
    Case "2" strDayName = "Mon"
    Case "3" strDayName = "Tue"
    Case "4" strDayName = "Wed"
    Case "5" strDayName = "Thu"
    Case "6" strDayName = "Fri"
    Case "7" strDayName = "Sat"
    END SELECT

    DayOfWeek = strDayName
End Function
</script>