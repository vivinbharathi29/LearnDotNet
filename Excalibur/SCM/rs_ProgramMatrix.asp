<!-- #include file="../includes/DataWrapper.asp" -->

<script runat="server" language="vbscript">

function GetPublishDates(strItem)
    on error resume next
    Dim returnValue
    Dim rs, dw, cn, cmd

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_ListProgramMatrixPublishDates")
    dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, strItem

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

    rs.Close
    set rs = nothing
    cn.Close
    set cn=nothing

    GetPublishDates = trim(returnValue)
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

dim intProductBrandID, strReturnValue

intProductBrandID = request.QueryString("ProductBrandID")
strReturnValue = GetPublishDates(intProductBrandID)

response.Write strReturnValue

