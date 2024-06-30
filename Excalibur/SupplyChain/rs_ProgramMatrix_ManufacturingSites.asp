<!-- #include file="../includes/DataWrapper.asp" -->

<script runat="server" language="vbscript">    

    dim intProductBrandID
    Dim returnValue
    Dim rs, dw, cn, cmd

    intProductBrandID = request.QueryString("ProductBrandID")

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    Set cmd = dw.CreateCommAndSP(cn, "usp_SCM_GetManufacturingSites")
    dw.CreateParameter cmd, "@p_intProductBrandId", adInteger, adParamInput, 8, intProductBrandID

    Set rs = dw.ExecuteCommAndReturnRS(cmd)

    returnValue = ""
    returnValue = "<select class=""form"" name=""selManufacturingSites"" id=""selManufacturingSites"">"
    returnValue = returnValue & "<option selected value=""-1""></option>"
            
    If Not rs.EOF Then
        returnValue = returnValue & "<option value=""0"">All</option>"
        returnValue = returnValue & "<option value=""" & rs("ManufacturingSiteId") & """>" & rs("Name") & "</option>"
        rs.MoveNext

        Do Until rs.EOF
            returnValue = returnValue & "<option value=""" & rs("ManufacturingSiteId") & """>" & rs("Name") & "</option>"
            rs.MoveNext
        Loop

        returnValue = returnValue & "</select>"

    End If    

    rs.Close
    set rs = nothing
    cn.Close
    set cn=nothing    
    
    response.Write returnValue

</script>

