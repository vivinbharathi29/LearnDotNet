<!-- #include file="../includes/DataWrapper.asp" -->

<script runat="server" language="vbscript">


function ClearHistory(ProductBrandID) 
	on error resume next 
    Dim rs, dw, cn, cmd

    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

    Set cmd = dw.CreateCommAndSP(cn, "usp_ClearAvHistory")
    dw.CreateParameter cmd, "@p_ProductBrandID", adInteger, adParamInput, 8, ProductBrandID
    ClearHistory = dw.ExecuteNonQuery(cmd)

    set cmd = nothing
    set cn = nothing
    set dw = nothing
    
end function

dim intProductBrandID

intProductBrandID = request.QueryString("ProductBrandID")
ClearHistory(intProductBrandID)


</script>
