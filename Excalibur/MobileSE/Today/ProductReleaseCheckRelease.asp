<%@  language="VBScript" %>

<%Option Explicit%>
<!-- #include file="../../includes/DataWrapper.asp" -->

<% 
    dim ReleaseID, ProductVersionID,returnValue
    ReleaseID = request.QueryString("ReleaseID")
    ProductVersionID = request.QueryString("ProductVersionID")
    
    Dim rs, dw, cn, cmd
    Set rs = Server.CreateObject("ADODB.RecordSet")
    Set cn = Server.CreateObject("ADODB.Connection")
    Set cmd = Server.CreateObject("ADODB.Command")
    Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

	Set cmd = dw.CreateCommandSP(cn, "usp_ProductVersion_Release_CheckonAVPRL")
	dw.CreateParameter cmd, "@ReleaseID", adInteger, adParamInput, 8, ReleaseID
    dw.CreateParameter cmd, "@ProductVersionID", adInteger, adParamInput, 8, ProductVersionID
	Set rs = dw.ExecuteCommandReturnRS(cmd)
    returnValue = rs("Message")
    rs.Close 

    response.Write returnValue

%>
