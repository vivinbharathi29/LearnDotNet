<%@ Language=VBScript %>


<%
	Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
	dim ProductID, ReleaseID
    if request("FusionRequirements") = 1 then
        dim cn, rs
        set cn = server.CreateObject("ADODB.Connection")
        set rs = server.CreateObject("ADODB.recordset")

	    cn.ConnectionString = Session("PDPIMS_ConnectionString")
	    cn.CommandTimeout=120
	    cn.Open
    
        rs.Open "Select ProductVersionID, ReleaseID from  ProductVersion_Release Where id = " & request("ID"),cn,adOpenStatic
        if not (rs.EOF and rs.bof) then
	        ProductID = rs("ProductVersionID")
            ReleaseID = rs("ReleaseID")
        end if
        rs.Close
        set rs = nothing
	    cn.Close
	    set cn = nothing
    else
        ProductID = request("ID")
        ReleaseID = 0
    end if	  
%>

<!-- #include file = "../../includes/noaccess.inc" -->

<HTML>
<TITLE>Exceptions</TITLE>
<HEAD>

</HEAD>
<FRAMESET ROWS="*,60" ID=TopWindow >
	<FRAME noresize ID="UpperWindow" Name="UpperWindow" SRC="ExceptionsMain.asp?ProductID=<%=ProductID%>&RootID=<%=request("RootID")%>&VersionIDList=<%=request("VersionIDList")%>&ReleaseID=<%=ReleaseID%>">
	<FRAME noresize ID="LowerWindow" Name="LowerWindow" SRC="ExceptionsButtons.asp">
</FRAMESET>

</HTML>