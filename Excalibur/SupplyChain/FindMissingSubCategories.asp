<%@  language="VBScript" %>
<% Option Explicit %>
<%
    dim cn
    dim cm
    dim p
    dim rowschanged
    dim rs

    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString")
    cn.Open

    set rs = server.CreateObject("ADODB.Recordset") 
    set cm = server.CreateObject("ADODB.Command")
    Set cm.ActiveConnection = cn
    cm.CommandType = 4
    cm.CommandText = "usp_FindMissingSubCategories"
		
    Set p = cm.CreateParameter("@p_AvParentID", 3, &H0001)
    p.Value = request("AvParentID")
    cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_NewFeatureID", 3, &H0001)
    p.Value = request("NewFeatureID")
    cm.Parameters.Append p

    rs.CursorType = adOpenForwardOnly
    rs.LockType=AdLockReadOnly
    Set rs = cm.Execute 
    Set cm=nothing

    Dim GEO
    GEO = ""
    if not (rs.EOF and rs.BOF) then
        do while not rs.EOF
            if GEO <> "" then
                GEO = GEO + ","
            end if
            GEO = GEO + rs("GEO")
            
		    rs.MoveNext
	    loop
    end if

    Response.Write GEO

    rs.Close

    cn.Close
    set cn = nothing
%>