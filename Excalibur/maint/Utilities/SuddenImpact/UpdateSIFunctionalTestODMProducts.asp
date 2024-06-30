<%@ Language=VBScript %>
<!-- #include file="../../../includes/emailwrapper.asp" -->

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	
<!-- #include file = "../../../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        if (txtAutoClose.value == "1") {
            window.parent.opener = 'X';
            window.parent.open('', '_parent', '')
            window.parent.close();
        }

    }

//-->
</SCRIPT>
</HEAD>
<STYLE>
TD{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
}
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
}
</STYLE>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%
    if request("UpdateOK") <> "1" then
        response.write "<font color=red><b>Preview Only.  Nothing in Sudden Impact was updated.</b></font><br><br>"
    end if
    response.write "<b>Update ODM Functional Test Components.</b></font><br><br>"
    response.write "<font size=1 face=verdana> - Searches Sudden Impact for Functional Test Deliverables with only one ODM product assigned.</b></font><br>"
    response.write "<font size=1 face=verdana> - Adds the ODM Products to each deliverable found.</b></font><br><br>"

    if request("UpdateOK") <> "1" then
        response.write "<a href=""UpdateSIFunctionalTestODMProducts.asp?UpdateOK=1"">Update Sudden Impact Now</a><br><br>"
    end if

    Server.ScriptTimeout = 5400

	dim cn
	dim cnSI
	dim rs 
	dim strSQL
	dim strUpdate
    dim strDeliverable

	set cn = server.CreateObject("ADODB.Connection")
	set cnSI = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")

	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
    cn.CommandTimeout=500
    cnSI.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=simpact;Server=gvv11651.auth.hpicorp.net,2048;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=PSG_SIMPACT;PASSWORD=@PSG!Pwd2%simpact;" 
	cnSI.IsolationLevel=256
	cnSI.Open

    if trim(request("ProductID")) = "" then
    strSQl = "spListSIDeliverablesMissingODMSFTRecord"
'    strSQl = "Select v.id as VersionID,o.ID as ProductID,o.active,r.OTSFVTOrganizationID, o.ProductVersion, v.deliverablename, v.version, v.revision, v.pass " & _
'             "from deliverableversion v, deliverableroot r, OTSVirtualProducts o " & _
'             "where r.OTSFVTOrganizationID <> 0 " & _
'             "and r.id = v.deliverablerootid " & _
'             "and r.OTSFVTOrganizationID = o.OTSFVTOrganizationID " & _
'             "and v.id in " & _
'				"( " & _
'				"Select Source_Comp_Part_Version_ID  " & _
'				"from OTSVirtualProducts p, HOUSIREPORT01.SIO.dbo.Product_Component pc " & _
'				"where pc.Source_Platform_Version_ID = p.id " & _
'				"and pc.valid_flg=1 " & _
'				"and pc.Platform_Active_Flg=1 " & _
'				"group by Source_Comp_Part_Version_ID " & _
'				"having count(otsfvtorganizationid) = 1 " & _
'				") " & _
 '            "order by r.OTSFVTOrganizationID, v.id, o.ID"
    else
    
    strSQl = "Select v.id as VersionID,o.ID as ProductID,o.active,r.OTSFVTOrganizationID, o.ProductVersion, v.deliverablename, v.version, v.revision, v.pass " & _
             "from deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK), OTSVirtualProducts o with (NOLOCK) " & _
             "where r.OTSFVTOrganizationID =  " & clng(request("ProductID"))  & " " & _
             "and r.id = v.deliverablerootid " & _
             "and (r.OTSFVTOrganizationID = o.OTSFVTOrganizationID or r.OTSFVTOrganizationID = o.id )" & _
             "order by r.OTSFVTOrganizationID, v.id, o.ID"             
'             "and o.id=100003 " & _

    end if
    rs.open strSQL,cn
    if not (rs.eof and rs.bof) then
    	response.write "<table bgcolor=ivory  width=""100%"" border=1 bordercolor=gainsboro cellpadding=2 cellspacing=0>"
	    response.write "<tr bgcolor=beige><td><b>Deliverable</td><td><b>Missing Product</td><td><b>Call to Fix</td></tr>"
    end if
    do while not rs.eof 
        strDeliverable = rs("deliverablename") & " " & rs("Version")
        if trim(rs("Revision") & "") <> "" then
            strDeliverable = strDeliverable & "," & rs("Revision") & ""
        end if
        if trim(rs("Pass") & "") <> "" then
            strDeliverable = strDeliverable & "," & rs("Pass") & ""
        end if
        strUpdate = "UpdatePlatformComponent " & rs("ProductID") & "," & rs("VersionID") & ",0,null,null," & abs(cint(rs("Active"))) 
        response.write "<tr><td>" & strDeliverable & "</td>"
        response.write "<td>" & rs("ProductVersion") & "</td>"
        response.Write "<td>" & strUpdate & "</td></tr>"
	    if request("UpdateOK") = "1" then
            response.write "<BR>" & strUpdate
            cnsi.execute  strUpdate
        end if
        response.Flush
        rs.movenext
    loop
    if not (rs.eof and rs.bof) then
        response.write "</table>"
    end if
    rs.close
    
    set rs = nothing
    cn.close
    cnSI.close
    Set cn = nothing
    set cnSI = nothing

%>
    <input id="txtAutoClose" type="hidden" value="<%=trim(request("autoclose"))%>">
</BODY>
</HTML>
