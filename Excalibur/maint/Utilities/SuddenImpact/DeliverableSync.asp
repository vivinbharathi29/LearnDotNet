<%@ Language=VBScript %>

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
    response.write "<b>Component Sync.</b></font><br><br>"
    response.write "<font size=1 face=verdana> - Compares deliverables versions in Excalibur to existing versions in Sudden Impact.</b></font><br>"
    response.write "<font size=1 face=verdana> - Looks for developers, devmanagers, and test leads that need to be updated.</b></font><br>"
    response.write "<font size=1 face=verdana> - Compares real deliverables (from the Excalibur deliverableversion table).</b></font><br>"
    response.write "<font size=1 face=verdana> - Looks up a list of products with generic components that need to be updated.</b></font><br><br>"

    if request("UpdateOK") <> "1" then
        response.write "<a href=""DeliverableSync.asp?UpdateOK=1"">Update Sudden Impact Now</a><br><br>"
    end if

	dim cn
	dim rs
    dim strSQL 
	set cn = server.CreateObject("ADODB.Connection")
	set rs = server.CreateObject("ADODB.Recordset")
	
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open 
	response.write "<table bgcolor=ivory  width=""100%"" border=1 bordercolor=gainsboro cellpadding=2 cellspacing=0>"
	response.write "<tr bgcolor=beige><td><b>VersionID</td><td><b>Issue</td><td><b>SI Value</td><td><b>Excalibur Value</td><td><b>Call to Fix</td></tr>"
    
    'Update Real Components
    strSQl = "spListSuddenImpactComponents2Sync"
	rs.open strSQL, cn
    do while not rs.eof
	    'response.Write  "spUpdateSuddenImpactDeliverable " & rs("ID") & ",0" & "<BR>"
        response.write "<tr>"
        response.write "<td>" & rs("ID") & "</td>"
        response.write "<td>" & rs("Issue") & "</td>"
        response.write "<td>" & rs("SI") & "</td>"
        response.write "<td>" & rs("Excalibur") & "</td>"
        response.write "<td>spUpdateSuddenImpactDeliverable " & rs("ID") & ",0</td>"
         
        response.write "</tr>"
        
	    if request("UpdateOK") = "1" then
            cn.execute "spUpdateSuddenImpactDeliverable " & rs("ID") & ",0"
        end if
        rs.movenext
    loop
    rs.close


    'Update Generic Components
    strSQl = "spListSIGenericComponents2Sync"
	rs.open strSQL, cn
    do while not rs.eof
        response.write "<tr>"
        response.write "<td>" & rs("ProductVersionID") & "</td>"
        response.write "<td>Generic Components Updated</td>"
        response.write "<td>&nbsp;</td>"
        response.write "<td>&nbsp;</td>"
        response.write "<td>spUpdateSuddenImpactProduct " & rs("ProductVersionID") & "</td>"
         
        response.write "</tr>"
        
	    if request("UpdateOK") = "1" then
            cn.execute "spUpdateSuddenImpactProduct " & rs("ProductVersionID") 
        end if
        rs.movenext
    loop
    rs.close

	response.write "</table>"
    set rs = nothing
	cn.Close
	set cn=nothing

%>
    <input id="txtAutoClose" type="hidden" value="<%=trim(request("autoclose"))%>">
</BODY>
</HTML>




