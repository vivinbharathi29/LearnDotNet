<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<TITLE>Reject Deliverable Request</TITLE>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function window_onload() {
    if (typeof(frmMain) != "undefined")
        frmMain.txtComments.focus();
}

function CheckTextSize(field, maxLength) {
	if (field.value.length > maxLength + 1)
		{
		field.value = field.value.substring(0, maxLength);
		alert("The maximum size of this field in 200 characters. You input has been truncated.");
		}
	else if (field.value.length >= maxLength)
		{
		window.event.keyCode=0;
		field.value = field.value.substring(0, maxLength);
		}
} 

//-->
</SCRIPT>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<%
    
	if trim(request("ID")) = "" or trim(request("NewValue")) = "" then
		Response.Write "Unable to process this request."
	else

        dim ItemArray
        dim strRoots
        dim strVersions
        dim strItem
        dim ValuePair

        ItemArray = split(request("ID"),",")

        strRoots = ""
        strVersions = ""
        for each strItem in ItemArray
            ValuePair = split(strItem,"_")
            if trim(Valuepair(0)) = "1" then
                strRoots = strRoots & "," & trim(Valuepair(1))
            else
                strVersions = strVersions & "," & trim(Valuepair(1))
            end if
        next
        if strRoots <> "" then
            strRoots = mid(strRoots,2)
        end if
        if strVersions <> "" then
            strVersions = mid(strVersions,2)
        end if


%>
<font size=3 face=verdana><b>Reject Deliverable Request</b></font><BR><BR>
<%if lbound(ItemArray)= ubound(ItemArray) then %>
    <font size=1 face=verdana>Enter comments explaining why this deliverable request is being rejected.</font><BR>
<%else%>
    <font size=1 face=verdana>Enter comments explaining why these deliverable requests are being rejected.</font><BR>
<%end if%>
<form id="frmMain" method="post" action="SaveDevNotificationStatus.asp">

<table border=1 width="100%" bordercolor=tan cellspacing=0 cellpadding=2>
    <TR bgcolor=cornsilk>
	<TD nowrap valign=top><b><font size=1 face=verdana>Comments:</font></b>&nbsp;<font color="#ff0000" size="1">*</font>
	<br>
	<font size=1 color=green>Max Length: 200&nbsp;&nbsp;</font>
	</td>
	<TD width="100%" ><TEXTAREA style="WIDTH:100%; HEIGHT:80px" id=txtComments name=txtComments onkeypress="CheckTextSize(this, 200)"  onchange="CheckTextSize(this, 200)"><%=strTestNotes%></TEXTAREA></td>
    </TR>
    <tr bgcolor=cornsilk><td valign=top><b><font size=1 face=verdana>Reject&nbsp;Request:</font></b></td><td>
        
        <table border=1 bgcolor=white bordercolor=gainsboro cellpadding=2 cellspacing=0 width="100%">
        <tr bgcolor=gainsboro>
            <td><font size=1 face=verdana><b>ID</b></font></td>
            <td><font size=1 face=verdana><b>Product</b></font></td>
            <td><font size=1 face=verdana><b>Vendor</b></font></td>
            <td><font size=1 face=verdana><b>Deliverable</b></font></td>
            <td><font size=1 face=verdana><b>Version</b></font></td>
            <td><font size=1 face=verdana><b>Model</b></font></td>
            <td><font size=1 face=verdana><b>Part</b></font></td>
        </tr>
<%
    dim strDeliverable
    dim strVersion
    dim strModel
    dim strVendor
    dim strPart
    dim strType
    dim strProduct
    dim strSQl
    dim strDeliverableID
    dim VersionArray
    dim arrRootIDs
    dim arrVersionIDs

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	
	set rs = server.CreateObject("ADODB.recordset")

    if strRoots <> "" then
        RootArray = split(strRoots,",")
        for each strItem in RootArray            
            arrRootIDs = split(strItem,":")            
		    rs.open "spGetProductDeliverableRootSummaryByID " & clng(arrRootIDs(0)) & "," & clng(arrRootIDs(1)),cn,adOpenForwardOnly
		    if (rs.EOF and rs.BOF) then
			    strProduct = ""
			    strDeliverable = ""
			    strVendor = ""
			    strDeliverableID = ""
                strType=""
		    else
			    strType=rs("TypeID") & ""
			    strProduct = rs("product") & ""
			    strDeliverable = rs("deliverable") & ""
			    strVendor = rs("Vendor") & ""
                strDeliverableID = rs("DeliverableID") & ""
		    end if
		    rs.Close

            if strVendor = "< Multiple Suppliers >" then
                strvendor = "[Multiple]"
            end if
            response.write "<td>" & strDeliverableID & "</td>"
            response.write "<td>" & strproduct & "</td>"
            response.write "<td>" & strVendor & "</td>"
            response.write "<td>" & strDeliverable & "</td>"
            response.write "<td>&nbsp;</td>"
            response.write "<td>&nbsp;</td>"
            response.write "<td>&nbsp;</td>"
            response.write "</tr>"
        next
    end if

    if strVersions <> "" then
        VersionArray = split(strversions,",")
        for each strItem in VersionArray            
            arrVersionIDs = split(strItem,":") 
		    rs.open "spGetProductDeliverableSummaryByID " & clng(arrVersionIDs(0)) & "," & clng(arrVersionIDs(1)),cn,adOpenForwardOnly
		    if (rs.EOF and rs.BOF) then
			    strProduct = ""
			    strDeliverable = ""
			    strVersion = ""
			    strPart = ""
			    strModel = ""
			    strVendor = ""
			    strDeliverableID = ""
                strType=""
		    else
			    strType=rs("TypeID") & ""
			    strProduct = rs("product") & ""
			    strDeliverable = rs("deliverable") & ""
			    strVersion = rs("version") & ""
			    if trim(rs("revision") & "") <> "" then
				    strVersion = strVersion & "," & rs("revision") & ""
			    end if
			    if trim(rs("pass") & "") <> "" then
				    strVersion = strVersion & "," & rs("pass") & ""
			    end if
			    strPart = rs("PartNumber") & ""
			    strModel = rs("ModelNumber") & ""
			    strVendor = rs("Vendor") & ""
                strDeliverableID = rs("DeliverableID") & ""
		    end if
		    rs.Close

            response.write "<td>" & strDeliverableID & "</td>"
            response.write "<td>" & strproduct & "</td>"
            response.write "<td>" & strVendor & "</td>"
            response.write "<td>" & strDeliverable & "</td>"
            response.write "<td>" & strversion & "</td>"
            response.write "<td>" & strModel & "&nbsp;</td>"
            response.write "<td>" & strpart & "&nbsp;</td>"
            response.write "</tr>"
        next
    end if
%>
        </table>
</td></tr>
</table>
<%end if%>

<input type="hidden" id="txtID" name="txtID" value="<%=request("ID")%>" />
<input type="hidden" id="txtNewValue" name="txtNewValue" value="<%=request("NewValue")%>" />
</form>

</BODY>
</HTML>
