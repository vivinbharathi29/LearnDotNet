<%@ Language=VBScript %>
	<%
	if trim(request("cboFormat"))= "1" then
		Response.ContentType = "application/vnd.ms-excel"
	elseif trim(request("cboFormat"))= "2" then
		Response.ContentType = "application/msword"
    else
      Response.Buffer = True
      Response.ExpiresAbsolute = Now() - 1
      Response.Expires = 0
      Response.CacheControl = "no-cache"
	end if	
	  
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
    <title></title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" language="javascript">
<!--
    <!-- #include file = "../_ScriptLibrary/sort.js" -->

    function HeaderMouseOver(){
        window.event.srcElement.style.cursor="hand";
        window.event.srcElement.style.color="red";
    }

    function HeaderMouseOut(){
        window.event.srcElement.style.color="black";
    }

    function ToggleProductDisplay(ID){
        if (window.document.all("divWhereUsed" + ID).style.display == "none")
            window.document.all("divWhereUsed" + ID).style.display = "";
        else
            window.document.all("divWhereUsed" + ID).style.display = "none";
    }


    //-->
</script>
<style>
    BODY {
        FONT-SIZE: x-small;
        FONT-FAMILY: Verdana;
    }

    A:visited {
        COLOR: blue;
    }

    A:hover {
        COLOR: red;
    }

    td {
        FONT-SIZE: xx-small;
        FONT-FAMILY: Verdana;
    }
</style>
</head>
<body bgcolor="Ivory">
    <%
    
	dim strDelName
	dim strprodname
	dim strRootID
	dim strProductID
    dim strReleaseID
	dim cn
	dim rs
	dim i
	dim strCategoryID
	dim blnLoadFailed
	dim strVersionsLoaded
	dim LockedCount
	
	strVersionsLoaded = ""
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	blnLoadFailed = false

	set rs = server.CreateObject("ADODB.recordset")
	if request("ProdRootID") <> "" then
		rs.Open "spGetProductIDbyProdRoot " & clng(request("ProdRootID")) & "," & clng(request("ProductDeliverableReleaseID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProductID = 0
            strReleaseID = 0
		else
			strProductID = rs("ID")
            strReleaseID = rs("ReleaseID")
		end if
		rs.Close
	else
		strProductID = request("ProductID")
        strReleaseID = request("ProductDeliverableReleaseID")
	end if
	

	rs.Open "spGetProductVersionName " & clng(strProductID) & "," & clng(strReleaseID),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strProdName = ""
		blnLoadFailed = true
	else
		strprodName = rs("name") & ""
	end if
	
	rs.Close
	
		
	if request("RootID") = "" or trim(request("RootID")) = "0" then
		rs.Open "spGetRootID " & clng(request("VersionID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strRootID = 0
			blnLoadFailed = true
		else
			strRootID = rs("ID")
		end if
	else
		strRootID = request("RootID")
	end if

	rs.Open "spGetDeliverableRootName " & clng(strRootID),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		strDelName = ""
		strType = ""
		strCategoryID = ""
		blnLoadFailed = true
	else
		strDelName = rs("name") & ""
		strType = rs("TypeID") & ""
		strCategoryID = rs("CategoryID") & ""
		strActive = rs("Active") & ""
	end if
	
	rs.Close

if trim(request("cboFormat"))= "" then 
    strRowBGColor = "cornsilk"
else
    strRowBGColor = ""
end if

	if blnLoadFailed then
		Response.Write "<BR><BR><font size=2 face=verdana>Not enough information supplied to display this page.</font>"
	else

		rs.Open "spListVersions4Supporting " & clng(strProductID) & "," & clng(request("RootID")) & "," & clng(strReleaseID), cn, adOpenForwardOnly

		if rs.EOF and rs.BOF then
			Response.Write "<BR><BR><font size=2 face=verdana>No deliverable versions found.</font>"
		else
            if trim(request("cboFormat")) = "" then

%>

    <h3>Select Supported Versions.</h3>
    <form id="frmTarget" name="frmTarget" action="AdvancedSupportSave.asp" method="post">
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
                <td><font face="verdana" size="2"><b><%=strDelName & " for " & strProdName%></font></td>
                <td align="right"><b>Export:</b> <a target="_blank" href="AdvancedSupportMain.asp?<%=request.querystring%>&cboFormat=1">Excel</a></td>
            </tr>
        </table>

        <table width="100%">
            <tr>
                <td><font color="red">Note: Only active versions are displayed.</font></td>
                <%if trim(strType) = "1" then%>
                <td align="right">
                    <table border="1" cellpadding="2" cellspacing="0" bordercolor="Tan">
                        <tr bgcolor="cornsilk">
                            <td><b>Product&nbsp;Status&nbsp;Key:</b></td>
                            <td><font color="red">Dropped</font></td>
                            <td><font color="green">QComplete</font></td>
                            <td><font color="blue">Investigating</font></td>
                            <td><font color="gold">QHold</font>
                                <td><font color="red">Fail</font></td>
                            <td><font color="Thistle">Not Used</font></td>
                            <td><font color="black">All Others</font></td>
                        </tr>
                    </table>
                </td>
                <%else%>
                <td align="right">
                    <table border="1" cellpadding="2" cellspacing="0" bordercolor="Tan">
                        <tr bgcolor="cornsilk">
                            <td><b>Product&nbsp;Status&nbsp;Key:</b></td>
                            <td><font color="Green">Targeted</font></td>
                            <td><font color="black">All Others</font></td>
                        </tr>
                    </table>
                </td>
                <%end if%>
        </table>
        <%end if %>
        <table id="MyTable" width="100%" bgcolor="<%=strRowBGColor%>" border="1" cellspacing="0" cellpadding="2" bordercolor="tan">
            <thead>
                <tr>
                    <%	if trim(request("cboFormat"))= "" then %>
                    <td><b>Supported</b></font></td>
                    <%end if%>
                    <td onclick="SortTable( 'MyTable', 1 ,1,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">ID</b></font></td>
                    <%if trim(strType) = "1" then%>
                    <td onclick="SortTable( 'MyTable', 2 ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">HW,FW,Rev</b></td>
                    <td onclick="SortTable( 'MyTable', 3 ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Vendor</b></td>
                    <td onclick="SortTable( 'MyTable', 4 ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Model</b></td>
                    <%VersionRows=4%>
                    <%else%>
                    <td onclick="SortTable( 'MyTable', 2 ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Version</b></td>
                    <%VersionRows=1%>
                    <%end if%>
                    <td onclick="SortTable( 'MyTable', <%=1+VersionRows%> ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">Part</b></td>
                    <td onclick="SortTable( 'MyTable', <%=2+VersionRows%> ,0,1);"><b onmouseout="javascript: HeaderMouseOut();" onmouseover="javascript: HeaderMouseOver();">EOA</b></td>
                    <%	if trim(request("cboFormat"))<> "" then %>
                    <td width="1000"><b>Where Used</b></td>
                    <%elseif trim(strType) = "1" then %>
                    <td nowrap><b>Where Used</b>&nbsp;-&nbsp;<a target="_blank" href="HardwareMatrix.asp?lstcategory=<%=strCategoryID%>&lstRoot=<%=strRootID%>">View Matrix</a></td>
                    <%else%>
                    <td><b>Where Used</b></td>
                    <%end if%>
                </tr>
            </thead>
            <%
		EOLCount = 0
		LockedCount = 0
		ProductCount = 0
		do while not rs.EOF	
			if rs("Active") = 0 then
				EOLCount = EOLCOunt + 1
			else
				strVersion = rs("Version") & ""
				if rs("Revision") & "" <> "" then
					strVersion = strVersion & "," & rs("Revision")
				end if
				if rs("Pass") & "" <> "" then
					strVersion = strVersion & "," & rs("Pass")
				end if
				if rs("Active") = 0 then
					strEOL = "EOL"
				elseif not isnull(rs("EOLDate")) then
					strEOL = rs("EOLDate")
				else
					strEOL = "&nbsp;"
				end if
		
				set rs2 = server.CreateObject("ADODB.recordset")
	
				strSQL = "spGetTargetedProductsForVersion " & rs("VersionID")
				rs2.Open strSQL, cn, adOpenStatic
	  
'				strWhereUsed = "<TABLE cellpadding=0 cellspacing=0>"
				strWhereUsed = ""
				strLastFamily = ""
				i=0
				ProductCount = 0
				do while not rs2.EOF
				    if trim(request("cboFormat")) <> "" then
				        strWhereUsed = strWhereUsed & ", " & rs2("Family") '& "&nbsp;" & rs2("Version")
				    else
					    if strLastFamily <> rs2("Family") then
						    if strLastFamily = "" then
							    'strWhereUsed = strWhereUsed & "<TR><Td nowrap><b>"					
							    strWhereUsed = strWhereUsed & "<NOBR><b>"					
						    else
							    'strWhereUsed = strWhereUsed & "</Td></TR><TR><Td nowrap><b>"
							    strWhereUsed = strWhereUsed & "</b></NOBR><BR><NOBR><b>"
						    end if
						    i=0
						    strLastFamily = rs2("Family")
					    end if
					    if i <> 0 then
						    strWhereUsed = strWhereUsed & ", " 
					    end if
					    if trim(strType) = "1" then
						    if isnull(rs2("TestStatus")) then
							    strWhereUsed = strWhereUsed & "<font title=""Not Used"" color=thistle>"
						    else
							    strWhereUsed = strWhereUsed & "<font title=""" & rs2("TestStatus") & """ color=" & rs2("BGCOLOR") & ">"
						    end if
						    strWhereUsed = strWhereUsed & rs2("Family") '& "&nbsp;" & rs2("Version")
						    'strWhereUsed = strWhereUsed & "&nbsp;("& rs2("TestStatus") & ")"
						    strWhereUsed = strWhereUsed & "</font>"
					    else
						    if rs2("Targeted") then
							    strWhereUsed = strWhereUsed & "<font color=green>"
						    else
							    strWhereUsed = strWhereUsed & "<font color=black>"
						    end if
						    strWhereUsed = strWhereUsed & rs2("Family") '& "&nbsp;" & rs2("Version")
						    strWhereUsed = strWhereUsed & "</font>"
					    end if
                    end if
					i=i+1
					ProductCount = ProductCount + 1
					rs2.MoveNext
				loop
				rs2.close
				set rs2=nothing
    		    if trim(request("cboFormat")) <> "" then
                    if strWhereUsed <> "" then
                        strWhereUsed = mid(strWhereUsed,3)
                    end if
	            else
				    strWhereUsed = strWhereUsed & "</b></NOBR>"
				    if ProductCount > 0 then
				        strWhereUsed = "Products&nbsp;Supported:&nbsp;<a href=""javascript: ToggleProductDisplay(" & rs("VersionID") & ");"">" & ProductCount & "</a><BR><span style=""display:none"" ID=""divWhereUsed" & trim(rs("VersionID")) & """>" & strWhereUsed & "</span>"
                    else
                        strWhereUsed = "&nbsp;"
                    end if				
			
                end if
		%>
            <tr>
                <%	if trim(request("cboFormat"))= "" then %>
                <%if instr(rs("Location"),"Workflow Complete")> 0 and not isnull(rs("ProdDelID")) then%>
                <td title="Can only be removed by the Product Team" bgcolor="LightSteelBlue" valign="top" align="left">Yes (Workflow Complete)</td>
                <%  LockedCount=LockedCount+1%>
                <%elseif not isnull(rs("ProdDelID")) then%>
                <td valign="top" align="center">
                    <input checked type="checkbox" id="chkSupport" name="chkSupport" style="width: 14px; height: 14px" size="14" value="<%=rs("VersionID")%>"></td>
                <% strVersionsLoaded = strVersionsLoaded & ", " & rs("VersionID")%>
                <%else%>
                <td valign="top" align="center">
                    <input type="checkbox" id="chkSupport" name="chkSupport" style="width: 14px; height: 14px" size="14" value="<%=rs("VersionID")%>"></td>
                <%end if%>
                <%end if%>
                <td valign="top"><%=rs("VersionID")%></td>
                <td nowrap width="100" valign="top"><%=strVersion%></td>
                <%if trim(strType) = "1" then%>
                <td valign="top"><%=rs("Vendor")%>&nbsp;</td>
                <td valign="top"><%=rs("ModelNumber")%>&nbsp;</td>
                <%end if%>
                <td nowrap valign="top"><%=rs("PartNumber")%>&nbsp;</td>
                <td valign="top"><%=strEOL%>&nbsp;</td>
                <td valign="top"><%=strWhereUsed%></td>

            </tr>
            <%
			end if
			rs.MoveNext
		loop
		rs.Close
	%>
        </table>
        <%
	if len(strVersionsLoaded) > 0 then
		strVersionsLoaded = mid(strVersionsLoaded,3)
	end if

%>
        <input type="hidden" id="txtProductID" name="txtProductID" value="<%=clng(strProductID)%>">
        <input type="hidden" id="txtReleaseID" name="txtReleaseID" value="<%=clng(strReleaseID)%>">
        <input type="hidden" id="txtProductDeliverableReleaseID" name="txtProductDeliverableReleaseID" value="<%=clng(request("ProductDeliverableReleaseID"))%>">
        <input type="hidden" id="txtRootID" name="txtRootID" value="<%=clng(strRootID)%>">
        <input type="hidden" id="txtVersionsLoaded" name="txtVersionsLoaded" value="<%=strVersionsLoaded%>">
        <input type="hidden" id="txtLockedCount" name="txtLockedCount" value="<%=LockedCount%>">
        <input type="hidden" id="txtRowID" name="txtRowID" value="<%=request("RowID")%>" />
    </form>


    <%
	end if
end if
set rs= nothing
set cn=nothing
%>
</body>
</html>