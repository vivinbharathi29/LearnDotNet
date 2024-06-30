<%@ Language=VBScript %>
<%
	  Response.Buffer = True
	  Response.ExpiresAbsolute = Now() - 1
	  Response.Expires = 0
	  Response.CacheControl = "no-cache"
%>
<html>
<head>
<META name=VI60_defaultClientScript content=JavaScript>
<title>Product Deliverable Alert Details - Confidential</title>
<STYLE>
td
{
    FONT-FAMILY: verdana;
    FONT-SIZE: xx-small;	
}
A:link,A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
}


function ROW_onmouseover() {
	event.srcElement.style.cursor="hand";

	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;

	if (srcElem.className =="Row")
		srcElem.style.backgroundColor = "Thistle";


}

function ROW_onmouseout() {
	var srcElem = window.event.srcElement;

	//crawl up the tree to find the table row
	while (srcElem.tagName != "TR" && srcElem.tagName != "TABLE")
		srcElem = srcElem.parentElement;

	if(srcElem.tagName != "TR") return;

	if (srcElem.className =="Row")
		srcElem.style.backgroundColor = "White";
	
}


function DelROW_onclick(ID, RootID){
	var strResult;
	strResult = window.showModalDialog("WizardFrames.asp?Type=1&RootID=" + RootID + "&ID=" + ID,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 

}

//-->
</SCRIPT>
</head>
<body LANGUAGE=javascript onload="return window_onload()">
<p align=center>
<font face=verdana size=3>
<%

    dim ProdID
    dim VersionID
    dim RootID
    dim strProductName
    dim strDeliverableName
    dim SEPMID
    dim strVersion
    dim strGeneralHeader
    
    strGeneralHeader = ""
    
    ProdID = clng(request("ProdID"))
    
    if trim(request("VersionID")) = "" then
        VersionID = 0
    else
        VersionID = clng(request("VersionID"))
    end if
    if trim(request("RootID")) = "" then
        RootID = 0
    else
        RootID = clng(request("RootID"))
    end if
    
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function
	
	Function GetWeekIndex(StartWeek, StartYear, GetWeek,GetYear)
		if StartYear = GetYear then
			GetWeekIndex = (GetWeek - StartWeek)
		else
			GetWeekIndex = GetWeek + (52-StartWeek) + (52* (GetYear - StartYear-1))
		end if
	end function

	dim cm
	dim cn
	dim p
	dim rs

  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout =120
	cn.ConnectionTimeout =120
	cn.Open

  'Create a recordset
  set rs = server.CreateObject("ADODB.recordset")
  rs.ActiveConnection = cn
	dim CurrentUser
	dim CurrentUSerID
	dim CurrentUserPartner
	'Get User
	dim CurrentDomain
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetUserInfo"
	

	Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
	p.Value = Currentuser
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
	p.Value = CurrentDomain
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	set cm=nothing
	
	CurrentUserID = 0
	if rs.EOF and rs.BOF then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=0"	
	else
		CurrentUserID = rs("ID")
		CurrentUserPartner = rs("PartnerID")
	end if		
	rs.Close
	
	if ProdID <> "" then
		rs.Open "spGetProductVersionName " & ProdID,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strproductName = ""
			SEPMID =  ""
		else
			strproductName = rs("Name") & ""
			SEPMID = rs("SEPMID") & ""
			'Verify Access is OK
			if trim(CurrentUserPartner) <> "1" then
				if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
					rs.Close
					set rs = nothing
					set cn=nothing
					
					Response.Redirect "../NoAccess.asp?Level=0"
				end if
			end if		
		end if
		rs.close
	else
		strproductName = ""
	end if

	if strProductName <> "" and trim(VersionID) <> "" and trim(VersionID) <> "0" then
		rs.Open "spGetDeliverableVersionProperties " & VersionID,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
            strDeliverableName = ""
            strVersion = ""
        else
            strDeliverableName = rs("name") & ""
            strVersion = rs("version") & ""
            if trim(rs("Revision") & "") <> "" then
                strVersion = strVersion & "," & rs("Revision") 
            end if
            if trim(rs("Pass") & "") <> "" then
                strVersion = strVersion & "," & rs("Pass") 
            end if
        end if
        rs.close
    end if	

	if strProductName <> "" and trim(RootID) <> "" and trim(RootID) <> "0" and(trim(VersionID) = ""  or trim(VersionID) = "0") then
		rs.Open "spGetDeliverableRootName " & RootID,cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
            strDeliverableName = ""
        else
            strDeliverableName = rs("name") & ""
        end if
        rs.close
    end if	
    
	if trim(strProductName) = "" then
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to find the selected product.</font>"
		set rs = nothing
		set cn = nothing
	elseif trim(strDeliverableName) = "" then
		Response.Write "<BR><BR><font size=2 face=verdana>Unable to find the selected product.</font>"
		set rs = nothing
		set cn = nothing
	else
	    response.write "<font family=Verdana size=2><b>"
	    if trim(strversion) <> "" then
            response.write strproductname & " Alerts for " &  strDeliverablename & " [" & strversion & "]<BR><BR></p>"
        else
            response.write strproductname & " Alerts for " &  strDeliverablename & "<BR><BR></p>"
        end if
	    response.write "</b></font>"
        strGeneralHeader = "<BR><font size=2 face=verdana><B>General Alerts</B></font><table ID=AlertTable border=1 borderColor=gainsboro cellpadding=2 cellSpacing=0 width=""100%""><tr  bgcolor=beige><TD><b>Alert&nbsp;&nbsp;&nbsp;&nbsp;</b></TD><TD width=""100%""><b>Details</b></TD></TR>"

        if trim(RootID) <> "" and trim(RootID) <> "" and (trim(VersionID) = "" or trim(VersionID) = "0") then
            if strGeneralHeader <> "" then
                response.Write strGeneralHeader
                strGeneralHeader = ""
            end if
            %>
            <tr bgcolor=ivory> 
                <TD nowrap>Choose Versions</TD>
                <TD width="100%">No versions have been targeted for this deliverable.</TD>
            </TR>
            <%
        else
            rs.open "spListDeliverableAlertDetails " & ProdID & "," & VersionID,cn,adOpenForwardOnly
            if not (rs.eof and rs.bof) then
                if rs("targeted") and trim(rs("Preinstall") & "") <> "True" and trim(rs("Patch") & "") = "0" and trim(rs("preload") & "") <> "True" and trim(rs("DropInBox") & "") <> "True" and rs("web") & "" <> "True" and trim(rs("SelectiveRestore") & "") <> "True" and trim(rs("arcd") & "") <> "True" and trim(rs("drdvd") & "") <> "True" and trim(rs("racd_EMEA") & "") <> "True" and trim(rs("racd_APD") & "") <> "True" and trim(rs("Racd_Americas") & "") <> "True" and trim(rs("doccd") & "") <> "True" and trim(rs("oscd") & "") <> "True" then
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                %>
		            <tr  bgcolor=ivory> 
		                <TD>No&nbsp;Distributions&nbsp;&nbsp;</TD>
		                <TD width="100%">The deliverable is targeted but has no distributions defined.</TD>
		            </TR>
                <%
                end if

                if trim(rs("CertificationStatus") & "") = "0" then
                    strWHQLStatus = "Required"
                elseif trim(rs("CertificationStatus") & "") = "1" then
                    strWHQLStatus = "Submitted"
                elseif trim(rs("CertificationStatus") & "") = "2" then
                    strWHQLStatus = "Approved"
                elseif trim(rs("CertificationStatus") & "") = "3" then
                    strWHQLStatus = "Failed"
                elseif trim(rs("CertificationStatus") & "") = "4" then
                    strWHQLStatus = "Waiver"
                else
                    strWHQLStatus = "Required"
                end if
                
                if trim(rs("LevelID") & "") = "3" or trim(rs("LevelID") & "") = "9" or trim(rs("LevelID") & "") = "10" or trim(rs("LevelID") & "") = "11" then
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                %>
		            <tr  bgcolor=ivory> 
		                <TD>Alpha&nbsp;&nbsp;</TD>
		                <TD width="100%">The deliverable has an "<%=rs("BuildLevel")%>" build level</TD>
		            </TR>
                <%
		        end if

                if trim(rs("LevelID") & "") = "4" or trim(rs("LevelID") & "") = "12" or trim(rs("LevelID") & "") = "13" or trim(rs("LevelID") & "") = "14" then
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                %>
		            <tr  bgcolor=ivory> 
		                <TD>Beta&nbsp;&nbsp;</TD>
		                <TD width="100%">The deliverable has an "<%=rs("BuildLevel")%>" build level</TD>
		            </TR>
                <%
		        end if

                if trim(rs("CertificationStatus") & "") <> "2" and trim(rs("CertificationStatus") & "") <> "4" and trim(rs("CertRequired") & "") = "1" and (trim(rs("LevelID") & "") = "7" or trim(rs("LevelID") & "") = "15" or trim(rs("LevelID") & "") = "16" or trim(rs("LevelID") & "") = "17"  or trim(rs("LevelID") & "") = "18") then 'RC or GM, Requires WHQL, WHQL Status <> 2 or 4
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                %>
		            <tr bgcolor=ivory> 
		                <TD>WHQL&nbsp;Issue&nbsp;&nbsp;</TD>
		                <TD width="100%">The Current WHQL status for this deliverable is "<%=strWHQLStatus%>". It should be "Approved" or "Waiver".</TD>
		            </TR>
                <%
		        end if

                if isdate(rs("EOLDate")) and clng(rs("ProductStatusID")) < 4  then
    		        if datediff("d",rs("EOLDate"),now) < 365 then
                        if strGeneralHeader <> "" then
                            response.Write strGeneralHeader
                            strGeneralHeader = ""
                        end if
                %>
		                <tr bgcolor=ivory> 
		                    <TD nowrap><%="Use Until: " & rs("EOLDate")%>&nbsp;</TD>
            		        <TD width="100%">The developer indicates that this deliverable can not be used after <%=rs("EOLDate") & ""%>.</TD>
		                </TR>
                <%
                    end if
		        end if
    		    
		        if trim(rs("DeveloperNotificationStatus") & "") = "2" or trim(rs("DeveloperNotificationStatus") & "") = "0" then
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                    if trim(rs("DeveloperNotificationStatus") & "") = "2" then
                %>
		                <tr bgcolor=ivory> 
		                    <TD>Dev:&nbsp;Disapproved</TD>
            		        <TD width="100%">The developer indicates that this version should not be used on this product.</TD>
		                </TR>
                <%
                    else
                %>
		                <tr bgcolor=ivory> 
		                    <TD>Dev:&nbsp;Awaiting&nbsp;Approval</TD>
            		        <TD width="100%">The developer has not approved or disapproved this version for use on this product.</TD>
		                </TR>
                <%
                    
                    end if
    		        
		        end if

		        if trim(rs("Location") & "") <> "Workflow Complete" then
                    if strGeneralHeader <> "" then
                        response.Write strGeneralHeader
                        strGeneralHeader = ""
                    end if
                %>
		                <tr bgcolor=ivory> 
		                    <TD nowrap><%=replace(replace(rs("Location")& "","Workflow Complete","Complete")," ","&nbsp;")%></TD>
            		        <TD width="100%">This deliverable is not workflow complete.</TD>
		                </TR>
                <%
    		        
		        end if
    		    
            
            end if
            rs.close
        end if
%>
        </table><BR><BR>

		<%
            if trim(RootID) <> "" and trim(RootID) <> "" and (trim(VersionID) = "" or trim(VersionID) = "0") then
		        rs.open "spListOTS4Root " &  RootID & ",0", cn,adOpenForwardOnly
		    else
		        rs.open "spListOTS4Version " &  VersionID & ",0", cn,adOpenForwardOnly
		    end if
		    
            if not (rs.eof and rs.bof) then
		  %>
                <font face=verdana size=2><b>OTS Alerts&nbsp;&nbsp;</b></font>
		        <table width="100%" cellspacing=0 cellpadding=2 border=1 bordercolor=gainsboro bgcolor=ivory>
		            <tr bgcolor=beige>
		                <td><b>ID</b></td>
		                <td><b>Product</b></td>
		                <td><b>Version</b></td>
		                <td><b>Pr</b></td>
		                <td><b>State</b></td>
		                <td width=100><b>Summary</b></td>
		            </tr>
        <%
            end if
		    do while not rs.eof
		%>
	            <tr>
	                <td><a target="_blank" href="search/ots/Report.asp?txtReportSections=1&txtObservationID=<%=rs("ObservationID") & ""%>"><%=rs("ObservationID") & ""%></a></td>
	                <td nowrap><%=rs("Product") & ""%></td>
	                <td><%=rs("OTSComponentVersion")%></td>
	                <td><%=rs("Priority")%></td>
	                <td><%=rs("State")%></td>
	                <td width="100%"><%=rs("Summary")%></td>
	            </tr>
<%
                rs.movenext
            loop
            rs.close

%>
		        </table>
		    </TD>
		</TR>

	    </TABLE>

<%
	    set rs = nothing
	    set cn = nothing

%>

        <br>
        <br>
        <font size="1">Report Generated <%=formatdatetime(date(),vblongdate) %></font>
        <br>
        <br>
        <font Size="2" Color="red"><p><strong>HP&nbsp;Confidential</strong></p></font>
<%  end if
%>

</body>
</html>
