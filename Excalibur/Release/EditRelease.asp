<%@ Language=VBScript %>

<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  

  Dim AppRoot
  AppRoot = Session("ApplicationRoot")
%>
	
<HTML>
<HEAD>
    <TITLE>Add Releases</TITLE>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--
        function ReleaseName_onchange() {
            if ($("#cboRelease").val() == "1C" || $("#cboRelease").val() == "2C" || $("#cboRelease").val() == "3C") {
                //Update Release format (Year + Cycle) for year post 2020
                if (($("#cboYear").val()) >= 2020) {
                    $("#lblReleaseName").text($("#cboYear").val().substr(2, 2) + $("#cboRelease").val().substr(1, 1) + $("#cboRelease").val().substr(0, 1));
                    $("#txtVersion").val($("#cboYear").val().substr(2, 2) + $("#cboRelease").val().substr(1, 1) + $("#cboRelease").val().substr(0, 1));
                }
                //Update Release format (Cycle + Year) for year before 2020
                else {
                    $("#lblReleaseName").text($("#cboRelease").val().substr(0, 2) + $("#cboYear").val().substr(2, 2));
                    $("#txtVersion").val($("#cboRelease").val().substr(0, 2) + $("#cboYear").val().substr(2, 2));
                }             
            }
            else {
                $("#lblReleaseName").text($("#cboRelease").val() + " " + $("#cboYear").val());
                $("#txtVersion").val($("#cboRelease").val() + " " + $("#cboYear").val());
            }
            return;
        }

        function cmdOK_onclick() {
            if (typeof ($("#txtBSOperation").val()) == "undefined" && $("#cboYear").val() == "") {
                    alert("Please select Year");
                    return;
            }

            if ($("#cboRelease").val() == "" || $("#cboYear").val() == "") {
                alert("Please select Release and Year");
                return;
            }

            if (!$("#txtDisclaimerNotes") || $("#txtDisclaimerNotes").val().length > 4000)
            {
                alert("Disclaimer Notes must be less than 4000 characters");
                return;
            }

            if ($("#txtDisclaimerNotes").val() === "" && $("#optActiveNotes").is(":checked")) {
                alert("State must be Inactive if no Disclaimer Notes provided");
                return;
            }

            var ID = 0;
            if ($("#txtID").val() != "") {
                ID = $("#txtID").val();
            }
            var ReleaseID = $("#txtReleaseID").val();
            var TypeID = $("#txtTypeID").val();
            var BSID = $("#txtBSID").val();
            var Notes = $("#txtDisclaimerNotes").val();
            var State = 0;
            if ($("#optActiveNotes").is(":checked")) {
                State = 1;
            }
            var User = $("#txtUser").val();
            var strRelease = $("#txtVersion").val();

            document.forms["frmUpdate"].action = "EditReleaseSave.asp?ID=" + ID + "&ProductTypeID=" + TypeID + "&BusinessSegmentID=" + BSID + "&Release=" + strRelease.replace(/^\s+|\s+$/gm, '') + "&DisclaimerNotes=" + Notes + "&State=" + State + "&User=" + User;
            document.forms["frmUpdate"].submit();
        }
    //-->
    </SCRIPT>
</HEAD>

<BODY bgcolor=ivory>
<%
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	dim strLoaded
    dim strTypeID
    dim strBusinessSegmentID
    dim strProductRelease
    dim strRelease
    dim strYear
	
	strLoaded = ""
    strTypeID = ""
    strProductRelease = ""
    strRelease = ""
    strYear = ""

    strTypeID = request("ProductTypeID")
    strBusinessSegmentID = request("BusinessSegmentID")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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
	
    dim CurrentUserFullName
	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
        CurrentUserFullName = rs("Name")
	end if
    
	rs.Close

    if trim(request("ID")) <> "0"  then
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spGetProductVersion_Pulsar"
		
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 

		strProductRelease = rs("ProductRelease") & ""
        dim Pos
        Pos = Instr(strProductRelease, ",")
        if Pos = 0 then
            if Len(strProductRelease) >= 4 then
                if Left(strProductRelease, 2) = "1C" OR Left(strProductRelease, 2) = "2C" OR Left(strProductRelease, 2) = "3C" then
                    strRelease = Left(strProductRelease, 2)
                    strYear = "20" + Right(strProductRelease, 2)
                elseif Left(strProductRelease, 3) = "NPI" OR Left(strProductRelease, 3) = "Jan" OR Left(strProductRelease, 3) = "Feb" OR Left(strProductRelease, 3) = "Mar" OR Left(strProductRelease, 3) = "Apr" OR Left(strProductRelease, 3) = "May" OR Left(strProductRelease, 3) = "Jun" OR _
                       Left(strProductRelease, 3) = "Jul" OR Left(strProductRelease, 3) = "Aug" OR Left(strProductRelease, 3) = "Sep" OR Left(strProductRelease, 3) = "Oct" OR Left(strProductRelease, 3) = "Nov" OR Left(strProductRelease, 3) = "Dec" then
                    strRelease = Left(strProductRelease, 3)
                    strYear = Right(strProductRelease, 4)
                else
                    strRelease = ""
                    strYear = ""            
                end if
            else
                strRelease = ""
                strYear = ""
            end if
        end if
		strProductRelease = ""
        
    	set cm=nothing	
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spPULSAR_Product_ListBusinessSegments"
		Set rs = cm.Execute 
        
    else
    	set cm=nothing	
        rs.open "spPULSAR_Product_ListBusinessSegments", cn 
    end if
%>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b>Add Release
</b></font>

<form id="frmUpdate" name="frmUpdate" method="post">

    <table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	    <tr>
		    <td width="150" nowrap valign=top><b>Release:</b>&nbsp;</td>
            <td width="100%"><label id="lblReleaseName"><%=strProductRelease%></label></td>
	    </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Product Type</b>&nbsp;</td>
            <td><label id="lblProductType">
                <%if strTypeID = "1" then%>PC (Notebook,Desktop,etc.)
                <%elseif strTypeID = "2" then%>Division Tools
                <%elseif strTypeID = "3" then%>Dock/Port Rep./Jacket
                <%elseif strTypeID = "4" then%>Other
                <%else%>PC (Notebook,Desktop,etc.)
                <%end if%>            
                </label>
            </td>

	    </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Business Segment</b>&nbsp;</td>
            <td><label id="lblBusinessSegment">
                <%  
                    do while not rs.eof
                    if trim(rs("BusinessSegmentID")) = trim(strBusinessSegmentID) then%>
                        <%=rs("name")%>
                    <%end if%>
                <%
                    rs.movenext
                    loop
                    rs.close
                %>
                </label>
            </td>        
	    </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Release</b><font color="red" size="1">&nbsp;*</font></td>
            <%  
                rs.open "spPULSAR_Product_ListBusinessSegments", cn 
                do while not rs.eof
                if trim(rs("BusinessSegmentID")) = trim(strBusinessSegmentID) then
                    if rs("Operation") <> "0" then %>
                        <font color="red" size="1">&nbsp;*</font>
                        <input type="hidden" id="txtBSOperation" name="txtBSOperation" value="<%=rs("Operation")%>">
            <%      end if  %>
            <%  end if
                rs.movenext
                loop
                rs.close
            %>
		    </td>
            <td>
                <select id="cboRelease" name="cboRelease" style="width: 200px; language="javascript" onchange="return ReleaseName_onchange()">
            <%  
                rs.open "spPULSAR_Product_ListBusinessSegments", cn 
                do while not rs.eof
                if trim(rs("BusinessSegmentID")) = trim(strBusinessSegmentID) then
                    if rs("BusinessId") = 2 then %>
                        <option selected value=""></option>
                        <option value="1C">1C</option>
                        <option value="2C">2C</option>
                        <option value="3C">3C</option>
            <%      else    %>
                        <option selected value=""></option>
                        <option value="NPI">NPI</option>
                        <option value="Jan">Jan</option>
                        <option value="Feb">Feb</option>
                        <option value="Mar">Mar</option>
                        <option value="Apr">Apr</option>
                        <option value="May">May</option>
                        <option value="Jun">Jun</option>
                        <option value="Jul">Jul</option>
                        <option value="Aug">Aug</option>
                        <option value="Sep">Sep</option>
                        <option value="Oct">Oct</option>
                        <option value="Nov">Nov</option>
                        <option value="Dec">Dec</option>
            <%      end if  %>
            <%  end if
                rs.movenext
                loop
                rs.close
            %>
                </select>
            </td>
        </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Year</b><font color="red" size="1">&nbsp;*</font></td>
            <td>
                <select id="cboYear" name="cboYear" style="width: 200px;" language="javascript" onchange="return ReleaseName_onchange()">
                    <option selected value=""></option>
                    <option value="2014">2014</option>
                    <option value="2015">2015</option>
                    <option value="2016">2016</option>
                    <option value="2017">2017</option>
                    <option value="2018">2018</option>
                    <option value="2019">2019</option>
                    <option value="2020">2020</option>
                    <option value="2021">2021</option>
                    <option value="2022">2022</option>
                    <option value="2023">2023</option>
                    <option value="2024">2024</option>
                    <option value="2025">2025</option>
                </select>
        </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>Disclaimer Notes</b><font color="red" size="1">&nbsp;*</font></td>
            <td colspan="3">
                <textarea id="txtDisclaimerNotes" name="txtDisclaimerNotes" style="width: 500px;
                    height: 120px;" language="javascript" ></textarea>
                <div style="margin-top: .5em">
                    <span class="p-Instruction">&nbsp;4000 maximum characters</span>
                </div>
            </td>
        </tr>
	    <tr>
		    <td width="150" nowrap valign=top><b>State</b><font color="red" size="1">&nbsp;*</font></td>
            <td>
                <input id="optActiveNotes" name="optState" type="radio" value="1" /> <font face="verdana" size="2">Active</font><br />
                <input id="optInactiveNotes" name="optState" type="radio" value="0" checked="checked" /> <font face="verdana" size="2">Inactive</font>
            </td>
        </tr>

    </table>

    <input type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
    <input type="hidden" id=txtTypeID name=txtTypeID value="<%=request("ProductTypeID")%>">
    <input type="hidden" id=txtBSID name=txtBSID value="<%=request("BusinessSegmentID")%>">
    <input type="hidden" id="txtVersion" name="txtVersion" value="<%=strProductRelease%>">
    <input type="hidden" id="txtUser" name="txtUser" value="<%=CurrentUserFullName%>">
</form>

<%
    'rs.Close
	set rs = nothing
	set cn = nothing
%>

</BODY>
</HTML>
