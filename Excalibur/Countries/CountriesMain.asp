<%@ language="VBScript" %>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link href="../style/TriStateCheckBox/Tri_State_Checkbox.css" rel="stylesheet" />
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>

    <script id="clientEventHandlersJS" language="javascript">
        $(document).ready(function () {
            var releaseIDs = $("#hidProductReleases").val().split(',');
            var msg = "";
            for (var i = 0; i < releaseIDs.length; i++) {
                if ($("#hidPDDLocked").val().split(',')[i] == "1") {
                    msg = msg + $("#hidProductReleasesName").val().split(',')[i] + " ";
                }
            }
            $("#LockedRelease").text("PRL-Locked Release(s): " + msg);
        });

        function chkAll_onclick(strRegion) {
            var isPulsarProduct = document.getElementById("IsPulsarProduct").value;
            var productReleases = document.getElementById("hidProductReleases").value;
            var selectedRegions = document.all("chkSelected" + strRegion);
            if (document.all("chkAll" + strRegion).checked) {
                for (var i = 0; i < selectedRegions.length; i++) {
                    frmCountries("chkSelected" + strRegion)(i).checked = true;
                    document.all("Row" + frmCountries("chkSelected" + strRegion)(i).value).bgColor = "lightsteelblue";
                    if (isPulsarProduct == "1") {
                        var releases = productReleases.split(",");
                        ID = frmCountries("chkSelected" + strRegion)(i).value;
                        for (var j = 0; j < releases.length; j++) {
                            var obj = document.getElementById("chk" + ID + "-" + releases[j]);
                            if (releases[j] != "" && (obj.className == "TriStateCheckbox")) {
                                obj.className = "TriStateCheckboxChecked";
                                document.getElementById("txtnew" + frmCountries("chkSelected" + strRegion)(i).value + "-" + releases[j]).value = "1";
                            }
                        }
                    }
                }
            }
            else {
                for (var i = 0; i < selectedRegions.length; i++) {
                    frmCountries("chkSelected" + strRegion)(i).checked = false;
                    document.all("Row" + frmCountries("chkSelected" + strRegion)(i).value).bgColor = "ivory";
                    if (isPulsarProduct == "1") {
                        var releases = productReleases.split(",");
                        ID = frmCountries("chkSelected" + strRegion)(i).value;
                        for (var j = 0; j < releases.length; j++) {
                            var obj = document.getElementById("chk" + ID + "-" + releases[j]);
                            if (releases[j] != "" && (obj.className == "TriStateCheckboxChecked")) {
                                obj.className = "TriStateCheckbox";
                                document.getElementById("txtnew" + frmCountries("chkSelected" + strRegion)(i).value + "-" + releases[j]).value = "0";
                            }
                        }
                    }
                }
            }
        }

        function chkMilestone_onclick(ID, RowID) {
            var isPulsarProduct = document.getElementById("IsPulsarProduct").value;
            var productReleases = document.getElementById("hidProductReleases").value;
            if (event.srcElement.checked) {
                document.all("Row" + ID).bgColor = "lightsteelblue";
                if (isPulsarProduct == "1") {
                    var releases = productReleases.split(",");
                    for (var i = 0; i < releases.length; i++) {
                        var obj = document.getElementById("chk" + ID + "-" + releases[i]);
                        if (releases[i] != "" && (obj.className == "TriStateCheckbox")) {
                            obj.className = "TriStateCheckboxChecked";
                            document.getElementById("txtnew" + ID + "-" + releases[i]).value = "1";
                        }
                    }
                }
            }
            else {
                document.all("Row" + ID).bgColor = "ivory";
                if (isPulsarProduct == "1") {
                    var releases = productReleases.split(",");
                    for (var i = 0; i < releases.length; i++) {
                        var obj = document.getElementById("chk" + ID + "-" + releases[i]);
                        if (releases[i] != "" && (obj.className == "TriStateCheckboxChecked")) {
                            obj.className = "TriStateCheckbox";
                            document.getElementById("txtnew" + ID + "-" + releases[i]).value = "0";
                        }
                    }
                }
            }

        }

        function chkrelease_onclick(obj, newtextboxobj, row) {
            if (obj.className == "TriStateCheckbox") {
                obj.className = "TriStateCheckboxChecked";
                document.getElementById(newtextboxobj).value = "1";
                row.getElementsByTagName("input")[0].checked = true;
            }
            else if (obj.className == "TriStateCheckboxPartial") {
                obj.className = "TriStateCheckboxChecked";
                document.getElementById(newtextboxobj).value = "1";
            }
            else if (obj.className == "TriStateCheckboxChecked") {
                obj.className = "TriStateCheckbox";
                document.getElementById(newtextboxobj).value = "0";
            }
            return false;
        }

        function VerifySave() {
            var success = true;
            var releaseIDs = $("#hidProductReleases").val().split(',');
            if ($("#cboDcr").val() == 0) {
                $("input[name='chkCountriesRelease']").each(function () {
                    for (var i = 0; i < releaseIDs.length; i++) {
                        if (this.id.split('-')[1] == releaseIDs[i]
                            && $("#hidPDDLocked").val().split(',')[i] == "1"
                            && $("#" + this.id.replace('chk', 'txtold')).val() != $("#" + this.id.replace('chk', 'txtnew')).val()) {
                            success = false;
                            return false;
                        }
                    }
                });

                if (!success) {
                    alert("Please Select the appropriate DCR");
                    $("#cboDcr").focus();
                    return false;
                }
            }
            return success;
        }
    </script>
</head>
<body bgcolor="White">

    <form id="frmCountries" action="CountriesSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method="post">
        <p><font id="LockedRelease" size="2" face="verdana" style="color: red"></font></p>

        <%	

	dim cn 
	dim rs
	dim p
	dim cm
	dim CurrentUser
	dim i
	dim strLastRegion
	dim strRegion
	dim strRegionID
    dim intIsPulsarProduct : intIsPulsarProduct = 0   

if request("ProdID") = "" Or Request("ProdBrandID") = "" then
	Response.Write "<P><font size=2 face=verdana>Not enough information to display this page</font></P>"
	Response.Write Request("ProdID") & "<BR>"
	Response.Write Request("ProdBrandID") & "<BR>"
else
    intIsPulsarProduct = Request("IsPulsarProduct")
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
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

	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../NoAccess.asp?Level=1"
	else
		CurrentUserPartner = rs("PartnerID")
	end if 
	rs.Close
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetProductVersionName"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ProdID")
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	'rs.open "spGetProductVersionName " & request("ProdID"),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>Unable to find the requested product</font>"
		rs.Close
	else
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../NoAccess.asp?Level=1"
			end if
		end if

		Dim PDDLocked
		PDDLocked = CheckPDDLockBit(request("ProdID"))

		If PDDLocked = -500 Then
			Response.Write "<font face=verdana size=3><b>Product Release Not Found</b></font><BR><BR>"
			Response.End
		ElseIf PDDLocked Then
			'Display the DCR List
			Response.Write "<p><font face=verdana size=2><b>Choose the Applicable DCR for PRL-Locked Release(s)</b></font><br>"
			Response.Write "<SELECT id=cboDcr name=cboDcr>"
			FillDcrList(request("prodid"))
			Response.Write "</SELECT></p>"
		End If
	
		Response.Write "<font size=3 face=verdana><b>Select " & rs("name") & " Countries</b></font><BR><BR>"
		rs.Close
		
        dim strReleaseInfo
        dim strRelease
        dim intTotalLocalization : intTotalLocalization = 0
        dim intReleaseSelected : intReleaseSelected = 0
        dim strCheckBoxID
        dim arrayRelease
        dim strReleaseIDs : strReleaseIDs = ","
        dim ReleaseNames : ReleaseNames = ","
        dim PDDLockeds : PDDLockeds = ","
        
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "usp_ListProductBrandAllCountries"		

		Set p = cm.CreateParameter("@p_ProductBrandID", 3, &H0001)
		p.Value = request("ProdBrandID")
		cm.Parameters.Append p
	
        Set p = cm.CreateParameter("@p_IsPulsarProduct", 3, &H0001)
		p.Value = intIsPulsarProduct
		cm.Parameters.Append p

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

		'rs.Open "spListCountries4ProductAll " & request("ProdID"),cn,adOpenForwardOnly
		Response.Write "<Table id=""localizedcountrylist"" width=100% border=0 cellpadding=1 cellspacing=0>"'<TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Country</b></font></TD></tr>"
		i=0
		strLastRegion = ""
		do while not rs.EOF 
			if strLastRegion <> rs("Region") & "" then
                if not IsNull(rs("Region")) then
				    strRegion = mid(left(rs("Region"),len(rs("Region"))-1),2)
				    select case lcase(trim(strRegion))
				    case "na"
					    strRegionID = 1
				    case "la"
					    strRegionID = 2
				    case "emea"
					    strRegionID = 3
				    case "apj"
					    strRegionID = 4
				    case else
					    strRegionID = 0
				    end select
				    Response.Write "<TR bgcolor=MediumAquamarine><td width=20 style=""BORDER-TOP: gray thin solid""><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL" & trim(strRegionID) & " name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick('" & strRegionID & "')""</td><TD colspan=""2"" style=""BORDER-TOP: gray thin solid""><font size=2 face=verdana><b>" & strRegion & "</b></font></TD></TR>"
				    strLastRegion = rs("Region") & ""
                end if
			end if
			if isnull(rs("ID")) then
				Response.Write "<TR bgcolor=Ivory ID=Row" & rs("CountryID") & "><TD style=""BORDER-TOP: gray thin solid; width:3%""><INPUT value=""" & rs("CountryID") & """ style=""width:16;height:16;"" type=""checkbox"" id=chkSelected" & trim(strRegionID) & " name=chkSelected LANGUAGE=javascript onclick=""return chkMilestone_onclick(" & rs("CountryID")& "," & i & ")""><INPUT value=""" & rs("CountryID") & """ style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
			else
				Response.Write "<TR bgcolor=lightsteelblue ID=Row" & rs("CountryID") & "><TD style=""BORDER-TOP: gray thin solid; width:3%""><INPUT value=""" & rs("CountryID") & """ checked style=""width:16;height:16;"" type=""checkbox"" id=chkSelected" & trim(strRegionID) & " name=chkSelected  LANGUAGE=javascript onclick=""return chkMilestone_onclick(" & rs("CountryID") & "," & i & ")""><INPUT value=""" & rs("CountryID") & """ checked style=""width:16;height:16;display:none"" type=""checkbox"" id=chkTag name=chkTag></td>"
			end if			
            'add releases column to the country row
            if intIsPulsarProduct = 1 then
                response.write "<TD style=""BORDER-TOP: gray thin solid; width:35%""><font size=1 face=verdana>" & rs("Country") & "</font></TD>"
                intTotalLocalization = rs("TotalLocalization")
                strReleaseInfo = rs("ReleaseSelected") & ""
                aReleaseInfo = Split(strReleaseInfo,";")
                response.write "<TD style=""BORDER-TOP: gray thin solid; width:62%""><font size=2 face=verdana>" 
                for each strRelease in aReleaseInfo
                    arrayRelease = split(strRelease,"-")                  
                    intReleaseSelected = arrayRelease(2)
                    if (InStr(strReleaseIDs , "," & trim(arrayRelease(0)) & ",") = 0) then
                        strReleaseIDs = strReleaseIDs + trim(arrayRelease(0)) + ","
                        ReleaseNames = ReleaseNames + trim(arrayRelease(1)) + ","
                        PDDLockeds = PDDLockeds + trim(arrayRelease(3)) + ","
                    end if
                    strCheckBoxID = rs("CountryID") & "-" & trim(arrayRelease(0))
                    if (cInt(intReleaseSelected) > 0) then
                        if (cInt(intReleaseSelected) >= cInt(intTotalLocalization)) then
                            Response.Write "<INPUT value=""" & strCheckBoxID & """ class=""TriStateCheckboxChecked"" onclick=""return chkrelease_onclick(this,'txtnew" & strCheckBoxID  & "',Row" & rs("CountryID") & ")"" style=""width:13;height:13"" type=""image"" id=""chk" & strCheckBoxID & """ name=""chkCountriesRelease"">"                        
                            Response.Write "<INPUT value=""1"" style=""display:none;width:10px"" type=""text"" id=""txtold" & strCheckBoxID & """ name=""txtold" & strCheckBoxID & """>"
                            Response.Write "<INPUT value=""1"" style=""display:none"" type=""text"" id=""txtnew" & strCheckBoxID & """ name=""txtnew" & strCheckBoxID & """>"
                        else
                            Response.Write "<INPUT value=""" & strCheckBoxID & """ class=""TriStateCheckboxPartial"" onclick=""return chkrelease_onclick(this, 'txtnew" & strCheckBoxID & "',Row" & rs("CountryID") & ")"" style=""width:13;height:13"" type=""image"" id=""chk" & strCheckBoxID & """ name=""chkCountriesRelease"">"                        
                            Response.Write "<INPUT value=""2"" style=""display:none;width:10px"" type=""text"" id=""txtold" & strCheckBoxID & """ name=""txtold" & strCheckBoxID & """>"
                            Response.Write "<INPUT value=""2"" style=""display:none;width:10px"" type=""text"" id=""txtnew" & strCheckBoxID & """ name=""txtnew" & strCheckBoxID & """>"
                        end if
                    else
                        Response.Write "<INPUT value=""" & strCheckBoxID & """ class=""TriStateCheckbox"" onclick=""return chkrelease_onclick(this,'txtnew" & strCheckBoxID  & "', Row" & rs("CountryID") & ")"" style=""width:13;height:13"" type=""image"" id=""chk" & strCheckBoxID & """ name=""chkCountriesRelease"">"    
                        Response.Write "<INPUT value=""0"" style=""display:none;width:10px"" type=""text"" id=""txtold" & strCheckBoxID & """ name=""txtold" & strCheckBoxID & """>"
                        Response.Write "<INPUT value=""0"" style=""display:none;width:10px"" type=""text"" id=""txtnew" & strCheckBoxID & """ name=""txtnew" & strCheckBoxID & """>"
                    end if
                    Response.Write "&nbsp<label for=chk" & strCheckBoxID & ">" & trim(arrayRelease(1)) & "</label>&nbsp&nbsp"
                next
                response.write "</font></TD>"
            else
                response.write "<TD style=""BORDER-TOP: gray thin solid;""><font size=1 face=verdana>" & rs("Country") & "</font></TD>"
            end if
            'end of adding releases
            response.write "</tr>"
			i=i+1
			rs.MoveNext
		loop
		rs.Close       
		Response.Write "</table>"
	  
	end if
	set rs = nothing   
	cn.Close
	set cn = nothing
end if


        %>
        <input type="hidden" id="ProdID" name="ProdID" value="<%= request("ProdID")%>">
        <input type="hidden" id="ProdBrandID" name="ProdBrandID" value="<%= request("ProdBrandID")%>">
        <input type="hidden" id="IsPulsarProduct" name="IsPulsarProduct" value="<%= request("IsPulsarProduct")%>">
        <input type="hidden" id="hidProductReleases" name="hidProductReleases" value="<%=strReleaseIDs%>">
        <input type="hidden" id="hidProductReleasesName" name="hidProductReleasesName" value="<%=ReleaseNames%>">
        <input type="hidden" id="hidPDDLocked" name="hidPDDLocked" value="<%=PDDLockeds%>">
    </form>
</body>
</html>
<%
Function CheckPDDLockBit(ProductVersionID)
	Dim cn, cmd, p, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	Set cm = Server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "usp_GetPDDLockStatus"

	Set p = cm.CreateParameter("@p_ProductVersionID", 3, &H0001)
	p.Value = ProductVersionID
	cm.Parameters.Append p

	Set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorType = adOpenStatic
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	Set cm = nothing
	Set p = nothing

	If rs.EOF and rs.BOF Then
		CheckPDDLockBit = -500
	Else
		CheckPDDLockBit = rs("PDDLocked")
	End If
	
	rs.Close
	
	Set rs = nothing
	Set cn = nothing
End Function

Sub FillDcrList(ProductVersionID)
	Dim Result
	Dim cn, cmd, p, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	Set cm = Server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListApprovedDCRs"

	Set p = cm.CreateParameter("@ProdID", 3, &H0001)
	p.Value = ProductVersionID
	cm.Parameters.Append p

	Set rs = Server.CreateObject("ADODB.recordset")
	rs.CursorType = adOpenStatic
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 

	Set cm = nothing
	Set p = nothing

	Response.Write "<option value=""0"">-- Please Select A DCR --</option>"					
	
	Do until rs.eof
		Response.Write "<option value=""" & rs("ID") & """>" & rs("ID") & ":" & server.HTMLEncode(rs("Summary")) & "</option>"					
		rs.movenext
	Loop

	rs.close

	set cn = nothing
	set cm = nothing
	set rs = nothing
	set p  = nothing
End Sub
%>
