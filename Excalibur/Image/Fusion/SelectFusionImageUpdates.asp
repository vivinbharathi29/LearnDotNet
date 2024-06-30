<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<STYLE>
TD
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: white;
}
TH
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: gainsboro;
}

.SummaryTH
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: LightSteelBlue;
}

.SummaryTD
{
	FONT-SIZE: x-small;
	FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: gainsboro;
}

.ChkMargin
{
	margin-left: 20px;
}
</STYLE>

<head>
<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
</HEAD>
<script src="/excalibur/includes/client/jquery.min.js" type="text/javascript"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    var addTotalCount, remTotalCount, repTotalCount;
    var addItemCount = 0, remItemCount = 0, repItemCount = 0;

    function window_onload() {
        lblProcessing.style.display = "none";
        addTotalCount = document.getElementById('hidAddCount').value;
        remTotalCount = document.getElementById('hidRemCount').value;
        repTotalCount = document.getElementById('hidRepCount').value;

        $("#InsOptDefaultSetting").change(function () {
            SelectAllSetting('IO', 2);
        });
        $("#InsOptCustomSetting").change(function () {
            SelectAllSetting('IO', 3);
        });
        $("#SortDefaultSetting").change(function () {
            SelectAllSetting('SO', 2);
        });
        $("#SortCustomSetting").change(function () {
            SelectAllSetting('SO', 1);
        });

        $("#DestDefaultSetting").change(function () {
            SelectAllSetting('DC', 2);
        });
        $("#DestCustomSetting").change(function () {
            SelectAllSetting('DC', 1);
        });
        
        $("#LocDefaultSetting").change(function () {
            SelectAllSetting('IL', 2);
        });
        $("#LocCustomSetting").change(function () {
           SelectAllSetting('IL', 1);
        });

        $("input[Name^='SO']").change(function () {
            CheckChooseItem('SO', this.value);
        });

        $("input[Name^='IO']").change(function () {
            CheckChooseItem('IO', this.value);
        });

        $("input[Name^='DC']").change(function () {
            CheckChooseItem('DC', this.value);
        });
        
        $("input[Name^='IL']").change(function () {
            CheckChooseItem('IL', this.value);
        });


    }

    function SelectAllSetting(setting, option) {
        var SettingList = $("input[Name^='" + setting + "']");
        for (var i = 0; i < SettingList.length; i++) {
            if (SettingList[i].value == option)
                SettingList[i].checked = true;
            else
                SettingList[i].checked = false;
        }
    }

    function DisplayTargetIssues() {
        TargetIssuesRow.style.display = "";
        ImageIssuesRow.style.display = "none";
    }

    function DisplayImageIssues() {
        TargetIssuesRow.style.display = "none";
        ImageIssuesRow.style.display = "";
    }


    function ProductDrop_onclick(DropID) {
        var list = document.getElementsByTagName("input");
        var itemValueList;
        for (var i = 0; i < list.length; i++)
            if (list[i].className == "PD" + DropID) {
                list[i].checked = event.srcElement.checked;
                itemValueList = list[i].value.split("|");
                if (itemValueList[4] == "add" && list[i].checked)
                    addItemCount = addItemCount + 1;
                else if (itemValueList[4] == "add" && !list[i].checked)
                    addItemCount = addItemCount - 1;
                else if (itemValueList[4] == "rem" && list[i].checked)
                    remItemCount = remItemCount + 1;
                else if (itemValueList[4] == "rem" && !list[i].checked)
                    remItemCount = remItemCount - 1;
                else if (itemValueList[4] == "rep" && list[i].checked)
                    repItemCount = repItemCount + 1;
                else if (itemValueList[4] == "rep" && !list[i].checked)
                    repItemCount = repItemCount - 1;
            }

        if (!event.srcElement.checked) {
            document.getElementById('CheckAll').checked = event.srcElement.checked;
            document.getElementById('CheckAllAdd').checked = event.srcElement.checked;
            document.getElementById('CheckAllRemove').checked = event.srcElement.checked;
            document.getElementById('CheckAllReplace').checked = event.srcElement.checked;
        }

    }

    function CheckAll_onclick() {
        var list = document.getElementsByTagName("input");
        for (var i = 0; i < list.length; i++)
            if (list[i].type == "checkbox")
                list[i].checked = event.srcElement.checked;

        if (!event.srcElement.checked) {
            addItemCount = 0;
            remItemCount = 0;
            repItemCount = 0;
        }

    }

    function CheckSameFunctionItem(actionName, chked) {
        var chkValue;
        var chkList = document.getElementsByName('ChangeDetails');
        for (var i = 0; i < chkList.length; i++) {
            chkValue = chkList[i].value.split("|");
            if (chkValue[4] == actionName)
                chkList[i].checked = chked;
        }

        if (!chked) {
            document.getElementById('CheckAll').checked = chked;
            var ProdList = document.getElementsByTagName("input");
            for (var i = 0; i < ProdList.length; i++) {
                if (ProdList[i].className == "prodDrop")
                    ProdList[i].checked = chked;
            }
        }

        if (actionName == "add" && chked)
            addItemCount = addTotalCount;
        else if (actionName == "add" && !chked)
            addItemCount = 0;
        else if (actionName == "rem" && chked)
            remItemCount = remTotalCount;
        else if (actionName == "rem" && !chked)
            remItemCount = 0;
        else if (actionName == "rep" && chked)
            repItemCount = repTotalCount;
        else if (actionName == "rep" && !chked)
            repItemCount = 0;
    }

    function AddItem(chked) {
        if (chked)
            addItemCount = addItemCount + 1;
        else
            addItemCount = addItemCount - 1;

        if (addItemCount != addTotalCount)
            document.getElementById('CheckAllAdd').checked = false;
        else
            document.getElementById('CheckAllAdd').checked = true;
    }

    function RemoveItem(chked) {
        if (chked)
            remItemCount = remItemCount + 1;
        else
            remItemCount = remItemCount - 1;

        if (remItemCount != remTotalCount)
            document.getElementById('CheckAllRemove').checked = false;
        else
            document.getElementById('CheckAllRemove').checked = true;
    }

    function ReplaceItem(chked) {
        if (chked)
            repItemCount = repItemCount + 1;
        else
            repItemCount = repItemCount - 1;

        if (repItemCount != repTotalCount)
            document.getElementById('CheckAllReplace').checked = false;
        else
            document.getElementById('CheckAllReplace').checked = true;
    }


    function CheckChooseItem(type, option) {
        if (type == 'IO' & $("#InsOptDefaultSetting").is(':checked') & option != 2)
            $("#InsOptDefaultSetting").prop('checked', false);
        else if (type == 'IO' & !$("#InsOptDefaultSetting").is(':checked') & option != 3)
            $("#InsOptCustomSetting").prop('checked', false);
        else if (type == 'SO' & $("#SortDefaultSetting").is(':checked') & option != 2)
            $("#SortDefaultSetting").prop('checked', false);
        else if (type == 'SO' & !$("#SortDefaultSetting").is(':checked') & option != 1)
            $("#SortCustomSetting").prop('checked', false);
        else if (type == 'DC' & $("#DestDefaultSetting").is(':checked') & option != 2)
            $("#DestDefaultSetting").prop('checked', false);
       else if (type == 'DC' & !$("#DestDefaultSetting").is(':checked') & option != 1)
           $("#DestCustomSetting").prop('checked', false);
       else if (type == 'IL' & $("#LocDefaultSetting").is(':checked') & option != 2)
           $("#LocDefaultSetting").prop('checked', false);
       else if (type == 'IL' & !$("#LocDefaultSetting").is(':checked') & option != 1)
           $("#LocCustomSetting").prop('checked', false);

    }
    //-->

</SCRIPT>
</head>
<body onload="return window_onload()">
<b><font id="lblProcessing" face="verdana" size="2">Processing. This may take several minutes.  Please wait...</font></b>
<%response.Flush %>
<font size="3" face="verdana"><b>Sync Images with IRS</b></font><br />
<form id="frmMain" action="SelectFusionImageUpdatesSave.asp" method="post" target="_blank" >
<input id="Submit1" type="submit" value="Send Selected Changes to IRS" /><br /><br />
<%
	
	Server.ScriptTimeout = 240

	dim StartDate
	StartDate = now()

	dim cn
	dim cn2
	dim cm
	dim p
	dim rs
	dim rs2
	dim strDash
	dim strSKU
	dim strOutBuffer
	dim strDelConveyor
	dim strSKURevision
	dim strSQL
	dim skuCount
	dim errorcount
	dim totalerrorcount
	dim TotalCompared
	dim strConveyor
	dim TableCount
	dim CurrentUserPartner
	dim strHideTables
	dim strPreinstallTeam
	dim strServerLocation
	dim SkuHeader
	dim CheckingSKUNumber
	dim strImageID
	dim TableHeaderDisplayed
	dim strCompareType
	dim CompareTypeUpdated
	dim blnUseSnapshotDeliverables
	dim ExcaliburSKUArray
	dim currentuserid
	dim ExtraConveyorHeaderDisplayed
	dim strProductDrops
	dim ProductDropArray
    dim LinesDisplayed
    dim LinesDisplayedTotal
    dim addCount
	dim remCount
	dim repCount
    dim installOptionList
    dim sortOrderList
    dim regionsList
    dim localesList
	dim lastPartNo
    dim productDropName
    dim productVersionID
    dim lastIRSProductDrop
    dim productName

	TableCount =0
	strProductDrops = ""
    LinesDisplayedTotal=0


	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")


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
	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../../NoAccess.asp?Level=0"
	else
		CurrentUserPartner = rs("PartnerID")
		currentuserid = rs("ID")
	end if 
	rs.Close		



	if trim(request("ProductDrop")) = "" and trim(request("ProductID")) = "" and trim(request("ImageDefinitionID")) = "" then
		Response.Write "<BR><font size=2 face=verdana>Unable to run report because all of the required information was not supplied.</font><BR><BR>"
	else

		if trim(request("ImageDefinitionID")) <> "" then 'Image Definition Supplied
			rs.open "spListProductDrops " & clng(request("ImageDefinitionID")),cn
			do while not rs.eof
                if instr(rs("ProductDrop") & "","##") > 0 and (rs("StatusID") = 1) then 
				    strProductDrops = strProductDrops & "," & rs("ProductDrop")
	            end if			
                rs.movenext 
			loop
			rs.close
		elseif trim(request("ProductDrop")) <> "" then 'Product Drop Supplied
            if instr(rs("ProductDrop") & "","##") > 0 then
			    strProductDrops = strProductDrops & "," & request("ProductDrop")
            end if
		else 'productID supplied
			rs.open "spListProductDrops4product " & clng(request("ProductID")),cn
			do while not rs.eof
                if instr(rs("ProductDrop") & "","##") > 0 and (rs("StatusID") = 1) then
                    strProductDrops = strProductDrops & "," & rs("ProductDrop") &"|" & rs("ProductVersionID") & "|" & rs("DotsName")
	            end if			
                rs.movenext 
			loop
			rs.close

		end if



	end if

	if strProductDrops <> "" then
		strProductDrops = mid(strProductDrops,2)
	end if

	ProductDropArray = split(strProductDrops,",")
	response.write "<table cellpadding=2 cellspacing=0 border=1>"
    response.write "<tr><td style=""background-color: lightsteelblue"" colspan=12><input id=""CheckAll"" type=""checkbox"" onclick=""javascript: CheckAll_onclick()""> All Product Drops" &_
    "<input id=""CheckAllAdd"" class=""ChkMargin"" type=""checkbox"" onclick=""javascript: CheckSameFunctionItem('add', checked)"" > Add " &_
	"<input id=""CheckAllRemove"" class=""ChkMargin"" type=""checkbox"" onclick=""javascript: CheckSameFunctionItem('rem', checked)""> Remove " &_
	"<input id=""CheckAllReplace"" class=""ChkMargin"" type=""checkbox"" onclick=""javascript: CheckSameFunctionItem('rep', checked)""> Replace " &_
	"</td></tr>" &_
    "<tr><td colspan=4>&nbsp;</td><td><input type='radio' ID='InsOptDefaultSetting' Name='InsOptSetting' checked>Default Install Option</input></td><td><input type='radio' ID='InsOptCustomSetting' Name='InsOptSetting'>Current Install Option</input></td><td><input type='radio' ID='SortDefaultSetting' Name='SortSetting' checked>Default Sort Order</input></td><td><input type='radio' ID='SortCustomSetting' Name='SortSetting'>Current Sort Order</input></td>"&_
    "<td><input type='radio' ID='DestDefaultSetting' Name='DestSetting' checked>Default Destination Country/Region</input></td><td><input type='radio' ID='DestCustomSetting' Name='DestSetting'>Current Destination Country/Region</input></td>"&_
    "<td><input type='radio' ID='LocDefaultSetting' Name='LocSetting' checked>Default Image Locale</input></td><td><input type='radio' ID='LocCustomSetting' Name='LocSetting'>Current Image Locale</input></td></tr>"
	for i = 0 to ubound(productdroparray)
		if trim(Productdroparray(i)) = "" then
			response.write "SKIPPED A PRODUCT DROPS WITH NO NUMBER ASSIGNED<br><br>"
		else
			ProductDropCount = ProductDropCount + 1
            tmparray = split(productdroparray(i),"|")
            productDropName = tmparray(0)
            productVersionID = tmparray(1)
            productName = tmparray(2)

			if (lastIRSProductDrop <> productDropName) then
                rs2.open "spGetProductDropID '" & FixProductDrop(trim(productDropName)) & "'",cn
			    if rs2.eof and rs2.bof then
				    strProductDropID = ""
			    else
				    strProductDropID = rs2("ProductDropID") & ""
			    end if
			    rs2.close

			    response.write "<tr><td colspan=12 style=""background-color: gainsboro"">"

			    if strProductDropID <> "" then
				    response.write  "<input id=""PD" & strProductDropID & """ type=""checkbox"" onclick=""javascript: ProductDrop_onclick(" & strProductDropID & ");"" class=""prodDrop""/> " & productDropName  
				    response.write "</td></tr>"
			    else
				    response.write  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & productDropName  
				    response.write "</td></tr><tr><td colspan=12>Product Drop Not Found in IRS</td></tr>"
			    end if
			    rs2.Open "spListImagesForProductDrop '" & productDropName & "'",cn,adOpenForwardOnly
			    strImageList = ""
			    do while not rs2.eof
				    strImageList = strImageList & "," & rs2("ImageID")
				    rs2.movenext
			    loop
			    rs2.close
			    if strImageArray <> "" then
				    strImageArray = mid(strImageArray,2)
			    end if
			    ImageArray = split(strImageList,",")

		       'response.write strImageList & "<br>"

			    strDelIRS = ""
			    strSQl = "spListComponentsInIRSImage '" & FixProductDrop(productDropName) & "',0"

			    rs2.Open strSQl,cn,adOpenForwardOnly
			    do while not rs2.EOF
				    strDelIRS = strDelIRS & "2|"
				    strDelIRS = strDelIRS  & trim(rs2("RootID")) & "|"
				    strDelIRS = strDelIRS  & trim(rs2("PartNo")) & "|"
				    strDelIRS = strDelIRS & trim(rs2("Name")) & " " &  trim(rs2("Version"))
				    if  rs2("Revision") & "" <> "" then
					    strDelIRS = strDelIRS & " " & trim(rs2("Revision"))
				    end if
				    if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
					    strDelIRS = strDelIRS & asc(lcase(trim(rs2("Pass")))) -87
				    else
					    strDelIRS = strDelIRS & trim(rs2("Pass"))
				    end if
                    if Not IsNull(rs2("DMIStr")) and rs2("DMIStr") <> "" then
					    installOptionList = installOptionList & "2|" & FixProductDrop(productDropName) & "|" & trim(rs2("PartNo")) & "|" & Left(rs2("DMIStr"), len(rs2("DMIStr"))-1) & "|" '& vbcrlf
				    end if
                    if Not IsNull(rs2("SortOrder")) then
					    sortOrderList = sortOrderList & "2|" & FixProductDrop(productDropName) & "|" & trim(rs2("PartNo")) & "|"  & rs2("SortOrder") & "|" '& vbcrlf
				    end if
                    if Not IsNull(rs2("SupportedRegions")) then
					    regionsList = regionsList & "2|" & FixProductDrop(productDropName) & "|" & trim(rs2("PartNo")) & "|"  & rs2("SupportedRegions") & "|" '& vbcrlf
				    end if
                    if Not IsNull(rs2("ImageLocale")) then
					    localesList = localesList & "2|" & FixProductDrop(productDropName) & "|" & trim(rs2("PartNo")) & "|"  & rs2("ImageLocale") & "|" '& vbcrlf
				    end if
				    strDelIRS = strDelIRS & vbcrlf
				    rs2.MoveNext
			    loop
			    rs2.Close
                lastIRSProductDrop = productDropName
            end if
			'response.write replace(strDelIRS,vbcrlf,"<br>") & "<br><br>"   
			'if request("CompareType") = "" then
			'    strSQl = "spListDeliverablesInProductDrop '" & productdroparray(i) & "'," & clng(request("ProductID"))
			'else
				strSQl = "spListDeliverablesInProductDrop '" & productDropName & "'," &  clng(ProductVersionID) & ",1,1" 
			'end if
			rs2.Open strSQl,cn,adOpenForwardOnly
			strOutbuffer = ""	
			do while not rs2.EOF
				if ( rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore") ) then 
					blnFound = false
					DeliverableImageArray = split(rs2("images") & "",",")
					for j = 0 to ubound(ImageArray)
						if inarray(DeliverableImageArray,imagearray(j)) then
							blnfound = true
							exit for
						end if
						if j > 3000 then
							response.write "<br>Infinate loop detected.  Exiting..."
							exit for
						end if
					next
					if trim(rs2("Images") & "") = "" or blnFound then
						strOutbuffer = strOutbuffer & "1|" 
						strOutbuffer = strOutbuffer & trim(rs2("RootID")) & "|" 
						strOutbuffer = strOutbuffer & trim(rs2("IRSPartNumber")) & "|" 
						strOutbuffer = strOutbuffer & trim(rs2("Name")) & " " &  rs2("Version")
						if  rs2("Revision") & "" <> "" then
							strOutbuffer = strOutbuffer & " " & rs2("Revision")
						end if
						if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
							strOutbuffer = strOutbuffer & asc(lcase(rs2("Pass"))) -87
						else
							strOutbuffer = strOutbuffer & rs2("Pass")
						end if
                        if Not IsNull(rs2("DMIStr")) and rs2("DMIStr") <> "" then
							installOptionList = installOptionList & "1|" & FixProductDrop(productDropName) & "|" & trim(rs2("IRSPartNumber")) & "|" & Left(rs2("DMIStr"), len(rs2("DMIStr"))-1) & "|"  '& vbcrlf
						end if
                        if Not IsNull(rs2("SortOrder")) then
							sortOrderList = sortOrderList & "1|" & FixProductDrop(productDropName) & "|" & trim(rs2("IRSPartNumber")) & "|" & rs2("SortOrder") & "|"  '& vbcrlf
						end if
                        if Not IsNull(rs2("Regions")) then
					        regionsList = regionsList & "1|" & FixProductDrop(productDropName) & "|" & trim(rs2("IRSPartNumber")) & "|"  & rs2("Regions") & "|" '& vbcrlf
				        end if
                        if Not IsNull(rs2("Locales")) then
					        localesList = localesList & "1|" & FixProductDrop(productDropName) & "|" & trim(rs2("IRSPartNumber")) & "|"  & rs2("Locales") & "|" '& vbcrlf
				        end if 
						strOutbuffer = strOutbuffer & vbcrlf
					end if
				end if
				rs2.MoveNext
			loop
			rs2.close		

			'Sort IRS Deliverables
			LineArray = Split(lcase(strDelIRS), vbcrlf)
            if (lastIRSProductDrop <> productDropName) then
			    strDelIRS = ""
                If (UBound(LineArray) > 1) Then
			        For j = UBound(LineArray) - 1 To 0 Step -1
				        For k = 0 To j
					        If mid(LineArray(k),2) > mid(LineArray(k + 1),2) Then
						        temp = LineArray(k + 1)
						        LineArray(k + 1) = LineArray(k)
						        LineArray(k) = temp
					        End If
				        Next
			        Next
                End If
            end if
						
			'Remove Dups and append to excalibur list
			for j = lbound(LineArray) to ubound(LineArray)					
				if j = ubound(LineArray) then
					strOutbuffer = strOutbuffer & LineArray(j) & vbcrlf
				elseif linearray(j) <> linearray(j+1) then
					strOutbuffer = strOutbuffer & LineArray(j) & vbcrlf
				end if
			next
			'Sort
			LineArray = Split(lcase(strOutBuffer), vbcrlf)
			if  UBound(LineArray) > 0 then
				TotalCompared = TotalCompared + UBound(LineArray)
			end if
            If (UBound(LineArray) > 1) Then
			    For j = UBound(LineArray) - 1 To 0 Step -1
				    For k = 0 To j
					    If mid(LineArray(k),2) > mid(LineArray(k + 1),2) Then
						    temp = LineArray(k + 1)
						    LineArray(k + 1) = LineArray(k)
						    LineArray(k) = temp
					    End If
				    Next
			    Next
            End If

			j=0
			TableHeaderDisplayed=false
			ErrorCount=0

			LastRoot = 0
			ExcaliburCount = 0
			IRSCount = 0
			ExcaliburPNList = ""
			IRSPNList = ""            
			RowOutputCount = 0
            LinesDisplayed = 0
            ExcaliburIO=""
			IRSIO=""
            ExcaliburSortOrder=""
			IRSSortOrder="" 
            ExcaliburRegion=""
            IRSRegion=""
            ExcaliburLocale=""
            IRSLocale=""
            lastPartNo=""

			for j = 0 to ubound(LineArray)
				if trim(LineArray(j)) <> "" then
					TempArray = split (LineArray(j),"|")
					if (clng(temparray(1)) <> lastroot) or ((len(lastPartNo) = 10) and (clng(temparray(1)) = lastroot) and (InStr(7,"b2",lastPartNo) = 0) and (Mid(temparray(2),8,2) <> Mid(lastPartNo, 8, 2))) then 
						OutputLine
						RowOutputCount  = RowOutputCount + 1
					end if
					'response.write TempArray(0) & "|" & TempArray(1) & "|" & TempArray(2) & "|" & TempArray(3) & "<br>" 
					if temparray(0) = "1" then 'Excalibur
						if instr(IRSPNList,"|" & temparray(2) & "|") > 0 then
							IRSPNList = replace(IRSPNList,"|" & temparray(2) & "|","")
							IRSCount = IRSCount - 1
						else
							ExcaliburPNList= ExcaliburPNList & "|" & temparray(2) & "|"
							ExcaliburCount = ExcaliburCount +1
						end if
					elseif temparray(0) = "2" then 'IRS
						if instr(ExcaliburPNList,"|" & temparray(2) & "|") > 0 then
							ExcaliburPNList = replace(ExcaliburPNList,"|" & temparray(2) & "|","")
							ExcaliburCount = ExcaliburCount -1
						else
							IRSPNList= IRSPNList & "|" & temparray(2) & "|"
							IRSCount = IRSCount + 1
						end if
					else
						response.write "Error<br>"
					end if


					LastRoot = clng(TempArray(1))
                    lastPartNo = temparray(2)
				end if
			next


			OutputLine
			RowOutputCount  = RowOutputCount + 1


		end if   
			if RowOutputCount = 1 and trim(strProductDropID & "") <> "" then
				response.write "<tr><td colspan=12>No changes needed. (" & productName & ") </td></tr>"
			end if
            if LinesDisplayed = 0 then
    		    response.write "<tr><td colspan=12>No changes needed. (" & productName & ") </td></tr>"
            end if
            LinesDisplayedTotal = LinesDisplayedTotal + LinesDisplayed
            LinesDisplayed = 0
	next





	set rs=nothing
	cn.close
	set cn = nothing


    if RowOutputCount = 0 then
        response.write "<tr><td colspan=8>No changes needed.</td></tr>"
    elseif LinesDisplayedTotal > 50 then
    	response.write "</table><br /><input id=""Submit2"" type=""submit"" value=""Send Selected Changes to IRS"" />"
    end if







function FixProductDrop(strProductDrop)
	if right(strProductDrop,2) = "##" then
		FixProductDrop = strProductDrop
	else
		FixProductDrop = left(strProductDrop,len(strProductDrop)-2) & "##"
	end if
end function





function InArray(MyArray,strFind)
	dim strElement
	dim blnFound
			
	blnFound = false
	for each strElement in MyArray
		if trim(strElement) = trim(strFind) and trim(strElement) <> "" then
			blnFound = true
			exit for
		end if
	next
	InArray = blnFound
end function

	function ScrubSQL(strWords) 
		dim badChars 
		dim newChars 
		dim i
		strWords=replace(strWords,"'","''")
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 

		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

    
     function getInstallOption(searchStr)
		if (Not IsNull(installOptionList) and Len(installOptionList) > 0 ) then
			strStartIndex = InStr(installOptionList,searchStr) + Len(searchStr)
			strEndIndex = InStr(strStartIndex,installOptionList,"|") 
			getInstallOption = mid(installOptionList,strStartIndex,(strEndIndex-strStartIndex))
        else
            getInstallOption=""
		end if

	end function

    function getSortOrder(searchStr)
		if (Not IsNull(SortOrderList) and Len(SortOrderList) > 0 ) then
			strStartIndex = InStr(SortOrderList,searchStr) + Len(searchStr)
			strEndIndex = InStr(strStartIndex,SortOrderList,"|") 
			getSortOrder = mid(SortOrderList,strStartIndex,(strEndIndex-strStartIndex))
        else
            getSortOrder=""
		end if
	end function

    function getRegion(searchStr)
		if (Not IsNull(regionsList) and Len(regionsList) > 0 ) then
			strStartIndex = InStr(regionsList,searchStr) + Len(searchStr)
			strEndIndex = InStr(strStartIndex,regionsList,"|") 
			getRegion = mid(regionsList,strStartIndex,(strEndIndex-strStartIndex))
	 
        else
            getRegion=""
		end if
    end function

    function getLocale(searchStr)
		if (Not IsNull(localesList) and Len(localesList) > 0 ) then
			strStartIndex = InStr(localesList,searchStr) + Len(searchStr)
			strEndIndex = InStr(strStartIndex,localesList,"|") 
			getLocale = mid(localesList,strStartIndex,(strEndIndex-strStartIndex))
        else
            getLocale=""
		end if
    end function

	function OutputLine
						ExcaliburPNList = replace(ExcaliburPNList,"||","|")
						if left(ExcaliburPNList,1) = "|" then
							ExcaliburPNList = mid(ExcaliburPNList,2)
						end if
						if right(ExcaliburPNList,1) = "|" then
							ExcaliburPNList = left(ExcaliburPNList,len(ExcaliburPNList)-1)
						end if

						IRSPNList = replace(IRSPNList,"||","|")
						if left(IRSPNList,1) = "|" then
							IRSPNList = mid(IRSPNList,2)
						end if
						if right(IRSPNList,1) = "|" then
							IRSPNList = left(IRSPNList,len(IRSPNList)-1)
						end if

						'Output results
						if ExcaliburCount = 0 and IRSCount > 0 then 'Remove
							PNArray = split(IRSPNList,"|")
							for each strPartNumber in PNArray
								response.write "<tr>"
								rs2.open "spGetProductAndDeliverableInfoForIRSFeed " & clng(ProductVersionID) & ",'" & strPartNumber & "'",cn
								if not (rs2.eof and rs2.bof) then
									if trim(strProductDropID) <> "" and (Not IsNull(rs2("IRSComponentPassID"))) then
										response.write "<td><input name=""ChangeDetails"" class=""PD" & strProductDropID & """ type=""checkbox"" value=""" & rs2("ProductName") & "|" & FixProductDrop(productDropName) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rem|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")&"") & "|" & cstr(rs2("NotInRCD")&"") & "|" & cstr(rs2("InAppRecovery")&"") & "|" & cstr(rs2("DIB")&"") & "|" & """ onclick=""javascript: RemoveItem(checked)"" />Remove</td>"
									    remCount = remCount + 1
                                    else
										response.write "<td>Remove&nbsp;</td>"
									end if
									'response.write "<td>" & rs2("ProductName") & "|" & FixProductDrop(Productdroparray(i)) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rem|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")&"") & "|" & cstr(rs2("NotInRCD")&"") & "|" & cstr(rs2("InAppRecovery")&"") & "|" & cstr(rs2("DIB")&"") & "|</td>" 
									strDeliverableName = rs2("DeliverableName") & " [" & rs2("Version") 
									if trim(rs2("revision") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("revision")
									end if
									if trim(rs2("pass") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("pass")
									end if
									strDeliverableName = strDeliverableName & "]"
									response.write  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(strPartNumber) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
	                                LinesDisplayed = LinesDisplayed + 1							
                                else
									response.write "<td colspan=12>Record not found for: " & strPartNumber & "</td>"
								end if

								rs2.close
								response.write "</tr>"
							next

						elseif ExcaliburCount > 0 and IRSCount = 0 then 'Add
							PNArray = split(ExcaliburPNList,"|")
							for each strPartNumber in PNArray
								response.write "<tr>"
								rs2.open "spGetProductAndDeliverableInfoForIRSFeed " & clng(ProductVersionID) & ",'" & strPartNumber & "'",cn
								if not (rs2.eof and rs2.bof) then
									if trim(strProductDropID) <> "" and (Not IsNull(rs2("IRSComponentPassID"))) then
										response.write "<td><input name=""ChangeDetails"" class=""PD" & strProductDropID & """ type=""checkbox""  value=""" & rs2("ProductName") & "|" & FixProductDrop(productDropName) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|add|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|"" onclick=""javascript: AddItem(checked)"" />Add</td>"
                                        addCount = addCount + 1
									else
										response.write "<td>Add</td>"
									end if
'                                    response.write "<td>" & rs2("ProductName") & "|" & FixProductDrop(Productdroparray(i)) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|add|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|</td>" 
									strDeliverableName = rs2("DeliverableName") & " [" & rs2("Version") 
									if trim(rs2("revision") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("revision")
									end if
									if trim(rs2("pass") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("pass")
									end if
									strDeliverableName = strDeliverableName & "]"
									response.write  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(strPartNumber) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
	                                LinesDisplayed = LinesDisplayed + 1							
								else
									response.write "<td>Add</td>"
									response.write "<td colspan=11>Unable to find info for " & strPartNumber & "</td>"
								end if
								rs2.close
								response.write "</tr>"
							next
						elseif ExcaliburCount =1 and IRSCount =1 then 'Replace
							response.write "<tr>"
							IRSPNList = replace(IRSPNList,"|","")
							ExcaliburPNList = replace(ExcaliburPNList,"|","")

							rs2.open "spGetProductAndDeliverableInfoForIRSFeed " & clng(ProductVersionID) & ",'" & ExcaliburPNList & "'",cn
							if not (rs2.eof and rs2.bof) then
								if trim(strProductDropID) <> "" and (Not IsNull(rs2("IRSComponentPassID"))) then
									response.write "<td><input  name=""ChangeDetails"" class=""PD" & strProductDropID & """ type=""checkbox""  value=""" & rs2("ProductName") & "|" & FixProductDrop(productDropName) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rep|" & ExcaliburPNList & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|" & IRSPNList & """ onclick=""javascript: ReplaceItem(checked)""  />Replace</td>"
								    repCount = repCount + 1
                                else
									response.write "<td>Replace</td>"
								end if
'                                response.write "<td>" & rs2("ProductName") & "|" & FixProductDrop(Productdroparray(i)) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rep|" & ExcaliburPNList & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|" & IRSPNList & "</td>"
                                    writeStr=""
									strDeliverableName = rs2("DeliverableName") & " [" & rs2("Version") 
									if trim(rs2("revision") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("revision")
									end if
									if trim(rs2("pass") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("pass")
									end if
									strDeliverableName = strDeliverableName & "]"

                                    tmpStr = "1|"&FixProductDrop(productDropName)&"|"& ucase(ExcaliburPNList) & "|"
									ExcaliburIO = getInstallOption(tmpStr)
                                    ExcaliburSortOrder = getSortOrder(tmpStr)
                                    ExcaliburRegion = getRegion(tmpStr)
                                    ExcaliburLocale = getLocale(tmpStr)
									tmpStr = "2|"&FixProductDrop(productDropName)&"|"& ucase(IRSPNList) & "|"
									IRSIO = getInstallOption(tmpStr)
                                    IRSSortOrder = getSortOrder(tmpStr)
                                    IRSRegion = getRegion(tmpStr)
                                    IRSLocale = getLocale(tmpStr)

                                    if ((ExcaliburIO <> IRSIO) and Not(ExcaliburIO="" and IRSIO="") and Not(IRSIO="override")) and ((ExcaliburSortOrder <> IRSSortOrder)) then
                                        writeStr =  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(ExcaliburPNList) & " replaces " &  ucase(IRSPNList) & "</td><td><input type=""radio"" name=""IO"& productDropName & "|" & ExcaliburPNList & """ id=""" & ExcaliburIO & """ value=""2"" checked />" & ucase(ExcaliburPNList) & "(" & ExcaliburIO & ")</td><td><input type=""radio"" id=""" & IRSIO & """ name=""IO"& productDropName & "|" & ExcaliburPNList &  """ value=""3"" />" & ucase(IRSPNList) & "(" & IRSIO &")</td><td><input type=""radio"" name=""SO"& productDropName & "|" & ExcaliburPNList & """ id="""& ExcaliburPNList& "|" & ExcaliburSortOrder & """ value=""2"" checked /> Default Sort Order Setting for "& ucase(ExcaliburPNList) & ": " & ExcaliburSortOrder & "</td><td><input type=""radio"" id=""" & ExcaliburPNList & "|" & IRSSortOrder  & """ name=""SO"& productDropName & "|" & ExcaliburPNList &  """ value=""1"" />Current Sort Order setting for " & ucase(IRSPNList) & " in " & productDropName & ": " & IRSSortOrder &"</td>"
                                    elseif ((ExcaliburIO <> IRSIO) and Not(ExcaliburIO="" and IRSIO="") and Not(IRSIO="override")) then
										writeStr =  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(ExcaliburPNList) & " replaces " &  ucase(IRSPNList) & "</td><td><input type=""radio"" name=""IO"& productDropName & "|" & ExcaliburPNList & """ id=""" & ExcaliburIO & """ value=""2"" checked />" & ucase(ExcaliburPNList) & "(" & ExcaliburIO & ")</td><td><input type=""radio"" id=""" & IRSIO & """ name=""IO"& productDropName & "|" & ExcaliburPNList &  """ value=""3"" />" & ucase(IRSPNList) & "(" & IRSIO &")</td><td>&nbsp;</td><td>&nbsp;</td>"
									elseif (ExcaliburSortOrder <> IRSSortOrder) then
										writeStr = "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(ExcaliburPNList) & " replaces " &  ucase(IRSPNList) & "</td><td>&nbsp;</td><td>&nbsp;</td><td><input type=""radio"" name=""SO"& productDropName & "|" & ExcaliburPNList & """ id="""& ExcaliburPNList& "|" & ExcaliburSortOrder & """ value=""2"" checked /> Default Sort Order Setting for "& ucase(ExcaliburPNList) & ": " & ExcaliburSortOrder & "</td><td><input type=""radio"" id=""" & ExcaliburPNList & "|" & IRSSortOrder  & """ name=""SO"& productDropName & "|" & ExcaliburPNList &  """ value=""1"" />Current Sort Order setting for " & ucase(IRSPNList) & " in " & productDropName & ": " & IRSSortOrder &"</td>"
									else
                                        writeStr =  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(ExcaliburPNList) & " replaces " &  ucase(IRSPNList) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
                                    end if

                                    if (ExcaliburRegion<>IRSRegion) and Not(ExcaliburRegion="") and Not(IRSRegion="override,") then
                                        writeStr = writeStr + "<td><input type=""radio"" name=""DC"& productDropName & "|" & ExcaliburPNList & """ id=""" & ExcaliburRegion & """ value=""2"" checked />" &  ExcaliburRegion & "</td><td><input type=""radio"" id=""" & IRSRegion & """ name=""DC"& productDropName & "|" & ExcaliburPNList &  """ value=""1"" />" & IRSRegion & "</td>"
                                    else
                                         writeStr = writeStr+ "<td>&nbsp;</td><td>&nbsp;</td>"
                                    end if
                                    
                                    if (ExcaliburLocale<>IRSLocale) and Not(ExcaliburLocale="") and Not(IRSLocale="override,") then
                                        writeStr = writeStr + "<td><input type=""radio"" name=""IL"& productDropName & "|" & ExcaliburPNList & """ id=""" & ExcaliburLocale & """ value=""2"" checked />" &  ExcaliburLocale & "</td><td><input type=""radio"" id=""" & IRSLocale & """ name=""IL"& productDropName & "|" & ExcaliburPNList &  """ value=""1"" />" & IRSLocale & "</td>"
                                    else
                                        writeStr = writeStr+ "<td>&nbsp;</td><td>&nbsp;</td>"
                                    end if
                             
				                    response.write writeStr
	                                LinesDisplayed = LinesDisplayed + 1							
							 else
								response.write "<td>Replace&nbsp;</td>"
								response.write "<td colspan=11>Unable to find info for " & strPartNumber & "</td>"
							end if
							rs2.close
							response.write "</tr>"

						elseif ExcaliburCount > 0 and IRSCount > 0 then 'Add All and Remove All

							PNArray = split(IRSPNList,"|")
							for each strPartNumber in PNArray
								response.write "<tr>"
								rs2.open "spGetProductAndDeliverableInfoForIRSFeed " & clng(ProductVersionID) & ",'" & strPartNumber & "'",cn
								if not (rs2.eof and rs2.bof) then
									if trim(strProductDropID) <> "" and (Not IsNull(rs2("IRSComponentPassID"))) then
										response.write "<td><input name=""ChangeDetails"" class=""PD" & strProductDropID & """ type=""checkbox""  value=""" & rs2("ProductName") & "|" & FixProductDrop(productDropName) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rem|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")&"") & "|" & cstr(rs2("NotInRCD")&"") & "|" & cstr(rs2("InAppRecovery")&"") & "|" & cstr(rs2("DIB")&"") & "|"" />Remove</td>"
									else
										response.write "<td>Remove&nbsp;</td>"
									end if
'                                    response.write "<td>" & rs2("ProductName") & "|" & FixProductDrop(Productdroparray(i)) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|rem|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")&"") & "|" & cstr(rs2("NotInRCD")&"") & "|" & cstr(rs2("InAppRecovery")&"") & "|" & cstr(rs2("DIB")&"") & "|</td>" 
									strDeliverableName = rs2("DeliverableName") & " [" & rs2("Version") 
									if trim(rs2("revision") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("revision")
									end if
									if trim(rs2("pass") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("pass")
									end if
									strDeliverableName = strDeliverableName & "]"
									response.write  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(strPartNumber) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
								else
									response.write "<td colspan=11>Record not found for: " & strPartNumber & "</td>"
								end if

								rs2.close
								response.write "</tr>"
							next

							PNArray = split(ExcaliburPNList,"|")
							for each strPartNumber in PNArray
								response.write "<tr>"
								rs2.open "spGetProductAndDeliverableInfoForIRSFeed " & clng(ProductVersionID) & ",'" & strPartNumber & "'",cn
								if not (rs2.eof and rs2.bof) then
									if trim(strProductDropID) <> "" and (Not IsNull(rs2("IRSComponentPassID"))) then
										response.write "<td><input name=""ChangeDetails"" class=""PD" & strProductDropID & """ type=""checkbox""  value=""" & rs2("ProductName") & "|" & FixProductDrop(productDropName) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|add|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|"" />Add</td>"
									else
										response.write "<td>Add</td>"
									end if
								   ' response.write "<td>" & rs2("ProductName") & "|" & FixProductDrop(Productdroparray(i)) & "|" & strProductDropID & "|" & replace(rs2("DeliverableName"),",","^") & "|add|" & strPartNumber & "|" & cstr(rs2("IRSComponentPassID")) & "|" & cstr(rs2("SRP")) & "|" & cstr(rs2("NotInRCD")) & "|" & cstr(rs2("InAppRecovery")) & "|" & cstr(rs2("DIB")) & "|</td>" 
									strDeliverableName = rs2("DeliverableName") & " [" & rs2("Version") 
									if trim(rs2("revision") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("revision")
									end if
									if trim(rs2("pass") & "") <> "" then
										strDeliverableName = strDeliverableName & "," & rs2("pass")
									end if
									strDeliverableName = strDeliverableName & "]"
									response.write  "<td>" & rs2("ProductName") & "</td><td>" & strDeliverableName & "</td><td>" & ucase(strPartNumber) & "</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"

								else
									response.write "<td colspan=11>Record not found for: " & strPartNumber & "</td>"
								end if
								rs2.close
								response.write "</tr>"
							next


						end if
										  
						'reset                   
						ExcaliburCount = 0
						IRSCount = 0
						ExcaliburPNList = ""
						IRSPNList = ""            

	end function

	
		set rs = nothing
		set rs2 = nothing
		set cn = nothing



%>


<INPUT type="hidden" id=txtHideTables name=txtHideTables value=<%=strHideTables%>>
</form>
<INPUT type="hidden" id=hidAddCount name=hidAddCount value=<%=addCount%>>
<INPUT type="hidden" id=hidRemCount name=hidRemCount value=<%=remCount%>>
<INPUT type="hidden" id=hidRepCount name=hidRepCount value=<%=repCount%>>
</BODY>
</HTML>
