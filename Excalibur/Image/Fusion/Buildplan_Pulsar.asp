<%@ Language=VBScript %>
<%
    Dim AppRoot : AppRoot = Session("ApplicationRoot")

	'if request("cboFormat")= 1 then
	'	Response.ContentType = "application/vnd.ms-excel"
	'elseif request("cboFormat")= 2 then
	'	Response.ContentType = "application/msword"
	'else
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
	'end if
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "../../_ScriptLibrary/sort.js" -->

function FilterMatrix(Function,strValue) {
	if (Function==1)
		window.location.href = document.location.href + "&Brand=" + strValue;
	else if (Function==2)
		window.location.href = document.location.href + "&Opsys=" + strValue;
	else if (Function==3)
		window.location.href = document.location.href + "&Region=" + strValue;
	else if (Function==4)
		window.location.href = document.location.href + "&Image=" + encodeURIComponent(strValue);
	else if (Function==5)
		window.location.href = document.location.href + "&Language=" + strValue;
	else if (Function==6)
		window.location.href = document.location.href + "&Priority=" + strValue;
	else if (Function==7)
		window.location.href = document.location.href + "&FCS=" + strValue;
	else if (Function==8)
		window.location.href = document.location.href + "&Actual=" + strValue;
	else if (Function==9)
		window.location.href = document.location.href + "&FAISKU=" + strValue;
	else if (Function==10)
		window.location.href = document.location.href + "&Comments=" + strValue;
	else if (Function==11)
		window.location.href = document.location.href + "&HPCode=" + strValue;
	else if (Function==12)
		window.location.href = document.location.href + "&Dash=" + strValue;
	else if (Function==13)
		window.location.href = document.location.href + "&SW=" + strValue;
	else if (Function==14)
	    window.location.href = document.location.href + "&Geo=" + strValue;
	else if (Function==15)
	    window.location.href = document.location.href + "&ProductReleaseID=" + strValue;

}

function FilterMenu(strValue){
	window.location.href = document.location.href + "&FCS=" + txtFilter.value;
}

function EditFCS(){
	var strRC;
	strRC = window.showModalDialog("FCSEdit_Pulsar.asp?ID=" + txtID.value,"","dialogWidth:700px;dialogHeight:600px;edge: Raised;center:Yes; help: No;resizable: Yes;status: No"); 
	if (typeof(strRC) != "undefined")
		{
			document.location.reload();
		}

}

function window_onload() {
	lblLoading.innerText = "Click any value to filter matrix"
}

function HeaderMouseOver(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function HeaderMouseOut(){
	window.event.srcElement.style.color="black";
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
a.Filter:visited
{
	COLOR: blue;
	TEXT-DECORATION: none;
}
a.Filter
{
	COLOR: blue;
	TEXT-DECORATION: none;
}
a.Filter:hover
{
	BACKGROUND-COLOR:thistle;
	COLOR: black;
}
a:visited
{
	COLOR:blue;
}
a:hover
{
	COLOR: red;
}

td.normal
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: VERDANA;
	BACKGROUND-COLOR:ivory;
}
td.selected
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: VERDANA;
	BACKGROUND-COLOR:thistle;
}
td.alert
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: VERDANA;
	BACKGROUND-COLOR:mistyrose;
}

td
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: VERDANA;
}

th
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: VERDANA;
	FONT-WEIGHT: bold;
	BACKGROUND-COLOR:cornsilk;
 }
</STYLE>

<BODY LANGUAGE=javascript onload="return window_onload()">
<P>

<%
	dim cn
	dim cm
	dim p
	dim rs
	dim strLanguage
	dim strImage
	dim strPriority
	dim strHPCode
	dim strDash
	dim strP1
	dim strP2
	dim strP3
	dim strP4
	dim strP5
	dim strP6
	dim strP7
	dim strP8
	dim strP9
	dim strP10
	dim strSQL
	dim strEditOK
	dim blnDateAlert
	dim CurrentUser
	dim CurrentUserID
	dim ColCount
	dim CurrentUserPartner
    dim RowCount
	dim strReleaseName

	strEditOK = 0
  	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

  'Create a recordset
	set rs = server.CreateObject("ADODB.recordset")
	
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
	

	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../../NoAccess.asp?Level=0"
	else
		CurrentUserID = rs("ID")
		CurrentUserPartner = rs("PartnerID")
	end if
	rs.Close

		
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetRampPlanUpdateAccessList"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing


	do while not rs.EOF	
		if trim(rs("ID")) = trim(CurrentUserID) then
			strEditOK = 1
			exit do
		end if
		rs.Movenext
	loop
	rs.Close


	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetProductVersionName"
	

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = request("ID")
	cm.Parameters.Append p


	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing



	if rs.EOF and rs.BOF then
		Response.Write "Unable to find the requested product."
		rs.Close
	else
	
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../../NoAccess.asp?Level=0"
			end if
		end if
	
	
		strFilters = ""

		Response.Write "<font size=3 face=verdana><b>" & rs("Name")
		if Request("Report") = "1" then
			Response.Write  " Ramp Plan</b><BR><BR></font>"
		else
			Response.Write  " Rollout Plan</b><BR><BR></font>"
		end if
		
		if request("Brand") <> "" then
			strFilters = strFilters & ", " & request("Brand")
		end if

		if request("Geo") <> "" then
			strFilters = strFilters & ", " & request("Geo")
		end if
		if request("Opsys") <> "" then
			strFilters = strFilters & ", " & request("Opsys")
		end if
		if request("Language") <> "" then
			strFilters = strFilters & ", " & request("Language")
		end if
		if request("Region") <> "" then
			strFilters = strFilters & ", " & request("Region")
		end if
		if request("Dash") <> "" then
			strFilters = strFilters & ", " & request("Dash")
		end if

		if request("HPCode") <> "" then
			strFilters = strFilters & ", " & request("HPCode")
		end if

		if request("Image") <> "" then
			strFilters = strFilters & ", " & request("Image")
		end if

		if request("Priority") <> "" then
			strFilters = strFilters & ", " & request("Priority")
		end if

		if request("FCS") <> "" then
			strFilters = strFilters & ", " & request("FCS")
		end if

		if request("Actual") <> "" then
			strFilters = strFilters & ", " & request("Actual")
		end if
		
		if request("FAISKU") <> "" then
			strFilters = strFilters & ", " & request("FAISKU")
		end if

		if request("Comments") <> "" then
			strFilters = strFilters & ", " & request("Comments")
		end if
		
		
		if strFilters <> "" then
			strFilters = mid(strFilters,3)
			Response.Write "<font size=2 color=black face=verdana><b>Filtered By:</b> " & strFilters & "&nbsp;&nbsp;<a href=""http://" & Application("Excalibur_ServerName") & AppRoot & "/image/fusion/Buildplan.asp?ID=" & Request("ID") & """><font size=1>Show All</font></a><BR><BR></font>"
		end if
		Response.Write "<font ID=lblLoading size=2 color=green face=verdana>Loading.  Please wait...<BR><BR></font>"
		
		rs.Close
		if strEditOK then
			Response.Write "<table width=""100%""><TR><TD align=right><a href=""javascript:EditFCS();"">Edit Rollout Plan</a></td></tr></table>"
		end if
		Response.Write "<table ID=plantable bordercolor=tan border=1 cellspacing=1 width=""100%"" cellpadding=2 >"
		Response.Write "<THEAD>"
        if trim(request("Geo")) <> "" then
    		Response.Write "<Th>Geo</Th>"
        else
    		Response.Write "<Th onclick=""SortTable( 'plantable', 0 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Geo</Th>"
        end if
        if trim(request("Brand")) <> "" then
		    Response.Write "<Th>Brand</Th>"
		else
		    Response.Write "<Th onclick=""SortTable( 'plantable', 1 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Brand</Th>"
        end if
        if trim(request("Region")) <> "" then
            Response.Write "<Th>Localization</Th>"
        else
            Response.Write "<Th onclick=""SortTable( 'plantable', 2 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Localization</Th>"
        end if
        if trim(request("HPCode")) <> "" then
		    Response.Write "<Th>HP Code</Th>"
		else
            Response.Write "<Th onclick=""SortTable( 'plantable', 3 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">HP Code</Th>"
        end if
        if trim(request("Opsys")) <> "" then
		    Response.Write "<Th>OS</Th>"
		else
            Response.Write "<Th onclick=""SortTable( 'plantable', 4 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">OS</Th>"
        end if
        
        if trim(request("ProductReleaseID")) <> "" then
		    Response.Write "<Th>Release</Th>"
		else
            Response.Write "<Th onclick=""SortTable( 'plantable', 5 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Release</Th>"
        end if

  
        if trim(request("Image")) <> "" then
        	Response.Write "<Th>Product&nbsp;Drop</Th>"
        else	
            Response.Write "<Th onclick=""SortTable( 'plantable', 6 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Product&nbsp;Drop</Th>"
        end if	
        if trim(request("Dash")) <> "" then
        	Response.Write "<Th>Dash Code</Th>"
	    else
        	Response.Write "<Th onclick=""SortTable( 'plantable', 7 ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Dash Code</Th>"
        end if
    	ColCount = 8
		if Request("Report") <> "1" then
            if trim(request("Language")) <> "" then
			    Response.Write "<Th>Lang</Th>"
			else
                Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">Lang</Th>"
            end if
			ColCount = ColCount + 1
		end if
        if trim(request("Priority")) <> "" then
    		Response.Write "<Th>RTM</Th>"
	    else
    		Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">RTM</Th>"
        end if
    	ColCount = ColCount + 1

        if trim(request("FCS")) <> "" then
    		Response.Write "<Th>FCS Target</Th>"
	    else
    		Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"">FCS Target</Th>"
        end if
    	ColCount = ColCount + 1

		if Request("Report") = "1" then
            if trim(request("Actual")) <> "" then
	    		Response.Write "<Th>FCS Actual</Th>"
	    	else	
                Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,2,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" nowrap>FCS Actual</Th>"
            end if		
        	ColCount = ColCount + 1
            if trim(request("FAISKU")) <> "" then
    			Response.Write "<Th>FAI SKU</Th>"
	        else
    			Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" nowrap>FAI SKU</Th>"
            end if
    		ColCount = ColCount + 1
		end if
        if trim(request("Comments")) <> "" then
		    Response.Write "<Th>Comments</Th>"
        else
		    Response.Write "<Th onclick=""SortTable( 'plantable', " & ColCount & " ,0,2);"" onmouseout=""javascript: HeaderMouseOut();"" onmouseover=""javascript: HeaderMouseOver();"" nowrap>Comments</Th>"
        end if
		ColCount = ColCount + 1

		Response.Write "</tHEAD>"
		
		strP1 = ""
		strP2 = ""
		strP3 = ""
		strP4 = ""
		strP5 = ""
		strP6 = ""
		strP7 = ""
		strP8 = ""
		strP9 = ""
		strP10 = ""
		strFCS=""

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "usp_Image_ListImageForProductRollout"
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p

        if (request("ProductReleaseID")<>"" and Not IsNull(request("ProductReleaseID"))) then
			Set p = cm.CreateParameter("@ProductReleaseID", 3, &H0001)
			p.Value = request("ProductReleaseID")
			cm.Parameters.Append p
		end if 


		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing
		
        RowCount = 0
	
		do while not rs.EOF

            set rs2 = server.CreateObject("ADODB.recordset")
            strImageBrandSummary = ""
		    rs2.open "spListImageDefinitionBrands " & rs("definitionID"),cn,adOpenForwardOnly
            do while not rs2.eof
                strImageBrandSummary = strImageBrandSummary & ", " & rs2("Brand")
                rs2.movenext
            loop
		    rs2.close
            if trim(strImageBrandSummary) <> "" then
                strImageBrandSummary = mid(strImageBrandSummary,3) 
            end if       
            set rs2=nothing

			strDash = rs("Dash")& ""
			
			strImage = ucase(rs("ProductDrop") & "")
			strLanguage = rs("OSLanguage") & ""
			if rs("OtherLanguage") <> "" then
				strLanguage = strLanguage & "," & rs("OtherLanguage")
			end if
			
			strFCS=trim(rs("FCSDate") & "")
			if strFCS = "" then
				strFCS="&nbsp;"
			end if
			strActual=trim(rs("FCSActual") & "")
			if strActual = "" then
				strActual="&nbsp;"
			end if
			strFAISKU=trim(rs("FAISKU") & "")
			if strFAISKU = "" then
				strFAISKU="&nbsp;"
			end if
			strComments=replace(trim(rs("DefinitionComments") & ""),"'","")
			if strComments = "" then
				strComments="&nbsp;"
			end if

			strHPCode = trim(rs("OptionConfig") & "")
			if strHPCode = "" then
				strHPCode="&nbsp;"
			end if
			
			strPriority = ""
			if (trim(rs("Priority")) = "1" and strP1 <> "") then
				strPriority = strP1
			elseif (trim(rs("Priority")) = "2" and strP2 <> "") then
				strPriority = strP2
			elseif (trim(rs("Priority")) = "3" and strP3 <> "") then
				strPriority = strP3
			elseif (trim(rs("Priority")) = "4" and strP4 <> "") then
				strPriority = strP4
			elseif (trim(rs("Priority")) = "5" and strP5 <> "") then
				strPriority = strP5
			elseif (trim(rs("Priority")) = "6" and strP6 <> "") then
				strPriority = strP6
			elseif (trim(rs("Priority")) = "7" and strP7 <> "") then
				strPriority = strP7
			elseif (trim(rs("Priority")) = "8" and strP8 <> "") then
				strPriority = strP8
			elseif (trim(rs("Priority")) = "9" and strP9 <> "") then
				strPriority = strP9
			elseif (trim(rs("Priority")) = "10" and strP10 <> "") then
				strPriority = strP10
			else
				set rs2 = server.CreateObject("ADODB.recordset")
				strSQl = ""
				
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spGetRolloutDate"
	

				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = request("ID")
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Priority", 3, &H0001)
				if isnumeric(rs("Priority")) then 'trim(rs("Priority")) = "1" or trim(rs("Priority")) = "2" or trim(rs("Priority")) = "3" or trim(rs("Priority")) = "4" or trim(rs("Priority")) = "5" or trim(rs("Priority")) = "6" then
					p.Value = rs("Priority")
				else
					p.value=0
				end if
				cm.Parameters.Append p

				if not isnumeric(rs("Priority")) then '(trim(rs("Priority")) = "1" or trim(rs("Priority")) = "2" or trim(rs("Priority")) = "3" or trim(rs("Priority")) = "4" or trim(rs("Priority")) = "5" or trim(rs("Priority")) = "6") then
					Set p = cm.CreateParameter("@Dash", 200, &H0001,10)
					p.Value = left(trim(rs("Priority")),10)
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@ImageDefID", 3, &H0001)
					p.Value = rs("DefinitionID")
					cm.Parameters.Append p
				end if
				
								
				rs2.CursorType = adOpenForwardOnly
				rs2.LockType=AdLockReadOnly
				Set rs2 = cm.Execute 
				Set cm=nothing
				
				if isdate(rs2("RTM")) then
					strPriority = formatdatetime(rs2("RTM") & "",vbshortdate)
					if trim(rs("Priority")) = "1" then
						strP1 = strPriority
					elseif trim(rs("Priority")) = "2" then
						strP2 = strPriority
					elseif trim(rs("Priority")) = "3" then
						strP3 = strPriority					
					elseif trim(rs("Priority")) = "4" then
						strP4 = strPriority					
					elseif trim(rs("Priority")) = "5" then
						strP5 = strPriority					
					elseif trim(rs("Priority")) = "6" then
						strP6 = strPriority					
					elseif trim(rs("Priority")) = "7" then
						strP7 = strPriority					
					elseif trim(rs("Priority")) = "8" then
						strP8 = strPriority					
					elseif trim(rs("Priority")) = "9" then
						strP9 = strPriority					
					elseif trim(rs("Priority")) = "10" then
						strP10 = strPriority					
					end if

				elseif rs2("RTM") & "" <> "" then
					if isnumeric(rs2("RTM") & "") then
						strPriority = "Tier " & trim(rs2("RTM"))
					else
						strPriority = rs2("RTM") & ""
					end if
				elseif isnumeric (rs("Priority")) then
					strPriority = "Tier " & trim(rs("Priority"))
				else
					strPriority =  trim(rs("Priority"))
				end if
				rs2.Close
				set rs2 = nothing
			end if
			if  (request("Geo") = rs("Geo") or  request("Geo") = "")  and (request("Opsys") = rs("OS") or  request("Opsys") = "") and (request("Region") = rs("Region") or  request("Region") = "") and (request("Image") = strImage or  request("Image") = "") and (request("Dash") = strDash or  request("Dash") = "") and (request("FCS") = strFCS or  request("FCS") = "" or ( strFCS="&nbsp;" and request("FCS") = "TBD" ) )and (request("Actual") = strActual or  request("Actual") = "" or ( strActual="&nbsp;" and request("Actual") = "TBD" ))and (request("HPCode") = strHPCode or  request("HPCode") = "" or ( strHPCode="&nbsp;" and request("HPCode") = "TBD" ))and (request("FAISKU") = strFAISKU or  request("FAISKU") = ""  or ( strFAISKU="&nbsp;" and request("FAISKU") = "TBD" ))and (request("Comments") = strComments or  request("Comments") = "" or ( strComments="&nbsp;" and request("Comments") = "TBD" )) and (request("Priority") = strPriority or  request("Priority") = "") and (instr(strLanguage,request("Language"))>0 or  request("Language") = "") then 'and (request("SW") = rs("SW") or  request("SW") = "") and (request("Brand") = rs("Brand") or  request("Brand") = "")
				RowCount = RowCount + 1
                Response.Write "<TR>"
				if request("Geo") = rs("Geo") then
					Response.Write "<TD nowrap class=selected>" & rs("Geo") & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(14,'" & rs("Geo") & "')"">" & rs("Geo") & "</a></TD>"
				end if
				if trim(request("Brand")) = trim(strImageBrandSummary) then
					Response.Write "<TD nowrap class=selected>" & strImageBrandSummary & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(1,'" & strImageBrandSummary & "')"">" & strImageBrandSummary & "</a></TD>"
				end if
				if request("Region") = rs("Region") then
					Response.Write "<TD nowrap class=selected>" & rs("Region") & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(3,'" & rs("Region") & "')"">" & rs("Region") & "</a></TD>"
				end if
				
				if request("HPCode") = strHPCode and strHPCode = "" then
					Response.Write "<TD nowrap class=normal>&nbsp;</TD>"
				elseif request("HPCode") = strHPCode or (request("HPCode") = "TBD" and strHPCode="&nbsp;" )then
					Response.Write "<TD nowrap class=selected>" & strHPCode & "</TD>"
				else
					if strHPCode = "&nbsp;" then
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(11,'TBD')"">" & strHPCode & "</a></TD>"
					else
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(11,'" & strHPCode & "')"">" & strHPCode & "</a></TD>"
					end if
				end if

				if request("Opsys") = rs("OS") then
					Response.Write "<TD nowrap class=selected>" & rs("OS") & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(2,'" & rs("OS") & "')"">" & rs("OS") & "</a></TD>"
				end if

                if request("ProductReleaseID") = rs("ReleaseName") then
					Response.Write "<TD nowrap class=selected>" & rs("ReleaseName") & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(15,'" & rs("ReleaseName") & "')"">" & rs("ReleaseName") & "</a></TD>"
				end if
				

				if request("Image") = strImage and strImage = "" then
					Response.Write "<TD nowrap class=normal>&nbsp;</TD>"
				elseif request("Image") = strImage then
					Response.Write "<TD nowrap class=selected>" & strImage & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(4,'" & strImage & "')"">" & strImage & "</a></TD>"
				end if

				if request("Dash") = strDash and strDash = "" then
					Response.Write "<TD nowrap class=normal>&nbsp;</TD>"
				elseif request("Dash") = strDash then
					Response.Write "<TD nowrap class=selected>" & strDash & "</TD>"
				else
					Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(12,'" & strDash & "')"">" & strDash & "</a></TD>"
				end if

				if Request("Report") <> "1" then

					if instr(strLanguage,request("Language"))>0 and request("Language") <> "" then
						Response.Write "<TD nowrap class=selected>" & strLanguage & "</TD>"
					else
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(5,'" & strLanguage & "')"">" & strLanguage & "</a></TD>"
					end if
				end if

				blnDateAlert = false				
				if isdate(strFCS) and isdate(strPriority)and request("Report") <> "1" then
					if datediff("d",strFCS, strPriority) > -14 then					
						blnDateAlert = true
					end if
				end if

				if request("Priority") = strPriority then
					if blnDateAlert then
						Response.Write "<TD class=alert>" & strPriority & "</TD>"
					else
						Response.Write "<TD class=selected>" & strPriority & "</TD>"
					end if
				else
					if blnDateAlert then
						Response.Write "<TD nowrap class=alert><a class=Filter href=""javascript:FilterMatrix(6,'" & strPriority & "')"">" & strPriority & "</a></TD>"
					else
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(6,'" & strPriority & "')"">" & strPriority & "</a></TD>"
					end if
				end if

				blnDateAlert = false				
				if isdate(strFCS) and trim(strActual)="" and request("Report") = "1" then
					if datediff("d",strFCS, Now()) > 0 then					
						blnDateAlert = true
					end if
				end if

				if (request("FCS") = strFCS and strFCS <> "" ) or (request("FCS") = "TBD" and strFCS = "&nbsp;" ) then
					if blnDateAlert then
						Response.Write "<TD class=alert>" & strFCS & "&nbsp;</TD>"
					else
						Response.Write "<TD class=selected>" & strFCS & "&nbsp;</TD>"
					end if
				else
					if strFCS = "&nbsp;" then
						if blnDateAlert then
							Response.Write "<TD class=alert><a class=Filter href=""javascript:FilterMatrix(7,'TBD')"">" & strFCS & "</a>&nbsp;</TD>"
						else
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(7,'TBD')"">" & strFCS & "</a>&nbsp;</TD>"
						end if
					else
						if blnDateAlert then
							Response.Write "<TD class=alert><a class=Filter href=""javascript:FilterMatrix(7,'" & strFCS & "')"">" & strFCS & "</a>&nbsp;</TD>"
						else
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(7,'" & strFCS & "')"">" & strFCS & "</a>&nbsp;</TD>"
						end if
					end if
				end if
				
				
				if Request("Report") = "1" then
				
					if request("Actual") = strActual and strActual = "" then
						Response.Write "<TD class=normal>&nbsp;</TD>"
					elseif request("Actual") = strActual or (request("Actual")="TBD" and strActual="&nbsp;") then
						Response.Write "<TD class=selected>" & strActual & "</TD>"
					else
						if strActual = "&nbsp;" then
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(8,'TBD')"">" & strActual & "</a></TD>"
						else
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(8,'" & strActual & "')"">" & strActual & "</a></TD>"
						end if
					end if

					if request("FAISKU") = strFAISKU and strFAISKU = "" then
						Response.Write "<TD class=normal>&nbsp;</TD>"
					elseif request("FAISKU") = strFAISKU or (request("FAISKU")="TBD" and strFAISKU="&nbsp;") then
						Response.Write "<TD class=selected>" & strFAISKU & "</TD>"
					else
						if strFAISKU = "&nbsp;" then
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(9,'TBD')"">" & strFAISKU & "</a></TD>"
						else
							Response.Write "<TD class=normal><a class=Filter href=""javascript:FilterMatrix(9,'" & strFAISKU & "')"">" & strFAISKU & "</a></TD>"
						end if
					end if
				end if

				if request("Comments") = strComments and strComments = "" then
					Response.Write "<TD nowrap class=normal>&nbsp;</TD>"
				elseif request("Comments") = strComments  or (request("Comments")="TBD" and strComments="&nbsp;") then
					Response.Write "<TD nowrap class=selected>" & strComments & "</TD>"
				else
					if strComments = "&nbsp;" then
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(10,'TBD')"">" & strComments & "</a></TD>"
					else
						Response.Write "<TD nowrap class=normal><a class=Filter href=""javascript:FilterMatrix(10,'" & strComments & "')"">" & strComments & "</a></TD>"
					end if
				end if

				Response.Write "</tr>"
			end if
			rs.MoveNext
		loop
		rs.Close
		Response.Write "</table>"
	
	end if
	
	set rs = nothing
	set cn = nothing
    
	Response.Write "<BR><font size=1 face=verdana>Rows: " & RowCount & "<BR><BR></font>"
	Response.Write "<BR><BR><BR><font size=1 face=verdana>Generated: " & now() & "<BR><BR></font>"

%>

</P>
<INPUT type="hidden" id=txtFilter name=txtFilter value="">
<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtAdmin name=txtAdmin value="<%=strEditOK%>">
</BODY>
</HTML>
