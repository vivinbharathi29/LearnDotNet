<%@ Language=VBScript %>
<HTML>
<HEAD>
<STYLE>
a:link
{
    COLOR: blue
}
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

</STYLE>

<META name="VI60_DefaultClientScript" content=JavaScript>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
var CurrentState;

function window_onload() {
	SearchBar.innerHTML = txtSearchBar.value;
	SearchBar2.innerHTML = txtSearchBar2.value + "";
	SearchBar3.innerHTML = txtSearchBar3.value + "";
	SearchBar4.innerHTML = txtSearchBar4.value + "";
	ActiveProgramFilter.innerHTML = txtActiveProgramFilter.value + "";

    if (typeof(txtProducts.value) != "undefined")
        txtProducts.focus();
}

function Goto(ProgID){
    window.location.href = "pmview.asp?Class=1&ID=" + ProgID;	
}

function GotoTool(ProgID){
	window.location.href= "pmview.asp?Class=1&ID=" + ProgID;	
}

function ShowInactive(){
	if (InactiveLink.innerText == "Hide Inactive Products")
		{
		InactiveTable.style.display="none";
		InactiveLink.innerText = "Show Inactive Products";
		}
	else
		{
		InactiveTable.style.display="";
		InactiveLink.innerText = "Hide Inactive Products";
		}
}


function ProcessState() {
	var steptext;
	
	switch (CurrentState)
	{
		case "Active":
			tabActive.style.display="";
			tabPostProduction.style.display="none";
            tabDCR.style.display="none";
            tabInactive.style.display="none";
            tabTools.style.display="none";

			window.scrollTo(0,0);		
		break;

		case "PostProduction":
			tabActive.style.display="none";
			tabPostProduction.style.display="";
            tabDCR.style.display="none";
            tabInactive.style.display="none";
            tabTools.style.display="none";

			window.scrollTo(0,0);		
		break;

		case "DCR":
			tabActive.style.display="none";
			tabPostProduction.style.display="none";
            tabDCR.style.display="";
            tabInactive.style.display="none";
            tabTools.style.display="none";

			window.scrollTo(0,0);		
		break;

		case "Inactive":
			tabActive.style.display="none";
			tabPostProduction.style.display="none";
            tabDCR.style.display="none";
            tabInactive.style.display="";
            tabTools.style.display="none";

			window.scrollTo(0,0);		
		break;

		case "Tools":
			tabActive.style.display="none";
			tabPostProduction.style.display="none";
            tabDCR.style.display="none";
            tabInactive.style.display="none";
            tabTools.style.display="";

			window.scrollTo(0,0);		
		break;

	}
}

function SelectTab(strStep) {
	var i;

	//Reset all tabs
	document.all("CellActiveb").style.display="none";
	document.all("CellActive").style.display="";
	document.all("CellPostproductionb").style.display="none";
	document.all("CellPostProduction").style.display="";
	document.all("CellDCRb").style.display="none";
	document.all("CellDCR").style.display="";
	document.all("CellInactiveb").style.display="none";
	document.all("CellInactive").style.display="";
	document.all("CellToolsb").style.display="none";
	document.all("CellTools").style.display="";

	//Highight the selected tab
	document.all("Cell"+strStep).style.display="none";
	document.all("Cell"+strStep+"b").style.display="";

	
	CurrentState = strStep;
	ProcessState();

}


function goproduct_onclick() {
	if (txtProducts.value != "")
	    window.location.href="mobilese/today/find.asp?Find=" + txtProducts.value.replace('&', '%26').replace('+', '%2B') + "&Type=Products";
	else
		{
		window.alert("Please enter search criteria first.");
		txtDeliverables.focus();
		}
}

function goID_onmouseover(){
	gochange.style.cursor = "hand";
}

function txtProducts_onkeypress() {
	if (window.event.keyCode == 13)
		goproduct_onclick();
}

//-->
</SCRIPT>
<link href="./style/wizard%20style.css" type="text/css" rel="stylesheet">
<STYLE>
th
{
 font-family:Verdana;
 font-size:xx-small;   
 font-weight:bold;
 text-align:left;
 background-color:cornsilk;
} 
    
.DisplayBar
{
    BORDER-RIGHT: gray thin solid;
    BORDER-TOP: gray thin solid;
    BORDER-LEFT: gray thin solid;
    BORDER-BOTTOM: gray thin solid;
    BACKGROUND-COLOR: gainsboro
}
</STYLE>

</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<FONT face=verdana><H3>Find Product</H3></FONT>
<%
    dim StatusArray
    StatusArray = split("Definition,Development,Production,Post-Production,Inactive,Current",",")

	on error resume next
	strTitleColor = "#0000cd"
	if Request.Cookies("TitleColor") <> "" then
		strTitleColor = Request.Cookies("TitleColor")
	else
		strTitleColor = "#0000cd"
	end if
	on error goto 0


	Dim LastProgram
	dim cm
	dim p
	dim cn
	dim CurrentUser 
	dim CurrentUserID
	dim txtSearch
	dim lastLetter
	dim strGoto
    dim CurrentUserPartnerType

	txtSearch = "<TABLE borderColor=gainsboro cellSpacing=0 cellPadding=2 border=1><TR bgColor=ivory>"
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
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
    CurrentUserPartnerType = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
		CurrentUserPartner = rs("PartnerID") & ""
        CurrentUserPartnerType = rs("PartnerTypeID") & ""
	else
		rs.Close
		set rs=nothing
		cn.Close
		set cn = nothing
		Response.Redirect "http://16.81.19.70/Excalibur.asp"
	end if

	rs.Close

'    CurrentUserPartnerType = 0
'    if CurrentUserPartner <> "1" then
'        rs.open "spgetPartnerType " & CurrentUserPartner,cn
'        if not(rs.eof and rs.bof) then
'            CurrentUserPartnerType = rs("PartnerTypeID") & ""
'        end if			'
'	    rs.close
'    end if	
%>
    <table cellpadding=0 cellspacing=0><tr><td><font face="Verdana" size="2"><strong>Search:&nbsp;&nbsp;</strong></font></td><td><input id="txtProducts" name="txtProducts" style="WIDTH: 92px; HEIGHT: 22px" size="13" LANGUAGE="javascript" onkeypress="return txtProducts_onkeypress()">&nbsp;<a><img id="gochange" border="0" src="images\go.gif" WIDTH="23" HEIGHT="20" LANGUAGE="javascript" onmouseover="return goID_onmouseover()" onclick="return goproduct_onclick()"></a>
    </td></tr> </table>
    <hr>

<table Class="MenuBar" border="1" bordercolor="Ivory" cellspacing="0" cellpadding="2">
	<tr bgcolor="<%=strTitleColor%>">
		<td id="CellActive" style="Display:none" width="10"><font size="1" color="black"><b>&nbsp;&nbsp;<a href="javascript:SelectTab('Active')">Current</a>&nbsp;&nbsp;</b></font></td>
		<td id="CellActiveb" style="Display:" width="10" bgcolor="wheat"><font size="1" color="black"><b>&nbsp;&nbsp;Current</b>&nbsp;&nbsp;</font></td>
		<td id="CellPostProduction" nowrap style="Display:" width="10"><font size="1" color="white"><b>&nbsp;&nbsp;<a href="javascript:SelectTab('PostProduction')">Post&#8722;Production</a>&nbsp;&nbsp;</b></font></td>
		<td id="CellPostProductionb"nowrap style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black"><b>&nbsp;&nbsp;Post&#8722;Production</b>&nbsp;&nbsp;</font></td>
		<td id="CellInactive" style="Display:" width="10"><font size="1" color="white"><b>&nbsp;&nbsp;<a href="javascript:SelectTab('Inactive')">Inactive</a>&nbsp;&nbsp;</b></font></td>
		<td id="CellInactiveb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black"><b>&nbsp;&nbsp;Inactive</b>&nbsp;&nbsp;</font></td>
		<%if trim(currentuserpartner)="1" then %>
        <td id="CellTools" style="Display:" width="10"><font size="1" color="white"><b>&nbsp;&nbsp;<a href="javascript:SelectTab('Tools')">Tools/Processes</a>&nbsp;&nbsp;</b></font></td>
		<td id="CellToolsb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black"><b>&nbsp;&nbsp;Tools/Processes</b>&nbsp;&nbsp;</font></td>
		<td id="CellDCR" style="Display:" width="10"><font size="1" color="white"><b>&nbsp;&nbsp;<a href="javascript:SelectTab('DCR')">DCR&nbsp;Projects</a>&nbsp;&nbsp;</b></font></td>
		<td id="CellDCRb" style="Display:none" width="10" bgcolor="wheat"><font size="1" color="black"><b>&nbsp;&nbsp;DCR&nbsp;Projects</b>&nbsp;&nbsp;</font></td>
	    <%else%>
            <div id="CellTools"></div>
            <div id="CellToolsb"></div>
            <div id="CellDCR"></div>
            <div id="CellDCRb"></div>
        <%end if%>
    </tr>
</table>
<!--<hr style="margin-top:0" color="Tan">-->
<br>
<div id=tabActive>
<Label ID=ActiveProgramFilter></Label>
<Label ID=SearchBar></Label>

  <TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory style="margin-top:4px">
    <tr><th>Name</th><th>Status</th>
    <% if (currentuserpartnertype) <> "1" then %>
        <th>ODM</th>
    <%end if%>
    <th>Dev&nbsp;Center</th><th style="width:100%">Cycles</th></tr>
    
<%

	rs.Open "spGetProducts",cn,adOpenForwardOnly
	
	LastProgram = ""
	dim strPrograms
	dim strPartner
	dim strPartnerList
	dim strProgramList
	dim strDevCenterList
	dim strDevCenterlinks
	dim strPartnerLinks
	dim strProgramLinks
	dim strProgramIDs
	strPartnerList=""
	strPartnerLinks = ""
	strProgramList=""
	strProgramLinks = ""
	strDevCenterList=""
	strDevCenterlinks = ""
    
    RowsWritten = 0

	do while not rs.EOF
		if rs("Name") <> "Not" and (trim(CurrentUserPartner) = trim(rs("PartnerID")) or trim(CurrentUserPartner)="1" or trim(CurrentUserPartnerType)="2" ) then

            if true then

				'GetProgram List
				strPrograms = ""
				strProgramIDs = ""
				set rs2 = server.CreateObject("ADODB.recordset")
				rs2.open "spListProgramsForProduct " & clng(rs("ID")),cn,adOpenStatic
				do while not rs2.EOF
					'if trim(rs2("OTSCycleName") & "") = "" then
						strPrograms = strPrograms & "," & rs2("FullName")
						strProgramIDs = strProgramIDs & "," & trim(rs2("ID"))
						if instr(strProgramList& "," , "," & rs2("FullName") & ",")=0 then
							strprogramList = strProgramlist & "," & rs2("FullName")
							if trim(request("Program")) = trim(rs2("ID")) then
								strProgramLinks = strProgramLinks & ", " & rs2("FullName")
							else
								strProgramLinks = strProgramLinks & ", <a href='findprogram.asp?Partner=" & request("Partner") & "&DevCenter=" & request("Devcenter") & "&Program=" & rs2("ID") & "'>"  & replace(rs2("FullName")," ","&nbsp;") & "</a>"
							end if
						end if
					'else
					'	strPrograms = strPrograms & ",BNB " & rs2("Name")  
					'	strProgramIDs = strProgramIDs & "," & trim(rs2("ID"))
					'	if instr(strProgramList& "," , ",BNB " & rs2("Name") & ",")=0 then
					'		strprogramList = strProgramlist & ",BNB " & rs2("Name")
					'		if trim(request("Program")) = trim(rs2("ID")) then
					'			strProgramLinks = strProgramLinks & ", " & "BNB " & rs2("Name")
					'		else
					'			strProgramLinks = strProgramLinks & ", <a href='findprogram.asp?Partner=" & request("Partner") & "&DevCenter=" & request("Devcenter") & "&Program=" & rs2("ID") & "'>" & "BNB&nbsp;" & replace(rs2("Name")," ","&nbsp;") & "</a>"
					'		end if
					'	end if
					'end if
					rs2.MoveNext
				loop
				rs2.Close	
				set rs2 = nothing
				if strPrograms="" then
					strPrograms = "&nbsp;"
				else
					strPrograms = ucase(mid(strPrograms,2))
					strProgramIDs =strProgramIDs & ","
				end if
				
				if instr( strPartnerList & ",","," & rs("Partner") & ",") = 0 then
					strPartnerList = strPartnerList & "," & rs("Partner") 
					if trim(request("Partner")) = trim(rs("PartnerID")) then
						strPartnerLinks = strPartnerLinks & ", " & rs("Partner") 
					else
						strPartnerLinks = strPartnerLinks & ", <a href='findprogram.asp?Partner=" & rs("PartnerID") & "&DevCenter=" & request("DevCenter") & "&Program=" & request("Program") & "'>" & rs("Partner") & "</a>"
					end if
				end if
				if instr( strDevCenterList & ",","," & rs("DevCenterName") & ",") = 0 then
					strDevCenterList = strDevCenterList & "," & rs("DevCenterName")
					if trim(request("DevCenter")) = trim(rs("Devcenter")) then
						strDevCenterlinks = strDevCenterlinks & ", " & rs("DevCenterName") 
					else
						strDevCenterlinks = strDevCenterlinks & ", <a href='findprogram.asp?Partner=" & request("Partner") & "&DevCenter=" & rs("Devcenter") & "&Program=" & request("Program") & "'>" & rs("DevCenterName") & "</a>"
					end if
				end if
                if lastProgram <> replace(rs("Name"),"""","&quot;") then
				lastProgram = replace(rs("Name"),"""","&quot;")

								strGoto = "<a name=""" & lastLetter & """></a>"

				'Header Row
				'strHeaderRow = "<TR><a name=" & lastProgram & "><TD width=100% colspan=4 nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>##GOTO##" & lastProgram & "</FONT></STRONG></TD></TR>"
			    end if
			end if

			if (trim(request("DevCenter")) = "" or trim(request("DevCenter")) = trim(rs("DevCenter"))) and (trim(request("Partner")) = "" or trim(request("Partner")) = trim(rs("PartnerID")))  and (trim(request("Program")) = "" or instr(strProgramIDs,","& trim(request("Program"))& ",") >0 )  then
				if lastLetter <> replace(left(rs("Name"),1),"""","&quot;") then
					lastLetter = replace(left(rs("Name"),1),"""","&quot;")
					txtsearch = txtsearch & "<TD width=5 align=middle><font size=2 face=verdana><a HREF=#" & lastletter & ">" & lastLetter & "</font></TD>"
					strGoto = "<a name=""" & lastLetter & """></a>"
				end if
				if strHeaderRow <> "" then
					Response.Write replace(strHeaderRow,"##GOTO##",strGoto)
					strHeaderRow=""
				end if
                RowsWritten = RowsWritten + 1

				Response.Write "<TR><TD nowrap>" & strGoto & "<FONT face=verdana size=1><a href=""javascript:Goto(" & rs("ID") & ")"">" & rs("Name") & " " & rs("Version") & "</a>&nbsp;&nbsp;</FONT></TD><TD><font size=1 face=verdana>" & statusarray(rs("ProductStatusID")-1) & "&nbsp;&nbsp;&nbsp;&nbsp;</font></TD>"
                if (currentuserpartnertype) <> "1" then
                    response.write "<TD><font size=1 face=verdana>" & rs("Partner") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                end if
                response.write "<TD nowrap><font size=1 face=verdana>" & rs("DevCentername") & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                response.write "<TD><font size=1 face=verdana>" & strprograms & "&nbsp;</TD></TR>"
			end if
		end if
		strGoto = ""
		rs.MoveNext
	loop
	rs.Close
    if RowsWritten = 0 then
        response.write "<tr><td colspan=5><font size=1 face=verdana>None</font></td></tr>"
    end if
	Response.Write "</Table>"
    response.write "<font size=1 face=verdana><br>Products Displayed: " & RowsWritten & "</font>"
    response.write "</div>"


	txtSearch = txtSearch & "</TR></TABLE>"
	
    if strProgramLinks <>""then
		if request("Program") = "" then
			strProgramLinks = "<TR><TD valign=top><font size=1 face=verdana><b>Cycle: </b></font></td><td><font size=1 face=verdana>" & mid(strProgramLinks,2) & ", All</font></td></tr>"
		else
			strProgramLinks = "<TR><TD valign=top><font size=1 face=verdana><b>Cycle: </b></font></td><td><font size=1 face=verdana>" & mid(strProgramLinks,2) & ", <a href='findprogram.asp?Partner=" & request("Partner") & "&DevCenter=" &  request("DevCenter") & "'>All</a></font></td></tr>"
		end if
	end if
	if strPartnerLinks <> "" then
		if request("Partner") = "" then
			strPartnerLinks = "<TR><TD valign=top><font size=1 face=verdana><b>ODM: </b></font></td><td><font size=1 face=verdana>" & mid(strPartnerLinks,2) & ", All</font></td></tr>"
		else
			strPartnerLinks = "<TR><TD valign=top><font size=1 face=verdana><b>ODM: </b></font></td><td><font size=1 face=verdana>" & mid(strPartnerLinks,2) & ", <a href='findprogram.asp?DevCenter=" &  request("DevCenter") & "&Program=" & request("Program") & "'>All</a></font></td></tr>"
		end if
	end if

	if strDevCenterlinks <> "" then
		if request("DevCenter") = "" then
			strDevCenterlinks = "<TR><TD valign=top><font size=1 face=verdana><b>Dev&nbsp;Center: </b></font></td><td><font size=1 face=verdana>" & mid(strDevCenterlinks,2) & ", All</font></td></tr>"
		else
			strDevCenterlinks = "<TR><TD valign=top><font size=1 face=verdana><b>Dev&nbsp;Center: </b></font></td><td><font size=1 face=verdana>" & mid(strDevCenterlinks,2) & ", <a href='findprogram.asp?Partner=" &  request("Partner") & "&Program=" & request("Program") & "'>All</a></font></td></tr>"
		end if
	end if
	if strprogramlist <> "" and strpartnerlist <> "" and strDevCenterlist <> "" then
    	if (currentuserpartnertype) = "1" then
            strPartnerLinks = ""
        end if
        strprogramlist = "<table class=DisplayBar cellspacing=0 cellpadding=2><tr><td><table>" & strProgramLinks  & strPartnerLinks & strDevCenterlinks & "</table></td></tr></table><BR>"
	end if
%>	  
  <INPUT type="hidden" id=txtSearchBar name=txtSearchBar value ="<%=txtSearch%>">    
  <INPUT type="hidden" id=txtActiveProgramFilter name=txtActiveProgramFilter value ="<%=strprogramlist%>">    

    <div id=tabDCR style=display:none>
<% if trim(CurrentUserPartner)="1" then %>	
	
	<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%

	Response.Write "<TR><TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Project Name</FONT></STRONG></TD></TR>"	'<TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Division</FONT></STRONG></TD>
	Response.Write "<TR><TD><FONT face=verdana size=1><a href=""javascript:GotoTool(344)"">SW Spec Changes</a></FONT></TD></tr>" 
	Response.Write "<TR><TD><FONT face=verdana size=1><a href=""javascript:GotoTool(347)"">Core BIOS Changes</a></FONT></TD></tr>" 
	Response.Write "<TR><TD><FONT face=verdana size=1><a href=""javascript:GotoTool(1107)"">ID Spec Changes</a></FONT></TD></tr>" 
	Response.Write "</Table>"
    end if
%>

    </div>

<% if trim(CurrentUserPartner)="1" then %>	

    <div id=tabTools style=display:none>
	<div style="display:none"  ID=SearchBar2></div>
	
    <TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory>

<%

	txtSearch = "<TABLE borderColor=gainsboro cellSpacing=0 cellPadding=2 border=1><TR bgColor=ivory>"

	Response.Write "<TR><TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Project Name</FONT></STRONG></TD></TR>"	'<TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Division</FONT></STRONG></TD>
	
	rs.Open "spGetProducts 2",cn,adOpenForwardOnly

	do while not rs.EOF
		if rs("Name") <> "Not" then
			if lastLetter <> replace(left(rs("Name"),1),"""","&quot;") then
				lastLetter = replace(left(rs("Name"),1),"""","&quot;")
				txtsearch = txtsearch & "<TD width=5 align=middle><font size=2 face=verdana><a HREF=#T" & lastletter & ">" & lastLetter & "</font></TD>"
				strGoto = "<a name=""T" & lastLetter & """></a>"
			end if
			if rs("Division") = "1" then
				strDivision = "Mobile"
			elseif rs("Division") = "2" then
				strDivision = "bPC"
			elseif rs("Division") = "3" then
				strDivision = "cPC"
			elseif rs("Division") = "4" then
				strDivision = "ISS"
			end if
			Response.Write "<TR><TD>" & strGoto & "<FONT face=verdana size=1><a href=""javascript:GotoTool(" & rs("ID") & ")"">" & rs("Name") & " " & rs("Version") & "</a></FONT></TD></tr>" '<TD><font size=1 face=verdana>" & strDivision & "&nbsp;</TD></TR>"
		end if
		strGoto = ""
		rs.MoveNext
	loop
	rs.Close
	Response.Write "</Table>"
	txtSearch = txtSearch & "</TR></TABLE>"
	

%>

    <%else%>
        <div style="display:none" ID=SearchBar2></div>
        <div id=tabTools style=display:none></div>
    <% end if %>	
    </div>
    <INPUT type="hidden" id=txtSearchBar2 name=txtSearchBar2 value ="<%=txtSearch%>">    
    
    <div id=tabInactive style="display:none">
	<font size=1 face=verdana color=red>Warning:  These products may contain incomplete or inaccurate data.<BR><BR></font>
    <div style="display:" ID=SearchBar4></div>

	<%
	txtSearch = "<TABLE borderColor=gainsboro cellSpacing=0 cellPadding=2 border=1><TR bgColor=ivory>"

	rs.Open "spListInactiveProducts",cn,adOpenForwardOnly
	if not( rs.EOF and rs.BOF) then
		Response.Write "<Table cellspacing=1 cellpadding=2 border=1 bordercolor=tan bgcolor=ivory width=""100%"" style=""margin-top:4px"">"
		Response.Write "<TR><TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Project Name</FONT></STRONG></TD>"
        if (currentuserpartnertype) <> "1" then
            response.write "<TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>ODM</FONT></STRONG></TD>"
        end if
        response.write "<TD nowrap bgColor=cornsilk  style=""width:100%""><STRONG><FONT face=verdana size=1>DevCenter</FONT></STRONG></TD></tr>" '<TD  style=""width:100%"" nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Cycles</FONT></STRONG></TD></TR>"	
	end if	
    lastLetter = ""

    RowsWritten = 0
	do while not rs.EOF
		if (trim(CurrentUserPartner) = trim(rs("PartnerID")) or trim(CurrentUserPartner)="1" or trim(CurrentUserPartnerType)="2") then
			if rs("ID") <> 347 and rs("ID") <> 344 and rs("ID") <> 1107 then
    		
       			RowsWritten = RowsWritten + 1

            	if lastLetter <> replace(left(rs("Name"),1),"""","&quot;") then
	    			lastLetter = replace(left(rs("Name"),1),"""","&quot;")
		    		txtsearch = txtsearch & "<TD width=5 align=middle><font size=2 face=verdana><a HREF=#I" & lastletter & ">" & lastLetter & "</font></TD>"
			    	strGoto = "<a name=""I" & lastLetter & """></a>"
			    end if

                            if false then

	    		strPrograms = ""
    			set rs2 = server.CreateObject("ADODB.recordset")
	    		rs2.open "spListProgramsForProduct " & clng(rs("ID")),cn,adOpenStatic
		    	do while not rs2.EOF
			    	'if trim(rs2("OTSCycleName") & "") = "" then
				    	strPrograms = strPrograms & "," & rs2("FullName")
    				'else
	    			'	strPrograms = strPrograms & ",BNB " & rs2("Name")  
    				'end if
	    			rs2.MoveNext
		    	loop
			    rs2.Close	
			    set rs2 = nothing
			    if strPrograms="" then
				    strPrograms = "&nbsp;"
			    else
				    strPrograms = ucase(mid(strPrograms,2))
				    strProgramIDs =strProgramIDs & ","
			    end if

                end if

			    Response.Write "<TR><TD>" & strGoto & "<FONT face=verdana size=1><a href=""javascript:Goto(" & rs("ID") & ")"">" & rs("Product") & "</a>&nbsp;&nbsp;</FONT></TD>"
                if (currentuserpartnertype) <> "1" then
                    response.write "<TD nowrap><font size=1 face=verdana>" & rs("Partner") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                end if
                response.write "<TD nowrap><font size=1 face=verdana>" & rs("DevCentername") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD></tr>" '<TD nowrap><font size=1 face=verdana>" & strprograms & "&nbsp;</TD></TR>"
		    end if
		end if
		rs.MoveNext
	loop
	rs.Close
	Response.Write "</Table>"
    response.write "<font size=1 face=verdana><br>Products Displayed: " & RowsWritten & "</font>"
	
	%>
    </div>
    <INPUT type="hidden" id=txtSearchBar4 name=txtSearchBar4 value ="<%=txtSearch%>">    

    
    <div id=tabPostProduction style=display:none>
	<div ID=SearchBar3></div>
	<TABLE cellSpacing=1 cellPadding=1 width="100%" border=1 borderColor=tan bgColor=ivory style="margin-top:4px">

<%
	txtSearch = "<TABLE borderColor=gainsboro cellSpacing=0 cellPadding=2 border=1><TR bgColor=ivory>"

	Response.Write "<TR><TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Project Name</FONT></STRONG></TD>"
    if (currentuserpartnertype) <> "1" then
        response.write "<TD nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>ODM</FONT></STRONG></TD>"
    end if
    response.write "<TD nowrap bgColor=cornsilk style=""width:100%""><STRONG><FONT face=verdana size=1>DevCenter</FONT></STRONG></TD></tr>"'<TD   nowrap bgColor=cornsilk><STRONG><FONT face=verdana size=1>Cycles</FONT></STRONG></TD></TR>"
    dim RowsWritten	
	rs.Open "spListProductsByStatus 4",cn,adOpenForwardOnly
    RowsWritten = 0
	do while not rs.EOF
		if rs("DotsName") <> "Not"  and (trim(CurrentUserPartner) = trim(rs("PartnerID")) or trim(CurrentUserPartner)="1" or trim(CurrentUserPartnerType)="2") then
			RowsWritten = RowsWritten + 1
			if lastLetter <> replace(left(rs("DotsName"),1),"""","&quot;") then
				lastLetter = replace(left(rs("DotsName"),1),"""","&quot;")
				txtsearch = txtsearch & "<TD width=5 align=middle><font size=2 face=verdana><a HREF=#P" & lastletter & ">" & lastLetter & "</font></TD>"
                strGoto = "<a name=""P" & lastLetter & """></a>"
			end if

            strPrograms = ""
            if false then
	    	
    		set rs2 = server.CreateObject("ADODB.recordset")
	    	rs2.open "spListProgramsForProduct " & clng(rs("ID")),cn,adOpenStatic
		    do while not rs2.EOF
			    'if trim(rs2("OTSCycleName") & "") = "" then
				    strPrograms = strPrograms & "," & rs2("FullName")
    			'else
	    		'	strPrograms = strPrograms & ",BNB " & rs2("Name")  
    			'end if
	    		rs2.MoveNext
		    loop
			rs2.Close	
			set rs2 = nothing
			if strPrograms="" then
				strPrograms = "&nbsp;"
			else
				strPrograms = ucase(mid(strPrograms,2))
				strProgramIDs =strProgramIDs & ","
			end if

            end if

			Response.Write "<TR><TD>" & strGoto & "<FONT face=verdana size=1><a href=""javascript:GotoTool(" & rs("ID") & ")"">" & rs("DotsName") & "</a>&nbsp;&nbsp;</FONT></TD>"
            if (currentuserpartnertype) <> "1" then
                response.write "<TD nowrap><font size=1 face=verdana>" & rs("Partner") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
            end if
            response.write "<TD nowrap><font size=1 face=verdana>" & rs("DevCentername") & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD></tr>" '<TD nowrap><font size=1 face=verdana>" & strprograms & "&nbsp;</TD></tr>" 
		end if
		strGoto = ""
		rs.MoveNext
	loop
	rs.Close
	if RowsWritten = 0 then
	    response.Write "<tr><td><font face=verdana size=1>none</font></td></tr>"
	end if
	Response.Write "</Table>"
	txtSearch = txtSearch & "</TR></TABLE>"
	
    response.write "<font size=1 face=verdana><br>Products Displayed: " & RowsWritten & "</font>"

%>

	</div>


	<%	
	
		set rs=nothing
		cn.Close
		set cn=nothing


	%>
	
	
	<INPUT style="Display:none" type="text" id=txtSearchBar3 name=txtSearchBar3 value ="<%=txtSearch%>">      
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  <BR>
  
</BODY>
</HTML>
