
<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<html>
<head>

  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
<META name=VI60_defaultClientScript content=JavaScript>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="Style/programoffice.css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function EditOTSPM(strID, strOwnerID){

	var strResult;
	strResult = window.showModalDialog("ChooseComponentOwner.asp?ID=" + strID + "&RoleID=1&OwnerID=" + strOwnerID,"","dialogWidth:400px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
		{
		if (strResult[1] != "" && strResult[0] != "")
			document.all("OTSPM" + strID).innerHTML = "<a href='javascript: EditOTSPM(" + strID + "," + strResult[0]+ ")'>" + strResult[1] + "</a>";
		else
			alert("Unable to update the selected owner.");
		}
		
}

function EditOTSDeveloper(strID, strOwnerID){

	var strResult;
	strResult = window.showModalDialog("ChooseComponentOwner.asp?ID=" + strID + "&RoleID=2&OwnerID=" + strOwnerID,"","dialogWidth:400px;dialogHeight:200px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
		{
		if (strResult[1] != "" && strResult[0] != "")
			document.all("OTSDeveloper" + strID).innerHTML = "<a href='javascript: EditOTSDeveloper(" + strID + "," + strResult[0]+ ")'>" + strResult[1] + "</a>";
		else
			alert("Unable to update the selected owner.");
		}
		
}

//-->
</SCRIPT>
</head>
<style>
.OTSComponentCell
{
    BORDER-TOP: gainsboro thin solid;
    FONT-SIZE: xx-small;
    VERTICAL-ALIGN: middle;
    LINE-HEIGHT: 15px;
    FONT-FAMILY: Verdana
}
</style>
<link rel="stylesheet" type="text/css" href="../../style/wizard%20style.css">
<body bgcolor=Ivory>
<font face=verdana>

<%
	dim cn
	dim rs
	dim cm
	dim p
	dim CnString
    dim blnItemsFound
    dim blnPM
    
    blnItemsFound = false
    blnPM = false
    
	cnString =Session("PDPIMS_ConnectionString")
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = cnString
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")
	rs.ActiveConnection = cn


' get current user info.
	dim CurrentDomain
	dim CurrentUser
	dim CurrentUserPhone
	dim CurrentUserID
	
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
	Set	rs = cm.Execute 
	
	set cm=nothing	
		
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	end if
	rs.Close



	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetProductVersion"
		
	Set p = cm.CreateParameter("@ID", 3, &H0001)
	p.Value = clng(request("ID"))
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	if rs.EOF and rs.BOF then
		strProductName= ""
	else
		strProductName= rs("DotsName") & ""
	end if
	rs.Close

    '***PBI 8650 / Task 16199 - Harris, Valerie - Remove hard coded names and UserIDs: ---
    'if currentuserid = 8 or currentuserid=31 then
        'blnPM = true
    'else
        blnPM = false
        rs.open "spListSystemTeamDropdowns " & clng(request("ID")),cn
        do while not rs.eof
            if trim(currentuserid) = trim(rs("ID")) and lcase(rs("role") & "") = "platform development" then
                blnPM = true
                exit do
            end if
            rs.movenext
        loop
        rs.close
    'end if
        
	dim blnOTSDown
	blnOTSDown = false
	
	if not blnPM then
		Response.Write "Please contact the SEPM to update components for this product."
	elseif request("ID") = "" then 'new Product
		Response.Write "Please contact the SEPM to have components added for this product."
	else
		on error resume next
		rs.Open "spGetOTSComponentCount " & clng(request("ID")),cn,adOpenStatic
		if cn.Errors.Count > 0 then
		    blnOTSDown = true
		end if
		on error goto 0
		if blnOTSDown then
            response.write "OTS Is Currently Down"
        else
            if rs.EOF and rs.BOF then
           		Response.Write "Please contact the SEPM to have components added for this product."
            elseif rs("Excalibur") <> 0 then
                blnItemsFound = true
            end if
	        rs.Close
        end if
    end if
%>	
</font>
<%if (not blnOTSDown) and blnItemsFound then %>
<table  border="1" cellPadding="2" cellSpacing="0" width="100%" bgcolor=cornsilk bordercolor=tan >
   <TR bgcolor=Wheat><TD><b><%=strProductName%> OTS HW Common Components</b></td></tr>
<%
	if request("ID") = "" then
		rs.Open "spListProductOTSComponents 0",cn,adOpenStatic
	else
		rs.Open "spListProductOTSComponents " & clng(request("ID")) & "",cn,adOpenStatic
	end if
	if rs.EOF and rs.BOF then
%>
		<tr>
		<td><font face=verdana size=2>
			<div ID=OTSAddComponentTable>none</div>
		</font></td>
		</tr>
<%	else%>
		<TR><TD>
			<div ID=OTSAddComponentTable>
			<Table style="width:100%" cellpadding=2 cellspacing=0>
			<TR>
				<td><font face=verdana size=1><b>Err&nbsp;Type</b></font></td>
				<td><font face=verdana size=1><b>Category</b></font></td>
				<td><font face=verdana size=1><b>Component</b></font></td>
				<td><font face=verdana size=1><b>PM</b></font></td>
				<td><font face=verdana size=1><b>Developer</b></font></td>
			</TR>

<%		do while not rs.EOF
		    if rs("ID") <> 0 and lcase(rs("ErrorType")&"") = "hw" then
%>
			
				<td class=OTSComponentCell><%=rs("ErrorType")%></td>
				<td class=OTSComponentCell><%=rs("category")%></td>
				<td class=OTSComponentCell><%=rs("Component")%></td>
				<td ID=OTSPM<%=trim(rs("ID"))%> class=OTSComponentCell><a href="javascript: EditOTSPM(<%=rs("ID")%>,<%=rs("PMID")%>)"><%=longname(rs("PM")&"")%></a></td>
				<td ID=OTSDeveloper<%=trim(rs("ID"))%> class=OTSComponentCell><a href="javascript: EditOTSDeveloper(<%=rs("ID")%>,<%=rs("DeveloperID")%>)"><%=longname(rs("Developer")&"")%></a></td>
			</TR>

    <%      end if
			rs.MoveNext
		loop
%>
		</table>
		</div>
		</td></tr>
<%		
	end if
	rs.Close
%>

		
</Table>
<%end if 'OTS Down Check%>

<%

	set rs = nothing
	cn.Close
	set cn = nothing
	
	function ShortName(strName)
		if instr(strName,",")>0 then
			ShortName=mid(strName,instr(strName,",")+2,1) & ".&nbsp;" &left(strName, instr(strName,",")-1)
		else
			ShortName = strName
		end if
	end function	
	
	function LongName(strName)
		dim FirstName
		dim LastName
		dim GroupName
'		LongName = strName
		if instr(strName,",")>0 then
			FirstName = mid(strName,instr(strName,",")+2)
			LastName = left(strName, instr(strName,",")-1)
			if instr(FirstName,"(")> 0 and instr(FirstName,")")> 0 and instr(FirstName,")") > instr(FirstName,"(")  then
				GroupName = "&nbsp;" & trim(mid(FirstName,instr(FirstName,"(")))
				FirstName  = left(FirstName,instr(FirstName,"(")-2)
			end if
			if right(Firstname,6) = "&nbsp;" then
				Firstname = left(firstname,len(firstname)-6)
			end if
			LongName = FirstName & "&nbsp;" & LastName & GroupName
		else
			LongName = strName
		end if
		
	end function
	
	function CleanIDList (strIDList)
		dim IDArray
		dim strID
		dim strOut
		IDArray = split(strIDList,",")
		strOut = ""
		for each strID in IDArray
			if isnumeric(strID) then
				strOut = strOut & "," & trim(clng(strID))
			end if
		next
		if strOut <> "" then
			strOut = mid(strOut,2)
		end if
		CleanIDList = strOut
	end function

%>

 </body>
</html>