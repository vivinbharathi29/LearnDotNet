<%@ Language=VBScript %>

<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdDate_onclick(strID){
	var strDate;
	strDate = window.showModalDialog("../../mobilese/today/calDraw1.asp", document.all("txtFCS" + strID).value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strDate) != "undefined")
		{
			document.all("txtFCS" + strID).value = strDate;
		}
}


function cmdActual_onclick(strID){
	var strDate;
	strDate = window.showModalDialog("../../mobilese/today/calDraw1.asp", document.all("txtActual" + strID).value, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strDate) != "undefined")
		{
			document.all("txtActual" + strID).value = strDate;
		}
}



//-->
</SCRIPT>
</HEAD>
<STYLE>

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

<BODY LANGUAGE=javascript bgcolor=Ivory>
<form id=frmUpdate action="FCSSave.asp?isFromPulsarPlus=<%=Request("isFromPulsarPlus")%>" method=post>
<%

	dim cn
	dim rs
	dim cm
	dim strLanguage
	dim strSKU
	dim strPriority
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
	dim strIDList
	dim strTags
	dim strActualTags
	dim strActual
	dim strComment
	dim strCommentTags
	dim strFAISKU
	dim strFAISKUTags
	
	strtags=""
	strActualTags=""
	strCommentTags=""
	strFAISKUTags=""
	strIDList = ""
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
		Response.Redirect "../../NoAccess.asp?Level=1"
    else
        CurrentUserPartner = rs("PartnerID")
    end if 
    rs.Close

	
	
	
	
  'Create a recordset
	set rs = server.CreateObject("ADODB.recordset")
	
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

	'rs.open "spGetProductVersionName " & request("ID"),cn,adOpenForwardOnly
	if rs.EOF and rs.BOF then
		Response.Write "Unable to find the requested product."
		rs.Close
	else
		'Verify Access is OK
		if trim(CurrentUserPartner) <> "1" then
			if trim(rs("PartnerID")) <> trim(CurrentUserPartner) then
				set rs = nothing
				set cn=nothing
				
				Response.Redirect "../../NoAccess.asp?Level=1"
			end if
		end if
			
		Response.Write "<font size=3 face=verdana><b>" & rs("Name") & " - Edit Rollout Plan</b><BR><BR></font>"
		rs.Close
		Response.Write "<table bordercolor=tan border=1 cellspacing=1 cellpadding=2 >"
		Response.Write "<TR>"
		Response.Write "<Th align=left>Brand</Th>"
		Response.Write "<Th align=left>Localization</Th>"
		Response.Write "<Th align=left>OS</Th>"
		Response.Write "<Th align=left>Image</Th>"
		'Response.Write "<Th>Lang</Th>"
		Response.Write "<Th align=left>RTM</Th>"
		Response.Write "<Th align=left>Target&nbsp;FCS</Th>"
		Response.Write "<Th style=""display:none"" align=left>Actual</Th>"
		Response.Write "<Th style=""display:none"" align=left>FAI SKU</Th>"
		Response.Write "<Th align=left>Comments</Th>"
		Response.Write "</tr>"
		
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
		strFCS = ""
		strActual = ""
		strFAISKU = ""
		strComment = ""
		
		dim i

		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListImagesForProductrolloutFusion" 'spListImagesForProductAll
		

		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = request("ID")
		cm.Parameters.Append p
	

		rs.CursorType = adOpenForwardOnly
		rs.LockType=AdLockReadOnly
		Set rs = cm.Execute 
		Set cm=nothing

'		rs.Open "spListImagesForProductAll " & request("ID"),cn,adOpenForwardOnly
		i=0
		do while not rs.EOF
		
			strSKU = ucase(rs("ProductDrop") & "")
			'strSKU = replace(strSKU,"-XX",replace(ucase(rs("Dash")& ""),"X",""))
			strLanguage = rs("OSLanguage") & ""
			if rs("OtherLanguage") <> "" then
				strLanguage = strLanguage & "," & rs("OtherLanguage")
			end if

			strFCS=trim(rs("FCSDate") & "")
			strActual=trim(rs("FCSActual") & "")
			strFAISKU = trim(rs("FAISKU") & "")
			strComment = trim(rs("Comments") & "")

			strtags=strtags & "," & strFCS
			strActualtags=strActualtags & "," & strActual
			strCommenttags=strCommenttags & "," & strComment
			strFAISKUtags=strFAISKUtags & "," & strFAISKU
			strIDList = strIDList & "," & rs("ID")

			
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
				if trim(rs("Priority")) = "1" or trim(rs("Priority")) = "2" or trim(rs("Priority")) = "3" or trim(rs("Priority")) = "4" or trim(rs("Priority")) = "5" or trim(rs("Priority")) = "6" or trim(rs("Priority")) = "7" or trim(rs("Priority")) = "8" or trim(rs("Priority")) = "9" or trim(rs("Priority")) = "10" then
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetRolloutDate"
		

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = request("ID")
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@Priority", 3, &H0001)
					p.Value = rs("Priority")
					cm.Parameters.Append p
	

					rs2.CursorType = adOpenForwardOnly
					rs2.LockType=AdLockReadOnly
					Set rs2 = cm.Execute 
					Set cm=nothing

					'strSQl = "spGetRolloutDate " & request("ID") & "," & rs("Priority")				
				else
					set cm = server.CreateObject("ADODB.Command")
					Set cm.ActiveConnection = cn
					cm.CommandType = 4
					cm.CommandText = "spGetRolloutDate"
		

					Set p = cm.CreateParameter("@ID", 3, &H0001)
					p.Value = request("ID")
					cm.Parameters.Append p

					Set p = cm.CreateParameter("@Priority", 3, &H0001)
					p.Value = 0
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@Dash", 200, &H0001,10)
					p.Value = trim(rs("Priority"))
					cm.Parameters.Append p
	
					Set p = cm.CreateParameter("@ImageDefID", 3, &H0001)
					p.Value = rs("DefinitionID")
					cm.Parameters.Append p
	

					rs2.CursorType = adOpenForwardOnly
					rs2.LockType=AdLockReadOnly
					Set rs2 = cm.Execute 
					Set cm=nothing

					'strSQl = "spGetRolloutDate " & request("ID") & ",0,'" & trim(rs("Priority")) & "'," & rs("DefinitionID")
				end if
				'rs2.open strSQL,cn,adOpenForwardOnly
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
						strPriority = "Tier " & rs2("RTM")
					else
						strPriority = rs2("RTM") & ""
					end if
				else
					strPriority = "Tier " & rs("Priority")
				end if
				rs2.Close
				set rs2 = nothing
			end if
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

				Response.Write "<TR><TD nowrap>" & strImageBrandSummary & "</TD>"
				Response.Write "<TD nowrap>" & rs("Region") & "</TD>"
				Response.Write "<TD nowrap>" & rs("OS") & "</TD>"
				Response.Write "<TD nowrap>" & strSKU & "&nbsp;</TD>"
				'Response.Write "<TD nowrap>" & strLanguage & "</TD>"
				Response.Write "<TD nowrap>" & strPriority & "</TD>"
				Response.Write "<TD nowrap><INPUT type=""text"" style=""MARGIN-TOP: -23px;WIDTH:80;FONT-SIZE: xx-small"" id=txtFCS" & trim(i) & " name=txtFCS value=""" & strFCS & """>"
				%>
				<a href="javascript: cmdDate_onclick(<%=i%>)" tabindex=-1><img ID="picTarget" SRC="../../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="20"></a>
				<%
				Response.Write "<TD style=""Display:none"" nowrap><INPUT type=""text"" style=""MARGIN-TOP: -23px;WIDTH:80;FONT-SIZE: xx-small"" id=txtActual" & trim(i) & " name=txtActual value=""" & strActual & """>"
				%>
				<a href="javascript: cmdActual_onclick(<%=i%>)" tabindex=-1><img style="display:none" ID="picTarget" SRC="../../images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="20"></a>
				<%
				Response.Write "<TD style=""Display:none"" nowrap><INPUT type=""text"" style=""WIDTH:80;FONT-SIZE: xx-small"" maxlength=20 id=txtFAISKU" & trim(i) & " name=txtFAISKU value=""" & strFAISKU & """>"
				Response.Write "<TD nowrap><INPUT type=""hidden"" style=""WIDTH:160;FONT-SIZE: xx-small"" maxlength=256 id=txtComments" & trim(i) & " name=txtComments value=""" & strComment & """>"
				response.write rs("DefinitionComments") & "&nbsp;</TD>"
				Response.Write "</tr>"
				i=i+1
			rs.MoveNext
		loop
		rs.Close
		Response.Write "</table>"
	
	end if
	
	set rs = nothing
	set cn = nothing

if strTags <> "" then
	strTags = mid(strTags,2)
end if
if strActualTags <> "" then
	strActualTags = mid(strActualTags,2)
end if
if strFAISKUTags <> "" then
	strFAISKUTags = mid(strFAISKUTags,2)
end if
if strCommentTags <> "" then
	strCommentTags = mid(strCommentTags,2)
end if
if strIDList <> "" then
	strIDList = mid(strIDList,2)
end if

%>

<INPUT type="hidden" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtDateTag name=txtDateTag value="<%=strTags%>">
<INPUT type="hidden" id=txtActualTag name=txtActualTag value="<%=strActualTags%>">
<TEXTAREA style="Display:none" rows=2 cols=20 id=txtCommentTag name=txtCommentTag><%=strCommentTags%></TEXTAREA>
<TEXTAREA style="Display:none" rows=2 cols=20 id=txtFAISKUTag name=txtFAISKUTag><%=strFAISKUTags%></TEXTAREA>



<INPUT type="hidden" id=txtIDList name=txtIDList value="<%=strIDList%>">
</form>
</BODY>
</HTML>
