<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <script type="text/javascript" src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function cmdReset_onclick() {
	var i;
	
	for (i=0;i<lstFrom.length;i++)
		lstFrom.options[i].selected=false;
	for (i=0;i<lstTo.length;i++)
		lstTo.options[i].selected=false;
		
	cmdOK_onclick();
}

function cmdOK_onclick() {
	var OutArray = new Array();
	var strReturn = "";
	var strTemp = "";
	var strNamesFrom = "";
	var strNamesTo = "";
	var strNames= "";
	var i;
	
	for (i=0;i<lstFrom.length;i++)
		{
		if (lstFrom.options[i].selected)
			{
			if (strReturn!="")
				{
				strReturn = strReturn + "," + lstFrom.options[i].value;
				strNamesFrom = strNamesFrom + " OR " + lstFrom.options[i].text;
				}
			else
				{
				strReturn = strReturn + lstFrom.options[i].value;
				strNamesFrom = strNamesFrom + lstFrom.options[i].text;
				}
			}
		}
			
	strReturn = strReturn + ":"
	
	for (i=0;i<lstTo.length;i++)
		{
		if (lstTo.options[i].selected)
			{
			if (strTemp != "" )
				{
				strTemp = strTemp + "," + lstTo.options[i].value;
				strNamesTo = strNamesTo + " OR " + lstTo.options[i].text;
				}
			else
				{
				strTemp = strTemp + lstTo.options[i].value;
				strNamesTo = strNamesTo + lstTo.options[i].text;
				}
			}
		}

	strReturn = strReturn + strTemp;
	if (strReturn ==":")
		strReturn = "";
	
	if (strNamesFrom == "" && strNamesTo == "")
		{
		strNames = "<a href='javascript:GetSpecificChange(" + txtType.value + ");'>All Changes</a>";
		}
	else
		{
		strNames = "<TABLE>"
		if (strNamesFrom != "")
			{
			strNames = strNames + "<tr><td>FROM:</TD><TD valign=top><a href='javascript:GetSpecificChange(" + txtType.value + ");'>" + strNamesFrom + "</a></TD></TR>"
			}
		if (strNamesTo != "")
			{
			strNames = strNames + "<tr><td>TO:</TD><TD valign=top><a href='javascript:GetSpecificChange(" + txtType.value + ");'>" + strNamesTo + "</a></TD></TR>"
			}
		strNames = strNames + "</TABLE>"
		}

	OutArray[0]=strReturn;
	OutArray[1]=strNames;
	
	if (IsFromPulsarPlus()) {
	    window.parent.parent.GetSpecificChangeResult(txtType.value, OutArray);
	    ClosePulsarPlusPopup();
	}
	else {
	    window.returnValue = OutArray;//strReturn;
	    window.parent.close();
	}
}

function cmdCancel_onclick() {
	window.parent.close();

}

//-->
</SCRIPT>
</HEAD>
<STYLE>
Body
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
}
TD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
}

</STYLE>
<BODY bgcolor=ivory>


<%

if request("TypeID") = "1" then
	Response.write "Only include changes where the Pilot Status changed:"
else
	Response.write "Only include changes where the Qualification Status changed:"
end if

	dim cn, rs
	dim strFROM
	dim strTO
	dim CurrentArray
	dim strFromCurrent
	dim strToCurrent
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	strFrom = ""
	strTo = ""
	if instr(request("Current"),":") = 0 then
		strFromCurrent = ""
		strToCurrent = ""
	else
		CurrentArray = split(request("Current"),":")
		strFromCurrent = "," & CurrentArray(0) & ","
		strToCurrent = "," & CurrentArray(1) & ","
	end if
	
	if request("TypeID") = "1" then
		rs.Open "spListPilotStatus",cn,adOpenStatic
	else
		rs.Open "spListTestStatus",cn,adOpenStatic
	end if
	do while not rs.eof
		if instr(strFromCurrent,"," & trim(rs("ID")) & ",") > 0 then
			strFROM = strFrom & "<option selected value=" & rs("ID") & ">" & rs("name") & "</option>"
		else
			strFROM = strFrom & "<option value=" & rs("ID") & ">" & rs("name") & "</option>"
		end if
		if instr(strToCurrent,"," & trim(rs("ID")) & ",") > 0 then
			strTO = strTO & "<option selected value=" & rs("ID") & ">" & rs("name") & "</option>"
		else
			strTO = strTO & "<option value=" & rs("ID") & ">" & rs("name") & "</option>"
		end if
		rs.MoveNext
	loop
	rs.Close
	
	set rs = nothing
	cn.Close
	set cn = nothing

%>

<TABLE>
	<TR>
		<TD><b>FROM:</b><BR>
<SELECT style="WIDTH: 114px; HEIGHT: 200px" size=2 id=lstFrom name=lstFrom multiple> 
  <%=strFROM%></SELECT>
	</TD>
		<TD><b>TO:</b><BR>
<SELECT style="WIDTH: 114px; HEIGHT: 200px" size=2 id=lstTo name=lstTo multiple> 
  <%=strTO%></SELECT>
	</TD>
</tr></table>
<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
<INPUT type="button" value="Reset" id=cmdReset name=cmdReset LANGUAGE=javascript onclick="return cmdReset_onclick()">
<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
<INPUT type="hidden" id=txtType name=txtType value="<%=request("TypeID")%>">
<BR><font size=1 color=green>Use CTRL or SHIFT keys to select multiple items in lists </font>
</BODY>
</HTML>
