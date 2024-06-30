<%@ Language=VBScript %>


<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }
    
function txtNew_onfocusout(strID) {
	var iTotal = 0;
	if (isNaN(trim(document.all("txtNew" + strID).value)) || trim(document.all("txtNew" + strID).value) == "")
		iTotal = 0;
	else
		iTotal = document.all("txtNew" + strID).value;

	iTotal = parseInt(iTotal) + parseInt(document.all("txtTotalLoaded" + strID).value);
	if(isNaN(trim(document.all("txtNew" + strID).value)))
		{
		alert("You must enter a number is this field.");
		document.all("txtNew" + strID).focus();
		}
	else if(iTotal < 0)
		{
		alert("You can not remove more that the previous total.");
		document.all("txtNew" + strID).focus();
		}
	else
		{
		document.all("lblTotal" + strID).innerText = iTotal;
		if ( trim(document.all("txtNew" + strID).value) == "" || trim(document.all("txtNew" + strID).value) == "0")
			document.all("txtUpdates" + strID).value = "";
		else
			document.all("txtUpdates" + strID).value = strID + "_" + iTotal + "_" + trim(document.all("txtNew" + strID).value); 
		}
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}

.DelTable TBODY TD{
	BORDER-TOP: gray thin solid;
}

TD
{
	Font-Family:Verdana;
	FONT-Size:xx-small;
}
</STYLE>
<BODY bgcolor="ivory">


<%

	dim cn
	dim rs
	dim cm
	dim p
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim strproductname
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
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
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	if request("ID") = "" then
		Response.Write "Not enough information supplied to process your request."
	elseif CurrentUserID = 0 then
		Response.Write "You do not have access to this page."
	else

		rs.Open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
		if not (rs.EOF and rs.BOF) then
			strproductname = rs("Name") & ""
		else
			strproductname = ""
		end if
		rs.Close
	
%>



<font face=verdana size=3><b><%=strproductname%> Integration Test Commodities Received </b><BR></font>

<form id="frmMain" method="post" action="SICommoditiesSave.asp">
<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
<%
	rs.Open "spListTargetedDel4Product " & request("ID") & ",1,1"	,cn,adOpenForwardOnly

	if rs.EOF and rs.BOF then
		Response.Write "<BR><font size=2 face=verdana>No Version Targeted for this Product.</font>"
	else
%>
	<TR bgcolor=Wheat>
		<TD><b>Add</b></TD>
		<TD><b>Total</b></TD>
		<TD><b>History</b></TD>
		<TD><b>Vendor</b></TD>
		<TD><b>Deliverable</b></TD>
		<TD><b>HW Version</b></TD>
		<TD><b>FW Version</b></TD>
		<TD><b>Rev</b></TD>
		<TD><b>Part</b></TD>
		<TD><b>Model</b></TD>
	</TR>

<%
	dim strTotal
	dim strCount
	dim HistoryArray
		strCount = 0
		do while not rs.EOF
			if trim(rs("SICommodityHistory") & "") = "" then
				strTotal = "0"
			else
				HistoryArray = Split(rs("SICommodityHistory") & "",";")
				strTotal = HistoryArray(0)
			end if
			if trim(rs("SICommodityHistory") & "") = "" then
				strHistory = ""
			else
				strHistory =replace(replace(mid(rs("SiCommodityHistory"),len(HistoryArray(0))+2),";","<BR>"),"-","&#8209;")
			end if
			Response.Write "<TR>"
			Response.Write "<TD><INPUT style=""Width:30"" type=""text"" id=txtNew" & trim(rs("LinkID")) & " name=txtNew" & trim(rs("LinkID")) & " LANGUAGE=javascript onfocusout=""return txtNew_onfocusout(" & rs("LinkID") & ")""><INPUT type=""hidden"" id=txtTotalLoaded" & trim(rs("LinkID")) & " name=txtTotalLoaded" & trim(rs("LinkID")) & " value=""" & strTotal & """></TD>"
			Response.Write "<TD align=middle ID=lblTotal" & trim(rs("LinkID")) & ">" & strTotal & "</TD>"
			Response.Write "<TD>" & strHistory & "&nbsp;"
			Response.Write "<INPUT type=""hidden"" id=txtUpdates" & trim(rs("LinkID")) & " name=txtUpdates></TD>"
			Response.Write "<TD>" & rs("Vendor") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Name") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Version") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Revision") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("Pass") & "&nbsp;</TD>"
			Response.Write "<TD nowrap>" & rs("PartNumber") & "&nbsp;</TD>"
			Response.Write "<TD>" & rs("ModelNumber") & "&nbsp;</TD>"
			Response.Write "</TR>"
			
			strCount = strCount + 1
			rs.MoveNext	
		loop

	end if
	
	rs.Close
%>

</table>

<BR><font size=1 face=verdana>Commodities Displayed: <%=strCount%></font>

<INPUT style="Display:none" type="text" id=txtID name=txtID value="<%=request("ID")%>">
<INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
</form>
<%

	end if

	cn.Close
	set cn = nothing
	set rs = nothing

%>
</BODY>
</HTML>


