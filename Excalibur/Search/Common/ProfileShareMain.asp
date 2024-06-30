<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../../_ScriptLibrary/sort.js" -->

function mouseover_Column(){
	event.srcElement.style.color="red";
	event.srcElement.style.cursor="hand";
	
}
function mouseout_Column(){
	event.srcElement.style.color="black";
}

function mouseover_Cell(){
	event.srcElement.parentElement.style.color="red";
	event.srcElement.parentElement.style.cursor="hand";
}

function mouseout_Cell(){
	event.srcElement.parentElement.style.color="black";
}

function onclick_Row(EmployeeID, CanEdit, CanDelete,AddType){
	var strResult;
	strResult = window.showModalDialog("ProfileShareProperties.asp?AddType=" + AddType + "&ProfileID=" + txtProfileID.value + "&EmployeeID=" + EmployeeID + "&CanEdit=" + CanEdit + "&CanDelete=" + CanDelete,"","dialogWidth:450px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
		window.location.href="ProfileShare.asp?ID=" + txtProfileID.value;
}

function cmdAdd_onclick(){
	var strResult;
	strResult = window.showModalDialog("ProfileShareProperties.asp?ProfileID=" + txtProfileID.value + "&EmployeeID=0&CanEdit=&CanDelete=","","dialogWidth:450px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
		window.location.href = "ProfileShare.asp?ID=" + txtProfileID.value;
}

function cmdAddGroup_onclick(){
	var strResult;
	strResult = window.showModalDialog("ProfileShareProperties.asp?AddType=2&ProfileID=" + txtProfileID.value + "&EmployeeID=0&CanEdit=&CanDelete=","","dialogWidth:540px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
	if (typeof(strResult) != "undefined")
		window.location.href = "ProfileShare.asp?ID=" + txtProfileID.value;
}


function cmdDone_onclick(){
    top.close();
}



function preSortTable(chkBox, strTable, col, dataType, depth){
	var i;
	var strChecked="";
	
	for (i=0;i<chkBox.length;i++)
		{
		if (chkBox(i).checked)
			strChecked = strChecked + "," + chkBox(i).value;
		}
	if (strChecked.length > 0)
		strChecked = strChecked + ","
	
	
	SortTable( strTable, col ,dataType,depth);


	for (i=0;i<chkBox.length;i++)
		{
		if (strChecked.indexOf("," + chkBox(i).value + ",",0)> -1)
			chkBox(i).checked=true;
		}
}



//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=ivory>
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<%
	dim cn, rs
	dim strProfileName
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") '"Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	rs.Open "spGetReportProfile " & clng(Request("ID")),cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strProfileName=""
	else 
		strProfileName= rs("ProfileName") & ""
	end if
	rs.Close
	
%>


<!--<%=Request("ID")%>-->

<font size=3 face=verdana><b>Share <%=strProfileName%> Profile<BR></font>
<font size=2 face=verdana><b><BR>Shared With:</font>


<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 250px; BACKGROUND-COLOR: white" id=DIV1>
	<TABLE ID=TableShare width=100%>
	<THEAD bgcolor=LightSteelBlue  ><TD  onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstShare, 'TableShare', 0 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Name&nbsp;</TD><TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstShare, 'TableShare', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Can&nbsp;Edit&nbsp;</TD><TD  style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstShare, 'TableShare', 2 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Can&nbsp;Delete&nbsp;</TD></THEAD>
<%
	rs.Open "spGetReportProfileShared " & clng(Request("ID")),cn,adOpenStatic
	do while not rs.EOF	
%>

			<TR onclick="onclick_Row(<%=rs("SharedEmployeeID")%>,'<%=rs("CanEdit")%>','<%=rs("CanDelete")%>',1);" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();">
				<TD nowrap><INPUT style="Display:none;" type="checkbox" id=lstShare name=lstShare>&nbsp;<%=rs("SharedEmployeeName") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=replace(replace(rs("CanEdit")& "","True","Yes"),"False","No") %>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=replace(replace(rs("CanDelete") & "","True","Yes"),"False","No")%>&nbsp;</TD>
			</TR>
<%
		rs.MoveNext 
	loop
	rs.Close

	rs.Open "spGetReportProfileGroups " & clng(Request("ID")),cn,adOpenStatic
	do while not rs.EOF	
%>

			<TR onclick="onclick_Row(<%=rs("ID")%>,'<%=rs("CanEdit")%>','<%=rs("CanDelete")%>',2);" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();">
				<TD nowrap><INPUT style="Display:none;" type="checkbox" id=lstShare name=lstShare>&nbsp;<b><%=left(rs("Setting") & "",instr(rs("Setting") & "","|")-1)%>&nbsp;</b></TD>
				<TD nowrap>&nbsp;<%=replace(replace(rs("CanEdit")& "","True","Yes"),"False","No") %>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=replace(replace(rs("CanDelete") & "","True","Yes"),"False","No")%>&nbsp;</TD>
			</TR>
<%
		rs.MoveNext 
	loop
	rs.Close
%>

	</TABLE>
            
</div>
<table width=100%>
	<TR>
		<TD align=right>
			<INPUT type="button" value="Add Group" id=cmdAddGroup name=cmdAddGroup LANGUAGE=javascript onclick="return cmdAddGroup_onclick()">
			<INPUT type="button" value="Add Person" id=cmdAdd name=cmdAdd LANGUAGE=javascript onclick="return cmdAdd_onclick()">
			<INPUT type="button" value="Done" id=cmdDone name=cmdDone LANGUAGE=javascript onclick="return cmdDone_onclick()">
		</TD>
	</TR>
</table>
<%
	set rs = nothing
	cn.Close
	set cn = nothing

%>
<INPUT type="hidden" id=txtProfileID name=txtProfileID value="<%=request("ID")%>">

</BODY>
</HTML>
