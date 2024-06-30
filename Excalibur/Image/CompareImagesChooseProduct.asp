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
<!-- #include file = "../_ScriptLibrary/sort.js" -->
<!--



function optImageChoose_onclick() {
//	ChooseImages.style.display="";
//	frmMain.ImageDefinitionID.value = frmMain.tagImageDefinitionID.value;

}

function optImageAll_onclick() {
//	ChooseImages.style.display="none";
	ClearImages();
//	frmMain.ImageDefinitionID.value = frmMain.tagImageDefinitionID.value;
}

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

function preSortTable(chkBox, strTable, col, dataType, depth){
	var i;
	var strChecked="";
	
	if (typeof(chkBox.length) == "undefined")
		{
		if (chkBox.checked)
			strChecked = strChecked + "," + chkBox.value;
		}
	else
		{
		for (i=0;i<chkBox.length;i++)
			{
			if (chkBox(i).checked)
				strChecked = strChecked + "," + chkBox(i).value;
			}
		}
	if (strChecked.length > 0)
		strChecked = strChecked + ","
	
	SortTable( strTable, col ,dataType,depth);


	if (typeof(chkBox.length) == "undefined")
		{
		if (strChecked.indexOf("," + chkBox.value + ",",0)> -1)
			//chkBox.checked=true;
			frmMain.lstImageDefinitions.checked=true;

		}
	else
		{
		for (i=0;i<chkBox.length;i++)
			{
			if (strChecked.indexOf("," + chkBox(i).value + ",",0)> -1)
				chkBox(i).checked=true;
			}
		}
}



function onclick_CheckAll(chkBox,chkBoxAll) {
	frmMain.optImageChoose.checked=true;
	if (typeof(chkBox.length) == "undefined")
		chkBox.checked = chkBoxAll.checked;
	else
		{
		for (i=0;i<chkBox.length;i++)
			{
			chkBox(i).checked = chkBoxAll.checked;
			}
		}
}

function onclick_lstImageDefinitions() {
	frmMain.optImageChoose.checked=true;
}

function cmdOK_onclick() {
	var blnFound=true;
	
	if (!frmMain.optImageChoose.checked)
		ClearImages();
	else
		{
		blnFound = false;
		if (typeof(frmMain.lstImageDefinitions.length)=="undefined");
			if (frmMain.lstImageDefinitions.checked)
				blnFound=true;
		else
			for (i=0;i<frmMain.lstImageDefinitions.length;i++)
				if (frmMain.lstImageDefinitions(i).checked)
					blnFound=true;
		}
	
	if (!blnFound)
		alert("You must select at least one image to compare.");
	else
		frmMain.submit();
}

function ClearImages(){
	frmMain.chkAllImages.checked = false;
	if (typeof(frmMain.lstImageDefinitions.length)=="undefined")
		frmMain.lstImageDefinitions.checked = false;
	else
		for (i=0;i<frmMain.lstImageDefinitions.length;i++)
			frmMain.lstImageDefinitions(i).checked = false;
}

function SelectOS(ID){
	frmMain.optImageChoose.checked = true;
	if (typeof(frmMain.lstImageDefinitions.length)=="undefined")
		{
		if (frmMain.lstImageDefinitions.OSID==ID)
			frmMain.lstImageDefinitions.checked=true;
		}
	else
		{
		for (i=0;i<frmMain.lstImageDefinitions.length;i++)
			if (frmMain.lstImageDefinitions(i).OSID==ID)
				frmMain.lstImageDefinitions(i).checked = true;
		}
}

function cmdCancel_onclick() {
	window.close();
}

function ChangeServer(){
    if (frmMain.chkServer.checked)
        frmMain.action="CompareImageTestServer.asp";
    else
        frmMain.action="CompareImage.asp";

}
//-->
</SCRIPT>
</HEAD>
<BODY bgColor=Ivory>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font face=verdana size=2>
<%if trim(request("ProdID")) = "" then%>
	Not enough information supplied to display this page.
<%else%>

<%
	dim cn, rs
	dim strProductName
	dim strImages

	strProductName = ""	

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.Recordset")

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
	
	if rs.EOF and rs.BOF then
		Response.Write "Unable to find the product you specified."
	else
		strProductName = rs("Name")
	end if
	rs.Close
	
	

	if strProductName = "" then
		Response.Write "Not enough information provided to display this page."
	else
%>


	<form ID=frmMain action="CompareImage.asp" method=post>
	<Font size=3>
	<%if request("PINTest")="1" then%>
		<b>Validate <%=strProductName%>&nbsp;PIN&nbsp;Test&nbsp;Images</b>
	<%else%>
		<b>Validate <%=strProductName%>&nbsp;Images</b>
	<%end if%>
	<BR><BR></font>
	<font size=2 face=verdana><b>Choose Images to Validate:</b><BR></font>
	<INPUT checked type="radio" id=optImageAll name=optImages LANGUAGE=javascript onclick="return optImageAll_onclick()">All <%=strProductName%> images in development. <BR>
	<!--<INPUT type="radio" id=optImageOS name=optImages LANGUAGE=javascript onclick="return optImageProd_onclick()">Choose images by operating system.-->
	
	<%
	dim OSLinks
	OSLinks=""
	rs.Open "spListImageDefinitionsOSByProduct " & cint(request("ProdID")) & ",3",cn,adOpenStatic
	'Response.Write "<table cellpadding=0 cellspacing=0 bgcolor=gainsboro border =0 width=""100%""><tr><TD bgcolor=ivory width=10>&nbsp;</td>"
'	cellCount=0
	do while not rs.eof
		if rs("ID") <> "21" and rs("ID") <> "22" and rs("ID") <> "23" then
'			cellcount = cellcount + 1
'			if cellcount =7 then
'				cellcount =1
'				Response.write "<TD width=""100%"">&nbsp;</td></tr><tr><TD width=10>&nbsp;</td>"
'			end if
'			Response.Write "<TD nowrap><INPUT checked type=""checkbox"" id=""checkbox1"" name=""checkbox1"">" & rs("OS") & "&nbsp;&nbsp;</TD>"
			OSLinks = OSLinks & " | <a href=""javascript: SelectOS(" & rs("ID") & ");"">" & rs("OS") & "</a>"
		end if
		rs.MoveNext
	loop
	rs.Close
	if OSLinks <> "" then
		OSLinks = mid(OSLinks,4)
	end if
'	Response.Write "<TD width=""100%"" colspan=" & 7-i & ">&nbsp;</td></tr></table>"
	%>
	<INPUT type="radio" id=optImageChoose name=optImages LANGUAGE=javascript onclick="return optImageChoose_onclick()">Choose individual image definitions. 
	<table ID=ChooseImages width=100% style="Display:"><tr><TD width=20>&nbsp;</TD><TD>
	<font size=2 face=verdana><b>Select:&nbsp;</b><%=OSLinks%></font></td></tr>
	<tr><TD>&nbsp;</TD><TD>
		<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 250px; BACKGROUND-COLOR: white" id=DIV1>
			<TABLE ID=TableImage width=100%>
				<THEAD bgcolor=LightSteelBlue  >
					<TD style="Width:1;BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT type="checkbox" id=chkAllImages name=chkAllImages LANGUAGE=javascript onclick="return onclick_CheckAll(lstImageDefinitions,chkAllImages)">&nbsp;</TD>
					<TD  style="Width=60" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;SKU&nbsp;Number&nbsp;</TD>
					<TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 2 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Brand&nbsp;</TD>
					<TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 3 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Software&nbsp;</TD>
					<TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 4 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;OS&nbsp;</TD>
					<TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 5 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Type&nbsp;</TD>
					<TD style="Width=100%" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImageDefinitions, 'TableImage', 6 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Comments&nbsp;</TD></THEAD>
<%
	rs.Open "spListImageDefinitionsByProduct " & cint(request("ProdID")) & ",3",cn,adOpenStatic
	do while not rs.EOF	
		if rs("OSID") <> "21" and rs("OSID") <> "22" and rs("OSID") <> "23" then
	
%>

			<TR>
				<TD><INPUT class="check" type="checkbox" id=lstImageDefinitions name=lstImageDefinitions value="<%=rs("ID")%>" OSID="<%=rs("OSID")%>" LANGUAGE=javascript onclick="return onclick_lstImageDefinitions();">&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("SkuNumber") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("Brand") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("OS") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("SWType") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("ImageType") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("Comments") & ""%>&nbsp;</TD>
			</TR>
<%		end if
		rs.MoveNext
	loop
	rs.Close
%>

	</TABLE>
            
</div>
</TD></TR></table>
	
	<font size=1 face=verdana color=red>Note:  The validate function only supports images in development and does not include RedFlag Linux images, SuSE Linux images, or FreeDOS images.</font>
	<HR>
	<table width=100%><TR><TD width=100% align=right>
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
	</td></tr></table>
	
	<INPUT type="hidden" id=ProdID name=ProdID value="<%=request("ProdID")%>">
	<INPUT style="display:none" type="text" id=PINTest name=PINTest value="<%=request("PINTest")%>">
	    <input style="display:none" id="chkServer" onclick="javascript:ChangeServer();" type="checkbox" /> Use Test Server

	
	</form>
	
	
<%

	end if

	set rs = nothing
	cn.close
	set cn = nothing

%>	
	
	
<%end if%>

</font>
</BODY>
</HTML>
