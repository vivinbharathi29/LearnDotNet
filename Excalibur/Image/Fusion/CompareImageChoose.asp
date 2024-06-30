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
<!-- #include file = "../../_ScriptLibrary/sort.js" -->
<!--



function optImageChoose_onclick() {
//	ChooseImages.style.display="";
	frmMain.ImageDefinitionID.value = frmMain.tagImageDefinitionID.value;

}

function optImageProd_onclick() {
///	ChooseImages.style.display="none";
	ClearImages();
	frmMain.ImageDefinitionID.value = ""
}

function optImageDef_onclick() {
//	ChooseImages.style.display="none";
	ClearImages();
	frmMain.ImageDefinitionID.value = frmMain.tagImageDefinitionID.value;

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

/*function onclick_Row(EmployeeID, CanEdit, CanDelete){
	if (event.srcElement.className != "check")
		{
		var strResult;
		strResult = window.showModalDialog("ProfileShareProperties.asp?ProfileID=" + txtProfileID.value + "&EmployeeID=" + EmployeeID + "&CanEdit=" + CanEdit + "&CanDelete=" + CanDelete,"","dialogWidth:450px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No"); 
		if (typeof(strResult) != "undefined")
			window.location.href =  "ProfileShare.asp?ID=" + txtProfileID.value;
		}
}
*/
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



function onclick_CheckAll(chkBox,chkBoxAll) {
	frmMain.optImageChoose.checked=true;
	for (i=0;i<chkBox.length;i++)
		{
		chkBox(i).checked = chkBoxAll.checked;
		}
}

function onclick_lstImage() {
	frmMain.optImageChoose.checked=true;
}

function cmdOK_onclick() {
	var blnFound=true;
	
	if (!frmMain.optImageChoose.checked)
		ClearImages();
	else
		{
		blnFound = false;
		if (typeof(lstImage)=="undefined");
			if (frmMain.lstImage.checked)
				blnFound=true;
		else
			for (i=0;i<frmMain.lstImage.length;i++)
				if (frmMain.lstImage(i).checked)
					blnFound=true;
		}
	
	if (!blnFound)
		alert("You must select at least one image to compare.");
	else
		frmMain.submit();
}

function ClearImages(){
	frmMain.chkAllImages.checked = false;
	for (i=0;i<frmMain.lstImage.length;i++)
		frmMain.lstImage(i).checked = false;
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
<link href="../../style/wizard%20style.css" type="text/css" rel="stylesheet">

<font face=verdana size=2>
<%if trim(request("ImageDefinitionID")) = "" or trim(request("ProdID")) = "" then%>
	Not enough information supplied to display this page.
<%else%>

<%
	dim cn, rs
	dim strProductName
	dim strDefinitionName
	dim strImages

	strProductName = ""	
	strDefinitionName = ""	

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
	
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spGetImageDefinition"
			
		
	Set p = cm.CreateParameter("@DefinitionID", 3, &H0001)
	p.Value = request("ImageDefinitionID")
	cm.Parameters.Append p
		
		
	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing	
	
	if rs.EOF and rs.BOF then
		Response.Write "Unable to find the product you specified."
	else
		if trim(rs("SKUNumber") & "") = "" then
			strDefinitionName = "Image Definition " & rs("ID")
		else
			strDefinitionName = rs("SKUNumber")
		end if
	end if
	rs.Close
	
	
	
	
	

	if strProductName = "" or strDefinitionName = "" then
		Response.Write "Not enough information provided to display this page."
	else
%>


	<form ID=frmMain action="CompareImage.asp" method=post>
	<Font size=3><b>Validate <%=strProductName%>&nbsp;(<%=strDefinitionName%>) Images</b><BR><BR></font>
	<font size=2 face=verdana><b>Choose Images to Validate:</b><BR></font>
	<INPUT checked type="radio" id=optImageDef name=optImages LANGUAGE=javascript onclick="return optImageDef_onclick()">All <%=strDefinitionName%> images. <BR>
	<INPUT type="radio" id=optImageProd name=optImages LANGUAGE=javascript onclick="return optImageProd_onclick()">All <%=strProductName%> images. <BR>
	<INPUT type="radio" id=optImageChoose name=optImages LANGUAGE=javascript onclick="return optImageChoose_onclick()">Choose individual <%=strDefinitionName%> images. <BR>
	<table ID=ChooseImages width=100% style="Display:"><tr><TD>&nbsp;</TD><TD>
		<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 250px; BACKGROUND-COLOR: white" id=DIV1>
			<TABLE ID=TableImage width=100%>
				<THEAD bgcolor=LightSteelBlue  >
					<TD style="Width:1;BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><INPUT type="checkbox" id=chkAllImages name=chkAllImages LANGUAGE=javascript onclick="return onclick_CheckAll(lstImage,chkAllImages)">&nbsp;</TD>
					<TD  style="Width=60" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImage, 'TableImage', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;SKU&nbsp;Number&nbsp;</TD>
					<TD style="Width=100" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImage, 'TableImage', 2 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Config&nbsp;</TD>
					<TD style="Width=100%" onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstImage, 'TableImage', 3 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Name&nbsp;</TD></THEAD>
<%
	rs.Open "spListImagesForDefinition " & cint(request("ImageDefinitionID")),cn,adOpenStatic
	do while not rs.EOF	
	
		if isnumeric(rs("Priority")) then
%>

			<TR>
				<TD><INPUT class="check" type="checkbox" id=lstImage name=lstImage value="<%=rs("ImageID")%>" LANGUAGE=javascript onclick="return onclick_lstImage();">&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("FullSkuNumber") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("OptionConfig") & ""%>&nbsp;</TD>
				<TD nowrap>&nbsp;<%=rs("DisplayName") & ""%>&nbsp;</TD>
			</TR>
<%
		end if
		rs.MoveNext
	loop
	rs.Close
%>

	</TABLE>
            
</div>
</TD></TR></table>
	
	
	<HR>
	<table width=100%><TR><TD width=100% align=right>
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">
	</td></tr></table>
	<INPUT type="hidden" id=ProdID name=ProdID value="<%=request("ProdID")%>">
	<INPUT type="hidden" id=ImageDefinitionID name=ImageDefinitionID value="<%=request("ImageDefinitionID")%>">
	<INPUT type="hidden" id=tagImageDefinitionID name=tagImageDefinitionID value="<%=request("ImageDefinitionID")%>">
	<INPUT style="display:none" type="text" id=PINTest name=PINTest value="0">
    <input style="display:none" id="chkServer" onclick="javascript:ChangeServer();" type="checkbox" /> <!--Use Test Server-->
	
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
