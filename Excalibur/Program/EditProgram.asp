<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
<!-- #include file = "../_ScriptLibrary/sort.js" -->

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
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

function onclick_Cell(chkBox, strID){
	var i;
	if(event.srcElement.name==chkBox(0).name)
		return;
	for (i=0;i<chkBox.length;i++)
		{
			if (chkBox(i).value==strID)
				{
				if (chkBox(i).checked)
					{
					chkBox(i).checked = false;
					}
				else
					{
					chkBox(i).checked = true;
					}
				}
		}
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
function onclick_CheckAll(chkBox,chkBoxAll) {
	for (i=0;i<chkBox.length;i++)
		{
		chkBox(i).checked = chkBoxAll.checked;
		}
}


function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


function chkActive_onclick() {
	if (! frmUpdate.chkActive.checked)
		frmUpdate.chkDefault.checked =false;
}

function activeValid(chkActive) {
	if (!frmUpdate.chkActive.checked)
	{
		var Products = frmUpdate.cboLead.options.length;
		if(Products > 1)
		{
			document.getElementById('inactiveRule').style.display='inline-block';
			chkActive.checked = true;
		}
		else
		{
			document.getElementById('inactiveRule').style.display='none';
		}
	}

}

function window_onload() {
	frmUpdate.txtName.focus();
}

function cboProgramGroup_onchange() {
	if (frmUpdate.cboProgramGroup.value!="0" && frmUpdate.cboProgramGroup.value!="" && frmUpdate.cboProgramGroup.selectedIndex!=-1)
		{
		frmUpdate.txtFullName.value = frmUpdate.cboGroupAbbreviations.options[frmUpdate.cboProgramGroup.selectedIndex-1].value + " " + frmUpdate.txtName.value;
		spnFullName.innerText = frmUpdate.txtFullName.value;
		}
	else
		{
		frmUpdate.txtFullName.value = "";
		spnFullName.innerText = "";
		}
}

function txtName_onkeyup() {
    cboProgramGroup_onchange();
}

function ProductChecked(ID){
    var i;
    
    if(window.document.all("lstProduct" + ID).checked)
	{
		frmUpdate.cboLead.options[frmUpdate.cboLead.length] = new Option(window.document.all("lstProduct" + ID).ProductName,window.document.all("lstProduct" + ID).value);
		frmUpdate.chkActive.checked = true;
	}
    else
		for (i=0;i<frmUpdate.cboLead.options.length;i++)
		    if (frmUpdate.cboLead.options[i].value==ID)
		        frmUpdate.cboLead.options[i] = null;
    
}
//-->
</SCRIPT>

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


</STYLE>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
</HEAD>
<BODY bgcolor="ivory"  LANGUAGE=javascript onload="return window_onload()">


<%

	dim cn
	dim cm
	dim p
	dim rs
	dim i
	dim CurrentUser
	dim CurrentUserID
	dim strName
	dim strActive
	dim strCurrent
	dim strProducts
	dim strGroupID
	dim strOTSCycleName
	dim strFullName
	dim strLeadProductID
    dim strProgramGroups
    dim strProgramGroupAbbreviations
	
	strLeadProductID = 0

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

	set cm=nothing
	
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID") & ""
	else
		CurrentUserID = 0
	end if
	rs.Close
	
	dim blnFound

	if request("ID") <> "" then
		rs.Open "spGetProgramProperties " & clng(request("ID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			Response.Write "<br><font size=2 face=verdana>Program not found.</font>"
			blnFound = false
		else
			strName = rs("Name") & ""
			strCurrent = rs("CurrentProgram") & ""
			strActive = rs("Active") & ""
			strCurrent = replace(replace (strCurrent ,"True","checked") ,"False","")
			strActive = replace(replace (strActive ,"1","checked") ,"0","")
			strGroupID = trim(rs("ProgramGroupID") & "")
			strOTSCycleName = rs("OTSCycleName") & ""
            strFullName = rs("FullName") & ""
			rs.Close
			strProducts = ""
			rs.Open "spGetProductsInProgram " & request("ID"),cn,adOpenForwardOnly
			do while not rs.EOF
				strproducts = strproducts & "," & rs("ID")
				if rs("CommonBucketLink") then
				    strLeadProductID = rs("ID")
				end if
				rs.MoveNext
			loop
	        
			blnFound = true

		end if
		rs.Close
	else
		strName = ""
		strCurrent =  ""
		strActive =  "checked"
		strProducts=""
	end if


    strProgramGroups=""
    strProgramGroupAbbreviations=""
    rs.open "spListProgramGroups null",cn
    do while not rs.eof
        if trim(strGroupID) = trim(rs("ID")) then
            strProgramGroups = strProgramGroups & "<option selected value=""" & rs("ID") &""">" & rs("Name") & "</option>"
            strProgramGroupAbbreviations = strProgramGroupAbbreviations & "<option value=""" & rs("Abbreviation") & """></option>"
        elseif rs("Active") then
            strProgramGroups = strProgramGroups & "<option value=""" & rs("ID") &""">" & rs("Name") & "</option>"
            strProgramGroupAbbreviations = strProgramGroupAbbreviations & "<option value=""" & rs("Abbreviation") & """></option>"
        end if
        rs.movenext
    loop
    rs.close


	if blnFound or request("ID") = "" then
	
%>



<font face=verdana size=><b>
<label ID="lblTitle">
<%if request("ID") <> "" then%>
	Update Product Group
<%else%>
	Add Product Group
<%end if%>
</label></b></font>





<form id="frmUpdate" method="post" action="EditProgramSave.asp?pulsarplus=<%=Request("pulsarplus")%>">

<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<td valign=top width=120 nowrap><b>Type:&nbsp;<font size=1 face=verdana color=red>*</font></b></td>
		<TD>
			<SELECT  style="width:100%" id=cboProgramGroup name=cboProgramGroup LANGUAGE=javascript onchange="return cboProgramGroup_onchange()">
				<OPTION selected value="0"></OPTION>
                <%=strProgramGroups%>
			</SELECT>
			<SELECT style="display:none" id=cboGroupAbbreviations>
                <%=strProgramGroupAbbreviations%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Short&nbsp;Name:</b> <font size=1 face=verdana color=red><b>*</b></font></td>
		<TD><INPUT style="width:150" type="text" id=txtName name=txtName maxlength=30 value="<%=strName%>" LANGUAGE=javascript onkeyup="return txtName_onkeyup()" onfocusout="return txtName_onkeyup()">
        <font size=1 face=verdana color=green> - Examples:  2012, 3C12, 2012 Slate,...</font>
        </TD>
	</TR>

	<TR>
		<td valign=top width=120 nowrap><b>Full&nbsp;Name:</b></td>
		<TD><span id=spnFullName><%=strFullName%></span>
		<INPUT type="hidden" id=txtFullName name=txtFullName value="<%=strFullName%>">
		</TD>
	</TR>

	<TR style=display:none>
		<td valign=top width=120 nowrap><b>OTS&nbsp;Cycle&nbsp;Name:</b></td>
		<TD><INPUT type="hidden" id=txtOTSCycleName name=txtOTSCycleName value="<%=strOTSCycleName%>">
		</TD>
	</TR>


  <tr>
    <td width="160" style="VERTICAL-ALIGN: top"><strong><font size=2>Products:</font></strong></td>
	<TD>
			<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 200px; BACKGROUND-COLOR: white" id=DIV1>
	<TABLE ID=TableProduct width=100%>
	<THEAD bgcolor=LightSteelBlue  ><tr><TD nowrap style="width:20px;BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset"><!--<INPUT id=chkProductAll type=checkbox name=chkProductAll LANGUAGE=javascript onclick="return onclick_CheckAll(lstProduct,chkProductAll)">-->&nbsp;</TD><TD  onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="preSortTable( lstProduct, 'TableProduct', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Product&nbsp;Version&nbsp;</TD></tr></THEAD>
		<% 
			rs.Open "spGetProductsAll",cn,adOpenForwardOnly
			'rs.Open "spGetProducts",cn,adOpenForwardOnly
			dim strRest
			dim strSelectedProducts
			strRest=""
			do while not rs.eof
			%>
				<%if instr(strproducts & ",","," & trim(rs("ID")) & ",")>0 then%>
				<TR onclick="onclick_Cell(lstProduct,'<%=rs("Product")%>');" onmouseover="mouseover_Cell();" onmouseout="mouseout_Cell();">
				<TD nowrap>
				<INPUT id=lstProduct<%=rs("ID")%> type=checkbox name=lstProduct ProductName="<%=rs("Product")%>" checked value="<%=rs("ID")%>" onclick="ProductChecked(<%=rs("ID") %>);"></TD>			
				<TD nowrap><%=rs("Product")%></TD>
				</TR>
				<%
				    if trim(strLeadProductID) = trim(rs("ID")) then
				        strSelectedProducts = strSelectedProducts & "<option selected value=""" & rs("ID") & """>" & rs("Product") & "</option>"
				    else
				        strSelectedProducts = strSelectedProducts & "<option value=""" & rs("ID") & """>" & rs("Product") & "</option>"
				    end if
				else
					strRest = strRest & "<TR onclick=""onclick_Cell(lstProduct,'" & rs("ID") & "');"" onmouseover=""mouseover_Cell();"" onmouseout=""mouseout_Cell();""><TD nowrap><INPUT id=lstProduct" & rs("ID") & " type=checkbox name=lstProduct onclick=""ProductChecked(" & rs("ID") & ");"" ProductName=""" & rs("Product") & """ value=""" & rs("ID") & """></TD><TD nowrap>" &  rs("Product")  & "</TD></TR>"			
				end if%>
				
			<%
				rs.MoveNext
			loop
			rs.Close
			if strRest <> "" then
				Response.Write strRest
			end if
			%>
	</TABLE>
            
</div>
		
		
		</TD>
	</TR>
	<TR>
		<td valign=top width=120 nowrap><b>Lead Product:</b></td>
		<TD>
            <select id="cboLead" name="cboLead" style="width:150px">
                <option selected value="0"></option>
                <%=strSelectedproducts%>
            </select>
		</TD>
	</TR>
	<TR style="display:none">
		<td valign=top width=120 nowrap><b>Status:</b></td>
		<TD><INPUT type="checkbox" <%=strCurrent%> id=chkCurrent name=chkCurrent> This is the Current Cycle</TD>
	</TR>


	<TR>
		<td valign=top width=120 nowrap><b>State:</b></td>
		<TD><INPUT type=checkbox <%=strActive%> id=chkActive name=chkActive  onclick="activeValid(chkActive);" > Active
		    <font id=inactiveRule size=1 face=verdana color=green style=display:none> *Please empty Product Group before set it to be inactive.</font>
		</TD>
	</TR>

</table>


<INPUT type="hidden" id=txtDisplayedID name=txtDisplayedID value="<%=request("ID")%>">
<INPUT type="hidden" id=txtCurrentUserID name=txtCurrentUserID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=tagProduct name=tagProduct value="<%=strProducts%>">
<INPUT type="hidden" id=tagLead name=tagLead value="<%=trim(strLeadProductID)%>">

</form>
<%
	end if
	cn.Close
	set cn = nothing
	set rs = nothing


%>


</BODY>
</HTML>


