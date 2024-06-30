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
<script language="javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    
function lstAvailable_ondblclick() {
	if (lstAvailable.selectedIndex != "undefined")
		{
			if (lstAvailable.selectedIndex != -1 )
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[lstAvailable.selectedIndex].text,lstAvailable.options[lstAvailable.selectedIndex].value);
				lstAvailable.options[lstAvailable.options.selectedIndex]=null;
			}
		}

}

function lstNew_ondblclick() {
	if (lstNew.selectedIndex != "undefined")
		{
			if (lstNew.selectedIndex != -1 )
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[lstNew.selectedIndex].text,lstNew.options[lstNew.selectedIndex].value);
				lstNew.options[lstNew.options.selectedIndex]=null;
			}
		}
}

function cmdAdd_onclick() {
	var i;
	for (i=0;i<lstAvailable.length;i++)
		{
			if (lstAvailable.options[i].selected)
			{
				lstNew.options[lstNew.length] = new Option(lstAvailable.options[i].text,lstAvailable.options[i].value);
			}
		}

	for (i=lstAvailable.length-1;i>-1;i--)
		{
			if (lstAvailable.options[i].selected)
			{
				lstAvailable.options[i]=null;
			}
		}
		
		
}

function cmdRemove_onclick() {
	var i;
	for (i=0;i<lstNew.length;i++)
		{
			if (lstNew.options[i].selected)
			{
				lstAvailable.options[lstAvailable.length] = new Option(lstNew.options[i].text,lstNew.options[i].value);
			}
		}

	for (i=lstNew.length-1;i>-1;i--)
		{
			if (lstNew.options[i].selected)
			{
				lstNew.options[i]=null;
			}
		}
		


}

function sortSelect(mySelect) {
    var textArray=new Array()
    var valueArray=new Array()

    for (i=0;i<mySelect.options.length;i++) {
        textArray[i] = mySelect.options[i].text;
        valueArray[i] = mySelect.options[i].value;
    }

    textArray.sort()
    valueArray.sort()

    for (i=0;i<textArray.length;i++) {
        mySelect.options[i].text = textArray[i];
        mySelect.options[i].value = valueArray[i];
    }
}


function cmdOK_onclick() {
	var i;
	var strNewOrder="";
	
    sortSelect(lstAvailable);

	for(i=0;i<lstAvailable.length;i++)
		lstAvailable.options[i].selected=true;
	cmdAdd_onclick();
	
	for(i=0;i<lstNew.length;i++)
	{
		if (strNewOrder=="")
			strNewOrder = lstNew.options[i].value  
		else
			strNewOrder = strNewOrder + "," + lstNew.options[i].value  
	}
	frmUpdate.txtNewOrder.value = strNewOrder;
	frmUpdate.submit();
}

function cmdCancel_onclick() {
	window.returnValue = 0;
	window.parent.close();
}


function cmdImageButton_onmouseover() {
	window.event.srcElement.style.cursor = "default";
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
}

function cmdImageButton_onmouseout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
}

function cmdImageButton_onmousedown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";

}

function cmdImageButton_onmouseup() {
	window.event.srcElement.style.borderColor = "gold";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	ImageButton_Pressed(window.event.srcElement.name);
}

function cmdImageButton_onkeydown() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "inset";
	window.event.srcElement.style.backgroundColor = "LightGrey";
}

function cmdImageButton_onkeyup() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "solid";
	window.event.srcElement.style.backgroundColor = "gainsboro";
	if (window.event.keyCode !=9)
		ImageButton_Pressed(window.event.srcElement.name);
}

function ImageButton_Pressed(ID){
	if (ID=="cmdAdd")
		cmdAdd_onclick();
	else if (ID=="cmdRemove")
		cmdRemove_onclick();
	else if (ID=="cmdUp")
		MoveItemsUp();
	else if (ID=="cmdDown")
		MoveItemsDown();

}


function MoveItemsUp(){
    var strTempValue;
    var strTempText;
    for (i=1;i<lstNew.options.length;i++)
        if (lstNew.options[i].selected && !lstNew.options[i-1].selected)
            {
                strTempValue = lstNew.options[i-1].value;
                strTempText = lstNew.options[i-1].text;
                isTempSelected = lstNew.options[i-1].selected;
                lstNew.options[i-1].value=lstNew.options[i].value;
                lstNew.options[i-1].text=lstNew.options[i].text;
                lstNew.options[i-1].selected=true;
                lstNew.options[i].value=strTempValue;
                lstNew.options[i].text=strTempText;
                lstNew.options[i].selected = false; 
            }
}


function MoveItemsDown(){
    var strTempValue;
    var strTempText;
    for (i=lstNew.options.length-2;i>=0;i--)
        if (lstNew.options[i].selected && !lstNew.options[i+1].selected)
            {
                strTempValue = lstNew.options[i+1].value;
                strTempText = lstNew.options[i+1].text;
                isTempSelected = lstNew.options[i+1].selected;
                lstNew.options[i+1].value=lstNew.options[i].value;
                lstNew.options[i+1].text=lstNew.options[i].text;
                lstNew.options[i+1].selected=true;
                lstNew.options[i].value=strTempValue;
                lstNew.options[i].text=strTempText;
                lstNew.options[i].selected = false; 
            }
}

function cmdImageButton_onfocusout() {
	window.event.srcElement.style.borderColor = "";
	window.event.srcElement.style.borderStyle = "outset";
	window.event.srcElement.style.backgroundColor = "gainsboro";

}

function lstNew_onfocusin(){
    if( lstNew.options.length > 1)
        {
        cmdUp.style.display="";
        cmdDown.style.display="";
        }
}


//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=lavender>
<BR>
<table style="width:100%"><TR><TD width=10>&nbsp;</TD><TD>
<%

dim cn
dim rs
dim p
dim cm
dim ListArray
dim ListItem


set cn = server.CreateObject("ADODB.Connection")
cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
cn.Open

set rs = server.CreateObject("ADODB.recordset")

dim CurrentUSer
dim CurrentUSerID

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
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
	end if
	rs.Close



%>


<font size=3 face=verdana><b>Reorder Columns</b></font><BR><BR>
<!--<font size=2 face=verdana>This screen changes the order of the columns in the Summary Columns listbox. You must select the columns to include</font><BR><BR>-->
<table border=0 width=100%>
<tr>
<td nowrap width=50%>
	<font face=verdana size=2><b>Old Column Order:</b></font><BR>	<SELECT style="WIDTH:100%" size=20 id=lstAvailable name=lstAvailable  multiple LANGUAGE=javascript ondblclick="return lstAvailable_ondblclick()">
		<%
		ListArray=split(request("lstColumns"),",")
		for each ListItem in ListArray
			if trim(Listitem) <> "" then
				Response.Write "<option value=""" & trim(ListItem) & """>" & trim(ListItem) & "</option>"
			end if
		next
		%>	</SELECT>
</td>
	<td valign=top width=30 align=center><BR>
        <input type="image" src="../../images/arrowadd.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdAdd" name="cmdAdd" title="Add item" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="30" HEIGHT="23">
        <input type="image" src="../../images/arrowremove.gif" style="BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdRemove" name="cmdRemove" title="Remove item" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="30" HEIGHT="22">&nbsp; 
        <br>    
        <input type="image" src="../../images/arrowup.gif" style="display:none;BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdUp" name="cmdUp" title="Move item up in list" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="30" HEIGHT="22">
        <input type="image" src="../../images/arrowdown.gif" style="display:none;BORDER-RIGHT: thin outset; BORDER-TOP: thin outset; BORDER-LEFT: thin outset; BORDER-BOTTOM: thin outset; BACKGROUND-COLOR: gainsboro" onfocusout="return cmdImageButton_onfocusout()" id="cmdDown" name="cmdDown" title="Move item down in list" LANGUAGE="javascript" onmouseover="return cmdImageButton_onmouseover()" onmouseout="return cmdImageButton_onmouseout()" onmouseup="return cmdImageButton_onmouseup()" onmousedown="return cmdImageButton_onmousedown()" onkeydown="return cmdImageButton_onkeydown()" onkeyup="return cmdImageButton_onkeyup()" WIDTH="30" HEIGHT="23">&nbsp; 
    
	</td>
<td nowrap width=50%>
	<font face=verdana size=2><b>New Column Order:</b></font><BR>	<SELECT style="WIDTH:100%" size=20 id=lstNew name=lstNew multiple LANGUAGE=javascript onfocusin="lstNew_onfocusin();"  ondblclick="return lstNew_ondblclick()">
	</SELECT>
</td>
</tr>
</table>
<form id=frmUpdate action="ReorderColumnsSave.asp" method=post>

<table width=100%>
<TR><TD align=right>
	<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()">
	<INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()"></TD></TR>
</table>

<%
set rs = nothing
cn.Close
set cn = nothing
%>
<INPUT type="hidden" id=txtEmployeeID name=txtEmployeeID value="<%=CurrentUserID%>">
<INPUT type="hidden" id=txtReportType name=txtUserSettingsID value="<%=request("UserSettingsID")%>">
<INPUT type="hidden" id=txtNewOrder name=txtNewOrder value="">
</form>
</td>
</tr>
</TABLE>

</BODY>
</HTML>
