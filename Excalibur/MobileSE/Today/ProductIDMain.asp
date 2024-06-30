<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script type="text/javascript" src="../../includes/client/jquery.min.js"></script>
<script type="text/javascript" src="../../includes/client/json2.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdOK_onclick() {

	var i;
	var IDList="";
	var CommentsList="";
	
	if (ValidValues())
		{
		for (i=0;i<30;i++)
			{
			if (txtID[i].value!="")
				{
				IDList = IDList + "," + txtID[i].value.toUpperCase();
				CommentsList = CommentsList + "|" + txtID[i].value.toUpperCase() + "^" + txtComments[i].value + "^" + txtCHID[i].value; 
				}
			}
		if (IDList.length > 0)
			IDList = IDList.substring(1)
		if (CommentsList.length > 0)
			CommentsList = CommentsList.substring(1)

		var OutArray = new Array();
		OutArray[0]= IDList;
		OutArray[1] = CommentsList;

		if (parent.window.parent.document.getElementById('modal_dialog')) {
		    //save array value and return to parent page: ---
		    parent.window.parent.modalDialog.passArgument(JSON.stringify(OutArray), 'product_id_array');
		    parent.window.parent.GetProductIDResult();
		    parent.window.parent.modalDialog.cancel();
		} else {
		    window.returnValue = OutArray;
		    window.close();
		}
	}
}

function ValidValues(){
	var i;

	for (i=0;i<15;i++)
		{
		
		if (txtID[i].value.length != 0 && txtID[i].value.length != 4)
			{
			alert("ID must be 4 characters long"); 
			txtID[i].focus();
			return false;
			}
		else if (txtID[i].value != "" && (! isHex(txtID[i].value) ) )
			{
			alert("ID must be a Hex number."); 
			txtID[i].focus();
			return false;
			}
		else if (txtID[i].value == "" && txtComments[i].value != "")
			{
			alert("You must enter and ID number for each comment you enter."); 
			txtID[i].focus();
			return false;
			}
        else if (txtID[i].value == "" && txtCHID[i].value != "")
			{
			alert("You must enter and ID number for each CHID you enter."); 
			txtID[i].focus();
			return false;
			}

		}
	
		return true;
}




function window_onload() {
	txtID[0].focus();
}


function isHex(strValue){
	var strGood="0123456789ABCDEF";
	var strLength= strValue.length;
	
	if(strLength == 0)
		return false;

	strValue=strValue.toUpperCase();
	
	for (i=0;i<strLength;i++)
		if(strGood.indexOf(strValue.charAt(i))<0)
			return false;

	return true;
}

//-->
</SCRIPT>
<STYLE>
td{
	FONT-FAMILY=Verdana;
	FONT-SIZE=x-small;
}
</STYLE>
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana><b>
<% if request("TypeID") = "1" then%>
System Board ID List
<%else%>
Machine PNP ID List
<%end if%>

</b></font><BR><BR>
<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<TR>
	<TD><b>ID</b>&nbsp;<font size=2 color=green>(4&nbsp;digit&nbsp;hex)</font>&nbsp;</TD>
    <TD><b>Comments</b>&nbsp;<font size=2 color=green>(optional)</font>&nbsp;&nbsp;</TD>
	<TD><b>CHID</b>&nbsp;<font size=2 color=green>(optional)</font>&nbsp;&nbsp;</TD>
</TR>
<%
	dim IDArray
	dim IDRow
	dim CellArray
	redim OutputID(30)
	redim OutputComments(30)
    redim OutputCHID(30)
	dim Row

	for i = 0 to 29
		OutputID(i) = ""
		OutputComments(i) = ""
        OutputCHID(i) = ""
	next
	
	IDArray = split(request("IDList"),"|")
	
	Row=0
	for each IDRow in IDArray
		CellArray = split(IDRow,"^")
		OutputID(Row) = CellArray(0) 
		if ubound(CellArray) = 1 then
			OutputComments(Row) = CellArray(1)
		elseif ubound(CellArray) = 2 then
            OutputComments(Row) = CellArray(1)
            OutputCHID(Row) = CellArray(2)
        else
			OutputComments(Row) = ""
            OutputCHID(Row) = ""
		end if
		Row=Row+1
	next

    for i = 0 to 29
%>
<TR>
	<TD nowrap width=100>0x&nbsp;<INPUT type="text" style="width:60" id=txtID name=txtID value="<%=OutputID(i)%>" maxlength=4>&nbsp;h</TD>
	<TD><INPUT type="text" style="width:100%" id=txtComments name=txtComments value="<%=server.htmlencode(OutputComments(i))%>" maxlength=24></TD>
    <TD><INPUT type="text" style="width:100%" id=txtCHID name=txtCHID value="<%=server.htmlencode(OutputCHID(i))%>" maxlength=38></TD>
</TR>
    <%next%>
</Table>

</BODY>
</HTML>
