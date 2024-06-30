<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "../_ScriptLibrary/sort.js" -->

var KeyString = "";
var SelectedRow = "";
var SelectedArray;
var i;


function mouseover_Column(){
	window.event.srcElement.style.cursor="hand";
	window.event.srcElement.style.color="red";
}

function mouseout_Column(){
	window.event.srcElement.style.color="black";
}

function txtLookup_onkeydown(){


	if (txtLookup.value == "")
		return;

	//Reset any previously selected items
	if (txtSelected.value != "")
		{
		SelectedArray = txtSelected.value.split(";");	
		for (i=0;i<SelectedArray.length;i++)		
			{
			if (SelectedArray[i] != "")
				{
				document.getElementById(SelectedArray[i]).style.backgroundColor = "White";
				document.getElementById(SelectedArray[i]).style.color = "Black";
				}
			}
		txtSelected.value = "";
		SelectedRow = "";
		}
	
	

	//Continue selecting the current one.	
	KeyString= txtLookup.value;
	KeyString = ";" + KeyString.toLowerCase();
	var strHit;
	var strHitEnd;
	var strNames = text1.value.toLowerCase();
	var strHitAt=strNames.indexOf(KeyString);
	if (strHitAt > -1)
		{
			strHit = strNames.substring(strHitAt+1);
			strHitEnd= strHit.indexOf(";");
			if (strHitEnd > -1)
				{
				strHit = text1.value.substr(strHitAt + 1,strHitEnd);

				RowLocation = document.getElementById(strHit).offsetTop -20;
				DIV1.scrollTop = RowLocation;
				document.getElementById(strHit).style.backgroundColor = "MediumBlue";
				document.getElementById(strHit).style.color = "white";
				if (SelectedRow != strHit)
					{
					if (SelectedRow != "")
						{
						document.all(SelectedRow).style.backgroundColor = "white";
						document.all(SelectedRow).style.color = "black";
						}
					SelectedRow = strHit;
					txtSelected.value = strHit;
					}
				}
		}
	if (event.keyCode==13)
		{
		cmdTo_onclick();
		txtLookup.value = "";
		}
}


function ondblclick_Row(){
	cmdTo_onclick();
	txtLookup.value = "";

}


function cmdTo_onclick() {
	var SelectedArray;
	var i;
	
	if (txtSelected.value == "")
		txtTo.value= "";
	else
		{
		SelectedArray = txtSelected.value.split(";");
		for (i=0;i<SelectedArray.length;i++)
			if (txtTo.value == "")
			    txtTo.value = document.getElementById(SelectedArray[i]).getAttribute("email");
			else
				{
			    if (txtTo.value.indexOf(document.getElementById(SelectedArray[i]).getAttribute("email")) == -1)
				    txtTo.value = txtTo.value + "; " + document.getElementById(SelectedArray[i]).getAttribute("email");
				}
		}
}

function onmouseover_Row() {
	var RowElement;
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
	
//	if (RowElement.style.backgroundColor == "")
//		{
		//RowElement.style.backgroundColor="#99ccff";
		RowElement.style.cursor="hand";
//		}
}

function onmouseout_Row() {
	var RowElement;
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
	
//	if (RowElement.style.backgroundColor == "")
//		{
		//RowElement.style.backgroundColor="#99ccff";
		RowElement.style.cursor="default";
//		}
}

function onmouseup_Row(){
	txtLookup.focus();

}

function onmousedown_Row(){
	var RowElement;
	var SelectedArray;
	var i;
	
	txtLookup.value = "";
	
	RowElement = window.event.srcElement;
	while (RowElement.className != "Row")
	{
		RowElement = RowElement.parentElement;
	}
	
	if (event.ctrlKey)
		{
		if (RowElement.style.backgroundColor.toLowerCase() != "mediumblue")
			{
			RowElement.style.backgroundColor = "MediumBlue";
			RowElement.style.color = "white";
			if (txtSelected.value=="")
				txtSelected.value = RowElement.id;
			else
				txtSelected.value = txtSelected.value + ";" + RowElement.id;
			}
			document.focus();
		}
	else
		{
		RowElement.style.backgroundColor = "MediumBlue";
		RowElement.style.color = "white";

		if (txtSelected.value != "" && RowElement.id != txtSelected.value)
			{
			SelectedArray = txtSelected.value.split(";");	
			for (i=0;i<SelectedArray.length;i++)		
				{
				if (SelectedArray[i] != "")
					{
					document.getElementById(SelectedArray[i]).style.backgroundColor = "White";
					document.getElementById(SelectedArray[i]).style.color = "Black";
					}
				}
			}
		txtSelected.value = RowElement.id
		}
		
}

function window_onload() {
	txtLookup.focus();
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
TD{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}
BODY{
	FONT-FAMILY: Verdana;
	FONT-SIZE: x-small;
}
</STYLE>
<BODY bgcolor=ivory LANGUAGE=javascript onload="return window_onload()">
<P>


<%
	dim cn, rs
	dim strEmployees
	
	set cn = server.CreateObject("ADODB.connection")
	set rs = server.CreateObject("ADODB.recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	strEmployees = ""
	strEmployeesID = ""
	
%>

<B>Excalibur Address Book</B><BR><BR>
Lookup:&nbsp;<INPUT type="text" id=txtLookup name=txtLookup value="" LANGUAGE=javascript onkeyup="return txtLookup_onkeydown()"></STRONG><BR><font size=1><BR></FONT>


<div style="BORDER-RIGHT: steelblue 1px solid; BORDER-TOP: steelblue 1px solid; OVERFLOW-Y: scroll; OVERFLOW-X: scroll; BORDER-LEFT: steelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: steelblue 1px solid; HEIGHT: 220px; BACKGROUND-COLOR: white" id=DIV1>
	<TABLE ID=TableEmployee width="100%">
	<THEAD><TR style="POSITION: relative;TOP: expression(document.getElementById('DIV1').scrollTop-2)">
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 0 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Name&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 1 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Group&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 2 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Division&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 3 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Manager&nbsp;</TD>
		<TD bgcolor=lightsteelblue onmouseover="mouseover_Column();" onmouseout="mouseout_Column();" onclick="SortTable( 'TableEmployee', 4 ,0,1);" style="BORDER-RIGHT: 1px outset; BORDER-TOP: 1px outset; BORDER-LEFT: 1px outset; BORDER-BOTTOM: 1px outset">&nbsp;Email&nbsp;</TD>
	</TR></THEAD>
	<%
	dim strDivision
	rs.Open "splistEmployees",cn,adOpenStatic
	do while not rs.EOF
	
		if trim(rs("PartnerID")) <> "1" then
			strDivision =  rs("Partner") & ""
		else
			select case trim(rs("Division") & "")
			case "1"
				strDivision = "Mobile"
			case "2"
				strDivision = "bPC"
			case "3"
				strDivision = "Mobile"
			case "4"
				strDivision = "Workstations"
			case "5"
				strDivision = "cPC"
			case else
				strDivision = "Other"
			end select
		end if
		if rs("Active") = 1 or trim(request("ShowAll")) = "1" then
			Response.Write "<TR class=""Row"" Email=""" & rs("Email") & """ ID=""" & rs("Name") & """ language=javascript ondblclick=""ondblclick_Row();"" onmouseover=""onmouseover_Row();"" onmouseout=""onmouseout_Row();"" onmousedown=""onmousedown_Row();"" onmouseup=""onmouseup_Row();"">"
			if trim(request("ShowAll")) = "1" and rs("Active") = 0 then
    			Response.Write "<TD nowrap>" & rs("Name") & " [Inactive]</TD>"
	        else
    			Response.Write "<TD nowrap>" & rs("Name") & "</TD>"
	        end if
			Response.Write "<TD nowrap>" & rs("Workgroup") & "</TD>"
			Response.Write "<TD nowrap>" & strDivision & "</TD>"
			if trim(rs("Manager") & "") = "" then
			    Response.Write "<TD nowrap>&nbsp;</TD>"
			else
			    Response.Write "<TD nowrap>" & rs("Manager") & "</TD>"
			end if
			Response.Write "<TD nowrap>" & rs("Email") & "</TD>"
			Response.Write "</TR>"
			strEmployees = strEmployees & ";" & rs("Name")
		end if
		rs.MoveNext
	loop
	rs.Close
	%>
	
	</TABLE>
            
</div>

<%
	set rs= nothing
	cn.Close
	set cn = nothing

%>
<TABLE width=100%><TR><TD valign=top><INPUT type="button" value="To -->" id=cmdTo name=cmdTo LANGUAGE=javascript onclick="return cmdTo_onclick()"></TD>
			<TD width=100%>
				<TEXTAREA style="WIDTH=100%" rows=3 cols=80 id=txtTo name=txtTo><%=request("AddressList")%></TEXTAREA>
				</TD>
		</TR>
</TABLE>


<INPUT type="hidden" id=text1 name=text1 value="<%=";" & strEmployees%>">

<BR>
<INPUT style="WIDTH:100%" type="hidden" id=txtSelected name=txtSelected value="">
</BODY>
</HTML>
