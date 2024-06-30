<%@  language="VBScript" %>
<%
Dim AppRoot
AppRoot = Session("ApplicationRoot")
%>
<html>
<head>
    <title></title>
    <script type="text/javascript" src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js"></script>
    <script type="text/javascript">
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


function ondblclick_Row(){
	cmdTo_onclick();
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
			if (txtTo.value == ""){
				txtTo.value = SelectedArray[i];
				hidReturn.value = document.getElementById(SelectedArray[i]).SKNO + "|" + SelectedArray[i];
				}
			else
				{
				if (txtTo.value.indexOf(document.getElementById(SelectedArray[i]).SKNO) == -1){
				    txtTo.value = SelectedArray[i];
				    hidReturn.value = document.getElementById(SelectedArray[i]).SKNO + "|" + SelectedArray[i];
				    }
				}
		}
}

function onmouseup_Row(){
}

function onmousedown_Row(){
	var RowElement;
	var SelectedArray;
	var i;
	
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
		txtSelected.value = RowElement.id;
		}
		
}

function window_onload() {
}

function GetSpareKits( CategoryId, ProductVersionId ) {
    var divSpareKits = document.getElementById("divSpareKits");
    divSpareKits.innerHTML = "";
    jsrsExecute("<%=AppRoot %>/Service/rsService.asp", GetSpareKitsCallBack, "GetKitsForProductCategory", Array(CategoryId, ProductVersionId));
}

function GetSpareKitsCallBack(result) {
    var divSpareKits = document.getElementById("divSpareKits");
    divSpareKits.innerHTML = result;
}

function selProducts_selectedIndexChanged() {
    var selProducts = document.getElementById("selProducts");
    var hidCategoryId = document.getElementById("hidCategoryId");
    GetSpareKits(hidCategoryId.value, selProducts.value);
}

//-->
    </script>
    <style type="text/css">
        TD
        {
            font-family: Verdana;
            font-size: x-small;
        }
        BODY
        {
            font-family: Verdana;
            font-size: x-small;
        }
        .header
        {
            background-color: #004874;
            color: White;
            border: solid 1px;
            border-bottom-style: outset;
            font: bold x-small;
            white-space: nowrap;
        }
        .header:Hover
        {
            color: Red;
            cursor: hand;
        }
        .row{}
        .row:Hover
        {
        	cursor: hand;
        }
    </style>
</head>
<body onload="return window_onload()">
        <%
	dim cn, rs
	dim strEmployees
	Dim strPartnerId
	
	set cn = server.CreateObject("ADODB.connection")
	set rs = server.CreateObject("ADODB.recordset")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	strEmployees = ""
	strEmployeesID = ""
	strPartnerId = ""

        %>
        <b>Existing Spare Kits</b><br />
        <br />
        Lookup:&nbsp;<select id="selProducts" onchange="selProducts_selectedIndexChanged()"><option value="0">-- Select A Product --</option>
<%
rs.Open "spGetProductPartner " & Request.QueryString("PVID"), cn, adOpenStatic
If Not rs.EOF Then
    strPartnerId = rs("PartnerId")
End If
rs.Close

If strPartnerId <> "" Then
rs.Open "usp_ListProductsByPartner " & strPartnerId, cn, adOpenStatic
Do While Not rs.EOF
    Response.Write "<option value=""" & rs("Id") & """>" & rs("DotsName") & "</option>"
    rs.MoveNext
Loop
rs.Close
End If
%>        
        </select>
            <br />
        <font size="1">
            <br />
        </font>
        <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
            overflow-x: scroll; border-left: steelblue 1px solid; width: 100%; border-bottom: steelblue 1px solid;
            height: 320px; background-color: white" id="divSpareKits">
            <table id="tblSpareKits" width="100%">
                    <tr style="position: relative; top: expression(document.getElementById('divSpareKits').scrollTop-2)">
                        <td class="header" onclick="SortTable( 'tblSpareKits', 0 ,0,1);">
                            &nbsp;Spare Kit No&nbsp;
                        </td>
                        <td class="header" onclick="SortTable( 'tblSpareKits', 1 ,0,1);">
                            &nbsp;Description&nbsp;
                        </td>
                    </tr>
            </table>
        </div>
        <%
	set rs= nothing
	cn.Close
	set cn = nothing

        %>
        <table width="100%">
            <tr>
                <td valign="top">
                    <input type="button" value="Add -->" id="cmdTo" name="cmdTo" language="javascript"
                        onclick="return cmdTo_onclick()">
                </td>
                <td width="100%">
                    <textarea style="width=100%" rows="3" cols="80" id="txtTo" name="txtTo"></textarea>
                </td>
            </tr>
        </table>
        <input type="hidden" id="hidReturn" />
        <input style="width: 100%" type="hidden" id="txtSelected" name="txtSelected" value="" />
        <input type="hidden" id="hidCategoryId" value="<%= Server.HtmlEncode(Request.QueryString("Category")) %>" />
</body>
</html>
