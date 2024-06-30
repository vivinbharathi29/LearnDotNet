<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {
    var i;
    var strID="";

    for (i = 0; i < FeatureBox.length; i++)
        if (FeatureBox[i].checked)
            strID = strID + "|" + FeatureBox[i].value;

    if (strID != "")
        strID = strID.substr(1)
    window.returnValue = strID;
    window.parent.opener = 'X';
    window.parent.open('', '_parent', '')
    window.parent.close();
}

function chkAll_onclick() {
    var i;

    for (i = 0; i < FeatureBox.length; i++)
        FeatureBox[i].checked = chkAll.checked;

}

//-->
</SCRIPT>
</HEAD>
<STYLE>
BODY
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}
TD
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}

H1
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: small;
}
.DateInput
{
	font-family: Verdana;
	font-size: x-small;
	height: 20;
	width: 80;
	border: solid 1px gray;
}
</STYLE>
<body style="background-color:ivory;" id="MainBody">
<h1>Deliverable Features</h1>
<font size=2><b>Choose Features:</b><BR></font>
    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: auto; border-left: steelblue 1px solid; width: 100%;height:200px; border-bottom: steelblue 1px solid; background-color: white" id="DIV1">
    <table id="tabSections" width="100%">
        <thead>
            <tr style="position: relative; top: expression(document.getElementById('DIV1').scrollTop-2);">
                <td bgcolor="lightsteelblue" nowrap style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset"><input type="checkbox" id="chkAll" name="chkAll" language=javascript onclick="chkAll_onclick();" ></td>
                <td bgcolor="lightsteelblue" style="width:100%; border-right: 1px outset; border-top: 1px outset;border-left: 1px outset; border-bottom: 1px outset">Features</td>
            </tr>
        </thead>
        <tbody>
<%
    dim rs, cn,strItems
	
    set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	strItems = ""

	set rs = server.CreateObject("ADODB.recordset")

    rs.open "spListDeliverableRootFeatures 3113,1" , cn,adOpenStatic

    do while not rs.EOF
        if instr("," & trim(request("IDList")) & ",","," & trim(rs("ID")) & ",") > 0 then
            response.write "<tr><td width=10><input id=""FeatureBox"" checked type=""checkbox"" value=""" & rs("ID") & "^" &  rs("Name") & """>&nbsp;</td><td>" & rs("name") & "</td></tr>"
        else
            strItems = strItems & "<tr><td width=10><input id=""FeatureBox"" type=""checkbox"" value=""" & rs("ID") & "^" &  rs("Name") & """>&nbsp;</td><td>" & rs("name") & "</td></tr>"
        end if
        rs.MoveNext
    loop
    rs.Close
    response.Write strItems

    set rs = nothing
    cn.Close
    set cn = nothing

%>


        </tbody>
    </table>
    </div>

</BODY>
</HTML>
