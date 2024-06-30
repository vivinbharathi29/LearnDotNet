<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function keyPress() {
        if (window.event.keyCode == 13) {
            event.returnValue = false;
            event.cancel = true;
            window.parent.frames("LowerWindow").cmdOK_onclick();            
        }
    }

    function cmdOK_onclick() {
        frmMain.action = "ChooseVersionsSave.asp"
        frmMain.submit();
    }


    function chkAllVersionsChecked() {
        
        if (typeof(frmMain.chkVersion.length)=="undefined")
            frmMain.chkVersion.checked = frmMain.chkAllVersions.checked;
        else
            {
               for (i=0;i<frmMain.chkVersion.length;i++)
                    frmMain.chkVersion[i].checked = frmMain.chkAllVersions.checked;

            }
       }

       function window_onload() {
           frmMain.txtID.focus();
       }
//-->
</SCRIPT>

<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:x-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
h3{
    FONT-FAMILY: Verdana;
    FONT-SIZE:small;
  }
</STYLE>

</HEAD>


<BODY bgcolor="ivory" onload="window_onload();" onkeypress="javascript: keyPress();">
    <h3>Choose Deliverable Versions</h3>
<%

    if request("optType") = "2" then
        strOptChooseChecked = " checked "
        strOptIDChecked = ""
    else
        strOptChooseChecked = ""
        strOptIDChecked = " checked "
    end if

    dim cn, rs

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
%>
<form id=frmMain method="post" action="ChooseVersionsMain.asp">
<table width="100%">
    <tr>
        <td width=10 valign=top><input id="optID" name="optType" type="radio" <%=strOptIDChecked%> value="1" /></td>
        <td>Enter Version ID Numbers: <font color="green" size="1"> (Enter a comma separated list to select multiple versions)</font><br>
            <input id="txtID" name="txtID"  type="text" style="width:100%" onfocus="javascript: optID.checked=true;" value="<%=server.htmlencode(request("txtID"))%>" /><br /><br />
        </td>
    </tr>
    <tr>
        <td width=10 valign=top><input id="optChoose" name="optType" type="radio" <%=strOptChooseChecked%> value="2" /></td>
        <td>Lookup Versions:<br>
            <select style="width:100%" id="cboRoot" name="cboRoot" onchange="javascript: frmMain.submit();" onfocus="javascript: optChoose.checked=true;">
                <option value="0"></option>
                <%
                    rs.open "Select ID, Name from deliverableroot with (NOLOCK) where active=1 order by name",cn
                    do while not rs.eof
                        if request("cboRoot") = "" or request("cboRoot") = "0"  or trim(strOptChooseChecked) = "" then
                            response.write "<option value=""" & rs("ID") & """>" &  rs("name") & "</option>"
                        else
                            if clng(request("cboRoot")) = rs("ID") then
                                response.write "<option selected value=""" & rs("ID") & """>" &  rs("name") & "</option>"
                            else
                                response.write "<option value=""" & rs("ID") & """>" &  rs("name") & "</option>"
                            end if
                        end if
                        rs.movenext 
                    loop
                    rs.close
                
                %>
            </select><br /><br />
        <div style="border-right: gainsboro 1px solid; border-top: gainsboro 1px solid; overflow-y: scroll;
            border-left: gainsboro 1px solid; width: 100%; border-bottom: gainsboro 1px solid;
            height: 143px; background-color: white" id="DIV1">
            <%if request("cboRoot") <> "" and request("optType") = "2" then
                rs.open "spGetDelDeliverableList " & clng(request("cboRoot")),cn
                if not(rs.eof and rs.bof) then
            %>
            <table>
                <thead bgcolor="LightSteelBlue">
                    <tr>
                    <td nowrap style="width:10px; border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset"><input id="chkAllVersions" type="checkbox" onclick="chkAllVersionsChecked();"  /></td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset">&nbsp;ID&nbsp;</td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset">&nbsp;Version&nbsp;</td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset">&nbsp;Model&nbsp;Number&nbsp;</td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset">&nbsp;Part&nbsp;Number&nbsp;</td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset">&nbsp;Active&nbsp;</td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;  border-bottom: 1px outset" width="100%">&nbsp;Vendor&nbsp;</td>
                    </tr>
                </thead>
                <%do while not rs.eof
                    strVersion = rs("Version") & ""
                    if trim(rs("revision") & "") <> "" then
                        strVersion = strVersion & "," & rs("revision")
                    end if
                    if trim(rs("pass") & "") <> "" then
                        strVersion = strVersion & "," & rs("pass")
                    end if

                    if rs("Active") then
                        strActive = "Yes"
                    else
                        strActive = "No"
                    end if

                    strAllVersions = request("txtAllVersions")
                    if request("chkVersion") <> "" then
                        if trim(strAllVersions) = "" then
                            strAllVersions = request("chkVersion")
                        else
                            strAllVersions = strAllVersions & ", " & request("chkVersion")
                        end if
                    end if

                    
                    if inarray(split(strAllversions,","),rs("ID") & "") then
                        strVersionChecked = " checked "
                    else
                        strVersionChecked = " "
                    end if
                %>
                    <tr>
                        <td>
                            <input id="chkVersion" name="chkVersion" type="checkbox" <%=strVersionChecked%> value="<%=rs("ID")%>" />
                        </td>
                        <td nowrap>&nbsp;<%=rs("ID")%>&nbsp;&nbsp;</td>
                        <td nowrap>&nbsp;<%=strversion%>&nbsp;&nbsp;</td>
                        <td nowrap>&nbsp;<%=rs("PartNumber") & ""%>&nbsp;&nbsp;</td>
                        <td nowrap>&nbsp;<%=rs("ModelNumber") & ""%>&nbsp;&nbsp;</td>
                        <td nowrap>&nbsp;<%=strActive%>&nbsp;&nbsp;</td>
                        <td nowrap>&nbsp;<%=rs("Vendor") & ""%>&nbsp;&nbsp;</td>
                    </tr>
                <%
                    rs.movenext 
                loop
                %>
            </table>
            <%
                else
                    response.write "No versions found"
                end if
            else
                response.write "&nbsp;"
            end if %>
        </div>

        </td>
    </tr>

</table>
<textarea style="display:none" id="txtAllVersions" name="txtAllVersions" cols="20" rows="2"><%=server.HTMLEncode(strAllVersions)%></textarea>
</form>
<%

    set rs = nothing
    cn.Close
    set cn = nothing

		function InArray(MyArray,strFind)
			dim strElement
			dim blnFound
			
			blnFound = false
			for each strElement in MyArray
				if trim(strElement) = trim(strFind) then
					blnFound = true
					exit for
				end if
			next
			InArray = blnFound
		end function

%>
</BODY>
</HTML>




