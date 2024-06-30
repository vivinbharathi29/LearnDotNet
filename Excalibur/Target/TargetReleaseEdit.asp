<%@ Language=VBScript %>
 <% OPTION EXPLICIT %>
    <!-- #include file = "../includes/Security.asp" -->
    <!-- #include file="../includes/DataWrapper.asp" -->
    <!-- #include file="../includes/no-cache.asp" -->
    <!-- #include file="../includes/lib_debug.inc" -->

    <%
        
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"	  

    Dim cn, dw, cmd, rs	, cm
    dim p
    dim bTargetRelease, returnValue, bTargeted, strReleaseInfo
    dim pvID, RootID, VersionID, ViewOnly  
    dim RowsChanged  
    dim CurrentUserID
            
    Set rs = Server.CreateObject("ADODB.RecordSet")
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

    'Get User
	dim CurrentDomain,CurrentUser
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
		CurrentUserID = rs("ID") 
	end if
	rs.Close
      
   
    pvID = 0
    if Request.QueryString("ProductID") <> "" then
        pvID = clng(Request.QueryString("ProductID"))
    end if
        
    RootID=0
    if Request.QueryString("RootID") <> "" then
	    RootID = clng(Request.QueryString("RootID"))
    end if

    VersionID=0
    if Request.QueryString("VersionID") <> "" then
	    VersionID = clng(Request.QueryString("VersionID"))
    end if  
        
    ViewOnly=0
    if Request.QueryString("ViewOnly") <> "" then
	   ViewOnly = clng(Request.QueryString("ViewOnly"))
    end if    

    if len(Request.form("bTargetRelease")) > 0 then
		bTargetRelease = true
	end if

    strReleaseInfo = Request.form("txtReleaseInfo")

    bTargeted = ""

    dim SelectedRelease

    if bTargetRelease then        	
    'section to execute when save/OK button is clicked
        Set dw = New DataWrapper
        Set cmd = dw.CreateCommAndSP(cn, "usp_ProductDeliverable_TargetReleases")
                dw.CreateParameter cmd, "@p_intDeliverableVersionID", adInteger, adParamInput, 8, VersionID    
                dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, pvID
                dw.CreateParameter cmd, "@p_xmlReleaseInfo", adLongVarChar, adParamInput, len(strReleaseInfo), strReleaseInfo     
                dw.CreateParameter cmd, "@p_UserID", adInteger, adParamInput, 8, CurrentUserID                    
            returnValue = dw.ExecuteNonQuery(cmd)
        bTargeted = "1"
        'set cm = nothing

     end if

        set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
	    cm.CommandType = 4	    
       
        cm.CommandText = "usp_ProductDeliverable_ViewReleases"
	
        Set p = cm.CreateParameter("@p_intProductVersioID", 3, &H0001)
	    p.Value = pvID
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@p_intRootID", 3, &H0001)
	    p.Value = RootID
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@p_intVersionID", 3, &H0001)
	    p.Value = VersionID
	    cm.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 

	    set cm=nothing
       
   

       

	%>


<HTML>
    <head>
        
        <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
        <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
        <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
        <script type="text/javascript" LANGUAGE="javascript">
            $(function () {
                $("input:button").button();
            });
            function body_onload() {
                var strbTargeted = document.getElementById("bTargeted").value;                
                var targetReleases,targetReleaseIDs;
                if (strbTargeted != "") {
                    if ('<%=Request("pulsarplusDivId")%>' != undefined && '<%=Request("pulsarplusDivId")%>' != "") {
                        window.parent.closeExternalPopup();
                    }
                    else
                      window.parent.CloseTargetReleasePopup(true,'','');
                }

            }
            function btnOKClick() {
                if (!$('input:checkbox').is(':checked'))
                {
                    alert("Please select at least one release");
                    return false;
                }
                var targetReleases = "";
                var targetReleaseIDs = "", strTargetNotes = "";
                var xmlReleaseInfo = '<?xml version="1.0" encoding="iso-8859-1" ?><ReleaseInfo>';
                $('input[name="chkRelease"]:checked').each(function () {
                    targetReleaseIDs = targetReleaseIDs == "" ? this.value : targetReleaseIDs + "," + this.value;
                    targetReleases = targetReleases == "" ? this.getAttribute("ReleaseName") : targetReleases + ", " + this.getAttribute("ReleaseName");
                    strTargetNotes = $("#txtNotes" + this.value).val();
                    strTargetNotes = strTargetNotes.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;").replace("'", "&apos;");//replace the xml string
                    xmlReleaseInfo = xmlReleaseInfo + '<Release><ReleaseID>' + this.value + '</ReleaseID><TargetNotes>' + strTargetNotes + '</TargetNotes></Release>';
                });
                xmlReleaseInfo = xmlReleaseInfo + '</ReleaseInfo>';
                $("#txtReleaseInfo").val(xmlReleaseInfo);
                window.parent.SetTargetInfomation(targetReleases, targetReleaseIDs);
                document.getElementById("cmdOK").disabled = true;
                document.targetreleaseform.action = 'TargetReleaseEdit.asp?ProductID=<%=cstr(pvID)%>&RootID=<%=cstr(RootID)%>&VersionID=<%=cstr(VersionID)%>&pulsarplusDivId=<%=Request.QueryString("pulsarplusDivId")%>';
                $("#bTargetRelease").val("1");
                document.targetreleaseform.submit();
            }            
            function Cancel_onclick(pulsarplusDivId) {
                if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                    // For Closing current popup if Called from pulsarplus
                    window.parent.closeExternalPopup();
                }
                else
                 window.parent.CloseTargetReleasePopup(false,'','');
            }
            function SelectAll()
            {
                var bChecked = $('#chkSelectAll').is(":checked");
                var checkBoxes = document.getElementsByTagName("input");
                for (i = 0; i < checkBoxes.length; i++) {
                    if (checkBoxes[i].name == "chkRelease") {
                        checkBoxes[i].checked = bChecked;
                    }
                }                              
            }
        </script>
    </head>

    <BODY id=bdy onload="return body_onload();" style="background: ivory;">
        <form name=targetreleaseform method=post>	        
	        <table border="1" BGCOLOR="Ivory" CELLSPACING="1" CELLPADDING="2" bordercolor=tan border="1" style="width:98%;">
                <colgroup>
                    <col style="width:10%" />
                    <col style="width:25%" />
                    <col style="width:65%" />
                </colgroup>        
                <tr bgcolor=cornsilk>
                    <%if ViewOnly = 0 then%>
                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold; text-align:center;"><input type="checkbox" id="chkSelectAll" onclick="return SelectAll();" /></td>
                    <%end if%>
                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Release</td>
		            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target Notes</td>
                </tr>                  
                <%do while not rs.eof%>
                    <!--add row for each release-->
                    <tr>
                        <%if ViewOnly = 0 then
                            if trim(rs("Targeted")) = "False" then%>
                                <%if rs("NoOfReleases") = "1" then %>
                                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold; text-align:center;"><input checked id="chkRelease" value='<%=rs("ID")%>' ReleaseName='<%=rs("Name")%>' name="chkRelease" type="checkbox"></td>
                                <%else%>
                                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold; text-align:center;"><input id="chkRelease" value='<%=rs("ID")%>' ReleaseName='<%=rs("Name")%>' name="chkRelease" type="checkbox" ></td>                                    
                                <%end if%>
                            <%else%>
                                <td style="font-size:xx-small; font-family:Verdana; font-weight:bold; text-align:center;"><input checked id="chkRelease" value='<%=rs("ID")%>' ReleaseName='<%=rs("Name")%>' name="chkRelease" type="checkbox"></td>
                            <%end if
                        end if%>
                        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;"><%=rs("Name")%></td>
                        <%if ViewOnly = 0 then%>
                            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">
                                <INPUT type="text" id="txtNotes<%=rs("ID")%>" name="txtNotes<%=rs("ID")%>" style="WIDTH:100%" value="<%=rs("TargetNotes")%>" maxlength="255" class="text-option">
			                    <INPUT type="hidden" id="txtNotesTag<%=rs("ID")%>" name="txtNotesTag" value="<%=rs("TargetNotes")%>" >
                            </td>
                        <%else%>
                            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">
                                <%=server.htmlencode(rs("TargetNotes"))%> &nbsp;
                            </td>
                        <%end if%>
                    </tr>
                <% rs.movenext
                loop
                rs.close
                if ViewOnly = 0 then%>                
                <tr>
		            <TD colspan="3" style="text-align:right;">
                        <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return Cancel_onclick('<%=Request("pulsarplusDivId")%>')"  >
                        <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return btnOKClick()">		                
                     </TD>
                </TR>
                <%end if%>
            </table> 
            <input type="hidden" id="txtReleaseInfo" name="txtReleaseInfo" value="" />                         
            <INPUT type="hidden" id="bTargetRelease" name="bTargetRelease" value=""/>
            <INPUT type="hidden" id="bTargeted" name="bTargeted" value="<%=bTargeted%>"/>
        </form>
    </body>
</HTML>