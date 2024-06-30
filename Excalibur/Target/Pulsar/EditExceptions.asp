<%@ Language=VBScript %>
 <% OPTION EXPLICIT %>
    <!-- #include file = "../../includes/Security.asp" -->
    <!-- #include file="../../includes/DataWrapper.asp" -->
    <!-- #include file="../../includes/no-cache.asp" -->
    <!-- #include file="../../includes/lib_debug.inc" -->

    <%
        
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"	
        
    Dim cn, dw, cmd, rs	, cm
    dim p
    dim bTargetRelease, returnValue, bTargeted, strReleaseInfo
    dim pvID, RootID, VersionID, ViewOnly
    Dim Security, m_UserFullName
    dim RowsChanged
    dim ReleaseID

   	Set Security = New ExcaliburSecurity
    m_UserFullName = Security.CurrentUserFullName()
   
    ReleaseID = 0        
    if Request.QueryString("ReleaseID") <> "" then
        ReleaseID = clng(Request.QueryString("ReleaseID"))
    end if

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

    Set rs = Server.CreateObject("ADODB.RecordSet")
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

    if bTargetRelease then        	
    'section to execute when save/OK button is clicked
        Set dw = New DataWrapper
        Set cmd = dw.CreateCommAndSP(cn, "usp_ProductDeliverable_UpdateExceptions")
                dw.CreateParameter cmd, "@p_intDeliverableVersionID", adInteger, adParamInput, 8, VersionID    
                dw.CreateParameter cmd, "@p_intProductVersionID", adInteger, adParamInput, 8, pvID
                dw.CreateParameter cmd, "@p_xmlReleaseInfo", adLongVarChar, adParamInput, len(strReleaseInfo), strReleaseInfo                       
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

        Set p = cm.CreateParameter("@p_intTargetedOnly", 3, &H0001)
	    p.Value = 1
	    cm.Parameters.Append p

        Set p = cm.CreateParameter("@p_intReleaseID", 3, &H0001)
	    p.Value = ReleaseID
	    cm.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 

	    set cm=nothing
       
   

       

	%>


<HTML>
    <head>
        
        <link href="../../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
        <script src="../../includes/client/jquery.min.js" type="text/javascript"></script>
        <script src="../../includes/client/jquery-ui.min.js" type="text/javascript"></script>
        <script type="text/javascript" LANGUAGE="javascript">
            $(function () {
                $("input:button").button();
            });
            function body_onload() {
                var strbTargeted = document.getElementById("bTargeted").value;               
                if (strbTargeted != "") {
                    if (parent.window.parent.loadDatatodiv != undefined) {
                        parent.window.parent.closeExternalPopup();
                    }
                    else {
                        parent.window.ChangeTargetNotesResult_Pulsar();
                        parent.window.modalDialog.cancel();
                    }
                }

            }
            function btnOKClick() {                
                var targetReleases = "";
                var targetReleaseIDs = "";
                var xmlReleaseInfo = '<?xml version="1.0" encoding="iso-8859-1" ?><ReleaseInfo>';
                var bNoteExists = 0, strTargetNotes = "";
                $('input[name="chkRelease"]:checked').each(function () {
                    targetReleaseIDs = targetReleaseIDs == "" ? this.value : targetReleaseIDs + "," + this.value;
                    targetReleases = targetReleases == "" ? this.ReleaseName : targetReleases + ", " + this.ReleaseName;
                    if ($("#txtNotes" + this.value).val() != "")
                        bNoteExists = "1";
                    strTargetNotes = $("#txtNotes" + this.value).val();
                    strTargetNotes = strTargetNotes.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\"", "&quot;").replace("'", "&apos;");//replace the xml string
                    xmlReleaseInfo = xmlReleaseInfo + '<Release><ReleaseID>' + this.value + '</ReleaseID><TargetNotes>' + strTargetNotes + '</TargetNotes>';
                    if ($('#chkOOC' + this.value).is(":checked"))
                        xmlReleaseInfo = xmlReleaseInfo + '<OOCRelease>1</OOCRelease>';
                    else
                        xmlReleaseInfo = xmlReleaseInfo + '<OOCRelease>0</OOCRelease>';
                    xmlReleaseInfo = xmlReleaseInfo + '<Type>' + $('input[name="optScope' + this.value + '"]:checked').val() + '</Type>';
                    xmlReleaseInfo = xmlReleaseInfo + '</Release>';
                });
                xmlReleaseInfo = xmlReleaseInfo + '</ReleaseInfo>';
                $("#txtReleaseInfo").val(xmlReleaseInfo);
                parent.window.SetNoteExists(bNoteExists, strTargetNotes);
                $("#TargetNotes").val(strTargetNotes);
                document.getElementById("cmdOK").disabled = true;
                document.updatetargetreleaseform.action = "EditExceptions.asp?ProductID=<%=cstr(pvID)%>&RootID=<%=cstr(RootID)%>&VersionID=<%=cstr(VersionID)%>&ReleaseID=<%=cstr(ReleaseID)%>";
                $("#bTargetRelease").val("1");
                document.updatetargetreleaseform.submit();
            }            
            function Cancel_onclick() {
                if (parent.window.parent.loadDatatodiv != undefined) {
                    parent.window.parent.closeExternalPopup();
                }
                else { 
                    parent.window.modalDialog.cancel();
                }
            }
            function ChangeThis_onclick(ReleaseID) {
                document.getElementById("optThis" + ReleaseID).checked = true;
                document.getElementById("optFuture" + ReleaseID).checked = false;
            }

            function ChangeDefault_onclick(ReleaseID) {
                document.getElementById("optThis" + ReleaseID).checked = false;
                document.getElementById("optFuture" + ReleaseID).checked = true;
            }

            function ChangeThis_onmouseover() {
                window.event.srcElement.style.cursor = "hand";
            }

            function ChangeDefault_onmouseover() {
                window.event.srcElement.style.cursor = "hand";
            }
        </script>
    </head>

    <BODY id=bdy onload="return body_onload();" style="background: ivory;">
        <form name=updatetargetreleaseform method=post>	        
	        <table border="1" BGCOLOR="Ivory" CELLSPACING="1" CELLPADDING="2" bordercolor=tan border="1" style="width:99%;">
                <colgroup>
                    <col style="width:10%" />
                    <col style="width:40%" />
                    <col style="width:13%" />
                    <col style="width:32%" />
                </colgroup>        
                <tr bgcolor=cornsilk>
                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Release</td>
		            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Target Notes</td>
                    <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Out&nbsp;of<br />Cycle&nbsp;Release</td>
		            <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">Scope</td>
                </tr>                  
                <%do while not rs.eof%>
                    <!--add row for each release-->
                    <tr>  
                        <td style="display:none"><input checked id="chkRelease" value='<%=rs("ID")%>' ReleaseName='<%=rs("Name")%>' name="chkRelease" type="checkbox"></td>                      
                        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;"><%=rs("Name")%></td>                        
                        <td style="font-size:xx-small; font-family:Verdana; font-weight:bold;">
                            <textarea id="txtNotes<%=rs("ID")%>" name="txtNotes<%=rs("ID")%>" style="WIDTH: 100%; height:50px" rows="2" maxlength="255"><%=rs("TargetNotes")%></textarea>
			                <INPUT type="hidden" id="txtNotesTag<%=rs("ID")%>" name="txtNotesTag" value="<%=rs("TargetNotes")%>" >
                        </td>
                        <% if (rs("OOCRelease") = "True") then %>
                            <td style="text-align:center"><INPUT type="checkbox" id="chkOOC<%=rs("ID")%>" checked name="chkOOC<%=rs("ID")%>"></td>     
                        <%else%>
                            <td style="text-align:center"><INPUT type="checkbox" id="chkOOC<%=rs("ID")%>" name="chkOOC<%=rs("ID")%>"></td>    
                        <%end if%>
                        <td>
                            <INPUT type="radio" id="optThis<%=rs("ID")%>" name="optScope<%=rs("ID")%>" value="1">&nbsp;<font size=2 face=verdana ID="ChangeThis" LANGUAGE=javascript onclick="return ChangeThis_onclick(<%=rs("ID")%>)" onmouseover="return ChangeThis_onmouseover()">Change this version of this product release only</font><BR>
		                    <INPUT type="radio" id="optFuture<%=rs("ID")%>" name="optScope<%=rs("ID")%>" checked value="2">&nbsp;<font size=2 face=verdana ID="ChangeDefault" LANGUAGE=javascript onclick="return ChangeDefault_onclick(<%=rs("ID")%>)" onmouseover="return ChangeDefault_onmouseover()">Change this version and all future versions of this product release</font><BR>
                            <INPUT type="hidden" id="txtScopeType<%=rs("ID")%>" name="txtNotesTag" value="2" >
                        </td>                   
                    </tr>
                <% rs.movenext
                loop
                rs.close %>                
                <tr>
		            <TD colspan="4" style="text-align:right;">
                        <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return Cancel_onclick()"  >
                        <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return btnOKClick()">		                
                     </TD>
                </TR>                
            </table> 
            <input type="hidden" id="txtReleaseInfo" name="txtReleaseInfo" value="" />                         
            <INPUT type="hidden" id="bTargetRelease" name="bTargetRelease" value=""/>
            <INPUT type="hidden" id="bTargeted" name="bTargeted" value="<%=bTargeted%>"/>
            <INPUT type="hidden" id="TargetNotes" name="TargetNotes" value=""/>
        </form>
    </body>
</HTML>