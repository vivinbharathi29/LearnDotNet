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
    dim bSaveRelease, returnValue, bSaved
    dim pvID, ProdBrandCountryLocalizationID, ProdBrandCountryID, LocalizationID
    Dim Security, m_UserFullName
   	Set Security = New ExcaliburSecurity
        m_UserFullName = Security.CurrentUserFullName()
   
    ProdBrandCountryLocalizationID = 0
    if Request.QueryString("ProdBrandCountryLocalizationID") <> "" then
	    ProdBrandCountryLocalizationID = clng(Request.QueryString("ProdBrandCountryLocalizationID"))
    end if
    if Request.QueryString("PVID") <> "" then
        pvID = clng(Request.QueryString("PVID"))
    end if
        
    ProdBrandCountryID=0
    if Request.QueryString("ProdBrandCountryID") <> "" then
	    ProdBrandCountryID = clng(Request.QueryString("ProdBrandCountryID"))
    end if
    LocalizationID = 0 
    if Request.QueryString("LocalizationID") <> "" then
	    LocalizationID = clng(Request.QueryString("LocalizationID"))
    end if
   
    Set rs = Server.CreateObject("ADODB.RecordSet")
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	
   ' Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

      
    if len(Request.form("bSaveRelease")) > 0 then
		bSaveRelease = true
	end if

    bSaved = ""
        
	if bSaveRelease then
    'section to execute when save/OK button is clicked
  
         Set dw = New DataWrapper
           Set cmd = dw.CreateCommAndSP(cn, "usp_ProductBrandLocalization_SaveReleases")
                    dw.CreateParameter cmd, "@p_ProdBrandCountryLocalizationID", adInteger, adParamInput, 8, ProdBrandCountryLocalizationID    
                    dw.CreateParameter cmd, "@p_ProdBrandCountryID", adInteger, adParamInput, 8, ProdBrandCountryID
                    dw.CreateParameter cmd, "@p_LocalizationID", adInteger, adParamInput, 8, LocalizationID 
                    dw.CreateParameter cmd, "@p_chrReleaseIDs", adVarchar, adParamInput, 256, Request.Form("chkRelease")
                    dw.CreateParameter cmd, "@p_chrUserName", adVarchar, adParamInput, 50, m_UserFullName        
               returnValue = dw.ExecuteNonQuery(cmd)
            bSaved = "1"
    end if
   
	
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "usp_ProductBrandLocalization_ViewReleases"
	
    Set p = cm.CreateParameter("@p_pvID", 3, &H0001)
	p.Value = pvID
	cm.Parameters.Append p

	Set p = cm.CreateParameter("@p_ProdBrandCountryLocalizationID", 3, &H0001)
	p.Value = ProdBrandCountryLocalizationID
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
            function btnOKClick() {
                if (!$('input:checkbox').is(':checked'))
                {
                    alert("Please select at least one release");
                    return false;
                }
                //disable the OK button once it is clicked
                document.getElementById("cmdOK").disabled = true;
                document.thisform.action = "LocalizationReleaseEdit.asp?ProdBrandCountryLocalizationID=<%=cstr(ProdBrandCountryLocalizationID)%>&PVID=<%=cstr(pvID)%>&ProdBrandCountryID=<%=cstr(ProdBrandCountryID)%>&LocalizationID=<%=cstr(LocalizationID)%>";
		        document.thisform.bSaveRelease.value = "1";
		        document.thisform.submit();
	       
            }
            
            function Cancel_onclick() {
   
                window.parent.ClosePropertiesDialog();
    
            }

            function body_onload()
            {
                var strbSaved = document.getElementById("bSaved").value;
                if (strbSaved != "")
                {
                    window.parent.ClosePropertiesDialog(true);                
                }    
    
            }
        </script>
    </head>

    <BODY id=bdy onload="return body_onload();" style="background: ivory;">
        <form name=thisform method=post>	        
	        <table style="width:98%;">
                <tr> <td style="width:20px;">Releases:<font color=red>*</font> &nbsp;</td>
                    <td style="text-align:left;">
                        <% 
                                 
                        dim strName
                        'rs.open "usp_ProductBrandLocalization_ViewReleases " & clng(request("PVID")) & "," & clng(request("ProdBrandCountryLocalizationID")), cn, adOpenForwardOnly
                        do while not rs.eof
                            strname = rs("Name")   

                            if trim(rs("bSelected")) = "1" then
                                response.write "<input checked id=""chkRelease"" name=""chkRelease"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """  > " & strname & "&nbsp;"
                                    
                            else
                                response.write "<input id=""chkRelease"" name=""chkRelease"" type=""checkbox"" value=""" & rs("ReleaseID") & """ ReleaseID=""" & rs("ReleaseID") &  """   > " & strname & "&nbsp;"
                            end if

                            rs.movenext
                        loop

                        rs.close
                            %>
            	     
                     </td>          

                </tr>
                <tr> <td><font color=red>*</font>Required field</td>
		            <TD style="text-align:right;">
                        <INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return btnOKClick()">
		                <INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return Cancel_onclick()"  >
                     </TD>
                </TR>
            </table>         

            <INPUT type="hidden" id="bSaveRelease" name="bSaveRelease" value=""/>
            <INPUT type="hidden" id="bSaved" name="bSaved" value="<%=bSaved%>"/>
        </form>
    </body>
</HTML>