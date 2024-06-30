<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	if request("cboFormat")= "1" then
		Response.ContentType = "application/vnd.ms-excel"
	end if	  
	%>

<HTML>
<HEAD>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        window.parent.ProfileSaved(txtType.value, txtID.value, txtResults.value, txtError.value);
    }
//-->
</SCRIPT>
<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  	vertical-align:top;
  }
thead
{
    background-color:beige;
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
h1{
    FONT-FAMILY: Verdana;
    FONT-SIZE:x-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
A:link
{
    COLOR: Blue;
}
A:visited
{
    COLOR: Blue;
}
A:hover
{
    COLOR: red;
}  
</STYLE>
</HEAD>


<BODY onload=window_onload();>




<%
    dim strSuccess
    dim strErrorMessage
    dim cm
    dim rs
    dim cn
    dim p

    on error resume next
    strSuccess = "1"
    strErrorMessage = ""

    strProfileData = ""

    for each strfield in request.Form
        if trim(request(strField)) <> "" then
            if trim(strField) <> "txtNewProfileName" and trim(strField) <> "txtPageLayout"  and trim(strField) <> "txtProfileUpdateType"  and trim(strField) <> "txtProfileUpdateID"  and trim(strField) <> "txtNewReportFormat" and trim(strField) <> "txtUserID"  and trim(strField) <> "txtNewTodayLink" then
                strProfileData = strProfileData & strField & "=" & server.URLEncode(request(strField)) & "&"
            end if
        end if
    next
    if strProfileData <> "" then
        strProfileData = left(strProfileData,len(strProfileData)-1)
    end if

    Response.write "Saving Profile..."
    response.write  "<BR><BR>____________<BR><BR>" & strProfileData
    
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

    if trim(request("txtProfileUpdateType")) = "3" then 'Renaming
	    set cm = server.CreateObject("ADODB.Command")
	    cm.ActiveConnection = cn

	    cm.CommandText = "spRenameProfile"
	    cm.CommandType =  &H0004
		
	    Set p = cm.CreateParameter("@ID", 3,  &H0001)
	    p.value =clng(request("txtProfileUpdateID"))
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@ProfileName", 200, &H0001, 120)
	    p.value = left(request("txtNewProfileName"),120)
	    cm.Parameters.Append p
	
	    cm.Execute
		set cm=nothing

		if cn.Errors.count > 0 or  err.number <> 0 then
			strSuccess = "0"
            strErrorMessage = err.Description
		end if	

    elseif trim(request("txtProfileUpdateType")) = "4" then 'Deleting
	    set cm = server.CreateObject("ADODB.Command")
	    cm.ActiveConnection = cn

        cm.CommandText = "spDeleteProfile"
        cm.CommandType =  &H0004
		
	    Set p = cm.CreateParameter("@ID", 3,  &H0001)
	    p.value =clng(request("txtProfileUpdateID"))
	    cm.Parameters.Append p

	    cm.Execute
		set cm=nothing

		if cn.Errors.count > 0 or  err.number <> 0 then
			strSuccess = "0"
            strErrorMessage = err.Description
		end if	
    elseif trim(request("txtProfileUpdateType")) = "5" then 'Remove
        response.write ".." & clng(request("txtProfileUpdateID")) & ".."
        cn.execute "spRemoveSharedProfile " & clng(request("txtProfileUpdateID"))

		if cn.Errors.count > 0 or  err.number <> 0 then
			strSuccess = "0"
            strErrorMessage = err.Description
		end if	

    elseif trim(request("txtProfileUpdateType")) = "1" then
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		cm.CommandText = "spAddReportProfile"	

		Set p = cm.CreateParameter("@EmployeeID", 3,  &H0001)
		p.Value = clng(request("txtUserID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ProfileName", 200,  &H0001,120)
		p.Value = left(request("txtNewProfileName"),120)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ProfileType", 3,  &H0001)
		if request("txtProfileType") <> "" then
            p.Value = clng(request("txtProfileType"))
        else
            p.Value = 6
        end if
		cm.Parameters.Append p
            
		Set p = cm.CreateParameter("@PageLayout", 200,  &H0001,2147483647)
		p.Value = request("txtPageLayout")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SelectedFilters", 200,  &H0001,2147483647)
        if request("txtFieldFilters") <> "" or trim(request("txtProfileType"))="7" then
    		p.Value = request("txtFieldFilters")
	    else
    	    p.Value = strProfileData
        end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@TodayPageLink", 3,  &H0001)
	    if request("txtNewTodayLink") = "" then
    	    p.Value = 0
        else
    	    p.Value = clng(request("txtNewTodayLink") )
        end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ReportFormat", 3,  &H0001)
	    if request("txtNewReportFormat") = "" then
    	    p.Value = 0
        else
    	    p.Value = clng(request("txtNewReportFormat") )
        end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@NewID", 3,  &H0002)
		cm.Parameters.Append p

		cm.Execute rowschanged
		NewID = cm("@NewID")
		set cm=nothing

		if cn.Errors.count > 0 or  err.number <> 0 then
			strSuccess = "0"
            strErrorMessage = err.Description
        else
            strSuccess = NewID
		end if	

    else
		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		cm.CommandText = "spUpdateReportProfile"	

		Set p = cm.CreateParameter("ID", 3,  &H0001)
		p.Value = clng(request("txtProfileUpdateID"))
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ProfileName", 200,  &H0001,120)
		p.Value = left(request("txtNewProfileName"),120)
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PageLayout", 200,  &H0001,2147483647)
		p.Value = request("txtPageLayout")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SelectedFilters", 200,  &H0001,2147483647)
        if request("txtFieldFilters") <> "" then
    		p.Value = request("txtFieldFilters")
	    else
    	    p.Value = strProfileData
        end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@TodayPageLink", 3,  &H0001)
	    if request("txtNewTodayLink") = "" then
    	    p.Value = 0
        else
    	    p.Value = clng(request("txtNewTodayLink") )
        end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ReportFormat", 3,  &H0001)
	    if request("txtNewReportFormat") = "" then
    	    p.Value = 0
        else
    	    p.Value = clng(request("txtNewReportFormat") )
        end if
		cm.Parameters.Append p

		cm.Execute rowschanged
		set cm=nothing

		if cn.Errors.count > 0 or  err.number <> 0 then
			strSuccess = "0"
            strErrorMessage = err.Description
		end if	

    end if

    
    
    
    set rs=nothing
    cn.Close
    set cn = nothing

    if err.number <> 0 then
        strSuccess = "0"
        strErrorMessage = err.Description
    end if

%>
<input id="txtResults" type="text" value="<%=strSuccess%>">
<input id="txtError" type="text" value="<%=strErrorMessage%>">
<input id="txtType" type="text" value="<%=request("txtProfileUpdateType")%>">
<input id="txtID" type="text" value="<%=request("txtProfileUpdateID")%>">

</BODY>
</HTML>




