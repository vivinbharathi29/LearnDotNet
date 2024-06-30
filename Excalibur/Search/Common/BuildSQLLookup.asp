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

    function cmdOK_onclick(){
        cmdOK.disabled = true;
        cmdCancel.disabled = true;
        
        var i;
       
        if(txtReturnFormat.value=="")
            {
            var strOutput="";
            for (i=0;i<lstLookup.options.length;i++)
                if (lstLookup.options[i].selected)
                {
                    if (strOutput=="")
                       strOutput = "" + lstLookup.options[i].value + "";
                    else
                        strOutput = strOutput + ", " + lstLookup.options[i].value + "";
                };
		    window.parent.returnValue=' in (' + strOutput + ') ';
            }
        else
            {
            var strOutput= new Array();
            strOutput[0] = "";
            strOutput[1] = "";
            for (i=0;i<lstLookup.options.length;i++)
                if (lstLookup.options[i].selected)
                {
                    if (strOutput[0]=="")
                        {
                        strOutput[0] = lstLookup.options[i].value ;
                        strOutput[1] = lstLookup.options[i].text ;
                        }
                    else
                        {
                        strOutput[0] = strOutput[0] + "," + lstLookup.options[i].value;
                        strOutput[1] = strOutput[1] + "; " + lstLookup.options[i].text;
                        }
                };
            if (optAll.checked)
                {
                strOutput[0] = "";
                strOutput[1] = "All " + txtFieldName2.value;

                window.parent.returnValue=strOutput;
                }
            else
		        window.parent.returnValue= strOutput;
            }

		window.parent.close();
    }


    function cmdOK_onclick_old(){
        cmdOK.disabled = true;
        cmdCancel.disabled = true;
        
        var i;
       
        if(txtReturnFormat.value=="")
            {
            var strOutput="";
            for (i=0;i<lstLookup.options.length;i++)
                if (lstLookup.options[i].selected)
                {
                    if (strOutput=="")
                        strOutput = "" + lstLookup.options[i].value + "";
                    else
                        strOutput = strOutput + ", " + lstLookup.options[i].value + "";
                };
		    window.parent.returnValue=' in (' + strOutput + ') ';
            }
        else
            {
            var strOutput= new Array();
            strOutput[0] = "";
            strOutput[1] = "";
            for (i=0;i<lstLookup.options.length;i++)
                if (lstLookup.options[i].selected)
                {
                    if (strOutput[0]=="")
                        {
                        strOutput[0] = lstLookup.options[i].value ;
                        strOutput[1] = lstLookup.options[i].text ;
                        }
                    else
                        {
                        strOutput[0] = strOutput[0] + "," + lstLookup.options[i].value;
                        strOutput[1] = strOutput[1] + "; " + lstLookup.options[i].text;
                        }
                };
		    window.parent.returnValue= strOutput;
            }

		window.parent.close();
    }

    function optAll_onclick(){
        if (optAll.checked)
            lstLookup.disabled=true;
        else
            lstLookup.disabled=false;
    }

    function optChoose_onclick(){
        if (optChoose.checked)
            lstLookup.disabled=false;
        else
            lstLookup.disabled=true;
    }

    function window_onload(){
        lstLookup.focus();
    }
//-->
</SCRIPT>
<STYLE>
td{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
body{
    FONT-FAMILY: Verdana;
    FONT-SIZE:xx-small;
  }
</STYLE>
<title>Lookup</title>
</HEAD>


<BODY bgcolor="lavender" onload="window_onload();">

<%
    dim strTitle
    dim strItems
    dim strFieldName
    
    select case (lcase(trim(request("txtField"))))
    case "email"
        strTitle = "Select People"
        strFieldName = "People"
    case "group"
        strTitle = "Select Groups"
        strFieldName = "Groups"
    case "product"
        strTitle = "Select Products"
        strFieldName = "Products"
    case "productfamily"
        strTitle = "Select Product Families"
        strFieldName = "Product Families"
    case "subsystem"
        strTitle = "Select Sub Systems"
        strFieldName = "Sub Systems"
    case "state"
        strTitle = "Select States"
        strFieldName = "States"
    case "gatingmilestone"
        strTitle = "Select GatingMilestones"
        strFieldName = "Gating Milestones"
    case "frequency"
        strTitle = "Select Frequencies"
        strFieldName = "Frequencies"
    case "feature"
        strTitle = "Select Features"
        strFieldName = "Features"
    case "coreteam"
        strTitle = "Select Core Teams"
        strFieldName = "Core Teams"
    case "component"
        strTitle = "Select Components"
        strFieldName = "Components"
    case "affectedstate"
        strTitle = "Select Affected States"
        strFieldName = "Affected States"
    case "type"
        strTitle = "Lookup Component Types"
        strFieldName = "Component Types"
    case "developer"
        strTitle = "Lookup Developers"
        strFieldName = "Developers"
    case "componentpm"
        strTitle = "Lookup Component PMs"
        strFieldName = "Component PMs"
    case else
        strTitle = "Lookup Items"
        strFieldName = "Items"
    end select


%>

<table style="margin-left:10px;margin-right:0px;width:100%"><tr>
<td><font size=2 face=verdana><b><%=strTitle%></b></font><br></td></tr>
<tr>
<%

    if request("ReturnFormat") <> "" then
        response.write "<td><input id=""optAll"" name=""optScope"" type=""radio"" onclick=""optAll_onclick();""/> All " & strFieldName & "<br>"
        response.write "<input checked id=""optChoose"" name=""optScope"" type=""radio"" onclick=""optChoose_onclick();""/> Choose Specific " & strFieldName & "</td></tr>"
        response.write "<tr>"
    end if

    dim cn, rs, strSQL, ListArray, strValue, ValuePair
    dim strDivisionFilter
    dim blnNoCoreTeamLoaded

    strDivisionFilter = ""
    strDivisionFilter = " and (MobileConsumer=1 or MobileCommercial=1 or MobileFunctional=1) "

    strSavedItem = scrubsql(request("SelectedValues"))

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	        response.write "<td style=""width:100%"">"
            if request("ReturnFormat") <> "" then
                response.write "<select id=lstLookup name=lstLookup multiple style=""height:225px;width:100%"">"
            else
                response.write "<select id=lstLookup name=lstLookup multiple style=""height:250px;width:100%"">"
            end if
           
           select case (lcase(trim(request("txtField"))))
           case "type"
                if strSavedItem <> "" then
                    strSQl = "Select id, Name from HOUSIREPORT01.SIO.dbo.List_Type with (NOLOCK) where id in (" & strSavedItem & ") order by Name"
                    rs.open strSQL,cn
                    do while not rs.EOF
    	                response.write "<Option selected value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
                        rs.MoveNext
                    loop
                    rs.Close
                end if

                if strSavedItem <> "" then
                    strSQL = "Select ID, Name " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_type with (NOLOCK) " & _
                             "where active=1 " &_
                             "and id not in (" & strSavedItem & ") "
                else
                    strSQL = "Select ID, Name " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_type with (NOLOCK) " & _
                             "where active=1 " 
                end if
                strSQl = strSQL & " order by name;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    response.write "<Option value= """ & rs("ID") & """>" & rs("Name") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close

           case "developer"
                if strSavedItem <> "" then
                    rs.open "Select u.user_id, u.User_Name from HOUSIREPORT01.SIO.dbo.Users u with (NOLOCK) where user_id in (" & strSavedItem & ") order by u.User_Name",cn
                    do while not rs.EOF
    	                response.write "<Option selected value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
                        rs.MoveNext
                    loop
                    rs.Close
                end if

                if strSavedItem <> "" then
                    strSQL = "Select UserID, DisplayName, Email " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) " & _
                             "where active=1 " & _
                             "and developer=1 " & _
                             "and userid not in (" & strSavedItem & ") " & _
                            strDivisionFilter
                else
                    strSQL = "Select UserID, DisplayName, Email " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) " & _
                             "where active=1 " & _
                             "and developer=1 " & _
                            strDivisionFilter
                end if
                strSQl = strSQL & " order by Displayname;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    if request("ReturnFormat") = "" then
                        response.write "<Option value= ""'" & rs("Email") & "'"">" & rs("DisplayName") & "</OPTION>"
                    else
                        response.write "<Option value= """ & rs("UserID") & """>" & rs("DisplayName") & "</OPTION>"
                    end if
		            rs.MoveNext
	            loop
	            rs.Close

           case "componentpm"
                if strSavedItem <> "" then
                    rs.open "Select u.user_id, u.User_Name from HOUSIREPORT01.SIO.dbo.Users u with (NOLOCK) where user_id in (" & strSavedItem & ") order by u.User_Name",cn
                    do while not rs.EOF
    	                response.write "<Option selected value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
                        rs.MoveNext
                    loop
                    rs.Close
                end if

                if strSavedItem <> "" then
                    strSQL = "Select UserID, DisplayName, Email " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) " & _
                             "where active=1 " & _
                             "and componentpm=1 " & _
                             "and userid not in (" & strSavedItem & ") " & _
                            strDivisionFilter
                else
                    strSQL = "Select UserID, DisplayName, Email " & _
                             "FROM HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) " & _
                             "where active=1 " & _
                             "and componentpm=1 " & _
                            strDivisionFilter
                end if

                strSQl = strSQL & " order by Displayname;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    if request("ReturnFormat") = "" then
                        response.write "<Option value= ""'" & rs("Email") & "'"">" & rs("DisplayName") & "</OPTION>"
                    else
                        response.write "<Option value= """ & rs("UserID") & """>" & rs("DisplayName") & "</OPTION>"
                    end if
		            rs.MoveNext
	            loop
	            rs.Close
           case "email"
 
                if strSavedItem <> "" then
                    rs.open "Select u.user_id, u.User_Name from HOUSIREPORT01.SIO.dbo.Users u with (NOLOCK) where user_id in (" & strSavedItem & ") order by u.User_Name",cn
                    do while not rs.EOF
    	                response.write "<Option selected value=""" & rs("user_id") & """>" & rs("User_Name") & "</OPTION>"
                        rs.MoveNext
                    loop
                    rs.Close
                end if
'                strSQL = "Select distinct u.user_id, u.User_Name,u.Email_Address, user_name + ' (' + replace(replace(replace(replace(replace(SUBSTRING(u.Email_Address,CHARINDEX('@',u.email_address)+1,LEN(u.email_address)),'.cn',''),'.com','') ,'.tw','') ,'cn.','') ,'tw.','')  + ')' as DistinctUserName " & _
'                         "from HOUSIREPORT01.SIO.dbo.Observation o with (NOLOCK), HOUSIREPORT01.SIO.dbo.Users u with (NOLOCK) " & _
'                         "where u.User_ID = o.Owner_User_ID " & _
'                         "and o.Status_Name <> 'Closed' " & _
'                         "and u.active_flg = 1 " & _
'                         "and Division_ID = 6 " & _
'                         "order by u.User_Name"

                strSQL = "Select UserID, DisplayName, Email " & _
                         "FROM HOUSIREPORT01.SIO.dbo.list_actor with (NOLOCK) " & _
                         "where active=1 " & _
                        strDivisionFilter
                strSQl = strSQL & " order by Displayname;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    if request("ReturnFormat") = "" then
                        response.write "<Option value= ""'" & rs("Email") & "'"">" & rs("DisplayName") & "</OPTION>"
                    else
                        response.write "<Option value= """ & rs("UserID") & """>" & rs("DisplayName") & "</OPTION>"
                    end if
		            rs.MoveNext
	            loop
	            rs.Close
            case "group"
'                strSQL = "SELECT [XLS_Org_ID] as ID,[XLS_Org_Name] as OwnerGroup " & _
'                         "FROM HOUSIREPORT01.SIO.dbo.[vWorkgroupPrimary]  with (NOLOCK) " & _
'                         "where division_name = 'Mobile' " & _
'                         "and [Active_Flg] = 1 " & _
'                         "and xls_org_id not in (5359,5351,5269,5130,5131) " & _
'                         "order by [XLS_Org_Name]"
                strSQL = "Select GroupID as ID, name " & _
                         "from HOUSIREPORT01.SIO.dbo.list_group with (NOLOCK) " & _
                        "where active=1 " & _
                        strDivisionFilter

                strSQL = strSQl & " order by name;"
	            rs.Open strSQL,cn,adOpenForwardOnly

	            do while not rs.EOF
                    response.write "<Option value= ""'" & rs("Name") & "'"">" & rs("Name") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close
            case "product"

        	    strSQL = "SELECT Name " & _
                         "FROM HOUSIREPORT01.SIO.dbo.[List_Product] with (NOLOCK) " & _
	                     "Where active=1 " & _
	                    strDivisionFilter
    
                strSql = strSql & " order by Name;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    response.write "<Option value=""'" & rs("Name") & "'"">" & rs("Name") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close
            case "productfamily"

        	    strSQL = "SELECT Distinct FamilyName " & _
                         "FROM HOUSIREPORT01.SIO.dbo.[List_Product] with (NOLOCK) " & _
	                     "Where active=1 " & _
	                    strDivisionFilter
    
                strSql = strSql & " order by FamilyName;"

	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    response.write "<Option value=""'" & rs("FamilyName") & "'"">" & rs("FamilyName") & "</OPTION>"
    	            rs.MoveNext
	            loop
	            rs.Close
            case "subsystem"

                if strSavedItem <> "" then
                    rs.open "Select id, Name from HOUSIREPORT01.SIO.dbo.list_subsystem with (NOLOCK) where id in (" & strSavedItem & ") order by name",cn
                    do while not rs.EOF
      	                response.write "<Option selected value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
                        rs.MoveNext
                    loop
                    rs.Close
                end if

                if strSavedItem <> "" then
                    strSQL = " Select id, name, active " & _
                            "from HOUSIREPORT01.SIO.dbo.List_SubSystem with (NOLOCK) " & _
                            "where active =1 " & _
                            "and id not in (" & strSavedItem & ") " & _
	                        strDivisionFilter
                else
                    strSQL = " Select id, name, active " & _
                            "from HOUSIREPORT01.SIO.dbo.List_SubSystem with (NOLOCK) " & _
                            "where active =1 " & _
	                        strDivisionFilter
                end if    
            strSql = strSql & " order by Name;"
	            rs.Open strSQL,cn,adOpenForwardOnly
                strItems = ""
	            do while not rs.EOF
                    if request("ReturnFormat") = "" then
                        strItems = strItems & "<Option value=""'" & rs("Name") & "'"">" & rs("name") & "</OPTION>"
                    else
                        strItems = strItems & "<Option value=""" & rs("ID") & """>" & rs("name") & "</OPTION>"
                    end if
                    rs.MoveNext
	            loop
	            rs.Close
                response.write strItems
            case "state"
                 strSQl = "Select Name " & _
                        "from HOUSIREPORT01.SIO.dbo.list_state with (NOLOCK) " & _
                         "where active=1 " & _
                        "order by Name;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    response.write "<Option value=""'" & rs("Name") & "'"">" & rs("name") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close
            case "gatingmilestone"

            strSQL = " Select Name " & _
                     "from HOUSIREPORT01.SIO.dbo.List_GatingMilestone with (NOLOCK) " & _
                     "where active=1 " & _
	                 strDivisionFilter
    
            strSql = strSql & " order by Name;"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    if trim(rs("Name")) = "" then
                        response.write "<Option value=""'" & rs("Name") & "'"">Not Specified</OPTION>"
                    else
                        response.write "<Option value=""'" & rs("Name") & "'"">" & rs("Name") & "</OPTION>"
                    end if
		            rs.MoveNext
	            loop
	            rs.Close
            case "frequency"
                ListArray = split("621|Always: 100%,622|Intermittent: <1%,623|Seen Once,780|Intermittent: 1-5%,781|Intermittent: 5-25%,782|Intermittent: 25-99%,783|Single Unit Failure,10247|Related Case",",")

                for each strValue in ListArray
                    ValuePair = split(strValue,"|")
                    response.write "<Option value=""'" & ValuePair(1) & "'"">" & ValuePair(1) & "</OPTION>"
			    next
            case "feature"
                strSQL = "Select Name as Feature " & _
	                     "from prs.dbo.DeliverableFeatures with (NOLOCK) " & _
	                     "where Active=1 " & _
	                     "order by Name"
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    response.write "<Option value=""'" & rs("Feature")  & "'"">" & rs("Feature") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close
            case "coreteam"
                if strSavedItem <> "" then
                    blnNoCoreTeamLoaded = false
                    rs.open "Select ID, Name from PRS.dbo.DeliverableCoreTeam with (NOLOCK) where ID in (" & strSavedItem & ") order by Name",cn
                    do while not rs.EOF
        	            if rs("ID") = 0 then
                            response.write "<Option selected value= ""0"">No Core Team Assigned</OPTION>"
                            blnNoCoreTeamLoaded = true
                        else
                            response.write "<Option selected value=""" & rs("id") & """>" & rs("Name") & "</OPTION>"
                        end if
                        rs.MoveNext
                    loop
                    rs.Close
                end if
                if strSavedItem <> "" then
                    strSQL = "Select ID, Name " & _
                             "from PRS.dbo.DeliverableCoreTeam with (NOLOCK) " & _
                             "where ID not in (" & strSavedItem & ") " & _
                             "and active=1 " & _
                             "order by Name;"
                else
                    strSQL = "Select ID, Name " & _
                             "from PRS.dbo.DeliverableCoreTeam with (NOLOCK) " & _
                             "where active=1 " & _
                             "order by Name;"
                end if
	            rs.Open strSQL,cn,adOpenForwardOnly
	            do while not rs.EOF
                    if rs("ID") <> 0 then
                        if request("ReturnFormat") = "" then
          	                response.write "<Option value= ""'" & rs("Name") & "'"">" & rs("Name") & "</OPTION>"
                        else
          	                response.write "<Option value= """ & rs("id") & """>" & rs("Name") & "</OPTION>"
                        end if
                    end if
                    rs.MoveNext
	            loop
	            rs.Close
                if not blnNoCoreTeamLoaded then
                    response.write "<Option value= ""0"">No Core Team Assigned</OPTION>"
                end if
            case "component"

                strSQL = " Select name " & _
                         "from HOUSIREPORT01.SIO.dbo.List_Component with (NOLOCK) " & _
                         "where active=1 " & _
	                     strDivisionFilter 
                         
                strSql = strSql & " order by Name;"
	            
                rs.Open strSQL,cn,adOpenForwardOnly
	            
                do while not rs.EOF
                    response.write "<Option value=""'" & replace(replace(rs("Name"),"""",""""""),"'","''") & "'"">" & rs("name") & "</OPTION>"
		            rs.MoveNext
	            loop
	            rs.Close        
            case "affectedstate"
                ListArray = split("Affected,Test Required,Waiver Requested,Untested,Deferred,Fix Implemented,Fix Verified,Module/Feature Constraint,Module/Feature Dropped,Not Affected,Waiver Approved,Will Not Fix",",")

                for each strValue in ListArray
                    response.write "<Option value=""'" & strValue & "'"">" & strValue & "</Option>"
			    next
            end select
            response.write "</select></td>"




    set rs = nothing
    cn.Close
    set cn = nothing


    	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i

		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 
%>
</tr>
<tr><td align=right><hr><input id="cmdOK" type="button" value="Ok" onclick="javascript:cmdOK_onclick();"/><input id="cmdCancel" type="button" value="Cancel" onclick="javascript: window.close();"/></td></tr>
</table>
<input style="display:none" id="txtReturnFormat" type="text" value="<%=server.HTMLEncode(request("ReturnFormat"))%>">
<input style="display:none" id="txtFieldName" type="text" value="<%=server.HTMLEncode((lcase(trim(request("txtField")))))%>">
<input style="display:none" id="txtFieldName2" type="text" value="<%=strFieldName%>">


</BODY>
</HTML>




