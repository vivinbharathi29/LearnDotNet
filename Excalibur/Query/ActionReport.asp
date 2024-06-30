<%@ Language=VBScript %>


	<%
	
	if request("cboFormat")= 1 then
		Response.ContentType = "application/vnd.ms-excel"
	elseif request("cboFormat")= 2 then
		Response.ContentType = "application/msword"
	end if    
    
	%>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<Title>Action Item Query Results</title>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

<!-- #include file = "../_ScriptLibrary/sort.js" -->

function trim( varText)
    {
    var i = 0;
    var j = varText.length - 1;
    
	for( i = 0; i < varText.length; i++ )
		{
		if( varText.substr( i, 1 ) != " " &&
			varText.substr( i, 1 ) != "\t")
		break;
		}
		
   
	for( j = varText.length - 1; j >= 0; j-- )
		{
		if( varText.substr( j, 1 ) != " " &&
			varText.substr( j, 1 ) != "\t")
		break;
		}

    if( i <= j )
		return( varText.substr( i, (j+1)-i ) );
	else
		return("");
    }

function window_onload() {
	//lblLoad.style.display = "none";
}


function row_onmouseover() {
	window.event.srcElement.parentElement.style.cursor = "hand"
	window.event.srcElement.parentElement.style.color = "red"

}
function row_onmouseout() {
	window.event.srcElement.parentElement.style.color = "black"

}

function row_onclick() {
	var strID;
	var strResult;
	strResult = window.showModalDialog("../mobilese/today/action.asp?" + window.event.srcElement.parentElement.className,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
//	if (typeof(strResult) != "undefined")
//		{
//			window.location.reload(true);
//		}

}
function DisplayAction(strID, strType) {
	var strResult;
	strResult = window.showModalDialog("../mobilese/today/action.asp?ID=" + strID + "&Type=" + strType,"","dialogWidth:655px;dialogHeight:650px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No") 
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
TABLE
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana, Tahoma, Arial
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
<BODY LANGUAGE=javascript onload="return window_onload()">

<H3><font face=verdana><%=request("txtTitle")%></font></H3>
<!--<span ID=lblLoad><font size=2 face=verdana>Loading.  Please wait...</font></span>-->

<%

	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union",  "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 



	dim strSQL
	dim cn
	dim strProducts
	dim strOwners
	dim strApprovers
	dim strStatus
	dim rs 
	dim rs2
	dim strType
	dim strRange1
	dim strRange2
	dim strApproverStatus
	dim strApproverComments
	dim strBusiness
	dim strProductType
	dim LineCount
	dim strBaseSQL
    dim strAVRequired
    dim strQualificationRequired
    dim strLegacyProducts
    dim strCombineProducts
    dim PulsarProductIDs
    dim IDs
    dim strProductReleases
    dim strPulsarProductReleases

	LineCount=0
	strBaseSQL=""
    strPulsarProductReleases = ""

    strPulsarProductReleases = request("lstProductsPulsar")

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.CommandTimeout = 20
	cn.Open


	set rs = server.CreateObject("ADODB.recordset")

	strSQl = "SELECT d.PendingImplementation, d.ecndate, v.Division, d.Type, v.typeid, d.AvailableNotes, d.priority, d.AffectsCustomers, cr.name as CoreTeamRep, d.Submitter, e.Name as Owner, f.name as ProductFamily, v.DOTSName as Product, d.ID, d.Status, d.TargetDate, d.Summary, d.Description, d.Details, d.consumer, d.commercial, d.smb, d.Approvals, d.Justification, d.Created, d.ActualDate, d.ReleaseNotification, d.Resolution, d.Actions, d.ZsrpReadyTargetDt, d.ZsrpReadyActualDt, d.ZsrpRequired, d.AVRequired, d.QualificationRequired,d.ProductVersionRelease " & _	
			 "FROM DeliverableIssues d with (NOLOCK), Employee e with (NOLOCK), productFamily f with (NOLOCK), ProductVersion v with (NOLOCK), CoreTeamRep cr with (NOLOCK) " & _
			 "Where d.OwnerID = e.ID " & _
			 "and f.Id = v.ProductFamilyID " & _
			 "and cr.Id = d.CoreTeamRep " & _
			 "and v.ID = d.ProductVersionID "
	strBaseSQL = strSQL
	if request("ID") <> "" then
		strSQl = strSQl & " and d.ID= " & request("ID") & " " 
	else
    
    strProducts = request("lstProducts")


    '*******Process Product Groups 
	    if request("lstProductGroups") <> "" then
		    dim ProductGroupsArray
		    dim ProductGroupArray
		    dim strProductGroup
		    dim lastProductGroup
		    dim strProductGroupFilter
		    dim strCycleList
		    ProductGroupsArray = split(request("lstProductGroups"),",")
		    lastProductGroup = 0
		    strProductGroupFilter = ""
		    strCycleList = ""

		    for each strProductGroup in ProductGroupsArray
			    if instr(strProductGroup,":")>0 then
				    ProductGroupArray = split(strProductGroup,":")
				    if trim(lastproductgroup) <> "0" and trim(ProductGroupArray(0)) <> "2" and lastproductgroup <> trim(ProductGroupArray(0)) then
					    strProductGroupFilter = strProductGroupFilter & " ) and  "
				    end if
				    if trim(lastproductgroup) <> trim(ProductGroupArray(0)) then
					    if trim(ProductGroupArray(0)) = "1" then
						    strProductGroupFilter = strProductGroupFilter & " ( partnerid = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    elseif trim(ProductGroupArray(0)) = "2" then
						    strCycleList = strCycleList & "," & clng(ProductGroupArray(1))
					    elseif trim(ProductGroupArray(0)) = "3" then
						    strProductGroupFilter = strProductGroupFilter & " ( devcenter = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    elseif trim(ProductGroupArray(0)) = "4" then
						    strProductGroupFilter = strProductGroupFilter & " ( productstatusid = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    end if
				    else
					    if trim(ProductGroupArray(0)) = "1" then
						    strProductGroupFilter = strProductGroupFilter & " or partnerid = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    elseif trim(ProductGroupArray(0)) = "2" then
						    strCycleList = strCycleList & "," & clng(ProductGroupArray(1))
					    elseif trim(ProductGroupArray(0)) = "3" then
						    strProductGroupFilter = strProductGroupFilter & " or devcenter = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    elseif trim(ProductGroupArray(0)) = "4" then
						    strProductGroupFilter = strProductGroupFilter & " or productstatusid = " & trim(ProductGroupArray(1))
						    lastproductgroup = trim(ProductGroupArray(0))
					    end if
				    end if
			    end if
		    next
		    if strProductGroupFilter <> "" then
			    strGroupSQl = strGroupSQL & " and ( " & ScrubSQL(strProductGroupFilter) &  ") ) "
		    end if
		    if strCycleList <> "" then
			    strGroupSQl = strGroupSQL & " and id in (Select ProductVersionid from product_program with (NOLOCK) where programid in (" & mid(strCycleList,2) &  ")) "
		    end if
		    if strGroupSQl <> "" then
		        strGroupSQl = mid(strGroupSQL,5)
		        rs.open "Select ID from productversion with (NOLOCK) where " & strgroupSQL,cn
		        do while not rs.eof
	                strProducts = strProducts & ", " & rs("ID") 
		            rs.movenext
		        loop
		        rs.close    
		    end if
		    if strProducts = ""  then 
	            strProducts = "0"
	        elseif left(strproducts,2) = ", " then
	            strproducts = mid(strproducts,3)
	        end if
		    
'		    if strProductGroupFilter <> "" then
'			    strSQl = strSQL & " and ( " & ScrubSQL(strProductGroupFilter) &  ") ) "
'		    end if
'		    if strCycleList <> "" then
'			    strSQl = strSQL & " and v.id in (Select ProductVersionid from product_program where programid in (" & mid(strCycleList,2) &  ")) "
'		    end if
	    end if
    '*******End Product Groups


		if left(strProducts,1) = "," then
			strProducts = mid(strProducts,2)
		end if
		strOwners = request("lstOwners")
		if left(strOwners,1) = "," then
			strOwners = mid(strOwners,2)
		end if
	
		strApprovers = request("lstApprovers")
		if left(strApprovers,1) = "," then
			strApprovers = mid(strApprovers,2)
		end if
	
		if request("cboApproverStatus") <> "" and request("cboApproverStatus") <> "0" then
			strApproverStatus = " and Status = " & scrubsql(request("cboApproverStatus")) & " "
		else
			strApproverStatus = ""
		end if

        If request("txtApproverComments") <> "" Then
            strApproverComments = " AND Comments LIKE '%" & scrubsql(request("txtApproverComments")) & "%'"
        Else
            strApproverComments = ""
        End If

		strStatus = request("lstStatus")
		if left(strStatus,1) = "," then
			strStatus = mid(strStatus,2)
		end if
		strType = request("lstType")
		if left(strType,1) = "," then
			strType = mid(strType,2)
		end if
	
    'process selected pulsar products
        dim strPulsarproducts 
        dim streachPulsarproduct
        dim pulsarproductID
        dim pulsarRelease
        dim strpulsarprodCondition
        dim strpulsarprodSub
        strPulsarproducts = strPulsarProductReleases '"1910_Aug 2017, 1910_Jun 2017" 'replace this with the value u passed in from parent asp page
        dim PulsarproductsArray
        dim pulsarproductcnt
        strpulsarprodCondition = ""
        PulsarproductsArray = split(strPulsarproducts,",")
        
        for pulsarproductcnt = lbound(PulsarproductsArray) to ubound(PulsarproductsArray)
                streachPulsarproduct = PulsarproductsArray(pulsarproductcnt)
                pulsarproductID = left(streachPulsarproduct, instr(streachPulsarproduct, "_")-1)
                pulsarRelease = mid(streachPulsarproduct, instr(streachPulsarproduct, "_") + 1, len(streachPulsarproduct) )
                strpulsarprodSub = " or ( d.productversionID = " & pulsarproductID & " and (charindex(REplace('" + pulsarRelease + "','NPI ',''),d.productversionRelease) > 0 or d.productversionRelease is null or d.productversionRelease =''))"           
               strpulsarprodCondition = strpulsarprodCondition  & strpulsarprodSub
         next     
    
        if strProducts <> "" then
            if strProducts = "0" or instr(strProducts,", 0,") or left(trim(strProducts),2) = "0,"  or right(trim(strProducts),3) = ", 0" then
                if strpulsarprodCondition <>"" then
                    strSQl = strSQL & " and ( v.Sustaining = 1 or d.CoreTeamRep = 12 or d.ProductVersionID in ( " & scrubsql(strProducts) &  " )   " & strpulsarprodCondition & " ) "
                ELSE
                    strSQl = strSQL & " and ( v.Sustaining = 1 or d.CoreTeamRep = 12 or d.ProductVersionID in ( " & scrubsql(strProducts) &  " ) ) "
                END IF 
           else
                if strpulsarprodCondition <>"" then
                    strSQl = strSQL & " and (d.ProductVersionID in ( " & scrubsql(strProducts) &  " )   " & strpulsarprodCondition & " ) "
                ELSE
                    strSQl = strSQL & " and d.ProductVersionID in ( " & scrubsql(strProducts) &  " ) "
                END IF                                                     
          end if
   
       else
            if strpulsarprodCondition <>"" then
                  strSQl = strSQL & " and (1<>1 " & strpulsarprodCondition & " ) "
            end if
      end if 
         
     if strOwners <> "" and request("lstApprovers") <> "" then
			strSQl = strSQL & " and ( d.OwnerID in ( " & scrubsql(strOwners) &  " ) or  d.Id in (Select ActionID from ActionApproval with (NOLOCK) where ApproverID in ( " & scrubsql(strApprovers) & ") " & scrubsql(strApproverStatus)  & scrubsql(strApproverComments) & "  ) )"
		elseif strOwners <> "" then
			strSQl = strSQL & " and d.OwnerID in ( " & scrubsql(strOwners) &  " ) "
		elseif strApprovers <> "" then
			strSQl = strSQL & " and d.Id in (Select ActionID from ActionApproval with (NOLOCK) where ApproverID in ( " & scrubsql(strApprovers) & ") " & strApproverStatus  & strApproverComments & " ) "
		end if

        if trim(request("chkWorkingList")) = "1" then
			strSQl = strSQL & " and d.pendingimplementation = 1 "
        end if
		
		if strType <> "" then
			strSQl = strSQL & " and d.Type in ( " & scrubsql(strType) &  " ) "
		end if
		
		if request("lstSubmitter") <> "" then
			dim SubmitterArray
			SubmitterArray = split(scrubsql(request("lstSubmitter")),",")
			strSQl = strSQL & " and ( " 
			for i = lbound(SubmitterArray) to ubound(SubmitterArray)
				if i <> lbound(SubmitterArray) then
					strSQl = strSQL & " or "
				end if
				strSQl = strSQL & "d.Submitter='" & trim(replace(SubmitterArray(i),"|",",")) & "'"
			next
			strSQl = strSQL & " ) "
		end if	
		
		if request("WorkingList") = "1" then
			strSQl = strSQL & " and d.PendingImplementation = 1  "
		end if
		
		if request("cboCategory") = "1" then
			strSQl = strSQL & " and SKUChange = 1 " 
		elseif request("cboCategory") = "2" then
			strSQl = strSQL & " and ImageChange = 1 " 
		elseif request("cboCategory") = "3" then
			strSQl = strSQL & " and ReqChange = 1 " 
		elseif request("cboCategory") = "4" then
			strSQl = strSQL & " and OtherChange = 1 " 
		elseif request("cboCategory") = "5" then
			strSQl = strSQL & " and DocChange = 1 " 
		elseif request("cboCategory") = "6" then
			strSQl = strSQL & " and CommodityChange = 1 " 
		end if
	
		if strStatus <> "" then
				strSQl = strSQL & " and d.Status in ( " & scrubsql(strStatus) &  " ) "
		end if

		if trim(request("txtNumbers")) <> "" then
				strSQl = strSQL & " and d.ID in ( " & scrubsql(request("txtNumbers")) &  " ) "
		end if
		
		dim strSearch
		dim strTemp
		strSearch = scrubsql(replace(replace(replace(replace(request("txtSearch"),"""",""),"'",""),"%",""),"*",""))
		if request("txtSearch") <> "" then
			strTemp = ""
			if request("chkSummarySearch") = "on" then
				strTemp = strTemp & " or Summary Like '%" & strSearch & "%' " 
			end if
			if request("chkActionSearch") = "on" then
				strTemp = strTemp &  " or Actions Like '%" & strSearch & "%' " 
			end if
			if request("chkDescriptionSearch") = "on" then
				strTemp = strTemp &  " or d.Description Like '%" & strSearch & "%' " 
			end if
			if request("chkApproverComments") = "on" then
			    strTemp = strTemp & " or Approvals Like '%" & strSearch & "%' "
			end if
			
			if strTemp = "" then 'Default to Summary
				strSQL = strSQL & " and Summary Like '%" & strSearch & "%' " 
			else
				strSQl = strSQL &	" and (" & mid(strTemp,5) & " ) "
			end if
			
			
		end if	
	
		if request("cboDaysOpenCompare") = "Range" then
			if request("txtOpenRange1") = "" then
				strRange1 = now
			else
				strRange1 =  scrubsql(request("txtOpenRange1"))
			end if
			if request("txtOpenRange2") = "" then
				strRange2 = now
			else
			strRange2 =  scrubsql(request("txtOpenRange2"))
			end if
			
			if datediff("d",strRange1,strRange2)> 0 then
				strSQL = strSQL & " and (d.Created between '" & strRange1 & "' and '" & dateadd("d",1,strRange2) & "') " 
			elseif datediff("d",strRange1,strRange2)< 0 then
				strSQL = strSQL & " and (d.Created between '" & strRange2 & "' and '" & dateadd("d",1,strRange1) & "') " 
			else
				strSQL = strSQL & " and Month(d.Created) = " & month(strRange2) & " and  Day(d.Created) = "   & day(strRange2) & " and  Year(d.Created) = " & year(strRange2) & " "  
			end if
		else
			if request("txtDaysOpen") <> "0" and request("txtDaysOpen") <> "" then
				strRange1 = scrubsql(DateAdd("d",clng("-" & request("txtDaysOpen")),now))
				if request("cboDaysOpenCompare") = "=" then
					strSQL = strSQL & " and Month(d.Created) = " & month(strRange1) & " and  Day(d.Created) = "   & day(strRange1) & " and  Year(d.Created) = " & year(strRange1) & " "  
				else
					strsql = strsql & " and '" &  strRange1 & "' " & scrubsql(request("cboDaysOpenCompare")) & " d.Created " 	
				end if
			end if
		end if
	
		if request("cboDaysClosedCompare") = "Range" then
			if request("txtClosedRange1") = "" then
				strRange1 = now
			else
				strRange1 =  scrubsql(request("txtClosedRange1"))
			end if
			if request("txtClosedRange2") = "" then
				strRange2 = now
			else
				strRange2 =  scrubsql(request("txtClosedRange2"))
			end if
			
			if datediff("d",strRange1,strRange2)> 0 then
				strSQL = strSQL & " and (d.ActualDate between '" & strRange1 & "' and '" & dateadd("d",1,strRange2) & "') " 
			elseif datediff("d",strRange1,strRange2)< 0 then
				strSQL = strSQL & " and (d.ActualDate between '" & strRange2 & "' and '" & dateadd("d",1,strRange1) & "') " 
			else
				strSQL = strSQL & " and Month(d.ActualDate) = " & month(strRange2) & " and  Day(d.ActualDate) = "   & day(strRange2) & " and  Year(d.ActualDate) = " & year(strRange2) & " "  
			end if
		else
			if request("txtDaysClosed") <> "0" and request("txtDaysClosed") <> "" then
				strRange1 = scrubsql(DateAdd("d",clng("-" & request("txtDaysClosed")),now))
				if request("cboDaysClosedCompare") = "=" then
					strSQL = strSQL & " and Month(d.ActualDate) = " & month(strRange1) & " and  Day(d.ActualDate) = "   & day(strRange1) & " and  Year(d.ActualDate) = " & year(strRange1) & " "  
				else
					strsql = strsql & " and '" &  strRange1 & "' " & scrubsql(request("cboDaysClosedCompare")) & " d.ActualDate " 	
				end if
			end if
		end if
	

		if request("cboDaysTargetCompare") = "Range" then
			if request("txtTargetRange1") = "" then
				strRange1 = now
			else
				strRange1 =  scrubsql(request("txtTargetRange1"))
			end if
			if request("txtTargetRange2") = "" then
				strRange2 = now
			else
				strRange2 =  scrubsql(request("txtTargetRange2"))
			end if
			
			if datediff("d",strRange1,strRange2)> 0 then
				strSQL = strSQL & " and (d.TargetDate between '" & strRange1 & "' and '" & strRange2 & "') " 
			elseif datediff("d",strRange1,strRange2)< 0 then
				strSQL = strSQL & " and (d.TargetDate between '" & strRange2 & "' and '" & strRange1 & "') " 
			else
				strSQL = strSQL & " and Month(d.TargetDate) = " & month(strRange2) & " and  Day(d.TargetDate) = "   & day(strRange2) & " and  Year(d.TargetDate) = " & year(strRange2) & " "  
			end if
		else
			if request("txtDaysTarget") <> "0" and request("txtDaysTarget") <> "" then
				strRange1 = scrubsql(DateAdd("d",clng(request("txtDaysTarget")),now))
				if request("cboDaysTargetCompare") = "=" then
					strSQL = strSQL & " and Month(d.TargetDate) = " & month(strRange1) & " and  Day(d.TargetDate) = "   & day(strRange1) & " and  Year(d.TargetDate) = " & year(strRange1) & " "  
				elseif request("cboDaysTargetCompare") = "<=" then
					strsql = strsql & " and d.TargetDate between '" & now & "' and '" &  strRange1 & "' "
				else
					strsql = strsql & " and d.TargetDate " & scrubsql(request("cboDaysTargetCompare")) & "'" &  strRange1 & "' "
				end if
			end if
		end if
	
		if request("cboECN") = "1" then
			strSQl = strSQL & " and d.ECNDate < '" & now & "' "
		elseif request("cboECN") = "2" then
			strSQl = strSQL & " and (d.ECNDate > '" & now & "' or d.ECNDate is null) "
		end if
	end if
	
    SELECT CASE Request("rdoDcrCategory")
        CASE "2"
            strSql = strSql & " AND BiosChange = 0 AND SwChange = 1"
        CASE "1"
            strSQL = strSQL & " AND BiosChange = 1 AND SwChange = 0"
        CASE "0"
            strSQL = strSQL & " AND BiosChange = 0 AND SwChange = 0"
    END SELECT

	if strBaseSQL = strSQL then
		Response.Write "<font size=2 face=verdana>No filters selected. Please select at least one filter on the previous screen to continue.</font>"
		blnFiltersFound = false
	else
		blnFiltersFound = true
	end if		
	
	'if request("txtDivision") = "1" then
	'	strSQl = strSQL & " and v.Division = 1  "
	'elseif request("txtDivision") = "2" then
	'	strSQl = strSQL & " and v.Division = 2  "
	'end if
	

	if request("Sort1Column") <> "" or request("Sort2Column") <> ""  or request("Sort3Column") <> "" then
		strSQl = strSQl & " Order By "
		if request("Sort1Column") <> "" then
			strSQl = strSQL &  ScrubSQL(request("Sort1Column"))
			if request("Sort1Direction") <> "" then
				strSQl = strSQL & " " &  ScrubSQL(request("Sort1Direction"))
			end if
		end if
			
		if request("Sort2Column") <> "" then
			if request("Sort1Column") <> "" then
				strSQl = strSQL & "," &  ScrubSQL(request("Sort2Column"))
				if request("Sort2Direction") <> "" then
					strSQl = strSQL & " " &  ScrubSQL(request("Sort2Direction"))
				end if
			else
				strSQl = strSQL &  ScrubSQL(request("Sort2Column"))
				if request("Sort2Direction") <> "" then
					strSQl = strSQL & " " &  ScrubSQL(request("Sort2Direction"))
				end if
			end if
		end if	

		if request("Sort3Column") <> "" then
			if request("Sort1Column") <> "" or request("Sort2Column") <> "" then
				strSQl = strSQL & "," &  ScrubSQL(request("Sort3Column"))
				if request("Sort3Direction") <> "" then
					strSQl = strSQL & " " &  ScrubSQL(request("Sort3Direction"))
				end if
			else
				strSQl = strSQL &  ScrubSQL(request("Sort3Column"))
				if request("Sort3Direction") <> "" then
					strSQl = strSQL & " " &  ScrubSQL(request("Sort3Direction"))
				end if
			end if
		end if	
	else
		strSQL = strSQL & " ORDER BY d.ID;"
	end if		
    
  '  if lcase(Session("LoggedInUser")) = "auth\dwhorton" then
  '      Response.Write strSQL
   '     Response.Flush
  '  end if

if blnFiltersFound then	
	dim strActions
	dim strResolution
	dim strApprovals
	dim strActual
	dim strTarget

	'Response.Write strSQL
	'Response.flush

	rs.Open strSQl , cn,adOpenForwardOnly
	if not (rs.EOF and rs.BOF) then
		if request("txtFunction") = "1" then
			Response.Write "<TABLE width=150% ID=tblResults cellspacing=0 cellpadding=2 bordercolor=tan bgcolor=ivory border=1>"
				if request("cboFormat")= 1 or request("cboFormat")= 2 then
				Response.Write "<THEAD bgcolor=cornsilk><TH width=60 nowrap><font size=2 face=verdana>Number</font></TH><TH nowrap width=100><font size=2 face=verdana>Product Family</font></TH><TH nowrap width=80><font size=2 face=verdana>Product</font><TH nowrap width=80><font size=2 face=verdana>Release</font></TH><TH nowrap width=80><font size=2 face=verdana>Type</font></TH><TH nowrap width=80><font size=2 face=verdana>Submitter</font></TH><TH nowrap width=80><font size=2 face=verdana>ZSRP Ready</font></TH><TH nowrap width=80><font size=2 face=verdana>ZSRP Target Date</font></TH><TH nowrap width=80><font size=2 face=verdana>ZSRP Actual Date</font></TH><TH nowrap width=80><font size=2 face=verdana>AV Required</font></TH><TH nowrap width=80><font size=2 face=verdana>Qualification Required</font></TH><TH nowrap width=80><font size=2 face=verdana>Owner</font></TH><TH nowrap width=80><font size=2 face=verdana>Status</font></TH><TH nowrap width=80><font size=2 face=verdana>Age</font></TH><TH nowrap width=80><font size=2 face=verdana>Submitted</font></TH><TH nowrap width=80><font size=2 face=verdana>TargetDate</font></TH><TH nowrap width=80><font size=2 face=verdana>ActualDate</font></TH><TH nowrap width=80><font size=2 face=verdana>Business</font></TH><TH nowrap width=800><font size=2 face=verdana><b>Summary</b></font></TH>"
			else
				Response.Write "<THEAD bgcolor=cornsilk><TD width=60><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 0,1,1);"">Number</a></font></TD><TD width=100><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 1,0,1);"">Product Family</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 2,0,1);"">Product</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 3,0,1);"">Release</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 4,0,1);"">Type</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 5,0,1);"">Submitter</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 6,0,1);"">ZSRP&nbsp;Ready</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 7,0,1);"">ZSRP&nbsp;Target&nbsp;Date</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 8,0,1);"">ZSRP&nbsp;Actual&nbsp;Date</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 9,0,1);"">AV&nbsp;Required</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 10,0,1);"">Qualification&nbsp;Required</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 11,0,1);"">Owner</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults',12,0,1);"">Status</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 13,1,1);"">Age</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 14,2,1);"">Submitted</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 15,2,1);"">TargetDate</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 16,2,1);"">ActualDate</a></font></TD><TD width=80><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 17,2,1);"">Business</a></font></TD><TD width=800><font size=1 face=verdana><a href=""javascript: SortTable( 'tblResults', 18,0,1);"">Summary</a></font></TD>"
			end if
			if request("chkIncludeActions") = "on" then
				Response.Write "<TH Width=200><font size=1 face=verdana><b>Actions</b></font></TH>"
			end if			
			if request("chkIncludeApprovers") = "on" then
				Response.Write "<TH Width=200 align=left><font size=1 face=verdana><b>Approvers</b></font></TH>"
			end if			
			if request("chkIncludeDescription") = "on" then
				Response.Write "<TH Width=200><font size=1 face=verdana><b>Description</b></font></TH>"
			end if			
			if request("chkIncludeJustification") = "on" then
				Response.Write "<TH Width=200><font size=1 face=verdana><b>Justification</b></font></TH>"
			end if			
			if request("chkIncludeResolution") = "on" then
				Response.Write "<TH Width=200><font size=1 face=verdana><b>Resolution</b></font></TH>"
			end if			
			Response.Write "</THEAD>"
		end if
	else
		Response.Write "<br><font size=2 face=verdana>No items match your query criteria</font>"
	end if
	Response.Write "<TBODY>"
	do while not rs.EOF
		if linecount >=12000 and  request("txtFunction") = "1" then
			Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Summary Reports are limited to 12000 records.</b><BR><BR></font>"
			exit do
		elseif linecount >=5000 and  request("txtFunction") <> "1" then
			Response.Write "<font size=2 face=verdana color=red><b>Some items were not displayed because Detailed Reports are limited to 5000 records.</b><BR><BR></font>"
			exit do
		end if
	if Not IsNull(rs("ZsrpRequired")) Then
        If rs("ZsrpRequired") Then
            If IsNull(rs("ZsrpReadyActualDt")) Then
                strZsrpReady = rs("ZsrpReadyTargetDt") & ""
                If DateDiff("d", NOW(), strZsrpReady) < 0 Then
                    strZsrpReady = "<span style=""color:red;"">" & strZsrpReady & "</span>"
                End If
            Else
                strZsrpReady = "Ready"
            End If
        Else
            strZsrpReady = "N/A"
        End If
    Else
        strZsrpReady = "N/A"
    End If

    'AV Required and Qualification Required
      strAVRequired = ""
      strQualificationRequired = ""

      If Not IsNull(rs("AVRequired")) Then
         If rs("AVRequired") Then
            strAVRequired = "<span class=""text"">Yes</span>"
         Else
            strAVRequired = "<span class=""text"">No</span>"
         End If
      Else
          strAVRequired = "<span class=""text"">No</span>"
      End If

      If Not IsNull(rs("QualificationRequired")) Then
         If rs("QualificationRequired") Then
            strQualificationRequired = "<span class=""text"">Yes</span>"
         Else
           strQualificationRequired = "<span class=""text"">No</span>"
         End If
      Else
          strQualificationRequired = "<span class=""text"">No</span>"
      End If
     'END AV Required and Qualification Required


		strResolution = rs("Resolution") & ""
		strActions = rs("actions") & ""
		strApprovals = rs("Approvals") & ""
		strProductType = rs("TypeID")
	
		if strApprovals = "" then
			strApprovals = "No Approvers Assigned"
		else
			strApprovals = replace(strApprovals,vbcrlf,"<BR>")
		end if


		strBusiness = ""
		if (not isnull(rs("Consumer"))) and (not isnull(rs("Consumer"))) and (not isnull(rs("Consumer"))) then
			if rs("Consumer") then
				strBusiness = strBusiness & ", Consumer"
			end if
			if rs("Commercial") then
				strBusiness = strBusiness & ", Commercial"
			end if
			if rs("SMB") then
				strBusiness = strBusiness & ", SMB"
			end if
			if strBusiness <> "" then
				strBusiness = mid(strBusiness,3)
			end if
		else
			strBusiness = "&nbsp;"
		end if


		
		select case rs("Type")
		case 1
			strType = "Issue"
		case 2
			if rs("ReleaseNotification") = "1" then
				strType = "Deliverable Release"
			else
				strType = "Action Item"
			end if
		case 3
			strType = "Change Request"
		case 4
			strType = "Status Note"
		case 5
			strType = "Improvement Opportunity"
		case 6
			strType = "Test Request"
		case 7
			strType = "Service ECR"
		case else
				strType = ""
		end select	
	
	
		if rs("ActualDate") & "" = "" then
			strActual = "&nbsp;"
		elseif rs("Status") = 1 or rs("Status") = 3 or rs("Status") = 6 then
			strActual = "&nbsp;"
		else
			strActual = formatdatetime(rs("ActualDate"),vbshortdate)
		end if

		if rs("Created") & "" = "" then
			strSubmitted = "&nbsp;"
		else
			strSubmitted = formatdatetime(rs("Created"),vbshortdate)
		end if

		if rs("TargetDate") & "" = "" then
			strTarget = "&nbsp;"
		else
			strTarget = formatdatetime(rs("targetDate"),vbshortdate)
		end if
		
		select case rs("Status")
		case 1
			strStatus = "Open"
		case 2
			strStatus = "Closed"
		case 3
			strStatus = "Need More info"
		case 4
			if not isnull(rs("ECNDate"))then
				strStatus = "ECN&nbsp;Complete"
			else
				strStatus = "Approved"
			end if
		case 5
			strStatus = "Disapproved"
		case 6
			strStatus = "Investigating"
		
		end select

		if isnull(rs("ActualDate")) then
			strAge = datediff("d",rs("Created"),now)
		else
			strAge = datediff("d",rs("Created"),rs("ActualDate"))		
		end if
	
		if request("txtFunction") = "1" then
				Response.write   "<TR class=""ID=" & rs("ID")& "&Type=" & rs("Type") & """>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()"" valign=top>" & rs("ID") & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()"" valign=top nowrap>" & rs("ProductFamily") & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()"" valign=top nowrap>" & rs("Product") & "</font></TD>"
                Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()"" valign=top nowrap>" & rs("ProductVersionRelease") & "&nbsp</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & strType & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & rs("Submitter") & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & strZsrpReady & "</font></TD>"
				Response.Write   "<TD valign=top nowrap>" & rs("ZsrpReadyTargetDt") & "&nbsp;</font></TD>"
				Response.Write   "<TD valign=top nowrap>" & rs("ZsrpReadyActualDt") & "&nbsp;</font></TD>"
                Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & strAVRequired & "</font></TD>"
                Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & strQualificationRequired & "</font></TD>"               
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & rs("Owner") & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()"" valign=top nowrap>" & strStatus & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & strAGE & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & strSubmitted & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & strtarget  & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & strActual  & "</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & strBusiness  & "&nbsp;</font></TD>"
				Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & rs("Summary") & "</font></TD>"
				if request("chkIncludeActions") = "on" then
					Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & rs("Actions") & "&nbsp;</font></TD>"
				end if
				if request("chkIncludeApprovers") = "on" then
					Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top nowrap>" & replace(rs("Approvals") & "",vbcrlf,"<BR>") & "&nbsp;</font></TD>"
				end if
				if request("chkIncludeDescription") = "on" then
					Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & rs("Description") & "&nbsp;</font></TD>"
				end if
				if request("chkIncludeJustification") = "on" then
					Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & rs("Justification") & "&nbsp;</font></TD>"
				end if
				if request("chkIncludeResolution") = "on" then
					Response.Write   "<TD LANGUAGE=javascript onmouseover=""return row_onmouseover()"" onmouseout=""return row_onmouseout()""  onclick=""return row_onclick()""  valign=top>" & rs("Resolution") & "&nbsp;</font></TD>"
				end if
				Response.Write   "</TR>"
				linecount=linecount +1
		else
			'if trim(strProductType) = "2" then
			'	Response.Write "<FONT face=verdana size=2><b>ID#: " & rs("ID") & "</a></b></font><BR>"
			'else
				Response.Write "<FONT face=verdana size=2><b>ID#: <a href=""javascript:DisplayAction(" & rs("ID") & "," & rs("Type") & ");"">" & rs("ID") & "</a></b></font><BR>"
			'end if
			if trim(rs("Type") & "") = "5" then
				Response.Write "<FONT face=verdana size=1><b>Issue/Accomplishment: " & rs("Summary") & "</b></font><BR>"
			else	
				Response.Write "<FONT face=verdana size=1><b>Summary: " & rs("Summary") & "</b></font><BR>"
			end if
			if rs("Type") = 4 then
				Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%"">"
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td><font size=1 face=verdana>" & rs("Description") &"</font></td></tr></table></TD></TR></table><BR><BR>"
			else
				Response.Write "<TABLE bordercolor=black cellspacing=0 Border=1 width=""100%""><TR>"
				Response.Write "<TD><TABLE width=""100%""><TR><TD nowrap><font face=verdana size=1>Product:</font></TD><TD><font face=verdana size=1>" & rs("Product") & "</font></TD></TR><TR><TD nowrap><font face=verdana size=1>Release:</font></TD><TD><font face=verdana size=1>" & rs("productversionRelease") & "</font></TD></TR><TR><TD><font face=verdana size=1>Type:</font></TD><TD><font face=verdana size=1>" & strType & "</font></TD></TR><TR><TD><font face=verdana size=1>Status:</font><TD><font face=verdana size=1>" & strStatus & "</font></TD></TR></table></TD>"
				Response.Write "<TD valign=top><TABLE width=""100%""><TR><TD nowrap><font face=verdana size=1>Date Created:</font></TD><TD><font face=verdana size=1>" & rs("Created") & "</font></TD></TR><TR><TD><font face=verdana size=1>Days Open:</font></TD><TD><font face=verdana size=1>" & strAge & "</font></TD></TR>"
				if trim(strProductType) = "2" then
					Response.Write "<TR><TD><font face=verdana size=1>Priority:</font><TD><font face=verdana size=1>" & rs("Priority") & "</font></TD></TR>"
				else
					'Response.Write "<TR><TD><font face=verdana size=1>Target Date:</font><TD><font face=verdana size=1>" & rs("TargetDate") & "</font></TD></TR>"
				end if				
				Response.Write "</table></TD>"
				
				if  rs("PendingImplementation") and strStatus <> "Closed" then			
					strWorkingList = "Yes"
				else				
					strWorkingList = "No"
				end if
				Response.Write "<TD valign=top><TABLE width=""100%"" ><TR><TD nowrap><font face=verdana size=1>Submitter:</font></TD><TD><font face=verdana size=1>" & rs("Submitter") & "</font></TD></TR><TR><TD><font face=verdana size=1>Owner:</font></TD><TD><font face=verdana size=1>" & rs("Owner") & "</font></TD></TR>"
				if trim(strProductType) = "2" then
					Response.Write "<TR><TD><font face=verdana size=1>On&nbsp;Working&nbsp;List:</font><TD><font face=verdana size=1>" & strWorkingList & "</font></TD></TR>"
				else
					Response.Write "<TR><TD><font face=verdana size=1>Core Team Rep:</font><TD><font face=verdana size=1>" & rs("CoreTeamRep") & "</font></TD></TR>"
				end if
				Response.Write "</table></TD></TR>"
				
				if trim(rs("Type") & "") = "5" then
					if rs("AffectsCustomers")=1  then
						strCustomers = "Positive"
					elseif rs("AffectsCustomers")=0 then
						strCustomers = "&nbsp;"
					else
						strCustomers = "Negative"
					end if
				
					select case trim(rs("Priority") & "")
					case "1"
						strPriority="High"
					case "2"
						strPriority="Medium"
					case "3"
						strPriority="Low"
					case else
						strPriority=""
					end select			
				
				
				
					Response.Write "<TR><TD><table width=""100%"" ><tr><td width=70 nowrap valign=top><font size=1 face=verdana>Impact: </font></td><td><font size=1 face=verdana>" & strPriority &"</font></td></tr></table></TD><TD><table width=""100%"" ><tr><td width=110 nowrap valign=top><font size=1 face=verdana>Net Affect: </font></td><td><font size=1 face=verdana>" & strCustomers &"</font></td></tr></table></TD><TD><table width=""100%"" ><tr><td width=130 nowrap valign=top><font size=1 face=verdana>Metric Impacted: </font></td><td><font size=1 face=verdana>" & rs("AvailableNotes") &"</font></td></tr></table></TD></TR>"
				end if
				strDescription = trim(replace(replace(rs("Description"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR>"))
				strDetails = trim(replace(replace(rs("Details")&"", vbcrlf & vbcrlf,"<BR><BR>"), vbcrlf, "<BR>"))
				
				if trim(rs("Type") & "") = "5" then
					StringArray = split(strDescription,chr(1))
					if ubound(StringArray) > -1 then
						if trim(StringArray(0)) <> "" then
							strDescription = "<b>POSITIVE IMPACT:</b><br>" & StringArray(0)
						else
							strDescription = ""				
						end if
					end if
					if ubound(StringArray) > 0 then
						if trim(StringArray(0)) <> "" and trim(StringArray(1)) <> ""  then
							strDescription = strDescription & "<BR><BR>"
						end if
						if trim(StringArray(1)) <> "" then
							strDescription = strDescription & "<b>NEGATIVE IMPACT:</b><br>" & StringArray(1)
						end if
					end if
				end if
				
				if trim(strProductType) = "2" then
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=80 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & strDescription &"</font></td></tr></table></TD></TR>"
				else
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Description: </font></td><td><font size=1 face=verdana>" & strDescription &"</font></td></tr></table></TD></TR>"
				end if
				Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Details: </font></td><td><font size=1 face=verdana>" & strDetails &"</font></td></tr></table></TD></TR>"
				if strBusiness <> "&nbsp;"  and trim(strBusiness) <> ""  then
					if trim(rs("Type") & "") = "5" then
						set rs2 = server.CreateObject("ADODB.recordset")
						rs2.Open "spListGroups4Action " & rs("ID"),cn,adOpenForwardOnly
						strFunctionalGroup=""
						do while not rs2.EOF
							if trim(rs2("ID") & "") <> "" then
								strFunctionalGroup = strFunctionalGroup & "," & rs2("GroupName") 
							end if
							rs2.MoveNext
						loop
						rs2.Close
						set rs2 = nothing
						if strFunctionalGroup <> "" then
							Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Sub-Group Owner: </font></td><td width=""100%""><font size=1 face=verdana>" & mid(strFunctionalGroup,2) & "</font></td></tr></table></TD></TR>"
						end if
					else
						Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Business: </font></td><td width=""100%""><font size=1 face=verdana>" & strBusiness & "</font></td></tr></table></TD></TR>"
					end if
				end if

				if trim(rs("Justification") & "") <> "" then
					if trim(rs("Type") & "") = "5" then
						Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Root Cause: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR>") & "</font></td></tr></table></TD></TR>"
					else
						Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Justification: </font></td><td width=""100%""><font size=1 face=verdana>" & replace(replace(rs("Justification"),vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR>") & "</font></td></tr></table></TD></TR>"
					end if
				end if
				
				If cbool(rs("ZsrpRequired") & "") Then
				    Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>ZSRP Ready: </font></td><td width=""100%""><font size=1 face=verdana>" & strZsrpReady & "<br>ZSRP Target: " & rs("ZsrpReadyTargetDt") & "<br>ZSRP Actual: " & rs("ZsrpReadyActualDt") & "</font></td></tr></table></TD></TR>"
				End If
				
				strDescription = trim(replace(replace(strActions,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR>"))
				
				if trim(rs("Type") & "") = "5" then
					StringArray = split(strDescription,chr(1))
					if ubound(StringArray) > -1 then
						if trim(StringArray(0)) <> "" then
							strDescription = "<b>Corrective Actions:</b><br>" & StringArray(0)
						else
							strDescription = ""				
						end if
					end if
					if ubound(StringArray) > 0 then
						if trim(StringArray(0)) <> "" and trim(StringArray(1)) <> ""  then
							strDescription = strDescription & "<BR><BR>"
						end if
						if trim(StringArray(1)) <> "" then
							strDescription = strDescription & "<b>Preventive Actions:</b><br>" & StringArray(1)
						end if
					end if
				end if
				if trim(strProductType) <> "2" then
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Actions Required: </font></td><td width=""100%""><font size=1 face=verdana>" & strDescription & "</font></td></tr></table></TD></TR>"
				end if
				if trim(strResolution) <> "" then
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Resolution: </font></td><td><font size=1 face=verdana>" & replace(replace(strResolution,vbcrlf & vbcrlf,"<BR><BR>"),vbcrlf,"<BR>") &"</font></td></tr></table></TD></TR>"
				end if
								if trim(strProductType) = "2" then
					set rs2 = server.CreateObject("ADODB.recordset")
					strRoadmap = ""
					rs2.Open "spGetActionRoadmapItem4Task " & rs("ID"),cn,adOpenForwardOnly
					if rs2.eof and rs2.bof then					
						strRoadmap = "TBD"
					else
						if trim(rs2("Summary") & "") = "" then
							strRoadmap = "TBD"
						else
							strRoadmap = "<Table>"
							strRoadmap = strRoadmap & "<TR><TD valign=top><font size=1 face=verdana><b>Summary:&nbsp;&nbsp;</b></font></TD><TD>" & rs2("Summary") & "</TD></TR>"
							if trim(rs2("Notes") & "") <> "" then
								strRoadmap = strRoadmap & "<TR><TD valign=top><font size=1 face=verdana><b>Notes:</b></font></TD><TD>" & rs2("Notes") & "</TD></TR>"
							end if
							if trim(rs2("Details") & "") <> "" then
								strRoadmap = strRoadmap & "<TR><TD valign=top><font size=1 face=verdana><b>Details:</b></font></TD><TD>" & rs2("Details") & "</TD></TR>"
							end if
							strRoadmap = strRoadmap & "</Table>"
						end if
					end if						
					rs2.close
					set rs2 = nothing
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=80 nowrap valign=top><font size=1 face=verdana>Roadmap: </font></td><td><font size=1 face=verdana>" & strRoadmap &"</font></td></tr></table></TD></TR>"
				else
					Response.Write "<TR><TD colspan=3><table width=""100%"" ><tr><td width=150 nowrap valign=top><font size=1 face=verdana>Approvals: </font></td><td><font size=1 face=verdana>" & strApprovals &"</font></td></tr></table></TD></TR>"
				end if
				Response.Write "</TABLE><BR><BR>"
				linecount=linecount +1
			end if		
		end if
	
		Response.flush
	
		rs.MoveNext
	loop

	if not (rs.EOF and rs.BOF) then
		Response.Write "</TBODY>"
		Response.Write "</table>"
	end if

	rs.Close

	Response.Write "<BR><font size=1 face=verdana>Items Displayed: " & LineCount & "</font>"
'	Response.Write "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>" & strSQL

end if
	cn.Close
	set rs=nothing
	set cn=nothing

%>

</BODY>
</HTML>