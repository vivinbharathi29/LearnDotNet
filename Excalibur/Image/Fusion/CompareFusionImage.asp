<%@ Language=VBScript %>
<%
		Response.Buffer = True
		Response.ExpiresAbsolute = Now() - 1
		Response.Expires = 0
		Response.CacheControl = "no-cache"	
%>

<HTML>
<STYLE>
TD
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: ivory;
}
TH
{
    FONT-SIZE: x-small;
    FONT-FAMILY: Verdana;
	BACKGROUND-COLOR: cornsilk;
}

</STYLE>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        lblProcessing.style.display = "none";
    }

    function DisplayTargetIssues() {
        TargetIssuesRow.style.display = "";
        ImageIssuesRow.style.display = "none";
    }

    function DisplayImageIssues() {
        TargetIssuesRow.style.display = "none";
        ImageIssuesRow.style.display = "";
    }

    function CompareLines(strTable) {
        var i;
        document.all("frmCompare" + strTable).submit();
    }

    function SwitchType(ID, ProductID, PINTest, IncludeIRSOnly) {
        var strIncludeIRSOnly = "";
        if (IncludeIRSOnly == "1")
            strIncludeIRSOnly = "&IncludeIRSOnly=1";

        if (ID == 1)
            window.location.href = "CompareFusionImage.asp?ProductID=" + ProductID + "&PINTest=" + PINTest + "&CompareType=1" + strIncludeIRSOnly;
        else
            window.location.href = "CompareFusionImage.asp?ProductID=" + ProductID + "&PINTest=" + PINTest + strIncludeIRSOnly;
    }

    function SwitchIRSOnly(ID, ProductID, PINTest, CompareType) {
        var strCompareType = "";
        if (CompareType == "1")
            strCompareType = "&CompareType=1";

        if (ID == 1)
            window.location.href = "CompareFusionImage.asp?ProductID=" + ProductID + "&PINTest=" + PINTest + strCompareType + "&IncludeIRSOnly=1";
        else
            window.location.href = "CompareFusionImage.asp?ProductID=" + ProductID + "&PINTest=" + PINTest + strCompareType;

    }

    //-->

</SCRIPT>
</HEAD>
<BODY  LANGUAGE=javascript onload="return window_onload()">
<b><font ID=lblProcessing face=verdana size=2>Processing. This may take several minutes.  Please wait...</font></b>
<font size=3 face=verdana><b>Compare Excalibur Images to IRS</b></font><br /><br />
<font size=2 face=verdana>
    <b>Compare Type:</b>
    <%if request("CompareType") = "" then %>
        In Image | <a href="javascript:SwitchType(1,<%=request("ProductID")%>,<%=request("PINTest")%>,'<%=request("IncludeIRSOnly")%>')">Targeted</a>
    <%else%>
        <a href="javascript:SwitchType(0,<%=request("ProductID")%>,<%=request("PINTest")%>,'<%=request("IncludeIRSOnly")%>')">In Image</a> | Targeted
    <%end if%>
    <br />
    <b>IRS-only Components:</b> 
    <%if trim(request("IncludeIRSOnly")) = "" then %>
        Exclude |  <a href="javascript:SwitchIRSOnly(1,<%=request("ProductID")%>,<%=request("PINTest")%>,'<%=request("CompareType")%>')">Include</a>
    <%else%>
        <a href="javascript:SwitchIRSOnly(0,<%=request("ProductID")%>,<%=request("PINTest")%>,'<%=request("CompareType")%>')">Exclude</a> | Include
    <%end if%>
    <br /><br />
</font>
<%
    
	Server.ScriptTimeout = 1200

	dim StartDate
	StartDate = now()

	dim cn
	
	dim cm
	dim p
	dim rs
	dim rs2
	dim strDash
	dim strOutBuffer
	dim strSQL
	dim errorcount
	dim totalerrorcount
	dim TotalCompared
	dim TableCount
	dim CurrentUserPartner
	dim TableHeaderDisplayed
	dim currentuserid
    dim strProductDrops
    dim ProductDropArray
	dim tmp
	
	TableCount =0
    strProductDrops = ""


   	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")

	'Get User
	dim CurrentDomain

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
	if (rs.EOF and rs.BOF) then
		set rs = nothing
		set cn=nothing
		Response.Redirect "../../NoAccess.asp?Level=0"
	else
		CurrentUserPartner = rs("PartnerID")
		currentuserid = rs("ID")
	end if 
	rs.Close		



    if trim(request("ProductDrop")) = "" and trim(request("ProductID")) = "" and trim(request("ImageDefinitionID")) = "" then
        Response.Write "<BR><font size=2 face=verdana>Unable to run report because all of the required information was not supplied.</font><BR><BR>"
    else

        if trim(request("ImageDefinitionID")) <> "" then 'Image Definition Supplied
            rs.open "spListProductDrops " & clng(request("ImageDefinitionID")),cn
            do while not rs.eof
                strProductDrops = strProductDrops & "," & rs("ProductDrop")
                rs.movenext 
            loop
            rs.close
        elseif trim(request("ProductDrop")) <> "" then 'Product Drop Supplied
            strProductDrops = strProductDrops & "," & request("ProductDrop")
        else 'productID supplied
            if trim(request("IncludeIRSOnly")) = "1" then
                rs.open "spListProductDrops4product " & clng(request("ProductID")) & ",1",cn
            else
                rs.open "spListProductDrops4product " & clng(request("ProductID")) & ",0",cn
                do while not rs.eof
                strProductDrops = strProductDrops & "," & rs("ProductDrop")
                rs.movenext 
                loop 
            rs.close
            end if
        end if
    end if

    if strProductDrops <> "" then
        strProductDrops = mid(strProductDrops,2)
    end if
	
    ProductDropArray = split(strProductDrops,",")
    for i = 0 to ubound(productdroparray)
        if trim(Productdroparray(i)) = "" then
            response.write "SKIPPED A PRODUCT DROP WITH NO NUMBER ASSIGNED<br><br>"
        else
          '  response.write "+++" & Productdroparray(i)  & "<BR>"
            rs2.Open "spListImagesForProductDrop '" & Productdroparray(i) & "'",cn,adOpenForwardOnly
            strImageList = ""
            do while not rs2.eof
                strImageList = strImageList & "," & rs2("ImageID")
                rs2.movenext
            loop
            rs2.close
            if strImageArray <> "" then
                strImageArray = mid(strImageArray,2)
            end if
			
	
            ImageArray = split(trim(strImageList),",")
            'response.write strImageList & "<br>"

            strDelIRS = ""
            if trim(request("IncludeIRSOnly")) = "1" then
                strSQl = "spListComponentsInIRSImage '" & FixProductDrop(productdroparray(i)) & "'"
            else
                strSQl = "spListComponentsInIRSImage '" & FixProductDrop(productdroparray(i)) & "'"
            end if
            rs2.Open strSQl,cn,adOpenForwardOnly
		    do while not rs2.EOF
			    strDelIRS = strDelIRS & "2" & trim(rs2("CompName")) & " " &  trim(rs2("Version"))
			    if  rs2("Revision") & "" <> "" then
				    strDelIRS = strDelIRS & " " & trim(rs2("Revision"))
			    end if
			    if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
				    strDelIRS = strDelIRS & asc(lcase(trim(rs2("Pass")))) -87
			    else
				    strDelIRS = strDelIRS & trim(rs2("Pass"))
			    end if
			    if  rs2("PartNo") & "" <> "" then
				    strDelIRS = strDelIRS & " - " & trim(rs2("PartNo"))
			    end if
			    strDelIRS = strDelIRS & vbcrlf
			    rs2.MoveNext
		    loop
		    rs2.Close
            'response.write replace(strDelIRS,vbcrlf,"<br>") & "<br><br>"   
            if request("CompareType") = "" then
                strSQl = "spListDeliverablesInProductDrop '" & productdroparray(i) & "'," & clng(request("ProductID"))
            else
                strSQl = "spListDeliverablesInProductDrop '" & productdroparray(i) & "'," &  clng(request("ProductID")) & ",1" 
            end if
            rs2.Open strSQl,cn,adOpenForwardOnly
	        strOutbuffer = ""	
            do while not rs2.EOF
			    if (rs2("Preinstall") or rs2("Preload") or rs2("ARCD") or rs2("SelectiveRestore")) then 
					DeliverableImageArray=""
                    blnFound = false
					if len(rs2("Images")) > 0 then
						DeliverableImageArray = "," & Replace(rs2("Images")," ","") & ","
					end if
                    
                    for j = 1 to ubound(ImageArray)				
						 If InStr(DeliverableImageArray, "," & imagearray(j) & ",") > 0 then
						   blnfound = true
						   exit for
						end if
                    next
                    if trim(rs2("Images") & "") = "" or blnFound then
				        strOutbuffer = strOutbuffer & "1" & trim(rs2("Name")) & " " &  rs2("Version")
				        if  rs2("Revision") & "" <> "" then
					        strOutbuffer = strOutbuffer & " " & rs2("Revision")
				        end if
				        if lcase(rs2("Pass")) >="a" and lcase(rs2("Pass")) <= "z" then
					        strOutbuffer = strOutbuffer & asc(lcase(rs2("Pass"))) -87
				        else
					        strOutbuffer = strOutbuffer & rs2("Pass")
				        end if
				        if  rs2("IRSPartNumber") & "" <> "" then
					        strOutbuffer = strOutbuffer & " - " & trim(rs2("IRSPartNumber"))
				        end if
				        strOutbuffer = strOutbuffer & vbcrlf
                    end if
			    end if
			    rs2.MoveNext
		    loop
		    rs2.close		


            'Sort IRS Deliverables         
			LineArray = Split(lcase(strDelIRS), vbcrlf)
			strDelIRS = ""
			For j = UBound(LineArray) - 1 To 0 Step -1
				If mid(LineArray(j),2) > mid(LineArray(j + 1),2) Then
					temp = LineArray(j + 1)
					LineArray(j + 1) = LineArray(j)
					LineArray(j) = temp
				End If
			Next
           
						
		    'Remove Dups and append to excalibur list
		    for j = lbound(LineArray) to ubound(LineArray)					
			    if j = ubound(LineArray) then
				    strOutbuffer = strOutbuffer & LineArray(j) & vbcrlf
			    elseif linearray(j) <> linearray(j+1) then
				    strOutbuffer = strOutbuffer & LineArray(j) & vbcrlf
			    end if
		    next
    	    'Sort
		    LineArray = Split(lcase(strOutBuffer), vbcrlf)
		    if  UBound(LineArray) > 0 then
			    TotalCompared = TotalCompared + UBound(LineArray)
		    end if
		    For j = UBound(LineArray) - 1 To 0 Step -1
			    For k = 0 To j
				    If mid(LineArray(k),2) > mid(LineArray(k + 1),2) Then
					    temp = LineArray(k + 1)
    				    LineArray(k + 1) = LineArray(k)
					    LineArray(k) = temp
				    End If
			    Next
		    Next

            j=0
            TableHeaderDisplayed=false
            ErrorCount=0
		    do while j < UBound(LineArray)

			    if mid(lcase(LineArray(j)),2) = mid(lcase(LineArray(j+1)),2) then
				    j=j+2
			    else
				    if not TableHeaderDisplayed then
					    Response.Write "<DIV ID=DIV" & TableCount & "><form action=""ShowDifference.asp""  method=post target=""_blank"" id=frmCompare" & TableCount & ">"
					    Response.Write "<font size=2 face=verdana><b>Product Drop: " & Productdroparray(i) & " - Discrepancies Found</b></font>" 
					    Response.write "<TABLE width=100% border=1 bordercolor=tan bgcolor=ivory>"
					    Response.write  "<TR><TH align=left><a href=""javascript: CompareLines(" & TableCount & ");"">Compare</a></TH><TH align=left>System</TH><TH align=left>Deliverable</TH></TR>"
					    TableHeaderDisplayed = true
				    end if
				    if left(LineArray(j),1) = "1" then
					    Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""E" & lcase(mid(LineArray(j),2))  & """></td>"
					    Response.Write "<TD>Excalibur</td>"
					    Response.Write "<TD>" & lcase(mid(LineArray(j),2))  & "</td></TR>"
				    elseif left(LineArray(j),1) = "2" then
					    Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""C" & lcase(mid(LineArray(j),2))  & """></td>"
					    Response.Write "<TD>IRS</td>"
					    Response.Write "<TD>" & lcase(mid(LineArray(j),2))  & "</td></TR>"
				    end if
                    ErrorCount = ErrorCount + 1
				    totalErrorCount = totalErrorCount + 1
				    j=j+1
			    end if
		    loop


               if ubound(linearray) <> -1 and j <=ubound(linearray) then
				    'Response.Write TableHeaderDisplayed & "<BR>"
				    if not TableHeaderDisplayed then
					    Response.Write "<DIV ID=DIV" & TableCount & "><form action=""ShowDifference.asp""  method=post target=""_blank"" id=frmCompare" & TableCount & ">"
					    Response.Write "<font size=2 face=verdana><b>Product Drop: " & Productdroparray(i) & " - Discrepancies Found</b></font>" 
					    Response.write "<TABLE width=100% border=1 bordercolor=tan bgcolor=ivory>"
					    Response.write  "<TR><TH align=left><a href=""javascript: CompareLines(" & TableCount & ");"">Compare</a></TH><TH align=left>System</TH><TH align=left>Deliverable</TH></TR>"
					    TableHeaderDisplayed = true
				    end if
				    if left(LineArray(j),1) = "1" then
					    Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""E" & lcase(mid(LineArray(j),2))  & """></td>"
					    Response.Write "<TD>Excalibur</td>"
					    Response.Write "<TD>" & lcase(mid(LineArray(j),2))  & "</td></TR>"
				    elseif left(LineArray(j),1) = "2" then
					    Response.Write "<TR><TD><INPUT type=""checkbox"" id=chkResult name=chkResult value=""C" & lcase(mid(LineArray(j),2))  & """></td>"
					    Response.Write "<TD>IRS</td>"
					    Response.Write "<TD>" & lcase(mid(LineArray(j),2))  & "</td></TR>"
				    end if
                    ErrorCount = ErrorCount + 1
				    totalErrorCount = totalErrorCount + 1
                end if
        if ErrorCount = 0 and strOutbuffer= "" then
            Response.Write "<font size=2 face=verdana><b>Product Drop: " & Productdroparray(i) & " - No Components found for this product drop in IRS or Excalbur</b></font><br><br>"
        elseif ErrorCount = 0 then
            Response.Write "<font size=2 face=verdana><b>Product Drop: " & Productdroparray(i) & " - No Discrepancies Found</b></font><br><br>"
        else
            response.Write "</table><br><br>"
        end if

        end if   
    next





    set rs=nothing
    cn.close
    set cn = nothing


function FixProductDrop(strProductDrop)
    if right(strProductDrop,2) = "##" then
        FixProductDrop = strProductDrop
    else
        FixProductDrop = left(strProductDrop,len(strProductDrop)-2) & "##"
    end if
end function


%>

</BODY>
</HTML>
