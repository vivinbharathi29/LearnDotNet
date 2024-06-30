<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<STYLE>
TD
{
	FONT-SIZE: xx-small;
	Font-Family: Verdana;	
}
BODY
{
	FONT-SIZE: x-small;
	Font-Family: Verdana;	
}
	
	
</STYLE>
<BODY>

<%
	dim cn
	dim rs
	dim strLeadName
	dim strProductName
	dim strLeadID
	dim strFollowID

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")
    
    if request("FusionRequirements") = 0 then
	    strSQL =  "Select p.dotsname as product, p.id as FollowID, p2.dotsname as Lead, p2.id as LeadID " & _
			    "From ProductVersion p with (NOLOCK), ProductVersion p2 with (NOLOCK) " & _
			    "Where p.id = " & clng(request("ProductID")) & " " & _
			    "and p.referenceid = p2.id;"
    else 
        strSQL =  "Select p.dotsname + ' (' + pvr.Name + ')' as product, pv_r.id as FollowID, p2.dotsname + ' (' + lpvr.Name + ')' as Lead, lpv_r.id as LeadID " & _
			      "From ProductVersion p with (NOLOCK) inner join " & _
                  "ProductVersion_Release pv_r with (NOLOCK) on pv_r.ProductVersionID = p.ID inner join " & _
                  "ProductVersion_Release lpv_r with (NOLOCK) on pv_r.LeadProductReleaseID = lpv_r.ID inner join " & _
                  "ProductVersion p2 with (NOLOCK) on p2.id = lpv_r.ProductVersionID inner join " & _
                  "ProductVersionRelease lpvr with (NOLOCK) on lpvr.id = lpv_r.ReleaseID inner join " & _
                  "ProductVersionRelease pvr with (NOLOCK) on pvr.id = pv_r.ReleaseID " & _
			      "Where pv_r.id = " & clng(request("ID"))
    end if

	rs.Open strSQL,cn,adOpenStatic
	if rs.EOF and rs.BOF then
		strLeadName = ""
		strProductName = ""
		strLeadID=0
		strFollowID=0
	else
		strLeadName = trim(rs("Lead") & "")
		strProductName = trim(rs("Product") & "")
		strLeadID=trim(rs("LeadID")&"")
		strFollowID=trim(rs("FollowID")&"")
	end if
	rs.Close
		
	if strLeadName = "" or strproductName = "" or request("RootID") = "" then
		Response.Write "Unable to find the requested products."
	else
	
        if request("FusionRequirements") = 0 then
		    strSQL = "Select distinct v.ID " & _
				     "FROM product_deliverable pd with (NOLOCK), deliverableversion v with (NOLOCK) " & _
				     "Where pd.targeted=1 " & _
				     "and v.id = pd.deliverableversionid " & _
				     "and v.deliverablerootid = " & clng(request("RootID")) & " " & _
				     "and productVersionID in (" & strLeadID & "," & strFollowID & ");"
        else
            strSQL = "Select distinct v.ID " & _
				     "FROM product_deliverable pd with (NOLOCK) inner join " & _
                     "Product_Deliverable_Release pdr with (NOLOCK) on pd.id = pdr.ProductDeliverableID inner join " & _
                     "ProductVersion_Release pvr with (NOLOCK) on pvr.ProductVersionID = pd.ProductVersionID and pvr.ReleaseID = pdr.ReleaseID inner join " & _
                     "deliverableversion v with (NOLOCK) on v.id = pd.deliverableversionid " & _
				     "Where pdr.targeted=1 " & _
				     "and v.deliverablerootid = " & clng(request("RootID")) & " " & _
				     "and pvr.ID in (" & strLeadID & "," & strFollowID & ");"
        end if
		rs.Open strSQL,cn,adOpenStatic
		strIDList = ""
		do while not rs.EOF
			strIDList = strIDList & "," &  rs("ID") 
			rs.MoveNext
		loop
		rs.Close
		if strIDList = "" then
			Response.Write "Unable to find the requested deliverables."
		else
			strIDList = mid(strIDList,2)
			
            if request("FusionRequirements") = 0 then
			    strSQL = "Select 1 as Lead, v.id, v.deliverablename, v.version, v.revision, v.pass, p.id as productID, p.dotsname as Product, pd.targeted, pd.imagesummary, pd.targetnotes,pd.preinstall, pd.preload, pd.web,pd.dropinbox,pd.arcd as drcd, pd.drdvd, pd.selectiverestore, pd.preinstallbrand, pd.preloadbrand, pd.patch, pd.doccd, pd.oscd, pd.RACD_EMEA, pd.RACD_APD, pd.RACD_Americas, pd.DIBHWReq " & _
					     "from product_deliverable pd with (NOLOCK) inner join " & _
                         "deliverableversion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _
                         "productversion p with (NOLOCK) on p.id = pd.productversionid " & _
					     "where pd.productversionid  = " & clng(strLeadID) & " " & _
					     "and v.id in (" & scrubsql(strIDList) & ") " & _
					     " union " & _
					     "Select 0 as Lead, v.id, v.deliverablename, v.version, v.revision, v.pass, p.id as productID, p.dotsname as Product, pd.targeted, pd.imagesummary, pd.targetnotes,pd.preinstall, pd.preload, pd.web,pd.dropinbox,pd.arcd as drcd, pd.drdvd, pd.selectiverestore, pd.preinstallbrand, pd.preloadbrand, pd.patch, pd.doccd, pd.oscd, pd.RACD_EMEA, pd.RACD_APD, pd.RACD_Americas, pd.DIBHWReq " & _
					     "from product_deliverable pd with (NOLOCK) inner join " & _
                         "deliverableversion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _ 
                         "productversion p with (NOLOCK) on p.id = pd.productversionid " & _
					     "where pd.productversionid = " & clng(strFollowID) & " " & _
					     "and v.id in (" & scrubsql(strIDList) & ") " & _
					     "order by v.id desc , p.dotsname"
            else
                strSQL = "Select 1 as Lead, v.id, v.deliverablename, v.version, v.revision, v.pass, p.id as productID, p.dotsname + ' (' + pvr.Name + ')' as Product, pdr.targeted, pdr.imagesummary, pdr.targetnotes, pdr.preinstall, pdr.preload, pdr.web, pdr.dropinbox, pdr.arcd as drcd, pdr.drdvd, pdr.selectiverestore, pdr.preinstallbrand, pdr.preloadbrand, pdr.patch, pdr.doccd, pdr.oscd, pdr.RACD_EMEA, pdr.RACD_APD, pdr.RACD_Americas, pdr.DIBHWReq, pvr.ReleaseYear, pvr.ReleaseMonth, p.dotsname " & _
					     "from product_deliverable pd with (NOLOCK) inner join " & _
                         "product_deliverable_release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.id inner join " & _
                         "ProductVersion_Release pv_r with (NOLOCK) on pv_r.ProductVersionID = pd.ProductVersionID and pv_r.ReleaseID = pdr.ReleaseID inner join " & _
                         "ProductVersionRelease pvr with (NOLOCK) on pvr.id = pv_r.ReleaseID inner join " & _
                         "deliverableversion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _
                         "productversion p with (NOLOCK) on p.id = pd.productversionid " & _
					     "where pv_r.id  = " & clng(strLeadID) & " " & _
					     "and v.id in (" & scrubsql(strIDList) & ") " & _
					     " union " & _
					     "Select 0 as Lead, v.id, v.deliverablename, v.version, v.revision, v.pass, p.id as productID, p.dotsname + ' (' + pvr.Name + ')' as Product, pdr.targeted, pdr.imagesummary, pdr.targetnotes, pdr.preinstall, pdr.preload, pdr.web, pdr.dropinbox, pdr.arcd as drcd, pdr.drdvd, pdr.selectiverestore, pdr.preinstallbrand, pdr.preloadbrand, pdr.patch, pdr.doccd, pdr.oscd, pdr.RACD_EMEA, pdr.RACD_APD, pdr.RACD_Americas, pdr.DIBHWReq, pvr.ReleaseYear, pvr.ReleaseMonth, p.dotsname " & _
					     "from product_deliverable pd with (NOLOCK) inner join " & _
                         "product_deliverable_release pdr with (NOLOCK) on pdr.ProductDeliverableID = pd.id inner join " & _
                         "ProductVersion_Release pv_r with (NOLOCK) on pv_r.ProductVersionID = pd.ProductVersionID and pv_r.ReleaseID = pdr.ReleaseID inner join " & _
                         "ProductVersionRelease pvr with (NOLOCK) on pvr.id = pv_r.ReleaseID inner join " & _
                         "deliverableversion v with (NOLOCK) on pd.deliverableversionid = v.id inner join " & _ 
                         "productversion p with (NOLOCK) on p.id = pd.productversionid " & _
					     "where pv_r.id = " & clng(strFollowID) & " " & _
					     "and v.id in (" & scrubsql(strIDList) & ") " & _
					     "order by v.id desc , p.dotsname, pvr.ReleaseYear desc, pvr.ReleaseMonth desc"
            end if

			rs.Open strsql,cn,adOpenStatic
			LastVersion = 0
			TargetedLead = ""
			DistributionLead=""
			TargetNotesLead = ""
			ImageSummaryLead = ""
			TargetedProduct = ""
			DistributionProduct=""
			TargetNotesProduct = ""
			ImageSummaryProduct = ""
			VersionBGColor = "gainsboro"
			do while not rs.EOF
				if LastVersion= 0 then
					response.write "Compare Targeted Versions of " & rs("DeliverableName") & "<BR><BR>" 
					Response.Write "<TABLE bgcolor=ivory border=1 width=""100%"" cellpadding=2 cellspacing=0>"
					Response.Write "<TR bgcolor=beige><TD><b>Version</b></td><TD><b>Product</b></td><TD><b>Targeted</b></td><TD><b>Distribution</b></td><TD><b>Target Notes</b></td><TD><b>Image Summary</b></td></TR>"
				elseif LastVersion <> rs("ID") then
					if VersionBGColor = "ivory" then
						VersionBGColor = "gainsboro"
					else					
						VersionBGColor = "ivory"
					end if
					Response.Write "<TR bgcolor=" & VersionBGColor & ">"
					Response.write "<TD valign=top rowspan=2>" & strVersion & "</TD>"
					Response.write "<TD>" & strLeadName & "</TD>"
					if trim(TargetedLead) = "" then
						response.write "<td colspan=4 bgcolor=Thistle><b>Not Supported</b></td>"
					else
						if TargetedLead <> TargetedProduct then
							Response.write "<TD><b><font color=red>" & TargetedLead & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & TargetedLead & "&nbsp;</TD>"
						end if
						if trim(DistributionLead) <> trim(DistributionProduct) then
							Response.write "<TD><b><font color=red>" & DistributionLead & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & DistributionLead & "&nbsp;</TD>"
						end if
						if trim(TargetNotesLead) <> trim(TargetNotesProduct) then
							Response.write "<TD><b><font color=red>" & TargetNotesLead & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & TargetNotesLead & "&nbsp;</TD>"
						end if
						if trim(ImageSummaryLead) <> trim(ImageSummaryProduct) then
							if trim(ImageSummaryLead)= "" then
								Response.write "<TD><b><font color=red>All&nbsp;</font></b></TD>"
							else
								Response.write "<TD><b><font color=red>" & ImageSummaryLead & "&nbsp;</font></b></TD>"
							end if
						else
							if trim(ImageSummaryLead)= "" then
								Response.write "<TD>All&nbsp;</TD>"
							else
								Response.write "<TD>" & ImageSummaryLead & "&nbsp;</TD>"
							end if
						end if
					end if
					Response.Write "</TR>"
					Response.Write "<TR bgcolor=" & VersionBGColor & ">"
					Response.write "<TD>" & strProductName & "</TD>"
					if trim(TargetedProduct) = "" then
						response.write "<td colspan=4 bgcolor=Thistle><b>Not Supported</b></td>"
					else
						if TargetedLead <> TargetedProduct then
							Response.write "<TD><b><font color=red>" & TargetedProduct & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & TargetedProduct & "&nbsp;</TD>"
						end if
						if trim(DistributionLead) <> trim(DistributionProduct) then
							Response.write "<TD><b><font color=red>" & DistributionProduct & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & DistributionProduct & "&nbsp;</TD>"
						end if
						if trim(TargetNotesLead) <> trim(TargetNotesProduct) then
							Response.write "<TD><b><font color=red>" & TargetNotesProduct & "&nbsp;</font></b></TD>"
						else
							Response.write "<TD>" & TargetNotesProduct & "&nbsp;</TD>"
						end if
						if trim(ImageSummaryLead) <> trim(ImageSummaryProduct) then
							if trim(ImageSummaryProduct)= "" then
								Response.write "<TD><b><font color=red>All&nbsp;</font></b></TD>"
							else
								Response.write "<TD><b><font color=red>" & ImageSummaryProduct & "&nbsp;</font></b></TD>"
							end if
						else
							if trim(ImageSummaryProduct)= "" then
								Response.write "<TD>All&nbsp;</TD>"
							else
								Response.write "<TD>" & ImageSummaryProduct & "&nbsp;</TD>"
							end if
						end if
					end if
					Response.Write "</TR>"
					TargetedLead = ""
					DistributionLead=""
					TargetNotesLead = ""
					ImageSummaryLead = ""
					TargetedProduct = ""
					DistributionProduct=""
					TargetNotesProduct = ""
					ImageSummaryProduct = ""

				end if
				LastVersion = rs("ID")
				strversion = rs("Version") & ""
				if trim(rs("Revision")) <> "" then
					strversion = strVersion & "," & rs("Revision")
				end if
				if trim(rs("Pass")) <> "" then
					strversion = strVersion & "," & rs("Pass")
				end if
				
				strDistribution = ""
				if rs("Preinstall") then
					strDistribution = strDistribution & ", Preinstall"
				end if
				if rs("Preload") then
					strDistribution = strDistribution & ", Preload"
				end if
				if rs("Web") then
					strDistribution = strDistribution & ", Web"
				end if
				if rs("DropInBox") then
					strDistribution = strDistribution & ", DIB"
				end if
				if rs("SelectiveRestore") then
					strDistribution = strDistribution & ", SelectiveRestore"
				end if
				if rs("DRCD") then
					strDistribution = strDistribution & ", DRCD"
				end if
				if rs("DRDVD") then
					strDistribution = strDistribution & ", DRDVD"
				end if
				if trim(rs("Patch") & "") <> "0" then
					strDistribution = strDistribution & ", Patch"
				end if
				if strDistribution <> "" then
					strDistribution = mid(strDistribution,3)
				end if

				
				
				if rs("Lead") then
					TargetedLead = rs("Targeted") & ""
					DistributionLead = strDistribution
					TargetNotesLead = rs("TargetNotes") &""
					ImageSummaryLead = rs("ImageSummary") & ""
				else
					TargetedProduct = rs("Targeted") & ""
					DistributionProduct = strDistribution
					TargetNotesProduct = rs("TargetNotes") & ""
					ImageSummaryProduct = rs("ImageSummary") & ""
				end if
				rs.MoveNext
			loop
			rs.Close
			if VersionBGColor = "ivory" then
				VersionBGColor = "gainsboro"
			else					
				VersionBGColor = "ivory"
			end if
			Response.Write "<TR bgcolor=" & VersionBGColor & ">"
			Response.write "<TD valign=top rowspan=2>" & strVersion & "</TD>"
			Response.write "<TD>" & strLeadName & "</TD>"
			if trim(TargetedLead) = "" then
				response.write "<td colspan=4 bgcolor=Thistle><b>Not Supported</b></td>"
			else
				if TargetedLead <> TargetedProduct then
					Response.write "<TD><b><font color=red>" & TargetedLead & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & TargetedLead & "&nbsp;</TD>"
				end if
				if trim(DistributionLead) <> trim(DistributionProduct) then
					Response.write "<TD><b><font color=red>" & DistributionLead & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & DistributionLead & "&nbsp;</TD>"
				end if
				if trim(TargetNotesLead) <> trim(TargetNotesProduct) then
					Response.write "<TD><b><font color=red>" & TargetNotesLead & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & TargetNotesLead & "&nbsp;</TD>"
				end if
				if trim(ImageSummaryLead) <> trim(ImageSummaryProduct) then
					if trim(ImageSummaryLead)= "" then
						Response.write "<TD><b><font color=red>All&nbsp;</font></b></TD>"
					else
						Response.write "<TD><b><font color=red>" & ImageSummaryLead & "&nbsp;</font></b></TD>"
					end if
				else
					if trim(ImageSummaryLead)= "" then
						Response.write "<TD>All&nbsp;</TD>"
					else
						Response.write "<TD>" & ImageSummaryLead & "&nbsp;</TD>"
					end if
				end if
			end if
			Response.Write "</TR>"
			Response.Write "<TR bgcolor=" & VersionBGColor & ">"
			Response.write "<TD>" & strProductName & "</TD>"
			if trim(TargetedProduct) = "" then
				response.write "<td colspan=4 bgcolor=Thistle><b>Not Supported</b></td>"
			else
				if TargetedLead <> TargetedProduct then
					Response.write "<TD><b><font color=red>" & TargetedProduct & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & TargetedProduct & "&nbsp;</TD>"
				end if
				if trim(DistributionLead) <> trim(DistributionProduct) then
					Response.write "<TD><b><font color=red>" & DistributionProduct & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & DistributionProduct & "&nbsp;</TD>"
				end if
				if trim(TargetNotesLead) <> trim(TargetNotesProduct) then
					Response.write "<TD><b><font color=red>" & TargetNotesProduct & "&nbsp;</font></b></TD>"
				else
					Response.write "<TD>" & TargetNotesProduct & "&nbsp;</TD>"
				end if
				if trim(ImageSummaryLead) <> trim(ImageSummaryProduct) then
					if trim(ImageSummaryProduct)= "" then
						Response.write "<TD><b><font color=red>All&nbsp;</font></b></TD>"
					else
						Response.write "<TD><b><font color=red>" & ImageSummaryProduct & "&nbsp;</font></b></TD>"
					end if
				else
					if trim(ImageSummaryProduct)= "" then
						Response.write "<TD>All&nbsp;</TD>"
					else
						Response.write "<TD>" & ImageSummaryProduct & "&nbsp;</TD>"
					end if
				end if
			end if
			Response.Write "</TR>"
			Response.Write "</TABLE>"
		
		end if

	end if


%>


<%
	set rs = nothing
	cn.Close
	set cn = nothing
%>
</BODY>
</HTML>
<%
	function ScrubSQL(strWords) 

		dim badChars 
		dim newChars 
		dim i
		
		strWords=replace(strWords,"'","''")
		
		badChars = array("select", "drop", ";", "--", "insert", "delete", "xp_", "union", "=", "update") 
		newChars = strWords 
		
		for i = 0 to uBound(badChars) 
			newChars = replace(newChars, badChars(i), "") 
		next 
		
		ScrubSQL = newChars 
	
	end function 

%>