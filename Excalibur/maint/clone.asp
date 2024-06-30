<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<%

	Server.ScriptTimeout = 5400
	
	function NewImages(strImages)
		if trim(strImages)="" or isnull(strIMages) then
			NewImages = ""
		else
			dim IDArray	
			dim i
			dim j
			
			IDArray = split(strImages,",")

			for i = lbound(IDArray) to ubound(IDArray) 
				if trim(IDArray(i)) <> "" then
					for j = lbound(OldImageArray) to ubound(OldImageArray)
						if trim(OldImageArray(j)) =  trim(IDArray(i)) then
							NewImages=NewImages & ", " & trim(NewImageArray(j))
							exit for
						end if
					next
				end if
			next 
			if NewImages <> "" then
				NewImages = trim(mid(NewIMages,3)) & ","
			else
				NewImages=""
			end if
		end if
	
	end function

	dim cn
	dim rs
	dim blnErrors
	dim RowsUpdated
	dim OldProduct
	dim NewProduct
	dim OldImageIDList
	dim NewImageIDList
	dim OldImageArray
	dim NewImageArray
	dim i

'oldProduct=1225
'NewProduct=1246
'NOTE: Update tmpGetImageCopyID if you only need to copy some images
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	cn.CommandTimeout = 5400
	set rs = server.CreateObject("ADODB.recordset")

	OldImageIDList = ""
	NewImageIDList = ""
	rs.Open "tmpGetImageCopyID " & OldProduct & "," & NewProduct,cn,adOpenForwardOnly
	do while not rs.EOF
		OldImageIDList = OldImageIDList & "," & rs("FromID")
		NewImageIDList = NewImageIDList & "," & rs("ToID")
		rs.MoveNext
	loop
	rs.Close
	
	if OldImageIDList <> "" then
		OldImageArray = split(mid(OldImageIDList,2),",")
		NewImageArray = split(mid(NewImageIDList,2),",")
	end if

	strSQl = "Select pd.Preinstall, pd.PreinstallBrand,pd.Preload, pd.PreloadBrand, pd.Patch, pd.Web,pd.DropInBox,pd.ARCD,pd.DRDVD, pd.RACD_Americas, pd.RACD_EMEA, pd.RACD_APD,pd.DocCD, pd.OSCD, pd.SelectiveRestore,  pd.TargetNotes, pd.ImageSummary, pd.Images, pd.Deliverablerootid as ID from product_DelRoot pd with (NOLOCK), deliverableroot r with (NOLOCK) where r.id = pd.deliverablerootid and r.typeid <> 1 and productversionID = " & OldProduct

	'cn.BeginTrans
	
	rs.Open strSQl,cn,adOpenStatic
	Response.Write "<TABLE border=1>"
	blnErrors = false
	do while not rs.EOF
	    if OldImageIDList <> "" then
		    strNewImages=NewImages(rs("Images") & "")
		else
		    strNewImages = ""
		end if
		Response.Write "<TR><TD>" & rs("ID") & "</TD><TD>" & rs("Images") & "<HR>" & strNewImages & "</TD><TD>" & rs("ImageSummary") & "</TD></TR>"

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		cm.CommandText = "tmpProductCloneRoot"	

		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = NewProduct
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RootID", 3,  &H0001)
		p.Value = rs("ID")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Preinstall", 16,  &H0001)
		if rs("Preinstall") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PreinstallBrand", 200,  &H0001,50)
		p.Value = rs("PreinstallBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Preload", 16,  &H0001)
		if rs("Preload") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PreloadBrand", 200,  &H0001,50)
		p.Value = rs("PreloadBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DropInBox", 16,  &H0001)
		if rs("DropInBox") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Web", 16,  &H0001)
		if rs("Web") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Patch", 16,  &H0001)
		if trim(rs("Patch") & "") <> "0" then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ARCD", 16,  &H0001)
		if rs("ARCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DRDVD", 16,  &H0001)
		if rs("DRDVD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_Americas", 16,  &H0001)
		if rs("RACD_Americas") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_EMEA", 16,  &H0001)
		if rs("RACD_EMEA") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_APD", 16,  &H0001)
		if rs("RACD_APD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DocCD", 16,  &H0001)
		if rs("DocCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSCD", 16,  &H0001)
		if rs("OSCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SelectiveRestore", 16,  &H0001)
		if rs("SelectiveRestore") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@TargetNotes", 200,  &H0001,256)
		p.Value = rs("TargetNotes") & ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Summary", 200,  &H0001,80)
		p.Value = rs("ImageSummary") & ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Images", 201,  &H0001,2147483647)
		p.Value = strNewImages
		cm.Parameters.Append p
		'Response.Flush
		cm.Execute RowsUpdated
		if RowsUpdated > 1 then
			blnErrors = true
			exit do
		end if
		set cm = nothing
		rs.MoveNext
	loop
	rs.Close
	
	Response.Write "</TABLE>"
	if blnErrors then
		'cn.rollbackTrans
		Response.Write "Failed<BR>"
	else
		'cn.CommitTrans
		Response.Write "OK"
	end if

	blnerrors = false
	Response.flush
'---------------------------------------------

	strSQl = "Select pd.PreinstallBrand, pd.PreloadBrand, pd.Preinstall,pd.Preload, pd.patch,pd.Web,pd.DropInBox,pd.ARCD,DRDVD, pd.RACD_Americas, pd.RACD_EMEA, pd.RACD_APD,pd.DocCD, pd.OSCD,pd.SelectiveRestore,pd.PartNumber, pd.TargetNotes, pd.targeted, pd.InImage, pd.ImageSummary, pd.Images, pd.Deliverableversionid as ID from product_Deliverable pd with (NOLOCK), deliverableversion v with (NOLOCK), deliverableroot r with (NOLOCK) where r.id = v.deliverablerootid and v.id = pd.deliverableversionid and r.typeid <> 1 and r.active=1 and v.active=1 and (pd.targeted=1 or pd.inimage=1) and pd.productversionID = " & OldProduct

	'cn.BeginTrans
	
	rs.Open strSQl,cn,adOpenStatic
	Response.Write "<TABLE border=1>"
	blnErrors = false
	cn.CommandTimeout = 10
	do while not rs.EOF
	    if OldImageIDList <> "" then
    		strNewImages=NewImages(rs("Images") & "")
	    else
    		strNewImages= ""
	    end if
	  	Response.Write "<TR><TD>" & rs("ID") & "</TD><TD>" & rs("Preinstall") & "</TD><TD>" & rs("Preload") & "</TD><TD>" & rs("DropInBox") & "</TD><TD>" & rs("Web") & "</TD><TD>" & rs("ARCD") & "</TD><TD>" & rs("SelectiveRestore") & "</TD><TD>" & rs("Targeted") & "</TD><TD>" & rs("InImage") & "</TD><TD>" & rs("TargetNotes") & "</TD><TD>" & rs("Images") & "<HR>" & strNewImages & "</TD><TD>" & rs("ImageSummary") & "</TD></TR>"

		set cm = server.CreateObject("ADODB.Command")
		cm.CommandType =  &H0004
		cm.ActiveConnection = cn
		
		
		cm.CommandText = "tmpProductClone"	

		Set p = cm.CreateParameter("@ProductID", 3,  &H0001)
		p.Value = NewProduct
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@VersionID", 3,  &H0001)
		p.Value = rs("ID")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Preinstall", 16,  &H0001)
		if rs("Preinstall") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PreinstallBrand", 200,  &H0001,50)
		p.Value = rs("PreinstallBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Preload", 16,  &H0001)
		if rs("Preload") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@PreloadBrand", 200,  &H0001,50)
		p.Value = rs("PreloadBrand")
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DropInBox", 16,  &H0001)
		if rs("DropInBox") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Web", 16,  &H0001)
		if rs("Web") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Patch", 16,  &H0001)
		if trim(rs("Patch") & "") <> 0 then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@ARCD", 16,  &H0001)
		if rs("ARCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DRDVD", 16,  &H0001)
		if rs("DRDVD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_Americas", 16,  &H0001)
		if rs("RACD_Americas") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_EMEA", 16,  &H0001)
		if rs("RACD_EMEA") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@RACD_APD", 16,  &H0001)
		if rs("RACD_APD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@DocCD", 16,  &H0001)
		if rs("DocCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@OSCD", 16,  &H0001)
		if rs("OSCD") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@SelectiveRestore", 16,  &H0001)
		if rs("SelectiveRestore") then
			p.Value = 1
		else
			p.Value = 0
		end if
		
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Targeted", 16,  &H0001)
		if rs("Targeted") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@InImage", 16,  &H0001)
		if rs("InImage") then
			p.Value = 1
		else
			p.Value = 0
		end if
		cm.Parameters.Append p
		
		Set p = cm.CreateParameter("@PartNumber", 200,  &H0001,50)
		p.Value = rs("PartNumber") & ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@TargetNotes", 200,  &H0001,256)
		p.Value = rs("TargetNotes") & ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Summary", 200,  &H0001,80)
		p.Value = rs("ImageSummary") & ""
		cm.Parameters.Append p

		Set p = cm.CreateParameter("@Images", 201,  &H0001,2147483647)
		p.Value = strNewImages
		cm.Parameters.Append p

		cm.commandtimeout=20
	 	cm.Execute RowsUpdated

		set cm = nothing		
		if RowsUpdated > 1 then
			blnErrors = true
			Response.write cn.Errors.count & ">" & rs("ID") & strNewImages
			exit do
		end if
		
		rs.MoveNext
	loop
	rs.Close
	
	Response.Write "</TABLE>"
	if blnErrors then
		'cn.RollbackTrans
		Response.Write "Failed"
	else
		'cn.CommitTrans
		Response.Write "OK"
	end if
	set rs= nothing
	set cn = nothing
	
%>

</BODY>
</HTML>