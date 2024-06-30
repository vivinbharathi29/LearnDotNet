<%@ Language=VBScript %>
<!-- #include file = "../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->
<!-- #include file="../includes/DataWrapper.asp" -->
<!-- #include file = "../includes/Security.asp" --> 
<%
  dim rs
  dim cn
  dim firstRow
  dim businessID
  dim familyID 
  dim geoID
  dim LookupID
  dim sFunction	            : sFunction = Request.Form("hidFunction")
  dim sAvFeatureCategoryID	: sAvFeatureCategoryID = Request.Form("hidAvFeatureCategoryID")
  dim sSelectedBusinessID	: sSelectedBusinessID = Request.Form("hidBusinessID")
  dim sSubassemblyID    	: sSubassemblyID = Request.Form("hidSubassemblyID") 
  dim AvCategoriesVisible
  dim RegionsVisible
  dim sExistingRegions    : sExistingRegions = Request.Form("hidExistingRegions") 
  dim sAllRegions         : sAllRegions = Request.Form("hidAllRegions")    
  
  If Request("FeatureCategoryID") <> "" Then
    sAvFeatureCategoryID = Request("FeatureCategoryID")
    sSubassemblyID = Request("SubassemblyID")
    AvCategoriesVisible = "none"
    RegionsVisible = ""
  End If

  set cn = server.CreateObject("ADODB.Connection")
  cn.ConnectionString = Session("PDPIMS_ConnectionString")
  cn.CommandTimeout = 20
  cn.Open
  
' Create Security Object to get User Info
Dim Security
	
Set Security = New ExcaliburSecurity
strUserID = Security.CurrentUserID()
strUserName = Security.CurrentUserFullName()
	
Set Security = Nothing
	 
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

If LCase(sFunction) = "save" Then
    Dim txtSuccess
    Dim ErrorCount
    'Dim AvCategoryID : AvCategoryID = ""
    'Dim PartNumberID : PartNumberID = ""
    Dim SelectedRegionIDs : SelectedRegionIDs = ""
    Dim arSeletedRegions
    Dim arAllRegions
    For Each item in Request.Form
        If instr(item, "chkRegions") > 0 Then
                SelectedRegionIDs = SelectedRegionIDs & "," & Replace(item, "chkRegions", "")
       End If
    Next
    SelectedRegionIDs = SelectedRegionIDs + ","
    
    arSeletedRegions = split(SelectedRegionIDs,",")
    For i = lbound(arSeletedRegions) To ubound(arSeletedRegions)
      If RTRIM(LTRIM(arSeletedRegions(i))) <> "" Then
        set cm = server.CreateObject("ADODB.connection")
	    cm.ConnectionString = Session("PDPIMS_ConnectionString")
	    cm.Open
	    cm.BeginTrans()
	    cm.Execute "usp_InsertSubassemblyMapping " & clng(sAvFeatureCategoryID) & "," & clng(sSubassemblyID) & "," & clng(arSeletedRegions(i))
	    cm.CommitTrans() 
	    set cm=nothing
	   'Response.Write(sAvFeatureCategoryID & "," & sSubassemblyID & "," & arSeletedRegions(i))
	   'Response.Write("<BR>")
	 End If
    Next
    
    set rs = server.CreateObject("ADODB.recordset")
    rs.Open "usp_SelectAllRegionsByBusiness " ,cn,adOpenForwardOnly
    do while not rs.EOF
	    sAllRegions = sAllRegions & "," & trim(rs("ID"))
	    rs.MoveNext
	loop
	sAllRegions = sAllRegions & ","
	rs.Close
	
    'Response.Write(sAllRegions)
    arAllRegions = split(sAllRegions,",")
    For i = lbound(arAllRegions) To ubound(arAllRegions)
        If instr(SelectedRegionIDs, "," & arAllRegions(i) & ",") = 0  And RTRIM(LTRIM(arAllRegions(i))) <> "" Then
            set cm = server.CreateObject("ADODB.connection")
	        cm.ConnectionString = Session("PDPIMS_ConnectionString")
	        cm.Open
	        cm.BeginTrans()
	        'Response.Write("Region: " & arAllRegions(i))
	        cm.Execute "usp_DeleteSubassemblyMapping " & clng(sAvFeatureCategoryID) & "," & clng(sSubassemblyID) & "," & clng(arAllRegions(i)) 
	        cm.CommitTrans()
	        set cm=nothing
	  'Response.Write(sAvFeatureCategoryID & "," & sSubassemblyID & "," & arAllRegions(i))
	  'Response.Write("<BR>")
        End If
    Next
    txtSuccess = "YES"
End If
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    var SelectedAvCategory;
    var SelectedBusinessID;
    var SelectedPartNo;
    var SelectedRegions;
    var iStep;
    iStep = 1;

    function ProcessStep() {
        if (iStep == 1) { //AvFeatureCategory
            divAvCategories.style.display = "";
            divAvSaDescriptions.style.display = "none";
            divRegions.style.display = "none";
            window.parent.frames["LowerWindow"].cmdPrevious.disabled = true;
            window.parent.frames["LowerWindow"].cmdNext.disabled = false;
            window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
        }
        else if (iStep == 2) { //PartNo 
            divAvCategories.style.display = "none";
            divAvSaDescriptions.style.display = "";
            divRegions.style.display = "none";
            window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
            window.parent.frames["LowerWindow"].cmdNext.disabled = false;
            window.parent.frames["LowerWindow"].cmdFinish.disabled = true;
        }
        else if (iStep == 3) { //Regions
            divAvCategories.style.display = "none";
            divAvSaDescriptions.style.display = "none";
            divRegions.style.display = "";
            window.parent.frames["LowerWindow"].cmdPrevious.disabled = false;
            window.parent.frames["LowerWindow"].cmdNext.disabled = true;
            window.parent.frames["LowerWindow"].cmdFinish.disabled = false;
        }
    }
    
    function body_onload(pulsarplusDivId) {
        var txtSuccess = document.getElementById("txtSuccess");

        if (txtSuccess.value == "YES") {
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                parent.window.parent.reloadFromPopUp(pulsarplusDivId);
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.closeExternalPopup();
            }
            else {
                window.parent.opener = 'X';
                window.parent.open('', '_parent', '')
                window.parent.close();
            }
       }
    }

    function ValidateAvCategorySelection() {
        SelectedAvCategory = "";
        SelectedBusinessID = "";
        var i;
        var checkBoxes = document.getElementsByTagName("input");

        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].id == "chkAv" && checkBoxes[i].checked == true) {
                if (SelectedAvCategory == "") {
                    SelectedBusinessID = checkBoxes[i].className;
                    //alert(SelectedBusinessID);
                    SelectedAvCategory = checkBoxes[i].name.substring(5);
                    frmAddNew.hidBusinessID.value = SelectedBusinessID;
                    frmAddNew.hidAvFeatureCategoryID.value = SelectedAvCategory;
               } else {
               SelectedAvCategory = SelectedAvCategory + "," + checkBoxes[i].name.substring(5);
               }
          }
      }
 }
 
  function ValidatePartNoSelection() {
        SelectedPartNo = "";
        var i;
        var checkBoxes = document.getElementsByTagName("input");

        for (i = 0; i < checkBoxes.length; i++) {
            if (checkBoxes[i].id == "chkPartNo" && checkBoxes[i].checked == true) {
                SelectedPartNo = checkBoxes[i].name.substring(9);
                frmAddNew.hidSubassemblyID.value = SelectedPartNo;
          }
      }
  }

  function ValidateRegionsSelection() {
      SelectedRegions = "";
      var i;
      var checkBoxes = document.getElementsByTagName("input");

      for (i = 0; i < checkBoxes.length; i++) {
          if (checkBoxes[i].id == "chkRegions" && checkBoxes[i].checked == true) {
              if (SelectedRegions == "") {
                  SelectedRegions = checkBoxes[i].name.substring(10);
              } else {
              SelectedRegions = SelectedRegions + "," + checkBoxes[i].name.substring(10);
              }
          }
      }
  }

  function PopulateExistingRegions() {
      var i;
      var checkBoxes = document.getElementsByTagName("input");
      var ExistingRegions;
      ExistingRegions = frmAddNew.hidExistingRegions.value;
      //alert(ExistingRegions);
      for (i = 0; i < checkBoxes.length; i++) {
          if (checkBoxes[i].id == "chkRegions" && ExistingRegions.contains(checkBoxes[i].name.substring(10)) == true) {
                  checkBoxes[i].checked = true;
              } else {
                  checkBoxes[i].checked = false;
              }
          }
      }

 function PopulatePartNoSelection(PartNo) {
     var i;
     var checkBoxes = document.getElementsByTagName("input");
     var checkBoxName;
     for (i = 0; i < checkBoxes.length; i++) {
           checkBoxName = checkBoxes[i].name
           if (checkBoxes[i].id == "chkPartNo" && checkBoxName.substring(9) == PartNo) {
             checkBoxes[i].checked = true;
         } else {
             checkBoxes[i].checked = false;
         }
     }
 }

 function PopulateAvFeatureCategortySelection(AvFeatureCategory) {
     var i;
     var checkBoxes = document.getElementsByTagName("input");
     var checkBoxName;
     for (i = 0; i < checkBoxes.length; i++) {
         checkBoxName = checkBoxes[i].name
         if (checkBoxes[i].id == "chkAv" && checkBoxName.substring(5) == AvFeatureCategory) {
             checkBoxes[i].checked = true;
         } else {
             checkBoxes[i].checked = false;
         }
     }
 }

  function chkAv_onclick() {
      var i;
      var checkBoxes = document.getElementsByTagName("input");
      var chkAv = window.event.srcElement.name;
      for (i = 0; i < checkBoxes.length; i++) {
          if (checkBoxes[i].id == "chkAv" && checkBoxes[i].name != chkAv) {
              checkBoxes[i].checked = false;   
          }
      }

  }

  function chkPartNo_onclick() {
      var i;
      var checkBoxes = document.getElementsByTagName("input");
      var chkPartNo = window.event.srcElement.name;
      for (i = 0; i < checkBoxes.length; i++) {
          if (checkBoxes[i].id == "chkPartNo" && checkBoxes[i].name != chkPartNo) {
              checkBoxes[i].checked = false;
          }
      }

  }
    
 </SCRIPT>
</HEAD>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
<STYLE>
TH
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}
TD
{
    FONT-SIZE: xx-small;
    FONT-FAMILY: Verdana
}

.ImageTable TBODY TD{
	BORDER-TOP: gray thin solid;
}

.imagerows TBODY TD{
	BORDER-TOP: none;
}

.imagerows THEAD TD{
	BORDER-TOP: none;
}

A:visited
{
    COLOR: blue
}

A:hover
{
    COLOR: red
}

</STYLE>
<body bgcolor="Ivory" onload="body_onload('<%Request("pulsarplusDivId")%>')">

<BR>
<form ID=frmAddNew method=post>
<input id="hidFunction" name="hidFunction" type=HIDDEN value=<%= LCase(sFunction)%>>
<input id="hidAvFeatureCategoryID" name="hidAvFeatureCategoryID" type=HIDDEN value=<%= LCase(sAvFeatureCategoryID)%>>
<input id="hidBusinessID" name="hidBusinessID" type=HIDDEN value=<%= LCase(sSelectedBusinessID)%>>
<input id="hidSubassemblyID" name="hidAvSubassemblyID" type=HIDDEN value=<%= LCase(sSubassemblyID)%>>
<input id="hidExistingRegions" name="hidExistingRegions" type=HIDDEN value=<%= LCase(sExistingRegions)%>>
<input id="hidAllRegions" name="hidAllRegions" type=HIDDEN value=<%= LCase(sAllRegions)%>>
<input type="hidden" id=txtSuccess name=txtSuccess value="<%= txtSuccess %>">
<%
    set rs = server.CreateObject("ADODB.recordset")
    rs.Open "usp_SelectAvFeatureCategoriesForNewSubassembly " & clng(request("SAType")),cn,adOpenForwardOnly
    businessID = 0
    firstRow = 0
    
   If Request("FeatureCategoryID") <> "" Then 
 %>
    <div id="divAvCategories" style="display:none">
 <%Else%>
    <div id="divAvCategories" style="display:""">
 <% 
 End If  
 
  If Request("FeatureCategoryText") <> "" Then %>
    <table id="PreviousSelection" width=100% border=0 cellspacing=0 cellpadding=1 >
        <thead bgcolor=Wheat>
            <th align=left></TD>
			<th align=left>AV Feature Category</th>
			<th align=left>SA Description</th>
		</thead>
		<tbody>
		    <tr>
                <TD><%=Request("FeatureCategoryText")%></TD>
			    <TD><%=Request("ProductNoText")%></TD>
			</tr>
        </tbody>
    </table>
 <%End If %>
 
<h3 align="center"><TH align=center>Select A Subassembly AV Category</TH></h3>
    <table id="AVFeatureCategoryTable" class="AVFeatureCategoryTable" width=100% border=0 cellspacing=0 cellpadding=1 >
        <thead bgcolor=Wheat>
            <th align=left></TD>
			<th align=left>AV Feature Category</th>
			<th align=left>Business</th>
		</thead>
		<tbody>
			 <%do while not rs.EOF%>
			    <%If firstRow = 0 Then 
			        businessID = trim(rs("BusinessID"))
			        firstRow = 1 
			      End If
			      If businessID <> trim(rs("BusinessID")) Then 
			         businessID = trim(rs("BusinessID"))
			         Response.Write"<tr bgcolor=Wheat><TD>&nbsp;</TD><TD>&nbsp;</TD><TD>&nbsp;</TD></tr>"
			      End If%>
			    <tr>
			        <%str = "<TD><input type=""checkbox"" class=" & rs("BusinessID") & " id=""chkAv"" name=chkAv" & rs("AvFeatureCategoryID") & " style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkAv_onclick()""></TD>"
			        Response.Write(str)%>
			        <TD><%=rs("AvFeatureCategory")%></TD>
			        <TD><%=rs("BusinessName")%></TD>
			    </tr>
			    <%
			    rs.MoveNext
			 loop
			 %>
		</tbody>
    </table>
</div>

<%set rs = server.CreateObject("ADODB.recordset")
  rs.Open "usp_SelectAvSaLookupDescription " & clng(request("SAType")),cn,adOpenForwardOnly
  familyID = 0
  firstRow = 0 %>
  <div id="divAvSaDescriptions" style="display:none">
    <h3 align="center"><TH align=center>Select A Subassembly Part Number</TH></h3>
    <table id="AvSaDescriptions" class="AvSaDescriptions" width=100% border=0 cellspacing=0 cellpadding=1 >
        <thead bgcolor=Wheat>
            <th align=left></TD>
			<th align=left>Subassembly Descriptions</th>
			<th align=left>Part Number</th>
		</thead>
		<tbody>
			 <%do while not rs.EOF%>
			    <tr>
			        <%str = "<TD><input type=""checkbox"" id=""chkPartNo"" name=chkPartNo" & rs("DeliverableRootID") & " style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chkPartNo_onclick()""></TD>"
			        Response.Write(str)%>
			        <TD><%=rs("Name")%></TD>
			        <TD><%=rs("Subassembly")%></TD>
			    </tr>
			    <%
			    rs.MoveNext
			 loop
			 rs.Close%>
		</tbody>
    </table>
</div>
<%
If Request("BusinessID") <> "" Then
  If sAvFeatureCategoryID <> "" And sSubassemblyID <> "" Then
    set rs = server.CreateObject("ADODB.recordset")
    rs.Open "usp_SelectExistingAvSaRegions " & clng(sAvFeatureCategoryID) & "," & clng(sSubassemblyID),cn,adOpenForwardOnly
	do while not rs.EOF
	    sExistingRegions = sExistingRegions & "," & trim(rs("ID"))
	    rs.MoveNext
	loop
	sExistingRegions = sExistingRegions & ","
	rs.Close
  End If
  geoID = 0
  firstRow = 0 
  
  set rs = server.CreateObject("ADODB.recordset")
  rs.Open "usp_SelectAllRegionsByBusiness " ,cn,adOpenForwardOnly
 If Request("FeatureCategoryID") <> "" Then 
 %>
    <div id="divRegions" style="display:""">
 <%Else%>
    <div id="divRegions" style="display:none">
 <% 
 End If  
 %>
    <h3 align="center"><TH align=center>Select Subassembly Regions</TH></h3>
    <table id="Regions" class="Regions" width=100% border=0 cellspacing=0 cellpadding=1 >
        <thead bgcolor=Wheat>
            <th align=left></TD>
			<th align=left>Config</th>
			<th align=left>Region Name</th>
			<th align=left>Commercial</th>
			<th align=left>Consumer</th>
			<th align=right>*Americas*</th>
		</thead>
		<tbody>
			 <%do while not rs.EOF
			    sAllRegions = sAllRegions & "," & trim(rs("ID"))%>
			    <%If firstRow = 0 Then 
			        geoID = trim(rs("GeoID"))
			        firstRow = 1 
			      End If
			      If geoID <> trim(rs("GeoID")) Then 
			         geoID = trim(rs("GeoID"))
			         Response.Write"<tr bgcolor=Wheat><th align=left>&nbsp;</Th><th align=left>Config</Th><th align=left>Region Name</Th><th align=left>Commercial</Th><th align=left>Consumer</Th><th align=right>*" & rs("Geo") & "*</Th></tr>"
			      End If%>
			    <tr>
			        <%If (request("BusinessID") = "1" And rs("Commercial") = "True") Or (request("BusinessID") = "2" And rs("Consumer") = "True") Then
			            If instr(sExistingRegions, "," & trim(rs("ID")) & ",") > 0 Then
			                str = "<TD><input type=""checkbox"" id=""chkRegions"" name=chkRegions" & rs("ID") & " checked style=""WIDTH:16;HEIGHT:16"" ></TD>"
			            Else
    			             str = "<TD><input type=""checkbox"" id=""chkRegions"" name=chkRegions" & rs("ID") & " style=""WIDTH:16;HEIGHT:16"" ></TD>"
                        End If
			            Response.Write(str)
			          End If  
			        If (request("BusinessID") = "1" And rs("Commercial") = "True") Or (request("BusinessID") = "2" And rs("Consumer") = "True") Then%>
			            <TD><%=rs("OptionConfig")%></TD>
			            <TD><%=rs("DisplayName")%></TD>
			            <TD><%=rs("Commercial")%></TD>
			            <TD><%=rs("Consumer")%></TD>
			        <%End If %>
			    </tr>
			    <%
			    rs.MoveNext
			 loop
			 rs.Close
			 sAllRegions = sAllRegions & ","%>
		</tbody>
    </table>
</div>
<%Else
    Response.Write("<div id=""divRegions"" style=""display:none"">")
  End If%>
</form>

<%
	set rs= nothing
	set cn=nothing
%>

</body>
</HTML>


