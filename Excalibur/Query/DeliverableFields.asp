<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
	

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<Title>Available Fields</Title>
    <script type="text/javascript" src="../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>

<!--



function window_onload() {
	//var s = window.dialogArguments;
	//txtSQL.value = s;
}




function cboSelect_onclick() {
	//window.returnValue = cboField.value;
    //window.close();

    if (IsFromPulsarPlus()) {
        window.parent.parent.BuildSqlResult(cboField.value);
        ClosePulsarPlusPopup();
    }
    else {
        window.returnValue = cboField.value;
        window.parent.close();
    }

}

//function cboSelect_onclick() {
//	window.returnValue = cboField.value;
//	window.close();

//}


function cboOp_onclick() {
	if (cboOp.options[cboOp.selectedIndex].text == "Is Null")
		{
		ValueRow.style.display = "none";
		txtValue.value = "";
		}
	else
		ValueRow.style.display = "";
}


function cboField_onclick() {

}

//-->
</SCRIPT>
</HEAD>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">

<%

	function SortArray(arrShort)
	
		for i = UBound(arrShort) - 1 To 0 Step -1
			for j= 0 to i
				if arrShort(j)>arrShort(j+1) then
					temp=arrShort(j+1)
					arrShort(j+1)=arrShort(j)
					arrShort(j)=temp
				end if
			next
		next 
		SortArray = arrShort
	end function



	dim cn
	dim rs
	dim i 
	dim strFields
	dim strValues
	
	strFields = "ID,Name,Version,Revision,Pass,VendorVersion,Location,DevManager,Developer,ImageSummary,Targeted,InImage,Languages,Vendor,Category,SelectiveRestore,ARCD,DRDVD,RACD_Americas,RACD_APD,RACD_EMEA,Preinstall,Preload,DropInBox,Web,Patch,OSCD,DOCCD,Product,TargetNotes"
	
	FieldArray = SortArray(Split(strFields,","))
	
	'for i = UBound(FieldArray) - 1 To 0 Step -1
	'	for j= 0 to i
	'		if FieldArray(j)>FieldArray(j+1) then
	'			temp=FieldArray(j+1)
	'			FieldArray(j+1)=FieldArray(j)
	'			FieldArray(j)=temp
	'		end if
	'	next
	'next 


	for i = lbound(FieldArray) to ubound(FieldArray) 
		if FieldArray(i) = "ID" then
			strFields = strFields & "<option value=""v." & FieldArray(i) & """>" & FieldArray(i) & "</option>" 
		else
			strFields = strFields & "<option value=""" & FieldArray(i) & """>" & FieldArray(i) & "</option>" 
		end if
	next

%>
<table border=0 width=100%><TR><TD width=8>&nbsp;</TD><TD>
<font size=2 face=verdana><b>Select Field Name</b><BR><BR></font>
<TABLE width="100%" cellspacing=0 cellpadding=2 bordercolor=tan bgcolor=ivory border=1><TR bgcolor=cornsilk><td width=80><font size=2 face=verdana><b>Field:</b></font></td>
<td>
<SELECT style="Width:250" id=cboField name=cboField LANGUAGE=javascript onclick="return cboField_onclick()"><%=strFields%></SELECT>
</td></tr>

</table>
<table width=100% border=0>
<TR>
<TD colspan=2 align=right>
	<INPUT type="button" value="Insert" id=cboSelect name=cboSelect LANGUAGE=javascript onclick="return cboSelect_onclick()">
</td></tr>
</table>
</td></tr></table>
</BODY>
</HTML>
