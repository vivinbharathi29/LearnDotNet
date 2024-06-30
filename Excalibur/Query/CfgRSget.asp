<!--#include file="../_ScriptLibrary/jsrsServer.inc"-->
<% jsrsDispatch( "TemplateStrings" ) %>


<script runat="server" language="vbscript">


function TemplateStrings(ID)

on error resume next
  dim cn
  dim rs
  dim i

  set cn = server.createobject("ADODB.Connection")
  set rs = server.createobject("ADODB.Recordset")

  cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
  cn.open

  rs.open "up_tmt_ct_GetSummary " & clng(ID) ,cn,adOpenForwardOnly

  if rs.EOF and rs.BOF then
    TemplateStrings = "<input type=""hidden"" id=txtQTemplateName name=txtQTemplateName value="""">"
  else
    TemplateStrings = "<input type=""hidden"" id=txtQTemplateName name=txtQTemplateName value=""" & rs("Name") & """><br>"
    TemplateStrings = TemplateStrings & "<input type=""hidden"" id=ProductFamilyID name=ProductFamilyID value=""" & rs("ProductFamilyID") & """><br>"
    TemplateStrings = TemplateStrings & "<input type=""hidden"" id=ProductVersionID name=ProductVersionID value=""" & rs("ProductVersionID") & """><br>"
    'TemplateStrings = TemplateStrings & "<input type=""hidden"" id=DelRoots name=DelRoots value=""" & rs("DelRoots") & """><br>"
  end if
  rs.Close
  set rs = nothing
  cn.Close
  set cn=nothing

end function

</script>
