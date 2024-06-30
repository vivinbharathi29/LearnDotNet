<%@ Language=VBScript %>
<%
option explicit
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function Test(){
        alert(AppendNewColumns("1,2,3,4","5,6,7"));
    }
    
    
    function AppendNewColumns(strSavedList,strMasterList){
        var SavedArray;
        var MasterArray;
        var i;
        var strOutput="";
        var strTemp="";
        
        SavedArray =  strSavedList.split(",");
        MasterArray = strMasterList.split(",");
        
        for (i=0;i< SavedArray.length;i++)
            if (SavedArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') != "")
                strOutput = strOutput + "," + SavedArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '');
 
        for (i=0;i< MasterArray.length;i++)
           if (MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') != "")
                {
                strTemp = "," + strOutput + "," 
                if (strTemp.indexOf("," + MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '') + ",") == -1)
                    strOutput = strOutput + "," + MasterArray[i].replace(/^\s\s*/, '').replace(/\s\s*$/, '')

                }
        if (strOutput != "")
            strOutput = strOutput.substring(1);

        return strOutput;
}    

//-->
</SCRIPT>
</HEAD>
<BODY onload="javascript:Test();">



</BODY>
</HTML>
