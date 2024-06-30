<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function window_onload() {
	if (typeof(txtPath) != "undefined")
		{
		window.open (txtPath.value);
		window.parent.close();
		}
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<font size=2 face=verdana>Generating Files.  Please wait...</font>
<%
	dim FileArray
	dim i
	dim strLang
	dim strSoftpaq
	dim strSupersedes
	dim LangArray
	dim strFilename
	
	if request("txtFilename") <> "" then
		strFilename = request("txtFilename")
	else
		strFilename = request("txtID")
	end if
	
	Dim fs,f
	dim strBuffer
	Set fs=Server.CreateObject("Scripting.FileSystemObject")

	FileArray = split(request("txtSoftpaqList"),",")
	for i = lbound(FileArray) to ubound(filearray)
		LangArray = split(FileArray(i),"=")

		if not fs.FolderExists("c:\Program Office\SoftpaqFiles\" & strFilename) then
			fs.CreateFolder("c:\Program Office\SoftpaqFiles\" & strFilename) 
		end if
		if not fs.FolderExists("c:\Program Office\SoftpaqFiles\" & strFilename & "\" & LangArray(0)) then
			fs.CreateFolder("c:\Program Office\SoftpaqFiles\" & strFilename & "\" & LangArray(0)) 
		end if
		
		'generate Softpaq File
		Set f = fs.CreateTextFile("c:\Program Office\SoftpaqFiles\" & strFilename & "\" & LangArray(0) & "\" & LangArray(1) & ".txt",true,true)
		if LangArray(0) = "LI" then
			f.WriteLine(request("txtPreview"))
		else
			strBuffer = request("txtPreview")
			strBuffer = replace(strBuffer,"SOFTPAQ NUMBER: [VARIES BY LANGUAGE]","SOFTPAQ NUMBER: " & LangArray(1))
			strBuffer = replace(strBuffer,"SUPERSEDES: [VARIES BY LANGUAGE]","SUPERSEDES: " & LangArray(2))
			f.WriteLine(strBuffer)
		end if
		f.close

		'generate Welcome File
		Set f = fs.CreateTextFile("c:\Program Office\SoftpaqFiles\" & strFilename & "\" & LangArray(0)  & "\welcome.txt",true,true)
		if LangArray(0) = "LI" then
			f.WriteLine(request("txtPreviewWelcome"))
		else
			strBuffer = request("txtPreviewWelcome")
			strBuffer = replace(strBuffer,"SOFTPAQ NUMBER: [VARIES BY LANGUAGE]","SOFTPAQ NUMBER: " & LangArray(1))
			strBuffer = replace(strBuffer,"SUPERSEDES: [VARIES BY LANGUAGE]","SUPERSEDES: " & LangArray(2))
			f.WriteLine(strBuffer)
		end if
		f.close
		
		'generate CVA File
		Set f = fs.CreateTextFile("c:\Program Office\SoftpaqFiles\" & strFilename & "\" & LangArray(0) & "\" & LangArray(1) & ".cva",true,true)
		if LangArray(0) = "LI" then
			f.WriteLine(request("txtPreviewCVA"))
		else
			strBuffer = request("txtPreviewCVA")
			strBuffer = replace(strBuffer,"SoftpaqNumber=[VARIES BY LANGUAGE]","SoftpaqNumber=" & LangArray(1))
			strBuffer = replace(strBuffer,"SupercededSoftpaqNumber=[VARIES BY LANGUAGE]","SupercededSoftpaqNumber=" & LangArray(2))
			f.WriteLine(strBuffer)
		end if
		f.close

	next
	
	set f=nothing
	set fs=nothing

%>
<INPUT type="hidden" id=txtPath name=txtPath value="<%="\\" & Application("Excalibur_File_Server") & "\progoffice\SoftpaqFiles\" & strFilename%>">
</BODY>
</HTML>
