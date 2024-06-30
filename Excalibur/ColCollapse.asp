<HTML>
<HEAD>
<!--<style type="text/css">
..ver{display:inline}
..nover{display:none}
</style>-->

<STYLE>
TD{
	Font-Size:xx-small;
}
</STYLE>

<script language="JavaScript">
	function ShowHide(cual, estilo) {
		var xx = eval(cual);
	
		if (xx.length)
			for (x=0;x<xx.length;x++)
				xx[x].className = estilo
		else 
			xx.className = estilo 
			

}

</script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {
//document.getElementsByName('C1') ;
alert(document.getElementsByTagName('STYLE').style);

}

//-->
</SCRIPT>
</HEAD>
<body>
<INPUT type="button" value="Button" id=button1 name=button1 LANGUAGE=javascript onclick="return button1_onclick()">
<form>
	<table border=1>
	<%for i = 1 to 2%>
		<tr>
			<td id=C1 name="C1">first</td>
			<td ID=C2 name="C2">second</td>
		</tr>
	<%next%>
	</table>

<input type="checkbox" name="checkbox" value="checkbox" onclick="ShowHide('C1', 'nover')">


<p>
  <button type="button" onclick="
   var x, opt, txt, tbl = document.getElementById('myTable');
 
   if (tbl.rows[2].style && tbl.rows[2].style.display == 'none') {
     opt = 'block';
     txt = 'Hide Rows 3 thru 5';
   } else {
     opt = 'none';
     txt = 'Show Rows 3 thru 5';
   }
   for (x=2; x<5; x++) {
     tbl.rows[x].style.display = opt;
   }
   this.value = txt;
   return true;">Hide Rows 3 thru 5</button>
</p>

<table id="myTable">
  <tr>
    <td>row1,col1</td>
    <td>row1,col2</td>
  </tr>
  <tr>
    <td>row2,col1</td>
    <td>row2,col2</td>
  </tr>
  <tr>
    <td>row3,col1</td>
    <td>row3,col2</td>
  </tr>
  <tr>
    <td>row4,col1</td>
    <td>row4,col2</td>
  </tr>
  <tr>
    <td>row5,col1</td>
    <td>row5,col2</td>
  </tr>
  <tr>
    <td>row6,col1</td>
    <td>row6,col2</td>
  </tr>
</table>




</body>
</HTML>

