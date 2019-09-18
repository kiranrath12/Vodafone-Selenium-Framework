<?xml version="1.0" encoding="ISO-8859-1"?>

<xsl:stylesheet version="1.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

<xsl:template match="/">
  <html>
  
 <SCRIPT language="JavaScript"><![CDATA[
 
 

 
function ExpandCollapse(varDecide)
{
	var elem = document.getElementsByTagName('tr');	
	for (var x = 0; x <= elem.length-1; x++)
   {
		if (elem(x).id == 'Steps' || elem(x).id == 'Data')
	{		if (varDecide == 'Expand') 
			{
				elem(x).style.display = '';
			}
			if (varDecide == 'Collapse') 
			{
				elem(x).style.display = 'none';
			}
			
	}	
   }
   
   var elem = document.getElementsByTagName('table');	
	for (var x = 0; x <= elem.length-1; x++)
   {
		if (elem(x).id == 'FailureDetails')
	{		if (varDecide == 'Expand') 
			{
				elem(x).style.display = '';
			}
			if (varDecide == 'Collapse') 
			{
				elem(x).style.display = 'none';
			}
			
	}	
   }
	var elemIMG = document.getElementsByTagName('IMG');	
	for (var x = 0; x <= elemIMG.length-1; x++)
   {
		if (elemIMG(x).id == 'PlusMinus')
	{		if (varDecide == 'Expand') 
			{
				elemIMG(x).src = "minus.gif";
			}
			if (varDecide == 'Collapse') 
			{
				elemIMG(x).src = "plus.gif";
			}
			
	}	
   }

} 
 
function Toggle(node)
{
	// Unfold the branch if it isn't visible
	//var elem = document.getElementById('SSS');
	//var output = elem.nextSibling.nodeValue;
	var td = node.parentNode
	//alert(node.tagName);
	//alert(par.tagName);
	tr = td.parentNode;
	trnxt = tr.nextSibling;
	//alert(trnxt.tagName);
	//alert(trnxt.style.display);
	tbody = tr.parentNode;
	//alert(tbody.nodeType);
	tbl=tbody.parentNode;
	//alert('Hi');
	//alert(tbl.nextSibling.tagName);
	tblnxt = tbl.nextSibling;
	//tblnxt.style.display = 'none';
	//abc=tbl.parentNode.childNodes;
	//alert(abc(5).nodeType);
	 //alert (elem.nextSibling.nodeValue);
	 //alert (node.parentNode.tagName);
	 //alert('Hi');
	 
	if (trnxt.style.display == 'none')
	{
		// Change the image (if there is an image)
		if (node.children.length > 0)
		{
			if (node.children.item(0).tagName == "IMG")
			{
				node.children.item(0).src = "minus.gif";
			}
		}

		//node.nextSibling.style.display = '';
		//document.getElementById('TC1').style.display = '';
		trnxt.style.display = '';
		
	}
	// Collapse the branch if it IS visible
	else
	{
		// Change the image (if there is an image)
		if (node.children.length > 0)
		{
			if (node.children.item(0).tagName == "IMG")
			{
				node.children.item(0).src = "plus.gif";
			}
		}

		//node.nextSibling.style.display = 'none';
		//document.getElementById('TC1').style.display = 'none';
		trnxt.style.display = 'none';
		
		//bgColor="#E0FFFF" ,bgColor="#ffffe0"
	}
//onLoad="Collapse()"
}

function ShowDetails(node)
{
	var detailsTable = node.nextSibling;
	detailsTable = detailsTable.nextSibling;
	
	
	if (detailsTable.style.display == 'none')
	{
		// Change the image (if there is an image)
		if (node.children.length > 0)
		{
			if (node.children.item(0).tagName == "IMG")
			{
				node.children.item(0).src = "minus.gif";
			}
		}

		
		detailsTable.style.display = '';
		
	}
	// Collapse the branch if it IS visible
	else
	{
		// Change the image (if there is an image)
		if (node.children.length > 0)
		{
			if (node.children.item(0).tagName == "IMG")
			{
				node.children.item(0).src = "plus.gif";
			}
		}

		
		detailsTable.style.display = 'none';
		
		//bgColor="#E0FFFF" ,bgColor="#ffffe0"
	}
//onLoad="Collapse()"
}


function ShowPop1(node, ImagePath)
{
  
  var oPopup = window.createPopup();
  var oPopupBody = oPopup.document.body;
  
  
  
  var T0 = "<div id='layer1' style='position:absolute; left:20; width:200; height:10'>";
  var T1 = "<table id = 'FailureDetails' bgColor='#ffffe0' border='1'>";
  var T2 = "<tr><td><b>Image:</b></td><td><a href=" + ImagePath + "> Image Link</a></td></tr>";
  var T3 ="</table>";
  var T4 = "</div>";
  
  oPopupBody.innerHTML = T0 + T1 + T2 + T3 + T4;
  
  oPopup.show(0, 0, 300,0);
  var realHeight = oPopupBody.scrollHeight;
  var realWidth = oPopupBody.scrollWidth;
  //alert(findPosX(node));
  oPopup.hide();
	
  //oPopup.show(findPosX(node), findPosY(node) + 20, 300,realHeight, document.body);
  oPopup.show(findPosX(node),250, realWidth,realHeight, document.body);
  }
  
  
function ShowPop(node, Description, Object, Name, Method, Arguments, ErrImagePath)
{
  
  var oPopup = window.createPopup();
  var oPopupBody = oPopup.document.body;
  
  
  
  var T0 = "<div id='layer1' style='position:absolute; left:20; width:200; height:10'>";
  var T1 = "<table id = 'FailureDetails' bgColor='#ffffe0' border='1'>";
  var T2 = "<tr><td><b>Description:</b></td><td>" + Description + "</td></tr>";
  var T3 = "<tr><td><b>Object:</b></td><td>" + Object + "</td></tr>";
  var T4 = "<tr><td><b>Name:</b></td><td>" + Name + "</td></tr>";	 
  var T5 = "<tr><td><b>Method:</b></td><td>" + Method + "</td></tr>";
  var T6 = "<tr><td><b>Arguments:</b></td><td>" + Arguments + "</td></tr>";
  var T7 = "<tr><td><b>Image:</b></td><td><a href=" + ErrImagePath + "> Image Link</a></td></tr>";
  var T8 ="</table>";
  var T9 = "</div>";
  
  oPopupBody.innerHTML = T0 + T1 + T2 + T3 + T4 + T5 + T6 + T7 + T8 + T9;
  
  oPopup.show(0, 0, 300,0);
  var realHeight = oPopupBody.scrollHeight;
  var realWidth = oPopupBody.scrollWidth;
  //alert(findPosX(node));
  oPopup.hide();
	
  //oPopup.show(findPosX(node), findPosY(node) + 20, 300,realHeight, document.body);
  oPopup.show(findPosX(node),250, realWidth,realHeight, document.body);
  }
  function findPosX(obj)
  {
    var curleft = 0;
    
    if(obj.offsetParent)
        while(1) 
        {
          curleft += obj.offsetLeft;
          if(!obj.offsetParent)
            break;
          obj = obj.offsetParent;
        }
    else if(obj.x)
        curleft += obj.x;
    return curleft;
  }
  function findPosY(obj)
  {
    var curtop = 0;
    if(obj.offsetParent)
        while(1)
        {
          curtop += obj.offsetTop;
          if(!obj.offsetParent)
            break;
          obj = obj.offsetParent;
        }
    else if(obj.y)
        curtop += obj.y;
    return curtop;
  }
  //<A onClick="ShowPop(this,'Item not found', '<xsl:value-of select='OBJECT'/>','cmbCity','Select','Cresit card')"><IMG SRC="plus.gif" id = "PlusMinus">--</IMG></A>
      ]]></SCRIPT>

  
  <body onLoad="ExpandCollapse('Collapse')"  >
 
    <h2>Test Case Results for <xsl:value-of select="RESULT/TITLE"/></h2>
	 <font size="2" color="blue" face = "Verdana"><A onClick="ExpandCollapse('Collapse')" align="right">Collapse All</A> ................
								       <A onClick="ExpandCollapse('Expand')" align="right">Expand All</A> ...............
	
	
	 <b><font color="brown">TestCases: <xsl:value-of select="RESULT/SUMMARY/TOTAL"/></font></b>,  
	 <b><font color="green">Pass: <xsl:value-of select="RESULT/SUMMARY/TOTALPASS"/></font></b>, 
	 <b><font color="red">Fail: <xsl:value-of select="RESULT/SUMMARY/TOTALFAIL"/></font></b>
	 
	 </font>
	 
	<table border="0" width="100%" cellSpacing="0" cellPadding="3">
		<tr bgColor="#C0C0C0" >    
			<th align="left" width="10%"><font size="2"  face = "Verdana">TCID </font></th>
			<th align="left" width="60%"><font size="2"  face = "Verdana">NAME</font></th>
			<th align="left"><font size="2"  face = "Verdana">RESULT</font></th>
		</tr>    
	</table>
  

  
    <xsl:for-each select="RESULT/TESTCASE">
	<table border="1" width="100%" cellSpacing="0" cellPadding="3" >
	<tr ><td width="10%"><A onClick="Toggle(this)"><IMG SRC="plus.gif" id = "PlusMinus">--<b><xsl:value-of select="TCID"/></b></IMG></A></td>
		<td width="60%" valign = "top"><b><font size="2" color="blue" face = "Verdana"><xsl:value-of select="NAME"/></font ></b></td>		
		<xsl:choose>
			<xsl:when test="RESULT='FAIL'">
				<td ><b><font size="2" color="red" face = "Verdana"><xsl:value-of select="RESULT"/></font ></b></td>
			</xsl:when>
			<xsl:when test="RESULT='NOT EXECUTED'">
				<td ><b><font size="2" color="blue" face = "Verdana"><xsl:value-of select="RESULT"/></font ></b></td>
			</xsl:when>
			<xsl:otherwise>
				<td >	<b><font size="2" color="green" face = "Verdana"><xsl:value-of select="RESULT"/></font ></b></td>
			</xsl:otherwise>
		</xsl:choose>
		
	</tr>
	
	<tr id='Steps'>	
	<td COLSPAN = "4" >
	
		
			
			<xsl:for-each select="STEPS">
					<table border="0" width="100%" cellSpacing="0" cellPadding="0" >
					<tr >
					<td  width="10%">--<A onClick="Toggle(this)"><IMG SRC="plus.gif" id = "PlusMinus">--<b> <xsl:value-of select="STEPNO"/></b></IMG></A></td>
					<td width="60%"><font color="blue">    |______</font><b><font size="2" color="brown" face = "Verdana"><xsl:value-of select="NAME"/></font></b></td>
					
					<xsl:choose>
						<xsl:when test="RESULT='Pass'">
							<td>  
								<a>
									<xsl:attribute name="onClick">
										javascript:ShowPop1(this,'<xsl:value-of select="IMAGE"/>');
									</xsl:attribute>
									<b><font size="2" color="red" face = "Verdana"><xsl:value-of select="RESULT"/>  </font ></b>
									
								</a>
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="IMAGE" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Path</font ></b>									
								</a>
								+
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="EXCEL" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Data</font ></b>									
								</a>
								
								

							</td>
						</xsl:when>
						<xsl:when test="RESULT='Fail'">
							<td>  
								<a>
									<xsl:attribute name="onClick">
										javascript:ShowPop(this,'<xsl:value-of select='DESC'/>', '<xsl:value-of select='OBJECT'/>','<xsl:value-of select="OBJNAME"/>','<xsl:value-of select="METHOD"/>','<xsl:value-of select="ARGUMENTS"/>','<xsl:value-of select="IMAGE"/>');
									</xsl:attribute>
									<b><font size="2" color="red" face = "Verdana"><xsl:value-of select="RESULT"/>  </font ></b>
									
								</a>
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="IMAGE" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Path</font ></b>									
								</a>
								+
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="EXCEL" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Data</font ></b>									
								</a>
								
								

							</td>
						</xsl:when>
						<xsl:when test="RESULT='Not Executed'">
							<td ><b><font size="2" color="brown" face = "Verdana"><xsl:value-of select="RESULT"/></font ></b>
							+
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="EXCEL" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Data</font ></b>									
								</a></td>
						</xsl:when>
						<xsl:otherwise>
							<td >	<b><font size="2" color="green" face = "Verdana"><xsl:value-of select="RESULT"/></font ></b>
							+
								<a>
									<xsl:attribute name="href">
									<xsl:value-of select="EXCEL" />
									</xsl:attribute>
									<xsl:attribute name="target">
										"_blank"
									</xsl:attribute>
									
									<b><font size="2" color="red" face = "Verdana">Data</font ></b>									
								</a></td>
						</xsl:otherwise>
					</xsl:choose>
					</tr>
					
					<tr id='Data'>
					<td COLSPAN = "5">					
						
							<table border="0" width ="100%" >
							<xsl:for-each select="VERIFICATION">
				
								<tr height = "0" bgcolor = "#CCFF99">
								<td width="15%" bgcolor ="white"> </td>
								<td width="30%"><b><font size="2"  face = "Verdana" color = "black"><xsl:value-of select="NAME"/> : </font></b></td>
								<td ><font size="2"  face = "Verdana"><xsl:value-of select="VALUE"/></font></td>
								</tr>
							</xsl:for-each>
							<tr ><td></td><td colspan = "2" bgcolor = "green"> </td> </tr>
							
							<xsl:for-each select="DATA">
				
								<tr height = "0" bgcolor = "#CCFF99">
								<td width="15%" bgcolor = "white"></td>
								<td width="30%"><b><font size="2"  face = "Verdana"><xsl:value-of select="NAME"/> : </font></b></td>
								<td ><font size="2"  face = "Verdana"><xsl:value-of select="VALUE"/></font></td>
								</tr>
							</xsl:for-each>
							</table>
					</td>
					</tr>
					</table>	
						
			</xsl:for-each>
	</td>
	</tr>		
		
	</table>
	
    </xsl:for-each>
    

  </body>
  </html>
</xsl:template>

</xsl:stylesheet>