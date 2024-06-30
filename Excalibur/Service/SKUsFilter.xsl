<!--?xml version='1.0' encoding='UTF-8'?-->          
 <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"> 
                                      
   
    <xsl:template match = 'Table'>
       	<tr class="MatrixAvRow" style="height:20px;cursor: pointer;">
          <xsl:attribute name='id'><xsl:value-of select ="SKU"/>Row</xsl:attribute>
          
	      <td colspan="9">
          <a><xsl:attribute name='name' ><xsl:value-of select = 'SKU' /></xsl:attribute></a>
          
		      <span title="Click to Perform a QuickSearch on this SKU" style="font: bold x-small verdana; float:left">
            <xsl:attribute name="onclick">QuickSearch('<xsl:value-of select="SKU"/>');</xsl:attribute>
			      <xsl:value-of select = 'SKU' />
		      </span>
          <span>  <div title="Click to display the Spare Kits for this SKU." mode="0" onclick="javascript:showSKUSpareKits(this);" style="font: normal xx-small verdana; float:left">
            <xsl:attribute name='id'><xsl:value-of select = 'SKU' /></xsl:attribute>+ Show Spare Kits
            </div>
          </span>   
          
	      </td>
       </tr>
        <tr>
          <td>
            <xsl:attribute name='id'>[<xsl:value-of select='SKU'/>]</xsl:attribute>
          </td>
        </tr>
    </xsl:template>
    <xsl:template match = '/'>
         <table class='MatrixTable'>
           <xsl:apply-templates select = 'NewDataSet' />
         </table>
    </xsl:template>
</xsl:stylesheet>

