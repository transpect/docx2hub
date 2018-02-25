<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"  xmlns:docx2hub="http://transpect.io/docx2hub"
  version="2.0" 
  exclude-result-prefixes="docx2hub">

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- CATCH-ALL (for all modes) -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:template match="node()|@*" mode="#all" priority="-1">
    <xsl:copy>
      <xsl:apply-templates select="@*"     mode="#current"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="*" mode="wml-to-dbk docx2hub:add-props" priority="-0.9">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, *" mode="#current"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="  w:t 
                       | m:t
                       | *[@xml:space eq 'preserve']
                       | w:instrText
                       | wp:posOffset
                       | wp14:pctWidth
                       | wp14:pctHeight" mode="wml-to-dbk docx2hub:add-props" priority="-0.8">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>  

</xsl:stylesheet>