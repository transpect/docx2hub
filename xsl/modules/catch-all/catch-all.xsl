<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  version="2.0" 
  exclude-result-prefixes="docx2hub w">

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- CATCH-ALL (for all modes) -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:template match="node()|@*" mode="#all" priority="-1">
    <xsl:copy copy-namespaces="no">
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
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  

</xsl:stylesheet>