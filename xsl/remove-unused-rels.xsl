<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:xs="http://www.w3.org/2001/XMLSchema" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  exclude-result-prefixes="xs" 
  version="2.0">

  <xsl:param name="active"/>
  <xsl:param name="word-container-cleanup" select="'yes'"/>

  <xsl:variable name="former-ole-objects" as="attribute(docx2hub:rel-ole-id)*"
                select="collection()[2]//mml:math/@docx2hub:rel-ole-id"/>
  
  <xsl:variable name="former-image-wmf-objects" as="attribute(docx2hub:rel-wmf-id)*" 
                select="collection()[2]//mml:math/@docx2hub:rel-wmf-id"/>

  <xsl:template match="rel:Relationship[@Id = $former-ole-objects]">
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:if test="$word-container-cleanup = 'yes' 
                    and matches($active, 'yes|wmf|ole')">
        <xsl:attribute name="remove" select="'yes'"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="rel:Relationship[@Id = $former-image-wmf-objects]">
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:if test="$word-container-cleanup = 'yes' and
                    contains($active, '+try-all-pict-wmf')">
        <xsl:attribute name="remove" select="'yes'"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="*|@*">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>  
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>