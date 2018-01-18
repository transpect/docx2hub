<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:xs="http://www.w3.org/2001/XMLSchema" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  exclude-result-prefixes="xs" 
  version="2.0">

  <xsl:param name="active"/>
  <xsl:param name="word-container-cleanup" select="'yes'"/>

  <xsl:variable name="ole-objects" as="element()*"
                select="if (ancestor::w:docRels)
                          then collection()[2]//w:document//o:OLEObject
                        else if (ancestor::w:footnoteRels)
                          then collection()[2]//w:footnotes//o:OLEObject
                        else if (ancestor::w:endnoteRels)
                          then collection()[2]//w:endnotes//o:OLEObject
                        else if (ancestor::w:commentRels)
                          then collection()[2]//w:comments//o:OLEObject
                        else ()"/>
  
  <xsl:variable name="image-wmf-objects" as="element()*" 
                select="if (ancestor::w:docRels) 
                          then collection()[2]//w:document//w:drawing//a:blip
                        else if (ancestor::w:footnoteRels)
                          then collection()[2]//w:footnotes//w:drawing//a:blip
                        else if (ancestor::w:endnoteRels)
                          then collection()[2]//w:endnotes//w:drawing//a:blip
                        else if (ancestor::w:commentRels)
                          then collection()[2]//w:comments//w:drawing//a:blip
                        else ()"/>

  <xsl:template match="rel:Relationship[@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject']">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:if test="
          $word-container-cleanup = 'yes' and
          matches($active, 'yes|wmf|ole') and
          not(@Id = $ole-objects/@r:id)">
        <xsl:attribute name="remove" select="'yes'"/>
      </xsl:if>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="rel:Relationship[@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image']">
    <xsl:copy>
      <xsl:apply-templates select="@*"/>
      <xsl:if test="
          $word-container-cleanup = 'yes' and
          contains($active, '+try-all-pict-wmf') and
          not(@Id = $image-wmf-objects/@r:embed)">
        <xsl:attribute name="remove" select="'yes'"/>
      </xsl:if>
      <xsl:apply-templates/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="node() | @*">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>