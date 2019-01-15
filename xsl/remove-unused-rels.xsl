<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" 
  xmlns:xs="http://www.w3.org/2001/XMLSchema" 
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:v="urn:schemas-microsoft-com:vml"
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

  <xsl:template match="rel:Relationship[$word-container-cleanup eq 'yes' ]
                                        [@Type = ('http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject', 
                                                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')]                                                 
                                        [@Id = ($former-ole-objects, $former-image-wmf-objects)]
                                        [rel:find-rel-element-by-ref(., collection()[2]/w:root)]
                                        [not(rel:find-rel-element-by-non-mml-ref(., collection()[2]/w:root))]">
    <xsl:copy>
      <xsl:copy-of select="@*"/>
      <xsl:attribute name="remove" select="'yes'"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:function name="rel:find-rel-element-by-ref" as="xs:boolean">
    <xsl:param name="rel" as="element(rel:Relationship)"/>
    <xsl:param name="root" as="element(w:root)"/>
    <xsl:variable name="rel-element" as="element()*"
      select="if($rel/ancestor::*[2]/name() eq 'w:docRels')      then $root/w:document//mml:math[(@docx2hub:rel-wmf-id, @docx2hub:rel-ole-id) = $rel/@Id]
         else if($rel/ancestor::*[2]/name() eq 'w:footnoteRels') then $root/w:footnotes//mml:math[(@docx2hub:rel-wmf-id, @docx2hub:rel-ole-id) = $rel/@Id]
         else if($rel/ancestor::*[2]/name() eq 'w:endnoteRels')  then $root/w:endnotes//mml:math[(@docx2hub:rel-wmf-id, @docx2hub:rel-ole-id) = $rel/@Id]
         else if($rel/ancestor::*[2]/name() eq 'w:commentRels')  then $root/w:comments//mml:math[(@docx2hub:rel-wmf-id, @docx2hub:rel-ole-id) = $rel/@Id]
                                 else ()"/>
    <xsl:sequence select="exists($rel-element)"/>
  </xsl:function>
  
  <xsl:function name="rel:find-rel-element-by-non-mml-ref" as="xs:boolean">
    <xsl:param name="rel" as="element(rel:Relationship)"/>
    <xsl:param name="root" as="element(w:root)"/>
    <xsl:sequence select="exists($root/w:*//v:imagedata[@r:id = $rel/@Id])"/>
  </xsl:function>

  <xsl:template match="*|@*">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()"/>  
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>
