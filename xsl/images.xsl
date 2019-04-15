<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:tr="http://transpect.io"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "w v xs r rel tr"
  version="2.0">
  
  <xsl:template match="w:drawing[descendant::a:blip]" mode="wml-to-dbk">
    <mediaobject>
      <xsl:apply-templates select="@srcpath" mode="#current"/>
      <xsl:if test=".//wp:docPr/@title[. ne '']">
        <info>
          <xsl:value-of select=".//wp:docPr/@title[. ne '']"/>
        </info>
      </xsl:if>
      <xsl:if test=".//wp:docPr/@descr[. ne '']">
        <alt>
          <xsl:value-of select=".//wp:docPr/@descr[. ne '']"/>
        </alt>
      </xsl:if>
      <xsl:apply-templates select="descendant::a:blip" mode="wml-to-dbk"/>
    </mediaobject>
  </xsl:template>
  
 <xsl:template match="w:drawing[not(descendant::a:blip) or descendant::*:webVideoPr]" mode="wml-to-dbk">
    <phrase role="hub:foreign">
      <xsl:apply-templates select="." mode="foreign"/>
    </phrase>
  </xsl:template>
  
  <!-- images embedded in word zip container, usually stored in {docx}/word/media/ -->
  <xsl:template match="a:blip[@r:embed]" mode="wml-to-dbk">
    <xsl:call-template name="create-imageobject">
      <xsl:with-param name="image-id" select="@r:embed"/>
    </xsl:call-template>
  </xsl:template>
  
  <xsl:template match="pic:blipFill | a:srcRect" mode="wml-to-dbk"/>
  
  <!-- parent v:shape is processed in vml mode, see objects.xsl -->
  <xsl:template match="v:imagedata" mode="vml">
    <xsl:param name="inline" select="false()" tunnel="yes"/>
    <xsl:element name="{if ($inline) then 'inlinemediaobject' else 'mediaobject'}">
      <xsl:if test="ancestor::w:object">
        <xsl:attribute name="annotations" select="concat('object_',generate-id(ancestor::w:object[1]))"/>
      </xsl:if>
      <xsl:apply-templates select="@srcpath, parent::v:shape/@style" mode="#current"/>
      <xsl:if test="ancestor::v:shape[1]/@docx2hub:generated-alt">
        <alt>
          <xsl:value-of select="ancestor::v:shape[1]/@docx2hub:generated-alt"/>
        </alt>
      </xsl:if>
      <xsl:variable name="image-id" select="(@r:href, @r:id)[1]" as="xs:string?"/>
      <xsl:choose>
        <xsl:when test="normalize-space($image-id) = ''">
          <xsl:message select="'! Empty image-id found:', ."/>
        </xsl:when>
        <xsl:otherwise>
          <xsl:call-template name="create-imageobject">
            <xsl:with-param name="image-id" select="$image-id"/>
          </xsl:call-template>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:element>
  </xsl:template>

  <xsl:template match="@style" mode="vml" priority="4">
    <xsl:analyze-string select="." regex="\s*;\s*">
      <xsl:non-matching-substring>
        <xsl:analyze-string select="." regex="(.+)\s*:\s*(.+)">
          <xsl:matching-substring>
            <xsl:attribute name="css:{regex-group(1)}" select="regex-group(2)"/>
          </xsl:matching-substring>
        </xsl:analyze-string>
      </xsl:non-matching-substring>
    </xsl:analyze-string>
  </xsl:template>

  <!-- externally referenced images -->
  <xsl:template match="a:blip[@r:link]" mode="wml-to-dbk">
    <xsl:call-template name="create-imageobject">
      <xsl:with-param name="image-id" select="@r:link"/>
    </xsl:call-template>
  </xsl:template>
  
  <xsl:key name="docrel" match="rel:Relationship" use="@Id"/>
  
  <xsl:template name="create-imageobject">
    <xsl:param name="image-id" as="xs:string"/>
    <xsl:variable name="rels" as="element(rel:Relationships)"
      select="if (ancestor::w:footnote) 
              then $root/*/w:footnoteRels/rel:Relationships
              else 
                if (ancestor::w:comment) 
                then $root/*/w:commentRels/rel:Relationships 
                else 
                  if (ancestor::w:endnote) 
                  then $root/*/w:endnoteRels/rel:Relationships 
                  else 
                    $root/*/w:docRels/rel:Relationships"/>
    <xsl:variable name="rel" as="element(rel:Relationship)"
      select="$rels/rel:Relationship[@Id = $image-id]"/>
    <xsl:variable name="patched-file-uri" select="replace($rel/@Target, '\\', '/')" as="xs:string"/>
    <imageobject>
      <xsl:apply-templates select="../a:srcRect" mode="wml-to-dbk"/>
      <imagedata fileref="{if ($rel/@TargetMode = 'External') 
                           then $patched-file-uri
                           else concat('container:word/', $patched-file-uri)}">
        <xsl:apply-templates select="ancestor-or-self::w:drawing//wp:extent/@*" mode="wml-to-dbk"/>
      </imagedata>
    </imageobject>
  </xsl:template>
  
  <!-- despite ISO 29100-1, which states that the clipping coordinates are always percentages, you’ll find
    large numbers as in <a:srcRect l="57262" t="34190" r="7344" b="42485"/> 
    These are called emu. 1 pt = 12700 emu, 1 mm = 36000 emu  -->
  <xsl:template match="a:srcRect[@l][@t][@r][@b][every $a in (@l, @t, @r, @b) satisfies (matches($a, '^\d+$'))]" mode="wml-to-dbk" priority="2">
    <xsl:attribute name="css:clip" select="concat('rect(', string-join(for $a in (@t, @r, @b, @l) return  concat(round(number($a) * 0.0015748) * 0.05, 'pt'), ', '), ')')"/>
  </xsl:template>
  
  <xsl:template match="a:srcRect[not(@*)]" mode="wml-to-dbk">
    <!-- don’t know whether the empty a:srcRect conveys some meaning; just wanted to get rid of W2D_020 messages -->
  </xsl:template>
  
</xsl:stylesheet>
