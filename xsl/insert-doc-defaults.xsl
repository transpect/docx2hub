<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:extendedProps="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
  xmlns:w10="urn:schemas-microsoft-com:office:word" 
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"
  xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  exclude-result-prefixes="docx2hub mml tr dbk cp"
  version="2.0">

  <xsl:variable name="docRels-uri" as="xs:anyURI"
    select="if (doc-available(resolve-uri(concat($base-dir,'_rels/document2.xml.rels'))))
                then resolve-uri(concat($base-dir,'_rels/document2.xml.rels'))
                else resolve-uri(concat($base-dir,'_rels/document.xml.rels'))"/>
  <xsl:variable name="docRels" as="document-node(element(rel:Relationships))"
    select="document($docRels-uri)"/>
  
  <xsl:variable name="themes" as="document-node(element(a:theme))*"
    select="for $t in $docRels/rel:Relationships/rel:Relationship[@Type eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme']/@Target
            return 
              if(doc-available(resolve-uri(concat($container-base-uri, $t)))) 
              then document(resolve-uri(concat($container-base-uri, $t)))
              else document(resolve-uri($t, $base-dir))"/>

  <!-- theme support incomplete … -->
  <xsl:function name="tr:theme-font" as="xs:string">
    <xsl:param name="rFonts" as="element(w:rFonts)?"/>
    <xsl:param name="themes" as="document-node(element(a:theme))*"/>
    <xsl:choose>
      <xsl:when test="not($themes | $rFonts)">
        <xsl:sequence select="'Arial'"/>
      </xsl:when>
      <xsl:when test="$rFonts/@w:asciiTheme">
        <!-- minor font is for the bulk text (major is for the headings).
             Spec sez dat w:asciiTheme has precedence over w:ascii (I don’t find it now, and it wasn’t all clear there) -->
        <xsl:sequence select="($themes/a:theme/a:themeElements/a:fontScheme/a:minorFont/a:latin/@typeface)[1]"/>
      </xsl:when>
      <xsl:when test="not($rFonts/@w:ascii)">
        <xsl:sequence select="'Arial'"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$rFonts/@w:ascii"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:template match="w:styles" mode="insert-doc-defaults">
    <xsl:copy>
      <!-- Font name of the default text -->
      <xsl:variable name="normal" as="element(w:style)?"
        select="(
                  w:style[@w:type = 'paragraph'][@w:default = '1'],
                  w:style[w:name[@w:val = 'Normal']]
                )[1]"/>
      <xsl:variable name="default-font" as="xs:string"
        select="if ($normal/w:rPr/w:rFonts/@w:ascii)
                then $normal/w:rPr/w:rFonts/@w:ascii
                else tr:theme-font(
                      (
                        w:docDefaults/w:rPrDefault/w:rPr/w:rFonts,
                        w:docDefaults/w:rPrDefault/w:rFonts
                      )[1], $themes)"/>
      <!-- Font size of the default text -->
      <xsl:variable name="default-font-size" as="xs:string"
        select="if ($normal/w:rPr/w:sz/@w:val)
                then ($normal/w:rPr/w:sz/@w:val)[1]
                else '20'"/>
      <xsl:variable name="default-lang" as="xs:string?"
        select="if ($normal/w:rPr/w:lang/@w:val)
                then $normal/w:rPr/w:lang/@w:val
                else (
                        w:docDefaults/w:rPrDefault/w:rPr/w:lang/@w:val,
                        w:docDefaults/w:rPrDefault/w:lang/@w:val
                      )[1]"/>
      <xsl:if test="exists($default-lang)">
        <xsl:attribute name="xml:lang" select="$default-lang"/>
      </xsl:if>
      <xsl:apply-templates select="@*, * except w:latentStyles" mode="#current">
        <xsl:with-param name="default-font" select="$default-font" tunnel="yes"/>
        <xsl:with-param name="default-font-size" select="$default-font-size" tunnel="yes"/>
        <xsl:with-param name="default-lang" select="$default-lang" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:style[@w:type = 'paragraph']
                              [not(w:basedOn)]/w:rPr" mode="insert-doc-defaults">
    <xsl:param name="default-font" as="xs:string?" tunnel="yes"/>
    <xsl:param name="default-font-size" as="xs:string" tunnel="yes"/>
    <xsl:param name="default-lang" as="xs:string?" tunnel="yes"/>
    <xsl:copy>
      <xsl:apply-templates select="@*, *" mode="#current"/>
      <xsl:if test="not(w:sz) and $default-font-size">
        <w:sz w:val="{$default-font-size}"/>
      </xsl:if>
      <xsl:if test="not(w:rFonts) and $default-font">
        <w:rFonts w:ascii="{$default-font}"/>
      </xsl:if>
      <xsl:if test="not(w:lang) and $default-lang">
        <w:lang w:val="{$default-lang}"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="w:style[@w:type = 'paragraph']
                              [not(w:basedOn)]/w:rPr/w:lang[not(@w:val)]" mode="insert-doc-defaults">
    <xsl:param name="default-lang" as="xs:string?" tunnel="yes"/>
    <xsl:copy>
      <xsl:attribute name="w:val" select="$default-lang"/>
      <xsl:sequence select="@*"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:style[@w:type = 'paragraph']
                              [not(w:basedOn)]
                              [empty(w:rPr)]" mode="insert-doc-defaults">
    <xsl:param name="default-font" as="xs:string?" tunnel="yes"/>
    <xsl:param name="default-font-size" as="xs:string" tunnel="yes"/>
    <xsl:param name="default-lang" as="xs:string?" tunnel="yes"/>
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
      <w:rPr>
        <xsl:if test="$default-font-size">
          <w:sz w:val="{$default-font-size}"/>
        </xsl:if>
        <xsl:if test="$default-font">
          <w:rFonts w:ascii="{$default-font}"/>
        </xsl:if>
        <xsl:if test="$default-lang">
          <w:lang w:val="{$default-lang}"/>
        </xsl:if>
      </w:rPr>
    </xsl:copy>
  </xsl:template>
  
</xsl:stylesheet>
