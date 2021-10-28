<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:cat="urn:oasis:names:tc:entity:xmlns:xml:catalog"
  exclude-result-prefixes="xs"
  version="2.0">
  
  <xsl:import href="../main.xsl"/>
  
  <xsl:param name="mml-space-handling" as="xs:string" select="'mspace'"/>
  
  <xsl:output method="xml" doctype-system="" doctype-public=""></xsl:output>

  <xsl:param name="fail-on-error" select="'no'"/>
  <xsl:param name="charmap-policy" select="'unicode'"/>
  
  <xsl:variable name="symbol-font-map" as="document-node(element(symbols))"
    select="document('http://transpect.io/fontmaps/Symbol.xml')"/>
  <xsl:variable name="catalog" as="document-node(element(cat:catalog))?">
    <xsl:sequence select="doc('http://this.transpect.io/xmlcatalog/catalog.xml')"/>
  </xsl:variable>

  <xsl:key name="symbol-by-number" match="symbol" use="upper-case(replace(@number, '^0*(.+?)$', '$1'))" />
  <xsl:key name="symbol-by-entity" match="symbol" use="@entity" />
  <xsl:key name="style-by-id" match="w:style" use="@w:styleId" />  
  
  <xsl:template match="m:oMathPara">
    <equation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:variable name="prelim" as="item()*">
        <xsl:apply-templates select="node()" mode="omml2mml"/>
      </xsl:variable>
      <xsl:apply-templates select="$prelim" mode="docx2hub:join-runs"/>
    </equation>
  </xsl:template>
  
  <xsl:template match="m:oMath">
    <inlineequation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:variable name="prelim" as="item()*">
        <xsl:apply-templates select="." mode="omml2mml"/>
      </xsl:variable>
      <xsl:apply-templates select="$prelim" mode="docx2hub:join-runs"/>
    </inlineequation>
  </xsl:template>
  
  <xsl:template match="w:sym" mode="omml2mml" priority="120">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>
  
  <xsl:function name="docx2hub:based-on-chain" as="document-node()">
    <xsl:param name="initial" as="element(w:style)*"/>
    <xsl:variable name="next" as="element(w:style)?" 
      select="if (exists($initial)) 
      then key('docx2hub:style', $initial[last()]/w:basedOn/@w:val, root($initial[last()]))
      else ()"/>
    <xsl:choose>
      <xsl:when test="exists($next)">
        <xsl:document>
          <xsl:sequence select="docx2hub:based-on-chain(($initial, $next))/*"/>  
        </xsl:document>
      </xsl:when>
      <xsl:otherwise>
        <xsl:document>
          <xsl:sequence select="$initial"/>  
        </xsl:document>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
</xsl:stylesheet>
