<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  exclude-result-prefixes="xs"
  version="2.0">
  
  <xsl:import href="../sym.xsl"/>
  <xsl:import href="../modules/error-handler/error-handler.xsl"/>
  <xsl:import href="omml2mml.xsl"/>
  
  <xsl:param name="fail-on-error" select="'no'"/>
  
  <xsl:variable name="symbol-font-map" as="document-node(element(symbols))"
    select="if (doc-available('Symbol.xml')) then document('Symbol.xml') else document('../Symbol.xml')"/>
  
  <xsl:template match="m:oMathPara">
    <equation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:apply-templates select="node()" mode="omml2mml"/>
    </equation>
  </xsl:template>
  
  <xsl:template match="m:oMath">
    <inlineequation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:apply-templates select="." mode="omml2mml"/>
    </inlineequation>
  </xsl:template>
  
  <xsl:template match="w:sym" mode="omml2mml" priority="120">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>
  
</xsl:stylesheet>