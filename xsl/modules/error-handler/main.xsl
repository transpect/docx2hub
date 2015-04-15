<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:saxon="http://saxon.sf.net/"
  xmlns:tr="http://transpect.io"
  version="2.0"
  exclude-result-prefixes = "xs saxon fn tr">

  <!-- ================================================================================ -->
  <!-- IMPORT OF OTHER STYLESHEETS -->
  <!-- ================================================================================ -->

  <xsl:import href="error-handler.xsl"/>

  <!-- ================================================================================ -->
  <!-- OUTPUT FORMAT -->
  <!-- ================================================================================ -->
  
  <xsl:output
    method="xml"
    encoding="utf-8"
    indent="yes"
    />

  <xsl:preserve-space elements="*"/>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ main template ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  
  <xsl:template name="main">
    <xsl:call-template name="signal-error">
      <xsl:with-param name="exit" select="'no'"/>
      <xsl:with-param name="hash">
        <value key="info-text">Testtext</value>
        <value key="xpath"><xsl:value-of select="tr:xpath-for-node(.)"/></value>
        <value key="level">WRN</value>
        <value key="mode">test-1</value>
        <value key="pi">test-1</value>
      </xsl:with-param>
    </xsl:call-template>
    <xsl:call-template name="signal-error">
      <xsl:with-param name="exit" select="'yes'"/>
      <xsl:with-param name="error-code" select="'W2D_test'"/>
      <xsl:with-param name="hash">
        <value key="info-text">Testtext 2</value>
        <value key="xpath"><xsl:value-of select="tr:xpath-for-node(node())"/></value>
        <value key="level">ERR</value>
        <value key="mode">test-2</value>
        <value key="pi">test-2</value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>
  
  <!-- tr:xpath-for-node() replaces saxon:path() -->
  <xsl:function name="tr:xpath-for-node" as="xs:string">
    <xsl:param name="current-node" as="node()"/>
    <xsl:variable name="names" as="xs:string*">
      <xsl:for-each select="$current-node/ancestor-or-self::*">
        <xsl:variable name="ancestor-of-current-node" select="."/>
        <xsl:variable name="siblings-with-equal-names" select="$ancestor-of-current-node/../*[name() = name($ancestor-of-current-node)]"/>
        <xsl:sequence select="concat( name( $ancestor-of-current-node ), concat( '[',tr:node-position($siblings-with-equal-names,$ancestor-of-current-node),']' ) )"/>
      </xsl:for-each>
    </xsl:variable>
    <xsl:sequence select="concat( '/', string-join($names,'/') )"/>
  </xsl:function>
  
  <xsl:function name="tr:node-position" as="xs:integer*">
    <xsl:param name="sequence" as="node()*"/> 
    <xsl:param name="node" as="node()"/> 
    <xsl:sequence select=" for $i in (1 to count($sequence)) return $i[$sequence[$i] is $node]"/>
  </xsl:function>
  
</xsl:stylesheet>