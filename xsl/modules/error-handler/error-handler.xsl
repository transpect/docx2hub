<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:fn="http://www.w3.org/2005/xpath-functions"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:word200x="http://schemas.microsoft.com/office/word/2003/wordml"
    xmlns:v="urn:schemas-microsoft-com:vml" 
    xmlns:dbk="http://docbook.org/ns/docbook"
    xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:exsl='http://exslt.org/common'
    xmlns:saxon="http://saxon.sf.net/"
    xmlns:docx2hub="http://transpect.io/docx2hub"
    xmlns:tr="http://transpect.io"
    exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon tr fn docx2hub"
    version="2.0">

  <xsl:param name="error-mode" select="'debug'"/>
  <xsl:param name="language-localization" select="'de'"/>
  <xsl:param name="error-msg-file" select="'xslt_error-de.yml'"/>
  <xsl:param name="create-pis" select="'yes'"/>
  
  <xsl:variable name="error-info">
    <xsl:choose>
      <xsl:when test="$error-mode = 'debug'">
        <xsl:sequence 
          select="unparsed-text($error-msg-file, 'utf-8')"/>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:variable>

  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->
  <!-- template to handle error messages -->
  <!-- ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ -->

  <xsl:template name="signal-error">
    <xsl:param name="error-code" select="'W2D_999'"/>
    <xsl:param name="fail-on-error"/>
    <xsl:param name="hash" select="()" as="document-node()?"/>
    <xsl:choose>
      <xsl:when test="$error-mode = 'debug'">
        <xsl:variable name="msg" select="normalize-space(tokenize($error-info, '\n')[matches(., concat('^\s*', $error-code, ':'))])"/>       
        <xsl:variable name="msg-text">
          <xsl:analyze-string select="$msg" regex="(\[\[)?\{{\{{(.+?)\}}\}}(\]\])?">
            <xsl:matching-substring>
              <xsl:value-of select="$hash/*:value[@key = regex-group(2)]"/>
            </xsl:matching-substring>
            <xsl:non-matching-substring>
              <xsl:value-of select="."/>
            </xsl:non-matching-substring>
          </xsl:analyze-string>
        </xsl:variable>
        <xsl:message select="concat('##  Error Code: ', $error-code)"/>
        <xsl:message select="concat('##  Error Msg: ', $msg-text)"/>
        <xsl:message select="$hash"></xsl:message>
        <xsl:if test="$hash/*:value[@key = 'mode']">
          <xsl:message select="concat('##  Mode: ', $hash/*:value[@key = 'mode'])"/>
        </xsl:if>
        <xsl:if test="$hash/*:value[@key = 'level']">
          <xsl:message select="concat('##  Level: ', $hash/*:value[@key = 'level'])"/>
        </xsl:if>
        <xsl:if test="$hash/*:value[@key = 'comment']">
          <xsl:comment>
            <xsl:value-of select="$hash/*:value[@key = 'xpath'], $msg-text"></xsl:value-of>
          </xsl:comment>
        </xsl:if>
        <xsl:if test="$hash/*:value[@key = 'xpath']">
          <xsl:message select="concat('##  XPath: ', $hash/*:value[@key = 'xpath'])"/>
        </xsl:if>
        <xsl:if test="$hash/*:value[@key = 'info-text']">
          <xsl:message select="concat('##  Info: ', $hash/*:value[@key = 'info-text'])"/>
        </xsl:if>
        <xsl:if test="$hash/*:value[@key = 'pi'] and $create-pis = 'yes'">
          <xsl:processing-instruction name="tr">
            <xsl:value-of select="$hash/*:value[@key = 'pi']"/>
          </xsl:processing-instruction>
        </xsl:if>
      </xsl:when>
      <xsl:otherwise>
        <xsl:message select="concat('##', $error-code, ':')"/>
        <xsl:for-each select="$hash/*:value[not(@key = 'pi')]">
          <xsl:message select="concat('##  ', ./@key, ': ', .)"/>
        </xsl:for-each>
        <xsl:if test="$hash/value[@key = 'pi'] and $create-pis = 'yes'">
          <xsl:processing-instruction name="tr">
            <xsl:value-of select="$hash/*:value[@key = 'pi']"/>
          </xsl:processing-instruction>
        </xsl:if>
      </xsl:otherwise>
    </xsl:choose>
    <xsl:if test="$fail-on-error eq 'yes'">
      <xsl:message terminate="{$fail-on-error}"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:function name="docx2hub:message" as="processing-instruction()">
    <xsl:param name="context" as="node()"/>
    <xsl:param name="terminate-on-error" as="xs:boolean"/>
    <xsl:param name="terminate-on-warning" as="xs:boolean"/>
    <xsl:param name="code" as="xs:string"/>
    <xsl:param name="severity" as="xs:string"/><!-- ('INFO', 'WRN', 'ERR', 'NRE') -->
    <xsl:param name="mode" as="xs:string"/>
    <xsl:param name="message" as="xs:string"/>
    <xsl:variable name="srcpath" as="attribute(srcpath)?" select="$context/ancestor-or-self::*[@srcpath][1]/@srcpath"/>
    <xsl:message terminate="{('yes'[$terminate-on-error][$severity = ('ERR', 'NRE')],
                              'yes'[$terminate-on-warning][$severity = ('WRN', 'ERR', 'NRE')],
                              'no')[1]}" 
      select="string-join(('&#xa;///////=====================', $severity, $code, $srcpath, $message, '=====================///////'), '&#xa;')"/>
    <xsl:processing-instruction name="tr" select="string-join(($code, $severity, $message), ' ')"/>
  </xsl:function>

</xsl:stylesheet>