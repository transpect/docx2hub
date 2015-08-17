<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
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
  xmlns:exsl="http://exslt.org/common"
  xmlns:saxon="http://saxon.sf.net/"
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:docx2hub ="http://transpect.io/docx2hub"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr mml">

  <!-- merge changemarkup paragraphs:
       an element w:p[w:pPr[w:rPr[w:del]]] has to be merged with the following w:p
  -->
  <xsl:template match="*[w:p]" mode="docx2hub:apply-changemarkup">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="*" 
        group-adjacent="docx2hub:is-merged-changemarkup-para(.) or 
                        preceding-sibling::*[1][docx2hub:is-merged-changemarkup-para(.)]">
        <xsl:choose>
          <xsl:when test="current-grouping-key()">
            <xsl:variable name="merged-para" as="element(w:p)">
              <xsl:element name="w:p">
                <xsl:copy-of select="current-group()[1]/@*"/>
                <xsl:copy-of select="current-group()[1]/node()"/>
                <xsl:copy-of select="current-group()[position() != 1][not(docx2hub:is-changemarkup-removed-para(.))]/node() except w:pPr"/>
              </xsl:element>
            </xsl:variable>
            <xsl:apply-templates select="$merged-para" mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <!-- changemarkup: remove a merged paragraph -->
  <xsl:template match="w:p[docx2hub:is-changemarkup-inserted-para-break(.)]" mode="docx2hub:apply-changemarkup"/>
  
  <xsl:function name="docx2hub:is-merged-changemarkup-para" as="xs:boolean">
    <xsl:param name="el" as="element()"/>
    <xsl:sequence select="not(docx2hub:is-changemarkup-inserted-para-break($el)) and
                          exists(
                            $el/self::w:p[
                              docx2hub:is-changemarkup-removed-para-break(.) or
                              docx2hub:is-changemarkup-removed-para(.)
                            ]
                          )"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:is-changemarkup-inserted-para-break" as="xs:boolean">
    <xsl:param name="el" as="element()"/>
    <xsl:sequence 
      select="exists($el/self::w:p[
                       w:pPr[w:rPr[w:ins]] or
                       w:ins[
                         every $r in descendant::* satisfies 
                         name($r) = ('w:r', 'w:rPr', 'w:bookmarkStart', 'w:bookmarkEnd')
                       ]
                     ][
                       not(w:pPr/w:pPrChange)
                     ][
                       every $e in * satisfies $e[
                         self::w:pPr[w:rPr[w:ins]] or
                         self::w:ins[
                           every $r in descendant::* satisfies 
                           name($r) = ('w:r', 'w:rPr', 'w:bookmarkStart', 'w:bookmarkEnd')
                         ] or
                         self::w:bookmarkStart[@w:name eq '_GoBack'] or
                         self::w:bookmarkEnd
                       ]
                     ])"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:is-changemarkup-removed-para-break" as="xs:boolean">
    <xsl:param name="el" as="element(*)"/>
    <xsl:sequence select="exists($el/self::w:p[w:pPr[w:rPr[w:del]]])"/>
  </xsl:function>

  <xsl:function name="docx2hub:is-changemarkup-removed-para" as="xs:boolean">
    <xsl:param name="para" as="element()"/>
    <xsl:sequence 
      select="exists(
                $para/self::w:p[w:del or w:moveFrom]
                     [every $e in * 
                      satisfies $e[name() = ('w:del', 'w:pPr', 'w:moveFromRangeStart', 'w:moveFromRangeEnd') or 
                      self::w:moveFrom[every $m in * satisfies $m/self::w:del]]
                     ]
              )"/>
  </xsl:function>

  <!-- changemarkup: remove deleted paragraphs -->
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:p[docx2hub:is-changemarkup-removed-para(.)]"/>

  <xsl:template match="w:del" mode="docx2hub:apply-changemarkup"/>

  <xsl:template match="w:ins | w:moveTo | w:moveFrom" mode="docx2hub:apply-changemarkup">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>
  
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:moveFromRangeStart | w:moveFromRangeEnd | w:moveToRangeStart | w:moveToRangeEnd"/>

</xsl:stylesheet>