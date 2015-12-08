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
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:docx2hub ="http://transpect.io/docx2hub"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr mml">

  <!-- mode docx2hub:changemarkup is for applying user`s tracked changes -->

  <!-- merge changemarkup paragraphs:
       an element w:p[w:pPr[w:rPr[w:del]]] has to be merged with the following w:p
  -->
  <xsl:template match="*[w:p]" mode="docx2hub:apply-changemarkup">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="*" group-starting-with="*[docx2hub:is-changemarkup-inserted-para-break(.)]">
        <xsl:for-each-group select="current-group()[not(name() = ('w:moveFromRangeEnd', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd'))]" 
        group-adjacent="docx2hub:is-merged-changemarkup-para(.) or 
                        preceding-sibling::*[1][docx2hub:is-merged-changemarkup-para(.)]">
        <xsl:choose>
            <!-- deleted para without following mergable paragraphs, i.e. single para in footnote -->
            <xsl:when test="current-grouping-key() and (
                              every $el in current-group() satisfies docx2hub:is-changemarkup-removed-para($el)
                            )"/>
            <!-- default behaviour: more than one paragraph, merge them -->
          <xsl:when test="current-grouping-key()">
              <xsl:variable name="first-non-merged-blockelement" as="element()?"
                select="current-group()[not(docx2hub:is-merged-changemarkup-para(.))][1]"/>
            <xsl:variable name="merged-para" as="element(*)">
              <!-- element name (15-09-14):
                     in case of a para merged with a table (para deleted, table alive) we need w:tbl as element name 
                     otherwise w:p will be set, usually -->
                <xsl:element name="{($first-non-merged-blockelement[1]/name(), 'w:p')[1]}">
                  <xsl:if test="current-group()/@srcpath">
                    <xsl:attribute name="srcpath" select="current-group()/@srcpath" separator="&#x20;"/>
                  </xsl:if>
                  <xsl:apply-templates select="current-group()[not(docx2hub:is-merged-changemarkup-para(.))][1]/@*" mode="#current"/>
                  <xsl:apply-templates mode="#current"
                    select="$first-non-merged-blockelement/w:pPr,
                            current-group()[not(docx2hub:is-changemarkup-removed-para(.))]/node()[not(self::w:pPr)]"/>
              </xsl:element>
            </xsl:variable>
              <xsl:sequence select="$merged-para"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
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
                       every $e in * satisfies $e[
                         self::w:pPr[w:rPr[w:ins]] or
                         self::w:pPrChange or
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
    <xsl:sequence select="exists($el/self::w:p[not(w:moveTo)][w:pPr[w:rPr[w:del]]])"/>
  </xsl:function>

  <xsl:function name="docx2hub:is-changemarkup-removed-para" as="xs:boolean">
    <xsl:param name="para" as="element()"/>
    <xsl:sequence 
      select="exists(
                $para/self::w:p[w:del or w:moveFrom]
                     [
                       every $e in * satisfies $e[
                         name() = ('m:oMath', 'w:del', 'w:pPr', 'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveFrom')
                       ]
                     ]
              )"/>
  </xsl:function>

  <!-- changemarkup: remove deleted paragraphs -->
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:p[docx2hub:is-changemarkup-removed-para(.)]"/>

  <xsl:template match="w:del" mode="docx2hub:apply-changemarkup"/>

  <xsl:template match="w:pPrChange" mode="docx2hub:apply-changemarkup"/>

  <!-- some magic: let some deleted end-fldChar elements stay, if begin and end were not equal -->
  <xsl:template match="w:del[*][every $r in * satisfies $r[self::w:r[*][every $e in * satisfies $e[self::w:rPr or self::w:fldChar[@w:fldCharType eq 'end']]]]]" mode="docx2hub:apply-changemarkup" priority="1">
    <xsl:variable name="start-elements" select="preceding-sibling::w:r/w:fldChar[@w:fldCharType eq 'begin']"/>
    <xsl:variable name="end-elements" select="preceding-sibling::w:r/w:fldChar[@w:fldCharType eq 'end'] 
                                                union
                                                preceding-sibling::w:del/w:r/w:fldChar[@w:fldCharType eq 'end']"/>
    <xsl:if test="count($start-elements) &gt; count($end-elements)">
      <xsl:apply-templates select="w:r" mode="#current"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:ins | w:moveTo | w:moveFrom" mode="docx2hub:apply-changemarkup">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>
  
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:moveFromRangeStart | w:moveFromRangeEnd | w:moveToRangeStart | w:moveToRangeEnd"/>

</xsl:stylesheet>