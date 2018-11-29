<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet version="2.0"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
  xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
  xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
  xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:w10="urn:schemas-microsoft-com:office:word"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
  xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
  xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types"
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
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:m = "http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:docx2hub ="http://transpect.io/docx2hub"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "xs dbk docx2hub tr word200x wx pkg mml">

  <!-- mode docx2hub:changemarkup is for applying user`s tracked changes -->

  <!-- changemarkup: remove deleted table row -->
  <!--<xsl:template match="w:tr[w:trPr[w:del]]" mode="docx2hub:apply-changemarkup"/>-->
 
  <xsl:function name="docx2hub:is-changemarkup-removed-para" as="xs:boolean">
    <xsl:param name="para" as="element()"/>
    <xsl:sequence 
      select="exists(
                      $para/self::w:p[w:del or w:moveFrom or m:oMath[every $i in descendant::text() satisfies $i/ancestor::w:del]]
                     [
                       every $e in * satisfies $e[
                       name() = ('w:del', 'w:pPr', 'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveFrom') or self::m:oMath[every $i in descendant::text() satisfies $i/ancestor::w:del]
                       ]
                     ]
                    )
              or
              exists( (: yes, I created something like this in Word 2016: an ins with nothing but a del inside :)
                      $para/self::w:p[w:ins/w:del]
                                     [every $e in (* except (w:pPr | w:del | w:moveFrom | m:oMath[every $i in descendant::text() satisfies $i/ancestor::w:del])) 
                                      satisfies ($e/self::w:ins
                                                 and
                                                 count($e/(* except (w:del | w:moveFrom | m:oMath[every $i in descendant::text() satisfies $i/ancestor::w:del]))) = 0)]
                    )" />
  </xsl:function>

  <!-- changemarkup: remove deleted paragraphs -->
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:p[docx2hub:is-changemarkup-removed-para(.)]">
    <xsl:if test=". is parent::w:tc/w:p[1]">
      <!-- Word file will be corrupt without w:p in w:tc -->
      <w:p/>
    </xsl:if>
  </xsl:template>

  <!-- All cell contents have been deleted → Word considers row to be deleted. What about whole tables? -->
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:tr[every $p in .//w:tc/w:p satisfies docx2hub:is-changemarkup-removed-para($p)]"/>
  <xsl:template mode="docx2hub:apply-changemarkup"
    match="w:tbl[every $p in .//w:tc/w:p satisfies docx2hub:is-changemarkup-removed-para($p)]"/>

  <xsl:template match="w:del" mode="docx2hub:apply-changemarkup"/>
  <xsl:template match="m:oMath[every $i in descendant::text() satisfies $i/ancestor::w:del]" mode="docx2hub:apply-changemarkup"/>
  <xsl:template match="w:pPrChange" mode="docx2hub:apply-changemarkup"/>

  <!-- some magic: let some deleted end-fldChar elements stay, if begin and end were not equal -->
  <xsl:template match="w:del[*][every $r in * satisfies $r[self::w:r[*][every $e in * satisfies $e[self::w:rPr or self::w:fldChar[@w:fldCharType eq 'end']]]]]" mode="docx2hub:apply-changemarkup" priority="1">
    <xsl:variable name="start-elements" select="preceding-sibling::w:r/w:fldChar[@w:fldCharType eq 'begin']
                                                  union
                                                  preceding-sibling::w:ins/w:r/w:fldChar[@w:fldCharType eq 'begin']
                                                  "/>
    <xsl:variable name="end-elements" select="preceding-sibling::w:r/w:fldChar[@w:fldCharType eq 'end'] 
                                                union
                                                preceding-sibling::w:del/w:r/w:fldChar[@w:fldCharType eq 'end']
                                                union
                                                preceding-sibling::w:ins/w:r/w:fldChar[@w:fldCharType eq 'end']"/>
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