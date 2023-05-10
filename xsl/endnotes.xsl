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
  xmlns:r= "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:exsl="http://exslt.org/common"
  xmlns:saxon="http://saxon.sf.net/"
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:functx="http://www.functx.com"
  xmlns:tr="http://transpect.io"
  version="2.0"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "fn xs w word200x v dbk wx o pkg r rel exsl saxon mml css docx2hub tr">
  
  <xsl:variable name="endnote-reference-styles" as="xs:string+"
                select="('Endnotenanker', 
                         'EmdnoteAnchor',   (: LibreOffice de/en :)
                         'EndnoteReference', 
                         'Endnotenzeichen'    (: MS Word en/de :))"/>
  
  <xsl:function name="docx2hub:is-endnote-reference-style" as="xs:boolean">
    <xsl:param name="style" as="attribute(role)?"/>
    <!-- It is important that even temporary trees contain the complete css:rules --> 
    <xsl:choose>
      <xsl:when test="empty($style)">
        <xsl:sequence select="false()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$style = $endnote-reference-styles
                              or
                              (: MS Word reserved native style name: :)
                              key('docx2hub:style', $style, root($style))/@native-name = 'endnote reference'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:function name="docx2hub:element-is-endnoteref" as="xs:boolean">
    <xsl:param name="el" as="element()"/>
    <xsl:sequence select="$el/self::*[name() = ('w:r', 'superscript')] and
                          (
                            docx2hub:is-endnote-reference-style($el/@role)
                            or 
                            $el/w:endnoteRef
                          )"/>
  </xsl:function>

  <xsl:template match="w:endnoteReference" mode="wml-to-dbk">
    <footnote role="endnote">
      <xsl:variable name="id" select="@w:id"/>
      <xsl:attribute name="xml:id" select="string-join(('en', $id), '-')"/>
      <xsl:variable name="fn-mark" select="key('endnote-by-id', $id)//dbk:phrase[@role eq 'hub:identifier'][1]" as="xs:string?"/>
      <xsl:variable name="fn-mark-fallback" select="key('endnote-by-id', $id)//w:r[1][following-sibling::node()[1][self::w:tab]]" as="xs:string?"/>
      <xsl:variable name="xreflabel" select="if (@w:customMarkFollows=('1','on','true')) 
                                             then normalize-space(($fn-mark, $fn-mark-fallback)[. ne ''][1])
                                             else ()" as="xs:string?"/>
      <xsl:if test="$xreflabel">
        <xsl:attribute name="xreflabel" select="$xreflabel"/>
      </xsl:if>
      <xsl:apply-templates select="/*/w:endnotes/w:endnote[@w:id = $id]/@srcpath" mode="#current"/>
      <xsl:apply-templates select="/*/w:endnotes/w:endnote[@w:id = $id]" mode="#current"/>
    </footnote>
  </xsl:template>
  
  <xsl:template match="w:t[preceding-sibling::node()[1][self::w:endnoteReference]/@w:customMarkFollows=('1', 'on', 'true')]//text()[1]" mode="wml-to-dbk">
    <xsl:variable name="fnref" as="element(w:endnoteReference)"
                  select="parent::w:t/preceding-sibling::node()[1][self::w:endnoteReference]"/>
    <xsl:variable name="fn" as="element(w:endnote)" select="key('endnote-by-id', $fnref/@w:id)"/>
    <xsl:variable name="fn-mark" select="normalize-space($fn//dbk:phrase[@role eq 'hub:identifier'][1])" as="xs:string"/>
    <xsl:variable name="fn-mark-fallback" select="$fn//w:r[1][following-sibling::node()[1][self::w:tab]]" as="xs:string?"/>
    <xsl:sequence select="if(($fn-mark, $fn-mark-fallback)[. ne ''][1])
                          then replace(., functx:escape-for-regex(($fn-mark, $fn-mark-fallback)[. ne ''][1]), '')
                          else ."/>
  </xsl:template>

  <xsl:template match="w:endnote" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:endnoteRef" mode="docx2hub:join-instrText-runs">
    <xsl:param name="identifier" select="false()" tunnel="yes"/>
    <xsl:param name="footnotePrs" as="element(w:footnotePr)*" tunnel="yes"/>
    <xsl:param name="sect-boundaries" as="element(*)*" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="$identifier">
        <xsl:variable name="fnref" as="element(w:endnoteReference)*"
                      select="key('endnoteReference-by-id', ancestor::w:endnote/@w:id)"/>
        <xsl:variable name="preceding-boundary" as="element(*)?" 
                      select="($sect-boundaries[. &lt;&lt; $fnref/..])[last()]"/>
        <xsl:variable name="following-boundary" as="element(*)?" 
                      select="($sect-boundaries[. >> $fnref/..])[1]"/>
        <xsl:variable name="fnpr" as="element(w:footnotePr)?" 
                      select="($footnotePrs[. >> $fnref/..]
                                           [.. is $following-boundary],
                               if ($docx2hub:use-document-footnotePr-settings) 
                               then /*/w:settings/w:footnotePr else ())[1]"/>
        <xsl:variable name="section-reset" as="xs:boolean" 
                      select="exists($fnpr/w:numRestart[@w:val='eachSect'])"/>
        <xsl:variable name="endnote-num-format" 
                      select="$fnpr/w:numFmt/@w:val" as="xs:string?"/>
        <xsl:variable name="startnum" as="xs:integer" select="xs:integer(($fnpr/w:numStart/@w:val, 1)[1])"/>
        <xsl:variable name="provisional-endnote-number" as="xs:string">
          <xsl:number value="if (exists($fnref)) 
                             then count(distinct-values($fnref[1]/preceding::w:endnoteReference[not(@w:customMarkFollows = 
                                                                                                     ('1','on','true'))]
                                                                                                [if ($section-reset and 
                                                                                                     exists($preceding-boundary)) 
                                                                                                 then . >> $preceding-boundary 
                                                                                                 else true()]/@w:id)) + 
                                  $startnum
                             else (count(preceding::w:endnoteRef[if ($section-reset and exists($preceding-boundary)) 
                                                                  then . >> $preceding-boundary 
                                                                  else true()]) + 
                                   $startnum)" 
                      format="{if ($endnote-num-format)
                               then tr:get-numbering-format($endnote-num-format, '') 
                               else '1'}"/>
        </xsl:variable>
        <xsl:variable name="cardinality" select="if (matches($provisional-endnote-number,'^\*†‡§[0-9]+\*†‡§$'))
                                                 then xs:integer(replace($provisional-endnote-number, '^\*†‡§([0-9]+)\*†‡§$', '$1'))
                                                 else 0"/>
        <xsl:variable name="endnote-number">
          <xsl:value-of select="if (matches($provisional-endnote-number,'^\*†‡§[0-9]+\*†‡§$')) 
                                then string-join((for $i 
                                                  in (1 to xs:integer(ceiling($cardinality div 4))) 
                                                  return substring($provisional-endnote-number,
                                                  if (($cardinality mod 4) ne 0) 
                                                  then ($cardinality mod 4) 
                                                  else 4,1)),'') 
                                else if (matches($provisional-endnote-number,'^a[a-z]$')) 
                                     then replace($provisional-endnote-number,'^a([a-z])$','$1$1')
                                     else $provisional-endnote-number"/>
        </xsl:variable>
        <w:t>
          <xsl:value-of select="$endnote-number"/>
        </w:t>
      </xsl:when>
      <xsl:when test="not(ancestor::*[matches(local-name(),'endnote')])">
        <w:t>
          <xsl:value-of select="'1'"/>
        </w:t>
      </xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>