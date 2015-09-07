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
  xmlns:tr="http://transpect.io"
  xmlns:mml="http://www.w3.org/Math/DTD/mathml2/mathml2.dtd"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0" 
  exclude-result-prefixes = "w o v wx xs dbk pkg r rel word200x exsl saxon fn tr mml">

  <xsl:variable name="footnote-reference-style-regex" select="'^(FootnoteReference|Funotenzeichen)$'"/>

  <xsl:template match="w:footnoteReference" mode="wml-to-dbk">
    <footnote>
      <xsl:variable name="id" select="@w:id"/>
      <xsl:apply-templates select="/*/w:footnotes/w:footnote[@w:id = $id]/@srcpath" mode="#current"/>
      <xsl:apply-templates select="/*/w:footnotes/w:footnote[@w:id = $id]" mode="#current"/>
    </footnote>
  </xsl:template>

  <xsl:template match="w:footnote" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:function name="docx2hub:element-is-footnoteref" as="xs:boolean">
    <xsl:param name="el" as="element()"/>
    <xsl:sequence select="$el/self::*[name() = ('w:r', 'superscript')] and
                          (
                            matches($el/@role, $footnote-reference-style-regex) or 
                            $el/w:footnoteRef
                          )"/>
  </xsl:function>

  <xsl:template match="w:footnote/w:p[*[docx2hub:element-is-footnoteref(.)]]" mode="wml-to-dbk" priority="+1">
    <xsl:param name="identifier" select="false()"/>
    <para>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:for-each-group select="*" 
            group-adjacent="if (
                              (
                                w:footnoteRef and
                                docx2hub:element-is-footnoteref(.) and
                                not(preceding-sibling::*[docx2hub:element-is-footnoteref(.)])
                              )
                              or 
                              (
                                matches(.,'^[\s&#160;]*$') and 
                                following-sibling::*[1][docx2hub:element-is-footnoteref(.)][w:footnoteRef]) and
                                not(preceding-sibling::*[docx2hub:element-is-footnoteref(.)])
                              ) 
                            then true() else false()">
            <xsl:choose>
              <xsl:when test="current-grouping-key()">
                <phrase role="hub:identifier">
                  <xsl:apply-templates select="current-group()/node() except (current-group()/tab, 
                                                                              current-group()[matches(.,'^[\s&#160;]*$')][not(w:footnoteRef)]
                                                                                             [not(exists(following-sibling::w:r[matches(@role,$footnote-reference-style-regex)]
                                                                                                                               [not(matches(.,'^[\s&#160;]*$'))])
                                                                                                  or 
                                                                                                  exists(following-sibling::w:r[w:footnoteRef]))
                                                                                             ]/node())" mode="#current">
                    <xsl:with-param name="identifier" select="true()"/>
                  </xsl:apply-templates>
                </phrase>    
              </xsl:when>
              <xsl:otherwise/>
            </xsl:choose>
          </xsl:for-each-group>
      <xsl:apply-templates select="node() except *[
                                     docx2hub:element-is-footnoteref(.) and 
                                     not(preceding-sibling::*[docx2hub:element-is-footnoteref(.)])
                                   ]" mode="#current"/>
    </para>
  </xsl:template>

  <!--KW 2014-08-14: 
    Nicht footnoteRefs werden gezaehlt, sondern footnoteReference[not(@customMarkFollows='1')]
    Ausserdem sollte zum identifier alles in footnote zaehlen, was die Formatvorlage FootnoteReference hat (ausser whitespace).-->
  <!-- setzt die Nummer der Fußnote. Prüfen!! -->
  <!-- GI 2013-05-23: Apparently both Word 2013 and LibreOffice 4.0.3 generate a number even if the 
      footnote doesn’t contain a footnoteRef. See for example DIN EN 419251-1, Sect. 6.1 -->
  <xsl:template match="w:footnoteRef" mode="wml-to-dbk">
    <xsl:param name="identifier" select="false()"/>
    <xsl:choose>
      <xsl:when test="$identifier">
        <xsl:variable name="footnote-num-format" select="/*/w:settings/w:footnotePr/w:numFmt/@w:val" as="xs:string?"/>
        <xsl:variable name="footnote-number">
          <xsl:number value="(count(preceding::w:footnoteRef) + 1)" 
            format="{
            if ($footnote-num-format)
            then tr:get-numbering-format($footnote-num-format, '') 
            else '1'
            }"/>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check']">
            <xsl:choose>
              <xsl:when test="some $i in (tokenize(//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check'],'&#xD;')) satisfies tokenize($i,',')[1]=$footnote-number">
                <xsl:value-of select="if (matches(tokenize(tokenize(//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check'],'&#xD;')[tokenize(.,',')[1]=$footnote-number],',')[2],'\)$') and ancestor::w:footnote//w:r[matches(@role,$footnote-reference-style-regex)][matches(.,'^[\s&#160;]*\)[\s&#160;]*$')]) then replace(tokenize(tokenize(//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check'],'&#xD;')[tokenize(.,',')[1]=$footnote-number],',')[2],'\)$','') else tokenize(tokenize(//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check'],'&#xD;')[tokenize(.,',')[1]=$footnote-number],',')[2]"/>
              </xsl:when>
              <xsl:otherwise>
                <xsl:value-of select="$footnote-number"/>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$footnote-number"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise/>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:t[ancestor::w:footnote]
                          [following-sibling::w:tab or parent::w:r/following-sibling::w:r[w:tab]]
                          [not(preceding-sibling::w:t[not(matches(.,'^[\s\)]*$'))]) and not(parent::w:r/preceding-sibling::w:r[w:t[not(matches(.,'^[\s\)]*$'))]])][not(matches(parent::w:r/@role,$footnote-reference-style-regex))]
                          [matches(.,'^[\s\)]*$')]"
                mode="wml-to-dbk"/>

  <xsl:template match="*[*]
                        [
                          self::dbk:superscript or 
                          self::w:r[
                            matches(@role,$footnote-reference-style-regex)
                            or
                            key('style-by-name', @role, root(.))/@remap eq 'superscript'
                          ]
                        ]
                        [
                          every $n in node() 
                          satisfies $n/self::w:*[local-name() = ('footnoteRef', 'footnoteReference')]
                        ]" 
                mode="wml-to-dbk" priority="3">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

</xsl:stylesheet>