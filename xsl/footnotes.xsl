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
  xmlns:tr="http://transpect.io"
  version="2.0"
  xmlns="http://docbook.org/ns/docbook"
  exclude-result-prefixes = "fn xs w word200x v dbk wx o pkg r rel exsl saxon mml css docx2hub tr">

  <!-- We don’t need to include MS Word localized style names if we trust it to
    always use 'footnote reference' as the native name -->
  <xsl:variable name="footnote-reference-styles" as="xs:string+"
                select="('Funotenanker', 
                         'FootnoteAnchor',   (: LibreOffice de/en :)
                         'FootnoteReference', 
                         'Funotenzeichen'    (: MS Word en/de :))"/>
  
  <xsl:function name="docx2hub:is-footnote-reference-style" as="xs:boolean">
    <xsl:param name="style" as="attribute(role)?"/>
    <!-- It is important that even temporary trees contain the complete css:rules --> 
    <xsl:choose>
      <xsl:when test="empty($style)">
        <xsl:sequence select="false()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$style = $footnote-reference-styles
                              or
                              (: MS Word reserved native style name: :)
                              key('docx2hub:style', $style, root($style))/@native-name = 'footnote reference'"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:function name="docx2hub:element-is-footnoteref" as="xs:boolean">
    <xsl:param name="el" as="element()"/>
    <xsl:sequence select="$el/self::*[name() = ('w:r', 'superscript')] and
                          (
                            docx2hub:is-footnote-reference-style($el/@role)
                            or 
                            $el/w:footnoteRef
                          )"/>
  </xsl:function>

  <xsl:template match="w:continuationSeparator" mode="wml-to-dbk">
    <!-- preliminary mapping for 30078588.docx, https://mantis.le-tex.de/mantis/view.php?id=23991
      It creates a horizontal line in the content in Word. We’ll just have an empty para in Hub for now. --> 
  </xsl:template>

  <xsl:template match="w:footnoteReference" mode="wml-to-dbk">
    <footnote>
      <xsl:variable name="id" select="@w:id"/>
      <xsl:attribute name="xml:id" select="string-join(('fn', $id), '-')"/>
      <xsl:variable name="xreflabel" select="if (@w:customMarkFollows=('1','on','true')) 
                                             then following-sibling::w:t[1]/text() 
                                             else ''" as="xs:string?"/>
      <xsl:if test="not($xreflabel = '')">
        <xsl:attribute name="xreflabel" select="$xreflabel"/>
      </xsl:if>
      <xsl:apply-templates select="/*/w:footnotes/w:footnote[@w:id = $id]/@srcpath" mode="#current"/>
      <xsl:apply-templates select="/*/w:footnotes/w:footnote[@w:id = $id]" mode="#current"/>
    </footnote>
  </xsl:template>
  
  <xsl:template match="w:t[preceding-sibling::w:footnoteReference/@w:customMarkFollows=('1', 'on', 'true')]" mode="wml-to-dbk"/>

  <xsl:template match="w:footnote" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:variable name="docx2hub:footnote-marker-embellishment-regex" as="xs:string" select="'^[\p{P}\s\p{Zs}]*$'"/>

  <!-- collateral, has to run before the templates below. They currently match in docx2hub:join-instrText-runs --> 
  <xsl:template match="w:footnote/w:p[1]/w:r[w:tab]" mode="docx2hub:remove-redundant-run-atts">
    <xsl:variable name="r" select="." as="element(w:r)"/>
    <xsl:for-each-group select="* except w:rPr" group-starting-with="*[self::w:tab]">
      <xsl:sequence select="current-group()[self::w:tab]"/>
      <xsl:for-each select="$r">
        <xsl:copy>
          <xsl:sequence select="@*, w:rPr, current-group()[not(self::w:tab)]"/>
        </xsl:copy>
      </xsl:for-each>
    </xsl:for-each-group>
  </xsl:template>


  <!-- The normal case is a run with rStyle='FootnoteReference' (or al localized name) with an element w:footnoteRef,
    followed by a run in the same style with a w:t that consists of a single space.
    The remainder of the paragraph should not have the 'FootnoteReference' rStyle.
    w:footnoteRef will lead to the footnote number being calculated with the corresponding formatting. 
    However, if there is additional 'FootnoteReference' formatting in the paragraph that is directly adjacent to the
    first w:r with that style (modulo runs that consist only of whitespace or punctuation in between), it will also
    become part of the <phrase role="hub:identifier"> that will be put around the footnote marker.
    The first space after the marker (the whole marker according to the previous sentence, not just <w:footnoteRef/>)
    will be put in a <phrase role="hub:separator">.
    If the marker erroneously stetches across the whole paragraph (which is a common error that authors do), the common 
    case that they insert a tab between the actually intended marker and the footnote text (that also has the 
    'FootnoteReference' style) should be handled as follows: 
    The tab should receive an attribute role="hub:separator", its content should be a single &#9; character, and the
    content after the tab must not be put in a 'hub:identifier' phrase.
    It is still an open question how the following common error should be handled: The whole paragraph contains
    'FootnoteReference' formatting, there is no tab, and there is a space after <w:footnoteRef/>. 
    Currently, the whole paragraph will be treated as a hub:identifier. I think it will be more reasonable to split
    at the space that comes after <w:footnoteRef/>, converting it into a separator.
    
    In the first paragraph of a footnote, 
    – there should be phrase in DocBook namespace that has the attribute role="hub:identifier" and it should
      be the first phrase in the para;
    – there should be an element (phrase or tab, both in DocBook namespace) that has an attribute role="hub:separator"
      and that should immediately follow hub:identifier. If it doesn’t follow immediately, there is probably manually 
      added content between the footnote marker and the tab. The hub:separator should not be enclosed by the 
      hub:identifier.
      
    The absence of either feature should evoke a Schematron warning. A missing separator, especially if there is no
    following text content and the hub:identifier does not match an enumeration regex, is an indicator that they 
    should remove the footnote reference character style from the footnote text.
    
    For docx content that is generated by LibreOffice, a tab (with role 'hub:separator') is inserted after the marker.
    In LibreOffice-generated docx, there is no default style name as MS Word’s 'footnote reference'. Therefore for each
    localized version, the variable $footnote-reference-styles needs to be extended.
    -->

  <xsl:template match="w:footnote/w:p[1][*[docx2hub:element-is-footnoteref(.)]]" mode="docx2hub:join-instrText-runs" 
    name="docx2hub:first-note-para" priority="1">
    <!-- This template is also provided as a named variant because it will not match otherwise due to import precedence 
      rules. It has to be called explicitly in join-runs.xsl. We don’t want to introduce an additional XSLT pass
    just for first paras in footnotes. -->
    <xsl:param name="identifier" select="false()" tunnel="yes"/>
    <xsl:variable name="root" select="/" as="document-node(element(dbk:hub))"/>
    <xsl:copy>
      <xsl:call-template name="docx2hub:adjust-lang"/>
      <xsl:apply-templates select="@*, dbk:tabs" mode="#current"/>
      <xsl:for-each-group select="* except dbk:tabs" 
            group-adjacent="(
                              docx2hub:element-is-footnoteref(.) 
                              or 
                              ( matches(., $docx2hub:footnote-marker-embellishment-regex) and 
                                following-sibling::*[1][docx2hub:element-is-footnoteref(.)
                                                        or
                                                        self::w:tab[not(preceding-sibling::w:tab)]]
                              )
                            )
                            and not(preceding-sibling::w:tab)">
        <xsl:choose>
          <xsl:when test="current-grouping-key()">
            <xsl:variable name="tab" as="element(w:tab)*" 
              select="current-group()/descendant-or-self::w:tab[every $n in (current-group()//* except .) satisfies ($n &lt;&lt; .)]"/>
            <phrase role="hub:identifier">
              <xsl:apply-templates select="current-group() except $tab" mode="#current">
                <xsl:with-param name="identifier" select="true()" tunnel="yes"/>
                <xsl:with-param name="tab" select="$tab" tunnel="yes"/>
              </xsl:apply-templates>
            </phrase>
            <xsl:apply-templates select="$tab" mode="docx2hub:join-instrText-runs_footnote-tabs">
              <xsl:with-param name="last" select="$tab[last()]"/>
            </xsl:apply-templates>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current">
              <xsl:with-param name="tab" select="current-group()/descendant-or-self::w:tab" tunnel="yes"/>
            </xsl:apply-templates>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>

  <!-- If there are erroneous identifier phrases in the middle of the paragraph, remove them: -->
  <xsl:template match="w:footnote/w:p/dbk:phrase[@role = 'hub:identifier']
                                                [preceding-sibling::*:phrase[@role = 'hub:identifier']]" 
                mode="wml-to-dbk">
    <superscript>
      <xsl:apply-templates mode="#current"/>  
    </superscript>
  </xsl:template>

  <xsl:template match="w:tab[ancestor::w:footnote | ancestor::w:endnote]" mode="docx2hub:join-instrText-runs">
    <xsl:param name="tab" as="element(w:tab)*" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test=". is $tab[1]">
        <xsl:copy>
          <xsl:apply-templates select="@*" mode="#current"/>
          <xsl:attribute name="role" select="string-join(('hub:separator', @role), ' ')"/>
          <xsl:apply-templates mode="#current"/>
        </xsl:copy>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:tab[ancestor::w:footnote | ancestor::w:endnote]" mode="docx2hub:join-instrText-runs_footnote-tabs">
    <xsl:param name="last" as="element(w:tab)"/>
    <xsl:choose>
      <xsl:when test=". is $last">
        <xsl:copy>
          <xsl:apply-templates select="@*" mode="#current"/>
          <xsl:attribute name="role" select="string-join(('hub:separator', @role), ' ')"/>
          <xsl:apply-templates mode="#current"/>
        </xsl:copy>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:footnote/w:p/w:r//w:tab" mode="docx2hub:join-instrText-runs">
    <xsl:param name="tab" as="element(w:tab)*" tunnel="yes"/>
    <xsl:if test="not(some $t in $tab satisfies ($t is current()))">
      <xsl:next-match/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:footnote/w:p/w:r/@role[docx2hub:is-footnote-reference-style(.)]" 
                mode="docx2hub:join-instrText-runs"/>

  <!-- There is a space after the marker in each Word-generated footnote. Convert it to a separator if there is no
    following separator tab. -->
  <xsl:template match="w:footnote/w:p/w:r[preceding-sibling::w:r[1]/w:footnoteRef]/w:t/text()" mode="docx2hub:join-instrText-runs">
    <xsl:param name="tab" as="element(w:tab)*" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="empty($tab)">
        <xsl:analyze-string select="." regex="^ ">
          <xsl:matching-substring>
            <phrase role="hub:separator">
              <xsl:attribute name="xml:space" select="'preserve'"/>
              <xsl:value-of select="."/>
            </phrase>
          </xsl:matching-substring>
          <xsl:non-matching-substring>
            <xsl:value-of select="."/>
          </xsl:non-matching-substring>
        </xsl:analyze-string>
      </xsl:when>
      <xsl:otherwise>
        <xsl:next-match/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- correct (remove) an embellishment that has been marked as hub:separator when there was also a regular tab separator: -->
  <xsl:template match="dbk:phrase[@role = 'hub:identifier']
                                 [following-sibling::*[1]/self::w:tab[@role = 'hub:separator']]
                         /w:r[w:t[dbk:phrase/@role = 'hub:separator']
                                 [count(node()) = 1]
                             ]" mode="wml-to-dbk"/>
  
  <xsl:template match="w:footnoteRef" mode="docx2hub:join-instrText-runs">
    <xsl:param name="identifier" select="false()" tunnel="yes"/>
    <xsl:if test="$identifier">
      <xsl:variable name="fnref" as="element(w:footnoteReference)*"
        select="key('footnoteReference-by-id', ancestor::w:footnote/@w:id)"/>
      <xsl:variable name="footnote-num-format" 
                    select="( $fnref/following::w:footnotePr[ancestor::w:p | ancestor::w:sectPr]/w:numFmt/@w:val,
                             /*/w:settings/w:footnotePr/w:numFmt/@w:val)[1]" as="xs:string?"/>
      <xsl:variable name="provisional-footnote-number">
        <xsl:number value="if (exists($fnref)) 
                           then count(
                             distinct-values(
                               $fnref[1]
                                  /preceding::w:footnoteReference[not(@w:customMarkFollows = ('1','on','true'))]/@w:id
                             )
                           ) + 1
                           else (count(preceding::w:footnoteRef) + 1)" 
                    format="{if ($footnote-num-format)
                             then tr:get-numbering-format($footnote-num-format, '') 
                             else '1'}"/>
      </xsl:variable>
      <xsl:variable name="cardinality" select="if (matches($provisional-footnote-number,'^\*†‡§[0-9]+\*†‡§$'))
                                               then xs:integer(replace($provisional-footnote-number, '^\*†‡§([0-9]+)\*†‡§$', '$1'))
                                               else 0"/>
      <xsl:variable name="footnote-number">
        <xsl:value-of select="if (matches($provisional-footnote-number,'^\*†‡§[0-9]+\*†‡§$')) 
                              then string-join((for $i 
                                                in (1 to xs:integer(ceiling($cardinality div 4))) 
                                                return substring($provisional-footnote-number,if (($cardinality mod 4) ne 0) 
                                                                                              then ($cardinality mod 4) 
                                                                                              else 4,1)),'') 
                              else if (matches($provisional-footnote-number,'^a[a-z]$')) 
                                   then replace($provisional-footnote-number,'^a([a-z])$','$1$1')
                                   else $provisional-footnote-number"/>
      </xsl:variable>
      <w:t>
        <xsl:choose>
          <!-- This seems to be very specific to a particular workflow / set of conventions / preprocessing: -->
          <xsl:when test="//*:keywordset[@role='docVars']/*:keyword[@role='footnote_check']">
            <xsl:variable name="footnote-check-docvar" as="xs:string*" 
              select="tokenize(
                                //*:keywordset[@role='docVars']/*:keyword[@role='footnote_check'],
                                '&#xD;'
                              )[tokenize(.,',')[1]=$footnote-number]"/>
            <xsl:choose>
              <xsl:when test="exists($footnote-check-docvar)">
                <xsl:variable name="after-comma" select="tokenize($footnote-check-docvar,',')[2]" as="xs:string"/>
                <xsl:value-of select="if (
                                           matches($after-comma,'\)$') 
                                           and 
                                           ancestor::w:footnote//w:r[docx2hub:is-footnote-reference-style(@role)]
                                                                    [matches(.,'^[\s&#160;]*\)[\s&#160;]*$')]
                                         )
                                      then replace($after-comma,'\)$','') 
                                      else $after-comma"/>
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
      </w:t>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="*[*]
                        [
                          self::dbk:superscript or 
                          self::w:r[
                            docx2hub:is-footnote-reference-style(@role)
                            or
                            key('style-by-name', @role, root(.))/@remap eq 'superscript'
                            or
                            (matches(@css:top,'^-') and @css:position eq 'relative')
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