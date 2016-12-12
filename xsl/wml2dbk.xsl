<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:w= "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:dbk="http://docbook.org/ns/docbook"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
  xmlns:tr="http://transpect.io"
  xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:docx2hub="http://transpect.io/docx2hub"
  xmlns:css="http://www.w3.org/1996/css"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0" 
  exclude-result-prefixes = "w xs dbk r rel tr m mc xlink docx2hub v wp">

  <!-- ================================================================================ -->
  <!-- IMPORT OF OTHER STYLESHEETS -->
  <!-- ================================================================================ -->

  <xsl:import href="modules/catch-all/catch-all.xsl"/>
  <xsl:import href="modules/error-handler/error-handler.xsl"/>

  <xsl:param name="debug" select="'yes'" as="xs:string?"/>
  <xsl:param name="fail-on-error" select="'no'" as="xs:string?"/>
  <xsl:param name="field-vars" select="'no'" as="xs:string?"/>
  <xsl:param name="srcpaths" select="'no'" as="xs:string?"/>
  <xsl:param name="base-dir" select="replace(base-uri(), '[^/]+$', '')"/>
  <xsl:param name="extract-dir-uri" select="''" as="xs:string"/><!-- tmp unzip dir URI -->
  <xsl:param name="local-href" select="''" as="xs:string"/><!-- docx file URI -->
  <xsl:variable name="debug-dir" select="concat(replace($base-dir, '^(.+/)(.+?/)$', '$1'), 'debug')"/>
  <!-- Links that probably have been inserted by Word without user consent: -->
  <xsl:param name="unwrap-tooltip-links" select="'no'" as="xs:string?"/>
  <xsl:param name="hub-version" select="'1.0'" as="xs:string"/>
  <xsl:param name="discard-alternate-choices" select="'yes'" as="xs:string"/>
  <xsl:param name="convert-footer" select="false()" as="xs:boolean"/>
  
  <xsl:variable name="docx2hub:discard-alternate-choices" as="xs:boolean"
    select="$discard-alternate-choices = ('yes', 'true', '1')"/>
  
  <xsl:variable name="symbol-font-map" as="document-node(element(symbols))"
                select="document('../fontmaps/Symbol.xml')"/>

  <xsl:key name="style-by-id" match="w:style" use="@w:styleId" />
  <xsl:key name="numbering-by-id" match="w:num" use="@w:numId" />
  <xsl:key name="abstract-numbering-by-id" match="w:abstractNum" use="@w:abstractNumId" />
  <xsl:key name="footnote-by-id" match="w:footnote" use="@w:id" />
  <xsl:key name="endnote-by-id" match="w:endnote" use="@w:id" />
  <xsl:key name="comment-by-id" match="w:comment" use="@w:id" />
  <xsl:key name="doc-rel-by-id" match="w:docRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="footnote-rel-by-id" match="w:footnoteRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="endnote-rel-by-id" match="w:endnoteRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="comment-rel-by-id" match="w:commentRels/rel:Relationships/rel:Relationship" use="@Id" />
  <xsl:key name="symbol-by-number" match="symbol" use="upper-case(replace(@number, '^0*(.+?)$', '$1'))" />
  <xsl:key name="symbol-by-entity" match="symbol" use="@entity" />
  <xsl:key name="style-by-name" match="css:rule | dbk:style" use="@name | @role"/>


  <!-- sorted includes -->
  <xsl:include href="changemarkup.xsl"/>
  <xsl:include href="comments.xsl"/>
  <xsl:include href="endnotes.xsl"/>
  <xsl:include href="footnotes.xsl"/>
  <xsl:include href="index.xsl"/>
  <xsl:include href="numbering.xsl"/>
  <xsl:include href="objects.xsl"/>
  <xsl:include href="images.xsl"/>
  <xsl:include href="omml2mml/omml2mml.xsl"/>
  <xsl:include href="sym.xsl"/>
  <xsl:include href="tables.xsl"/>

  <xsl:function name="tr:node-index-of" as="xs:integer?">
    <xsl:param name="nodes" as="node()*"/>
    <xsl:param name="node" as="node()"/>
    <xsl:sequence select="index-of(for $n in $nodes return generate-id($n), generate-id($node))"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:rel-lookup" as="element(rel:Relationship)">
    <xsl:param name="rid" as="attribute(r:id)"/>
    <xsl:variable name="key-name" as="xs:string"
      select="if ($rid/../ancestor::w:footnote)
              then 'footnote-rel-by-id'
              else 
                if ($rid/../ancestor::w:endnote)
                then 'endnote-rel-by-id'
                else
                  if ($rid/../ancestor::w:comment)
                  then 'comment-rel-by-id'
                  else 'doc-rel-by-id'" />
    <xsl:sequence select="key($key-name, $rid, root($rid))"/>
  </xsl:function>

  <!-- ================================================================================ -->
  <!-- Mode: docx2hub:field-functions -->
  <!-- ================================================================================ -->

  <!-- Each field function will be replaced with an XML element with the same name as the field function -->

  <xsl:template match="/" mode="docx2hub:field-functions">
    <xsl:variable name="field-begins" as="element(w:fldChar)*" 
      select="for $it in .//w:instrText
              return $it/preceding::w:fldChar[1][@w:fldCharType = 'begin']"/>
    <xsl:variable name="field-ends" as="element(w:fldChar)*" 
      select=".//w:fldChar[@w:fldCharType = 'end']
                          [exists(key('docx2hub:item-by-id', @linkend) intersect $field-begins)]"/>
    <xsl:next-match>
      <xsl:with-param name="field-begins" select="$field-begins" tunnel="yes"/>
      <xsl:with-param name="field-ends" select="$field-ends" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:function name="docx2hub:field-function" as="xs:string+">
    <!-- $result[1]: field function name, $result[2]: field function args -->
    <xsl:param name="begin" as="element(w:fldChar)"/>
    <xsl:variable name="prelim" as="xs:string+">
      <xsl:analyze-string select="$begin/following::w:instrText[1]" regex="^\s*\\?(\i\c*)\s+">
        <xsl:matching-substring>
          <xsl:sequence select="regex-group(1)"/>
          <xsl:if test="empty(regex-group(1))">
            <xsl:sequence select="'BROKEN1'"/>
            <xsl:sequence select="."/>
          </xsl:if>
        </xsl:matching-substring>
        <xsl:non-matching-substring>
          <xsl:sequence select="normalize-space(.)"/>
        </xsl:non-matching-substring>
      </xsl:analyze-string>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="count($prelim) = 1 and not(matches($prelim[1], '^\i\c*$'))">
        <xsl:sequence select="'BROKEN2'"/>
        <xsl:sequence select="$prelim"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="$prelim"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:variable name="docx2hub:block-field-functions" as="xs:string+" 
    select="('ADDRESSBLOCK', 'BIBLIOGRAPHY', 'COMMENTS', 'DATABASE', 'INDEX', 'RD', 'TOA', 'TOC')"/>
  
  <xsl:variable name="docx2hub:hybrid-field-functions" as="xs:string+" 
    select="('IF')"/>
  
  <!-- Handle block field functions. The inline field functions will be handled when processing
    the individual current-group()s in docx2hub:field-functions mode, with tunneled begin/end
    field chars.  -->
  <xsl:template match="*[w:p]" mode="docx2hub:field-functions">
    <xsl:param name="field-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="field-ends" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:variable name="block-begins" as="element(w:fldChar)*"
      select="$field-begins[docx2hub:field-function(.)[1] = $docx2hub:block-field-functions]"/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <!-- Assumption 1: Block field functions are not nested. There may be inline field functions
      nested within block field functions, but there are no block field functions within other block field functions.
      Therefore, we may use group-starting-with/group-ending-with to balance them, without recurring to recursion.
      Assumption 2: There is at most only one block field function begin per paragraph.
      Caveat: Since block field functions may be contained in table cells, we should only consider the 
      w:p/w:r/w:fldChars in this context, not arbitrarily deep .//w:fldChars (those in w:tbl/w:tc/w:p). -->
      <xsl:for-each-group select="*" group-starting-with="w:p[.//w:fldChar/generate-id() = $block-begins/generate-id()]">
        <xsl:variable name="begin-fldChar" as="element(w:fldChar)?"
          select=".//w:fldChar[generate-id() = $block-begins/generate-id()]"/>
        <xsl:choose>
          <xsl:when test="$begin-fldChar">
            <xsl:variable name="end-fldChar" as="element(w:fldChar)?" select="docx2hub:corresponding-end-fldChar($begin-fldChar)"/>
            <xsl:variable name="name-and-args" as="xs:string+" select="docx2hub:field-function($begin-fldChar)"/>
            <!-- GI 2010-16-13: It turns out that if the block end field function is at the beginning of a paragraph, then this
    paragraph must be excluded from the block. -->
            <xsl:variable name="end-p" as="element(w:p)?"
              select="for $p in current-group()/self::w:p[.//w:fldChar/generate-id() = $end-fldChar/generate-id()]
                      return if (
                                  empty($end-fldChar/ancestor::w:p[1]//w:t intersect $end-fldChar/parent::w:r/preceding::w:t)
                                  and
                                  exists($p/preceding-sibling::*[1])
                                )
                             then $p/preceding-sibling::*[1]
                             else $p"/>
            <xsl:for-each-group select="current-group()" group-ending-with="w:p[. is $end-p]">
              <xsl:choose>
                <xsl:when test="current-group()[last()] is $end-p">
                  <xsl:element name="{$name-and-args[1]}">
                    <xsl:attribute name="fldArgs" select="$name-and-args[2]"/>
                    <xsl:apply-templates select="current-group()" mode="#current">
                      <xsl:with-param name="field-begins" select="$field-begins except $begin-fldChar" tunnel="yes"/>
                      <xsl:with-param name="field-ends" select="$field-ends except $end-fldChar" tunnel="yes"/>
                    </xsl:apply-templates>
                  </xsl:element>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:apply-templates select="current-group()" mode="#current"/>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:for-each-group>
          </xsl:when>
          <xsl:otherwise>
            <xsl:apply-templates select="current-group()" mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:for-each-group>
    </xsl:copy>
  </xsl:template>
  
  <xsl:function name="docx2hub:corresponding-end-fldChar" as="element(w:fldChar)?">
    <xsl:param name="begin" as="element(w:fldChar)"/>
    <xsl:variable name="prelim" as="element(w:fldChar)*"
      select="key('docx2hub:linking-item-by-id', $begin/@xml:id, root($begin))[@w:fldCharType = 'end']"/>
    <xsl:if test="count($prelim) gt 1">
      <xsl:message terminate="no" select="'More than one docx2hub:corresponding-end-fldChar() result for ', $begin"/>
    </xsl:if>
    <xsl:if test="count($prelim) eq 0">
      <xsl:message terminate="no" select="'No docx2hub:corresponding-end-fldChar() result for ', $begin"/>
    </xsl:if>
    <xsl:sequence select="$prelim[1]"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:corresponding-begin-fldChar" as="element(w:fldChar)">
    <xsl:param name="end" as="element(w:fldChar)"/>
    <xsl:sequence select="key('docx2hub:item-by-id', $end/@linkend, root($end))"/>
  </xsl:function>

  <!-- convert inline field functions to elements of the same names -->
  <xsl:template match="w:p | w:hyperlink" mode="docx2hub:field-functions">
    <xsl:param name="field-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="field-ends" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:variable name="begins-before-para" as="element(w:r)*">
      <xsl:for-each select="$field-begins[. &lt;&lt; current()]
                                         [not(docx2hub:field-function(.)[1] = $docx2hub:block-field-functions)]
                                         [not(docx2hub:corresponding-end-fldChar(.) &lt;&lt; current())]">
        <w:r>
          <xsl:sequence select="., following::w:instrText[1]"/>
        </w:r>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="ends-after-para" as="element(w:r)*"
      select="for $last-in-para in (current()/*[last()]/w:fldChar, current()/*[last()], current())[1]
              return $field-ends[. &gt;&gt; $last-in-para]
                                [not(docx2hub:corresponding-begin-fldChar(.) &gt;&gt; $last-in-para)]
                /.."/>
    <xsl:variable name="props" as="element(*)*" select="w:numPr, w:pPr"/>
    <xsl:copy>
      <xsl:apply-templates select="@*, $props" mode="#current"/>
      <xsl:call-template name="docx2hub:nest-inline-field-function">
        <xsl:with-param name="para-contents" as="document-node()">
          <xsl:document>
            <xsl:sequence select="$begins-before-para, * except $props, $ends-after-para"/>
          </xsl:document>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template name="docx2hub:nest-inline-field-function" as="element(*)*">
    <xsl:param name="para-contents" as="document-node()"/>
    <xsl:variable name="innermost-nesting-begins" as="element(w:fldChar)*" 
      select="for $f in $para-contents/w:r/w:fldChar[@w:fldCharType = 'begin']
                                                    [not(docx2hub:field-function(.)[1] = $docx2hub:block-field-functions)]
              return $f[docx2hub:corresponding-end-fldChar(.) 
                        is 
                        ($para-contents/w:r/w:fldChar[not(@w:fldCharType = 'separate')]
                                                     [. &gt;&gt; $f]
                        )[1]
                       ]"/>
    <xsl:choose>
      <xsl:when test="empty($para-contents/*)"/>
      <xsl:when test="exists($innermost-nesting-begins)">
        <xsl:variable name="innermost-nesting-begin" as="element(w:fldChar)" select="$innermost-nesting-begins[1]"/>
        <xsl:variable name="innermost-nesting-end" as="element(w:fldChar)" 
          select="docx2hub:corresponding-end-fldChar($innermost-nesting-begin)"/>
        <xsl:variable name="name-and-args" as="xs:string+" select="docx2hub:field-function($innermost-nesting-begin)"/>
        <xsl:call-template name="docx2hub:nest-inline-field-function">
          <xsl:with-param name="para-contents">
            <xsl:document>
              <xsl:sequence select="$para-contents/*[. &lt;&lt; $innermost-nesting-begin/..]"/>
              <xsl:element name="{upper-case(replace($name-and-args[1], '\\', ''))}" xmlns="">
                <!-- upper-case: for the rare (and maybe user error) case of 'xe' for index terms -->
                <xsl:attribute name="fldArgs" select="$name-and-args[2]"/>
                <xsl:sequence select="$para-contents/*[. &gt;&gt; $innermost-nesting-begin/..]
                                                      [. &lt;&lt; $innermost-nesting-end/..]"/>
              </xsl:element>
              <xsl:sequence select="$para-contents/*[. &gt;&gt; $innermost-nesting-end/..]"/>
            </xsl:document>
          </xsl:with-param>
        </xsl:call-template>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates select="$para-contents" mode="#current"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:r[every $child in * satisfies $child/(self::w:instrText | self::w:fldChar)]" mode="docx2hub:field-functions"/>

  <!-- ================================================================================ -->
  <!-- Mode: wml-to-dbk -->
  <!-- ================================================================================ -->

  <!-- default for elements -->
  <xsl:template match="*" mode="wml-to-dbk" priority="-1">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_020'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="concat('Element: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- GI 2012-10-08 §§§
       Want to get rid of the warnings. Does that hurt? Not tested.
       KW 2013-03-28: It hurts in case of w:numPr. Indentation properties are missing.
       -->
  <xsl:template match="  w:p/w:numPr 
                       | css:rule/w:numPr 
                       | *:style/w:numPr 
                       | /*/w:numbering 
                       | /*/w:docRels
                       | /*/w:footnoteRels
                       | /*/w:endnoteRels
                       | /*/w:commentRels
                       | /*/w:fonts 
                       | /*/w:comments 
                       | /*/w:footnotes
                       | /*/w:endnotes
                       | mc:AlternateContent
                       | w:fldChar" mode="wml-to-dbk" priority="-0.25"/>    

  <xsl:template match="css:rule/w:tblPr" mode="wml-to-dbk">
    <xsl:apply-templates select="@*" mode="#current"/>
  </xsl:template>

  <xsl:template match="dbk:* | css:*" mode="wml-to-dbk" priority="-0.1">
     <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current" />
      <xsl:apply-templates select="node()" mode="#current" />
    </xsl:copy>
  </xsl:template>

  <xsl:template match="css:rule | *:style" mode="wml-to-dbk">
    <xsl:param name="content" as="element(*)*">
      <!-- linked-style, css:attic -->
    </xsl:param>
    <xsl:copy copy-namespaces="no">
      <xsl:if test="w:numPr">
        <xsl:variable name="ilvl" select="w:numPr/w:ilvl/@w:val"/>
        <xsl:variable name="lvl-properties" select="key('abstract-numbering-by-id',key('numbering-by-id',w:numPr/w:numId/@w:val)/w:abstractNumId/@w:val)/w:lvl[@w:ilvl=$ilvl]"/>
        <xsl:apply-templates select="$lvl-properties/@* except $lvl-properties/@w:ilvl" mode="#current"/>
      </xsl:if>
      <xsl:apply-templates select="@*, w:tblPr, *[not(self::w:tblPr)], $content" mode="#current" />
    </xsl:copy>   
  </xsl:template>
  
  <xsl:template match="css:rule[w:tblPr[@*[contains(local-name(), 'inside')]]]" mode="wml-to-dbk">
    <xsl:copy copy-namespaces="no">
      <xsl:attribute name="layout-type" select="'cell'"/>
      <xsl:attribute name="name" select="docx2hub:linked-cell-style-name(@name)"/>
      <xsl:apply-templates select="w:tblPr/@*[contains(local-name(), 'inside')]" mode="#current">
        <xsl:with-param name="is-first-cell" select="false()" tunnel="yes"/>
        <xsl:with-param name="is-last-cell" select="false()" tunnel="yes"/>
        <xsl:with-param name="is-first-row-in-group" select="false()" tunnel="yes"/>
        <xsl:with-param name="is-last-row-in-group" select="false()" tunnel="yes"/>
      </xsl:apply-templates>
    </xsl:copy>
    <!-- Cell style will be generated before processing the table style. Reason: table border properties
      will have CSS-precedence over cell border properties. We just need to make sure that a table with 
      no outer borders will explicitly override the cell style borders if they are present. --> 
    <xsl:next-match>
      <xsl:with-param name="content" as="element(dbk:linked-style)">
        <linked-style xmlns="http://docbook.org/ns/docbook" layout-type="cell" name="{docx2hub:linked-cell-style-name(@name)}"/>
      </xsl:with-param>
    </xsl:next-match>
  </xsl:template>

  <xsl:template match="@srcpath" mode="wml-to-dbk">
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@*:paraId" mode="wml-to-dbk" priority="3" />
  <xsl:template match="@*:textId" mode="wml-to-dbk" priority="3" />

  <!-- default for attributes -->
  <xsl:template match="@*" mode="wml-to-dbk" priority="1">
    <xsl:copy/>
  </xsl:template>

  <xsl:template match="@*[not(starts-with(name(), 'docx2hub:generated'))]" mode="wml-to-dbk" priority="-0.5">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_021'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="../@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="concat('Attribut: ', name(), '     Parent: ', ../name())"/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- default for comments -->
  <xsl:template match="comment()" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_022'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <!-- standard Kommentar von Aspose nicht warnen -->
  <xsl:template match="comment()[matches(., '^ Generated by Aspose')]" mode="wml-to-dbk"/>

  <!-- default for PIs -->
  <xsl:template match="processing-instruction()" mode="wml-to-dbk">
    <xsl:call-template name="signal-error" xmlns="">
      <xsl:with-param name="error-code" select="'W2D_023'"/>
      <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
      <xsl:with-param name="hash">
        <value key="xpath"><xsl:value-of select="preceding::*[1]/@srcpath"/></value>
        <value key="level">INT</value>
        <value key="mode">wml-to-dbk</value>
        <value key="info-text"><xsl:value-of select="."/></value>
      </xsl:with-param>
    </xsl:call-template>
  </xsl:template>

  <xsl:template match="w:commentRels" mode="wml-to-dbk"/>    

  <!-- element section -->

  <xsl:template match="w:document" mode="wml-to-dbk">
    <xsl:apply-templates select="@* except @srcpath, *" mode="wml-to-dbk"/>
  </xsl:template>

  <xsl:template match="@mc:Ignorable" mode="wml-to-dbk"/>

  <!-- collateral, has to happen before wml-to-hub because redundant xml:lang attributes will
    be eliminated then (redundant = same as this top-level language) -->
  <xsl:template match="/dbk:*" mode="docx2hub:join-instrText-runs">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:variable name="most-frequent-lang" select="docx2hub:most-frequent-lang(.)" as="xs:string?"/>
      <xsl:if test="exists($most-frequent-lang)">
        <xsl:attribute name="xml:lang" select="$most-frequent-lang"/>
      </xsl:if>
      <xsl:apply-templates mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:function name="docx2hub:most-frequent-lang" as="xs:string?">
    <xsl:param name="context" as="element(*)"/>
    <xsl:variable name="langs" as="xs:string*">
      <xsl:for-each-group select="$context//w:t" group-by="docx2hub:text-lang(.)">
        <xsl:sort select="string-length(string-join(current-group(), ''))" order="descending"/>
        <xsl:sequence select="current-grouping-key()"/>
      </xsl:for-each-group>
    </xsl:variable>
    <xsl:sequence select="$langs[1]"/>
  </xsl:function>
  
  <xsl:function name="docx2hub:text-lang" as="xs:string?">
    <xsl:param name="text" as="element(w:t)"/>
    <xsl:variable name="closest" select="$text/ancestor::*[@xml:lang | @role[key('style-by-name', ., $text)/@xml:lang]][1]" as="element(*)?"/>
    <xsl:sequence select="($closest/@xml:lang, key('style-by-name', $closest/@role, root($text))/@xml:lang)[1]"/>
  </xsl:function>


  <xsl:template match="/dbk:*" mode="wml-to-dbk">
    <xsl:copy copy-namespaces="no">
      <xsl:apply-templates select="@*, *" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <!-- paragraphs (w:p) -->

  <xsl:variable name="docx2hub:allowed-para-element-names" as="xs:string+"
    select="('w:r', 'w:pPr', 'w:bookmarkStart', 'w:bookmarkEnd', 'w:smartTag', 'w:commentRangeStart', 'w:commentRangeEnd', 'w:proofErr', 'w:hyperlink', 'w:del', 'w:ins', 'w:fldSimple', 'm:oMathPara', 'm:oMath')" />

  <xsl:template match="w:p" mode="wml-to-dbk">
    <xsl:element name="para">
      <xsl:apply-templates select="@* except @*[matches(name(),'^w:rsid')]" mode="#current"/>
      <xsl:if test="w:r[1][count(*)=1][w:br[@w:type='page']]">
        <xsl:attribute name="css:page-break-before" select="'always'"/>
      </xsl:if>
      <xsl:if test="w:r[last()][count(*)=1][w:br[@w:type='page']] and count(w:r[count(*)=1][w:br[@w:type='page']]) gt 1">
        <xsl:attribute name="css:page-break-after" select="'always'"/>
      </xsl:if>
<!--      <xsl:if test=".//w:r">-->
        <xsl:sequence select="tr:insert-numbering(.)"/>
      <!--</xsl:if>-->
      <!-- Only necessary in tables? They'll get lost otherwise. -->     
      <xsl:variable name="bookmarkstart-before-p" as="element(w:bookmarkStart)*"
        select="preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/preceding-sibling::*[not(self::w:bookmarkStart or self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkstart-before-tc" as="element(w:bookmarkStart)*"
        select="parent::w:tc[current() is w:p[1]]/preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/parent::w:tc/preceding-sibling::*[not(self::w:bookmarkStart or self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkstart-before-tr" as="element(w:bookmarkStart)*"
        select="parent::w:tc/parent::w:tr[current() is (w:tc/w:p)[1]]/preceding-sibling::w:bookmarkStart[. &gt;&gt; current()/parent::w:tc/parent::w:tr/preceding-sibling::*[not(self::w:bookmarkStart or self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkend-after-p" as="element(w:bookmarkEnd)*"
        select="following-sibling::w:bookmarkEnd[. &lt;&lt; current()/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkend-after-tc" as="element(w:bookmarkEnd)*"
        select="parent::w:tc[current() is w:p[1]]/following-sibling::w:bookmarkEnd[. &lt;&lt; current()/parent::w:tc/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>
      <xsl:variable name="bookmarkend-after-tr" as="element(w:bookmarkEnd)*"
        select="parent::w:tc/parent::w:tr[current() is (w:tc/w:p)[1]]/following-sibling::w:bookmarkEnd[. &lt;&lt; current()/parent::w:tc/parent::w:tr/following-sibling::*[not(self::w:bookmarkEnd)][1]]"/>

      <xsl:apply-templates select="$bookmarkstart-before-p | $bookmarkstart-before-tc | $bookmarkstart-before-tr" mode="wml-to-dbk-bookmarkStart"/>
      <xsl:apply-templates select="node() except dbk:tabs" mode="#current"/>
      <xsl:apply-templates select="$bookmarkend-after-p | $bookmarkend-after-tc | $bookmarkend-after-tr" mode="wml-to-dbk-bookmarkEnd"/>
    </xsl:element>
  </xsl:template>

  
  <!-- Verlauf -->
  <xsl:template match="w:del" mode="wml-to-dbk">
    <!-- gelöschten Text wegwerfen -->
  </xsl:template>

  <xsl:template match="w:ins" mode="wml-to-dbk">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>

  <!-- footer (w:ftr) -->
  <xsl:template match="w:ftr[$convert-footer]" mode="wml-to-dbk">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>

  <!-- bookmarks -->

  <xsl:key name="docx2hub:bookmarkStart" match="w:bookmarkStart" use="@w:id"/>

  <!-- mode wml-to-dbk-bookmarkStart is for transforming bookmarkStarts that used to be in between w:ps 
    within the paras where they belong -->
  <xsl:template match="w:bookmarkStart" mode="wml-to-dbk wml-to-dbk-bookmarkStart">
    <anchor role="start">
      <xsl:apply-templates select="@w:name" mode="bookmark-id"/>
    </anchor>
  </xsl:template>
  
  <xsl:key name="docx2hub:bookmarkStart-by-name" match="w:bookmarkStart[@w:name]" 
    use="for $n in docx2hub:normalize-name-for-id(@w:name) return ($n, upper-case($n), lower-case($n))"/>
  
  <xsl:function name="docx2hub:normalize-name-for-id" as="xs:string?">
    <xsl:param name="name" as="xs:string?"/>
    <xsl:choose>
      <xsl:when test="not($name)"/>
      <xsl:when test="not(matches($name, '^\i'))">
        <xsl:sequence select="docx2hub:normalize-name-for-id(concat('_', $name))"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:sequence select="replace($name, '[^-_.a-z\d]', '_', 'i')"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>
  
  <xsl:template match="w:bookmarkStart/@w:name" 
    mode="bookmark-id" as="attribute(xml:id)">
    <xsl:param name="end" select="false()"/>
    <xsl:variable name="normalized-string" as="xs:string" select="docx2hub:normalize-name-for-id(.)"/>
    <xsl:attribute name="xml:id" 
      select="  replace(
                  replace(
                    string-join(
                      (
                        $normalized-string, 
                        if (count(key('docx2hub:bookmarkStart-by-name', $normalized-string)) = 1) then () else ../@w:id, 
                        if ($end) then 'end' else ()
                      ), 
                      '_'
                    ), 
                    '%20', 
                    '_'),
                  '^(\I)',
                  '_$1'
                )"/>
  </xsl:template>

  <!-- remove $bookmarkstart-before-p (see template for w:p above) outside of tables --> 
  <xsl:template match="w:bookmarkStart[following-sibling::w:p]" mode="wml-to-dbk"/>

  <!-- remove $bookmarkend-after-p (see template for w:p above) outside of tables --> 
  <xsl:template match="w:bookmarkEnd[preceding-sibling::w:p]" mode="wml-to-dbk"/>
  
  <xsl:template match="w:bookmarkEnd" mode="wml-to-dbk wml-to-dbk-bookmarkEnd">
    <xsl:if test="exists(key('docx2hub:bookmarkStart', @w:id)[not(@w:name='_GoBack')])">
      <xsl:variable name="start" select="key('docx2hub:bookmarkStart', @w:id)[not(@w:name='_GoBack')]" as="element(w:bookmarkStart)+"/>
      <xsl:if test="count($start) gt 1">
        <xsl:message select="'Multiple bookmarkStart IDs ', @w:id, $start"/>
      </xsl:if>
      <anchor role="end">
        <xsl:variable name="id" as="attribute(xml:id)">
          <xsl:apply-templates select="$start[1]/@w:name" mode="bookmark-id"/> 
        </xsl:variable>
        <xsl:apply-templates select="$start[1]/@w:name" mode="bookmark-id">
          <xsl:with-param name="end" select="true()"/>
        </xsl:apply-templates>
        <xsl:attribute name="linkend" select="$id"/>
      </anchor>
    </xsl:if>
  </xsl:template>

  <xsl:template match="w:bookmarkStart[@w:name eq '_GoBack']" mode="wml-to-dbk wml-to-dbk-bookmarkStart" priority="2"/>
  
  <xsl:template match="w:bookmarkEnd[key('docx2hub:bookmarkStart', @w:id)/@w:name = '_GoBack']" mode="wml-to-dbk wml-to-dbk-bookmarkEnd" priority="2">
    <xsl:if test="not(preceding::w:bookmarkStart[@w:id=current()/@w:id][1]/@w:name = '_GoBack')">
      <xsl:next-match/>
    </xsl:if>
  </xsl:template>
  
  <!-- comments -->
  <xsl:template match="w:commentRangeStart" mode="wml-to-dbk"/>
  
  <xsl:template match="w:commentRangeEnd" mode="wml-to-dbk"/>
  
  <xsl:template match="w:proofErr" mode="wml-to-dbk"/>
 
  <!-- paragraph properties (w:pPr) -->

  <xsl:template match="w:pPr" mode="wml-to-dbk">
     <!-- para properties are collected by para-props.xsl -->
  </xsl:template>

  <!-- run properties (w:rPr) -->

  <xsl:template match="w:rPr" mode="wml-to-dbk">
    <!-- run properties are collected when text nodes are handled -->
  </xsl:template>

  <!-- smartTags (w:smartTag) -->
  <xsl:template match="w:smartTag" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <!-- Hyperlinks -->
  <xsl:template match="w:hyperlink" mode="wml-to-dbk">
    <link>
      <xsl:apply-templates select="@*, *" mode="#current"/>
    </link>
  </xsl:template>

  <xsl:template match="w:hyperlink[@w:tooltip][$unwrap-tooltip-links = ('yes', 'true')]" mode="wml-to-dbk">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <xsl:template match="@w:anchor[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
     <xsl:choose>
      <xsl:when test="exists(parent::w:hyperlink/@r:id)"/>
      <xsl:otherwise>
        <xsl:attribute name="linkend" select="docx2hub:normalize-name-for-id(.)"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="@w:tooltip[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
    <!-- p1609: The method by which this string is surfaced by an application is outside the scope of this Office Open XML Standard. -->
  </xsl:template>

  <xsl:template match="@r:id[parent::w:hyperlink]" mode="wml-to-dbk" priority="1.5">
    <xsl:variable name="value" select="."/>
    <xsl:variable name="rel-item" select="docx2hub:rel-lookup(.)" as="element(rel:Relationship)" />
    <xsl:choose>
      <xsl:when test="exists(parent::w:hyperlink/@w:anchor)">
        <xsl:attribute name="xlink:href" select="concat(
                                                $rel-item/@Target,
                                                '#',
                                                parent::w:hyperlink/@w:anchor
                                              )"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:call-template name="insert-target">
          <xsl:with-param name="rel-item" select="$rel-item"/>
        </xsl:call-template>        
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="insert-target">
    <xsl:param name="rel-item" as="element(rel:Relationship)"/>
    <xsl:choose>
      <xsl:when test="$rel-item/@Type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'">
        <xsl:attribute name="xlink:href" select="$rel-item/@Target" />
      </xsl:when>
      <xsl:otherwise>
        <xsl:attribute name="linkend" select="$rel-item/@Target"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="w:hyperlink/@w:history | w:hyperlink/@w:tgtFrame" mode="wml-to-dbk" priority="1.5"/>
 
  <xsl:template match="w:smartTagPr" mode="wml-to-dbk"/>

  <!-- textbox -->
  <xsl:template match="w:txbxContent" mode="wml-to-dbk">
    <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>


  <!-- runs (w:r) -->
  <xsl:template match="w:r[@* except (@srcpath,@xml:lang[matches(.,'^$')])][not(count(*)=count(w:instrText))]" mode="wml-to-dbk">
    <xsl:element name="phrase">
      <xsl:apply-templates select="@* except @*[matches(name(),'^w:rsid')]" mode="#current"/>
      <xsl:apply-templates select="*" mode="#current"/>
    </xsl:element>
  </xsl:template>

  <xsl:template match="w:r" mode="wml-to-dbk">
    <xsl:apply-templates select="*" mode="#current"/>
  </xsl:template>

  <!-- text (w:t) -->
  <xsl:template match="w:t" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="@xml:lang" mode="wml-to-dbk" priority="5">
    <xsl:variable name="context" select="." as="attribute(*)"/>
    <xsl:variable name="ancestors-with-langs" as="element(*)+">
      <xsl:for-each select="ancestor::*">
        <xsl:copy>
          <xsl:sequence select="key('style-by-name', @role)/(@xml:lang, @css:direction, @docx2hub:rtl-lang)"/>
          <xsl:sequence select="@xml:lang except $context, ../@css:direction, ../@docx2hub:rtl-lang"/>
        </xsl:copy>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="dir" as="xs:string?" select="($ancestors-with-langs, ..)[@css:direction][last()]/@css:direction"/>
    <xsl:variable name="last-lang" select="($ancestors-with-langs[@docx2hub:rtl-lang][@css:direction = 'rtl'][last()]/@docx2hub:rtl-lang,
                                           $ancestors-with-langs[@xml:lang][not(@css:direction = 'rtl')][last()]/@xml:lang)[last()]"/>
    <!-- Only output the next specific xml:lang if its string value differs from the current one’s: -->
    <xsl:if test="not($dir = 'rtl') and not($last-lang = $context)">
      <xsl:attribute name="xml:lang" select="$context"/>
    </xsl:if>
  </xsl:template>

  <xsl:template match="@docx2hub:rtl-lang" mode="wml-to-dbk" priority="5">
    <xsl:variable name="context" select="." as="attribute(*)"/>
    <xsl:variable name="ancestors-with-langs" as="element(*)+">
      <xsl:for-each select="ancestor::*">
        <xsl:copy>
          <xsl:sequence select="key('style-by-name', @role)/(@xml:lang, @css:direction, @docx2hub:rtl-lang)"/>
          <xsl:sequence select="@docx2hub:rtl-lang except $context, ../@css:direction, ../@xml:lang"/>
        </xsl:copy>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="dir" as="xs:string?" select="($ancestors-with-langs, ..)[@css:direction][last()]/@css:direction"/>
    <xsl:variable name="last-lang" select="($ancestors-with-langs[@docx2hub:rtl-lang][@css:direction = 'rtl'][last()]/@docx2hub:rtl-lang,
                                           $ancestors-with-langs[@xml:lang][not(@css:direction = 'rtl')][last()]/@xml:lang)[last()]"/>
    <!-- Only output the next specific xml:lang if its string value differs from the current one’s: -->
    <xsl:if test="$dir = 'rtl' and not($last-lang = $context)">
      <xsl:attribute name="xml:lang" select="$context"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:dir" mode="wml-to-dbk">
    <xsl:if test="@w:val eq 'rtl'">
      <xsl:message select="'WRN: unimplemented rtl direction element with nested children'"></xsl:message>
    </xsl:if>
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <!-- This occured in a file without styles. It had w:position/@w:val="0" all over the place, which
  is particularly bad for docx2tex, where each span will be converted to a \raisebox.
  The first template removes it if it is in a phrase within a para that has the same property. --> 
  <xsl:template match="@css:top[not(ancestor::css:rule)]
                               [not(local-name(..) = ('para', 'simpara', 'title'))]
                               [some $para in ancestor::*[local-name() = ('para', 'simpara', 'title')][1] 
                                satisfies ((key('style-by-name', $para/@role)/@css:top, $para/@css:top)[last()] = current())]" 
    mode="wml-to-dbk" priority="1.5"/>

  <!-- I think removing this is justified even for this superscript in DIN_EN_12602_tr_25461149 that previously read:
    <superscript role="TableFootNoteXref" css:top="0pt" css:position="relative" css:font-size="9.5pt" xml:lang="de">a</superscript>
    It can be seen in Word that it is just plain superscript without any additional shift. -->
  <xsl:template match="@css:top[not(ancestor::css:rule)][. = '0pt']" mode="wml-to-dbk" priority="1"/>
  
  <xsl:template match="@css:position[. = 'relative']" mode="wml-to-dbk" priority="1">
    <!-- only keep this if the corresponding offset is also kept -->
    <xsl:variable name="top" as="attribute(css:top)?">
      <xsl:apply-templates select="../@css:top" mode="#current"/>
    </xsl:variable>
    <xsl:if test="exists($top)">
      <xsl:next-match/>
    </xsl:if>
  </xsl:template>

  <!-- Field functions -->

  <xsl:template match="REF[@fldArgs]" mode="wml-to-dbk" priority="2">
    <xsl:variable name="linkend" select="replace(@fldArgs, '^([_A-Za-z\d]+)(.*)$', '$1')" as="xs:string"/>
    <xsl:choose>
    <!-- Unwrap, don’t link, a REF like this: 
    <REF xmlns="" fldArgs="DINDOKYEAR \* CHARFORMAT \* MERGEFORMAT">
      <w:r css:font-weight="normal" css:color="#000000" css:font-size="10pt">
        <w:t>2015</w:t>
      </w:r>
    </REF>
    if it points to within a SET field:
    <SET xmlns="" fldArgs="DINDOKYEAR &#34;2015&#34;">
      <w:bookmarkStart w:id="12" w:name="DINDOKYEAR"/>
      <w:r css:color="#000000">
        <w:t>2015</w:t>
      </w:r>
      <w:bookmarkEnd w:id="12"/>
    </SET> -->
      <xsl:when test="key('docx2hub:bookmarkStart-by-name', ($linkend, upper-case($linkend)))/ancestor::SET">
        <xsl:apply-templates mode="#current"/>
      </xsl:when>
      <xsl:otherwise>
        <link linkend="{$linkend}">
          <xsl:apply-templates mode="#current"/>
        </link>    
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:variable name="mail-regex" as="xs:string" select="'^[-a-zA-Z0-9.!#$~_]+@[-a-zA-Z0-9]+\.[a-zA-Z0-9]+$'"/>
  
  <xsl:template match="*[@fldArgs]" mode="wml-to-dbk tables">
    <xsl:variable name="tokens" as="xs:string*">
      <xsl:analyze-string select="(@fldArgs, ' ')[ . ne ''][1]" regex="&quot;(.*?)&quot;">
        <xsl:matching-substring>
          <xsl:sequence select="regex-group(1)"/>
        </xsl:matching-substring>
        <xsl:non-matching-substring>
          <xsl:sequence select="tokenize(., '\s+')[normalize-space(.)]"/>
        </xsl:non-matching-substring>
      </xsl:analyze-string>
    </xsl:variable>
    <xsl:variable name="func" select="$tr:field-functions/tr:field-functions/tr:field-function[@name = current()/name()]" as="element(tr:field-function)?"/>
    <xsl:choose>
      <xsl:when test="not($func)">
        <xsl:choose>
          <!-- Should rewrite this case switch to matching templates --> 
          <xsl:when test="name() = 'SYMBOL'">
            <!-- Template in sym.xsl -->
            <xsl:call-template name="create-symbol">
              <xsl:with-param name="tokens" select="$tokens"/>
              <xsl:with-param name="context" select="."/>
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="name() = ('EQ','eq','FORMCHECKBOX')">
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
          <xsl:when test="name() = 'INCLUDEPICTURE'">
            <xsl:choose>
              <!-- figures are preferably handled by looking at the relationships 
              because INCLUDEPICTURE is more like a history of all locations where
              the image was once included from.  
              Because there may be multiple INCLUDEPICTUREs, we ignore them not only
              if the w:pict is contained in a field function, but if there is any 
              w:pict in the current paragraph. Is this assumption justified?
              -->
              <xsl:when test="ancestor::w:p//w:pict[.//v:imagedata]">
                <xsl:apply-templates mode="#current"/>    
              </xsl:when>
              <xsl:otherwise>
                <mediaobject>
                  <xsl:attribute name="docx2hub:field-function" select="'yes'"/>
                  <xsl:apply-templates select="(.//@srcpath)[1]" mode="#current"/>
                  <imageobject>
                    <imagedata fileref="{if (tokenize(@fldArgs, ' ')[matches(.,'^&#x22;.*&#x22;$')]) 
                                         then replace(tokenize(@fldArgs, ' ')[matches(.,'^&#x22;.*&#x22;$')][1],'&#x22;','') 
                                         else if (matches(@fldArgs,'&#x22;.*&#x22;')) 
                                              then tokenize(@fldArgs,'&#x22;')[2] 
                                              else tokenize(@fldArgs, ' ')[2]}"/>
                  </imageobject>
                </mediaobject>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="name() = 'HYPERLINK'">
            <xsl:variable name="link-content" as="node()*">
              <xsl:apply-templates select="*" mode="#current"/>
            </xsl:variable>
            <xsl:variable name="without-options" select="$tokens[not(matches(., '\\[lo]'))]" as="xs:string*"/>
            <xsl:variable name="local" as="xs:boolean" select="$tokens = '\l'"/>
            <xsl:variable name="target" select="if(exists($without-options)) then replace($without-options[1], '(^&quot;|&quot;$)', '') 
                                                else string-join($link-content/descendant-or-self::text(), '')"/>
            <xsl:variable name="tooltip" select="replace($without-options[2], '(^&quot;|&quot;$)', '')"/>
            <link docx2hub:field-function="yes">
              <xsl:attribute name="{if ($local) then 'linkend' else 'xlink:href'}" 
                select="if(matches($target, $mail-regex)) then concat('mailto:', $target) else $target"/>
              <xsl:if test="$tooltip">
                <xsl:attribute name="xlink:title" select="$tooltip"/>
              </xsl:if>
              <xsl:sequence select="(.//@srcpath)[1], $link-content"/>
            </link>
          </xsl:when>
          <xsl:when test="name() = 'SET'">
            <xsl:if test="$field-vars='yes'">
              <keyword role="{concat('fieldVar_',$tokens[1])}" docx2hub:field-function="yes">
                <xsl:value-of select="$tokens[2]"/>    
              </keyword>
              </xsl:if>
          </xsl:when>
          <xsl:when test="matches(@fldArgs,'^[\s&#160;]*$')">
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
          <xsl:when test="name() = 'PRINT'">
            <xsl:processing-instruction name="PRINT" select="string-join($tokens, ' ')"/>
          </xsl:when>
          <xsl:when test="name() = 'AUTOTEXT'">
            <xsl:call-template name="signal-error" xmlns="">
              <xsl:with-param name="error-code" select="'W2D_045'"/>
              <xsl:with-param name="fail-on-error" select="$fail-on-error"/>
              <xsl:with-param name="hash">
                <value key="xpath"><xsl:value-of select="@srcpath"/></value>
                <value key="level">WRN</value>
                <value key="info-text"><xsl:value-of select="@fldArgs"/></value>
              </xsl:with-param>
            </xsl:call-template>
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_040', 'WRN', 'wml-to-dbk', 
                                                   concat('Unrecognized field function in ''', name(), ' ', @fldArgs, ''''))"/>
            <xsl:apply-templates mode="#current"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$func/@element">
            <xsl:element name="{$func/@element}">
              <xsl:attribute name="docx2hub:field-function" select="'yes'"/>
              <xsl:choose>
                <xsl:when test="exists($func/@role) and not($func/@attrib)">
                  <xsl:attribute name="role" select="$func/@role"/>
                  <xsl:apply-templates mode="#current"/>
                </xsl:when>
                <xsl:when test="$func/@attrib">
                  <xsl:attribute name="{$func/@attrib}" select="replace($tokens[position() = $func/@value], '&quot;', '')"/>
                  <xsl:if test="$func/@role">
                    <xsl:attribute name="role" select="$func/@role"/>
                  </xsl:if>
                  <xsl:apply-templates mode="#current"/>
                </xsl:when>
                <xsl:otherwise>
                  <!-- no apply-templates? -->
                </xsl:otherwise>
              </xsl:choose>
            </xsl:element>
          </xsl:when>
          <xsl:when test="$func/@destroy = 'yes'">
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
          <xsl:when test="$func[@name][count(@*) = 1]">
            <!-- Same handling as for @destroy = 'yes'? At least necessary for SEQ.
              Probably a consequence of the 2016-08 change in handling field functions -->
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:variable name="tr:field-functions" as="document-node(element(tr:field-functions))">
    <xsl:document>
      <tr:field-functions>
        <tr:field-function name="BIBLIOGRAPHY" element="div" role="hub:bibliography"/>
        <tr:field-function name="INDEX" element="div" role="hub:index"/>
        <tr:field-function name="NOTEREF" element="link" attrib="linkend" value="1"/>
        <tr:field-function name="PAGE"/>
        <tr:field-function name="PAGEREF" element="link" attrib="linkend" role="page" value="1"/>
        <tr:field-function name="RD"/>
        <tr:field-function name="REF"/>
        <tr:field-function name="ADVANCE"/>
        <tr:field-function name="QUOTE"/>
        <tr:field-function name="SEQ"/>
        <tr:field-function name="STYLEREF"/>
        <tr:field-function name="USERPROPERTY" destroy="yes"/>
        <tr:field-function name="TOA" element="div" role="hub:toa"/>
        <tr:field-function name="TOC" element="div" role="hub:toc"/>
        <tr:field-function name="\IF"/>
      </tr:field-functions>
    </xsl:document>  
  </xsl:variable>
  

  <!-- w:sectPr ignorieren -->
  <xsl:template match="w:sectPr | w:pgSz | w:footnotePr" mode="wml-to-dbk"/>
 
  <xsl:template match="w:tcPr" mode="wml-to-dbk"/>

  <!-- Background -->
  <xsl:template match="w:background[parent::w:document]" mode="wml-to-dbk"/>

  <!-- fldSimple -->
  <xsl:template match="w:fldSimple" mode="wml-to-dbk">
    <!-- §17.16.19 p1592 gehört zu Feldfunktionen. Wenn w:t darunter, muss der geschrieben werden -->
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="w:fldSimple[
                         matches(@w:instr, '^\s*REF\s(_[A-Za-z]+\d+)\s?.+\\h.+$')
                       ]" mode="wml-to-dbk" priority="1">
    <xsl:variable name="linkend" select="replace(@w:instr, '^\s*REF\s(_[A-Za-z]+\d+)\s?.+$', '$1', 'i')" as="xs:string"/>
    <link linkend="{$linkend}">
      <xsl:apply-templates mode="#current"/>
    </link>
  </xsl:template>

  <!-- whitespace elements, etc. -->
  <xsl:template match="w:tab" mode="wml-to-dbk">
    <tab>
      <xsl:attribute name="xml:space" select="'preserve'"/>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:text>&#9;</xsl:text>
    </tab>
  </xsl:template>

  <xsl:template match="w:br" mode="wml-to-dbk">
    <br>
      <xsl:apply-templates select="@srcpath, @w:type" mode="#current"/>
    </br>
  </xsl:template>
  
  <xsl:template mode="wml-to-dbk" match="w:br/@w:type" priority="10">
    <xsl:attribute name="role" select="."/>
  </xsl:template>

  <xsl:template match="w:cr" mode="wml-to-dbk">
    <!-- carriage return -->
    <!-- ggf. als echte Absatzmarke behandeln. Dazu muss ein nachgelagerter neuer mode eingefuehrt werden. -->
    <phrase role="cr"/>
  </xsl:template>

  <xsl:template match="w:softHyphen" mode="wml-to-dbk">
  	<xsl:value-of select="'&#xad;'"/>
  </xsl:template>
 
  <xsl:template match="w:noBreakHyphen" mode="wml-to-dbk">
      <xsl:value-of select="'&#x2011;'"/>
  </xsl:template>

  <xsl:template match="w:pageBreakBefore" mode="wml-to-dbk">
    <xsl:if test="@w:val = ('true', '1', 'on') or not(@w:val)">
      <phrase role="pageBreakBefore"/>
    </xsl:if>
  </xsl:template>
  
  <xsl:template match="w:ruby" mode="wml-to-dbk">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:rt" mode="wml-to-dbk">
    <phrase role="ruby-text">
      <xsl:apply-templates select="parent::w:ruby/@* except parent::w:ruby/@docx2hub:*" mode="#current"/>
      <xsl:apply-templates mode="#current"/>
    </phrase>
  </xsl:template>
  
  <xsl:template match="w:rubyBase" mode="wml-to-dbk">
    <phrase role="ruby-base">
      <xsl:apply-templates mode="#current"/>
    </phrase>
  </xsl:template>

  <!-- math section -->
  <xsl:template match="m:oMathPara" mode="wml-to-dbk">
    <equation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:apply-templates select="node()" mode="omml2mml"/>
    </equation>
  </xsl:template>

  <xsl:template match="m:oMath" mode="wml-to-dbk">
    <inlineequation>
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:apply-templates select="." mode="omml2mml"/>
    </inlineequation>
  </xsl:template>

 <xsl:template match="w:sym" mode="omml2mml" priority="120">
    <xsl:apply-templates select="." mode="wml-to-dbk"/>
  </xsl:template>

  <!-- attribute section -->

  <xsl:template match="@*[(parent::w:p or parent::w:r) and matches(name(), '^w:rsid')]" mode="wml-to-dbk">
    <!-- IDs zur Kennzeichnung des Verlaufs im Word-Dokument ignorieren -->
  </xsl:template>

  <xsl:template match="@w:val[parent::*/name() = (
                      'w:pStyle'
                      )]" mode="wml-to-dbk">
    <!-- Attributswerte, die in anderem Kontext bereits ausgegeben werden -->
  </xsl:template>

  <xsl:template match="@*[parent::w:smartTag]" mode="wml-to-dbk">
    <!-- Attribute von smartTag vorerst ignoriert. -->
  </xsl:template>
  
  <xsl:template match="w:sdt" mode="wml-to-dbk">
    <xsl:apply-templates select="w:sdtContent/*" mode="#current"/>
  </xsl:template>
  
  <!-- The following template removes indentation if the document.xml was processed 
    earlier with libxml without indent flag.  If you miss breaks, it's dead certain 
    that this template is the cause. -->
  <xsl:template match="text()[not(parent::w:t)][matches(., '^&#xa;$')]" mode="wml-to-dbk" priority="2">
    <xsl:next-match/>
  </xsl:template>

  <xsl:template match="@role" mode="wml-to-dbk" priority="2">
    <xsl:attribute name="role" select="replace(., ' ', '_')" />
  </xsl:template>

  <xsl:template match="@srcpath[$srcpaths != 'yes']" mode="wml-to-dbk" priority="2" />

  <xsl:function name="docx2hub:twips2mm" as="xs:string">
    <xsl:param name="val" as="xs:integer"/>
    <xsl:sequence select="concat(xs:string($val * 0.01763889), 'mm')" />
  </xsl:function>

</xsl:stylesheet>