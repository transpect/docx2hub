<?xml version="1.0" encoding="UTF-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:xs="http://www.w3.org/2001/XMLSchema"
  xmlns:fn="http://www.w3.org/2005/xpath-functions"
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
  xmlns:mml="http://www.w3.org/1998/Math/MathML"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns="http://docbook.org/ns/docbook"
  version="2.0" 
  exclude-result-prefixes = "w xs dbk fn r rel tr m mc docx2hub v wp">

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
  <!-- Links that probably have been inserted by Word without user content: -->
  <xsl:param name="unwrap-tooltip-links" select="'no'" as="xs:string?"/>
  <xsl:param name="hub-version" select="'1.0'" as="xs:string"/>
  <xsl:param name="discard-alternate-choices" select="'yes'" as="xs:string"/>
  <xsl:param name="include-header-and-footer" select="'no'" as="xs:string"/>
  <xsl:param name="convert-footer" select="false()" as="xs:boolean"/>
  <xsl:param name="mathtype2mml" select="'no'" as="xs:string?"/>
  <xsl:param name="charmap-policy" select="'unicode'" as="xs:string">
    <!-- Values: unicode or xs:string. For xs:string, mapping attribute in the fashion @char-{xs:string} must exist in the symbols file -->
  </xsl:param>
  
  <xsl:variable name="docx2hub:discard-alternate-choices" as="xs:boolean"
    select="$discard-alternate-choices = ('yes', 'true', '1')"/>
  
  <xsl:variable name="symbol-font-map" as="document-node(element(symbols))"
                select="document('http://transpect.io/fontmaps/Symbol.xml')"/>

  <xsl:key name="style-by-id" match="w:style" use="@w:styleId" />
  <xsl:key name="numbering-by-id" match="w:num" use="@w:numId" />
  <xsl:key name="abstract-numbering-by-id" match="w:abstractNum" use="@w:abstractNumId" />
  <xsl:key name="footnote-by-id" match="w:footnote" use="@w:id" />
  <xsl:key name="footnoteReference-by-id" match="w:footnoteReference" use="@w:id" />
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

  <xsl:template match="/dbk:hub" mode="docx2hub:field-functions" priority="2">
    <xsl:variable name="field-begins" as="element(w:fldChar)*" 
      select="key('docx2hub:item-by-id', .//w:instrText/@docx2hub:fldChar-start-id)"/>
    <xsl:variable name="field-begin-ids" select="$field-begins/@xml:id" as="xs:string*"/>
    <xsl:variable name="field-ends" as="element(w:fldChar)*" 
      select=".//w:fldChar[@w:fldCharType = 'end']
                          [@linkend = $field-begin-ids]"/>
    <xsl:variable name="non-block-field-begins" as="element(w:fldChar)*" 
      select="$field-begins[not(
                              key('docx2hub:instrText-by-start-id', @xml:id)/@docx2hub:field-function-name 
                              = $docx2hub:block-field-functions
                           )]"/>
    <xsl:variable name="non-block-field-begin-ids" select="$non-block-field-begins/@xml:id" as="xs:string*"/>
    <xsl:variable name="non-block-field-ends" as="element(w:fldChar)*" 
      select="$field-ends[@linkend = $non-block-field-begin-ids]"/>
    <xsl:variable name="block-begins" as="element(w:fldChar)*"
      select="$field-begins[key('docx2hub:instrText-by-start-id', @xml:id)/@docx2hub:field-function-name = $docx2hub:block-field-functions]"/>
    <xsl:variable name="lookaround-count" as="xs:double*">
      <xsl:for-each select="$non-block-field-begins">
        <xsl:variable name="c" as="element(w:fldChar)?"
          select="key('docx2hub:linking-item-by-id', @xml:id)[@w:fldCharType = 'end']"/>
        <xsl:variable name="distance" as="xs:double" select="xs:double(count(ancestor::w:p/following-sibling::*[$c >> .]))"/>
        <xsl:if test="$distance gt 10">
          <xsl:message select="'Large lookaround distance', $distance, 'due to', ancestor::w:p"/>
        </xsl:if>
        <xsl:sequence select="$distance"/>
      </xsl:for-each>
    </xsl:variable>
    
    <xsl:variable name="max-lookaround-count" as="xs:integer" select="xs:integer((max($lookaround-count), 0)[1])"/>
    <xsl:if test="$debug = 'yes'">
      <xsl:message select="'Number of lookaround paragraphs for inline field functions: ', $max-lookaround-count"/>
    </xsl:if>
    <!--<xsl:message select="'CCCCCCCCCCCCC ', count($non-block-field-begins), count($non-block-field-ends)"></xsl:message>-->
    <xsl:next-match>
      <!-- Pre-calculating all these params for large documents with many field functions, such as prEN_16815 -->
      <xsl:with-param name="field-begins" select="$field-begins" tunnel="yes"/>
      <xsl:with-param name="field-ends" select="$field-ends" tunnel="yes"/>
      <xsl:with-param name="non-block-field-begin-ids" select="$non-block-field-begin-ids" tunnel="yes"/>
      <xsl:with-param name="non-block-field-begins" select="$non-block-field-begins" tunnel="yes"/>
      <xsl:with-param name="non-block-field-ends" select="$non-block-field-ends" tunnel="yes"/>
      <xsl:with-param name="block-begins" select="$block-begins" tunnel="yes"/>
      <xsl:with-param name="block-begin-ids" select="$block-begins/@xml:id" as="xs:string*" tunnel="yes"/>
      <xsl:with-param name="lookaround-count" select="$max-lookaround-count" tunnel="yes"/>
    </xsl:next-match>
  </xsl:template>
  
  <xsl:key name="docx2hub:instrText-by-start-id" match="w:instrText[@docx2hub:fldChar-start-id]" 
    use="@docx2hub:fldChar-start-id"/>

  <xsl:variable name="docx2hub:block-field-functions" as="xs:string+" 
    select="('ADDRESSBLOCK', 'BIBLIOGRAPHY', 'CITAVI_XML', 'COMMENTS', 'DATABASE', 'INDEX', 'RD', 'TOA', 'TOC')"/>
  
  <xsl:variable name="docx2hub:hybrid-field-functions" as="xs:string+" 
    select="('IF')"/>
  
  <!-- Handle block field functions. The inline field functions will be handled when processing
    the individual current-group()s in docx2hub:field-functions mode, with tunneled begin/end
    field chars.  -->
  <xsl:template match="*[w:p]" mode="docx2hub:field-functions">
    <xsl:param name="field-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="field-ends" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="block-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="block-begin-ids" as="xs:string*" tunnel="yes"/>
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <!-- Assumption 1: Block field functions are not nested. There may be inline field functions
      nested within block field functions, but there are no block field functions within other block field functions.
      Therefore, we may use group-starting-with/group-ending-with to balance them, without recurring to recursion.
      Assumption 2: There is at most only one block field function begin per paragraph.
      Caveat: Since block field functions may be contained in table cells, we should only consider the 
      w:p/w:r/w:fldChars in this context, not arbitrarily deep .//w:fldChars (those in w:tbl/w:tc/w:p). -->
      <xsl:for-each-group select="*" group-starting-with="w:p[w:r/w:fldChar/@xml:id = $block-begin-ids]">
        <xsl:variable name="begin-fldChar-candidates" as="element(w:fldChar)*"
          select="w:r/w:fldChar[@xml:id = $block-begin-ids]"/>
        <xsl:variable name="begin-fldChar" as="element(w:fldChar)?"
          select="($begin-fldChar-candidates)[1]"/>
        <xsl:if test="count($begin-fldChar-candidates) gt 1">
          <xsl:message select="'More than one begin w:fldChar found for ', $block-begins/@xml:id"/>
        </xsl:if>
        <xsl:choose>
          <xsl:when test="exists($begin-fldChar)">
            <xsl:variable name="instr-text" as="element(w:instrText)" 
              select="key('docx2hub:instrText-by-start-id', $begin-fldChar/@xml:id)"/>
            <xsl:variable name="end-fldChar" as="element(w:fldChar)?" select="$field-ends[@linkend = $begin-fldChar/@xml:id]"/>
            <xsl:variable name="ffname" as="xs:string" select="$instr-text/@docx2hub:field-function-name"/>
            <xsl:variable name="ffargs" as="xs:string?" select="$instr-text/@docx2hub:field-function-args"/>
            <!-- GI 2010-16-13: It turns out that if the block end field function is at the beginning of a paragraph, then this
    paragraph must be excluded from the block. -->
            <xsl:variable name="end-p" as="element(w:p)*"
              select="for $p in current-group()/self::w:p[.//w:fldChar[@w:fldCharType = 'end']/@linkend = $end-fldChar/@linkend]
                      return if (
                                  empty($end-fldChar/ancestor::w:p[1]//w:t intersect $end-fldChar/parent::w:r/preceding::w:t)
                                  and
                                  exists($p/preceding-sibling::w:p[1])
                                )
                             then $p/preceding-sibling::w:p[1]
                             else $p"/>
            <xsl:for-each-group select="current-group()" group-ending-with="w:p[. is $end-p]">
              <xsl:choose>
                <xsl:when test="current-group()[last()] is $end-p">
                  <xsl:element name="{$ffname}">
                    <xsl:attribute name="fldArgs" select="$ffargs"/>
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
  
  <!--<xsl:function name="docx2hub:corresponding-end-fldChar" as="element(w:fldChar)?">
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
  </xsl:function>-->
  
  <!--<xsl:function name="docx2hub:corresponding-begin-fldChar" as="element(w:fldChar)">
    <xsl:param name="end" as="element(w:fldChar)"/>
    <xsl:sequence select="key('docx2hub:item-by-id', $end/@linkend, root($end))"/>
  </xsl:function>-->
  
  

  <!-- convert inline field functions to elements of the same names -->
  <xsl:template match="w:p | w:hyperlink | w:sdtContent[w:r/w:instrText]" mode="docx2hub:field-functions">
    <xsl:param name="non-block-field-begins" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="non-block-field-ends" as="element(w:fldChar)*" tunnel="yes"/>
    <xsl:param name="lookaround-count" as="xs:integer" tunnel="yes">
      <!-- Number of elments to look back or ahead from the current paragraph in order to find 
      inline field functions that span multiple paragraphs (HYPERLINK would be a candidate).
      This limit has been introduced because it would be too expensive to consider the whole
      document when looking for field begins or ends that stretch across the current para. -->
    </xsl:param>

<!--    <xsl:variable name="preceding-field-begins" as="element(w:fldChar)*" select="$non-block-field-begins[. &lt;&lt; current()]"/>
    <xsl:variable name="following-or-contained-field-ends" as="element(w:fldChar)*" select="$non-block-field-ends[. >> current()]"/>
    <xsl:variable name="last-in-para" as="element(*)" select="(current()/*[last()]/w:fldChar, current()/*[last()], current())[1]"/>
    <xsl:variable name="following-field-ends" as="element(w:fldChar)*" select="$non-block-field-ends[. >> $last-in-para]"/>
    <xsl:variable name="contained-field-begins" as="element(w:fldChar)*" select="$non-block-field-begins intersect .//w:fldChar"/>-->

    <xsl:variable name="preceding-field-begins" as="element(w:fldChar)*" select="preceding-sibling::*[position() le $lookaround-count]//w:fldChar[@w:fldCharType='begin']"/>
    <xsl:variable name="contained-field-begins" as="element(w:fldChar)*" select=".//w:fldChar[@w:fldCharType='begin']"/>
    <xsl:variable name="following-field-ends" as="element(w:fldChar)*" select="following-sibling::*[position() le $lookaround-count]//w:fldChar[@w:fldCharType='end']"/>
    <xsl:variable name="following-or-contained-field-ends" as="element(w:fldChar)*" select=".//w:fldChar[@w:fldCharType='end'] union $following-field-ends"/>
    <!--<xsl:variable name="last-in-para" as="element(*)" select="(current()/*[last()]/w:fldChar, current()/*[last()], current())[1]"/>-->
    
    <xsl:variable name="begins-before-ids" as="xs:string*" 
      select="$following-or-contained-field-ends[@linkend = $preceding-field-begins/@xml:id]/@linkend"/>
    <xsl:variable name="ends-after-ids" as="xs:string*" 
      select="$following-field-ends[@linkend = ($preceding-field-begins | $contained-field-begins)/@xml:id]/@xml:id"/>
    <xsl:variable name="begins-before-para" as="element(w:r)*">
      <xsl:for-each select="key('docx2hub:item-by-id', $begins-before-ids)">
        <w:r>
          <xsl:sequence select="., key('docx2hub:instrText-by-start-id', @xml:id)"/>
        </w:r>
      </xsl:for-each>
    </xsl:variable>
    <xsl:variable name="ends-after-para" as="element(*)*"
      select="key('docx2hub:item-by-id', $ends-after-ids)/.."/>
    <xsl:for-each select="$ends-after-para[not(name() = ('w:r', 'm:r'))]">
      <xsl:sequence select="docx2hub:message(., $fail-on-error = 'yes', false(), 'W2D_096', 'WRN', 'wml-to-dbk', 
                                             concat('Unexpected field function context ''', name(), ''' with content ''', string(.), ''''))"/>
    </xsl:for-each>
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
    <xsl:param name="non-block-field-begin-ids" as="xs:string*" tunnel="yes"/>
    <xsl:variable name="innermost-nesting-begins" as="element(w:fldChar)*" 
      select="for $f in ($para-contents/w:r/w:fldChar[@xml:id = $non-block-field-begin-ids])
              return $f[@xml:id
                        =
                        ($para-contents/w:r/w:fldChar[not(@w:fldCharType = 'separate')]
                                                     [. &gt;&gt; $f]
                        )[1]/@linkend
                       ]"/>
    <xsl:choose>
      <xsl:when test="empty($para-contents/*)"/>
      <xsl:when test="exists($innermost-nesting-begins)">
        <xsl:variable name="innermost-nesting-begin" as="element(w:fldChar)" select="$innermost-nesting-begins[1]"/>
        <xsl:variable name="innermost-nesting-end" as="element(w:fldChar)" 
          select="$para-contents/w:r/w:fldChar[@w:fldCharType = 'end'][@linkend =$innermost-nesting-begin/@xml:id]">
          <!-- don’t try to find it by key – the nodes in $para-contents are copies of the original document nodes --> 
        </xsl:variable>
        <xsl:variable name="instr-text" as="element(w:instrText)*" 
          select="key('docx2hub:instrText-by-start-id', $innermost-nesting-begin/@xml:id)[not(@docx2hub:field-function-error)]"/>
        <xsl:if test="empty($instr-text[matches(@docx2hub:field-function-name, 'MACROBUTTON|GOTOBUTTON')]) and count($instr-text) gt 1">
          <xsl:message select="'More than one instrText: '"/>
          <xsl:for-each select="$instr-text">
            <xsl:message select="'ID: ', string(@docx2hub:fldChar-start-id), ', Error: ', string(@docx2hub:field-function-error),
              ', Name: ', string(@docx2hub:field-function-name), ', Args: ', string(@docx2hub:field-function-args)"/>
          </xsl:for-each>
        </xsl:if>
        <xsl:variable name="ffname" as="xs:string" select="$instr-text[1]/@docx2hub:field-function-name"/>
        <xsl:variable name="ffargs" as="xs:string?" select="$instr-text[1]/@docx2hub:field-function-args"/>
        <xsl:call-template name="docx2hub:nest-inline-field-function">
          <xsl:with-param name="para-contents">
            <xsl:document>
              <xsl:sequence select="$para-contents/*[. &lt;&lt; $innermost-nesting-begin/..]"/>
              <xsl:variable name="inner" as="node()*" select="$para-contents/*[. &gt;&gt; $innermost-nesting-begin/..]
                                                                              [. &lt;&lt; $innermost-nesting-end/..]"/>
              <xsl:choose>
                <xsl:when test="exists($ffargs)"><!-- no error -->
                  <xsl:element name="{replace($ffname, '\\', '')}" xmlns="">
                    <xsl:attribute name="fldArgs" select="$ffargs"/>
                    <xsl:copy-of select="$inner/@css:*"/>
                    <xsl:if test="exists($instr-text[1]/*)">
                      <xsl:attribute name="docx2hub:contains-markup" select="'yes'"/>
                      <xsl:sequence select="$instr-text[1]/node()"/>
                    </xsl:if>
                    <xsl:sequence select="$inner, $instr-text[position() gt 1]/node()"/>
                  </xsl:element>
                  
                </xsl:when>
                <xsl:otherwise>
                  <xsl:sequence select="$inner"/>
                </xsl:otherwise>
              </xsl:choose>
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
  
  <!-- used to be handled in wml-to-dbk, earlier translation due to problems when used in 
       fieldfunctions  -->
  <xsl:template match="w:r[w:noBreakHyphen[following-sibling::w:instrText]]" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*, node() except w:noBreakHyphen" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:fldSimple[@docx2hub:field-function-name]" mode="docx2hub:field-functions">
    <xsl:element name="{@docx2hub:field-function-name}" xmlns="">
      <xsl:attribute name="fldArgs" select="@docx2hub:field-function-args"/>
      <xsl:apply-templates mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <!-- change to normal hyphen when used in fieldfunction &#x2d;-->
  <xsl:template match="w:instrText[ancestor::w:r[w:noBreakHyphen[following-sibling::w:instrText]]]" mode="docx2hub:remove-redundant-run-atts">
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:value-of select="'&#x2d;'"/>
      <xsl:apply-templates select="node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template match="w:p/w:r/@*" mode="docx2hub:remove-redundant-run-atts">
    <xsl:variable name="self-name" select="name()"/>
    <xsl:variable name="p-style" select="../../@role/key('style-by-name', .)/@*[name() = $self-name]" as="attribute()?"/>
    <xsl:variable name="r-style" select="../@role/key('style-by-name', .)/@*[name() = $self-name]" as="attribute()?"/>
    <xsl:choose>
      <xsl:when test="exists($r-style) and . = $r-style">
        <!-- inherit from run-style -->
      </xsl:when>
      <xsl:when test="exists($p-style) and . = $p-style and exists($r-style) and not(. = $r-style)">
        <!-- style would be the same like para, but different from run -->
        <xsl:copy/>
      </xsl:when>
      <xsl:when test="exists($p-style) and . = $p-style">
        <!-- inherit from para-style -->
      </xsl:when>
      <xsl:otherwise>
        <!-- cant inherit -->
        <xsl:copy/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="w:p/@*" mode="docx2hub:remove-redundant-run-atts">
    <xsl:variable name="self-name" select="name()"/>
    <xsl:variable name="p-style" select="../@role/key('style-by-name', .)/@*[name() = $self-name]" as="attribute()?"/>
    <xsl:choose>
      <xsl:when test="exists($p-style) and . = $p-style">
        <!-- inherit from para-style -->
      </xsl:when>
      <xsl:otherwise>
        <!-- cant inherit -->
        <xsl:copy/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

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
                       | *:phrase/w:numPr
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
  
  <xsl:template match="w:permStart | w:permEnd" mode="wml-to-dbk">
    <xsl:variable name="tokens" as="xs:string*" select="for $att in @* return concat($att/name(),'=',$att)"/>
    <xsl:processing-instruction name="tr" select="string-join((name(), $tokens), ' ')"/>
  </xsl:template>

  <xsl:template match="css:rule/w:tblPr" mode="wml-to-dbk">
    <xsl:apply-templates select="@*[not(some $pa in ../@*/name() satisfies $pa = name())]" mode="#current"/>
  </xsl:template>

  <xsl:template match="dbk:* | css:*" mode="wml-to-dbk" priority="-0.1">
     <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current" />
      <xsl:apply-templates select="node()" mode="#current" />
    </xsl:copy>
  </xsl:template>

  <xsl:template match="css:rule | *:style" mode="wml-to-dbk">
    <xsl:param name="content" as="element(*)*">
      <!-- linked-style, css:attic -->
    </xsl:param>
    <xsl:copy>
      <xsl:if test="w:numPr">
        <xsl:variable name="ilvl" select="w:numPr/w:ilvl/@w:val"/>
        <xsl:variable name="lvl-properties" select="key('abstract-numbering-by-id',key('numbering-by-id',w:numPr/w:numId/@w:val)/w:abstractNumId/@w:val)/w:lvl[@w:ilvl=$ilvl]"/>
        <xsl:apply-templates select="$lvl-properties/@* except $lvl-properties/@w:ilvl" mode="#current"/>
      </xsl:if>
      <xsl:apply-templates select="@*, w:tblPr, *[not(self::w:tblPr)], $content" mode="#current" />
    </xsl:copy>   
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
    <xsl:copy>
      <xsl:apply-templates select="@*" mode="#current"/>
      <xsl:variable name="most-frequent-lang" select="docx2hub:most-frequent-lang(.)" as="xs:string?"/>
      <xsl:if test="exists($most-frequent-lang)">
        <xsl:attribute name="xml:lang" select="$most-frequent-lang"/>
      </xsl:if>
      <xsl:apply-templates mode="#current">
        <xsl:with-param name="most-frequent-lang" select="$most-frequent-lang" as="xs:string?" tunnel="yes"/>
      </xsl:apply-templates>
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
  
  <!-- changes in this commit: because of vr_SB_525-12345_NESTOR-Testdaten-01 ($most-frequent-lang) -->
  <xsl:template match="w:p | w:r" mode="docx2hub:join-instrText-runs">
    <xsl:copy>
      <xsl:call-template name="docx2hub:adjust-lang"/>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
  </xsl:template>
  
  <xsl:template name="docx2hub:adjust-lang">
    <xsl:param name="most-frequent-lang" as="xs:string?" tunnel="yes"/>
    <xsl:choose>
      <xsl:when test="empty(@role | ancestor::w:p[1]/@role)
                      and
                      not($most-frequent-lang = ancestor-or-self::*[@xml:lang][1]/@xml:lang)">
        <xsl:copy-of select="ancestor-or-self::*[@xml:lang][1]/@xml:lang"/>
      </xsl:when>
      <xsl:when test="every $run in w:r[w:t[matches(., '\w')]]
                      satisfies (($run/@xml:lang, key('docx2hub:style-by-role', $run/@role))[1] = $most-frequent-lang)">
        <xsl:attribute name="xml:lang" select="$most-frequent-lang"/>
      </xsl:when>
    </xsl:choose>
  </xsl:template>
  
  <xsl:template match="@css:font-stretch[.='normal']" mode="docx2hub:join-instrText-runs">
    <!-- <w:w w:val="98"/> is an override, but it will be mapped to 'normal'. Unless this is different
      from the style’s font-stretch property, remove this property -->
    <!-- assuming they are only on w:r, not also on w:p -->
    <xsl:variable name="rule" as="element(css:rule)?" select="key('docx2hub:style-by-role', ../@role)"/>
    <xsl:if test="exists($rule/@css:font-stretch[not(.='normal')])">
      <xsl:next-match/>
    </xsl:if>
  </xsl:template>


  <xsl:template match="/dbk:*" mode="wml-to-dbk">
    <xsl:variable name="citavi-refs" as="document-node()?">
      <xsl:call-template name="docx2hub:citavi-json-to-xml"/>
    </xsl:variable>
    <xsl:copy>
      <xsl:apply-templates select="@*, *" mode="#current">
        <xsl:with-param name="citavi-refs" as="document-node()?" select="$citavi-refs" tunnel="yes"/>
      </xsl:apply-templates>
      <xsl:variable name="citavi-bib" as="element(dbk:biblioentry)*">
        <xsl:for-each-group select="$citavi-refs/docx2hub:citavi-jsons/fn:map/fn:array/fn:map/fn:map[@key = 'Reference']"
          group-by="fn:string[@key = 'Id']">
          <xsl:apply-templates select="." mode="citavi"/>
        </xsl:for-each-group>
      </xsl:variable>
      <xsl:if test="exists($citavi-bib)">
        <bibliography role="Citavi">
          <xsl:sequence select="$citavi-bib"/>
        </bibliography>
      </xsl:if>
      <xsl:if test="$debug = 'yes'">
        <xsl:sequence select="$citavi-refs/*"/>
      </xsl:if>
    </xsl:copy>
  </xsl:template>
  
  
  <!-- paragraphs (w:p) -->

  <xsl:variable name="docx2hub:allowed-para-element-names" as="xs:string+"
                select="('w:r', 
                         'w:pPr', 
                         'w:bookmarkStart', 
                         'w:bookmarkEnd', 
                         'w:smartTag', 
                         'w:commentRangeStart', 
                         'w:commentRangeEnd', 
                         'w:proofErr', 
                         'w:hyperlink', 
                         'w:del', 
                         'w:ins', 
                         'w:fldSimple', 
                         'm:oMathPara', 
                         'm:oMath')" />

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
    <xsl:message select="'[WARNING] $convert-footer is DEPRECATED: use p:option or xsl:param $include-header-and-footer instead.'"/>
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
    <xsl:variable name="normalized-string" as="xs:string?" select="docx2hub:normalize-name-for-id(.)"/>
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
  <xsl:template match="w:commentRangeStart" mode="wml-to-dbk">
    <anchor role="start" xml:id="comment_{@w:id}"/>
  </xsl:template>
  
  <xsl:template match="w:commentRangeEnd" mode="wml-to-dbk">
    <anchor role="end" xml:id="comment_{@w:id}_end"/>
  </xsl:template>
  
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
  
  <xsl:template match="@css:position[. = 'relative']" mode="wml-to-dbk" priority="1.5">
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
    <xsl:variable name="linkend" select="replace(@fldArgs, '^([_A-Za-z\d\.-]+)(.*)$', '$1')" as="xs:string"/>
    <xsl:variable name="switches" select="tokenize(@fldArgs,'\s+\\')[string-length(.)=1]" as="xs:string *"/>
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
      <!-- REF that is nested in another REF is ignored by Word. So do we.-->
      <xsl:when test="parent::REF[@fldArgs]">
        <xsl:apply-templates mode="#current"/>
      </xsl:when>
      <xsl:otherwise>
        <link linkend="{$linkend}">
          <xsl:if test="$switches[.='t']">
            <xsl:attribute name="xrefstyle" select="'numbering-only'"/>
          </xsl:if>
          <xsl:apply-templates mode="#current"/>
        </link>    
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>
  
  <!-- SYMBOL field function with translated characters already present in separate-markup -->
  <xsl:template match="SYMBOL[@fldArgs[matches(., '^\d+.*\\f\s&quot;?Symbol&quot;?')]]
                             [string-length(normalize-space(.)) = 1]
                             [key(
                               'symbol-by-number', 
                               tr:dec-to-hex(xs:integer(replace(@fldArgs, '^(\d+).+$', '$1'))), 
                               $symbol-font-map
                              )/@char = normalize-space(.)
                             ]" mode="wml-to-dbk" priority="2">
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>
  
  <xsl:variable name="mail-regex" as="xs:string" select="'^[-a-zA-Z0-9.!#$~_]+@[-a-zA-Z0-9]+\.[a-zA-Z0-9]+$'"/>
  
  <xsl:template match="*[@fldArgs]" mode="wml-to-dbk tables" name="docx2hub:default-field-function-handler">
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
          <xsl:when test="name() = ('EQ','eq')">
            <phrase role="docx2hub:EQ">
              <xsl:apply-templates select="@fldArgs, node()" mode="#current"/>
            </phrase>
          </xsl:when>
          <xsl:when test="name() = ('FORMCHECKBOX')">
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
                                              else tokenize(@fldArgs, ' ')[2]}">
                      <xsl:apply-templates select=".//@css:width | .//@css:height" mode="#current"/>
                    </imagedata>
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
         <xsl:when test="name() = 'NUMPAGES'">
            <!-- Ignore silently like Conditionally calculated field function above. 
            	Prospectively we could add a phrase for that to create a field again in docx for a better roundtripping -->
          </xsl:when>
          <xsl:when test="matches(@fldArgs,'^[\s&#160;]*$')">
            <xsl:apply-templates mode="#current"/>
          </xsl:when>
          <xsl:when test="name() = ('IF', 'PRINT', 'MACROBUTTON', 'GOTOBUTTON')">
            <!-- Conditionally calculated field function values are sometimes seen in figure or
            table counters. These are also included as the calculated value in the docx file. Therefore we see
            no immediate pressure to evaluate these expressions. -->
            <xsl:processing-instruction name="{name()}" select="string-join($tokens, ' ')"/>
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
          <xsl:when test="name() = 'LISTNUM'">
            <xsl:variable name="stdStyle" select="('DezimalStandard','NummerStandard','GliederungStandard','OutlineDefault','LegalDefault','NumberDefault')"/>
            <xsl:variable name="style" select="replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','')"/>
            <xsl:variable name="level" select="tokenize(tokenize(@fldArgs,'[\s&#160;]*\\l[\s&#160;]*')[2],'[\s&#160;]+')[matches(.,'^[0-9]*$')][1]"/>
            <xsl:variable name="current-level" select="//w:numbering/w:abstractNum[w:name/@w:val=$style]/w:lvl[number(@w:ilvl) + 1=
                                                                                                               number(if (exists($level)) 
                                                                                                                      then $level 
                                                                                                                      else '1')]"/>
            <xsl:variable name="number" select="if (.[tokenize(tokenize(@fldArgs,'[\s&#160;]*\\l[\s&#160;]*')[2],'[\s&#160;]+')[matches(.,'^[0-9]*$')][1] = $level]
                                                     [count(tokenize(@fldArgs,'\\s'))=2]) 
                                                then . 
                                                else preceding::LISTNUM[if ($style = $stdStyle) 
                                                                        then replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle 
                                                                        else replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $style]
                                                                       [tokenize(tokenize(@fldArgs,'[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                 '[\s&#160;]+')[matches(.,'^[0-9]*$')][1] = $level]
                                                                       [count(tokenize(@fldArgs,'\\s'))=2]
                                                                       [1]"/>
            <xsl:variable name="num-value" select="(if (exists($number)) 
                                                    then count(preceding::LISTNUM[if ($style = $stdStyle) 
                                                                                  then replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle 
                                                                                  else replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $style]
                                                                                 [tokenize(tokenize(@fldArgs, '[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                           '[\s&#160;]')[matches(.,'^[0-9]*$')][1] = $level]
                                                                                 [. &gt;&gt; $number]) + 
                                                         number(tokenize($number/@fldArgs,'\\s')[2]) + 
                                                         (if (generate-id(.) eq generate-id($number)) then 0 else 1) 
                                                    else count(preceding::LISTNUM[if ($style = $stdStyle) 
                                                                                  then replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle 
                                                                                  else replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $style]
                                                                                 [tokenize(tokenize(@fldArgs,'[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                           '[\s&#160;]')[matches(.,'^[0-9]*$')][1] = $level]) + 1) +
                                                   (if (($style = $stdStyle) and 
                                                        ($level = '1') and 
                                                        ((exists(preceding::LISTNUM[replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle]
                                                                                   [tokenize(tokenize(@fldArgs, '[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                             '[\s&#160;]')[matches(.,'^[0-9]*$')][1] gt $level]) and
                                                          not(exists(preceding::LISTNUM[replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle]
                                                                                       [tokenize(tokenize(@fldArgs, '[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                                 '[\s&#160;]')[matches(.,'^[0-9]*$')][1] = $level]))) or
                                                         (not(exists($number)) and
                                                          preceding::LISTNUM[replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = $stdStyle]
                                                                            [tokenize(tokenize(@fldArgs, '[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                      '[\s&#160;]')[matches(.,'^[0-9]*$')][1] = $level]
                                                                            [last()]
                                                                            [exists(preceding::LISTNUM[replace(tokenize(@fldArgs,'[\s&#160;]*\\')[1],'&quot;','') = 
                                                                                                       $stdStyle]
                                                                                                      [tokenize(tokenize(@fldArgs, '[\s&#160;]*\\l[\s&#160;]*')[2],
                                                                                                                '[\s&#160;]')[matches(.,'^[0-9]*$')][1] gt $level])]))) 
                                                    then 1 
                                                    else 0)"/>
            <xsl:variable name="lvltext" select="if (exists($current-level/w:lvlText)) 
                                                 then $current-level/w:lvlText/@w:val
                                                 else if ($style = $stdStyle)
                                                      then tr:get-listnum-lvltext($level,$style)
                                                      else '%1'" as="xs:string"/>
            <xsl:variable name="numfmt" select="if (exists($current-level/w:numFmt)) 
                                                then tr:get-numbering-format($current-level/w:numFmt/@w:val,'decimal')
                                                else if ($style = $stdStyle)
                                                     then tr:get-listnum-numfmt($level,$style)
                                                     else '1'"/>
            <xsl:variable name="act-value">
              <xsl:number format="{$numfmt}" value="$num-value"/>
            </xsl:variable>
            <xsl:value-of select="replace($lvltext, '%1', xs:string($act-value))"/>
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

  <xsl:template match="MACROBUTTON[starts-with(@fldArgs, 'MTPlaceRef')]" mode="wml-to-dbk" priority="1.5">
    <!-- MathType equation numbers, for ex. 
<MACROBUTTON xmlns="" fldArgs="MTPlaceRef \* MERGEFORMAT"><SEQ fldArgs="MTEqn \h \* MERGEFORMAT"
    /><w:bookmarkStart xmlns="http://docbook.org/ns/docbook" w:id="37" w:name="ZEqnNum377305"
    /><SEQ fldArgs="MTChap \c \* Arabic \* MERGEFORMAT"/><SEQ
    fldArgs="MTEqn \c \* Arabic \* MERGEFORMAT"/><w:bookmarkEnd
    xmlns="http://docbook.org/ns/docbook" w:id="37"/>(2.21)</MACROBUTTON> -->      
    <phrase role="hub:equation-number">
      <xsl:apply-templates mode="#current"/>  
    </phrase>
  </xsl:template>
  
  <xsl:template match="GOTOBUTTON[REF]
                                 [count(node()) = 1]
                                 [tokenize(@fldArgs, '\s+\\')[1] = tokenize(REF/@fldArgs, '\s+\\')[1]]" 
    mode="wml-to-dbk" priority="2">
    <!-- Otherwise, duplicate nested links -->
    <xsl:apply-templates mode="#current"/>
  </xsl:template>
  
  <xsl:template match="GOTOBUTTON[REF]
                                 [count(node()) = 2](: text node with combined field function args, and REF :)
                                 [tokenize(., '\s+\\')[1] = tokenize(REF/@fldArgs, '\s+\\')[1]]
                                 [tokenize(@fldArgs, '\s+\\')[1] = tokenize(REF/@fldArgs, '\s+\\')[1]]" 
    mode="wml-to-dbk" priority="2">
    <!-- http://svn.le-tex.de/svn/ltxbase/word2tex/trunk/testset/mantis-18038.docx -->
    <xsl:apply-templates select="REF" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="GOTOBUTTON/REF[tokenize(@fldArgs, '\s+\\')[1] = tokenize(../@fldArgs, '\s+\\')[1]]
                                     [empty(node())]" 
    mode="wml-to-dbk" priority="2"/>
  

  <xsl:template match="*[name() = ('PRINT', 'GOTOBUTTON', 'MACROBUTTON')][@docx2hub:contains-markup]" 
    mode="wml-to-dbk" priority="1.5">
    <xsl:call-template name="docx2hub:default-field-function-handler"/>
  </xsl:template>
  
  <xsl:template match="EQ[@docx2hub:contains-markup]/@fldArgs 
                     | eq[@docx2hub:contains-markup]/@fldArgs" mode="wml-to-dbk" priority="5"/>

  <xsl:template match="EQ[@docx2hub:contains-markup]" mode="wml-to-dbk" priority="2">
    <phrase role="docx2hub:EQ">
      <xsl:apply-templates select="@fldArgs, node()" mode="#current"/>
    </phrase>
  </xsl:template>

  <xsl:template match="EQ[empty(@docx2hub:contains-markup)]/@fldArgs 
                     | EQ/text() 
                     | eq[empty(@docx2hub:contains-markup)]/@fldArgs 
                     | eq/text()" mode="wml-to-dbk" priority="5">
    <xsl:analyze-string select="tr:EQ-string-to-unicode(replace(., '^\s*EQ\s+', '', 'i'))" regex="(\\\p{{Lu}}|[\(\);])">
      <xsl:matching-substring>
        <xsl:choose>
          <xsl:when test=". = '('">
            <open-delim/>
          </xsl:when>
          <xsl:when test=". = ')'">
            <close-delim/>
          </xsl:when>
          <xsl:when test=". = ';'">
            <sep/>
          </xsl:when>
          <xsl:when test=". = '\F'">
            <frac/>
          </xsl:when>
          <xsl:when test=". = '\R'">
            <root/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:message select="'Unknown EQ markup: ', ."/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:matching-substring>
      <xsl:non-matching-substring>
        <xsl:value-of select="."/>
      </xsl:non-matching-substring>
    </xsl:analyze-string>
  </xsl:template>
  

  <xsl:function name="tr:get-listnum-numfmt" as="xs:string">
    <xsl:param name="level"/>
    <xsl:param name="style"/>
    
    <xsl:choose>
      <xsl:when test="$style = ('LegalDefault','DezimalStandard')">
        <xsl:value-of select="'1'"/>
      </xsl:when>
      <xsl:when test="$style = ('OutlineDefault','GliederungStandard')">
        <xsl:choose>
          <xsl:when test="$level = '2'">
            <xsl:value-of select="'A'"/>
          </xsl:when>
          <xsl:when test="$level = ('3','5')">
            <xsl:value-of select="'1'"/>
          </xsl:when>
          <xsl:when test="$level = ('4','6','8')">
            <xsl:value-of select="'a'"/>
          </xsl:when>
          <xsl:when test="$level = ('7','9')">
            <xsl:value-of select="'i'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'I'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$level = ('2','5','8')">
            <xsl:value-of select="'a'"/>
          </xsl:when>
          <xsl:when test="$level = ('3','6','9')">
            <xsl:value-of select="'i'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'1'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
    
  </xsl:function>
  
  <xsl:function name="tr:get-listnum-lvltext" as="xs:string">
    <xsl:param name="level"/>
    <xsl:param name="style"/>
    
    <xsl:choose>
      <xsl:when test="$style = ('LegalDefault','DezimalStandard')">
        <xsl:choose>
          <xsl:when test="$level = '2'">
            <xsl:value-of select="'1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '3'">
            <xsl:value-of select="'1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '4'">
            <xsl:value-of select="'1.1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '5'">
            <xsl:value-of select="'1.1.1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '6'">
            <xsl:value-of select="'1.1.1.1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '7'">
            <xsl:value-of select="'1.1.1.1.1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '8'">
            <xsl:value-of select="'1.1.1.1.1.1.1.%1.'"/>
          </xsl:when>
          <xsl:when test="$level = '9'">
            <xsl:value-of select="'1.1.1.1.1.1.1.1.%1.'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'%1.'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:when test="$style = ('OutlineDefault','GliederungStandard')">
        <xsl:choose>
          <xsl:when test="$level = '4'">
            <xsl:value-of select="'%1)'"/>
          </xsl:when>
          <xsl:when test="$level = ('5','6','7','8','9')">
            <xsl:value-of select="'(%1)'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'%1.'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:when>
      <xsl:otherwise>
        <xsl:choose>
          <xsl:when test="$level = ('4','5','6')">
            <xsl:value-of select="'(%1)'"/>
          </xsl:when>
          <xsl:when test="$level = ('7','8','9')">
            <xsl:value-of select="'%1.'"/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="'%1)'"/>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:function>

  <xsl:variable name="tr:field-functions" as="document-node(element(tr:field-functions))">
    <xsl:document>
      <tr:field-functions>
        <tr:field-function name="BIBLIOGRAPHY" element="div" role="hub:bibliography"/>
        <tr:field-function name="CITATION"><!-- not implemented yet --></tr:field-function>
        <tr:field-function name="CITAVI_XML"/>
        <tr:field-function name="CITAVI_JSON"/>
        <tr:field-function name="INDEX" element="div" role="hub:index"/>
        <tr:field-function name="NOTEREF" element="link" attrib="linkend" value="1"/>
        <tr:field-function name="GOTOBUTTON" element="link" attrib="linkend" value="1"/>
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

  <xsl:variable name="w:fldSimple-REF-regex" as="xs:string" select="'^\s*REF\s(_?[A-Za-z]+\d+)\s?.+$'"/>
<!--  <xsl:variable name="w:fldSimple-REF-regex" as="xs:string" select="'^\s*REF\s(_[A-Za-z]+\d+)\s?.+\\h.+$'"/>
  Remove the \h (creates a hyperlink) restriction so that also non-hyperlinked references may be generated.
  Example: MathType equations in http://svn.le-tex.de/svn/ltxbase/word2tex/trunk/testset/mantis-18038.docx -->

  <xsl:template match="w:fldSimple[matches(@w:instr, $w:fldSimple-REF-regex)]" mode="wml-to-dbk" priority="1">
    <xsl:variable name="linkend" select="replace(@w:instr, $w:fldSimple-REF-regex, '$1', 'i')" as="xs:string"/>
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
      <xsl:apply-templates select="node()" mode="omml2mml">
        <xsl:with-param name="inline" select="false()" tunnel="yes"/>
      </xsl:apply-templates>
    </equation>
  </xsl:template>
  
  <xsl:template match="m:oMathPara[.//m:aln or .//w:br]" mode="wml-to-dbk">
    <equation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <mml:math display="block">
        <xsl:apply-templates select="m:oMath/@*" mode="omml2mml"/>
        <mml:mtable>
          <xsl:apply-templates select="node()" mode="omml2mml"/>
        </mml:mtable>
      </mml:math>
    </equation>
  </xsl:template>

  <xsl:template match="m:oMath" mode="wml-to-dbk">
    <inlineequation role="omml">
      <xsl:apply-templates select="@* except @srcpath" mode="#current"/>
      <xsl:apply-templates select="." mode="omml2mml">
        <xsl:with-param name="inline" select="true()" tunnel="yes"/>
      </xsl:apply-templates>
    </inlineequation>
  </xsl:template>
  
  <xsl:template match="w:object[mml:math]" mode="wml-to-dbk">
    <inlineequation role="mtef">
      <xsl:attribute name="condition" select="mml:math/@class"/>
      <xsl:apply-templates select="mml:math" mode="mathml"/>
    </inlineequation>
  </xsl:template>

  <!-- we used @class just as placeholder, the value was moved to inlineequation/@condition -->
  <xsl:template match="mml:math/@class" mode="mathml"/>
  
  <!-- inlineequation? remove block display setting! -->
  <xsl:template match="w:object/mml:math/@display[. eq 'block']" mode="mathml"/>
  
  <!-- identity template to preserve mathml text nodes -->
  <xsl:template match="mml:*" mode="mathml">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
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
  
  <xsl:template match="w:sdt[descendant::w:tc]
                            [ancestor::w:tbl]" mode="docx2hub:field-functions">
  <xsl:element name="w:tc">
    <xsl:copy>
      <xsl:apply-templates select="@*, node()" mode="#current"/>
    </xsl:copy>
   </xsl:element>
  </xsl:template>
  
  <xsl:template match="w:tc[ancestor::w:sdt[w:sdtPr/w:alias/@w:val or w:sdtPr/w:citation]]
                           [ancestor::w:tbl]" mode="docx2hub:field-functions">
    <xsl:message select="'eeeee'"></xsl:message>
    <xsl:apply-templates select="node()" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:sdt" mode="wml-to-dbk" priority="-1">
    <xsl:apply-templates select="w:sdtContent/*" mode="#current"/>
  </xsl:template>
  
  <xsl:template match="w:sdt[w:sdtPr/w:alias/@w:val or w:sdtPr/w:citation]
                            [empty(.//*:CITAVI_XML)]" mode="wml-to-dbk tables">
    <xsl:element name="blockquote">
      <xsl:attribute name="role" select="if (w:sdtPr/w:citation) then 'hub:citation' else w:sdtPr/w:alias/@w:val"/>
      <xsl:apply-templates select="w:sdtContent/*" mode="#current"/>
    </xsl:element>
  </xsl:template>
  
  <xsl:template match="CITAVI_JSON" mode="wml-to-dbk tables" 
    use-when="xs:decimal(system-property('xsl:version')) lt 3.0">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template match="CITAVI_XML" mode="wml-to-dbk tables"
    use-when="xs:decimal(system-property('xsl:version')) lt 3.0">
    <xsl:apply-templates mode="#current"/>
  </xsl:template>

  <xsl:template name="docx2hub:citavi-json-to-xml" use-when="xs:decimal(system-property('xsl:version')) lt 3.0"/>


  
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
